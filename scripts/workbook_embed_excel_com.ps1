param(
  [Parameter(Mandatory=$true)][string]$SourceWorkbook,
  [Parameter(Mandatory=$true)][string]$PlanJson,
  [Parameter(Mandatory=$true)][string]$EvidenceDir,
  [Parameter(Mandatory=$true)][string]$OutputWorkbook,
  [string]$RunnerResultJson = "",
  [switch]$CleanupTmp,
  [switch]$Overwrite
)

$ErrorActionPreference = "Stop"

function New-UnicodeString([int[]]$CodePoints) {
  return -join ($CodePoints | ForEach-Object { [char]$_ })
}

function Assert-WorkbookWritable([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return }
  try {
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $stream.Close()
  } catch {
    throw "Output workbook is locked or open in Excel; refusing to write and claim success: $Path"
  } finally {
    if ($stream) { $stream.Dispose() }
  }
}

$AdminInterviewFinding = New-UnicodeString @(0x6D89,0x53CA,0x8BBF,0x8C08,0x7BA1,0x7406,0x5458,0xFF0C,0x672A,0x8FDB,0x884C,0x64CD,0x4F5C)
$AdminInterviewResult = New-UnicodeString @(0x672A,0x68C0,0x67E5)
$ScreenshotLabel = New-UnicodeString @(0x622A,0x56FE)
$EvidenceLabel = New-UnicodeString @(0x8BC1,0x636E)

function Convert-StatusToReportResult([string]$Status) {
  $text = $Status.Trim()
  switch ($text.ToLowerInvariant()) {
    "已检查" { return "符合" }
    "通过" { return "符合" }
    "合规" { return "符合" }
    "符合要求" { return "符合" }
    "不通过" { return "不符合" }
    "不合规" { return "不符合" }
    "不符合要求" { return "不符合" }
    "na" { return "不适用" }
    "n/a" { return "不适用" }
    "不涉及" { return "不适用" }
    "未执行" { return "未检查" }
    "未操作" { return "未检查" }
  }
  if ($text -in @("符合", "不符合", "不适用", "未检查")) {
    return $text
  }
  return ""
}

function ConvertTo-ColumnLetter([int]$ColumnNumber) {
  $letters = ""
  while ($ColumnNumber -gt 0) {
    $mod = ($ColumnNumber - 1) % 26
    $letters = [char](65 + $mod) + $letters
    $ColumnNumber = [math]::Floor(($ColumnNumber - $mod) / 26)
  }
  return $letters
}

function Remove-EvidenceText([string]$Text) {
  if ($null -eq $Text) { return "" }
  $clean = $Text -replace "`r`n", "`n"
  $clean = $clean -replace ("`n?" + [regex]::Escape($ScreenshotLabel) + "[:\uFF1A][\s\S]*$"), ""
  $clean = $clean -replace ("`n?" + [regex]::Escape($EvidenceLabel) + "[:\uFF1A][\s\S]*$"), ""
  $clean = $clean -replace "`n?row\d{2}_[^\s；;，,]+\.png", ""
  return $clean.Trim()
}

function ConvertTo-DeliveryFinding([string]$Text, [int]$Row) {
  $clean = Remove-EvidenceText $Text
  $clean = $clean -replace '^\s*(已检查|已通过|通过|使用|执行|检查方式|检查命令|命令输出)[^：:]{0,40}[：:]\s*', ''
  if ($clean -match '(?i)\brow\d{2}_.+\.png\b') {
    throw "row $Row finding still contains screenshot filename text"
  }
  if ($clean -match '[A-Za-z]:\\|\\\\[^\\]+\\') {
    throw "row $Row finding contains a filesystem/evidence path"
  }
  if ($clean -match '(?i)\b(audit|registry|reg|HKLM|HKCU|HKEY_|DWORD|REG_\w+|xxx)\s*=') {
    throw "row $Row finding contains raw machine field syntax; write a delivery conclusion instead"
  }
  return $clean.Trim()
}

function Get-ReportResult($Item) {
  if (Test-AdminInterviewSkip $Item) {
    return $AdminInterviewResult
  }
  foreach ($name in @("result", "compliance", "judgement", "judgment")) {
    if (($Item.PSObject.Properties.Name -contains $name) -and $Item.PSObject.Properties[$name].Value) {
      return [string]$Item.PSObject.Properties[$name].Value
    }
  }
  if (($Item.PSObject.Properties.Name -contains "status") -and $Item.status) {
    return Convert-StatusToReportResult ([string]$Item.status)
  }
  return ""
}

function Test-AdminInterviewSkip($Item) {
  if (($Item.PSObject.Properties.Name -contains "check_id") -and ([string]$Item.check_id -eq "admin_interview")) { return $true }
  if (($Item.PSObject.Properties.Name -contains "gui_action") -and ([string]$Item.gui_action -eq "skip_admin_interview")) { return $true }
  if (($Item.PSObject.Properties.Name -contains "skip_reason") -and ([string]$Item.skip_reason -eq "administrator_interview")) { return $true }
  return $false
}

function Merge-RunnerResultIntoPlan($Plan, [string]$RunnerResultPath) {
  if ([string]::IsNullOrWhiteSpace($RunnerResultPath)) { return }
  if (-not (Test-Path -LiteralPath $RunnerResultPath -PathType Leaf)) {
    throw "RunnerResultJson not found: $RunnerResultPath"
  }
  $runner = Get-Content -LiteralPath $RunnerResultPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $byRow = @{}
  foreach ($result in @($runner.results)) {
    if (($result.PSObject.Properties.Name -contains "row") -and $null -ne $result.row) {
      $byRow[[int]$result.row] = $result
    }
  }
  foreach ($item in @($Plan.items)) {
    $row = [int]$item.row
    if (-not $byRow.ContainsKey($row)) { continue }
    foreach ($property in $byRow[$row].PSObject.Properties) {
      if ($property.Name -eq "row") { continue }
      if ($item.PSObject.Properties.Name -contains $property.Name) {
        $item.PSObject.Properties[$property.Name].Value = $property.Value
      } else {
        $item | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value
      }
    }
  }
}

function Get-ExpectedFinding($Item) {
  if (Test-AdminInterviewSkip $Item) { return $AdminInterviewFinding }
  if (($Item.PSObject.Properties.Name -contains "finding") -and $Item.finding) { return [string]$Item.finding }
  if (($Item.PSObject.Properties.Name -contains "observed") -and $Item.observed) { return [string]$Item.observed }
  return ""
}

function Assert-WorkbookOutput($Worksheet, $Plan, [int]$FindingCol, [int]$ResultCol) {
  $errors = @()
  foreach ($item in @($Plan.items)) {
    $row = [int]$item.row
    $expectedFinding = Get-ExpectedFinding $item
    if (-not [string]::IsNullOrWhiteSpace($expectedFinding)) {
      $expectedFinding = ConvertTo-DeliveryFinding -Text $expectedFinding -Row $row
    }
    $expectedResult = Get-ReportResult $item
    $actualFinding = Remove-EvidenceText ([string]$Worksheet.Cells.Item($row, $FindingCol).Value2)
    $actualResult = ([string]$Worksheet.Cells.Item($row, $ResultCol).Value2).Trim()
    if (-not [string]::IsNullOrWhiteSpace($expectedFinding) -and -not $actualFinding.Contains($expectedFinding.Trim())) {
      $errors += "row $row finding mismatch"
    }
    if (-not [string]::IsNullOrWhiteSpace($expectedResult) -and $actualResult -ne $expectedResult) {
      $errors += "row $row result mismatch: expected '$expectedResult', got '$actualResult'"
    }
  }
  if ($errors.Count -gt 0) {
    throw "Workbook validation failed after write: $($errors[0..([Math]::Min($errors.Count, 8) - 1)] -join '; ')"
  }
}

function Assert-RemediationUnchanged($Excel, [string]$SourceWorkbook, $OutputWorkbook, $Plan) {
  if (-not (($Plan.PSObject.Properties.Name -contains "columns") -and $Plan.columns.remediation)) {
    return
  }
  $remediationCol = [int]$Plan.columns.remediation
  $sourceBook = $null
  $sourceSheet = $null
  $outputSheet = $null
  try {
    $sourceBook = $Excel.Workbooks.Open($SourceWorkbook, 3, $true)
    $sourceSheet = $sourceBook.Worksheets.Item($Plan.sheet)
    $outputSheet = $OutputWorkbook.Worksheets.Item($Plan.sheet)
    $sourceColumn = $sourceSheet.Columns.Item($remediationCol)
    $outputColumn = $outputSheet.Columns.Item($remediationCol)
    if ($sourceColumn.ColumnWidth -ne $outputColumn.ColumnWidth -or $sourceColumn.Hidden -ne $outputColumn.Hidden) {
      throw "整改建议 column dimensions changed"
    }
    foreach ($item in @($Plan.items)) {
      $row = [int]$item.row
      $sourceCell = $sourceSheet.Cells.Item($row, $remediationCol)
      $outputCell = $outputSheet.Cells.Item($row, $remediationCol)
      if ([string]$sourceCell.Value2 -ne [string]$outputCell.Value2) {
        throw "row $row 整改建议 value changed"
      }
      if ([string]$sourceCell.NumberFormat -ne [string]$outputCell.NumberFormat) {
        throw "row $row 整改建议 number format changed"
      }
      if ($sourceCell.Interior.Color -ne $outputCell.Interior.Color -or $sourceCell.Font.Color -ne $outputCell.Font.Color -or $sourceCell.Font.Bold -ne $outputCell.Font.Bold -or $sourceCell.Font.Name -ne $outputCell.Font.Name -or $sourceCell.Font.Size -ne $outputCell.Font.Size) {
        throw "row $row 整改建议 style changed"
      }
      if ($sourceCell.HorizontalAlignment -ne $outputCell.HorizontalAlignment -or $sourceCell.VerticalAlignment -ne $outputCell.VerticalAlignment -or $sourceCell.WrapText -ne $outputCell.WrapText) {
        throw "row $row 整改建议 alignment changed"
      }
      if ($sourceCell.Hyperlinks.Count -ne $outputCell.Hyperlinks.Count) {
        throw "row $row 整改建议 hyperlink count changed"
      }
    }
  }
  finally {
    if ($sourceBook) { $sourceBook.Close($false) | Out-Null }
    if ($sourceSheet) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sourceSheet) | Out-Null }
    if ($outputSheet) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outputSheet) | Out-Null }
    if ($sourceBook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sourceBook) | Out-Null }
  }
}

function Remove-ReportTmp([string]$PlanPath, [string]$EvidenceRoot) {
  $tmpRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PlanPath))
  $expectedTmp = [System.IO.Path]::GetFullPath((Join-Path $EvidenceRoot "tmp"))
  if (-not $tmpRoot.Equals($expectedTmp, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to remove tmp outside evidence dir: $tmpRoot"
  }
  if (Test-Path -LiteralPath $tmpRoot -PathType Container) {
    Remove-Item -LiteralPath $tmpRoot -Recurse -Force
  }
}

function Assert-EvidenceDirectory([object]$Plan, [string]$EvidenceRoot) {
  if (-not (Test-Path -LiteralPath $EvidenceRoot -PathType Container)) {
    throw "Evidence directory does not exist: $EvidenceRoot"
  }
  $missing = @()
  foreach ($item in @($Plan.items)) {
    if (Test-AdminInterviewSkip $item) {
      continue
    }
    foreach ($name in @($item.evidence)) {
      if ([string]::IsNullOrWhiteSpace([string]$name)) {
        continue
      }
      $path = Join-Path $EvidenceRoot ([string]$name)
      if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $missing += [string]$name
      }
    }
  }
  if ($missing.Count -gt 0) {
    throw "Evidence validation failed before tmp cleanup; missing files: $($missing[0..([Math]::Min($missing.Count, 8) - 1)] -join '; ')"
  }
}

$outputParent = Split-Path -Parent $OutputWorkbook
if ($outputParent) {
  New-Item -ItemType Directory -Force -Path $outputParent | Out-Null
}

if ((Test-Path -LiteralPath $OutputWorkbook) -and -not $Overwrite) {
  throw "Output workbook already exists. Pass -Overwrite to replace: $OutputWorkbook"
}
Assert-WorkbookWritable -Path $OutputWorkbook

Copy-Item -LiteralPath $SourceWorkbook -Destination $OutputWorkbook -Force

$plan = Get-Content -LiteralPath $PlanJson -Raw -Encoding UTF8 | ConvertFrom-Json
Merge-RunnerResultIntoPlan -Plan $plan -RunnerResultPath $RunnerResultJson
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  $workbook = $excel.Workbooks.Open($OutputWorkbook)
  $worksheet = $workbook.Worksheets.Item($plan.sheet)
  if ($null -eq $worksheet) {
    $worksheet = $workbook.Worksheets.Item(1)
  }

  $findingCol = if ($plan.columns.finding) { [int]$plan.columns.finding } else { 5 }
  $resultCol = if ($plan.columns.result) { [int]$plan.columns.result } else { 6 }

  foreach ($item in $plan.items) {
    $row = [int]$item.row
    $cell = $worksheet.Cells.Item($row, $findingCol)

    $finding = Remove-EvidenceText ([string]$cell.Value2)
    if ($item.PSObject.Properties.Name -contains "finding") { $finding = [string]$item.finding }
    elseif ($item.PSObject.Properties.Name -contains "observed") { $finding = [string]$item.observed }
    if (Test-AdminInterviewSkip $item) { $finding = $AdminInterviewFinding }

    $status = Get-ReportResult $item

    $evidence = @()
    if ((-not (Test-AdminInterviewSkip $item)) -and $item.evidence) {
      foreach ($entry in $item.evidence) {
        $evidence += [string]$entry
      }
    }

    $cell.Value2 = ConvertTo-DeliveryFinding -Text $finding -Row $row

    if ($status) {
      $resultCell = $worksheet.Cells.Item($row, $resultCol)
      $resultCell.Value2 = $status
    }

    if ($evidence.Count -eq 0) {
      continue
    }

    $multi = $evidence.Count -gt 1
    $maxWidth = if ($multi) { 303.0 } else { 420.0 }
    $maxHeight = if ($multi) { 245.0 } else { 340.0 }
    $textArea = 76.0
    $left = $cell.Left + 12.0
    $top = $cell.Top + $textArea

    foreach ($name in $evidence) {
      $imagePath = Join-Path $EvidenceDir $name
      if (-not (Test-Path -LiteralPath $imagePath)) {
        continue
      }
      $shape = $worksheet.Shapes.AddPicture($imagePath, $false, $true, $left, $top, -1, -1)
      $shape.AlternativeText = "vm-windows-security-audit evidence"
      $scale = [math]::Min($maxWidth / $shape.Width, $maxHeight / $shape.Height)
      if ($scale -lt 1) {
        $shape.Width = $shape.Width * $scale
        $shape.Height = $shape.Height * $scale
      }
      $shape.Placement = 1 # xlMoveAndSize
      $left += $shape.Width + 18.0
    }
  }

  $shapeCount = $worksheet.Shapes.Count
  $hyperlinkCount = $worksheet.Hyperlinks.Count
  $workbook.Save()
  $workbook.Close($true) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
  $workbook = $null
  $readbackWorkbook = $excel.Workbooks.Open($OutputWorkbook, 3, $true)
  $readbackWorksheet = $readbackWorkbook.Worksheets.Item($plan.sheet)
  Assert-WorkbookOutput -Worksheet $readbackWorksheet -Plan $plan -FindingCol $findingCol -ResultCol $resultCol
  Assert-RemediationUnchanged -Excel $excel -SourceWorkbook $SourceWorkbook -OutputWorkbook $readbackWorkbook -Plan $plan
  $readbackWorkbook.Close($false) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($readbackWorksheet) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($readbackWorkbook) | Out-Null
  $readbackWorksheet = $null
  $readbackWorkbook = $null
  if ($CleanupTmp) {
    Assert-EvidenceDirectory -Plan $plan -EvidenceRoot $EvidenceDir
    Remove-ReportTmp -PlanPath $PlanJson -EvidenceRoot $EvidenceDir
  }
  Write-Output (@{
    output_workbook = $OutputWorkbook
    mode = "embed-images-excel-com"
    tmp_removed = [bool]$CleanupTmp
    shapes = $shapeCount
    hyperlinks = $hyperlinkCount
  } | ConvertTo-Json -Compress)
}
finally {
  if ($readbackWorkbook) { $readbackWorkbook.Close($false) | Out-Null }
  if ($workbook) { $workbook.Close($true) | Out-Null }
  $excel.Quit()
  if ($readbackWorksheet) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($readbackWorksheet) | Out-Null }
  if ($readbackWorkbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($readbackWorkbook) | Out-Null }
  if ($worksheet) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null }
  if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
