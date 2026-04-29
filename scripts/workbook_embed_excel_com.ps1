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

$outputParent = Split-Path -Parent $OutputWorkbook
if ($outputParent) {
  New-Item -ItemType Directory -Force -Path $outputParent | Out-Null
}

if ((Test-Path -LiteralPath $OutputWorkbook) -and -not $Overwrite) {
  throw "Output workbook already exists. Pass -Overwrite to replace: $OutputWorkbook"
}

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

    $cell.Value2 = $finding.Trim()

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

  $workbook.Save()
  Assert-WorkbookOutput -Worksheet $worksheet -Plan $plan -FindingCol $findingCol -ResultCol $resultCol
  if ($CleanupTmp) {
    Remove-ReportTmp -PlanPath $PlanJson -EvidenceRoot $EvidenceDir
  }
  Write-Output (@{
    output_workbook = $OutputWorkbook
    mode = "embed-images-excel-com"
    tmp_removed = [bool]$CleanupTmp
    shapes = $worksheet.Shapes.Count
    hyperlinks = $worksheet.Hyperlinks.Count
  } | ConvertTo-Json -Compress)
}
finally {
  if ($workbook) { $workbook.Close($true) | Out-Null }
  $excel.Quit()
  if ($worksheet) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null }
  if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
