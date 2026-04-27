param(
  [Parameter(Mandatory=$true)][string]$SourceWorkbook,
  [Parameter(Mandatory=$true)][string]$PlanJson,
  [Parameter(Mandatory=$true)][string]$EvidenceDir,
  [Parameter(Mandatory=$true)][string]$OutputWorkbook,
  [switch]$Overwrite
)

$ErrorActionPreference = "Stop"

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
  $clean = $clean -replace "`n?截图[:：][\s\S]*$", ""
  $clean = $clean -replace "`n?证据[:：][\s\S]*$", ""
  $clean = $clean -replace "`n?row\d{2}_[^\s；;，,]+\.png", ""
  return $clean.Trim()
}

function Get-ReportResult($Item) {
  foreach ($name in @("result", "compliance", "judgement", "judgment")) {
    if (($Item.PSObject.Properties.Name -contains $name) -and $Item.PSObject.Properties[$name].Value) {
      return [string]$Item.PSObject.Properties[$name].Value
    }
  }
  if (($Item.PSObject.Properties.Name -contains "status") -and $Item.status) {
    $status = [string]$Item.status
    if ($status -notin @("captured", "error", "skipped")) {
      return $status
    }
  }
  return ""
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

  $worksheet.Columns.Item($findingCol).ColumnWidth = [math]::Max($worksheet.Columns.Item($findingCol).ColumnWidth, 101.5)

  # Remove existing links and drawing objects in the copied workbook so reruns
  # do not stack stale evidence.
  while ($worksheet.Hyperlinks.Count -gt 0) {
    $worksheet.Hyperlinks.Item(1).Delete()
  }
  for ($i = $worksheet.Shapes.Count; $i -ge 1; $i--) {
    $worksheet.Shapes.Item($i).Delete()
  }

  foreach ($item in $plan.items) {
    $row = [int]$item.row
    $cell = $worksheet.Cells.Item($row, $findingCol)

    $finding = Remove-EvidenceText ([string]$cell.Value2)
    if ($item.PSObject.Properties.Name -contains "finding") { $finding = [string]$item.finding }
    elseif ($item.PSObject.Properties.Name -contains "observed") { $finding = [string]$item.observed }

    $status = Get-ReportResult $item

    $evidence = @()
    if ($item.evidence) {
      foreach ($entry in $item.evidence) {
        $evidence += [string]$entry
      }
    }

    $cell.Value2 = $finding.Trim()
    $cell.WrapText = $true
    $cell.VerticalAlignment = -4160 # xlTop

    if ($status) {
      $resultCell = $worksheet.Cells.Item($row, $resultCol)
      $resultCell.Value2 = $status
      $resultCell.WrapText = $true
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
    $maxPlacedHeight = 0.0

    foreach ($name in $evidence) {
      $imagePath = Join-Path $EvidenceDir $name
      if (-not (Test-Path -LiteralPath $imagePath)) {
        continue
      }
      $shape = $worksheet.Shapes.AddPicture($imagePath, $false, $true, $left, $top, -1, -1)
      $scale = [math]::Min($maxWidth / $shape.Width, $maxHeight / $shape.Height)
      if ($scale -lt 1) {
        $shape.Width = $shape.Width * $scale
        $shape.Height = $shape.Height * $scale
      }
      $shape.Placement = 1 # xlMoveAndSize
      $left += $shape.Width + 18.0
      if ($shape.Height -gt $maxPlacedHeight) {
        $maxPlacedHeight = $shape.Height
      }
    }

    if ($maxPlacedHeight -gt 0) {
      $worksheet.Rows.Item($row).RowHeight = [math]::Max($worksheet.Rows.Item($row).RowHeight, ($textArea + $maxPlacedHeight + 18.0))
    }
  }

  $workbook.Save()
  Write-Output (@{
    output_workbook = $OutputWorkbook
    mode = "embed-images-excel-com"
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
