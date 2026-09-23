param(
  [Parameter(Mandatory=$true)][string]$JobDir,
  [Parameter(Mandatory=$true)][string]$TemplatePath
)

$ErrorActionPreference = 'Stop'
$inputDir = Join-Path $JobDir 'input'
$outputDir = Join-Path $JobDir 'output'
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
$inputs = @(Get-ChildItem -LiteralPath $inputDir -File | Where-Object { $_.Extension -ieq '.xlsx' })
if ($inputs.Count -eq 0) { throw 'Task 6 needs source .xlsx files.' }

function Get-OutputName([string]$BaseName) {
  if ($BaseName -match '^\d+\s*[A-Za-z]+') { return "$BaseName.xlsx" }
  $numbers = [regex]::Match($BaseName, '\d+')
  $letters = [regex]::Match($BaseName, '[A-Za-z]+')
  if ($numbers.Success -and $letters.Success) { return "$($numbers.Value) $($letters.Value.ToUpperInvariant()).xlsx" }
  return "$BaseName.xlsx"
}

function Release-Com($object) { if ($null -ne $object) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) } }

function Get-TemplateAssets([string]$WorkbookPath, [string]$WorkingDirectory) {
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  New-Item -ItemType Directory -Force -Path $WorkingDirectory | Out-Null
  $templateCopy = Join-Path $WorkingDirectory 'template-copy.xlsx'
  $input = [IO.File]::Open($WorkbookPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
  $output = [IO.File]::Open($templateCopy, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::Read)
  try { $input.CopyTo($output) } finally { $output.Dispose(); $input.Dispose() }
  $archive = [IO.Compression.ZipFile]::OpenRead($templateCopy)
  try {
    $entry = $archive.Entries | Where-Object { $_.FullName -like 'xl/media/*' } | Select-Object -First 1
    if ($null -eq $entry) { throw 'No stamp image was found in template 229 WOFENG.xlsx.' }
    $target = Join-Path $WorkingDirectory ([IO.Path]::GetFileName($entry.FullName))
    [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, $true)
    return [PSCustomObject]@{ TemplatePath = $templateCopy; StampPath = $target }
  } finally { $archive.Dispose() }
}

function Copy-TemplateBlock($SourceSheet, $TargetSheet, [int]$TargetStartRow) {
  $sourceStartRow = 21; $rowCount = 8; $columnCount = 8
  $targetRange = $TargetSheet.Range($TargetSheet.Cells.Item($TargetStartRow,1), $TargetSheet.Cells.Item($TargetStartRow + $rowCount - 1,$columnCount))
  $targetRange.UnMerge() | Out-Null
  $mergedAreas = @{}
  for ($rowOffset = 0; $rowOffset -lt $rowCount; $rowOffset++) {
    $TargetSheet.Rows.Item($TargetStartRow + $rowOffset).RowHeight = $SourceSheet.Rows.Item($sourceStartRow + $rowOffset).RowHeight
    for ($column = 1; $column -le $columnCount; $column++) {
      $source = $SourceSheet.Cells.Item($sourceStartRow + $rowOffset, $column)
      $target = $TargetSheet.Cells.Item($TargetStartRow + $rowOffset, $column)
      $target.Formula = $source.Formula
      $target.NumberFormat = $source.NumberFormat
      $target.VerticalAlignment = $source.VerticalAlignment
      $target.WrapText = $source.WrapText
      $target.Orientation = $source.Orientation
      $target.IndentLevel = $source.IndentLevel
      $target.Font.Name = $source.Font.Name
      $target.Font.Size = $source.Font.Size
      $target.Font.Bold = $source.Font.Bold
      $target.Font.Italic = $source.Font.Italic
      $target.Font.Color = $source.Font.Color
      $target.Interior.Pattern = $source.Interior.Pattern
      $target.Interior.Color = $source.Interior.Color
      foreach ($borderIndex in 7,8,9,10) {
        $target.Borders.Item($borderIndex).LineStyle = $source.Borders.Item($borderIndex).LineStyle
        $target.Borders.Item($borderIndex).Weight = $source.Borders.Item($borderIndex).Weight
        $target.Borders.Item($borderIndex).Color = $source.Borders.Item($borderIndex).Color
      }
      if ($source.MergeCells) {
        $area = $source.MergeArea
        $key = "$($area.Row):$($area.Column):$($area.Rows.Count):$($area.Columns.Count)"
        $mergedAreas[$key] = $area
      }
    }
  }
  foreach ($area in $mergedAreas.Values) {
    $row = $TargetStartRow + $area.Row - $sourceStartRow
    $column = $area.Column
    $TargetSheet.Range($TargetSheet.Cells.Item($row,$column), $TargetSheet.Cells.Item($row+$area.Rows.Count-1,$column+$area.Columns.Count-1)).Merge() | Out-Null
  }
  for ($rowOffset = 0; $rowOffset -lt $rowCount; $rowOffset++) {
    for ($column = 1; $column -le $columnCount; $column++) { $TargetSheet.Cells.Item($TargetStartRow + $rowOffset,$column).HorizontalAlignment = -4131 }
  }
}

$templateAssets = Get-TemplateAssets $TemplatePath (Join-Path $JobDir 'working')
$stampFile = $templateAssets.StampPath
$excel = $null; $templateBook = $null; $templateSheet = $null
$outputFiles = @(); $notes = @()
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false
  $templateBook = $excel.Workbooks.Open($templateAssets.TemplatePath)
  $templateSheet = $templateBook.Worksheets.Item('INV')
  $templateBlock = $templateSheet.Range('A21:H28')
  foreach ($input in $inputs) {
    $destinationName = Get-OutputName $input.BaseName
    $destination = Join-Path $outputDir $destinationName
    if (Test-Path -LiteralPath $destination) { throw "Output file already exists: $destinationName" }
    Copy-Item -LiteralPath $input.FullName -Destination $destination
    $book = $null; $sheet = $null
    try {
      $book = $excel.Workbooks.Open($destination)
      if ($book.Worksheets.Count -gt 1) {
        for ($sheetIndex = $book.Worksheets.Count; $sheetIndex -ge 1; $sheetIndex--) {
          $candidate = $book.Worksheets.Item($sheetIndex)
          if ($candidate.Name -eq 'PL') { $candidate.Delete() }
        }
      }
      $sheet = $book.Worksheets.Item('INV')
      if ($null -eq $sheet) { throw "INV sheet is missing in $($input.Name)." }
      for ($shapeIndex = $sheet.Shapes.Count; $shapeIndex -ge 1; $shapeIndex--) { $sheet.Shapes.Item($shapeIndex).Delete() }
      $lastRow = [Math]::Max(33, $sheet.UsedRange.Row + $sheet.UsedRange.Rows.Count - 1)
      $lastColumn = [Math]::Max(9, $sheet.UsedRange.Column + $sheet.UsedRange.Columns.Count - 1)
      if ($lastRow -gt 33) { $sheet.Range($sheet.Cells.Item(34,1), $sheet.Cells.Item($lastRow,$lastColumn)).ClearContents() }
      if ($lastColumn -gt 9) { $sheet.Range($sheet.Cells.Item(1,10), $sheet.Cells.Item(33,$lastColumn)).ClearContents() }
      $totalRow = $null
      for ($row = 1; $row -le 33 -and $null -eq $totalRow; $row++) {
        for ($column = 1; $column -le 9; $column++) {
          $value = $sheet.Cells.Item($row,$column).Text
          if ($null -ne $value -and "$value" -match '(?i)total') { $totalRow = $row; break }
        }
      }
      if ($null -eq $totalRow) { throw "Total row is missing in $($input.Name)." }
      $blockStart = $totalRow + 2
      if ($blockStart + 7 -gt 33) { throw "The block after Total does not fit A1:I33 in $($input.Name)." }
      Copy-TemplateBlock $templateSheet $sheet $blockStart
      $sheet.Columns.Item(1).ColumnWidth = 6
      $beneficiary = $null; $sellerTerms = $null
      for ($row = 1; $row -le 33; $row++) {
        for ($column = 1; $column -le 9; $column++) {
          $cellText = "$($sheet.Cells.Item($row,$column).Text)"
          if ($null -eq $beneficiary -and $cellText -match '(?i)Beneficiary\s+Address') { $beneficiary = $sheet.Cells.Item($row,$column) }
          if ($null -eq $sellerTerms -and ($cellText -replace '\s+',' ') -match '(?i)SELLER OTHER TERMS') { $sellerTerms = $sheet.Cells.Item($row,$column) }
        }
      }
      if ($null -eq $beneficiary) { throw "Beneficiary Address row is missing in $($input.Name)." }
      $beneficiaryRow = $beneficiary.Row
      $sheet.Range($sheet.Cells.Item($beneficiaryRow,1),$sheet.Cells.Item($beneficiaryRow,8)).UnMerge()
      $sheet.Range($sheet.Cells.Item($beneficiaryRow,1),$sheet.Cells.Item($beneficiaryRow,8)).Merge()
      $sheet.Rows.Item($beneficiaryRow).RowHeight = 29.25
      $sheet.Cells.Item($beneficiaryRow,1).HorizontalAlignment = -4131
      if ($null -eq $sellerTerms) { throw "SELLER OTHER TERMS cell is missing in $($input.Name)." }
      $sellerRow = $sellerTerms.Row; $sellerColumn = $sellerTerms.Column
      if ($sellerRow + 2 -gt 33) { throw 'SELLER OTHER TERMS cannot be merged with two rows below.' }
      $sellerRange = $sheet.Range($sheet.Cells.Item($sellerRow,$sellerColumn),$sheet.Cells.Item($sellerRow+2,$sellerColumn))
      $sellerRange.UnMerge(); $sellerRange.Merge(); $sellerRange.HorizontalAlignment = -4131
      $boxLeft = $sheet.Cells.Item(20,4).Left; $boxTop = $sheet.Cells.Item([Math]::Max(20,$totalRow+1),4).Top
      $boxRight = $sheet.Cells.Item(29,7).Left + $sheet.Cells.Item(29,7).Width
      $boxBottom = $sheet.Cells.Item(29,7).Top + $sheet.Cells.Item(29,7).Height
      if ($boxTop -ge $boxBottom) { throw "No room for the stamp in D20:G29 in $($input.Name)." }
      $stamp = $sheet.Shapes.AddPicture($stampFile, $false, $true, $boxLeft, $boxTop, -1, -1)
      $scale = [Math]::Min(($boxRight-$boxLeft)/$stamp.Width, ($boxBottom-$boxTop)/$stamp.Height)
      $stamp.LockAspectRatio = -1
      $stamp.Width = $stamp.Width * $scale
      $stamp.Left = $boxLeft + (($boxRight-$boxLeft-$stamp.Width)/2)
      $stamp.Top = $boxTop + (($boxBottom-$boxTop-$stamp.Height)/2)
      $excel.CutCopyMode = 0
      $book.Save() | Out-Null; $book.Close($true) | Out-Null
      Release-Com $sheet; $sheet = $null; Release-Com $book; $book = $null
      $checkBook = $excel.Workbooks.Open($destination)
      $checkSheet = $checkBook.Worksheets.Item('INV')
      if ($checkBook.Worksheets.Count -ne 1) { throw "Validation failed: extra sheets remain in $destinationName." }
      if ($checkSheet.Shapes.Count -ne 1) { throw "Validation failed: wrong number of images in $destinationName." }
      $verifiedStamp = $checkSheet.Shapes.Item(1)
      if ($verifiedStamp.Left -lt $boxLeft -or $verifiedStamp.Top -lt $boxTop -or ($verifiedStamp.Left + $verifiedStamp.Width) -gt $boxRight -or ($verifiedStamp.Top + $verifiedStamp.Height) -gt $boxBottom) { throw 'Validation failed: stamp is outside D20:G29.' }
      $checkBook.Close($false) | Out-Null; Release-Com $checkSheet; Release-Com $checkBook
      $outputFiles += $destinationName; $notes += "${destinationName}: INV prepared from template 229 WOFENG."
    } finally {
      if ($null -ne $sheet) { try { $sheet | Out-Null } catch {} }
      if ($null -ne $book) { try { $book.Close($false) } catch {}; Release-Com $book }
    }
  }
} finally {
  if ($null -ne $templateBook) { try { $templateBook.Close($false) } catch {} }
  Release-Com $templateSheet; Release-Com $templateBook
  if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-Com $excel }
}
[PSCustomObject]@{ outputFiles = $outputFiles; notes = $notes } | ConvertTo-Json -Compress
