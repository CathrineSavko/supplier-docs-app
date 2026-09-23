param([Parameter(Mandatory=$true)][string]$Root)

$ErrorActionPreference = 'Stop'
$JobsDir = Join-Path $Root 'data\jobs'
$TemplatePath = Join-Path $Root 'data\templates\payment-invoices\229 WOFENG.xlsx'
$PythonPath = 'C:\Users\e.savko\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Save-Job($Path, $Job) {
  [System.IO.File]::WriteAllText($Path, ($Job | ConvertTo-Json -Depth 12), $Utf8NoBom)
}

function Set-Progress($Job, $Percent, $Stage) {
  if (-not ($Job.PSObject.Properties.Name -contains 'progress')) { $Job | Add-Member -NotePropertyName progress -NotePropertyValue 0 }
  if (-not ($Job.PSObject.Properties.Name -contains 'stage')) { $Job | Add-Member -NotePropertyName stage -NotePropertyValue '' }
  $Job.progress = [int]$Percent
  $Job.stage = $Stage
}

function Invoke-Processor($TaskId, $JobFolder) {
  if ($TaskId -eq 'payment-invoices') {
    if (-not (Test-Path -LiteralPath $TemplatePath)) { throw 'Template 229 WOFENG.xlsx is not saved in Templates.' }
    $output = & $PythonPath (Join-Path $Root 'scripts\process_payment_invoices.py') --job-dir $JobFolder --template $TemplatePath 2>&1
  } elseif ($TaskId -eq 'check-final') {
    $output = & $PythonPath (Join-Path $Root 'scripts\check_final_documents.py') --job-dir $JobFolder 2>&1
  } elseif ($TaskId -eq 'standardize') {
    $output = & $PythonPath (Join-Path $Root 'scripts\standardize_final_documents.py') --job-dir $JobFolder 2>&1
  } else {
    throw "No processor is connected for task $TaskId."
  }
  if ($LASTEXITCODE -ne 0) { throw (($output | Out-String).Trim()) }
  return (($output | Out-String) | ConvertFrom-Json)
}

New-Item -ItemType Directory -Force -Path $JobsDir | Out-Null
Write-Host 'Supplier Docs worker is running.'
while ($true) {
  $folders = @(Get-ChildItem -LiteralPath $JobsDir -Directory -ErrorAction SilentlyContinue)
  foreach ($folder in $folders) {
    $requestPath = Join-Path $folder.FullName 'request.json'
    if (-not (Test-Path -LiteralPath $requestPath)) { continue }
    try { $job = Get-Content -LiteralPath $requestPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { continue }
    if ($job.status -ne 'queued') { continue }
    if (@('payment-invoices', 'check-final', 'standardize') -notcontains $job.taskId) { continue }
    $job.status = 'processing'
    Set-Progress $job 5 'Подготавливаю файлы…'
    Save-Job $requestPath $job
    try {
      $result = Invoke-Processor $job.taskId $folder.FullName
      if (-not ($job.PSObject.Properties.Name -contains 'notes')) { $job | Add-Member -NotePropertyName notes -NotePropertyValue @() }
      if (-not ($job.PSObject.Properties.Name -contains 'report')) { $job | Add-Member -NotePropertyName report -NotePropertyValue @() }
      if ($job.taskId -eq 'check-final') { $job.report = @($result.report) }
      $job.outputFiles = @($result.outputFiles | Where-Object { $_ -ne $null })
      $job.notes = @($result.notes | Where-Object { $_ -ne $null })
      $job.errors = @()
      $job.status = 'completed'
      Set-Progress $job 100 'Готово'
    } catch {
      $job.errors = @($_.Exception.Message)
      $job.status = 'failed'
      Set-Progress $job 100 'Не удалось обработать'
    }
    Save-Job $requestPath $job
  }
  Start-Sleep -Milliseconds 800
}
