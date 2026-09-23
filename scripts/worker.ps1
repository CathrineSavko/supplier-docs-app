param([Parameter(Mandatory=$true)][string]$Root)

$ErrorActionPreference = 'Stop'
$JobsDir = Join-Path $Root 'data\jobs'
$TemplatePath = Join-Path $Root 'data\templates\payment-invoices\229 WOFENG.xlsx'
$PythonPath = 'C:\Users\e.savko\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$utf8NoBom = New-Object Text.UTF8Encoding($false)

function Save-Job($Path, $Job) {
  [IO.File]::WriteAllText($Path, ($Job | ConvertTo-Json -Depth 8), $utf8NoBom)
}

function Set-Progress($Job, $Percent, $Stage) {
  if (-not ($Job.PSObject.Properties.Name -contains 'progress')) { $Job | Add-Member -NotePropertyName progress -NotePropertyValue 0 }
  if (-not ($Job.PSObject.Properties.Name -contains 'stage')) { $Job | Add-Member -NotePropertyName stage -NotePropertyValue '' }
  $Job.progress = $Percent
  $Job.stage = $Stage
}

New-Item -ItemType Directory -Force -Path $JobsDir | Out-Null
Write-Host 'Supplier Docs worker is running.'
while ($true) {
  $jobs = @(Get-ChildItem -LiteralPath $JobsDir -Directory -ErrorAction SilentlyContinue)
  foreach ($jobFolder in $jobs) {
    $requestPath = Join-Path $jobFolder.FullName 'request.json'
    if (-not (Test-Path -LiteralPath $requestPath)) { continue }
    try { $job = Get-Content -LiteralPath $requestPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { continue }
    if ($job.status -ne 'queued' -or $job.taskId -notin @('payment-invoices', 'check-final', 'standardize')) { continue }
    $job.status = 'processing'; Set-Progress $job 5 'Подготавливаю файлы…'; Save-Job $requestPath $job
    try {
      if ($job.taskId -eq 'payment-invoices') {
        if (-not (Test-Path -LiteralPath $TemplatePath)) { throw 'Template 229 WOFENG.xlsx is not saved in Templates.' }
        $rawResult = & $PythonPath (Join-Path $Root 'scripts\process_payment_invoices.py') --job-dir $jobFolder.FullName --template $TemplatePath 2>&1
      } elseif ($job.taskId -eq 'check-final') {
        $rawResult = & $PythonPath (Join-Path $Root 'scripts\check_final_documents.py') --job-dir $jobFolder.FullName 2>&1
      } else {
        $rawResult = & $PythonPath (Join-Path $Root 'scripts\standardize_final_documents.py') --job-dir $jobFolder.FullName 2>&1
      }
      if ($LASTEXITCODE -ne 0) { throw (($rawResult | Out-String).Trim()) }
      $result = $rawResult | ConvertFrom-Json
      if (-not ($job.PSObject.Properties.Name -contains 'notes')) { $job | Add-Member -NotePropertyName notes -NotePropertyValue @() }
      if ($job.taskId -eq 'check-final') {
        if (-not ($job.PSObject.Properties.Name -contains 'report')) { $job | Add-Member -NotePropertyName report -NotePropertyValue @() }
        $job.report = @($result.report)
      }
      $job.status = 'completed'; Set-Progress $job 100 'Готово'; $job.outputFiles = @($result.outputFiles | Where-Object { $_ }); $job.notes = @($result.notes | Where-Object { $_ }); $job.errors = @()
    } catch {
      $job.status = 'failed'; Set-Progress $job 100 'Не удалось обработать'; $job.errors = @($_.Exception.Message)
    }
    Save-Job $requestPath $job
  }
  Start-Sleep -Milliseconds 800
}
