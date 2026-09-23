param([Parameter(Mandatory=$true)][string]$Root)

$ErrorActionPreference = 'Stop'
$jobsDir = Join-Path $Root 'data\jobs'
$python = 'C:\Users\e.savko\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$template = Join-Path $Root 'data\templates\payment-invoices\229 WOFENG.xlsx'
$encoding = New-Object System.Text.UTF8Encoding($false)

New-Item -ItemType Directory -Force -Path $jobsDir | Out-Null
Write-Host 'Supplier Docs worker is running.'
while ($true) {
  $folders = Get-ChildItem -LiteralPath $jobsDir -Directory -ErrorAction SilentlyContinue
  foreach ($folder in $folders) {
    $requestPath = Join-Path $folder.FullName 'request.json'
    if (-not (Test-Path -LiteralPath $requestPath)) { continue }
    try { $job = Get-Content -LiteralPath $requestPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { continue }
    $task = [string]$job.taskId
    if ($job.status -ne 'queued') { continue }
    if (($task -ne 'check-final') -and ($task -ne 'check-final-corrections') -and ($task -ne 'standardize') -and ($task -ne 'commercial-offer') -and ($task -ne 'product-descriptions') -and ($task -ne 'translate-letters') -and ($task -ne 'payment-invoices')) { continue }
    if (-not ($job.PSObject.Properties.Name -contains 'progress')) { $job | Add-Member -NotePropertyName progress -NotePropertyValue 0 }
    if (-not ($job.PSObject.Properties.Name -contains 'stage')) { $job | Add-Member -NotePropertyName stage -NotePropertyValue 'Waiting...' }
    if (-not ($job.PSObject.Properties.Name -contains 'outputFiles')) { $job | Add-Member -NotePropertyName outputFiles -NotePropertyValue @() }
    if (-not ($job.PSObject.Properties.Name -contains 'errors')) { $job | Add-Member -NotePropertyName errors -NotePropertyValue @() }
    if (-not ($job.PSObject.Properties.Name -contains 'notes')) { $job | Add-Member -NotePropertyName notes -NotePropertyValue @() }
    $job.status = 'processing'
    $job.progress = 5
    $job.stage = 'Preparing files...'
    [System.IO.File]::WriteAllText($requestPath, ($job | ConvertTo-Json -Depth 12), $encoding)
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    if ($task -eq 'check-final') {
      $raw = & $python (Join-Path $Root 'scripts\check_final_documents.py') --job-dir $folder.FullName 2>&1
    } elseif ($task -eq 'check-final-corrections') {
      $raw = & $python (Join-Path $Root 'scripts\apply_final_corrections.py') --job-dir $folder.FullName 2>&1
    } elseif ($task -eq 'standardize') {
      $raw = & $python (Join-Path $Root 'scripts\standardize_final_documents.py') --job-dir $folder.FullName 2>&1
    } elseif ($task -eq 'commercial-offer') {
      $raw = & $python (Join-Path $Root 'scripts\process_commercial_offer.py') --job-dir $folder.FullName --templates-dir (Join-Path $Root 'data\templates') 2>&1
    } elseif ($task -eq 'product-descriptions') {
      $raw = & $python (Join-Path $Root 'scripts\process_product_descriptions.py') --job-dir $folder.FullName --templates-dir (Join-Path $Root 'data\templates') 2>&1
    } elseif ($task -eq 'translate-letters') {
      $raw = & $python (Join-Path $Root 'scripts\process_translation_letters.py') --job-dir $folder.FullName --templates-dir (Join-Path $Root 'data\templates') 2>&1
    } else {
      $raw = & $python (Join-Path $Root 'scripts\process_payment_invoices.py') --job-dir $folder.FullName --template $template 2>&1
    }
    $exitCode = $LASTEXITCODE
    $ErrorActionPreference = $previousErrorActionPreference
    $job = Get-Content -LiteralPath $requestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($exitCode -ne 0) {
      $errorText = ($raw | Out-String).Trim()
      $firstError = ($errorText -split '\r?\n' | Where-Object { $_.Trim() } | Select-Object -First 1).Trim()
      $job.status = 'failed'
      $job.progress = 100
      $job.stage = 'Processing failed'
      $job.errors = @($firstError)
    } else {
      $result = ($raw | Out-String) | ConvertFrom-Json
      if (-not ($job.PSObject.Properties.Name -contains 'notes')) { $job | Add-Member -NotePropertyName notes -NotePropertyValue @() }
      $job.status = 'completed'
      $job.progress = 100
      $job.stage = 'Done'
      $job.errors = @()
      $job.outputFiles = @($result.outputFiles | Where-Object { $_ })
      $job.notes = @($result.notes | Where-Object { $_ })
      if ($task -eq 'check-final') {
        if ($job.PSObject.Properties.Name -contains 'report') { $job.report = @($result.report) }
        else { $job | Add-Member -NotePropertyName report -NotePropertyValue @($result.report) }
      }
    }
    [System.IO.File]::WriteAllText($requestPath, ($job | ConvertTo-Json -Depth 12), $encoding)
  }
  Start-Sleep -Milliseconds 800
}
