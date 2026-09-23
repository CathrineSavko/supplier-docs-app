param(
  [Parameter(Mandatory=$true)][string]$Processor,
  [Parameter(Mandatory=$true)][string]$JobDir,
  [string]$Template
)

$Python = 'C:\Users\e.savko\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
if ($Template) {
  & $Python $Processor --job-dir $JobDir --template $Template
} else {
  & $Python $Processor --job-dir $JobDir
}
exit $LASTEXITCODE
