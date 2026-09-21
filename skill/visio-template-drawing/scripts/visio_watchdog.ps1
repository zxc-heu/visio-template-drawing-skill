param([Parameter(Mandatory=$true)][string]$RunDirectory)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\visio_session.ps1"
$request = Get-Content -LiteralPath (Join-Path $RunDirectory 'request.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$worker = Get-Content -LiteralPath (Join-Path $RunDirectory 'worker.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$cancelPath = Join-Path $RunDirectory 'cancel'
$cancelAt = $null
while (Test-VisioProcessIdentity $worker.pid $worker.start 'powershell') {
  $supervisor = Get-Process -Id $request.supervisor_pid -ErrorAction SilentlyContinue
  $supervisorAlive = ($supervisor -and $supervisor.StartTime.ToUniversalTime().ToString('o') -eq $request.supervisor_start)
  if (-not $supervisorAlive -or [DateTime]::UtcNow -gt [DateTime]::Parse($request.deadline).ToUniversalTime() -or (Test-Path -LiteralPath $cancelPath)) {
    if (-not $cancelAt) {
      $cancelAt = [DateTime]::UtcNow
      [IO.File]::WriteAllText($cancelPath, 'watchdog: cancellation, deadline, or supervisor exit')
    }
    if (([DateTime]::UtcNow - $cancelAt).TotalSeconds -gt 10) {
      if (Test-VisioProcessIdentity $worker.pid $worker.start 'powershell') { (Get-Process -Id $worker.pid).Kill() }
      break
    }
  }
  Start-Sleep -Milliseconds 250
}
Start-Sleep -Milliseconds 300
$recovery = Repair-VisioOrphan -RunDirectory $RunDirectory -Recover
Write-VisioJson (Join-Path $RunDirectory 'recovery.json') $recovery
