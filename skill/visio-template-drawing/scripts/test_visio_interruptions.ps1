param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
. "$PSScriptRoot\visio_session.ps1"
Initialize-VisioNative
$root=[IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $root) { throw 'Use a new test output directory.' }
[void][IO.Directory]::CreateDirectory($root)
$baseline=@(Get-Process VISIO -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
$draw=Join-Path $PSScriptRoot 'visio_draw_from_plan.ps1'
$single=Join-Path $PSScriptRoot '..\examples\case-08-background-flow.json'
$records=@()
$shapes=@(for ($i=0;$i -lt 3000;$i++) { @{id="s$i";kind='rectangle';x=(1+($i%50)*0.1);y=(1+[Math]::Floor($i/50)*0.05);width=0.08;height=0.04;text="s$i"} })
$large=Join-Path $root 'large.json'
Write-VisioJson $large @{page=@{name='Cancellation fixture';width=8;height=6};shapes=$shapes;connectors=@()}
function Assert-Test { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw $Message } }
function Save-Pass { param([string]$Name,$Data)
  $script:records+=@{test=$Name;passed=$true;data=$Data}
  Write-VisioJson (Join-Path $root 'interruptions.json') $script:records
  Write-Output "PASS $Name"
}
foreach ($mode in @('cancel','kill-worker','kill-supervisor','timeout')) {
  $run=Join-Path $root $mode
  $output=Join-Path $root "$mode.vsdx"
  $timeout=if ($mode -eq 'timeout') {5} else {120}
  $command="& '"+$draw.Replace("'","''")+"' -PlanJson '"+$large.Replace("'","''")+"' -OutputVsdx '"+$output.Replace("'","''")+"' -RunDirectory '"+$run.Replace("'","''")+"' -TimeoutSeconds $timeout"
  $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
  $supervisor=Start-Process powershell.exe -ArgumentList @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-EncodedCommand',$encoded) -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $root "$mode.stdout.log") -RedirectStandardError (Join-Path $root "$mode.stderr.log")
  $deadline=[DateTime]::UtcNow.AddSeconds(60)
  $state=$null
  try {
    while ([DateTime]::UtcNow -lt $deadline -and -not $supervisor.HasExited) {
      $log=Join-Path $run 'events.jsonl'
      if ((Test-Path -LiteralPath $log) -and (Get-Content -LiteralPath $log -Raw -Encoding UTF8) -match 'draw_page') {
        $state=Get-Content -LiteralPath (Join-Path $run 'state.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        break
      }
      Start-Sleep -Milliseconds 100
    }
    Assert-Test ($null -ne $state -and $state.owned) "Did not observe owned drawing stage: $mode"
    if ($mode -eq 'cancel') { [IO.File]::WriteAllText((Join-Path $run 'cancel'),'acceptance cancellation') }
    if ($mode -eq 'kill-worker') {
      Assert-Test (Test-VisioProcessIdentity $state.worker_pid $state.worker_start 'powershell') 'Worker identity changed.'
      (Get-Process -Id $state.worker_pid).Kill()
    }
    if ($mode -eq 'kill-supervisor') { $supervisor.Kill() }
    Assert-Test ($supervisor.WaitForExit(45000)) "Supervisor did not finish: $mode"
    $deadline=[DateTime]::UtcNow.AddSeconds(25)
    while (-not (Test-Path -LiteralPath (Join-Path $run 'recovery.json')) -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
    $recovery=Get-Content -LiteralPath (Join-Path $run 'recovery.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-Test (-not (Test-VisioProcessIdentity $state.visio_pid $state.visio_start)) "Task-owned Visio remained after $mode"
    if ($mode -eq 'kill-worker') { Assert-Test ($recovery.status -in @('recovered','exited_or_pid_reused')) 'Hard-interruption recovery failed.' }
    else {
      $result=Get-Content -LiteralPath (Join-Path $run 'result.json') -Raw -Encoding UTF8 | ConvertFrom-Json
      Assert-Test (-not $result.success -and $result.original_error) 'Cancellation/timeout was not reported as failure.'
    }
    Save-Pass $mode @{pid=$state.visio_pid;recovery=$recovery}
  } finally {
    if (-not $supervisor.HasExited -and (Test-Path -LiteralPath $run)) { [IO.File]::WriteAllText((Join-Path $run 'cancel'),'test teardown') }
    $supervisor.Dispose()
  }
}

# A locked template must fail without touching its content or leaking an instance.
$template=Join-Path $root 'template.vsdx'
& $draw -PlanJson $single -OutputVsdx $template -RunDirectory (Join-Path $root 'template-create') | Out-Null
$before=(Get-FileHash -LiteralPath $template).Hash
$locked=[IO.File]::Open($template,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
try {
  try { & $draw -PlanJson $single -TemplateVsdx $template -OutputVsdx (Join-Path $root 'locked-output.vsdx') -RunDirectory (Join-Path $root 'locked-template') | Out-Null } catch { }
} finally { $locked.Dispose() }
$result=Get-Content -LiteralPath (Join-Path $root 'locked-template\result.json') -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-Test (-not $result.success -and $result.original_error -and -not (Get-Process -Id $result.pid -ErrorAction SilentlyContinue)) 'Locked template was not handled safely.'
Assert-Test ((Get-FileHash -LiteralPath $template).Hash -eq $before) 'Locked template changed.'
Save-Pass 'locked-template' $result.original_error

# Recovery cannot act on a pre-existing process, reused PID, retained or visible instance.
$state=Get-Content -LiteralPath (Join-Path $root 'cancel\state.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$state.visio_pid=$PID; $state.visio_start=(Get-Process -Id $PID).StartTime.ToUniversalTime().ToString('o')
$fake=Join-Path $root 'identity-refusal';[void][IO.Directory]::CreateDirectory($fake)
Write-VisioJson (Join-Path $fake 'state.json') $state
$refusal=Repair-VisioOrphan -RunDirectory $fake -Recover
Assert-Test ($refusal.status -eq 'exited_or_pid_reused' -and -not $refusal.recovered) 'Recovery accepted a non-Visio identity.'
Save-Pass 'recovery-refuses-identity-mismatch' $refusal
$extra=@(Get-Process VISIO -ErrorAction SilentlyContinue | Where-Object {$baseline -notcontains $_.Id} | ForEach-Object {$_.Id})
Assert-Test ($extra.Count -eq 0) 'Unexpected Visio processes remain.'
Save-Pass 'no-orphans' @{extra=$extra}
