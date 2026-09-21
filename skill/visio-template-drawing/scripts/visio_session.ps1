# Shared lifecycle utilities. No active-object lookup, UI input, or clipboard calls.
function Initialize-VisioNative {
  if ('VisioSession.Native' -as [type]) { return }
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace VisioSession {
 public static class Native {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback, IntPtr param);
  public delegate bool EnumProc(IntPtr hwnd, IntPtr param);
  public static bool HasVisibleWindow(int pid) {
   bool visible = false;
   EnumWindows(delegate(IntPtr hwnd, IntPtr param) {
    uint owner; GetWindowThreadProcessId(hwnd, out owner);
    if (owner == pid && IsWindowVisible(hwnd)) visible = true;
    return true;
   }, IntPtr.Zero);
   return visible;
  }
 }
}
'@
}

function Write-VisioJson {
  param([string]$Path, $Value)
  # Replace atomically so the independent watchdog never reads half a state file.
  $tmp = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
  [IO.File]::WriteAllText($tmp, ($Value | ConvertTo-Json -Depth 16), (New-Object Text.UTF8Encoding($false)))
  if ([IO.File]::Exists($Path)) { [IO.File]::Replace($tmp, $Path, [System.Management.Automation.Language.NullString]::Value) }
  else { [IO.File]::Move($tmp, $Path) }
}

function Write-VisioEvent {
  param($Context, [string]$Stage, $Data = @{})
  $entry = [ordered]@{ run_id=$Context.run_id; time_utc=[DateTime]::UtcNow.ToString('o'); mode=$Context.mode; stage=$Stage; pid=$Context.visio_pid; data=$Data }
  [IO.File]::AppendAllText($Context.log_path, (($entry | ConvertTo-Json -Depth 16 -Compress) + [Environment]::NewLine), (New-Object Text.UTF8Encoding($false)))
}

function Get-VisioError {
  param($Record)
  return @{ type=$Record.Exception.GetType().FullName; message=$Record.Exception.Message; hresult=$Record.Exception.HResult; original=$Record.ToString(); exception=$Record.Exception.ToString(); stack=$Record.ScriptStackTrace; position=$Record.InvocationInfo.PositionMessage }
}

function Release-VisioCom {
  param($Object)
  if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
  }
}

function Test-VisioProcessIdentity {
  param([int]$ProcessId, [string]$StartUtc, [string]$Name = 'VISIO')
  $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
  return ($null -ne $process -and $process.ProcessName -ieq $Name -and $process.StartTime.ToUniversalTime().ToString('o') -eq $StartUtc)
}

function Register-VisioDocument {
  param($Document, [string]$Role = 'temporary')
  $key = [string]$Document.ID
  if (-not $script:VisioContext.documents.ContainsKey($key)) {
    $script:VisioContext.documents[$key] = @{ object=$Document; role=$Role; path=[string]$Document.FullName }
    Write-VisioEvent $script:VisioContext 'document_owned' @{ id=$key; role=$Role; path=[string]$Document.FullName }
  }
  return $Document
}

function Assert-VisioRunning {
  if (Test-Path -LiteralPath $script:VisioContext.cancel_path) {
    throw [OperationCanceledException]::new('Cancellation requested; closing task-owned documents.')
  }
  if ([DateTime]::UtcNow -gt $script:VisioContext.deadline) {
    throw [TimeoutException]::new('Visio job deadline exceeded.')
  }
}

function Invoke-VisioJob {
  param([hashtable]$Request, [string]$RunDirectory, [int]$TimeoutSeconds = 300)
  $ErrorActionPreference = 'Stop'
  if (-not $RunDirectory) {
    $base = if ($Request.output) { Split-Path -Parent ([IO.Path]::GetFullPath($Request.output)) } else { [IO.Path]::GetTempPath() }
    $RunDirectory = Join-Path $base ('.visio-runs\' + [guid]::NewGuid().ToString('N'))
  }
  $RunDirectory = [IO.Path]::GetFullPath($RunDirectory)
  if (Test-Path -LiteralPath $RunDirectory) { throw 'RunDirectory must be a new directory to protect earlier evidence.' }
  [void][IO.Directory]::CreateDirectory($RunDirectory)
  $Request.run_id = [guid]::NewGuid().ToString('N')
  $Request.run_directory = $RunDirectory
  $Request.timeout_seconds = $TimeoutSeconds
  $Request.supervisor_pid = $PID
  $Request.supervisor_start = (Get-Process -Id $PID).StartTime.ToUniversalTime().ToString('o')
  $Request.deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds).ToString('o')
  $requestPath = Join-Path $RunDirectory 'request.json'
  Write-VisioJson $requestPath $Request
  $hostExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
  $workerPath = Join-Path $PSScriptRoot 'visio_worker.ps1'
  $watchPath = Join-Path $PSScriptRoot 'visio_watchdog.ps1'
  # Encoded commands contain only literal, single-quoted paths, never user script text.
  $command = "& '" + $workerPath.Replace("'", "''") + "' -RequestPath '" + $requestPath.Replace("'", "''") + "'"
  $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
  $worker = Start-Process -FilePath $hostExe -ArgumentList @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-EncodedCommand',$encoded) -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $RunDirectory 'stdout.log') -RedirectStandardError (Join-Path $RunDirectory 'stderr.log')
  $workerIdentity = @{ pid=$worker.Id; start=$worker.StartTime.ToUniversalTime().ToString('o') }
  Write-VisioJson (Join-Path $RunDirectory 'worker.json') $workerIdentity
  $command = "& '" + $watchPath.Replace("'", "''") + "' -RunDirectory '" + $RunDirectory.Replace("'", "''") + "'"
  $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
  try {
    $watcher = Start-Process -FilePath $hostExe -ArgumentList @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-EncodedCommand',$encoded) -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $RunDirectory 'watchdog.stdout.log') -RedirectStandardError (Join-Path $RunDirectory 'watchdog.stderr.log')
    [IO.File]::WriteAllText((Join-Path $RunDirectory 'go'), 'watchdog started')
    while (-not $worker.WaitForExit(250)) { }
    if (-not $watcher.WaitForExit(30000)) { throw "Watchdog has not finished. Inspect $RunDirectory" }
    # Start-Process can expose a null ExitCode after WaitForExit on PS5.1.
    # The atomically written recovery record is the watchdog completion marker.
    $recoveryPath = Join-Path $RunDirectory 'recovery.json'
    if (-not (Test-Path -LiteralPath $recoveryPath)) { throw "Watchdog did not record completion. Inspect $RunDirectory" }
    $recovery = Get-Content -LiteralPath $recoveryPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $resultPath = Join-Path $RunDirectory 'result.json'
    if (-not (Test-Path -LiteralPath $resultPath)) { throw "Worker interrupted. Inspect recovery.json and logs in $RunDirectory" }
    $result = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $result.success) { throw "Visio job failed ($($result.status)). Original error and cleanup results: $resultPath" }
    if ($recovery.status -notin @('exited_or_pid_reused','intentionally_retained')) { throw "Unresolved watchdog disposition: $($recovery.status). Inspect $RunDirectory" }
    Write-Output "Visio job complete. Evidence: $RunDirectory"
    return $result
  } finally {
    if (-not $worker.HasExited) {
      [IO.File]::WriteAllText((Join-Path $RunDirectory 'cancel'), 'supervisor interrupted')
      # The independent watchdog handles stuck calls; never kill all Visio processes.
    }
    $worker.Dispose()
    if ($watcher) { $watcher.Dispose() }
  }
}

function Repair-VisioOrphan {
  param([string]$RunDirectory, [switch]$Recover)
  Initialize-VisioNative
  $statePath = Join-Path $RunDirectory 'state.json'
  if (-not (Test-Path -LiteralPath $statePath)) { return @{status='no_identity_record'; recovered=$false} }
  $state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
  if (-not $state.visio_pid) { return @{status='activation_not_identified'; recovered=$false; baseline=$state.baseline} }
  if (-not (Test-VisioProcessIdentity $state.visio_pid $state.visio_start)) { return @{status='exited_or_pid_reused'; recovered=$false; pid=$state.visio_pid} }
  if ($state.status -eq 'retained') { return @{status='intentionally_retained'; recovered=$false; pid=$state.visio_pid} }
  $ownerAlive = Test-VisioProcessIdentity $state.worker_pid $state.worker_start 'powershell'
  $eligible = ($state.owned -and $state.mode -eq 'background' -and -not $state.ever_visible -and -not $ownerAlive -and -not [VisioSession.Native]::HasVisibleWindow($state.visio_pid))
  if (-not $Recover -or -not $eligible) {
    return @{status='review_required'; eligible=$eligible; recovered=$false; pid=$state.visio_pid; reason='Recovery requires a dead worker and a verified exclusively owned, never-visible process. Visible or uncertain instances require user review.'}
  }
  # Recheck against PID reuse immediately before terminating this ONE owned process.
  if (-not (Test-VisioProcessIdentity $state.visio_pid $state.visio_start)) { throw 'Process identity changed during recovery.' }
  $process = Get-Process -Id $state.visio_pid
  $process.Kill()
  $exited = $process.WaitForExit(10000)
  $process.Dispose()
  return @{status= $(if ($exited) {'recovered'} else {'exit_timeout'}); recovered=$exited; pid=$state.visio_pid; unsaved_work='Task-only in-memory work may be lost; saved artifacts and logs are preserved.'}
}
