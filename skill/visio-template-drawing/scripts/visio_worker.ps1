param([Parameter(Mandatory=$true)][string]$RequestPath)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\visio_session.ps1"
. "$PSScriptRoot\visio_render.ps1"
Initialize-VisioNative
$request = Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$run = $request.run_directory
$script:VisioContext = @{
  run_id=$request.run_id; mode=$(if ($request.visible) {'visible'} else {'background'});
  log_path=(Join-Path $run 'events.jsonl'); cancel_path=(Join-Path $run 'cancel');
  deadline=[DateTime]::Parse($request.deadline).ToUniversalTime(); documents=@{}; visio_pid=$null
}
$ctx = $script:VisioContext
$state = [ordered]@{ run_id=$request.run_id; status='starting'; mode=$ctx.mode; owned=$false;
  worker_pid=$PID; worker_start=(Get-Process -Id $PID).StartTime.ToUniversalTime().ToString('o');
  visio_pid=$null; visio_start=$null; hwnd=$null; ever_visible=$false;
  baseline=@(Get-Process VISIO -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }); output=$request.output }
$statePath = Join-Path $run 'state.json'
Write-VisioJson $statePath $state
Write-VisioEvent $ctx 'starting' @{ runtime=$PSVersionTable.PSVersion.ToString(); request=$request; ownership=$state }
$visio=$null; $doc=$null; $originalError=$null; $cleanupErrors=@(); $saved=$false; $retained=$false; $pageResults=@(); $exports=@(); $closed=@(); $quit=$false
try {
  while (-not (Test-Path -LiteralPath (Join-Path $run 'go'))) { Assert-VisioRunning; Start-Sleep -Milliseconds 100 }
  Assert-VisioRunning
  # InvisibleApp is a new automation instance, never an arbitrary active application.
  $visio = New-Object -ComObject Visio.InvisibleApp
  $hwnd = [IntPtr]([long]$visio.WindowHandle32)
  [uint32]$nativePid = 0
  [void][VisioSession.Native]::GetWindowThreadProcessId($hwnd, [ref]$nativePid)
  $process = Get-Process -Id $nativePid
  if ($nativePid -le 0 -or $state.baseline -contains [int]$nativePid -or $process.ProcessName -ine 'VISIO') {
    throw 'Cannot establish exclusive Visio process ownership; refusing to modify or quit it.'
  }
  $state.visio_pid=[int]$nativePid; $ctx.visio_pid=[int]$nativePid
  $state.visio_start=$process.StartTime.ToUniversalTime().ToString('o'); $state.hwnd=$hwnd.ToInt64(); $state.owned=$true; $state.status='running'
  Write-VisioJson $statePath $state
  Write-VisioEvent $ctx 'instance_owned' $state
  # Cancel unexpected dialogs rather than accepting destructive defaults.
  $visio.AlertResponse=2
  $visio.EventsEnabled=$false
  if ($request.visible) { $state.ever_visible=$true; Write-VisioJson $statePath $state; $visio.Visible=$true }
  if ($request.action -eq 'probe') {
    Write-VisioEvent $ctx 'probe' @{ version=$visio.Version; caption=$visio.Caption }
  } elseif ($request.action -eq 'draw') {
    if ($request.template) {
      # Open a COPY with macros and workspace restoration disabled. Never modify the source.
      $flags=1+128+256+8
      $doc=Register-VisioDocument ($visio.Documents.OpenEx($request.template, $flags)) 'output'
    } else { $doc=Register-VisioDocument ($visio.Documents.Add('')) 'output' }
    $names=@{}
    $index=0
    foreach ($planPath in $request.plans) {
      $parsed=Get-Content -LiteralPath $planPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $plans=if ($parsed.pages) { @($parsed.pages) } else { @($parsed) }
      foreach ($plan in $plans) {
        Assert-VisioRunning
        $index++
        $name=if ($plan.page.name) { [string]$plan.page.name } else { "Page $index" }
        if ($names.ContainsKey($name)) { throw "Duplicate planned page name: $name" }; $names[$name]=$true
        $page=$null
        for ($i=1; $i -le $doc.Pages.Count; $i++) {
          $candidate=$doc.Pages.Item($i)
          if ($candidate.NameU -eq $name -or $candidate.Name -eq $name) { $page=$candidate; break }
          Release-VisioCom $candidate; $candidate=$null
        }
        if (-not $page) {
          if (-not $request.template -and $index -eq 1) { $page=$doc.Pages.Item(1) } else { $page=$doc.Pages.Add() }
          $page.Name=$name
        }
        Write-VisioEvent $ctx 'draw_page' @{ index=$index; name=$name; plan=$planPath }
        Invoke-VisioPlan -visio $visio -doc $doc -page $page -plan $plan
        $pageResults+=@{ name=$page.Name; shapes=$page.Shapes.Count; connections=$page.Connects.Count }
        Write-VisioEvent $ctx 'page_drawn' $pageResults[-1]
        Release-VisioCom $page; $page=$null; $candidate=$null
      }
    }
    Assert-VisioRunning
    Write-VisioEvent $ctx 'save_started' @{path=$request.output}
    $doc.SaveAs($request.output)
    if (-not [IO.File]::Exists($request.output) -or -not $doc.Saved) { throw 'Save did not produce a saved output document.' }
    $saved=$true
    Write-VisioEvent $ctx 'save_succeeded' @{path=$request.output; pages=$doc.Pages.Count}
  } elseif ($request.action -eq 'export') {
    $doc=Register-VisioDocument ($visio.Documents.OpenEx($request.input, (2+128+256+8))) 'readonly'
  } else { throw 'Unknown Visio job action.' }

  if ($request.preview) {
    Assert-VisioRunning
    if ($request.action -eq 'export') {
      $page=if ($request.page_name) { $doc.Pages.ItemU($request.page_name) } else { $doc.Pages.Item(1) }
      Write-VisioEvent $ctx 'export_started' @{path=$request.preview; page=$page.Name}
      $page.Export($request.preview)
      if (-not [IO.File]::Exists($request.preview)) { throw 'Export did not produce a file.' }
      $exports+=@{page=$page.Name;path=$request.preview}
      Release-VisioCom $page; $page=$null
    } else {
      # Explicit preview directory: all pages, named by index, without activating any window.
      [void][IO.Directory]::CreateDirectory($request.preview)
      for ($i=1; $i -le $doc.Pages.Count; $i++) {
        Assert-VisioRunning
        $page=$doc.Pages.Item($i)
        $path=Join-Path $request.preview ('page-{0:D2}.png' -f $i)
        if (Test-Path -LiteralPath $path) { throw "Export target already exists: $path" }
        Write-VisioEvent $ctx 'export_started' @{path=$path; page=$page.Name}
        $page.Export($path)
        if (-not [IO.File]::Exists($path)) { throw 'Export did not produce a file.' }
        $exports+=@{page=$page.Name;path=$path}
        Release-VisioCom $page; $page=$null
      }
    }
    Write-VisioEvent $ctx 'export_succeeded' @{outputs=$exports}
  }
  Assert-VisioRunning
  if ($request.keep_open -and $doc -and ($saved -or $request.action -eq 'export')) {
    # Keep only the final document, after all requested work succeeds.
    $retained=$true
  }
} catch {
  $originalError=Get-VisioError $_
  Write-VisioEvent $ctx 'original_error' $originalError
  if ($doc -and -not $saved -and $request.action -eq 'draw') {
    try {
      $failurePath=Join-Path $run 'failed-draft.vsdx'
      $doc.SaveAs($failurePath)
      Write-VisioEvent $ctx 'failure_draft_saved' @{path=$failurePath; complete=$false}
    } catch { Write-VisioEvent $ctx 'failure_draft_save_error' (Get-VisioError $_) }
  }
} finally {
  if ($visio -and $state.owned) {
    # An unregistered document could have been opened by the user in visible mode.
    # Never close it or quit its application. Never set Saved on an arbitrary document.
    $unknown=@()
    try {
      for ($i=1; $i -le $visio.Documents.Count; $i++) {
        $candidate=$visio.Documents.Item($i)
        if (-not $ctx.documents.ContainsKey([string]$candidate.ID)) { $unknown+=[string]$candidate.FullName }
        $candidate=$null
      }
    } catch { $cleanupErrors+=Get-VisioError $_; $unknown+= 'document inventory failed' }
    foreach ($key in @($ctx.documents.Keys)) {
      $entry=$ctx.documents[$key]; $ownedDoc=$entry.object
      if ($retained -and $doc -and $ownedDoc.ID -eq $doc.ID) { continue }
      try {
        # Only generated/copy documents or opened read-only stencils are disposable.
        # Original input files and pre-existing user documents are never registered here.
        $ownedDoc.Saved=$true
        $ownedDoc.Close()
        $closed+=@{id=$key;role=$entry.role;closed=$true}
        Write-VisioEvent $ctx 'document_closed' $closed[-1]
      } catch { $cleanupErrors+=Get-VisioError $_; Write-VisioEvent $ctx 'close_error' $cleanupErrors[-1] }
      try { Release-VisioCom $ownedDoc } catch { $cleanupErrors+=Get-VisioError $_ }
      $entry.object=$null; $ownedDoc=$null
    }
    if ($unknown.Count) {
      $cleanupErrors+=@{message='Unowned documents found; application left intact.'; documents=$unknown}
      $retained=$false; $state.status='review_required'; $state.ever_visible=$true
      try { $visio.AlertResponse=0; $visio.EventsEnabled=$true } catch { $cleanupErrors+=Get-VisioError $_ }
    } elseif ($retained -and $cleanupErrors.Count -eq 0) {
      try {
        $state.ever_visible=$true; Write-VisioJson $statePath $state
        $visio.EventsEnabled=$true; $visio.AlertResponse=0; $visio.Visible=$true
        $state.status='retained'
        Write-VisioEvent $ctx 'final_window_retained' @{path=$request.output; pid=$state.visio_pid}
      } catch { $cleanupErrors+=Get-VisioError $_; $retained=$false }
    } else {
      $retained=$false
      try { $visio.Quit(); $quit=$true; Write-VisioEvent $ctx 'quit_called' @{} }
      catch { $cleanupErrors+=Get-VisioError $_; Write-VisioEvent $ctx 'quit_error' $cleanupErrors[-1] }
    }
    if (-not $retained -and $unknown.Count -eq 0 -and -not $quit) {
      try { $visio.Quit(); $quit=$true; Write-VisioEvent $ctx 'quit_after_handoff_error' @{} }
      catch { $cleanupErrors+=Get-VisioError $_ }
    }
    # RCWs created inside the rendering function are now out of scope.
    $doc=$null; $page=$null; $candidate=$null; $ctx.documents.Clear()
    try { Release-VisioCom $visio } catch { $cleanupErrors+=Get-VisioError $_ }
    $visio=$null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if (-not $retained -and $unknown.Count -eq 0) {
      $until=[DateTime]::UtcNow.AddSeconds(10)
      while ((Test-VisioProcessIdentity $state.visio_pid $state.visio_start) -and [DateTime]::UtcNow -lt $until) { Start-Sleep -Milliseconds 100 }
      if (Test-VisioProcessIdentity $state.visio_pid $state.visio_start) {
        $cleanupErrors+=@{message='Owned process did not exit; independent watchdog will inspect it.';pid=$state.visio_pid}
      }
    }
  } elseif ($visio) {
    # Uncertain activation ownership must never turn into a Quit of user resources.
    try { Release-VisioCom $visio } catch { $cleanupErrors+=Get-VisioError $_ }; $visio=$null
    $cleanupErrors+=@{message='Ownership could not be established. Manual review required.'}
  }
  $alive=if ($state.visio_pid) { Test-VisioProcessIdentity $state.visio_pid $state.visio_start } else { $false }
  if ($retained -and -not $alive) { $cleanupErrors+=@{message='Requested retained instance is no longer running.'}; $retained=$false }
  $success=($null -eq $originalError -and $cleanupErrors.Count -eq 0 -and ($retained -or -not $alive))
  if (-not $retained) { $state.status=if ($alive) {'orphaned'} elseif ($success) {'completed'} else {'failed'} }
  Write-VisioJson $statePath $state
  $result=@{run_id=$request.run_id;success=$success;status=$state.status;saved=$saved;output=$request.output;pages=$pageResults;exports=$exports;original_error=$originalError;cleanup_errors=$cleanupErrors;closed=$closed;quit_called=$quit;pid=$state.visio_pid;retained=$retained;leftovers=@(if ($alive -and -not $retained) {$state.visio_pid})}
  Write-VisioEvent $ctx 'finished' $result
  Write-VisioJson (Join-Path $run 'result.json') $result
}
if (-not $success) { exit 1 }
