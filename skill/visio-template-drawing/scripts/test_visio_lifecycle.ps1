# Run in Windows PowerShell 5.1 -STA. Uses synthetic drawings only.
param([Parameter(Mandatory=$true)][string]$OutputDirectory, [switch]$IncludeVisible)
$ErrorActionPreference='Stop'
. "$PSScriptRoot\visio_session.ps1"
Initialize-VisioNative
if ($PSVersionTable.PSEdition -ne 'Desktop') { throw 'Run this COM test harness with powershell.exe -STA (Windows PowerShell 5.1).' }
$root=[IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $root) { throw 'Use a new OutputDirectory for each acceptance run.' }
[void][IO.Directory]::CreateDirectory($root)
$records=New-Object Collections.Generic.List[object]
$baseline=@(Get-Process VISIO -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$single=Join-Path $PSScriptRoot '..\examples\case-08-background-flow.json'
$atlas=Join-Path $PSScriptRoot '..\examples\case-09-eight-pages.json'
$draw=Join-Path $PSScriptRoot 'visio_draw_from_plan.ps1'
$export=Join-Path $PSScriptRoot 'visio_export_preview.ps1'
$control=$null; $sentinel=$null; $checkDoc=$null
function Assert-Test { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw $Message } }
function Assert-Control {
  Assert-Test ($sentinel.Pages.Item(1).Shapes.Item(1).Text -eq 'UNSAVED CONTROL') 'Controlled pre-existing text changed.'
  Assert-Test (-not $sentinel.Saved) 'Controlled pre-existing document lost its unsaved state.'
  Assert-Test ($control.Documents.Count -eq 1) 'Controlled document was closed or extra documents were left open.'
}
function Assert-Clean { param($Result)
  Assert-Test (-not (Get-Process -Id $Result.pid -ErrorAction SilentlyContinue)) 'Task process remained alive.'
  Assert-Test (@($Result.cleanup_errors).Count -eq 0) 'Cleanup errors reported.'
  Assert-Control
}
function Assert-Unlocked { param([string]$Path)
  $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None); $stream.Dispose()
}
function Read-Result { param([string]$Name)
  Get-Content -LiteralPath (Join-Path $root "$Name\result.json") -Raw -Encoding UTF8 | ConvertFrom-Json
}
function Record-Pass { param([string]$Name,$Data)
  $records.Add(@{test=$Name;passed=$true;data=$Data})
  Write-VisioJson (Join-Path $root 'acceptance.json') @($records.ToArray())
  Write-Output "PASS $Name"
}
function Check-Drawing { param([string]$Path,[int]$Pages)
  $checkDoc=$null
  try {
    $checkDoc=$control.Documents.OpenEx($Path, (1+128+256+8))
    Assert-Test ($checkDoc.Pages.Count -eq $Pages) 'Page count mismatch.'
    for ($i=1;$i -le $Pages;$i++) {
      $p=$checkDoc.Pages.Item($i)
      Assert-Test ($p.Shapes.Count -eq 3) 'Expected two native shapes and one connector.'
      Assert-Test ($p.Connects.Count -eq 2) 'Both connector endpoints must remain glued.'
      $a=$p.Shapes.ItemU('source'); $b=$p.Shapes.ItemU('target'); $c=$p.Shapes.ItemU('link')
      Assert-Test ($a.Type -eq 3 -and $b.Type -eq 3 -and $c.OneD -ne 0) 'Unexpected non-native object type.'
      Assert-Test ($a.Text -eq "输入 Input" -and $b.Text -eq "输出 Output") 'Editable text mismatch.'
      $a.Text='EDIT VERIFIED'; $a.CellsU('PinX').ResultIU=2.2
      Assert-Test ($a.Text -eq 'EDIT VERIFIED' -and [Math]::Abs($a.CellsU('PinX').ResultIU-2.2) -lt 0.001 -and $p.Connects.Count -eq 2) 'Editing broke text, coordinates, or glued endpoints.'
    }
  } finally { if ($checkDoc) { $checkDoc.Saved=$true; $checkDoc.Close(); Release-VisioCom $checkDoc } }
  Assert-Unlocked $Path
}
try {
  # Controlled analogue of a user's already-open, unsaved document. Never attach to their real work.
  $control=New-Object -ComObject Visio.InvisibleApp
  $control.AlertResponse=2
  $sentinel=$control.Documents.Add('')
  $sentinel.Pages.Item(1).DrawRectangle(1,1,3,2).Text='UNSAVED CONTROL'
  if ($IncludeVisible) { $control.Visible=$true }
  Assert-Control
  foreach ($round in 1..2) {
    foreach ($kind in @('single','atlas')) {
      $name="$kind-$round"; $path=Join-Path $root "$name.vsdx"
      $plan=if ($kind -eq 'single') {$single} else {$atlas}
      $args=@{PlanJson=$plan;OutputVsdx=$path;RunDirectory=(Join-Path $root $name)}
      if ($round -eq 1) { $args.PreviewDirectory=Join-Path $root "$name-previews" }
      & $draw @args | Out-Null
      $r=Read-Result $name; Assert-Test $r.success 'Success run failed.'; Assert-Clean $r
      Check-Drawing $path $(if ($kind -eq 'single') {1} else {8})
      Record-Pass $name @{pid=$r.pid;pages=$r.pages;editable=$true;unlocked=$true}
    }
  }

  $bad=Get-Content -LiteralPath $single -Raw -Encoding UTF8 | ConvertFrom-Json
  $bad.shapes[1].id=''
  $badPath=Join-Path $root 'invalid-mid-draw.json'; Write-VisioJson $badPath $bad
  try { & $draw -PlanJson $badPath -OutputVsdx (Join-Path $root 'mid-draw.vsdx') -RunDirectory (Join-Path $root 'mid-draw') | Out-Null; throw 'Expected draw failure.' } catch { }
  $r=Read-Result 'mid-draw'; Assert-Test (-not $r.success -and $r.original_error.message -match 'must have an id') 'Did not observe injected drawing failure.'; Assert-Clean $r
  Assert-Test (Test-Path -LiteralPath (Join-Path $root 'mid-draw\failed-draft.vsdx')) 'Incomplete failure draft missing.'
  Record-Pass 'mid-draw-failure' $r.original_error

  # A reserved Windows device name reaches Visio SaveAs and fails at the COM save boundary.
  try { & $draw -PlanJson $single -OutputVsdx (Join-Path $root 'CON.vsdx') -RunDirectory (Join-Path $root 'save-failure') | Out-Null } catch { }
  $r=Read-Result 'save-failure'; Assert-Test (-not $r.success -and -not $r.saved -and $r.original_error) 'Did not observe save failure.'; Assert-Clean $r
  Assert-Test ((Get-Content -LiteralPath (Join-Path $root 'save-failure\events.jsonl') -Raw) -match 'save_started') 'Failure happened before the save boundary.'
  Record-Pass 'save-failure' $r.original_error

  # Existing file where a preview directory is expected: drawing was saved before export failed.
  $blocked=Join-Path $root 'blocked-preview'; [IO.File]::WriteAllText($blocked,'not a directory')
  $savedPath=Join-Path $root 'export-failure.vsdx'
  try { & $draw -PlanJson $single -OutputVsdx $savedPath -PreviewDirectory $blocked -RunDirectory (Join-Path $root 'export-failure') | Out-Null } catch { }
  $r=Read-Result 'export-failure'; Assert-Test (-not $r.success -and $r.saved) 'Saved artifact was not retained on export failure.'; Assert-Clean $r; Check-Drawing $savedPath 1
  Record-Pass 'export-path-failure-preserves-saved-output' $r.original_error
  try { & $export -InputVsdx $savedPath -OutputPng (Join-Path $root 'unsupported.invalidformat') -RunDirectory (Join-Path $root 'export-com-failure') | Out-Null } catch { }
  $r=Read-Result 'export-com-failure'; Assert-Test (-not $r.success -and $r.original_error) 'Did not observe COM exporter failure.'; Assert-Clean $r
  Record-Pass 'export-com-failure' $r.original_error
  & $export -InputVsdx (Join-Path $root 'atlas-1.vsdx') -PageName 'Page 8' -OutputPng (Join-Path $root 'page-eight.png') -RunDirectory (Join-Path $root 'export-page-8') | Out-Null
  $r=Read-Result 'export-page-8'; Assert-Clean $r; Record-Pass 'explicit-page-export' $r.exports

  if ($IncludeVisible) {
    & $draw -PlanJson $single -OutputVsdx (Join-Path $root 'visible.vsdx') -Visible -RunDirectory (Join-Path $root 'visible') | Out-Null
    $r=Read-Result 'visible'; Assert-Clean $r; Record-Pass 'visible-closes' @{pid=$r.pid}
    try { & $draw -PlanJson $badPath -OutputVsdx (Join-Path $root 'visible-failure.vsdx') -Visible -KeepOpen -RunDirectory (Join-Path $root 'visible-failure') | Out-Null } catch { }
    $r=Read-Result 'visible-failure'; Assert-Test (-not $r.success -and -not $r.retained) 'Failed visible job was retained.'; Assert-Clean $r
    Record-Pass 'visible-keepopen-still-cleans-on-failure' @{pid=$r.pid;retained=$r.retained}
    $keepPath=Join-Path $root 'retained.vsdx'
    & $draw -PlanJson $single -OutputVsdx $keepPath -KeepOpen -RunDirectory (Join-Path $root 'retained') | Out-Null
    $r=Read-Result 'retained'; Assert-Test ($r.retained -and (Get-Process -Id $r.pid -ErrorAction SilentlyContinue)) 'Final window did not stay open.'
    Assert-Control
    # Bind only to this exact test output; validate its PID before touching it.
    $kept=[Runtime.InteropServices.Marshal]::BindToMoniker($keepPath)
    $keptApp=$kept.Application; [uint32]$keptPid=0
    [void][VisioSession.Native]::GetWindowThreadProcessId([IntPtr]([long]$keptApp.WindowHandle32),[ref]$keptPid)
    Assert-Test ($keptPid -eq $r.pid -and $kept.FullName -ieq $keepPath -and $keptApp.Documents.Count -eq 1) 'Retained document ownership mismatch.'
    $kept.Close(); $keptApp.Quit(); Release-VisioCom $kept; Release-VisioCom $keptApp; $kept=$null;$keptApp=$null
    Record-Pass 'keep-final-window' @{pid=$r.pid;visible=$true;closed_by_test_harness=$true}
  }
  Assert-Control
  Record-Pass 'pre-existing-unsaved-control-unmodified' @{text='UNSAVED CONTROL';saved=$sentinel.Saved;documents=$control.Documents.Count}
} catch {
  $records.Add(@{test='harness';passed=$false;error=(Get-VisioError $_)})
  Write-VisioJson (Join-Path $root 'acceptance.json') @($records.ToArray())
  throw
} finally {
  if ($sentinel) { $sentinel.Saved=$true; $sentinel.Close(); Release-VisioCom $sentinel; $sentinel=$null }
  if ($control) { $control.Quit(); Release-VisioCom $control; $control=$null }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Start-Sleep -Seconds 1
$extra=@(Get-Process VISIO -ErrorAction SilentlyContinue | Where-Object {$baseline -notcontains $_.Id} | ForEach-Object {$_.Id})
Assert-Test ($extra.Count -eq 0) 'Unexpected Visio processes remain after the suite. Inspect ownership records; do not kill by name.'
Record-Pass 'two-rounds-no-process-accumulation' @{baseline=$baseline;extra=$extra}
