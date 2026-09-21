# Managed Visio sessions

Use the public drawing, export, and probe entry points. They all call the same session launcher, Windows PowerShell 5.1 STA worker, and independent watchdog. Do not dot-source the worker as an ad hoc drawing script.

## Modes and outputs

| Options | While drawing | After successful completion |
| --- | --- | --- |
| none | Hidden `Visio.InvisibleApp` | Close owned documents and verify exit |
| `-Visible` | Visible | Close owned documents and verify exit |
| `-KeepOpen` | Hidden | Show and retain only the final document |
| `-Visible -KeepOpen` | Visible | Retain only the final document |

Failures always request cleanup, including with `-KeepOpen`. Do not edit the generated document while a visible job is running: it is still a task-owned draft. After a successful keep-open handoff, the user owns that window; automatic recovery refuses it. A visible mode can change focus; only background mode promises to avoid intentional UI interaction.

`-TemplateVsdx` is optional. A template is opened as a copy, with macros and workspace restoration disabled. Its pages are preserved, and a matching named page is reused. Without a template, the first plan uses the initial page, so an eight-page plan creates exactly eight pages without a blank cover page.

`-PlanJson` accepts an array when invoked from PowerShell, or a JSON atlas containing `pages`. These are complete per-page plans. Assembly redraws plans in one document and preserves native shapes and glued connections; it is not a general importer for arbitrary existing VSDX documents. Cross-page links are not implemented. Never fall back to clipboard Copy/Paste silently.

Existing output paths are refused. Use a new filename. PNG previews remain opt-in through `-PreviewDirectory` (all pages) or the separate export command (one explicit page). Geometry and labels remain native objects; image shapes are used only where the plan intentionally calls for raster content.

## Ownership and cleanup

The worker snapshots existing Visio PIDs before activation, creates `Visio.InvisibleApp`, derives the real Windows PID from `WindowHandle32` with `GetWindowThreadProcessId`, and records PID plus process start time before drawing. Visio's own `Application.ProcessID` is **not** a Windows PID.

All newly created/copy/read-only documents and stencils are registered immediately. Rendering uses explicit application, document, page and shape references. No `ActiveWindow`, `ActivePage`, selection, mouse, keyboard, or system clipboard is used. `ConnectorToolDataObject` is an in-process Visio data object, not a call to the system clipboard.

Cleanup closes only registered task documents. An incomplete draw is saved to the run directory as `failed-draft.vsdx` when possible; it is explicitly not a completed deliverable. A saved final file remains intact if subsequent export fails. Temporary copies are disposable. Setting `Saved=true` before closing is restricted to registered task copies/stencils after this policy is applied; it never changes the original template or a pre-existing user document.

`AlertResponse=2` cancels unexpected dialogs in the owned instance; it does not blindly accept them. Explicit save calls must succeed before a job is marked saved. Final handoff restores alerts and events. If an unregistered document appears in a visible instance, do not close it or quit the instance: report incomplete cleanup for review. Rendering references leave function scope, document/application RCWs are released, remaining scoped RCWs are collected, and process exit is checked. A file being saved is insufficient to declare success.

## Cancellation, deadlines, and hard interruption

`-TimeoutSeconds` defaults to 300. The worker checks cancellation between pages/shapes/connectors and expensive stages. To request cancellation, create a file named `cancel` inside its run directory. The independent watchdog also requests cancellation when the supervisor dies or the deadline passes. After a ten-second grace period it may terminate that exact worker; it never kills processes by name.

The watchdog then inspects the persisted ownership record. It can terminate one abandoned Visio process only if PID, process name, start time, exclusive ownership and hidden mode all match; the worker must have exited, the instance must never have been made visible, and it must currently have no visible top-level window. This may lose unsaved task-only work. Already saved outputs and logs remain. Recovery never targets pre-existing PIDs or an intentionally retained instance.

`finally` cannot cover forced termination of the worker, supervisor, watchdog, or machine. If both supervision processes disappear, inspect the specific run on the next task:

```powershell
powershell -NoProfile -File scripts\visio_recover.ps1 -RunDirectory path\to\run
# Execute only the narrow, verified hidden-instance recovery described above:
powershell -NoProfile -File scripts\visio_recover.ps1 -RunDirectory path\to\run -Recover
```

There is an unavoidable activation gap before a COM instance can be identified and durably recorded. `activation_not_identified` requires review; do not infer ownership from a process-name sweep. Visible, retained, or uncertain instances also require user review after a hard crash, because they could contain new user work. No automatic bulk process termination or arbitrary `GetActiveObject` recovery is allowed. Do not claim every force-stop or Office add-in hang has been tested.

## Logs and runtime

Each invocation writes to a new `-RunDirectory`, or `.visio-runs/<id>` beside the output. Do not reuse a run directory.

- `request.json`, `worker.json`, `state.json`: mode, deadline, inputs/outputs, supervisor/worker ownership, native PID, start time, handle, visibility and disposition.
- `events.jsonl`: run ID, UTC time, mode, stage, ownership, document roles, page/shape/connection counts, save/export results, original exception type/HRESULT/stack, close and quit results.
- `result.json`: success/failure, original error distinct from secondary cleanup errors, artifacts, retained status and leftovers.
- `recovery.json`: independent watchdog disposition; manual recovery writes `recovery.manual.json`.
- stdout/stderr logs: PowerShell diagnostics, including renderer fallback warnings. These logs may contain private file paths and document text; keep them local.

The public launchers work from PowerShell 7 or Windows PowerShell. COM runs in `powershell.exe` 5.1 STA to avoid PowerShell 7 COM numeric-binding differences seen with this workflow. Native PS7 COM execution is not a supported path. Keep scripts containing non-ASCII text in UTF-8 with BOM for PS5.1. Avoid `Document.Activate`, `doc.Windows`, and `Marshal.GetActiveObject`; none is needed here.

## Reproducible acceptance

Run with a new local output directory. The visible switch on the test suite deliberately opens controlled test windows; it never uses the user's real unsaved document.

```powershell
powershell -NoProfile -STA -File scripts\test_visio_lifecycle.ps1 -OutputDirectory path\to\acceptance -IncludeVisible
powershell -NoProfile -STA -File scripts\test_visio_interruptions.ps1 -OutputDirectory path\to\interruptions
```

These exercise single/eight-page output, native text and coordinate editing, endpoint glue, two rounds, locks, actual COM save/export failures, drawing failure, controlled unsaved-document isolation, visibility/retention, timeout, cancellation, worker termination, and supervisor termination. They do not simulate a human's concurrent keyboard/clipboard use. Perform that test with the user in another application and report their observation separately; a focus sample or static search alone does not prove non-interference.

Official API references: [InvisibleApp](https://learn.microsoft.com/en-us/office/vba/api/visio.invisibleapp), [ProcessID caveat](https://learn.microsoft.com/en-us/office/vba/api/visio.application.processid), [OpenEx flags](https://learn.microsoft.com/en-us/office/vba/api/visio.documents.openex), [AlertResponse values](https://learn.microsoft.com/en-us/office/vba/api/visio.application.alertresponse).
