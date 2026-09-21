# Visio Template Drawing Skill

This repository contains a Codex skill for creating editable Microsoft Visio drawings (`.vsdx`) from PNG references, text specifications, or existing Visio templates.

The skill is designed to automate real Visio drawing work through Microsoft Visio COM on Windows. It should create editable Visio shapes, connectors, and text instead of using a pasted screenshot as the final diagram.

<p align="center">
  <img src="docs/assets/visio-template-drawing-skill-flow.png" width="760" alt="Visio template drawing skill workflow illustration">
</p>

## Examples

| Reference | Codex-generated editable Visio output |
| --- | --- |
| <img src="docs/assets/examples/example-01-reference.png" width="360" alt="Example 1 reference diagram"> | <img src="docs/assets/examples/example-01-codex.png" width="360" alt="Example 1 Codex-generated Visio output"> |
| <img src="docs/assets/examples/example-02-reference.png" width="300" alt="Example 2 reference diagram"> | <img src="docs/assets/examples/example-02-codex.png" width="300" alt="Example 2 Codex-generated Visio output"> |
| <img src="docs/assets/examples/example-03-reference.png" width="360" alt="Example 3 reference diagram"> | <img src="docs/assets/examples/example-03-codex.png" width="360" alt="Example 3 Codex-generated Visio output"> |
| <img src="docs/assets/examples/example-04-reference.png" width="330" alt="Example 4 reference diagram"> | <img src="docs/assets/examples/example-04-codex.png" width="330" alt="Example 4 Codex-generated Visio output"> |
| <img src="docs/assets/examples/example-05-reference.png" width="220" alt="Example 5 reference diagram"> | <img src="docs/assets/examples/example-05-codex.png" width="220" alt="Example 5 Codex-generated Visio output"> |

## What It Does

- Converts diagram descriptions or reference images into editable `.vsdx` files.
- Draws in a new hidden Visio COM instance by default; closes task-owned documents and verifies process exit.
- Separates process visibility (`-Visible`) from keeping the final window (`-KeepOpen`).
- Builds multi-page atlases directly from plans without using the system clipboard.
- Records ownership, original errors, save/export results and cleanup, with an independent cancellation/timeout watchdog.
- Prefers Visio-native stencils, masters, connectors, shapes, and editable text.
- Supports layout validation before drawing non-trivial diagrams.
- Supports rich text spans for mathematical labels, including font, italic, subscript, and superscript formatting.
- Exports PNG previews only when explicitly requested.

## Requirements

- Windows.
- Microsoft Visio installed locally.
- Windows PowerShell 5.1 for STA COM workers. Launchers also accept PowerShell 7.
- Python 3 for layout validation and crop helper scripts.
- Pillow is required only when using `crop_reference_regions.py`.

## Skill Location

The publishable skill package is:

```text
skill/visio-template-drawing/
```

Key files:

```text
skill/visio-template-drawing/SKILL.md
skill/visio-template-drawing/agents/openai.yaml
skill/visio-template-drawing/scripts/
skill/visio-template-drawing/references/
skill/visio-template-drawing/examples/
```

## Installation

Copy the skill folder into your Codex skills directory:

```powershell
Copy-Item -Path "skill\visio-template-drawing" -Destination "$env:USERPROFILE\.codex\skills\visio-template-drawing" -Recurse
```

If a previous copy already exists, overwrite individual files carefully. Avoid deleting the whole installed skill folder unless you have reviewed what will be removed.

## Quick Checks

Check whether Visio COM is available:

```powershell
powershell -ExecutionPolicy Bypass -File skill\visio-template-drawing\scripts\visio_probe.ps1
```

Validate a drawing plan:

```powershell
python skill\visio-template-drawing\scripts\validate_plan_layout.py skill\visio-template-drawing\examples\case-06-polygon-spline.json
```

Draw from a plan:

```powershell
powershell -ExecutionPolicy Bypass -File skill\visio-template-drawing\scripts\visio_draw_from_plan.ps1 `
  -PlanJson skill\visio-template-drawing\examples\case-06-polygon-spline.json `
  -OutputVsdx path\to\output.vsdx
```

A template is optional (`-TemplateVsdx path\to\template.vsdx`). Existing output files are refused; choose a fresh name. Use `case-09-eight-pages.json` for the bundled eight-page atlas, or pass an array of single-page plans from a PowerShell session. Plan assembly creates native objects directly in one document; arbitrary VSDX page import and cross-page connectors are not supported.

| Options | Process display | Completion |
| --- | --- | --- |
| none | Hidden | Close owned resources |
| `-Visible` | Visible | Close owned resources |
| `-KeepOpen` | Hidden while drawing | Show and keep final document |
| `-Visible -KeepOpen` | Visible | Keep final document |

Add `-PreviewDirectory path\to\previews` only when PNG exports are wanted. Use `-RunDirectory path\to\new-run` to choose the evidence directory and `-TimeoutSeconds 300` to set a deadline. The default evidence location is `.visio-runs/<id>` beside the output. Logs contain local paths and must stay private.

Read [session lifecycle and recovery](skill/visio-template-drawing/references/session-lifecycle.md) before changing process ownership or recovering a crashed run. Recovery targets only a verified, abandoned, never-visible task process. A hard-crashed visible instance or an unidentified COM activation requires review; no blanket `VISIO.EXE` termination is used.

## Acceptance tests

```powershell
powershell -NoProfile -STA -File skill\visio-template-drawing\scripts\test_visio_lifecycle.ps1 `
  -OutputDirectory path\to\new-acceptance -IncludeVisible
powershell -NoProfile -STA -File skill\visio-template-drawing\scripts\test_visio_interruptions.ps1 `
  -OutputDirectory path\to\new-interruption-tests
```

The first suite deliberately opens controlled windows when `-IncludeVisible` is supplied. Test outputs and logs remain local. Real concurrent mouse/keyboard/clipboard use requires a separate human test; static checks cannot replace it.

## Release notes — 2026-09-21

- Replaced visibility-dependent cleanup with shared ownership, document registration, original-error logging, COM release, and exit verification for drawing, atlases, previews and probing.
- Added clipboard-free multi-page plans, optional templates, independent visibility/retention, cancellation, timeout monitoring and narrow orphan recovery.
- Verified on Visio 16 with Windows PowerShell 5.1 STA; the PowerShell 7.6 launcher also passed. Tests exercised two rounds of single/eight-page drawing, native editing and glue, actual save/export errors, controlled unsaved-document isolation, both visibility modes, locked input, cancellation, worker/supervisor termination and timeout. Concurrent operation in another application was also tested with user confirmation.
- Limitations: native PS7 COM execution, arbitrary existing-VSDX merging, cross-page links, machine power loss, and every third-party add-in/modal failure have not been validated. Never infer successful cleanup solely from a saved file.

## Publishing Boundary

This repository intentionally keeps local test outputs, one-off debugging scripts, generated previews, and conversation notes out of version control. The source of truth for public use is the `skill/visio-template-drawing/` folder plus this README.

## License

MIT. See `LICENSE`.
