# Visio Template Drawing Skill

This repository contains a Codex skill for creating editable Microsoft Visio drawings (`.vsdx`) from PNG references, text specifications, or existing Visio templates.

The skill is designed to automate real Visio drawing work through Microsoft Visio COM on Windows. It should create editable Visio shapes, connectors, and text instead of using a pasted screenshot as the final diagram.

## What It Does

- Converts diagram descriptions or reference images into editable `.vsdx` files.
- Uses visible Microsoft Visio COM automation for the final drawing pass.
- Prefers Visio-native stencils, masters, connectors, shapes, and editable text.
- Supports layout validation before drawing non-trivial diagrams.
- Supports rich text spans for mathematical labels, including font, italic, subscript, and superscript formatting.
- Exports PNG previews only when explicitly requested.

## Requirements

- Windows.
- Microsoft Visio installed locally.
- PowerShell.
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
  -TemplateVsdx path\to\template.vsdx `
  -OutputVsdx path\to\output.vsdx `
  -Visible
```

## Publishing Boundary

This repository intentionally keeps local test outputs, one-off debugging scripts, generated previews, and conversation notes out of version control. The source of truth for public use is the `skill/visio-template-drawing/` folder plus this README.

## License

MIT. See `LICENSE`.
