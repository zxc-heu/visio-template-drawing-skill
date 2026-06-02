---
name: Visio-template-drawing
description: Create editable Microsoft Visio drawings (.vsdx) from PNG references, text specifications, or existing Visio templates. Use when Codex needs visible Visio COM automation, VSDX output, template-based diagram recreation, screenshot-to-Visio conversion, or editable scientific/technical diagrams. Save .vsdx by default; export PNG previews only when explicitly requested.
---

# Visio Template Drawing

## Goal

Create real, editable `.vsdx` diagrams with Microsoft Visio COM. Use editable Visio shapes, connectors, and text; do not satisfy the task by pasting a reference PNG as the final diagram.

## Defaults

- Output format: `.vsdx`.
- Local input file: save beside the input as `<source-file-stem>_codex.vsdx`.
- Pasted chat image with no local path: save to the desktop with a descriptive `_codex.vsdx` name.
- User-specified output path or filename overrides defaults.
- Do not export PNG unless the user explicitly asks for a preview, screenshot, or visual QA.
- Keep Visio visible for the final drawing pass unless the user asks for background execution.

## Workflow

1. Confirm inputs and output path. Run `scripts/visio_probe.ps1` if Visio COM availability is uncertain.
2. Analyze the reference semantically: diagram type, layout, labels, shape families, connectors, grouping, colors, images, and any salient geometry.
3. Choose Visio-native stencils/masters before primitives. Start from `references/master-catalog.md`; if the needed master is missing, use local master discovery rather than guessing.
4. Build a drawing plan using `references/drawing-plan-schema.md`.
5. For non-trivial diagrams, run `scripts/validate_plan_layout.py` before drawing. Treat unintentional text/image overlap, unintentional image/image overlap, and page overflow as plan errors that must be fixed before Visio automation. If the reference clearly uses overlays or stacked images, mark the intended overlap explicitly in the plan.
6. Draft and iterate internally first. Do not show rough intermediate versions to the user.
7. Run the final Visio COM automation visibly, draw editable shapes/text/connectors, and save the final `.vsdx`.
8. Report the `.vsdx` path and any meaningful fallback, uncertainty, or manual cleanup area.

## Drawing Principles

- Think like a skilled Visio user, not a geometry renderer: classify the diagram's semantic family first, then prefer visible Visio UI tools, the left-side Shapes pane, template shapes, and Visio stencil masters before primitives or custom geometry. Use raster images only for photo-like content or user-approved non-editable elements.
- Do not hand-draw a known Visio shape family. Search built-in stencils and drop masters by `NameU`; run a small proof when a newly discovered master will be reused.
- Use native dynamic connectors with glued endpoints for semantic connections. Use straight lines only when the reference or user explicitly calls for straight geometry.
- For converging or bundled connectors, prioritize readable topology first; prefer smooth curves only when they do not reduce clarity.
- Route connectors as first-class layout objects. Connector paths must not run through independent labels or non-endpoint node bodies; reserve gaps around branch labels, formulas, and filled nodes. Preserve the reference line style first; if a line is visibly straight, prefer nudging nearby nodes/labels slightly over converting it into an angular detour.
- When the reference contains freehand-looking curves or irregular curved outlines, prefer Visio's Tools menu "Freeform/任意多边形(F)" UI tool as the first-class strategy. Use spline, closed-polyline, or ribbon fallbacks only when the native UI tool cannot be driven reliably in batch, and report that fallback.
- Distinguish Visio UI tools from COM fallbacks. In particular, closed `polygon`, `ribbon`, `DrawPolyline`, and `DrawSpline` outputs are not the same as the Tools menu "Freeform/任意多边形(F)" command.
- Preserve salient visual logic from the reference when it affects meaning or recognition, including connector style, perspective, grouping, repeated shape families, and arrow direction.
- Allocate layout regions before placing details. Photographic images, their labels, legend color chips, connector fan-in/fan-out corridors, and module blocks each need reserved space by default. A label, chip, or image may overlap another object only when the reference clearly uses an overlay/stacked design or the drawing plan explicitly marks the overlap as intentional.
- Use text for label symbols and formulas; do not draw label bars, dashes, formula operators, or mathematical marks as geometry.
- Use SimSun/宋体 for Chinese text and Times New Roman for English, numbers, and mathematical notation by default.
- For MathType-like labels, use `richText` spans so one Visio text box can contain character-level font, italic, subscript, and superscript formatting.
- If a known flowchart or diagram shape is visible in a screenshot, search template/stencil masters before approximating it with primitives.

## Verification

Before final delivery, check that:

- main diagram elements are editable Visio objects, not a pasted screenshot;
- labels and formulas remain editable text;
- connectors and repeated objects preserve the reference's visual logic where practical;
- layout preflight passes for non-trivial plans, especially no unintended label-on-image collisions and no page overflow;
- output path follows the defaults or the user's explicit instruction;
- no PNG preview was generated unless requested.

## Resources

- Use `scripts/visio_draw_from_plan.ps1` to draw from a JSON plan.
- Use `scripts/validate_plan_layout.py` before drawing non-trivial screenshot-derived plans.
- Use `scripts/visio_export_preview.ps1` only when PNG preview output is requested.
- Read `references/master-catalog.md` before selecting or discovering Visio masters.
- Read `references/drawing-plan-schema.md` for plan fields, including `richText`.
- Read `references/visio-com-notes.md` before changing automation, master lookup, connector behavior, or character formatting.
