---
name: visio-template-drawing
description: Create editable Microsoft Visio drawings (.vsdx) from PNG references, text specifications, or existing Visio templates. Use when Codex needs background Visio COM automation, VSDX output, template-based diagram recreation, screenshot-to-Visio conversion, or editable scientific/technical diagrams. Save .vsdx by default; export PNG previews only when explicitly requested.
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
- Draw completely in the background by default using a new owned `Visio.InvisibleApp`; close task documents and verify process exit after completion or failure.
- `-Visible` shows the process but still closes it; `-KeepOpen` independently retains only a successfully completed final document. Never use either switch without the corresponding user request.
- Do not use system clipboard, active selections, mouse, keystrokes, or arbitrary existing Visio instances for background work. Explain any operation requiring foreground interaction before starting it.

## Workflow

1. Confirm inputs and output path. Run `scripts/visio_probe.ps1` if Visio COM availability is uncertain.
2. Analyze the reference semantically: diagram type, layout, labels, shape families, connectors, grouping, colors, images, and any salient geometry.
3. Choose Visio-native stencils/masters before primitives. Start from `references/master-catalog.md`; if the needed master is missing, use local master discovery rather than guessing.
4. Build a drawing plan using `references/drawing-plan-schema.md`.
5. For non-trivial diagrams, run `scripts/validate_plan_layout.py` before drawing. Treat unintentional text/image overlap, unintentional image/image overlap, and page overflow as plan errors that must be fixed before Visio automation. If the reference clearly uses overlays or stacked images, mark the intended overlap explicitly in the plan.
6. Draft and iterate internally first. Do not show rough intermediate versions to the user.
7. Use the managed `visio_draw_from_plan.ps1` entry point for single or multi-page plans. Draw directly into one document; do not invent clipboard-based merge scripts. Read `references/session-lifecycle.md` for modes, timeout/cancellation, logs, and recovery. Save and optionally export in that same session.
8. Check `result.json` plus the watchdog recovery result. A saved file alone is not success: verify structure, text, connections, and cleanup. Report the artifact, run-log path, fallbacks, and any unresolved process or file lock.

## Drawing Principles

- Think like a skilled Visio user, not a geometry renderer: classify the diagram's semantic family first, then prefer template shapes and Visio stencil masters through explicit COM references before primitives or custom geometry. Use raster images only for photo-like content or user-approved non-editable elements.
- Do not hand-draw a known Visio shape family. Search built-in stencils and drop masters by `NameU`; run a small proof when a newly discovered master will be reused.
- Use native dynamic connectors with glued endpoints for semantic connections. Use straight lines only when the reference or user explicitly calls for straight geometry.
- For converging or bundled connectors, prioritize readable topology first; prefer smooth curves only when they do not reduce clarity.
- Route connectors as first-class layout objects. Connector paths must not run through independent labels or non-endpoint node bodies; reserve gaps around branch labels, formulas, and filled nodes. Preserve the reference line style first; if a line is visibly straight, prefer nudging nearby nodes/labels slightly over converting it into an angular detour.
- For freehand-looking curves, use editable COM spline, Bezier, or polygon geometry in background mode and state any approximation. Exact use of the native "Freeform/任意多边形(F)" UI tool requires prior explanation and explicit interactive authorization; never silently switch to foreground input.
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
- no PNG preview was generated unless requested;
- the owned PID exited (or the final document was explicitly retained), original errors are preserved, and unresolved cleanup is reported.

## Resources

- Use `scripts/visio_draw_from_plan.ps1` to draw from one or more JSON plans; an atlas may use a `pages` array.
- Read `references/session-lifecycle.md` before changing ownership, cleanup, cancellation, or recovery; use `scripts/visio_recover.ps1` to inspect a prior run.
- COM workers always use Windows PowerShell 5.1 STA, including when launched from PowerShell 7. Do not reuse `Document.Activate`, `doc.Windows`, or `Marshal.GetActiveObject` recovery snippets.
- Use `scripts/validate_plan_layout.py` before drawing non-trivial screenshot-derived plans.
- Use `scripts/visio_export_preview.ps1` only when PNG preview output is requested.
- Read `references/master-catalog.md` before selecting or discovering Visio masters.
- Read `references/drawing-plan-schema.md` for plan fields, including `richText`.
- Read `references/visio-com-notes.md` before changing automation, master lookup, connector behavior, or character formatting.
