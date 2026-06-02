# Visio COM Notes

Visio COM uses page units in inches for most drawing calls. `DrawRectangle(left, bottom, right, top)` uses page coordinates, while the drawing plan stores shape centers for easier layout reasoning.

## Visibility

Set `Visio.Application.Visible = $true` when the user wants to watch the process. Visible automation is slower but easier to trust and debug.

## Templates and Masters

Open the template with `Documents.Open()`, then save to a new output path with `SaveAs()` so the original template is not modified. If a shape plan specifies a `master`, search open documents and stencils for a matching master name. If no master is found, fall back to primitives and mention the fallback.

Prefer masters and stencil shapes the way a human Visio user would use the left-side Shapes pane. Reuse shapes from a user-provided `.vsdx` before drawing primitives. Use primitives as a practical fallback for screenshots or common boxes/arrows. Avoid constructing rectangles, frames, or icon-like shapes from many independent line segments unless the target is truly a custom line drawing.

Open stencils hidden and read-only with `Documents.OpenEx(path, 66)` when possible. This avoids visible stencil windows and save prompts. Prefer `NameU` for master lookup, because it is more stable than localized display names.

If the target shape is not listed in `master-catalog.md`, search installed Visio stencil files by query and preferred stencil regex before falling back. Use a small proof `.vsdx` for newly discovered masters, then record reusable findings in `master-catalog.md`.

## Connectors

Prefer dynamic connectors when available. Glue connector begin/end cells to target shape pin cells when possible. If glue fails, draw a line between shape centers as a fallback and report it.

For semantic connections, use native Visio dynamic connector masters and `GlueToPos` endpoints. Defaults should be right-middle of the source to left-middle of the target. Keep Visio's orthogonal/right-angle routing by default; set straight-routing ShapeSheet cells only when the user or reference explicitly requires straight connectors.

When the reference uses a visible curve, do not replace it with a straight dynamic connector by default. Use the plan connector `kind: "curve"` with explicit `points`. The drawing script should call Visio `Page.DrawSpline()` first so the result is a real editable spline. Only fall back to a rounded editable polyline if `DrawSpline()` fails, and report that fallback.

For curves that need to look hand-shaped and consistent, prefer `DrawBezier()` with degree 3. A cubic Bezier gives better manual control over the curve tangent and avoids the angular appearance of short polyline segments. Use `DrawSpline()` when the curve should be sampled through several visible points. Use `DrawPolyline()` only for intentionally angular routes.

For converging connector bundles, topology readability is the hard priority. A clear angular route is better than a smooth curve bundle whose paths overlap, tangle, or obscure which source connects to which target. If readability is comparable, prefer smooth curves over visibly angular polylines. Practical priority: clear smooth curves > clear polylines > tangled smooth curves > tangled polylines. When smooth curves reduce clarity, use clear polylines, staged merge points, slight endpoint offsets, or layered fan-in routing.

Use `kind: "polyline"` for deliberate broken/orthogonal paths. Use `kind: "straight"` for straight line segments. Use `kind: "dynamic"` when routing is not visually important.

When connector bundles should visually enter modules or pass behind feature blocks, set connector `zOrder` to `back` in the drawing plan. Do not globally send every connector behind all shapes, because diagrams with filled background regions may need connectors above the region fills.

Use shape `kind: "polygon"` for filled slanted regions, wedge backgrounds, and segmented color areas. The script creates a closed polyline-style filled shape and closes open point lists automatically. Draw background polygons before foreground objects so the shape order matches the intended visual stacking. Do not describe this fallback as Visio's native Tools menu "Freeform/任意多边形(F)" command.

## Images From Screenshots

If a screenshot contains photographic or raster subimages that are meaningful to the diagram, crop those regions into local image files and insert them with `kind: "image"`. Do not replace them with same-color blocks unless the user asks for placeholders or the source image cannot be accessed.

For direct chat-pasted screenshots, first determine whether the image is available as a local file in the conversation or workspace. If it is available, crop from that file. If it is not available as a file, explain the limitation and use editable placeholders only as a fallback.

Image-heavy diagrams need explicit layout reservation. Treat every inserted image as an occupied rectangle with a no-label zone by default, but inspect the reference before enforcing this. Some diagrams intentionally place captions, watermarks, callouts, or cropped panels on top of images. Captions such as `Support set` and `Query set` should be placed in external caption bands or reserved whitespace when the reference uses external captions; if the reference intentionally overlays the caption on an image, mark that intent in the plan with `allowOverlap` or `allowedOverlaps`. Small legend color chips should align to the image rows or columns without consuming the caption area unless the source design intentionally stacks them.

Run `scripts/validate_plan_layout.py` on the JSON plan before drawing. If it reports unintentional text-image overlap, image-image overlap, or page overflow, fix the plan coordinates first. If the overlap is intentional, encode that decision in the plan so the validator can distinguish deliberate stacking from accidental collision. Do not use z-order changes as a substitute for correcting accidental layout conflicts; an editable text box hidden above or below an image is still a bad Visio drawing.

## Cuboids and Perspective Blocks

For feature maps, tensors, stacked bars, and perspective blocks, search Visio stencils before drawing geometry. Common useful masters include `3-D box` from Blocks, `Cube` from Basic Shapes, and `Horizontal bar` from 3-D Blocks. Only draw custom editable front/side faces when no suitable stencil master can be opened or dropped. Report that fallback.

Some Visio masters are groups whose visible faces are child shapes. After dropping a master, apply fill and line colors recursively to child shapes; setting only the parent ShapeSheet may leave the visible block uncolored.

For screenshot replication, spline connectors are not always visually stable enough. When a curve needs precise contour, thickness, or arrow-body shape, use a thin filled closed-polyline/ribbon instead of a one-dimensional connector. A ribbon can be supplied as an explicit outline or as a centerline plus width. Keep it editable and use a separate filled polygon for the arrowhead if needed. This is not the same as Visio's native "任意多边形(F)" UI drawing tool.

## Native Freeform Tool

Visio's ribbon Tools menu "Freeform/任意多边形(F)" is a UI command. It can be selected through the Visio command system, but the deterministic COM page methods used for batch rendering (`DrawPolyline`, `DrawSpline`, `DrawBezier`) create path shapes and are not proof that the native UI tool was used. `DrawPolyline` is documented as equivalent to `DrawSpline` with zero tolerance and the abrupt flag, so it must not be treated as the native freeform tool.

If the reference contains freehand-looking curves, irregular curved outlines, thick curved bands, or shapes that a human Visio user would naturally draw from the Tools menu, treat the native "任意多边形(F)" UI tool as the preferred strategy. If exact use of that tool matters, use visible interactive Visio operation and state that it is interactive/manual.

If batch rendering must proceed without reliable UI-tool point entry, use a closed-polyline, spline, or ribbon fallback and explicitly report the fallback. Do not silently substitute COM geometry while claiming the Visio UI tool was used.

The broader principle is to behave like a skilled Visio user: use the UI tool palette and left-side Shapes pane first, then script those actions through COM where practical. Only drop to low-level custom geometry after stencil/master/UI-tool options have been considered.

## Text and Fonts

Set label content as text whenever it is part of the label, including mathematical bars, vertical separators, minus signs, dashes, arrows, and formula notation. Do not draw these text symbols as separate line geometry.

Default to SimSun for labels containing Chinese characters and Times New Roman for English, numbers, and formula-only labels. Use explicit `fontFamily` in the plan when the reference or template requires another font. Visio font support can vary by installation; report any font fallback.

For academic and mathematical labels, inspect italic, subscript, and superscript formatting instead of treating the text as plain characters. Use real Visio text formatting spans when possible. If COM rich-text formatting is unreliable, use Unicode fallback characters such as `x₁` and `x₂`, keep the label editable, and report the approximation.

Mixed Chinese/math labels need mixed typography when possible: Chinese segments in SimSun and English/math segments in Times New Roman with italic variables and proper subscript/superscript formatting.

Use ShapeSheet `FONT("Font Name")` formulas for whole-shape font assignment. Do not rely on `Document.Fonts.ItemU()` because some Visio COM font collections do not expose that method.

Short formula labels such as `g_θ`, `f_θ`, `E_W`, and `G_W(x_1)` should use `richText`, not plain text. Set Latin/math variables in Times New Roman italic and apply subscript formatting to Greek or numeric subscripts.

`crop_reference_regions.py` requires Pillow. In Codex desktop sessions, prefer the bundled Python runtime when available; otherwise ensure Pillow is installed before using the crop helper.

## Flowchart Shape Matching

For screenshot-derived flowchart-like objects, search for Visio masters before approximating. Common targets include Document, Process, Decision, Data, Predefined Process, and Display. A shape with a wavy lower edge should usually be a Document/文档 master. If the target master cannot be loaded or found, use an editable approximation and state that fallback in the final report.

Flowchart labels need collision-aware placement. Keep branch labels offset beside the connector path so the line remains visually continuous and readable. For layered diagrams, place container/layer titles with enough inset padding to avoid touching the container border. When using helper merge points for branch routing, keep horizontal merge lines below source boxes and above target boxes; do not let the route ride along a box border.

Connector routing is part of layout quality, not a per-case tweak. Before drawing, check that connector segments avoid independent text labels and avoid all non-endpoint node bodies. For filled circular nodes, prefer explicit line endpoints on the circle boundary; center-to-center lines drawn above nodes leave visible strokes through the node fill.

Collision repair should preserve the source line style when possible. If the source uses a straight dashed distance line, first adjust adjacent node or label positions slightly and keep a straight connector. Use a bent route only when a straight line cannot stay readable without changing the diagram semantics.

## Export

Use the active page's `Export()` method for PNG previews. Export is page-based, so activate or select the intended page first.

## Reliability Rules

- Never overwrite the template path.
- Read drawing-plan JSON as UTF-8 explicitly. Chinese labels can corrupt or break `ConvertFrom-Json` if PowerShell falls back to the local ANSI code page.
- Keep COM object cleanup explicit at the end of scripts.
- Use editable Visio shapes and text, not pasted screenshots, for final output.
- Report missing masters, unglued connectors, or text that may need manual adjustment.
