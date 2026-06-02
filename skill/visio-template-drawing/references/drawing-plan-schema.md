# Drawing Plan Schema

Use this JSON shape for `scripts/visio_draw_from_plan.ps1`.

```json
{
  "page": {
    "name": "Generated Diagram",
    "width": 11,
    "height": 8.5,
    "backgroundPage": null
  },
  "shapes": [
    {
      "id": "start",
      "shapeIntent": "process-box",
      "preferredSource": "stencil-or-primitive",
      "textRole": "label",
      "kind": "rectangle",
      "master": null,
      "imagePath": null,
      "x": 1.5,
      "y": 7.2,
      "width": 1.8,
      "height": 0.7,
      "text": "Start",
      "fill": "#EAF3FF",
      "line": "#1F4E79",
      "fontFamily": null,
      "fontSize": 10,
      "fontColor": "#000000",
      "bold": false,
      "allowOverlap": false,
      "allowedOverlaps": [],
      "italic": false,
      "mathText": false,
      "richText": [
        { "text": "G", "fontFamily": "Times New Roman", "italic": true },
        { "text": "W", "fontFamily": "Times New Roman", "italic": true, "subscript": true },
        { "text": "(x", "fontFamily": "Times New Roman", "italic": true },
        { "text": "1", "fontFamily": "Times New Roman", "subscript": true },
        { "text": ")", "fontFamily": "Times New Roman" }
      ],
      "rounded": false
    }
  ],
  "connectors": [
    {
      "id": "c1",
      "from": "start",
      "to": "process",
      "text": "Next",
      "line": "#404040",
      "arrow": true,
      "kind": "dynamic",
      "points": null
    }
  ]
}
```

## Page

- `name`: Page name to create or reuse.
- `width`, `height`: Page dimensions in inches. Omit to keep the template page size.
- `backgroundPage`: Reserved for future template background selection. Usually `null`.

## Defaults

- `defaults.connectorZOrder`: Optional default connector stacking directive for the whole plan. Use `back` for dense network diagrams where curved connector bundles should sit behind modules and feature blocks. Individual connectors can override with their own `zOrder`.

## Shapes

- `id`: Stable unique ID used by connectors.
- `shapeIntent`: Optional semantic purpose, such as `process-box`, `formula-box`, `module`, `legend-item`, `icon`, `caption`, or `container`.
- `preferredSource`: Optional source preference: `template`, `stencil`, `primitive`, `custom-lines`, or `raster`.
- `textRole`: Optional text type: `label`, `formula`, `caption`, or `annotation`.
- `kind`: One of `rectangle`, `rounded-rectangle`, `ellipse`, `diamond`, `polygon`, `text`, `image`, `image-placeholder`, or `cuboid`.
- `stencil`: Optional local Visio stencil filename/path/regex to open before master lookup, such as `BASFLO_M.VSSX`, `BPMN_M.VSSX`, `BLOCK_M.VSSX`, or `BASIC_M.VSSX`.
- `master`: Optional Visio master name. If provided and found, the script drops the master instead of drawing a primitive.
- `masterQuery`: Optional regex for local master discovery when `master` is unknown or not already open.
- `preferredStencilRegex`: Optional regex to constrain discovery to relevant stencil filenames, such as `BPMN`, `BASFLO`, `BLOCK`, `UML`, `VSM`, `AWS`, or `AZURE`.
- `imagePath`: Required for `kind: "image"`. Use a local cropped image path. The script imports the image into Visio and sizes it as an editable image object.
- `x`, `y`: Center point in page inches.
- `width`, `height`: Shape size in inches.
- `points`: For `kind: "polygon"`, an array of page coordinates `[x1, y1, x2, y2, ...]`. The script creates a closed polyline-style filled shape and closes the polygon automatically if the first and last points differ. This is **not** the same as Visio's ribbon Tools menu "Freeform/任意多边形(F)" command.
- `text`: Shape text.
- `fill`, `line`, `fontColor`: Hex colors such as `#4472C4`.
- `lineWeight`: Optional line weight in points for emphasized borders and connector trunks.
- `fontFamily`: Optional font family. Omit to use the script default: SimSun for Chinese labels and Times New Roman for non-Chinese labels.
- `fontSize`: Font size in points.
- `bold`: Optional boolean for whole-shape bold text. Use for emphasized labels visible in the reference, especially Chinese flowchart nodes.
- `allowOverlap`: Optional boolean. Use only when overlap is intentionally part of the reference design, such as an overlaid caption, stacked image collage, watermark, or label printed directly on an image.
- `allowedOverlaps`: Optional array of shape IDs or connector IDs allowed to overlap this shape. Use this instead of broad `allowOverlap` when only one known image/text pair or intentional connector pass-through should overlap.
- `italic`: Optional boolean for whole-shape italic text.
- `mathText`: Optional boolean. Use `true` when the label is a formula or formula-like variable expression; preserve italic variables and subscript/superscript formatting where possible.
- `richText`: Optional array for span-level formatting inside one Visio text box. When present, the script concatenates span `text` values and applies character-level `fontFamily`, `fontSize`, `italic`, `subscript`, and `superscript` formatting with Visio `Characters.CharProps`.
- `rounded`: Optional boolean for rectangle corner styling.

Prefer one semantic shape per visible object. A boxed label should be a rectangle/frame shape with text, not four independent lines plus a separate text box. Put mathematical and typographic marks that belong to labels directly in `text`, such as `|| G_W(x_1)-G_W(x_2) ||`, `|`, `->`, `-`, or em dashes, instead of drawing those marks as separate lines.

## Layout Preflight

Before running Visio automation for a non-trivial screenshot-derived plan, run `scripts/validate_plan_layout.py <plan.json>`. Treat these findings as plan errors, not cosmetic issues, unless the reference clearly uses an overlay/stacked design and the plan marks it with `allowOverlap` or `allowedOverlaps`:

- text or label boxes overlapping image boxes;
- image boxes overlapping each other;
- shapes extending beyond the page boundary;
- connector paths crossing unrelated node bodies or labels;
- labels that sit inside an image's visual area when the reference uses external captions.

Plan labels as anchored objects with their own clear area unless the reference intentionally overlays text. For image lists, place each image, its color chip, and its label in a small row or column layout with reserved spacing. Do not let captions drift into neighboring images when the source figure is dense or the page is scaled down.

For connector-heavy figures, reserve routing corridors before placing labels. If a connector fan-in or fan-out passes through the same region as an image label, move the label outside the corridor or adjust the corridor; do not rely on z-order to hide the conflict.

For flowcharts and layered architecture diagrams, keep container titles inside the container with visible padding from the border. Do not center a title box so its left/top edge crosses the container stroke. Place branch labels such as `是`, `否`, `Yes`, and `No` beside the connector path rather than centered directly on the line, unless the reference explicitly places the label on an intentional break in the line.

For network and clustering diagrams, draw connectors to node boundaries rather than node centers when the line is visible above the node. A connector may enter its own source/target node only at the intended attachment point; it must not pass through unrelated node bodies, independent labels, or formula labels. Use explicit `points` to terminate lines near circle/box boundaries when a center-to-center connection would leave visible strokes inside filled nodes.

When a reference shows a straight dashed or solid relation line, keep it straight if a small node/label adjustment can clear the route. Do not automatically replace a straight visual relation with a multi-bend polyline just to satisfy collision checks; first try nudging the nearby point, label, or endpoint while preserving the visual intent.

For invisible merge/route anchors used only to control connector geometry, set `shapeIntent` to `route-helper` or `connector-helper`. These helpers are not semantic nodes and should not be treated as obstacles during layout preflight.

Use `preferredSource` to make the intended drawing strategy explicit. For a user-provided `.vsdx`, mark reused-looking elements as `template` when a matching master or example shape should be searched first. For screenshot-only input, use `stencil` or `primitive` for approximate editable replacements unless exact reproduction is required.

When a shape has known semantics, prefer `stencil` + `master` or `masterQuery` over primitive `kind` fallback. For example, a feature-map cuboid should use a block master such as `3-D box` when available; a flowchart process should use `BASFLO_M.VSSX` `Process`; a BPMN gateway should use a BPMN `Gateway` master.

For formula-like labels, do not flatten typography into plain default text. Use `richText` spans so variables are italic and subscripts/superscripts are real character formatting within one Visio text box. Use Unicode fallback characters such as `x₁`, `x₂` only if rich text formatting fails, and report the limitation.

For known flowchart shapes, set `master` to the intended shape name when possible, such as `Document`, `Process`, `Decision`, or their localized equivalents. A document/sample block with a wavy edge should not be approximated as a plain rectangle unless the master cannot be found.

Use `kind: "cuboid"` for perspective feature bars, tensors, and 3D-looking blocks. The script first opens common Visio stencils such as Basic Shapes and Blocks, then tries masters such as `3-D box`, `Cube`, `Horizontal bar`, and `Square block`. It draws editable front/side fallback faces only when no suitable master is available.

Use `kind: "polygon"` for slanted regions, Voronoi-like areas, wedge backgrounds, and other filled areas that cannot be represented by rectangles. Provide explicit `points` in page coordinates and place polygons before foreground nodes/text in the shape list so later objects appear on top. Treat this as an editable closed-polyline/fill fallback; do not describe it as using Visio's native "任意多边形(F)" UI tool.

When source images are embedded in a screenshot, crop the relevant regions into local image assets and use `kind: "image"` rather than replacing them with color placeholders.

## Connectors

- `id`: Stable unique connector ID.
- `from`, `to`: Shape IDs from the `shapes` array.
- `text`: Optional connector label.
- `line`: Optional hex line color.
- `lineWeight`: Optional line weight in points.
- `arrow`: `true` for an arrow at the target end.
- `fontFamily`: Optional connector label font family.
- `kind`: Optional connector geometry: `dynamic`, `straight`, `polyline`, `curve`, `bezier`, or `ribbon`.
- `points`: Optional array for `polyline` or `curve`, expressed as page coordinates: `[x1, y1, x2, y2, ...]`.
- `zOrder`: Optional connector stacking directive: `back` sends the connector behind later foreground shapes; `front` brings it forward. Use this when bundles should pass behind modules, bars, or cards.
- `fromX`, `fromY`, `toX`, `toY`: Optional relative connection positions for native dynamic connectors. Defaults are right-middle of the source (`1.0, 0.5`) to left-middle of the target (`0.0, 0.5`). Dynamic connectors use Visio `GlueToPos` where possible.

Use `kind: "curve"` when the reference visibly uses curved connectors. Prefer explicit `points` when connector curvature matters. The script uses Visio `DrawSpline` to create an editable spline; if that COM call fails, it reports the fallback and creates an editable rounded polyline.

Use connector `kind: "bezier"` for hand-drawn smooth curves that must avoid angular bends and preserve a deliberate curve contour. Provide Bezier control points in `points` and optional `degree` (`3` by default). Prefer degree-3 Bezier for manually controlled curve shapes; use `curve`/spline for sampled paths that should pass through several points. Use `polyline` only when the reference is deliberately angular.

Use connector `kind: "ribbon"` when screenshot replication needs more visual control than a spline line can provide. Provide either a polygon outline in `points`, or provide centerline points plus `width` to let the script generate a filled closed-polyline ribbon. The script closes the polygon automatically and fills it with `fill` or, if omitted, `line`. This is useful for thick curves, curved arrow bodies, and precisely arranged connector bundles. Add a separate small polygon arrowhead when the reference needs a sharp filled arrow tip. This is also **not** Visio's native "任意多边形(F)" UI tool.

## Native Freeform Tool

The Visio Tools menu "Freeform/任意多边形(F)" is a UI drawing command, not the same as `DrawPolyline`, `DrawSpline`, or `ribbon`. In the Visio command enumeration it corresponds to a UI command such as `visCmdDrawRegion`; it can be activated with `Application.DoCmd`, but the point-by-point drawing interaction is not exposed as the same deterministic batch API used by `visio_draw_from_plan.ps1`.

When the reference contains curves or irregular outlines that a skilled Visio user would draw using the Tools menu, prefer the real Visio "任意多边形(F)" tool as the first strategy. Do not claim that `polygon` or `ribbon` satisfies that requirement. Either:

- run an interactive visible Visio step that activates the tool for manual drawing, or
- report that the batch renderer is using an editable closed-polyline fallback rather than the native UI tool.

For plan-only batch rendering, add a note in the final report when a curve has been rendered with `curve`, `polygon`, or `ribbon` fallback instead of the native UI command.

Keep plans small and explicit. For complex figures, generate multiple pages or group the diagram into phases before drawing.
