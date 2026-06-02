# Visio Master Catalog

This catalog records reusable stencil/master choices verified or expected for the local Visio drawing workflow. Treat it as a starting point, not a complete list.

## Selection Rule

1. Classify the diagram by meaning before choosing shapes: flowchart, BPMN, DFD, UML, network/cloud, org chart, value stream, scientific schematic, screenshot replication, or existing-file edit.
2. Prefer semantic Visio stencil masters over primitive geometry. Use basic shapes only for purely visual elements or when no suitable master exists.
3. Prefer `NameU` in scripts because it is more stable than localized display names.
4. If a needed master is not listed, search installed Visio stencil files with a domain-specific query and stencil regex, then prove the result in a tiny `.vsdx`.
5. Add reusable discoveries here after local verification.

## Generic Visual Replication

Stencil: `BASIC_M.VSSX`

| NameU | Typical use |
| --- | --- |
| Rectangle | Generic boxes and table cells |
| Rounded Rectangle | Function blocks and module boxes |
| Circle | Circular nodes |
| Ellipse | Oval nodes |
| Diamond | Decision or diamond node |
| Can | Database or cylinder |
| Cube | 3D block fallback |
| Left Brace / Right Brace | Grouping brace |
| Left Parenthesis / Right Parenthesis | Parenthesis grouping |

## Blocks And Feature Maps

Stencil: `BLOCK_M.VSSX`, `BLOCK3_M.VSSX`

| NameU | Typical use |
| --- | --- |
| 3-D box | Feature-map bars, cuboids, perspective blocks |
| Layered Box | Stacked block |
| Horizontal bar | 3D horizontal bar alternative |
| Square block | 3D square block |
| Dynamic connector | Native dynamic connector |

## Basic Flowchart

Stencil: `BASFLO_M.VSSX`

| NameU | Typical use |
| --- | --- |
| Process | Process step |
| Decision | Decision branch |
| Start/End | Terminator |
| Document | Document-like shape |
| Data | Input/output |
| Database | Database/storage |
| Dynamic connector | Native flow connector |

## Discovery Examples

Use `masterQuery` and `preferredStencilRegex` in a drawing plan when the catalog does not yet contain the needed shape:

```json
{
  "id": "gateway",
  "kind": "rectangle",
  "stencil": "BPMN_M.VSSX",
  "masterQuery": "^Gateway$",
  "preferredStencilRegex": "BPMN"
}
```

For local PowerShell discovery, search only stencil files first (`*.vssx`, `*.vss`) and constrain by domain (`BPMN`, `BASFLO`, `BLOCK`, `UML`, `VSM`, `AWS`, `AZURE`) to avoid slow broad scans.
