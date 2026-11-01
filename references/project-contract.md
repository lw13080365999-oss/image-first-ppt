# Project contract (schema version 1)

Use a separate task directory for each deck. Project manifest and static-audit helpers use only Python's standard library; rendering and authoring have additional requirements listed in [runtime.md](runtime.md). The project manifest is the source of truth; generated PPTX files and images are outputs, not source-content authority.

## Files and commands

```text
task/
  deck.json
  source/index.json
  source/src-001-report.docx
  assets/...
  slides/slide-001/
    design/v001.png
    design/v002.png
    objects.json
    assets/...
    renders/...
  output/deck.pptx
  output/static-audit.json
```

```bash
python scripts/project.py init /absolute/task --title "Title" --source /absolute/report.docx
python scripts/project.py validate /absolute/task
python scripts/project.py validate /absolute/task --require-maps
python scripts/project.py status /absolute/task
python scripts/audit_pptx.py /absolute/task/output/deck.pptx --project /absolute/task --out /absolute/task/output/static-audit.json
```

`init` copies inputs into `source/`, records their SHA-256, and creates an empty `slides` array. It does not infer content. Optional `--slides N` creates empty planning records that require actual titles and content before validation can pass. Initialization refuses to replace an existing `deck.json`. `source/index.json` is the intake snapshot; after intake, update the authoritative `deck.json` source records when adding materials.

Commands print JSON. Validation/audit hard errors produce exit status 1; malformed arguments produce the standard argparse nonzero status. A successful `init` means initialized, not a complete or validated deck.

## `deck.json`

```json
{
  "schema_version": 1,
  "title": "Product review",
  "canvas": {"width_in": 13.333333, "height_in": 7.5},
  "style": {"id": "editorial"},
  "requirements": {
    "native_text": true,
    "editable_tables": true,
    "editable_charts": true,
    "editable_formulas": true
  },
  "sources": [
    {"id": "src-001", "path": "source/report.txt", "kind": "text"}
  ],
  "assets": [
    {"id": "asset-hero", "path": "assets/hero.png", "kind": "image", "source_ids": ["src-001"]}
  ],
  "slides": [
    {
      "id": "slide-001",
      "title": "Growth and limits",
      "source_ids": ["src-001"],
      "texts": [{"id": "summary", "text": "The exact approved wording", "required": true}],
      "formulas": [{"id": "eq-1", "latex": "x^2", "display_text": "x2", "editable": true}],
      "tables": [{"id": "table-1", "cells": [["Item", "Value"], ["A", "2"]], "editable": true}],
      "charts": [{"id": "chart-1", "categories": ["A", "B"], "series": [{"name": "Count", "values": [2, 3]}], "editable": true}],
      "required_assets": ["asset-hero"],
      "notes": "Source context and speaker notes"
    }
  ],
  "outputs": {"pptx": "output/deck.pptx", "audit": "output/static-audit.json"}
}
```

Rules:

- Array order determines slide order. Stable slide IDs must not change when a page is reordered. IDs use letters, digits, hyphens or underscores, begin with a letter or digit, and have at most 80 characters. IDs are unique in their array; content IDs may repeat across different slides or different content kinds.
- `title` is actual visible slide text, not an internal caption. It is automatically checked for retention. Do not duplicate it in `texts` unless the slide intentionally contains the same title twice.
- Text/formula/table/chart items are required by default; set `required: false` only when omission is intentional. Each can carry `source_ids`. Tables use rectangular `cells`; charts use explicit categories and finite numeric series values of matching length.
- Per-item `editable` overrides the corresponding deck default. Do not silently set it false because reconstruction is difficult. Photos, illustrations and textures ordinarily remain images. Editable equations require native Office Math, not an SVG, outlined path, PNG or plain LaTeX string.
- `display_text` is optional expected Office Math token text, for example `x2` for an expression whose tokens are x and 2. It does not encode superscript structure. Preserve `latex` as the mathematical source and review structure visually and mathematically. Native equation presence and token matching do not prove equivalence.
- `style` is an extensible object for the selected style rules. Additional source-grounded content metadata is allowed. Helpers validate the core fields, not every custom field.
- Every source/asset record must refer to an existing regular file. Optional `sha256` must be lowercase 64-character hex and must match. Do not write API keys or account credentials into task metadata.
- All declared file paths use POSIX paths relative to the task root. Absolute paths, backslashes, empty/`.`/`..` segments, URI/drive prefixes and symlinks escaping the task are rejected. `outputs` may point to future files, but still must be safe project paths.

## Per-page `objects.json`

```json
{
  "schema_version": 1,
  "slide_id": "slide-001",
  "design": {"path": "slides/slide-001/design/v002.png", "sha256": "<actual SHA-256>"},
  "objects": [
    {
      "id": "obj-title",
      "kind": "text",
      "content_ref": "title",
      "editable": true,
      "bbox": [0.5, 0.4, 11.8, 0.7],
      "pptx_shape_id": "2"
    },
    {
      "id": "obj-summary",
      "kind": "text",
      "content_ref": "text:summary",
      "editable": true,
      "bbox": [0.5, 1.5, 5.8, 1.2],
      "pptx_shape_id": "3"
    }
  ]
}
```

- Select the exact current design file. Use versioned filenames; do not overwrite a user-selected design without retaining its prior version. A supplied design hash is verified.
- Object `kind` is one of `text`, `formula`, `table`, `chart`, `image`, `shape`, `connector`.
- `content_ref` is `title`, `text:<id>`, `formula:<id>`, `table:<id>`, `chart:<id>` or `asset:<id>`. The referenced entry must exist and its kind must match. Unbound decorative shapes/connectors are allowed. Bind every required content item and required asset before final reconstruction audit.
- `bbox` is `[x, y, width, height]` in slide inches. All numbers are finite; sizes are nonnegative and cannot both be zero. Connector extent may have one zero dimension.
- `pptx_shape_id` is the actual numeric `p:cNvPr/@id` of the exported object, serialized as a string. IDs are local to a slide. Add IDs after export when the authoring library assigns them automatically. Multiple objects may bind the same content reference when a text block must be split. Every mapped ID must exist in the final page.
- Missing exported IDs receive an audit warning and weaker content/type fallback checks. They cannot establish a precise source-to-object match. Populate them for final acceptance wherever supported.
- Default `project validate` permits pages whose design/map has not yet been created. `--require-maps` requires a valid selected design and mappings for all required content. It does not require the PPTX to exist.

## Audit and status interpretation

The static audit checks ZIP integrity, required package parts, XML parsing, internal relationship targets and undeclared relationship references. It follows presentation order instead of assuming that numbered filenames imply page order. Reports include per-page native text, table, chart, equation, image, shape and connector counts.

With `--project`, the audit also checks title/text retention, mapped object identities, table cells, chart cached series values/categories and required native object types. Formula checks establish native Office Math presence and optional display tokens. Hidden/negligibly small text does not satisfy native text retention. A whole-page raster with no useful native content is rejected when native text is required.

Translated/scaled groups are accounted for in object bounds. Ordinary shape rotation is included. Rotated/flipped group geometry, master-inherited positions, effects, shadows, text-frame internal clipping, glyph fallback and mathematical meaning require actual rendering and inspection. A successful audit means static checks passed; it never means visually approved.

`status` derives source/design/map/output presence from current files. A static audit is current only when its success flag is true, errors are empty, and PPTX, deck and per-page object-map hashes match current files. Its highest automatic stage is `static_checked_visual_review_pending`; it never reports `complete` or manufactures visual approval. Later content, mapping or PPTX changes invalidate that audit. Keep visual review evidence separately and base delivery on the exact final artifact.
