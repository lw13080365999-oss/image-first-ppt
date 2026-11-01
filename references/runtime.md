# Runtime and object-map adapter

## Capability discovery

Before consuming generation time, check the actual environment for an image backend, PPT authoring runtime and final-file renderer. Tool availability can differ between chats. This skill supplies planning, reconstruction and validation; it does not bundle image-model weights, an API account, or an image service.

In Codex, use the installed `imagegen` skill when available and follow its backend rules. Prefer the exposed image tool. If the user chooses a supported API/CLI route, use that route's maintained client and secure credential setup. Do not add keys to the skill, prompts or project. No backend means image generation is blocked, not permission to draw replacement slides in code. Preserve the content plan and report the precise capability needed.

Do not freeze model IDs or sampling controls into content/style prompts. Record a user-pinned model in project configuration and verify support before calling it. A tool that does not expose its model cannot substantiate an exact-model claim.

## Project helpers

`project.py`, `audit_pptx.py`, `embed_formulas.py`, `bind_shape_ids.py` use Python 3.10+ standard library. `render_review.py` additionally needs Pillow, LibreOffice and Poppler. The reconstruction adapter uses Node.js ESM and `@oai/artifact-tool`. Do not install arbitrary replacements into bundled runtimes.

For Codex desktop, call `load_workspace_dependencies` to locate the available Python/Node and package paths. Set the task-specific `PPT_NODE_MODULES` to the returned Node packages directory when bare module resolution is unavailable. Set `PPT_BIN_DIR` to its override binaries directory for the bundled renderer (or pass `--soffice`/`--pdftoppm` explicitly); a system PATH wrapper may point to an absent Office application. Use the returned Node/Python executables, not a hardcoded user path. Resolve this skill's folder from its actual installed location.

```bash
node scripts/build_pptx.mjs --check-runtime
python3 scripts/project.py init /absolute/task --title "Presentation" --source /absolute/input.txt
python3 scripts/project.py validate /absolute/task --require-maps
node scripts/build_pptx.mjs --project /absolute/task --output build/draft-v001.pptx
```

The project must contain real, selected design images and corresponding object maps before building. The adapter verifies design hashes. It places explicitly mapped objects, not generic slide templates. New output paths avoid accidental replacement. A `.build.json` report is saved beside the draft. A successful build is not final acceptance.

When the installed presentation tool requires operation markers/finalization, follow that tool's authoring and finalization instructions as well. A trusted alternative can author richer pages; do not force a limited adapter to approximate unsupported features silently. Keep maps aligned with the actual exported objects.

## Supported object properties

Core fields are defined in [project-contract.md](project-contract.md). The bundled adapter adds:

| Object | Additional fields |
|---|---|
| All positioned objects | `bbox: [x,y,width,height]` in inches; creation order is layer order |
| Text | `font_family`, `font_size_pt`, `color`, `bold`, `italic`, `align`, `valign`; `content_ref` required |
| Shape | `geometry` (Office preset), `fill`, `line_color`, `line_width_pt`; optional literal `text` for non-source decorative labels |
| Image | `path` or bound `asset:<id>`, `fit: contain|cover`, `alt`; use actual independent assets |
| Table | `table:<id>`, `font_size_pt`, `font_family`, `header_fill`, `fill`, `line_color` |
| Chart | `chart:<id>`, `chart_type` supported by authoring engine, optional documented `chart_options`; actual series remain source-bound |
| Connector | `from`/`to` object IDs, `from_side`, `to_side`, `connector_kind`, `line_color`, `line_width_pt`, `arrow` |
| Formula | `formula:<id>`, `omml_path`, `font_size_pt`, `font_family`; native equation patch required after draft export |

Map-level `background` supplies a clean solid background; richer textures/gradients can use separate clean image assets. Do not put a full text-bearing reference image beneath native text. Source-critical labels belong in content records even if drawn on a shape. Connectors bind to native shape endpoints; specify endpoints deliberately and inspect routing.

The adapter's direct text uses `typeface` and explicit point-to-pixel conversion at 96 DPI; no automatic font shrinking. The selected font also becomes the presentation theme default. Choose installed fonts with required glyph coverage. Check chart labels/axes/legend fonts separately in the final render; this engine's chart text style API does not expose a typeface override. Tables use native cells; advanced merges and chart styles may need the trusted authoring surface directly.

Name-bound text/shape/formula objects export as `ifppt:<slide_id>:<object_id>`. After the final export and equation patch, run `bind_shape_ids.py` to populate each map object's `pptx_shape_id` from actual `p:cNvPr/@id`; do not assume creator IDs equal PowerPoint numeric IDs. The helper follows presentation relationship order, prefers exact managed names, and otherwise requires a unique native type and bbox match (default per-coordinate tolerance 0.03 inches). Anonymous connectors require a unique directed `stCxn`/`endCxn` match to their already-bound `from`/`to` objects; a stale connector bbox does not determine identity. Ambiguous matches, missing endpoints and named type mismatches are hard errors; no map is changed when matching fails. Use `--dry-run` to inspect proposed bindings first.

Binding modifies maps only, never the PPTX. Changed maps invalidate older static audits and hash-bound reviews. Run binding before final audit/render/review; repeated binding is idempotent when IDs and maps are unchanged. Audit warnings about missing identity mappings require attention.

## Editable equations

The adapter intentionally does not pretend to be an arbitrary LaTeX converter. GPT supplies exact, structurally checked Office Math XML (OMML) from the saved equation source, or uses an available trusted converter. For complex expressions, validate terms, grouping, dimensions and math structure explicitly. Prefer tested converters when available.

Store each expression as a project-relative file, root `m:oMath` or `m:oMathPara`, using the standard Office Math namespace. A simple example for x squared:

```xml
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>
</m:oMath>
```

Map the formula with `omml_path`. A formula object's `font_family` or project `style.math_font_family` chooses missing math-run fonts, with `STIX Two Math` as a fallback choice that still requires availability checks. Existing explicit OMML font properties are preserved. Use a textbox proportioned to the actual equation, rather than stretching an equation across an oversized region. The builder creates a named temporary placeholder. The patcher replaces it with native `a14:m` Office Math without flattening the page:

```bash
python3 scripts/embed_formulas.py /absolute/task/build/draft-v001.pptx /absolute/task/output/deck-v001.pptx --project /absolute/task
```

It rejects missing placeholders, empty expressions, DTDs/entities and foreign XML namespaces. It sets missing math font/size properties but cannot prove LaTeX equivalence. Run the audit and render the patched final file, not the unpatched draft. Office Math existence establishes structural editability, not all-target-software compatibility. Do not silently downgrade to a bitmap when editable equations are required.

## Final rendering and review

```bash
python3 scripts/bind_shape_ids.py /absolute/task/output/deck-v001.pptx --project /absolute/task
python3 scripts/audit_pptx.py /absolute/task/output/deck-v001.pptx --project /absolute/task --out /absolute/task/output/static-audit.json
python3 scripts/render_review.py /absolute/task/output/deck-v001.pptx --project /absolute/task --out /absolute/task/review/v001
```

Use a new empty review directory for every revision. The renderer uses an isolated LibreOffice profile, exports the actual PPTX to PDF, and rasterizes all pages. It produces a contact sheet, side-by-side design/render comparisons and a hash-bound `review.json` whose status remains **pending** until an actual reviewer inspects it. Use PNG/JPEG designs for the comparison helper; SVG style specimens are not selected final page designs.

If the bundled renderer cannot see an existing CJK/math font, add `--font-file /existing/fonts/Chinese.ttf --font-file /existing/fonts/Math.otf`. The helper copies these files into its temporary Office profile, records their hashes and removes that profile afterward; it does not change installed fonts or the bundled runtime. `--font-dir /existing/fonts` loads font files directly inside one directory, without recursion. Use only required fonts. This affects the render environment and does not embed fonts into the PPTX or establish target-software availability. Confirm the resulting glyphs by viewing the actual pages.

Review every individual render; a montage cannot establish formula legibility. In a separate record, cite the reviewed PPTX hash, render hashes, reviewer, actual checks and unresolved issues. A changed file invalidates that record. `render_review.py` does not open PowerPoint/WPS or automatically approve appearances. If target software is available, inspect sensitive pages there as well.
