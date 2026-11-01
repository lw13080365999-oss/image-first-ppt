---
name: image-first-ppt
description: Create beautiful editable PowerPoint decks from documents, text, data, or reference slides by planning content, generating full-slide designs with an image model, and reconstructing native PPT objects. Use for general PPT creation, image-led redesign, or revising decks made by this skill. Automatically infer routine choices and proceed. Not for text-only outlines, read-only PPT analysis, or small edits that do not need image-led design.
---

# Image-first PPT

Turn supplied material into an attractive, editable deck. **The image model designs the pages; GPT reconstructs editable objects from those designs and the exact source content.** Preserve that production order. User instructions override these defaults.

## Working defaults

- Start from the supplied material. Infer audience, purpose, slide count, density and style; use source language and 16:9 unless specified. State useful assumptions briefly and continue. Do not impose an intake questionnaire, outline approval or per-image approval. Offer preview/selection only when requested or genuinely useful without blocking progress.
- Preserve requested slide count, order, evidence and branding. Do not invent missing facts, data, citations, comparisons or results to fit a layout.
- Default editability: visible text, data tables/charts and key diagram structure are native editable objects. Photos, complex illustration and texture remain separate images. Formulas require actual editable math when requested; an SVG is not proof of editability.
- Keep project state in the user's task folder, separate from this installed skill. Do not store client material or modify style assets here.

## 1. Read and plan

Read [content-planning.md](references/content-planning.md). Extract documents using available document/PDF/spreadsheet tools as appropriate; inspect original figures and tables when extraction loses meaning.

Create a project with `python3 scripts/project.py init <project-dir>`. Read [project-contract.md](references/project-contract.md) before writing `deck.json` and per-slide object maps. Preserve exact text, source IDs, data, units, qualifications, equation source and asset roles. Keep stable slide/content IDs across revisions. The project format records decisions; it is not a fixed page template.

Determine each page's message, role, information relationships and evidence. Plan rhythm across the deck. Use semantic layout choices (comparison, sequence, hierarchy, evidence-led image, etc.) as input to image design. Do not pre-render the final design with a programmatic PPT template.

## 2. Choose a visual system

Read [style-system.md](references/style-system.md); load only the selected style in [catalog.json](assets/styles/catalog.json). Default to automatic selection. User references and explicit choices take precedence. [gallery.html](assets/styles/gallery.html) offers eight comparable style specimens when the user wants to browse. These are design specimens, not image-model performance samples.

Lock palette roles, typography, illustration language, texture and density. Vary composition with content; do not repeat the same cards or copy a cover layout onto every page.

## 3. Generate full-slide designs

Read [image-design.md](references/image-design.md) and the image-generation portion of [runtime.md](references/runtime.md). Use the actual available image tool/model; respect a user-pinned backend. `image2`/`image2.5` names must resolve to a supported model ID when the backend requires one. Do not claim a model was used unless the call identifies it.

Generate a representative content page first and inspect it internally, then generate remaining pages with the same visual system. Each prompt receives deck context, exact page content, semantic relationships, asset roles and style. Keep all required pages; batching is optional and backend-dependent. Save design candidates and prompts in the project. Check missing components, wrong text and distorted evidence before reconstruction. If the user requests a full preview/revision loop, show an ordered contact sheet plus readable individual pages, revise only affected pages, and wait only at the confirmation points they requested.

If no image backend is available, preserve the completed content plan and explain that precise blocker. Do not silently substitute code drawings, ordinary PPT templates, another model or an image-only deck for this route. Do not build an API runner or solicit credentials in chat; follow an explicitly chosen backend's setup flow.

## 4. Reconstruct editable pages

Read [reconstruction.md](references/reconstruction.md). Use **exact content + design image + object map** together. Content supplies facts and text; the image supplies visual placement and hierarchy. Reuse original evidence assets or recreate quantitative objects from source data. Obtain clean backgrounds and independent illustration crops; do not leave flattened text underneath editable overlays.

GPT authors each page's `objects.json` from the image. It must identify required components and their native/image representations. The bundled `build_pptx.mjs` executes these explicit positions; it does not infer layout or design a deck. Read [runtime.md](references/runtime.md) for its supported fields and commands. An existing trusted authoring tool may be used for richer features, keeping the same object and QA contracts.

For native formulas, provide checked Office Math XML per formula and use `embed_formulas.py` after the draft export. Preserve LaTeX beside it. Do not treat generated XML as automatically equivalent to the source expression; verify the math and final rendering.

## 5. Verify the final artifact

Read [quality-assurance.md](references/quality-assurance.md). Run:

```bash
python3 scripts/project.py validate <project-dir> --require-maps
python3 scripts/bind_shape_ids.py <final.pptx> --project <project-dir>
python3 scripts/audit_pptx.py <final.pptx> --project <project-dir> --out <audit.json>
python3 scripts/render_review.py <final.pptx> --project <project-dir> --out <new-review-dir>
```

Confirm command syntax with `--help` in a new environment. A static audit is necessary but cannot establish visual or mathematical correctness. Inspect every rendered page at full size, and compare it with its design and content. Check formula glyphs, superscripts/subscripts, missing labels, overlaps, clipping and preserved hierarchy. Use a contact sheet for deck-wide rhythm. An independent reviewer can inspect rendered artifacts with the prompt in the QA reference; do not force it to invent a flaw.

Fix the source/map/assets and rebuild a new revision, then re-render affected pages. Reviews bind to exact PPTX/render hashes; stale previews cannot approve a new build. Never call a deck fully verified when final rendering or a required editability check is unavailable. Preserve usable artifacts and disclose the specific remaining limit. Resolve routine fixes autonomously; stop repeated non-improving retries rather than accumulating costs indefinitely.

## 6. Revise and deliver

Update only affected slides for local feedback. Small wording edits can change native text directly, followed by layout checks; a new composition or illustration normally goes back through image design. Global style changes require a deck-wide review. Keep concise version history and actual file state using `project.py status`.

Deliver the editable `.pptx` and useful preview, not internal scaffolds by default. Explain only limitations that affect use. Keep prompts, validation records and production commentary out of slides and speaker notes. Notes contain presentation content and sources when useful.

## Supporting material

- [Runtime and reconstruction adapter](references/runtime.md): dependencies, commands, supported object fields, formulas and rendering.
- [Open-source provenance](references/provenance.md): researched sources, reuse scope and licenses.

Load references when their stage applies; do not read the entire research archive or every style for each deck.
