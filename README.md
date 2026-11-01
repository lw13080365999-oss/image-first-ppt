# Image-first PPT

A Codex skill for turning source material into editable PowerPoint decks. It plans the content, has an image model design each slide, rebuilds native PowerPoint objects from the design and source material, then checks the exported deck.

## Install

Clone this repository into the Codex skills directory:

```bash
git clone https://github.com/lw13080365999-oss/image-first-ppt.git "$HOME/.codex/skills/image-first-ppt"
```

Restart Codex after installing the skill. If your Codex setup uses a different skills directory, place the `image-first-ppt` folder there instead.

## Use

Ask Codex to use `image-first-ppt` and provide the source files and any requirements for audience, slide count, style, branding, or editability. The skill keeps project materials in the task folder and does not store them in the installed skill directory.

The skill relies on capabilities supplied by the active environment. In Codex, it uses the `imagegen` skill when available and follows that skill's backend rules. It does not include image-model weights, an image service, API credentials, or a PowerPoint rendering runtime.

The bundled project and audit scripts use Python 3.10+ standard library. The editable deck adapter uses Node.js and the environment-provided `@oai/artifact-tool` package. Full render review additionally needs Pillow, LibreOffice, and Poppler. See [runtime.md](references/runtime.md) for environment discovery and commands.

## Contents

- `SKILL.md`: workflow and invocation guidance.
- `references/`: planning, project format, image design, reconstruction, runtime, provenance, and quality checks.
- `scripts/`: project validation, PowerPoint generation adapter, formula embedding, shape binding, auditing, and render review helpers.
- `assets/styles/`: eight visual style specimens and their catalog.
- `agents/openai.yaml`: Codex UI metadata.

## Licensing

No general license is currently granted for the original material in this repository. Third-party source material and its license notices are listed in [provenance.md](references/provenance.md) and retained under [references/licenses](references/licenses/). Contact the author for permission to reuse or redistribute the original material.
