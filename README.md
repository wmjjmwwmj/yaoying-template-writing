# yaoying-template-writing

This repository contains a Codex skill for drafting SinoUnited Health / yaoying policy and procedure `.docx` files from a locked 2026 Word template.

The skill keeps the bundled layout intact and writes content only through the provided template-filling script:

```bash
python3 scripts/write_from_template.py examples/minimal-input.json output.docx
```

See [SKILL.md](SKILL.md) for the required workflow and content rules.

