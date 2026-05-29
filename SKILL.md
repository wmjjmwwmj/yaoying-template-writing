---
name: yaoying-template-writing
description: "Use this skill to draft, fill, or revise a Word .docx policy/procedure document with the bundled SinoUnited Health/曜影医疗 2026制度文件模板. All models must follow the fixed workflow exactly: preserve the template layout byte-for-byte outside approved content fields, convert content to structured JSON, and fill via the provided script without changing the template layout."
---

# yaoying-template-writing

本 skill 用内置的曜影医疗 2026 制度文件模板生成 `.docx`。所有模型都必须严格遵循本文流程；不要根据模型能力自行放宽步骤、重建 Word 文档或直接手改 OOXML 版式。

最终原则只有一个：输出文件的样式、版式、页眉页脚、logo、边框、编号、表格结构和模板保持一致；模型只负责生成结构化内容，写入动作由脚本完成。

## Fixed Layout Contract

- Immutable layout asset: `assets/layout/yaoying-2026-policy-layout.docx`.
- Only write content through `scripts/write_from_template.py`; the script copies every non-content OOXML part unchanged.
- The script removes Word content-control shells from the final document after filling, so written content remains ordinary editable Word text.
- Do not use `python-docx`, `docx-js`, or a new blank document to recreate the template.
- Do not directly edit `word/styles.xml`, `word/numbering.xml`, headers, footers, images, relationships, or table borders.
- If a field is unknown, leave it as an empty string instead of inventing content.
- Treat `word/document.xml` content fields as the only editable layer. All other package parts are layout assets.

## Workflow

1. Read `references/template-map.md` before writing. If the user gives raw prose, convert it into the JSON structure in `references/input-schema.json`.
2. Create an input JSON file. Start from `examples/minimal-input.json` and fill only known fields.
3. Generate the Word file:

```bash
python3 scripts/write_from_template.py input.json output.docx
```

4. Deliver the final `.docx`; do not deliver the JSON unless the user asks.

## Content Rules

- Keep `sections.content` as a list of objects: `{"level": 1, "text": "..."}` or `{"level": 2, "text": "..."}`. Level 1 maps to `5.x`; level 2 maps to `5.x.x`.
- Keep top metadata short; these fields live inside narrow fixed table cells.
- In `training_plan`, `object`, `cycles`, and `methods` are fixed checkbox fields. Use only the template options and let the script check them. `remarks` is the only free-text cell.
- `resources` is fixed to five rows: `人力资源`, `财务预算`, `设施空间`, `资材物料`, `科技信息`.
- `revision_history` supports up to three rows because the template has exactly three blank history rows.
- Long narrative content can push the fixed shell onto more pages. Keep content concise and set `metadata.page_count` to the best expected final page count.

## Filling Discipline

- Write clean JSON values only. Do not put Markdown bullets, manual numbering, table syntax, HTML, XML, or decorative separators inside text fields.
- Avoid unnecessary newlines. Use a single string for normal paragraphs; use arrays only where the template expects multiple content items, such as `definitions`, `responsibilities`, `training_plan.cycles`, and `training_plan.methods`.
- For `training_plan.cycles`, use only `1次或岗前`, `1月`, `1季度`, `1年`. For `training_plan.methods`, use only `自学`, `在线学习`, `授课`, `开会`. Do not type checkbox symbols manually.
- Do not add leading/trailing spaces. Do not use blank lines to create visual spacing; the template already controls spacing.
- Do not include visible section numbers in content. For example, use `"服务对象包括..."`, not `"5.1 服务对象包括..."`.
- Keep narrow-cell fields concise: `department`, `document_number`, `version`, `writer`, `reviewer`, and dates should fit on one line when possible.
- If content is likely to wrap awkwardly in narrow cells, shorten the JSON value before generation.
