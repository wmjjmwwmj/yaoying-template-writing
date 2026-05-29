# Template Map

Use this file to decide where content belongs. The model should never infer layout from scratch; the bundled DOCX already contains the visual system.

## Locked Layout

- Page: A4 portrait, `11906 x 16838` DXA.
- Margins: top/right/bottom/left `720` DXA; header `851`; footer `992`.
- Main body is a large blue-bordered table containing the logo/title block, metadata rows, and the policy body.
- Header assets: `word/media/image1.png`, `word/media/image2.png`, `word/media/image3.png`.
- Footer text and page number are already in `word/footer1.xml`.
- Styles, numbering, table borders, shading, and page furniture are part of the layout layer. Do not edit them.

## Generation Boundary

`scripts/write_from_template.py` edits only `word/document.xml` and leaves all other DOCX package parts unchanged. It writes into:

- top metadata table cells,
- selected content controls in the policy body,
- fixed resource/training/revision tables,
- legacy checkbox state for document nature and selected section headings.

## Metadata Fields

Place these under `metadata`:

| JSON key | Template location | Notes |
|---|---|---|
| `file_name` | first row title content control | Main document title. |
| `file_category` | row 2 col 1 | Keep short. |
| `department` | row 2 col 2 | File owner department. |
| `document_number` | row 2 col 3 | Keep short. |
| `version` | row 3 col 1 | Example: `V1.0`. |
| `writer` | row 3 col 2 | Person or department. |
| `reviewer` | row 3 col 3 | Person or department. |
| `revision_date` | row 4 col 1 | Use final display text, e.g. `2026 年 5 月 28 日`. |
| `review_date` | row 4 col 2 | Same date style. |
| `effective_date` | row 4 col 3 | Same date style. |
| `document_nature` | row 5 col 1 checkboxes | One or more of `新增`, `修订`, `废止`, `回顾`. |
| `page_count` | row 5 col 2 | Best expected final page count. Keep content concise so this remains accurate. |

## Body Sections

Place these under `sections`:

| JSON key | Template section | Write behavior |
|---|---|---|
| `purpose` | `1. 目的` | Replaces the purpose placeholder. |
| `scope` | `2. 范围` | Replaces the scope placeholder. |
| `definitions` | `3. 定义` | String, list, or key-value object. |
| `responsibilities` | `4. 权责` | String, list, or key-value object. |
| `content` | `5. 内容` | List of numbered content items. |
| `process` | `6. 流程` | Short text; script inserts a paragraph if needed. |
| `quality_management` | `9. 质量管理` | Checks the heading box and writes text below it. |
| `risk_management` | `10. 风险管理` | Checks the heading box and writes text below it. |
| `attachments` | `11. 表单附件` | Checks the heading box and writes text below it. |
| `references` | `12. 参考文献` | Checks the heading box and writes text below it. |

`sections.content` items must use:

```json
{"level": 1, "text": "5.1 对应内容"}
{"level": 2, "text": "5.1.1 对应内容"}
```

Do not include the visible number (`5.1`) inside `text`; Word numbering supplies it.

## Resource Table

Place this under `resources`. The first column is fixed by the layout and should not be changed.

```json
"resources": {
  "人力资源": {"quantity": "", "description": ""},
  "财务预算": {"quantity": "", "description": ""},
  "设施空间": {"quantity": "", "description": ""},
  "资材物料": {"quantity": "", "description": ""},
  "科技信息": {"quantity": "", "description": ""}
}
```

## Training Table

Place up to two rows under `training_plan`. The first three columns are fixed checkbox fields in the template; do not rewrite them as plain text.

- `object`: use `全员` for the first row and `岗位` for the second row.
- `cycles`: choose from `1次或岗前`, `1月`, `1季度`, `1年`.
- `methods`: choose from `自学`, `在线学习`, `授课`, `开会`.
- `remarks`: free text, kept concise.

```json
"training_plan": [
  {"object": "全员", "cycles": ["1次或岗前", "1年"], "methods": ["自学"], "remarks": ""},
  {"object": "岗位", "cycles": ["1次或岗前"], "methods": ["授课"], "remarks": ""}
]
```

## Revision History

Place up to three rows under `revision_history`.

```json
"revision_history": [
  {"version": "V1.0", "summary": "新增制度文件", "date": "2026 年 5 月 28 日"}
]
```

If the user needs more than three history rows, ask whether to extend the fixed table. Extending it changes the layout and should be treated as a deliberate layout edit.
