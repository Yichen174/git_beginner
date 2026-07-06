---
name: meeting-minutes-docx-template
description: Generate Chinese sell-side meeting minutes DOCX files from raw notes using a sample Word document as the formatting template. Use when creating or formatting 会议纪要/卖方会议纪要/Q&A minutes from .docx notes while preserving a reference .docx template's styles, numbering, margins, paragraph prototypes, and Q&A bold rules.
---

# Meeting Minutes DOCX Template

## Core Rule

Use a sample `.docx` as the mother template. Do not create the final Word document from a blank file when the user provides or references a sample meeting-minutes document.

Preserve the sample document's Word-native format by copying prototype paragraphs and replacing text. Let the template carry `styles.xml`, `numbering.xml`, page setup, `Normal`, `List Paragraph`, numbering definitions, indentation, spacing, and font inheritance.

## Workflow

1. Read the raw notes document and the sample `.docx` template.
2. Draft structured content in this order: title, date, `要点总结：`, summary topics/details, `一、公司概要`, company overview, `二、Q&A`, Q&A topics/questions/answers.
3. Read `references/content_rules.md` before finalizing content judgments, especially when adding Q&A or inferring growth-curve labels.
4. Read `references/format_rules.md` before generating the DOCX.
5. Use `scripts/build_minutes_docx.py` or the same prototype-copy method. Prefer running the script over rewriting fragile python-docx logic.
6. Run structural validation after generation. If LibreOffice/soffice is available, render and inspect pages. If not available, say visual render QA could not be completed.

## Script Usage

The script expects a JSON content file so Codex can think through the content first, then hand deterministic formatting to the script.

```powershell
python scripts/build_minutes_docx.py --template sample.docx --content content.json --out 20260705公司会议纪要（请勿外发）.docx
```

Content JSON shape:

```json
{
  "title": "鹏翎股份交流纪要",
  "date": "时间：2026.7.5",
  "summary": [
    {"topic": "公司战略与热管理业务", "items": ["要点一", "要点二"]}
  ],
  "company_overview": ["公司概要第一段", "公司概要第二段"],
  "qa": [
    {"topic": "1. 公司战略与热管理业务", "items": [{"q": "问题？", "a": "回答。"}]}
  ]
}
```

## Validation Checklist

After generating the DOCX, verify at least:

- Summary paragraphs use Word numbering (`numPr`), not hand-typed `1、` / `1)` markers.
- Summary topic paragraphs use the template's first-level numbering prototype.
- Summary item paragraphs use the template's detail numbering prototypes.
- Company overview body paragraphs preserve first-line indent.
- Q paragraphs are fully bold.
- A paragraphs have only `A：` bold and body text not bold.
- No tables, images, or decorative page furniture were introduced.
- Counts match the drafted content: summary item count, Q count, A count.

## Content Discipline

Reasonable Q&A supplementation is allowed, but answers must be grounded in the source notes. Do not invent revenue, profit, order, customer, or timing data that is not present. Use conservative language for uncertain points.
