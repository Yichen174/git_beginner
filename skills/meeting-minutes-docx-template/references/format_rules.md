# Format Rules

Use these rules when generating Chinese sell-side meeting minutes from a sample Word document.

## Non-Negotiables

- Copy the sample `.docx` as the mother template.
- Do not start from a blank Word document.
- Do not manually rebuild numbering, margins, or styles unless template copying is impossible.
- Add new content by copying same-role prototype paragraphs from the sample and replacing text.

## Prototype Mapping

- Title: copy the sample company title paragraph.
- Date: copy the sample `时间：YYYY.M.D` paragraph.
- Summary title: copy `要点总结：`.
- Summary topic: copy a first-level auto-numbered summary topic paragraph.
- Summary detail: copy an auto-numbered detail paragraph under a topic.
- Company overview heading: copy `一、公司概要`.
- Company overview body: copy a body paragraph under company overview, preserving first-line indent.
- Q&A heading: copy `二、Q&A`.
- Q&A topic: copy a hand-typed `1. 主题` paragraph.
- Question: copy a `Q：...` paragraph and keep the full question bold.
- Answer: copy an `A：...` paragraph and keep only `A：` bold.
- Q&A separator: only inside the Q&A section, insert one blank paragraph between each complete Q/A pair. Do not add these separator blanks in the summary or company overview sections.

## Expected Structure

```text
公司简称交流纪要
时间：YYYY.M.D

要点总结：

1、主题一
1) 要点一
2) 要点二

一、公司概要
    公司概要正文。

二、Q&A

1. 主题一
Q：问题？
A：回答。

Q：问题？
A：回答。

2. 主题二
Q：问题？
A：回答。
```

## Validation

Use structural checks after saving:

- `List Paragraph` summary paragraphs have `numPr`.
- Company overview body paragraphs have first-line indent.
- Q runs are bold.
- A first run is bold and later body runs are not bold.
- QA pairs have one blank paragraph between adjacent Q/A pairs, and this separator rule is limited to the Q&A section.
- Table count is zero unless the user explicitly asks for tables.
