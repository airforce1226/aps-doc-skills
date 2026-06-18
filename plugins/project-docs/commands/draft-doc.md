---
description: 프로젝트 문서(착수 보고서·요구사항 정의서·회의록)를 초안 또는 대화로 .docx 작성
argument-hint: "[문서유형] [초안 파일 경로(선택)]"
---

You are starting the project-document drafting workflow.

**REQUIRED:** Use the `drafting-project-documents` skill and follow its workflow exactly (identify document type and reader → gather content from an existing draft or by interviewing the user → never invent unknowns, mark them `TBD` → match the reader's vocabulary → build the spec JSON → render with `build_doc.py` → read-back verify).

User request / arguments: $ARGUMENTS

Guidance for interpreting the arguments:
- If a document type is given (예: 착수 보고서 / 요구사항 정의서 / 회의록), use it. Otherwise ask which one.
- If a draft file path is given, fill that draft. If a draft clearly exists in the working folder but no path was given, ask whether to use it.
- If neither is provided, interview the user section by section using the standard outline for that document type.

Begin now by confirming the document type and the intended reader (임원 vs 실무), then proceed.
