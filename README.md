# aps-doc-skills

Claude Code **marketplace** for drafting Korean project deliverables — styled `.docx` documents and `.pptx` reports.

## Plugin: `docs-report`

Skill **`drafting-project-documents`** — writes a 착수 보고서(kickoff report), 요구사항 정의서(requirements),
or 회의록(meeting notes) either by filling an existing draft or by interviewing you. Unknown values are
left as "(미정 — 추후 확정)" rather than guessed, and the wording adapts to the reader (임원 vs 실무).

Rendering uses a single spec-driven engine (`scripts/build_doc.py`, requires `python-docx`).

## Plugin: `ppt-report`

Skill **`creating-ppt-reports`** + command **`/ppt-report`** — creates a Korean PPT report (`.pptx`) via a
docs-first workflow: build a source `.docx`, then derive the slide deck from it. Enforces team rules
(author = IT전략팀; filename with spaces + version suffix; `noProof`/`lang=ko-KR` on every run so no
spell-check red squiggles).

Rendering uses `scripts/build_doc.py` (.docx) and `scripts/build_ppt.py` (.pptx), requires
`python-docx` and `python-pptx`.

## Install

```bash
# add this marketplace (after pushing to GitHub)
claude plugin marketplace add airforce1226/aps-doc-skills
# install the plugins
claude plugin install docs-report@aps-doc-skills
claude plugin install ppt-report@aps-doc-skills
```

Then use them either way:

- **Slash commands:** `/draft-doc 착수 보고서`, `/ppt-report 착수 보고서` (optionally add a draft file path)
- **Natural language:** just ask "착수 보고서 초안 작성해줘" or "PPT 보고서 만들어줘" (the skills auto-trigger)

## Requirements

- Python 3 with `python-docx` and `python-pptx`
  (`python -m pip install python-docx python-pptx`)

## Layout

```
aps-doc-skills/
  .claude-plugin/marketplace.json
  plugins/docs-report/
    .claude-plugin/plugin.json
    commands/draft-doc.md
    skills/drafting-project-documents/
      SKILL.md
      references/document-types.md
      scripts/build_doc.py
  plugins/ppt-report/
    .claude-plugin/plugin.json
    commands/ppt-report.md
    skills/creating-ppt-reports/
      SKILL.md
      scripts/build_doc.py
      scripts/build_ppt.py
```

## License

MIT
