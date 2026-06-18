# aps-doc-skills

Claude Code **marketplace** for drafting Korean project documents as styled `.docx` files.

## Plugin: `project-docs`

Skill **`drafting-project-documents`** — writes a 착수 보고서(kickoff report), 요구사항 정의서(requirements),
or 회의록(meeting notes) either by filling an existing draft or by interviewing you. Unknown values are
left as "(미정 — 추후 확정)" rather than guessed, and the wording adapts to the reader (임원 vs 실무).

Rendering uses a single spec-driven engine (`scripts/build_doc.py`, requires `python-docx`).

## Install

```bash
# add this marketplace (after pushing to GitHub)
claude plugin marketplace add airforce1226/aps-doc-skills
# install the plugin
claude plugin install project-docs@aps-doc-skills
```

Then use it either way:

- **Slash command:** `/draft-doc 착수 보고서` (optionally add a draft file path)
- **Natural language:** just ask "착수 보고서 초안 작성해줘" (the skill auto-triggers)

## Requirements

- Python 3 with `python-docx` (`python -m pip install python-docx`)

## Layout

```
aps-doc-skills/
  .claude-plugin/marketplace.json
  plugins/project-docs/
    .claude-plugin/plugin.json
    skills/drafting-project-documents/
      SKILL.md
      references/document-types.md
      scripts/build_doc.py
```

## License

MIT
