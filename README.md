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

## Plugin: `design-ppt`

Skill **`creating-design-ppt`** + command **`/design-ppt`** — builds an **APS Brand Presentation System**
deck into a Korean `.pptx`. You assemble a `deck.html` from 14 section archetypes (`assets/sections/*.html`);
each `<section>` (1920×1080) becomes a slide. Default output is **editable PowerPoint native objects**
(textboxes/tables/shapes) measured from headless-Chrome layout; `--mode image` bakes pixel-faithful
screenshots instead. Page numbers are auto-stamped on every slide except the cover. Enforces team rules
(author = IT전략팀; 대외비 badge; `noProof`/`lang=ko-KR`).

**Browse the templates:** open **`plugins/design-ppt/skills/creating-design-ppt/assets/templates-gallery.html`**
in a browser to see every archetype at a glance (regenerate with `python scripts/build_gallery.py`).

Rendering uses `scripts/build_design_ppt.py` (+ `scripts/native_render.py`), requires `python-pptx`,
`Pillow`, and Chrome/Edge.

## Install

```bash
# add this marketplace (after pushing to GitHub)
claude plugin marketplace add airforce1226/aps-doc-skills
# install the plugins
claude plugin install docs-report@aps-doc-skills
claude plugin install ppt-report@aps-doc-skills
claude plugin install design-ppt@aps-doc-skills
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
  plugins/design-ppt/
    .claude-plugin/plugin.json
    commands/design-ppt.md
    skills/creating-design-ppt/
      SKILL.md
      assets/sections/*.html        # 14 slide archetypes
      assets/templates-gallery.html # browser preview of all archetypes
      assets/design-tokens.md
      scripts/build_design_ppt.py
      scripts/native_render.py
      scripts/build_gallery.py
```

## License

MIT
