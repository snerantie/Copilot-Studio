# For Manager — Manco Capture Agent

Documents for review by your manager and (where relevant) the Power Platform
admin. Both are written to be readable without technical context.

## Documents

| Document | Format | Audience | Purpose |
|---|---|---|---|
| [`business-case-upgrade.md`](./business-case-upgrade.md) | Markdown (canonical) | Manager (decision-maker) | Justify the Copilot Studio capability upgrade and AI Builder Prompts admin policy review |
| [`business-case-upgrade.docx`](./business-case-upgrade.docx) | Word | Manager | Same content, ready to email or attach |
| [`manco-capture-business-case.pptx`](./manco-capture-business-case.pptx) | PowerPoint (11 slides, 16:9) | Manager (meeting) | Distilled exec deck of the business case for a verbal walkthrough |
| [`status-and-blockers.md`](./status-and-blockers.md) | Markdown (canonical) | Manager + Power Platform admin | Snapshot of Sprint 1 progress, exact blockers, decisions needed |
| [`status-and-blockers.docx`](./status-and-blockers.docx) | Word | Manager + admin | Same content, ready to email or attach |

## How to use

- **Email a single document**: attach the `.docx` directly to Outlook. No
  reformatting needed.
- **Walk the manager through verbally**: open the `.pptx`. The slides are
  designed to be presented in ~10–15 minutes, with the TL;DR on slide 2
  and the decision asks on slide 11.
- **Skim on GitHub**: open the `.md` versions in the browser — they render
  with proper tables and links.
- **Edit before sharing**: the `.docx` and `.pptx` are generated from the
  Python script in [`_scripts/generate_docs.py`](./_scripts/generate_docs.py).
  Either edit the docx/pptx directly in Word/PowerPoint, or update the
  script and re-run (`python _scripts/generate_docs.py`) for fully
  regenerable artefacts.

## Placeholders to fill before sharing

- `[your name]` — your name as author of the business case
- `[your manager]` — your manager's name on the "For" line

Both placeholders appear in the title block of the Word doc and the title
slide of the deck. Either edit in place or update `AUTHOR_PLACEHOLDER` and
`MANAGER_PLACEHOLDER` in `_scripts/generate_docs.py` and re-run.

## Caveats to review before sharing

- **Pricing** uses Microsoft list pricing as of mid-2026 (~£164/user/month
  for Copilot Studio). Confirm with M365 procurement before quoting these
  numbers externally.
- **Payback estimate** uses a conservative 1 hour saved per senior leader
  per week. Adjust if you have a stronger anchor.

## Regenerating the Word and PowerPoint files

```bash
cd manco-capture/for-manager
pip install python-docx python-pptx
python _scripts/generate_docs.py
```

The script produces all three files in this folder.
