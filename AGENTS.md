# Agent Notes

## App Workflow
- `report.py` is the user-facing entrypoint; run it with `python3 report.py` after installing `openpyxl`.
- Install Python dependencies with `pip install -r requirements.txt`; this repo currently only declares `openpyxl`.
- PDF generation depends on a system LibreOffice binary (`libreoffice` or `soffice`), not a Python package.
- `report-template.xlsx` is the source template and must stay reusable; generated XLSX files are temporary conversion artifacts and should be deleted after PDF creation.
- Generated reports go under `temp/`, which is gitignored.

## Template Mapping
- Ticket fields are written to merged-cell anchors only: `A11`, `B11`, `D11`, `H11`, `A14`, `D14`, `F14`, `G14`, `H14`.
- Summary rows start at row `18`; extra issues insert rows and copy row `18` styling/merged ranges.
- Current summary inputs are `Issue`, `Action`, and `Status`; `Action` is merged across the former issue-type area, so do not re-add an `Issue Type` prompt unless the template changes again.

## Verification
- Quick syntax check: `python3 -m py_compile report.py`.
- Focused end-to-end check: pipe sample answers into `.venv/bin/python report.py`, verify `temp/<Ticket ID>.pdf` exists, and verify the temporary `.xlsx` was removed.

## Git Workflow
- Before starting work, check the branch with `git branch --show-current`; if it is `main`, create a feature/fix/docs branch before editing.
- Before committing, inspect `git status --short`, `git diff`, and `git log --oneline -10`.
- Always commit completed changes and create a PR with details; never approve or merge the PR yourself.
