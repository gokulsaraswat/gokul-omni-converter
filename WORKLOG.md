# WORKLOG: gokul-omni-converter

Master status of all repos: platform/docs/WORKLOG.md. Update this file after every step.

## Current state (2026-11-25, upgrade run phase A)
- GO Omni Pro 2.4.0, Python desktop app (CustomTkinter). Runs from source on Python 3.12 (`.venv`, gitignored).
- Contract: `smoke_test.py` green; `tests/test_cores.py` 15/15; `tests/test_ui_smoke.py` 33/33 (with `PYTHONUTF8=1`).
- Fixed: PDF overlay tools crashed on pypdf 7 (OMN-B-001).

## Done
- 2026-11-25 fix(pdf): attach page to the writer before merging overlays (OMN-B-001).

## In progress / next
- Phase A: pin customtkinter < 6, docs/FEATURES.md + baseline numbers, docs/ARCHITECTURE.md, kit fixes.

## Assumptions
- Python 3.12 is the dev/build interpreter (README says 3.9+; 3.12 has wheels for every dependency and PyInstaller 6).

## Blockers / requests to other repos
- none

## Reminders
- Scope: standalone desktop app; hub links to it only.

- 2026-11-26 Upgrade program: status of all 7 repos and how to resume is in platform/docs/UPGRADE-PROGRAM.md (read it first).
