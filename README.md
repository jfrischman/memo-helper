# Memo Helper

Local Python app for LP-led secondary exposure rollups and memo-linked project storage.

## Run

```powershell
python app.py
```

The app opens at `http://127.0.0.1:8765/`.

On Windows, you can also double-click [launch_project_balance_updater.cmd](C:/Users/jfrischman/Documents/Codex/2026-06-02/files-mentioned-by-the-user-project/launch_project_balance_updater.cmd) from this folder or use the desktop shortcut.

## What it does

- start a new project and attach it to a memo
- save uploaded workbooks and project settings on disk
- upload one workbook per fund
- choose the sheet for each fund
- let the app auto-detect whether the sheet has headers or is raw data
- map the important columns
- enter proposed bid amounts
- calculate project-level exposure by record date NAV
- view charts, tables, and a downloadable JSON snapshot
- add manual category overrides for funds that do not carry those fields in the workbook

## Expected workbook shape

The app works best when each fund workbook has one row per underlying investment and a header row with columns similar to:

- investment name
- record date NAV
- asset class
- security type
- geography
- sub-asset class

## Notes

- The app assumes the first row of the selected sheet is the header row.
- `record date NAV` is the weighting basis inside each fund.
- Project-level exposure is weighted by proposed bid amount across funds.
- `sub-asset class` mirrors `security type` by default.
- Cash-like rows are excluded from the invested asset normalization base, so a 98% invested tape rolls up to 100% across the invested assets.
