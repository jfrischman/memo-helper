# Update Buttons Guide

Plain-language guide to every button in the app that changes something — either in
the browser preview or in the actual memo `.docx`. For code-level detail (which
`doc.tables[N]` / chart XML part each section maps to), see
[MEMO_TABLE_MAPPING.md](MEMO_TABLE_MAPPING.md).

## Quick answer: does this button touch the memo file?

| Button | Writes to the memo .docx? |
|---|---|
| Calculate exposures | No — browser preview only |
| **Update Exposures** | **Yes**, overwrites the file at Memo file path |
| **Update Fund Information** | **Yes**, overwrites the file at Memo file path |
| **Update Model Outputs** | **Yes**, overwrites the file at Memo file path |
| **Update All** | **Yes** — runs Update Exposures + Update Fund Information only |
| Export memo | No — writes a brand-new copy to `exports/`, never touches the live memo |
| Save project | No — just saves your inputs to the project file |
| Download JSON / Download positions Excel | No — standalone file download |
| Parse project files / per-fund Parse | No — only fills in the Fund Information table in the browser |

The memo must be **closed in Word** before pressing any of the three bolded Update
buttons — if Word has it open, you'll get an error telling you to close it and retry.

---

## Update Exposures

**What it changes in the memo:**
- Investment Summary box: Top 1 / Top 5 position %, Positions > 20%, Underlying Funds, Underlying Investments
- Top Positions / concentration table
- Asset Type By Fund table
- Current Portfolio bullet list (names + weight %)
- The three exposure pie charts: **security type, asset class, geography**

**Where the data comes from:**
- Each fund's uploaded workbook, using the columns you mapped (asset class, security type, geography) and the record-date NAV column for within-fund weighting (cash-like rows like money market/T-bills excluded)
- Funds with no workbook use whatever manual category mix you typed in instead
- Across funds, each fund's contribution is weighted by the **bid amount** you entered for it, not by NAV
- Any normalization rules / issuer aliases / force-split text you've entered

**Needs in place first:** an open project, a Memo file path pointing to an existing `.docx`, and at least one fund with either an uploaded+mapped workbook or a manual mix, plus a positive bid amount.

**Does NOT touch:** the Fund Summary table, model outputs, IRR/MOIC, bid bridge, or cash-flow chart.

---

## Update Fund Information

**What it changes in the memo:**
- The Fund Summary table (Fund / LP NAV / Unfunded / Commits / Invest End / Term End / Extensions / IRR / TVPI / RVPI / DPI / Leverage)
- Adds a perpetuity footnote if any fund is flagged as perpetual

**Where the data comes from:**
- The Fund Information panel's per-field values — each field is either parsed by GPT-4o from documents you uploaded (quarterly letters, audited financials, capital account statements, LPAs) via the "Parse" buttons, or typed/overridden manually and locked
- Manual overrides always win over a parsed value

**Needs in place first:** Memo file path set; fund info fields populated (via parsing or manual entry) — otherwise the table just gets resized to match your fund count with mostly blank cells.

**Does NOT touch:** exposure tables, pie charts, or model outputs.

---

## Update Model Outputs

**What it changes in the memo:**
- Investment Summary box: Bid Price, Effective Price, Base IRR/MOIC, Upside/Downside, WAL
- The header text ("Base Case Expected Net Return: X% / Y.YYx")
- Underwritten IRR/MOIC table
- Effective Bid Price bridge table
- The Combined Cashflows image (replaced with a freshly rendered image)
- The Cash Flow Profile chart (Return of Capital bars + MOIC line)

**Where the data comes from:**
- A single **Excel combined model** file (separate field from the fund workbooks — set via "Browse" next to Model file) — not the fund workbooks and not the bid amounts entered elsewhere in the app
- Each sub-section (fund returns, bid bridge, cash-flow profile) is read independently and silently skipped if not found in that workbook; only the core Bid Price/IRR/MOIC/WAL read is required and will hard-fail the whole update if missing

**Needs in place first:** Memo file path set, **and** the Excel model file path set to an existing file.

**Does NOT touch:** exposure tables, pie charts, or the Fund Summary table. This button is fully independent of the fund workbooks and bid amounts.

---

## Update All

Runs **Update Exposures** then **Update Fund Information**, back to back.

**Important:** this does *not* include Update Model Outputs. If you need the IRR/MOIC table, bid bridge, or cash-flow chart refreshed, click Update Model Outputs separately.

---

## Export memo

Produces a **new, separate file** (downloaded, and saved under the project's `exports/` folder) — it never modifies the live memo.

**What it includes:** the same exposure tables, portfolio names, and pie charts as Update Exposures. It does **not** include Fund Information or Model Outputs sections.

**What it's built from:** if the Memo file path is set and that file exists, it's used as the starting template (so anything already in that file carries over). If no Memo file path is set, it silently falls back to a hardcoded default template on this machine — only correct for the "Project Balance" project, so always set the Memo file path for any other project before exporting.

---

## Calculate exposures

Browser-only preview — recomputes the same blend as Update Exposures and renders it as charts/tables/summary in the app, but writes nothing to disk. Use this to sanity-check numbers before committing to Update Exposures. Also runs automatically as you edit funds/bids if "Live chart updates" is checked.

---

## Save project

Saves your current inputs (fund list, bid amounts, column mappings, memo/model paths, rules, fund info) to the project's own file so you can close and reopen later. Does not touch the memo. Runs automatically in the background a moment after most edits, so you rarely need to click it manually.
