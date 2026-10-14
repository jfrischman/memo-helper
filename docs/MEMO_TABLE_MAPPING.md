# Memo Table & Chart Mapping

The app doesn't write into a semantically-tagged template — it opens the existing
`Project Balance IC Memo v1.docx` and finds each table/chart either **by fixed
position** (`doc.tables[N]`) or **by scanning cell/label text**. There is no tag in
the docx that says "this is the fund summary table" — the mapping below is the only
record of it. If the template's table order or wording ever changes, the matching
section here needs to be checked/updated too.

All of this lives in [memo_export.py](../memo_export.py) unless noted otherwise.

## Word tables

| `doc.tables` index | Section (as seen in the memo) | How it's found | Updater | Notes |
|---|---|---|---|---|
| `table[0]` | Investment Summary box (top of memo) | Fixed index; then a label→(row,col) map is built by scanning every cell's text for known labels | `_apply_model_outputs` | Also updated for entry dates, WAL, upside/downside, etc. |
| `table[1]` | Fund Summary table (Fund / LP NAV / Unf. / Commits / Invest End / Term End / Extensions / IRR / TVPI / RVPI / DPI / Leverage) | Fixed index | `_update_fund_summary_table` | Row 0 = header. Rows 1..n = one per fund. Rows are added by cloning the header row when there are more funds than existing rows — every data-row cell has its shading explicitly cleared on every rebuild (this is what the "blue row" bug was — a cloned/pre-existing row kept the header's dark-blue fill). |
| Found by title text, falls back to `table[2]` | Top Positions / Concentration table | Title search (`_find_concentration_table`) scans each table's first row for "Top Positions"/concentration wording; falls back to position 2 if not found | `_rebuild_concentration_table` | Fully wipes and rebuilds the grid each time — project-name row (`result["_project_name"]`, falls back to "Total" if unset) + one row per fund, labeled by each fund's `abbrev_name` (falls back to `fund_name` if no abbreviation is set). |
| Found by title text after a heading, falls back to `table[3]` | Asset Type By Fund table | `_find_table_after_title(doc, "Asset Type By Fund")` walks body elements looking for that heading, then returns the table right after it | `_rebuild_asset_type_table` | Dynamic row/column count — one column per fund (labeled by `abbrev_name`, falling back to `fund_name`), one row block per asset class present (with sub-rows per security type). The combined column is headed with the project name instead of "Total". `_remove_duplicate_asset_tables` cleans up leftover tables from older export runs. |
| Found by content, no fallback | Underwritten IRR / MOIC returns table | `_find_returns_table` looks for a table with 7 columns whose row 1 reads `SECONDARY / FUND` × 3 — **not** the title text, because some templates misspell it "Unerwritten" | `_apply_fund_returns_table` | Combined row first, then one row per *included* fund (per the model's left-to-right order). |
| Found by content, no fallback | Effective Bid Price bridge table | `_find_bid_bridge_table` looks for a table whose row-1, cell-0 text is exactly "effective bid price" | `_apply_bid_bridge_table` | Columns = "Balance (All)" + one per included fund; row order is fixed (`_BID_BRIDGE_ROWS`). |

## Native pie charts (chart XML parts inside the .docx zip)

Handled in [native_charts.py](../native_charts.py) — edits the chart's cached
`<c:numCache>`/`<c:strCache>` XML directly (no python-docx, no Office automation),
so Word re-renders from the same cache Excel would normally maintain.

| Chart part | Dimension | Notes |
|---|---|---|
| `word/charts/chart1.xml` | `security_type` | |
| `word/charts/chart2.xml` | `asset_class` | |
| `word/charts/chart3.xml` | `geography` | |
| `chart4`–`chart8` | *(not wired up)* | Second-cut ABS/Hard Assets pies (4–6) and cash-flow combo charts (7–8) in the template — `update_native_pies` only touches chart1–3 today. |

Colors: 3-slice palette is fixed — `09304F` (dark navy) / `4887B2` (medium blue) /
`A6A6A6` (gray) — assigned in canonical category order per dimension
(`_FAMILY_ORDER` in native_charts.py), not by insertion order.

## Section-key → function map

This is the `SECTION_UPDATERS` dict in memo_export.py, used by
`update_sections_in_file(memo_path, result, sections=(...))` — the section keys the
app.py routes pass in when they call "Update Exposures" / "Update Fund Information" /
"Update Model Outputs" / etc.

| Section key | Function | Also triggers |
|---|---|---|
| `exposures` | `_apply_exposure_tables` (concentration + asset-type tables) | Native pie chart refresh (chart1–3) |
| `portfolio_names` | `_apply_portfolio_names` | |
| `fund_info` | `_update_fund_summary_table` (`table[1]`) | |
| `model_outputs` | `_apply_model_outputs` (`table[0]`) | Header net return update + cash-flow chart image, applied post-save |
| `fund_returns` | `_apply_fund_returns_table` | |
| `bid_bridge` | `_apply_bid_bridge_table` | |

## Why this matters

Because most of these lookups are either a hardcoded index or a fragile string
match, a template change (reordering tables, rewording a heading, retyping "Fund"
as "Funds") can silently break a section — it'll either update the wrong table or
quietly find nothing and no-op. When something updates the wrong section of the
memo, or stops updating at all, this is the first place to check.
