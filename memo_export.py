from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

from native_charts import update_native_pies


PROJECT_BALANCE_TEMPLATE = Path(
    r"C:\Users\jfrischman\OneDrive - GCM Grosvenor\Credit\Credit Secondaries\1. Investments\1. IC Approved & Closed Transactions\Project Balance\Project Balance IC Memo v1.docx"
)

# Ordered asset classes for the Asset Type By Fund table.
_ASSET_CLASS_ORDER = ["Corporate Lending", "ABS", "Special Situations"]

# Security type → asset class (matches the app's dropdown)
_SEC_TO_AC = {
    "Direct Lending": "Corporate Lending",
    "Other Senior Lending": "Corporate Lending",
    "Opportunistic / Junior": "Corporate Lending",
    "Distressed": "Corporate Lending",
    "Corporate Equity": "Corporate Lending",
    "CLOs": "ABS",
    "Regulatory Capital": "ABS",
    "Commercial RE (Debt)": "ABS",
    "Residential RE": "ABS",
    "Consumer": "ABS",
    "Hard Assets": "ABS",
    "Specialty Lending": "ABS",
    "Commercial RE (Equity)": "Special Situations",
    "Commercial RE (Non-Perf)": "Special Situations",
    "Equity": "Special Situations",
}


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def _fmt_pct(value: float) -> str:
    return f"{value * 100:.1f}%"


def _fmt_pct0(value: float) -> str:
    return f"{value * 100:.0f}%"


def _fmt_money(value: float) -> str:
    if value is None:
        return "-"
    try:
        num = float(value)
    except Exception:
        return "-"
    if abs(num) >= 1_000_000:
        return f"${num / 1_000_000:.1f}m"
    return f"${num:,.0f}"


def _set_cell_text(cell, text: str, align: Optional[WD_ALIGN_PARAGRAPH] = None,
                   bold: bool = False, font_pt: int = 8):
    cell.text = str(text)
    for p in cell.paragraphs:
        if align is not None:
            p.alignment = align
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(font_pt)
            if bold:
                run.font.bold = True


def _set_cell_shading(cell, fill_hex: str = "D9D9D9"):
    """Apply a background fill color to a cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:shd")):
        tcPr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def _set_cell_no_wrap(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    for el in tcPr.findall(qn("w:noWrap")):
        tcPr.remove(el)
    tcPr.append(OxmlElement("w:noWrap"))


def _format_table(table, center_from_col: int, font_pt: int = 8,
                  wide_col: Optional[int] = None, wide_width_in: float = 1.9):
    for row in table.rows:
        cells = row.cells
        for ci, cell in enumerate(cells):
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(font_pt)
            if ci >= center_from_col:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if wide_col is not None and ci == wide_col:
                cell.width = Inches(wide_width_in)
                _set_cell_no_wrap(cell)


def _find_project_balance_template(project: Dict[str, Any]) -> Path:
    memo_path = project.get("memo_file_path") or ""
    if memo_path and Path(memo_path).exists():
        return Path(memo_path)
    return PROJECT_BALANCE_TEMPLATE


# ---------------------------------------------------------------------------
# Investment Summary box (table[0]) – label-search so merged cells don't bite
# ---------------------------------------------------------------------------

def _apply_summary_stats(doc: Document, result: Dict[str, Any]) -> None:
    """Update the Investment Summary box using label search (robust to merged cells)."""
    if not doc.tables:
        return
    t = doc.tables[0]
    top = result.get("top_concentration") or {}
    positions = result.get("top_positions") or []
    fund_profiles = result.get("fund_profiles") or []

    # Build a flat map: label_text -> (row_idx, col_idx) for every cell
    label_to_pos: Dict[str, tuple] = {}
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            txt = cell.text.strip()
            if txt:
                label_to_pos[txt] = (ri, ci)

    def set_next(label: str, value: str):
        """Find the cell with `label`, update the cell immediately to its right."""
        pos = label_to_pos.get(label)
        if pos is None:
            return
        ri, ci = pos
        try:
            _set_cell_text(t.rows[ri].cells[ci + 1], value)
        except (IndexError, Exception):
            pass

    set_next("Top 1 Position", _fmt_pct(top.get("top_1", 0)))
    set_next("Top 5 Position", _fmt_pct(top.get("top_5", 0)))
    gt20 = sum(1 for p in positions if float(p.get("value", 0)) > 0.20)
    set_next("Positions >20%", str(gt20))
    set_next("Underlying Funds", str(len(fund_profiles)))
    set_next("Underlying Investments", str(len(positions)))


# ---------------------------------------------------------------------------
# Concentration table (table[2]) – dynamic rows (Total + one per fund)
# ---------------------------------------------------------------------------

def _compute_top_concentration(items: Sequence[Dict[str, Any]]) -> Dict[str, float]:
    ordered = sorted(items, key=lambda x: float(x.get("value") or x.get("percentage") or 0), reverse=True)
    vals = [float(x.get("value") or x.get("percentage") or 0) for x in ordered]
    return {
        "top_1": sum(vals[:1]),
        "top_3": sum(vals[:3]),
        "top_5": sum(vals[:5]),
        "top_10": sum(vals[:10]),
        "remaining": max(0.0, 1.0 - sum(vals[:10])),
    }


def _rebuild_concentration_table(table, result: Dict[str, Any]) -> None:
    """Replace all data rows in the concentration table to match fund count."""
    fund_profiles = result.get("fund_profiles") or []
    top = result.get("top_concentration") or {}

    # Build data: [Total row, then one row per fund]
    data_rows = [["Total", _fmt_pct(top.get("top_1", 0)), _fmt_pct(top.get("top_3", 0)),
                  _fmt_pct(top.get("top_5", 0)), _fmt_pct(top.get("top_10", 0)),
                  _fmt_pct(top.get("remaining", 0))]]
    for fp in fund_profiles:
        p = _compute_top_concentration(fp.get("position_exposure") or [])
        data_rows.append([str(fp.get("fund_name") or "Fund"),
                          _fmt_pct(p["top_1"]), _fmt_pct(p["top_3"]),
                          _fmt_pct(p["top_5"]), _fmt_pct(p["top_10"]),
                          _fmt_pct(p["remaining"])])

    # Header row is row[0]; data starts at row[1]
    tbl = table._tbl
    existing_data_rows = list(tbl.tr_lst)[1:]   # all rows after header
    n_need = len(data_rows)
    n_have = len(existing_data_rows)

    # Remove surplus rows
    for tr in existing_data_rows[n_need:]:
        tbl.remove(tr)

    # Add missing rows (clone header row structure)
    header_tr = tbl.tr_lst[0]
    for _ in range(n_need - n_have):
        import copy
        new_tr = copy.deepcopy(header_tr)
        tbl.append(new_tr)

    # Write values into rows 1..n
    for i, row_vals in enumerate(data_rows):
        row = table.rows[i + 1]
        for ci, val in enumerate(row_vals):
            if ci < len(row.cells):
                _set_cell_text(row.cells[ci], val,
                               align=WD_ALIGN_PARAGRAPH.CENTER if ci > 0 else None)
                row.cells[ci].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    _format_table(table, center_from_col=1)


# ---------------------------------------------------------------------------
# Asset Type By Fund table (table[3]) – fully dynamic, ABS support, gray headers
# ---------------------------------------------------------------------------

def _rebuild_asset_type_table(table, result: Dict[str, Any]) -> None:
    """Rebuild the Asset Type By Fund table dynamically.
    Structure: Asset Class (gray header) | Sub-Asset Class | Total | fund1 | fund2 ...
    Shows only asset classes that exist in the blend; sub-rows show present sub-types.
    """
    categories = result.get("categories") or {}
    fund_profiles = result.get("fund_profiles") or []

    asset_vals = {it["label"]: float(it.get("value") or it.get("percentage") or 0)
                  for it in categories.get("asset_class", [])}
    sub_vals = {it["label"]: float(it.get("value") or it.get("percentage") or 0)
                for it in categories.get("sub_asset_class", [])}

    fund_weights = [float(f.get("weight") or 0) for f in fund_profiles]
    total_w = sum(fund_weights) or 1.0
    fund_weight_pcts = [w / total_w for w in fund_weights]

    def fund_ac(fi: int, label: str) -> str:
        cats = (fund_profiles[fi].get("categories") or {}).get("asset_class", [])
        m = {it["label"]: float(it.get("value") or it.get("percentage") or 0) for it in cats}
        return _fmt_pct0(m.get(label, 0.0))

    def fund_sub(fi: int, label: str) -> str:
        cats = (fund_profiles[fi].get("categories") or {}).get("sub_asset_class", [])
        m = {it["label"]: float(it.get("value") or it.get("percentage") or 0) for it in cats}
        return _fmt_pct0(m.get(label, 0.0))

    n_funds = len(fund_profiles)
    fund_names = [str(fp.get("fund_name") or "Fund") for fp in fund_profiles]

    # Build row specs: (is_header, col_values, is_gray)
    # Header row: Asset Class | Sub-Asset Class | Total | fund...
    specs: List[tuple] = []  # (is_gray, values)
    specs.append((False, ["Asset Class", "Sub-Asset Class", "Total"] + fund_names))

    # LP NAV row
    specs.append((False, ["LP NAV", "", "100%"] + [_fmt_pct0(w) for w in fund_weight_pcts]))

    for ac in _ASSET_CLASS_ORDER:
        if ac not in asset_vals and ac not in {_SEC_TO_AC.get(s, "") for s in sub_vals}:
            continue
        # Gray asset class header row
        specs.append((True, [ac, "", _fmt_pct0(asset_vals.get(ac, 0.0))]
                      + [fund_ac(fi, ac) for fi in range(n_funds)]))
        # Sub-asset rows for this asset class
        subs = sorted(
            [s for s, v in sub_vals.items() if _SEC_TO_AC.get(s) == ac and v > 0],
            key=lambda s: sub_vals.get(s, 0.0), reverse=True
        )
        for sub in subs:
            specs.append((False, [ac, sub, _fmt_pct0(sub_vals.get(sub, 0.0))]
                          + [fund_sub(fi, sub) for fi in range(n_funds)]))

    # Rebuild table rows
    tbl = table._tbl
    import copy
    header_tr = list(tbl.tr_lst)[0]

    # Remove all existing rows
    for tr in list(tbl.tr_lst):
        tbl.remove(tr)

    # Add new rows
    for is_gray, vals in specs:
        new_tr = copy.deepcopy(header_tr)
        tbl.append(new_tr)

    # Now fill values
    for row_i, (is_gray, vals) in enumerate(specs):
        row = table.rows[row_i]
        while len(row.cells) < len(vals):
            # extend row if needed (shouldn't happen normally)
            break
        for ci, val in enumerate(vals):
            if ci >= len(row.cells):
                break
            align = WD_ALIGN_PARAGRAPH.CENTER if ci >= 2 else None
            _set_cell_text(row.cells[ci], val, align=align)
            row.cells[ci].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if is_gray:
                _set_cell_shading(row.cells[ci], "D9D9D9")
            if ci == 1:
                row.cells[ci].width = Inches(1.9)
                _set_cell_no_wrap(row.cells[ci])


def _apply_exposure_tables(doc: Document, result: Dict[str, Any]) -> None:
    if len(doc.tables) < 4:
        return
    _rebuild_concentration_table(doc.tables[2], result)
    _rebuild_asset_type_table(doc.tables[3], result)
    _apply_summary_stats(doc, result)


# ---------------------------------------------------------------------------
# Section registry
# ---------------------------------------------------------------------------

SECTION_UPDATERS = {
    "exposures": (_apply_exposure_tables, True),
}


def build_memo_export(project: Dict[str, Any], result: Dict[str, Any], output_path: Path) -> Path:
    template = _find_project_balance_template(project)
    doc = Document(str(template))
    _apply_exposure_tables(doc, result)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    update_native_pies(output_path, result.get("categories", {}) or {})
    return output_path


def update_sections_in_file(memo_path, result: Dict[str, Any], sections=("exposures",)) -> Path:
    memo_path = Path(memo_path)
    if not memo_path.exists():
        raise FileNotFoundError(f"Memo file not found: {memo_path}")
    doc = Document(str(memo_path))
    refresh_pies = False
    for name in sections:
        updater = SECTION_UPDATERS.get(name)
        if not updater:
            continue
        fn, also_pies = updater
        fn(doc, result)
        refresh_pies = refresh_pies or also_pies
    doc.save(str(memo_path))
    if refresh_pies:
        update_native_pies(memo_path, result.get("categories", {}) or {})
    return memo_path
