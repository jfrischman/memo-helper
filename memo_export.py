from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

from native_charts import update_native_pies, parse_label_colors


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
                   bold: bool = False, italic: bool = False, font_pt: int = 8):
    cell.text = str(text)
    for p in cell.paragraphs:
        if align is not None:
            p.alignment = align
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(font_pt)
            if bold:
                run.font.bold = True
            if italic:
                run.font.italic = True


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

    # Apply alignment and Calibri 9 directly via XML (bypasses table-style inheritance).
    # Use row._tr.tc_lst to get actual cells with correct grid column positions.
    # Rule: title rows (0-1) → all LEFT; data rows → even cols LEFT, odd cols CENTER.
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    from lxml import etree as _et

    def _force_align(p_el, jc_val):
        pPr = p_el.find(f"{{{W}}}pPr")
        if pPr is None:
            pPr = _et.SubElement(p_el, f"{{{W}}}pPr")
        for old in pPr.findall(f"{{{W}}}jc"):
            pPr.remove(old)
        jc = _et.SubElement(pPr, f"{{{W}}}jc")
        jc.set(f"{{{W}}}val", jc_val)

    for ri, row in enumerate(t.rows):
        # Use tc index (not grid column) — avoids gridSpan miscounting.
        # Table pattern: i=0,2,4 → labels (LEFT); i=1,3,5 → values (CENTER).
        # Title rows (ri 0-1) and any merged title cell → always LEFT.
        for i, tc in enumerate(row._tr.tc_lst):
            jc_val = "left" if (ri <= 1 or i % 2 == 0) else "center"
            for p_el in tc.findall(f"{{{W}}}p"):
                _force_align(p_el, jc_val)
                for r_el in p_el.findall(f".//{{{W}}}r"):
                    rPr = r_el.find(f"{{{W}}}rPr")
                    if rPr is None:
                        rPr = _et.SubElement(r_el, f"{{{W}}}rPr")
                    for tag in ("rFonts", "sz", "szCs"):
                        for old in rPr.findall(f"{{{W}}}{tag}"):
                            rPr.remove(old)
                    rf = _et.SubElement(rPr, f"{{{W}}}rFonts")
                    rf.set(f"{{{W}}}ascii", "Calibri")
                    rf.set(f"{{{W}}}hAnsi", "Calibri")
                    sz = _et.SubElement(rPr, f"{{{W}}}sz")
                    sz.set(f"{{{W}}}val", "18")  # 9pt = 18 half-points

    # Update specific value cells (these calls also write the new text + CENTER)
    def set_next(label: str, value: str):
        pos = label_to_pos.get(label)
        if pos is None:
            return
        ri, ci = pos
        try:
            _set_cell_text(t.rows[ri].cells[ci + 1], value,
                           align=WD_ALIGN_PARAGRAPH.CENTER, font_pt=9)
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

def _compute_top_concentration(items: Sequence[Dict[str, Any]],
                               fund_weight: float = 1.0) -> Dict[str, float]:
    """Compute top-N concentration. If fund_weight < 1, values are project-shares;
    divide by fund_weight to get fund-level percentages for the per-fund rows."""
    ordered = sorted(items, key=lambda x: float(x.get("value") or x.get("percentage") or 0), reverse=True)
    scale = (1.0 / fund_weight) if fund_weight > 0 else 1.0
    vals = [float(x.get("value") or x.get("percentage") or 0) * scale for x in ordered]
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
        w = float(fp.get("weight") or 0)
        p = _compute_top_concentration(fp.get("position_exposure") or [], fund_weight=w)
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
    """Rebuild the Asset Type By Fund table to match memo format:
    - Column header row (Asset Class | Sub-Asset Class | Total | fund...)
    - LP NAV row
    - Gray + bold asset class header row for each present asset class
    - Italic sub-rows for each present security type under that asset class
    Dynamic: rows added/removed as security types change; fund columns follow fund count.
    """
    import copy
    categories = result.get("categories") or {}
    fund_profiles = result.get("fund_profiles") or []

    asset_vals = {it["label"]: float(it.get("value") or it.get("percentage") or 0)
                  for it in categories.get("asset_class", [])}
    # Use sub_asset_class; fall back to security_type if not populated
    sub_src = categories.get("sub_asset_class") or categories.get("security_type") or []
    sub_vals = {it["label"]: float(it.get("value") or it.get("percentage") or 0)
                for it in sub_src}

    fund_weights = [float(f.get("weight") or 0) for f in fund_profiles]
    total_w = sum(fund_weights) or 1.0
    fund_weight_pcts = [w / total_w for w in fund_weights]
    n_funds = len(fund_profiles)
    fund_names = [str(fp.get("fund_name") or "Fund") for fp in fund_profiles]

    def fund_ac(fi: int, label: str) -> str:
        cats = (fund_profiles[fi].get("categories") or {}).get("asset_class", [])
        m = {it["label"]: float(it.get("value") or it.get("percentage") or 0) for it in cats}
        return _fmt_pct0(m.get(label, 0.0))

    def fund_sub(fi: int, label: str) -> str:
        fp_cats = fund_profiles[fi].get("categories") or {}
        cats = fp_cats.get("sub_asset_class") or fp_cats.get("security_type") or []
        m = {it["label"]: float(it.get("value") or it.get("percentage") or 0) for it in cats}
        return _fmt_pct0(m.get(label, 0.0))

    # Row specs: (row_type, values)
    # row_type: "header" | "lpnav" | "ac_header" | "sub"
    specs: List[tuple] = []
    specs.append(("header", ["Asset Class", "Sub-Asset Class", "Total"] + fund_names))
    specs.append(("lpnav",  ["LP NAV", "", "100%"] + [_fmt_pct0(w) for w in fund_weight_pcts]))

    present_acs = {_SEC_TO_AC.get(s, "") for s in sub_vals}
    for ac in _ASSET_CLASS_ORDER:
        if ac not in asset_vals and ac not in present_acs:
            continue
        specs.append(("ac_header", [ac, "", _fmt_pct0(asset_vals.get(ac, 0.0))]
                      + [fund_ac(fi, ac) for fi in range(n_funds)]))
        _sub_order = [
            "Direct Lending", "Other Senior Lending", "Opportunistic / Junior",
            "Distressed", "Corporate Equity",
            "CLOs", "Regulatory Capital", "Commercial RE (Debt)", "Residential RE",
            "Consumer", "Hard Assets", "Specialty Lending",
            "Commercial RE (Equity)", "Commercial RE (Non-Perf)", "Equity",
        ]
        _sub_pos = {s: i for i, s in enumerate(_sub_order)}
        subs = sorted(
            [s for s, v in sub_vals.items() if _SEC_TO_AC.get(s) == ac and v > 0],
            key=lambda s: _sub_pos.get(s, 999)
        )
        for sub in subs:
            specs.append(("sub", [ac, sub, _fmt_pct0(sub_vals.get(sub, 0.0))]
                          + [fund_sub(fi, sub) for fi in range(n_funds)]))

    # --- Sync the table's column grid to match the required fund count ---
    from lxml import etree as _lxml_et
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tbl = table._tbl
    n_cols_needed = 3 + n_funds   # Asset Class | Sub-Asset Class | Total | fund...

    tblGrid = tbl.find(f"{{{W}}}tblGrid")
    if tblGrid is not None:
        existing = tblGrid.findall(f"{{{W}}}gridCol")
        n_cols_current = len(existing)
        if n_cols_needed != n_cols_current:
            # Derive a sensible width for the fund columns from the last existing one
            fund_w = existing[-1].get(f"{{{W}}}w", "1400") if existing else "1400"
            # Remove all and rebuild with correct count
            for gc in list(existing):
                tblGrid.remove(gc)
            # Approximate widths: issuer col wide, sub-asset wide, then narrower numerics
            col_widths = ["3200", "2600"] + ["1200"] * (1 + n_funds)
            for w in col_widths[:n_cols_needed]:
                gc = _lxml_et.SubElement(tblGrid, f"{{{W}}}gridCol")
                gc.set(f"{{{W}}}w", w)

    # --- Build a prototype cell (copy tcPr from a data cell in the original row) ---
    header_tr = list(tbl.tr_lst)[0]
    orig_tcs = header_tr.findall(f"{{{W}}}tc")
    proto_tc = copy.deepcopy(orig_tcs[2]) if len(orig_tcs) > 2 else copy.deepcopy(orig_tcs[-1])
    # Clear text content from the prototype, keep only tcPr
    for p_el in list(proto_tc.findall(f"{{{W}}}p")):
        proto_tc.remove(p_el)
    empty_p = _lxml_et.SubElement(proto_tc, f"{{{W}}}p")

    def _make_tr(n_cells: int) -> "_lxml_et._Element":
        tr = _lxml_et.Element(f"{{{W}}}tr")
        for _ in range(n_cells):
            tr.append(copy.deepcopy(proto_tc))
        return tr

    # Remove all existing rows, then add fresh rows with the correct cell count
    for tr in list(tbl.tr_lst):
        tbl.remove(tr)
    for _ in specs:
        tbl.append(_make_tr(n_cols_needed))

    # --- Fill values with appropriate formatting per row type ---
    for row_i, (rtype, vals) in enumerate(specs):
        row = table.rows[row_i]
        is_ac_header = (rtype == "ac_header")
        is_sub = (rtype == "sub")
        is_col_header = (rtype == "header")

        for ci, val in enumerate(vals):
            if ci >= len(row.cells):
                break
            align = WD_ALIGN_PARAGRAPH.CENTER if ci >= 2 else None
            bold = is_ac_header or is_col_header
            italic = is_sub
            _set_cell_text(row.cells[ci], val, align=align, bold=bold, italic=italic)
            row.cells[ci].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if is_ac_header:
                _set_cell_shading(row.cells[ci], "D9D9D9")
        if ci == 1 and ci < len(row.cells):
            row.cells[1].width = Inches(1.9)
            _set_cell_no_wrap(row.cells[1])


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
    lc = parse_label_colors(project.get("label_colors") or "")
    update_native_pies(output_path, result.get("categories", {}) or {}, label_colors=lc)
    return output_path


def update_sections_in_file(memo_path, result: Dict[str, Any], sections=("exposures",),
                            project: Dict[str, Any] = None) -> Path:
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
        lc = parse_label_colors((project or {}).get("label_colors") or "")
        update_native_pies(memo_path, result.get("categories", {}) or {}, label_colors=lc)
    return memo_path
