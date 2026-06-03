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
from fund_info import effective, fmt_money, fmt_ext


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

def _update_current_portfolio_names(doc: Document, result: Dict[str, Any],
                                    project_name: str = "Project") -> None:
    """Replace the bold company-name line in each Current Portfolio numbered entry
    with the live exposure format: 'Name (X.X% ProjectName, X.X% Fund1, ...)'.
    Finds entries by ilvl=0 List Paragraph paragraphs after the section heading.
    Works even when content is inside tracked-change elements (<w:ins>/<w:moveTo>).
    Updates the top 8 positions; skips paragraphs with no text content."""
    from lxml import etree as _et
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    XML_SPACE = "http://www.w3.org/XML/1998/namespace"

    fund_profiles = result.get("fund_profiles") or []
    top_positions = result.get("top_positions") or []

    # Per-fund position lookup: {(fund_idx, key): fund_level_%}
    fund_pos: Dict[tuple, float] = {}
    for fi, fp in enumerate(fund_profiles):
        w = float(fp.get("weight") or 0)
        for pos in (fp.get("position_exposure") or []):
            key = str(pos.get("key") or pos.get("label") or "").lower()
            val = float(pos.get("value") or 0)
            fund_pos[(fi, key)] = (val / w) if w > 0 else 0.0

    def _fmt(pct: float) -> str:
        return f"{pct * 100:.1f}%"

    def _build_label(pos_item: Dict[str, Any]) -> str:
        name = pos_item.get("label") or ""
        key = str(pos_item.get("key") or name).lower()
        proj_pct = float(pos_item.get("value") or 0)
        parts = [f"{_fmt(proj_pct)} {project_name}"]
        for fi, fp in enumerate(fund_profiles):
            fp_pct = fund_pos.get((fi, key), 0.0)
            if fp_pct > 0:
                abbrev = str(fp.get("abbrev_name") or fp.get("fund_name") or f"Fund {fi+1}")
                parts.append(f"{_fmt(fp_pct)} {abbrev}")
        return f"{name} ({', '.join(parts)})"

    def _all_text(p_el) -> str:
        return "".join(t.text or "" for t in p_el.iter() if t.tag == f"{{{W}}}t")

    def _replace_content(p_el, new_text: str):
        """Keep pPr + bookmarks, replace everything else with a single bold run."""
        keep_tags = {f"{{{W}}}pPr", f"{{{W}}}bookmarkStart", f"{{{W}}}bookmarkEnd"}
        for child in list(p_el):
            if child.tag not in keep_tags:
                p_el.remove(child)
        r = _et.SubElement(p_el, f"{{{W}}}r")
        rPr = _et.SubElement(r, f"{{{W}}}rPr")
        _et.SubElement(rPr, f"{{{W}}}b")
        rFonts = _et.SubElement(rPr, f"{{{W}}}rFonts")
        rFonts.set(f"{{{W}}}ascii", "Calibri")
        rFonts.set(f"{{{W}}}hAnsi", "Calibri")
        sz = _et.SubElement(rPr, f"{{{W}}}sz"); sz.set(f"{{{W}}}val", "18")
        t_el = _et.SubElement(r, f"{{{W}}}t")
        t_el.text = new_text
        t_el.set(f"{{{XML_SPACE}}}space", "preserve")

    # Find Current Portfolio section
    in_section = False
    label_idx = 0
    for para in doc.paragraphs:
        if not in_section:
            if "Current Portfolio" in para.text:
                in_section = True
            continue

        if label_idx >= min(8, len(top_positions)):
            break

        # Check for ilvl=0 (company entry paragraph)
        pPr = para._p.find(f"{{{W}}}pPr")
        if pPr is None:
            continue
        numPr = pPr.find(f"{{{W}}}numPr")
        if numPr is None:
            continue
        ilvl_el = numPr.find(f"{{{W}}}ilvl")
        if ilvl_el is None or ilvl_el.get(f"{{{W}}}val") != "0":
            continue

        # Skip truly empty paragraphs
        if not _all_text(para._p).strip():
            continue

        _replace_content(para._p, _build_label(top_positions[label_idx]))
        label_idx += 1


def _update_fund_summary_table(doc: Document, result: Dict[str, Any]) -> None:
    """Update the fund summary table (table[1]) with live fund information.
    Columns: Fund | LP NAV | Unf. | Commits | Invest End | Term End |
             Extensions | IRR | TVPI | RVPI | DPI | Leverage
    Rows are added/removed to match fund count. Values come from fund_infos
    stored in result['_fund_infos'] (effective value = override ?? parsed)."""
    if len(doc.tables) < 2:
        return
    import copy
    from lxml import etree as _et
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    fund_profiles = result.get("fund_profiles") or []
    fund_infos: Dict[str, Any] = result.get("_fund_infos") or {}
    n_funds = len(fund_profiles)
    table = doc.tables[1]
    tbl = table._tbl

    # Sync row count (header = row 0, data rows = 1..n)
    header_tr = list(tbl.tr_lst)[0]
    existing = list(tbl.tr_lst)[1:]
    for tr in existing[n_funds:]:
        tbl.remove(tr)
    for _ in range(n_funds - len(existing)):
        tbl.append(copy.deepcopy(header_tr))

    # Build and fill one row per fund
    any_perp = False
    perp_notes: List[str] = []

    for fi, fp in enumerate(fund_profiles):
        fname = fp.get("fund_name") or f"Fund {fi + 1}"
        finfo = fund_infos.get(fname) or {}
        fields = finfo.get("fields") or {}
        scale = float(finfo.get("scale_pct") or 100.0)

        def eff(k):
            v = effective(fields.get(k))
            return str(v) if v is not None else ""

        ext_field = fields.get("extensions") or {}
        if ext_field.get("perpetuity"):
            any_perp = True
            note = ext_field.get("perpetuity_note") or ""
            if note and note not in perp_notes:
                perp_notes.append(note)

        vals = [
            fname,
            fmt_money(fields.get("lp_nav"),   scale),
            fmt_money(fields.get("unfunded"),  scale),
            fmt_money(fields.get("commits"),   scale),
            eff("invest_end"),
            eff("term_end"),
            fmt_ext(ext_field),
            eff("irr"),
            eff("tvpi"),
            eff("rvpi"),
            eff("dpi"),
            eff("leverage"),
        ]

        row = table.rows[fi + 1]
        for ci, val in enumerate(vals):
            if ci >= len(row.cells):
                break
            align = WD_ALIGN_PARAGRAPH.CENTER if ci > 0 else WD_ALIGN_PARAGRAPH.LEFT
            _set_cell_text(row.cells[ci], val, align=align, font_pt=8)
            row.cells[ci].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    _format_table(table, center_from_col=1)

    # Add / update perpetuity footnote paragraph immediately after table
    footnote_text = ""
    if any_perp:
        note = perp_notes[0] if perp_notes else "after stated extensions, LPAC may approve additional extensions"
        footnote_text = f"* {note}"

    # Find the paragraph right after table[1] in the body XML
    body = doc.element.body
    tbl_el = table._tbl
    tbl_idx = list(body).index(tbl_el)
    # Check if next element is already a footnote paragraph (starts with *)
    next_el = body[tbl_idx + 1] if tbl_idx + 1 < len(body) else None
    next_text = ""
    if next_el is not None and next_el.tag == f"{{{W}}}p":
        next_text = "".join(t.text or "" for t in next_el.iter() if t.tag == f"{{{W}}}t")

    if footnote_text:
        if next_text.startswith("*"):
            # Update existing footnote
            for t in next_el.findall(f".//{{{W}}}t"):
                t.text = footnote_text
        else:
            # Insert new footnote paragraph
            from lxml import etree as _et2
            fn_p = _et2.Element(f"{{{W}}}p")
            fn_pPr = _et2.SubElement(fn_p, f"{{{W}}}pPr")
            fn_rPr = _et2.SubElement(fn_p, f"{{{W}}}r")
            fn_rPrInner = _et2.SubElement(fn_rPr, f"{{{W}}}rPr")
            fn_sz = _et2.SubElement(fn_rPrInner, f"{{{W}}}sz"); fn_sz.set(f"{{{W}}}val", "10")  # 5pt = 10 half-pts
            fn_szCs = _et2.SubElement(fn_rPrInner, f"{{{W}}}szCs"); fn_szCs.set(f"{{{W}}}val", "10")
            fn_font = _et2.SubElement(fn_rPrInner, f"{{{W}}}rFonts")
            fn_font.set(f"{{{W}}}ascii", "Calibri"); fn_font.set(f"{{{W}}}hAnsi", "Calibri")
            fn_t = _et2.SubElement(fn_rPr, f"{{{W}}}t")
            fn_t.text = footnote_text
            tbl_el.addnext(fn_p)
    elif next_text.startswith("*"):
        # Remove stale footnote (no longer perpetuity)
        body.remove(next_el)


def _apply_portfolio_names(doc: Document, result: Dict[str, Any]) -> None:
    """Wrapper that reads project_name from result['_project_name'] (injected before call)."""
    _update_current_portfolio_names(doc, result, result.get("_project_name") or "Project")


SECTION_UPDATERS = {
    "exposures": (_apply_exposure_tables, True),
    "portfolio_names": (_apply_portfolio_names, False),
    "fund_info": (_update_fund_summary_table, False),
}


def build_memo_export(project: Dict[str, Any], result: Dict[str, Any], output_path: Path) -> Path:
    template = _find_project_balance_template(project)
    doc = Document(str(template))
    result["_project_name"] = project.get("project_name") or "Project"
    result["_fund_infos"] = project.get("fund_infos") or {}
    _apply_exposure_tables(doc, result)
    _apply_portfolio_names(doc, result)
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
    result["_project_name"] = (project or {}).get("project_name") or "Project"
    result["_fund_infos"] = (project or {}).get("fund_infos") or {}
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
