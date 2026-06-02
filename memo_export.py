from __future__ import annotations

from collections import OrderedDict
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from native_charts import update_native_pies


PROJECT_BALANCE_TEMPLATE = Path(
    r"C:\Users\jfrischman\OneDrive - GCM Grosvenor\Credit\Credit Secondaries\1. Investments\1. IC Approved & Closed Transactions\Project Balance\Project Balance IC Memo v1.docx"
)


def _fmt_pct(value: float) -> str:
    return f"{value * 100:.1f}%"


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


def _set_cell_text(cell, text: str, align: Optional[WD_ALIGN_PARAGRAPH] = None):
    cell.text = str(text)
    for p in cell.paragraphs:
        if align is not None:
            p.alignment = align
        for run in p.runs:
            run.font.name = "Arial"
            run.font.size = Pt(9)


def _clear_paragraph(paragraph):
    p = paragraph._p
    for child in list(p):
        p.remove(child)


def _insert_paragraph_after(paragraph, text: str = ""):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    from docx.text.paragraph import Paragraph

    new_para = Paragraph(new_p, paragraph._parent)
    if text:
      new_para.add_run(text)
    return new_para


def _insert_table_after(paragraph, rows: int, cols: int):
    table = paragraph._parent.add_table(rows=rows, cols=cols)
    paragraph._p.addnext(table._tbl)
    return table


def _ASSET_CLASS_FOR_SECURITY_TYPE(label: str) -> str:
    mapping = {
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
    return mapping.get(label, "")


def _compute_top_concentration(items: Sequence[Dict[str, Any]]) -> Dict[str, float]:
    ordered = sorted(items, key=lambda item: float(item.get("value") or item.get("percentage") or 0), reverse=True)
    values = [float(item.get("value") or item.get("percentage") or 0) for item in ordered]
    out = {
        "top_1": sum(values[:1]),
        "top_3": sum(values[:3]),
        "top_5": sum(values[:5]),
        "top_10": sum(values[:10]),
        "remaining": max(0.0, 1.0 - sum(values[:10])),
    }
    return out


def _find_project_balance_template(project: Dict[str, Any]) -> Path:
    memo_path = project.get("memo_file_path") or ""
    if memo_path and Path(memo_path).exists():
        return Path(memo_path)
    return PROJECT_BALANCE_TEMPLATE


def _clear_and_set_table_row(row, values: Sequence[str]):
    cells = row.cells
    for idx, value in enumerate(values):
        if idx < len(cells):
            _set_cell_text(cells[idx], value)


def _set_table_if_present(doc: Document, index: int, rows: Sequence[Sequence[str]]):
    if index >= len(doc.tables):
        return
    table = doc.tables[index]
    for row_idx, values in enumerate(rows):
        if row_idx >= len(table.rows):
            break
        _clear_and_set_table_row(table.rows[row_idx], values)


def _table3_rows(result: Dict[str, Any]) -> List[List[str]]:
    categories = result.get("categories") or {}
    asset = {item["label"]: float(item["value"] or item["percentage"] or 0) for item in categories.get("asset_class", [])}
    subasset = {item["label"]: float(item["value"] or item["percentage"] or 0) for item in categories.get("sub_asset_class", [])}
    fund_profiles = result.get("fund_profiles") or []
    fund_weights = [float(f.get("weight") or 0) for f in fund_profiles]
    total_weight = sum(fund_weights) or 1.0
    fund_weight_pcts = [w / total_weight for w in fund_weights]

    def fund_pct(fund_index: int, label: str, family: str) -> str:
        profiles = fund_profiles[fund_index].get("categories") if fund_index < len(fund_profiles) else {}
        fam = profiles.get(family, [])
        mapping = {item["label"]: float(item["value"] or item["percentage"] or 0) for item in fam}
        return _fmt_pct(mapping.get(label, 0.0))

    corporate_rows = sorted(
        [label for label in subasset.keys() if label and "equity" not in label.lower() or label in subasset],
        key=lambda label: subasset.get(label, 0.0),
        reverse=True,
    )
    corp_rows = [label for label in corporate_rows if label in subasset and _ASSET_CLASS_FOR_SECURITY_TYPE(label) == "Corporate Lending"]
    special_rows = [label for label in subasset if _ASSET_CLASS_FOR_SECURITY_TYPE(label) == "Special Situations"]
    corp_top = corp_rows[:2] if corp_rows else ["Other Senior Lending", "Opportunistic / Junior"]
    special_top = special_rows[:1] if special_rows else ["Equity"]

    rows = [
        ["LP NAV", "", "100%", *[_fmt_pct(w) for w in fund_weight_pcts]],
        ["Corporate Lending", "", _fmt_pct(asset.get("Corporate Lending", 0.0)), *[fund_pct(i, "Corporate Lending", "asset_class") for i in range(len(fund_profiles))]],
    ]
    for label in corp_top:
        rows.append(["Corporate Lending", label, _fmt_pct(subasset.get(label, 0.0)), *[fund_pct(i, label, "sub_asset_class") for i in range(len(fund_profiles))]])
    rows.append(["Special Situations", "", _fmt_pct(asset.get("Special Situations", 0.0)), *[fund_pct(i, "Special Situations", "asset_class") for i in range(len(fund_profiles))]])
    for label in special_top:
        rows.append(["Special Situations", label, _fmt_pct(subasset.get(label, 0.0)), *[fund_pct(i, label, "sub_asset_class") for i in range(len(fund_profiles))]])
    return rows


def _table2_rows(result: Dict[str, Any]) -> List[List[str]]:
    top = result.get("top_concentration") or {}
    rows = [["Total", _fmt_pct(top.get("top_1", 0.0)), _fmt_pct(top.get("top_3", 0.0)), _fmt_pct(top.get("top_5", 0.0)), _fmt_pct(top.get("top_10", 0.0)), _fmt_pct(top.get("remaining", 0.0))]]
    for fund in result.get("fund_profiles") or []:
        p = _compute_top_concentration(fund.get("position_exposure") or [])
        rows.append([
            str(fund.get("fund_name") or fund.get("filename") or "Fund"),
            _fmt_pct(p["top_1"]),
            _fmt_pct(p["top_3"]),
            _fmt_pct(p["top_5"]),
            _fmt_pct(p["top_10"]),
            _fmt_pct(p["remaining"]),
        ])
    return rows


def build_memo_export(project: Dict[str, Any], result: Dict[str, Any], output_path: Path) -> Path:
    template = _find_project_balance_template(project)
    doc = Document(str(template))

    # Update exposure summary tables.
    if len(doc.tables) >= 4:
        table2 = _table2_rows(result)
        _clear_and_set_table_row(doc.tables[2].rows[1], table2[0])
        for idx, row_values in enumerate(table2[1:], start=2):
            if idx < len(doc.tables[2].rows):
                _clear_and_set_table_row(doc.tables[2].rows[idx], row_values)

        table3 = _table3_rows(result)
        for idx, row_values in enumerate(table3, start=1):
            if idx < len(doc.tables[3].rows):
                _clear_and_set_table_row(doc.tables[3].rows[idx], row_values)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))

    # Update the memo's *native* exposure pies in place (asset class / security
    # type / geography). python-docx cannot edit chart XML, so this runs on the
    # saved package, rewriting the chart caches. Preserves the memo's chart
    # styling instead of appending flat PNG images.
    update_native_pies(output_path, result.get("categories", {}) or {})
    return output_path
