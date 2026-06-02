from __future__ import annotations

from collections import OrderedDict
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from PIL import Image, ImageDraw, ImageFont


PROJECT_BALANCE_TEMPLATE = Path(
    r"C:\Users\jfrischman\OneDrive - GCM Grosvenor\Credit\Credit Secondaries\1. Investments\1. IC Approved & Closed Transactions\Project Balance\Project Balance IC Memo v1.docx"
)


ASSET_CLASS_COLORS = {
    "Corporate Lending": "#1f5f74",
    "ABS": "#8b5e3c",
    "Special Situations": "#5d7d4e",
}
GEOGRAPHY_COLORS = {
    "North America": "#4f6fb5",
    "Europe": "#c27c3d",
    "Other": "#8a4f69",
}
DEFAULT_COLORS = ["#1f5f74", "#8b5e3c", "#5d7d4e", "#4f6fb5", "#c27c3d", "#8a4f69", "#6d7a8a"]


def _font(size: int, bold: bool = False):
    candidates = [
        Path(r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf"),
        Path(r"C:\Windows\Fonts\calibri.ttf"),
    ]
    for path in candidates:
        if path.exists():
            try:
                return ImageFont.truetype(str(path), size=size)
            except Exception:
                continue
    return ImageFont.load_default()


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


def _image_fill(color: str):
    color = color.lstrip("#")
    return tuple(int(color[i : i + 2], 16) for i in (0, 2, 4))


def _resolve_color(family: str, label: str, index: int) -> str:
    if family == "geography":
        return GEOGRAPHY_COLORS.get(label, DEFAULT_COLORS[index % len(DEFAULT_COLORS)])
    asset_class = label
    if family in {"security_type", "sub_asset_class"}:
        asset_class = _ASSET_CLASS_FOR_SECURITY_TYPE(label)
    return ASSET_CLASS_COLORS.get(asset_class, DEFAULT_COLORS[index % len(DEFAULT_COLORS)])


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


def _top_items(items: Sequence[Dict[str, Any]], limit: int = 6) -> List[Dict[str, Any]]:
    ordered = sorted(items, key=lambda item: float(item.get("value") or item.get("percentage") or 0), reverse=True)
    if len(ordered) <= limit:
        return list(ordered)
    head = list(ordered[: limit - 1])
    other_value = sum(float(item.get("value") or item.get("percentage") or 0) for item in ordered[limit - 1 :])
    head.append({"label": "Other", "value": other_value, "percentage": other_value})
    return head


def _draw_donut_chart(title: str, items: Sequence[Dict[str, Any]], family: str, out_path: Path, width: int = 1400, height: int = 580):
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    title_font = _font(28, bold=True)
    body_font = _font(22)
    small_font = _font(18)
    draw.text((40, 24), title, fill="#1f2937", font=title_font)

    chart_box = (40, 90, 420, 470)
    cx = (chart_box[0] + chart_box[2]) // 2
    cy = (chart_box[1] + chart_box[3]) // 2
    radius = 160
    inner = 88
    total = sum(float(item.get("value") or item.get("percentage") or 0) for item in items) or 1.0
    angle = -90.0
    for idx, item in enumerate(items):
        value = float(item.get("value") or item.get("percentage") or 0)
        sweep = 360.0 * value / total
        color = _resolve_color(family, str(item.get("label") or ""), idx)
        draw.pieslice((cx - radius, cy - radius, cx + radius, cy + radius), start=angle, end=angle + sweep, fill=_image_fill(color))
        angle += sweep
    draw.ellipse((cx - inner, cy - inner, cx + inner, cy + inner), fill="white")

    legend_x = 500
    legend_y = 110
    legend_line_h = 60
    for idx, item in enumerate(items):
        y = legend_y + idx * legend_line_h
        color = _resolve_color(family, str(item.get("label") or ""), idx)
        draw.rectangle((legend_x, y + 6, legend_x + 24, y + 30), fill=_image_fill(color))
        label = str(item.get("label") or "")
        pct = _fmt_pct(float(item.get("value") or item.get("percentage") or 0))
        draw.text((legend_x + 36, y), label, fill="#1f2937", font=body_font)
        draw.text((legend_x + 36, y + 28), pct, fill="#6b7280", font=small_font)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(out_path)


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

    # Insert refreshed charts after the exposure summary table.
    if len(doc.tables) >= 4:
        anchor = doc.tables[3]
        after = anchor._tbl
        chart_specs = [
            ("Asset Class Exposure", result.get("categories", {}).get("asset_class", []), "asset_class"),
            ("Security Type Exposure", result.get("categories", {}).get("security_type", []), "security_type"),
            ("Geography Exposure", result.get("categories", {}).get("geography", []), "geography"),
        ]
        chart_dir = output_path.parent / "_charts"
        chart_dir.mkdir(parents=True, exist_ok=True)
        for idx, (title, items, family) in enumerate(chart_specs, start=1):
            img_path = chart_dir / f"{idx}_{family}.png"
            _draw_donut_chart(title, _top_items(items), family, img_path)
            p = OxmlElement("w:p")
            after.addnext(p)
            from docx.text.paragraph import Paragraph

            para = Paragraph(p, doc._body)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(str(img_path), width=Inches(6.5))
            after = p

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path
