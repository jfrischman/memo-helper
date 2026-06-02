"""
Update the memo's *native* Office charts in place by editing the chart-part XML
caches inside the .docx — no python-docx (it can't touch chart XML), no Office
automation. Word renders pie charts straight from the <c:numCache>/<c:strCache>
in word/charts/chartN.xml (these memos carry no embedded workbook), so rewriting
those caches is sufficient and preserves all designed styling (slice colors,
manual label positions, fonts).

Strategy: keep each chart's existing slice structure and map the blended exposure
onto the chart's existing categories *by name*, rewriting only the numeric value
and the label text (e.g. "Corporate Lending (94%)" -> "Corporate Lending (75%)").
Categories present in the chart but absent from the blend are set to 0%.
"""
from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Callable, Dict, List, Sequence

from lxml import etree

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS = {"c": C_NS}

# Which chart part renders which blended dimension (Project Balance memo).
DEFAULT_CHART_MAP: Dict[str, str] = {
    "word/charts/chart2.xml": "asset_class",
    "word/charts/chart1.xml": "security_type",
    "word/charts/chart3.xml": "geography",
}

_PCT_SUFFIX = re.compile(r"\s*\(\s*-?\d+(?:\.\d+)?\s*%\s*\)\s*$")


def base_label(text: str) -> str:
    """Strip a trailing ' (NN%)' from a category label."""
    return _PCT_SUFFIX.sub("", text or "").strip()


def _default_label(base: str, value: float) -> str:
    return f"{base} ({round(value * 100)}%)"


def _blend_lookup(items: Sequence[Dict[str, Any]]) -> Dict[str, float]:
    """label (lowercased) -> share. Uses 'percentage' if present, else 'value'."""
    out: Dict[str, float] = {}
    for item in items or []:
        label = base_label(str(item.get("label", "")))
        if not label:
            continue
        val = item.get("percentage")
        if val is None:
            val = item.get("value", 0.0)
        try:
            out[label.lower()] = float(val)
        except (TypeError, ValueError):
            continue
    return out


def _update_chart_xml(xml_bytes: bytes, items: Sequence[Dict[str, Any]],
                      label_fmt: Callable[[str, float], str]) -> tuple[bytes, List[str]]:
    lookup = _blend_lookup(items)
    root = etree.fromstring(xml_bytes)
    ser = root.find(".//c:ser", NS)
    if ser is None:
        return xml_bytes, ["no <c:ser>"]

    cat_pts = ser.findall(".//c:cat//c:strCache/c:pt", NS)
    val_pts = ser.findall(".//c:val//c:numCache/c:pt", NS)
    notes: List[str] = []

    for i, cat_pt in enumerate(cat_pts):
        cat_v = cat_pt.find("c:v", NS)
        base = base_label(cat_v.text if cat_v is not None else "")
        share = lookup.get(base.lower(), 0.0)
        # value
        if i < len(val_pts):
            val_v = val_pts[i].find("c:v", NS)
            if val_v is not None:
                val_v.text = repr(float(share))
        # label
        if cat_v is not None:
            cat_v.text = label_fmt(base, share)
        notes.append(f"{base}={share:.1%}")

    new_bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return new_bytes, notes


def update_native_pies(docx_path: str | Path,
                       blend_categories: Dict[str, Sequence[Dict[str, Any]]],
                       chart_map: Dict[str, str] = DEFAULT_CHART_MAP,
                       label_fmt: Callable[[str, float], str] = _default_label) -> Dict[str, List[str]]:
    """
    Rewrite the native pie caches in `docx_path` in place.
    `blend_categories` maps family name -> list of {label, value/percentage}.
    Returns {chart_part: [per-slice notes]} for logging.
    """
    docx_path = Path(docx_path)
    report: Dict[str, List[str]] = {}

    with zipfile.ZipFile(docx_path, "r") as zin:
        names = zin.namelist()
        contents = {name: zin.read(name) for name in names}

    for part, family in chart_map.items():
        if part not in contents:
            report[part] = ["MISSING part"]
            continue
        items = blend_categories.get(family) or []
        if not items:
            report[part] = [f"no blend data for '{family}' (left unchanged)"]
            continue
        new_xml, notes = _update_chart_xml(contents[part], items, label_fmt)
        contents[part] = new_xml
        report[part] = notes

    # zipfile can't edit in place; rewrite the archive preserving entry order.
    tmp_fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=str(docx_path.parent))
    import os
    os.close(tmp_fd)
    with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, contents[name])
    shutil.move(tmp_name, docx_path)
    return report
