"""
Update native Office pie charts in the .docx by editing chart-part XML caches.
No python-docx, no Office automation. Word renders from <c:numCache>/<c:strCache>.

Rebuild approach: instead of mapping onto existing template slots, we rebuild the
series data to exactly match the blend categories (non-zero values). This means
new categories (e.g. ABS) are added and removed categories vanish cleanly.
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
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"c": C_NS, "a": A_NS}

DEFAULT_CHART_MAP: Dict[str, str] = {
    "word/charts/chart2.xml": "asset_class",
    "word/charts/chart1.xml": "security_type",
    "word/charts/chart3.xml": "geography",
}

# The three memo colors in order: Corporate Lending / ABS / Special Situations.
# Geography and other dims use the same three in order.
_COLOR_1 = "09304F"   # RGB(9,48,79)   – dark navy
_COLOR_2 = "4887B2"   # RGB(72,135,178) – medium blue
_COLOR_3 = "A6A6A6"   # RGB(166,166,166)– gray
_DEFAULT_COLORS = [_COLOR_1, _COLOR_2, _COLOR_3]

# Canonical order for each dimension — controls which slice appears "first" (top/left)
_FAMILY_ORDER: Dict[str, List[str]] = {
    "asset_class": ["Corporate Lending", "ABS", "Special Situations"],
    "geography": ["North America", "Europe", "Other"],
    "security_type": [
        "Direct Lending", "Other Senior Lending", "Opportunistic / Junior",
        "Distressed", "Corporate Equity",
        "CLOs", "Regulatory Capital", "Commercial RE (Debt)", "Residential RE",
        "Consumer", "Hard Assets", "Specialty Lending",
        "Commercial RE (Equity)", "Commercial RE (Non-Perf)", "Equity",
    ],
    "sub_asset_class": [
        "Direct Lending", "Other Senior Lending", "Opportunistic / Junior",
        "Distressed", "Corporate Equity",
        "CLOs", "Regulatory Capital", "Commercial RE (Debt)", "Residential RE",
        "Consumer", "Hard Assets", "Specialty Lending",
        "Commercial RE (Equity)", "Commercial RE (Non-Perf)", "Equity",
    ],
}

# Security type → asset class (so security-type slices inherit asset-class colors)
_SEC_TO_AC = {
    "direct lending": "corporate lending",
    "other senior lending": "corporate lending",
    "opportunistic / junior": "corporate lending",
    "distressed": "corporate lending",
    "corporate equity": "corporate lending",
    "traditional dl": "corporate lending",
    "clos": "abs",
    "regulatory capital": "abs",
    "commercial re (debt)": "abs",
    "residential re": "abs",
    "consumer": "abs",
    "hard assets": "abs",
    "specialty lending": "abs",
    "commercial re (equity)": "special situations",
    "commercial re (non-perf)": "special situations",
    "equity": "special situations",
    "equity*": "special situations",
}

_AC_COLOR = {
    "corporate lending": _COLOR_1,
    "abs": _COLOR_2,
    "special situations": _COLOR_3,
}

# Explicit color map for top-level asset classes and geography
_CAT_COLORS: Dict[str, str] = {
    "corporate lending": _COLOR_1,
    "abs": _COLOR_2,
    "special situations": _COLOR_3,
    "north america": _COLOR_1,
    "europe": _COLOR_2,
    "other": _COLOR_3,
}

_PCT_SUFFIX = re.compile(r"\s*\(\s*-?\d+(?:\.\d+)?\s*%\s*\)\s*$")


def base_label(text: str) -> str:
    return _PCT_SUFFIX.sub("", text or "").strip()


def _default_label(base: str, value: float) -> str:
    return f"{base} ({round(value * 100)}%)"


def _slice_color(label: str, index: int) -> str:
    key = label.lower().strip()
    # Direct hit (asset class or geography)
    if key in _CAT_COLORS:
        return _CAT_COLORS[key]
    # Security type → inherit its asset class color
    ac = _SEC_TO_AC.get(key)
    if ac:
        return _AC_COLOR.get(ac, _DEFAULT_COLORS[index % len(_DEFAULT_COLORS)])
    # Fallback: sequential
    return _DEFAULT_COLORS[index % len(_DEFAULT_COLORS)]


def _make_dpt(idx: int, color: str) -> etree._Element:
    """Build a <c:dPt> element with a solid fill color."""
    C = C_NS
    dp = etree.Element(f"{{{C}}}dPt")
    etree.SubElement(dp, f"{{{C}}}idx").set("val", str(idx))
    etree.SubElement(dp, f"{{{C}}}bubble3D").set("val", "0")
    etree.SubElement(dp, f"{{{C}}}explosion").set("val", "0")
    spPr = etree.SubElement(dp, f"{{{C_NS}}}spPr")
    solidFill = etree.SubElement(spPr, f"{{{A_NS}}}solidFill")
    etree.SubElement(solidFill, f"{{{A_NS}}}srgbClr").set("val", color)
    ln = etree.SubElement(spPr, f"{{{A_NS}}}ln")
    ln.set("w", "19050")
    ln_fill = etree.SubElement(ln, f"{{{A_NS}}}solidFill")
    etree.SubElement(ln_fill, f"{{{A_NS}}}srgbClr").set("val", "FFFFFF")
    return dp


def _make_txPr(color_hex: str, sz: int = 800) -> etree._Element:
    """Build a <c:txPr> element with a solid fill color and size."""
    txPr = etree.Element(f"{{{C_NS}}}txPr")
    etree.SubElement(txPr, f"{{{A_NS}}}bodyPr")
    etree.SubElement(txPr, f"{{{A_NS}}}lstStyle")
    p = etree.SubElement(txPr, f"{{{A_NS}}}p")
    pPr = etree.SubElement(p, f"{{{A_NS}}}pPr")
    defRPr = etree.SubElement(pPr, f"{{{A_NS}}}defRPr")
    defRPr.set("sz", str(sz))
    solidFill = etree.SubElement(defRPr, f"{{{A_NS}}}solidFill")
    etree.SubElement(solidFill, f"{{{A_NS}}}srgbClr").set("val", color_hex)
    return txPr


def _apply_label_colors(dlbls: etree._Element, keep_items: list,
                        inside_threshold: float = 0.08) -> None:
    """Set default label color to black; override large slices (inside pie) with white.
    Slices >= inside_threshold are assumed to render inside the pie."""
    if dlbls is None:
        return
    # Default text: black (for outside/small labels)
    for existing_txPr in list(dlbls.findall("c:txPr", NS)):
        dlbls.remove(existing_txPr)
    dlbls.append(_make_txPr("000000"))

    # Per-slice white override for large slices (will be inside the pie)
    for i, (base, share) in enumerate(keep_items):
        if share >= inside_threshold:
            dLbl = etree.Element(f"{{{C_NS}}}dLbl")
            etree.SubElement(dLbl, f"{{{C_NS}}}idx").set("val", str(i))
            dLbl.append(_make_txPr("FFFFFF"))
            dlbls.insert(i, dLbl)


def _update_chart_xml(xml_bytes: bytes, items: Sequence[Dict[str, Any]],
                      label_fmt: Callable[[str, float], str]) -> tuple[bytes, List[str]]:
    """
    Rebuild the pie series to exactly match the non-zero blend categories.
    Adds new categories (e.g. ABS) and removes absent ones. Assigns colors by
    category name. A >0 value that rounds to 0% is kept and labeled '(0%)'.
    """
    keep = [(base_label(str(it.get("label", ""))),
             float(it.get("percentage") or it.get("value") or 0))
            for it in (items or [])
            if float(it.get("percentage") or it.get("value") or 0) > 0]

    root = etree.fromstring(xml_bytes)
    ser = root.find(".//c:ser", NS)
    if ser is None or not keep:
        return xml_bytes, ["no <c:ser> or no non-zero items"]

    cat_cache = ser.find(".//c:cat//c:strCache", NS)
    val_cache = ser.find(".//c:val//c:numCache", NS)
    if cat_cache is None or val_cache is None:
        return xml_bytes, ["no cache nodes"]

    # Clear existing cat/val points
    for pt in list(cat_cache.findall("c:pt", NS)):
        cat_cache.remove(pt)
    for pt in list(val_cache.findall("c:pt", NS)):
        val_cache.remove(pt)

    # Update ptCount
    for cache in (cat_cache, val_cache):
        pc = cache.find("c:ptCount", NS)
        if pc is not None:
            pc.set("val", str(len(keep)))

    # Add new points
    for i, (base, share) in enumerate(keep):
        cat_pt = etree.SubElement(cat_cache, f"{{{C_NS}}}pt")
        cat_pt.set("idx", str(i))
        cat_v = etree.SubElement(cat_pt, f"{{{C_NS}}}v")
        cat_v.text = label_fmt(base, share)

        val_pt = etree.SubElement(val_cache, f"{{{C_NS}}}pt")
        val_pt.set("idx", str(i))
        val_v = etree.SubElement(val_pt, f"{{{C_NS}}}v")
        val_v.text = repr(float(share))

    # Remove all existing dPt color nodes and rebuild
    for dp in list(ser.findall("c:dPt", NS)):
        ser.remove(dp)

    # Insert dPt nodes before dLbls (or at start of ser after idx/order/spPr)
    dlbls = ser.find("c:dLbls", NS)
    insert_before = dlbls if dlbls is not None else ser.find("c:cat", NS)
    for i, (base, _) in enumerate(keep):
        dp = _make_dpt(i, _slice_color(base, i))
        if insert_before is not None:
            insert_before.addprevious(dp)
        else:
            ser.append(dp)

    # Remove old per-slice dLbl overrides then rebuild with white/black text colors
    if dlbls is not None:
        for dl in list(dlbls.findall("c:dLbl", NS)):
            dlbls.remove(dl)
        _apply_label_colors(dlbls, keep)

    notes = [f"{b}={v:.1%}" for b, v in keep]
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), notes


def update_native_pies(docx_path, blend_categories: Dict[str, Sequence[Dict[str, Any]]],
                       chart_map: Dict[str, str] = DEFAULT_CHART_MAP,
                       label_fmt: Callable[[str, float], str] = _default_label) -> Dict[str, List[str]]:
    docx_path = Path(docx_path)
    with zipfile.ZipFile(docx_path, "r") as zin:
        names = zin.namelist()
        contents = {name: zin.read(name) for name in names}

    report: Dict[str, List[str]] = {}
    for part, family in chart_map.items():
        if part not in contents:
            report[part] = ["MISSING"]
            continue
        items = list(blend_categories.get(family) or [])
        if not items:
            report[part] = [f"no data for '{family}'"]
            continue
        # Sort by canonical order for this dimension; unknown labels go to the end
        order = _FAMILY_ORDER.get(family, [])
        order_lower = {lbl.lower(): i for i, lbl in enumerate(order)}
        items = sorted(items, key=lambda it: order_lower.get(
            base_label(str(it.get("label", ""))).lower(), len(order)))
        new_xml, notes = _update_chart_xml(contents[part], items, label_fmt)
        contents[part] = new_xml
        report[part] = notes

    tmp_fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=str(docx_path.parent))
    import os
    os.close(tmp_fd)
    with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, contents[name])
    shutil.move(tmp_name, docx_path)
    return report
