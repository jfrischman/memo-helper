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


def _make_label_txPr(color_hex: str) -> etree._Element:
    """<c:txPr> with a solid fill text color (for dLbls default or per-slice override)."""
    txPr = etree.Element(f"{{{C_NS}}}txPr")
    etree.SubElement(txPr, f"{{{A_NS}}}bodyPr")
    etree.SubElement(txPr, f"{{{A_NS}}}lstStyle")
    p = etree.SubElement(txPr, f"{{{A_NS}}}p")
    pPr = etree.SubElement(p, f"{{{A_NS}}}pPr")
    defRPr = etree.SubElement(pPr, f"{{{A_NS}}}defRPr")
    defRPr.set("sz", "800")
    solidFill = etree.SubElement(defRPr, f"{{{A_NS}}}solidFill")
    etree.SubElement(solidFill, f"{{{A_NS}}}srgbClr").set("val", color_hex)
    return txPr


def _make_dLbl(idx: int, color_hex: str = "FFFFFF") -> etree._Element:
    """Per-slice dLbl with explicit color and show-flags matching dLbls defaults.
    show-flags MUST be included — without them Word shows series name ('Series1...')."""
    C = C_NS
    dl = etree.Element(f"{{{C}}}dLbl")
    etree.SubElement(dl, f"{{{C}}}idx").set("val", str(idx))
    etree.SubElement(dl, f"{{{C}}}showLegendKey").set("val", "0")
    etree.SubElement(dl, f"{{{C}}}showVal").set("val", "1")
    etree.SubElement(dl, f"{{{C}}}showCatName").set("val", "0")
    etree.SubElement(dl, f"{{{C}}}showSerName").set("val", "0")
    etree.SubElement(dl, f"{{{C}}}showPercent").set("val", "0")
    etree.SubElement(dl, f"{{{C}}}showBubbleSize").set("val", "0")
    dl.append(_make_label_txPr(color_hex))
    return dl


def parse_label_colors(text: str) -> Dict[str, str]:
    """Parse 'Category = white/black' lines into {canonical_label: hex}.
    Case-insensitive. Returns {} if text is empty."""
    out: Dict[str, str] = {}
    for line in (text or "").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = re.split(r"\s*=\s*", line, maxsplit=1)
        if len(parts) != 2:
            continue
        label = parts[0].strip().lower()
        color = parts[1].strip().lower()
        if label:
            out[label] = "000000" if color == "black" else "FFFFFF"
    return out


def _update_chart_xml(xml_bytes: bytes, items: Sequence[Dict[str, Any]],
                      label_fmt: Callable[[str, float], str],
                      label_colors: Dict[str, str] = None) -> tuple[bytes, List[str]]:
    """
    Rebuild the pie series to exactly match the non-zero blend categories.
    Adds new categories (e.g. ABS) and removes absent ones. Assigns colors by
    category name. A >0 value that rounds to 0% is kept and labeled '(0%)'.
    """
    _label_colors_map = label_colors or {}
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

    # Update dLbls: disable leader lines, set default color to white, apply
    # per-slice overrides from label_colors (injected via closure below).
    if dlbls is not None:
        for dl in list(dlbls.findall("c:dLbl", NS)):
            dlbls.remove(dl)
        for ll in list(dlbls.findall("c:showLeaderLines", NS)):
            dlbls.remove(ll)
        etree.SubElement(dlbls, f"{{{C_NS}}}showLeaderLines").set("val", "0")
        # Default = white (unspecified categories stay white)
        for old_txPr in list(dlbls.findall("c:txPr", NS)):
            dlbls.remove(old_txPr)
        dlbls.append(_make_label_txPr("FFFFFF"))
        # Per-slice overrides from label_colors map
        for i, (base, _) in enumerate(keep):
            color = _label_colors_map.get(base.lower(), "FFFFFF")
            if color != "FFFFFF":   # only need dLbl if overriding default (white)
                dlbls.insert(i, _make_dLbl(i, color))

    notes = [f"{b}={v:.1%}" for b, v in keep]
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), notes


def _build_chart_map_from_document(contents: Dict[str, bytes]) -> Dict[str, str]:
    """
    Identify the three exposure pie charts by their visual left-to-right position.

    Reads the <wp:positionH> offset from each floating pie chart's <wp:anchor>
    element, groups charts into left/middle/right horizontal buckets, and maps
    them to asset_class/security_type/geography.

    When a document contains duplicate pie chart sets (e.g. a secondary "legend"
    set stacked above the main charts), multiple charts share the same H bucket.
    Within each bucket the chart with the largest V offset is selected — it sits
    furthest below its anchor paragraph and corresponds to the primary visible row.

    Falls back to DEFAULT_CHART_MAP if fewer than 3 pie charts are found or the
    horizontal positions cannot be parsed.
    """
    FAMILY_ORDER = ["asset_class", "security_type", "geography"]
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

    # Build rId → chart part path from word/_rels/document.xml.rels
    rels_bytes = contents.get("word/_rels/document.xml.rels")
    if not rels_bytes:
        return dict(DEFAULT_CHART_MAP)
    try:
        rels_root = etree.fromstring(rels_bytes)
    except Exception:
        return dict(DEFAULT_CHART_MAP)

    rid_to_part: Dict[str, str] = {}
    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        if "chart" not in rel.get("Type", "").lower():
            continue
        rid = rel.get("Id", "")
        target = rel.get("Target", "")
        if not rid or not target:
            continue
        if target.startswith("../"):
            target = "word/" + target[3:]
        elif not target.startswith("word/"):
            target = "word/" + target
        rid_to_part[rid] = target

    doc_bytes = contents.get("word/document.xml")
    if not doc_bytes:
        return dict(DEFAULT_CHART_MAP)
    try:
        doc_root = etree.fromstring(doc_bytes)
    except Exception:
        return dict(DEFAULT_CHART_MAP)

    C_CHART_TAG = f"{{{C_NS}}}chart"

    # Collect (H_norm, V_off, part) for every pie chart
    chart_positions: List[tuple] = []
    seen: set = set()

    for chart_ref in doc_root.iter(C_CHART_TAG):
        rid = chart_ref.get(f"{{{R_NS}}}id")
        if not rid or rid not in rid_to_part:
            continue
        part = rid_to_part[rid]
        if part in seen or part not in contents:
            continue
        try:
            chart_root = etree.fromstring(contents[part])
        except Exception:
            continue
        if not chart_root.findall(f".//{{{C_NS}}}pieChart"):
            continue
        seen.add(part)

        # Walk up to find the wp:anchor (floating) or wp:inline element
        anchor = None
        parent = chart_ref.getparent()
        while parent is not None:
            if parent.tag == f"{{{WP_NS}}}anchor":
                anchor = parent
                break
            if parent.tag == f"{{{WP_NS}}}inline":
                break
            parent = parent.getparent()

        H_norm = 0
        V_off = 0
        if anchor is not None:
            posH = anchor.find(f"{{{WP_NS}}}positionH")
            posV = anchor.find(f"{{{WP_NS}}}positionV")
            if posH is not None:
                h_offset = posH.find(f"{{{WP_NS}}}posOffset")
                h_align = posH.find(f"{{{WP_NS}}}align")
                if h_offset is not None and h_offset.text:
                    try:
                        H_norm = int(h_offset.text)
                    except ValueError:
                        pass
                elif h_align is not None:
                    # 'left' → 0, 'center' → middle of page, 'right' → far right
                    align = (h_align.text or "").lower()
                    H_norm = 0 if align == "left" else (3000000 if align == "center" else 6000000)
            if posV is not None:
                v_offset = posV.find(f"{{{WP_NS}}}posOffset")
                if v_offset is not None and v_offset.text:
                    try:
                        V_off = int(v_offset.text)
                    except ValueError:
                        pass

        chart_positions.append((H_norm, V_off, part))

    if len(chart_positions) < 3:
        return dict(DEFAULT_CHART_MAP)

    # Sort by H to find the full horizontal range
    chart_positions.sort(key=lambda x: x[0])
    min_H = chart_positions[0][0]
    max_H = chart_positions[-1][0]

    if max_H == min_H:
        # All at same H (e.g. vertical layout) — fall back
        return dict(DEFAULT_CHART_MAP)

    # Divide H range into 3 equal buckets → left / middle / right
    bucket_size = (max_H - min_H) / 3.0
    buckets: List[List[tuple]] = [[], [], []]
    for H, V, part in chart_positions:
        idx = min(2, int((H - min_H) / bucket_size))
        buckets[idx].append((H, V, part))

    # Map ALL charts in each H bucket to the corresponding family.
    # Both the primary (visible) and secondary (duplicate) chart in a bucket
    # get the same data, so whichever one is visible in the memo is correct.
    result: Dict[str, str] = {}
    for i, bucket in enumerate(buckets):
        if not bucket:
            return dict(DEFAULT_CHART_MAP)
        for _, _, part in bucket:
            result[part] = FAMILY_ORDER[i]

    return result


def update_native_pies(docx_path, blend_categories: Dict[str, Sequence[Dict[str, Any]]],
                       chart_map: Dict[str, str] = None,
                       label_fmt: Callable[[str, float], str] = _default_label,
                       label_colors: Dict[str, str] = None) -> Dict[str, List[str]]:
    docx_path = Path(docx_path)
    with zipfile.ZipFile(docx_path, "r") as zin:
        names = zin.namelist()
        contents = {name: zin.read(name) for name in names}

    # Discover which chart XML file corresponds to which exposure family from
    # the document itself, rather than relying on hardcoded chart1/2/3 positions.
    if chart_map is None:
        chart_map = _build_chart_map_from_document(contents)

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
        new_xml, notes = _update_chart_xml(contents[part], items, label_fmt, label_colors)
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
