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
    """
    Map the blend onto the pie's existing category slots by name, then:
      - drop any category whose value is *legitimately 0* (slice + its dPt color +
        its dLbl label are removed; surviving slices are re-indexed so they keep
        their original colors and label positions);
      - keep categories with value > 0 even if they round to 0% (label shows "(0%)").
    """
    lookup = _blend_lookup(items)
    root = etree.fromstring(xml_bytes)
    ser = root.find(".//c:ser", NS)
    if ser is None:
        return xml_bytes, ["no <c:ser>"]

    cat_cache = ser.find(".//c:cat//c:strCache", NS)
    val_cache = ser.find(".//c:val//c:numCache", NS)
    if cat_cache is None or val_cache is None:
        return xml_bytes, ["no cache"]
    cat_pts = cat_cache.findall("c:pt", NS)
    val_pts = val_cache.findall("c:pt", NS)
    dlbls = ser.find("c:dLbls", NS)

    notes: List[str] = []
    keep: List[tuple[int, str, float]] = []   # (orig_idx, base_label, share)
    for i, cp in enumerate(cat_pts):
        cv = cp.find("c:v", NS)
        base = base_label(cv.text if cv is not None else "")
        share = lookup.get(base.lower(), 0.0)
        if share > 0:
            keep.append((i, base, share))
            notes.append(f"{base}={share:.1%}")
        else:
            notes.append(f"{base}=drop(0)")

    drop_set = {i for i in range(len(cat_pts)) if i not in {oi for oi, _, _ in keep}}
    orig_to_new = {oi: new for new, (oi, _, _) in enumerate(keep)}

    # remove dropped category / value points
    for i in sorted(drop_set, reverse=True):
        cat_cache.remove(cat_pts[i])
        if i < len(val_pts):
            val_cache.remove(val_pts[i])

    # fix ptCount
    for cache in (cat_cache, val_cache):
        pc = cache.find("c:ptCount", NS)
        if pc is not None:
            pc.set("val", str(len(keep)))

    # rewrite values + labels and re-index surviving points (now in keep order)
    new_cat_pts = cat_cache.findall("c:pt", NS)
    new_val_pts = val_cache.findall("c:pt", NS)
    for new_i, (oi, base, share) in enumerate(keep):
        cp = new_cat_pts[new_i]
        cp.set("idx", str(new_i))
        cv = cp.find("c:v", NS)
        if cv is not None:
            cv.text = label_fmt(base, share)
        if new_i < len(new_val_pts):
            vp = new_val_pts[new_i]
            vp.set("idx", str(new_i))
            vv = vp.find("c:v", NS)
            if vv is not None:
                vv.text = repr(float(share))

    # drop dropped slices' colors/labels; re-index survivors to match
    for dp in list(ser.findall("c:dPt", NS)):
        ie = dp.find("c:idx", NS)
        if ie is None:
            continue
        oi = int(ie.get("val"))
        if oi in drop_set:
            ser.remove(dp)
        elif oi in orig_to_new:
            ie.set("val", str(orig_to_new[oi]))
    if dlbls is not None:
        for dl in list(dlbls.findall("c:dLbl", NS)):
            ie = dl.find("c:idx", NS)
            if ie is None:
                continue
            oi = int(ie.get("val"))
            if oi in drop_set:
                dlbls.remove(dl)
            elif oi in orig_to_new:
                ie.set("val", str(orig_to_new[oi]))

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
