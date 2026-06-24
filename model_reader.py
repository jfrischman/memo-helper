from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from typing import Optional

try:
    import openpyxl
    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False

_PREFERRED_SHEETS = ["Model vUpd", "Model", "Model v1", "Model vOld"]


def _load_ws(path: str):
    wb = openpyxl.load_workbook(path, data_only=True)
    for name in _PREFERRED_SHEETS:
        if name in wb.sheetnames:
            return wb, wb[name]
    # fallback: first sheet containing 'Net IRR' in first 50 rows
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows(max_row=50):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip().upper() == "NET IRR":
                    return wb, ws
    return wb, wb[wb.sheetnames[0]]


def read_model_outputs(path: str) -> dict:
    """
    Extract key deal model outputs from an Excel model file.

    Returns a dict with these keys (all float, None if not found):
      gross_bid   — Gross Bid Price as decimal (e.g. 0.83 for 83c)
      eff_bid     — Effective Bid Price as decimal
      base_irr    — Base Case Secondary Net IRR as decimal
      base_moic   — Base Case Secondary Net MOIC
      bear_irr    — Bear Case Secondary Net IRR as decimal (Downside)
      mgr_irr     — Manager Case Secondary Net IRR as decimal (Upside)
    """
    if not _HAS_OPENPYXL:
        raise RuntimeError("openpyxl is required to read Excel models.")

    tmp = Path(tempfile.mktemp(suffix=".xlsx"))
    shutil.copy2(path, str(tmp))
    try:
        wb, ws = _load_ws(str(tmp))
        cells = {}
        try:
            for row in ws.iter_rows():
                for cell in row:
                    if hasattr(cell, "row") and hasattr(cell, "column"):
                        cells[(cell.row, cell.column)] = cell.value
        finally:
            wb.close()
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass

    result: dict = {}

    # --- Bid prices ---
    for label, key in [("Gross Bid Price", "gross_bid"), ("Effective Bid Price", "eff_bid")]:
        for (r, c), v in cells.items():
            if isinstance(v, str) and v.strip() == label:
                adj = cells.get((r, c + 1))
                if isinstance(adj, (int, float)):
                    result[key] = float(adj)
                break

    # --- Returns table ---
    # Find the row with the most 'Sec' sub-headers (marks base/bear/mgr secondary columns)
    sec_by_row: dict = {}
    for (r, c), v in cells.items():
        if isinstance(v, str) and v.strip() == "Sec":
            sec_by_row.setdefault(r, []).append(c)

    sec_row: Optional[int] = None
    if sec_by_row:
        sec_row = max(sec_by_row, key=lambda r: len(sec_by_row[r]))

    if sec_row is not None:
        case_row = sec_row - 1

        def case_for_col(sc: int) -> Optional[str]:
            # Look leftward on case_row for the nearest BASE/BEAR/MANAGER header
            for c in range(sc, max(1, sc - 10) - 1, -1):
                v = cells.get((case_row, c))
                if isinstance(v, str):
                    u = v.strip().upper()
                    if "BASE" in u:
                        return "base"
                    if "BEAR" in u:
                        return "bear"
                    if "MANAGER" in u or "MGR" in u or "UPSIDE" in u:
                        return "mgr"
            return None

        case_sec_col: dict = {}
        for sc in sorted(sec_by_row[sec_row]):
            case = case_for_col(sc)
            if case and case not in case_sec_col:
                case_sec_col[case] = sc

        irr_row = moic_row = None
        for r in range(sec_row + 1, sec_row + 15):
            for c in range(1, 6):
                v = cells.get((r, c))
                if isinstance(v, str):
                    u = v.strip().upper()
                    if u == "NET IRR" and irr_row is None:
                        irr_row = r
                    elif u == "NET MOIC" and moic_row is None:
                        moic_row = r

        if irr_row:
            for case, col in case_sec_col.items():
                v = cells.get((irr_row, col))
                if isinstance(v, (int, float)):
                    result[f"{case}_irr"] = float(v)
        if moic_row:
            for case, col in case_sec_col.items():
                v = cells.get((moic_row, col))
                if isinstance(v, (int, float)):
                    result[f"{case}_moic"] = float(v)

    else:
        # Fallback: no Sec sub-headers — read directly from case header columns
        # Find rows with case headers and Net IRR/MOIC rows
        case_header_rows: dict = {}
        for (r, c), v in cells.items():
            if isinstance(v, str):
                u = v.strip().upper()
                if "BASE CASE" in u or u == "BASE":
                    case_header_rows.setdefault(r, {})["base"] = c
                elif "BEAR CASE" in u or u == "BEAR":
                    case_header_rows.setdefault(r, {})["bear"] = c
                elif "MANAGER CASE" in u or "MGR CASE" in u:
                    case_header_rows.setdefault(r, {})["mgr"] = c

        if case_header_rows:
            hdr_row = max(case_header_rows, key=lambda r: len(case_header_rows[r]))
            case_cols = case_header_rows[hdr_row]

            irr_row = moic_row = None
            for r in range(hdr_row + 1, hdr_row + 20):
                for c in range(1, 6):
                    v = cells.get((r, c))
                    if isinstance(v, str):
                        u = v.strip().upper()
                        if u == "NET IRR" and irr_row is None:
                            irr_row = r
                        elif u == "NET MOIC" and moic_row is None:
                            moic_row = r

            if irr_row:
                for case, col in case_cols.items():
                    v = cells.get((irr_row, col + 1))
                    if isinstance(v, (int, float)):
                        result[f"{case}_irr"] = float(v)
            if moic_row:
                for case, col in case_cols.items():
                    v = cells.get((moic_row, col + 1))
                    if isinstance(v, (int, float)):
                        result[f"{case}_moic"] = float(v)

    return result
