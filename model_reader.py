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

        irr_row = moic_row = dur_row = None
        for r in range(sec_row + 1, sec_row + 15):
            for c in range(1, 6):
                v = cells.get((r, c))
                if isinstance(v, str):
                    u = v.strip().upper()
                    if u == "NET IRR" and irr_row is None:
                        irr_row = r
                    elif u == "NET MOIC" and moic_row is None:
                        moic_row = r
                    elif ("DURATION" in u or "WAL" in u) and dur_row is None:
                        dur_row = r

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
        if dur_row and "base" in case_sec_col:
            # WAL is treated as the combined Base Case Secondary duration.
            v = cells.get((dur_row, case_sec_col["base"]))
            if isinstance(v, (int, float)):
                result["base_wal"] = float(v)

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


def read_cashflows(path: str) -> dict:
    """
    Read the Combined Cashflows table from the first sheet.

    Returns:
      {
        'header': {
          'base_irr_sec', 'base_irr_fund', 'bear_irr_sec', 'bear_irr_fund',
          'mgr_irr_sec',  'mgr_irr_fund',
          'base_moic_sec', ... (same pattern),
          'base_dur_sec',  ... (same pattern),
        },
        'rows': [
          {'date': 'Sep-25', 'base_s': float|None, 'base_f': float|None,
           'bear_s': ..., 'bear_f': ..., 'mgr_s': ..., 'mgr_f': ...},
          ...  (only rows where at least one column has abs >= 0.01)
        ]
      }
    """
    if not _HAS_OPENPYXL:
        raise RuntimeError("openpyxl is required to read Excel models.")

    import re as _re
    from datetime import datetime as _dt

    tmp = Path(tempfile.mktemp(suffix=".xlsx"))
    shutil.copy2(path, str(tmp))
    try:
        wb = openpyxl.load_workbook(str(tmp), data_only=True)
        ws = wb[wb.sheetnames[0]]  # first sheet always has bid info + combined cashflows
        cells: dict = {}
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

    def _fv(v):
        return float(v) if isinstance(v, (int, float)) else None

    # Locate quarterly date rows in column 2 (stored as datetime or "Sep-25" strings)
    _date_pat = _re.compile(r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2}$")
    date_rows: list = []
    for (r, c), v in sorted(cells.items()):
        if c != 2:
            continue
        if isinstance(v, _dt):
            date_rows.append((r, v.strftime("%b-%y")))
        elif isinstance(v, str) and _date_pat.match(v.strip()):
            date_rows.append((r, v.strip()))
    if not date_rows:
        return {}
    date_rows.sort()
    first_date_row = date_rows[0][0]

    # The S CFs / F CFs sub-header row sits directly above the first date row
    scf_row = first_date_row - 1

    # Find the case header row by scanning upward
    case_cols: dict = {}
    case_row: Optional[int] = None
    for r in range(scf_row - 1, max(1, scf_row - 15) - 1, -1):
        for c in range(1, 20):
            v = cells.get((r, c))
            if isinstance(v, str):
                u = v.strip().upper()
                if "BASE" in u and "CASE" in u and "base" not in case_cols:
                    case_cols["base"] = c
                    case_row = r
                elif "BEAR" in u and "CASE" in u and "bear" not in case_cols:
                    case_cols["bear"] = c
                elif ("MANAGER" in u or "MGR" in u) and "mgr" not in case_cols:
                    case_cols["mgr"] = c
        if case_cols:
            break

    if not case_cols or case_row is None:
        return {}

    # For each case column, find the F CFs column (first col to the right whose label starts with "F")
    # Note: "F CFs".upper() = "F CFS" which contains "S", so we cannot use "S" not in check —
    # use startswith("F") to distinguish "F CFs" from "S CFs" reliably.
    def get_f_col(start_c: int) -> int:
        for offset in range(1, 6):
            v = cells.get((scf_row, start_c + offset))
            if isinstance(v, str) and v.strip().upper().startswith("F"):
                return start_c + offset
        return start_c + 2  # fallback

    base_s = case_cols.get("base", 3)
    bear_s = case_cols.get("bear", 6)
    mgr_s  = case_cols.get("mgr",  9)
    base_f = get_f_col(base_s)
    bear_f = get_f_col(bear_s)
    mgr_f  = get_f_col(mgr_s)

    # Read Net IRR, Net MOIC, Duration header rows between case_row and scf_row
    h: dict = {}
    for r in range(case_row + 1, scf_row):
        label = cells.get((r, 2))
        if not isinstance(label, str):
            continue
        lu = label.strip().upper()
        if "NET IRR" in lu:
            h.update({
                "base_irr_sec":  _fv(cells.get((r, base_s))),
                "base_irr_fund": _fv(cells.get((r, base_f))),
                "bear_irr_sec":  _fv(cells.get((r, bear_s))),
                "bear_irr_fund": _fv(cells.get((r, bear_f))),
                "mgr_irr_sec":   _fv(cells.get((r, mgr_s))),
                "mgr_irr_fund":  _fv(cells.get((r, mgr_f))),
            })
        elif "NET MOIC" in lu:
            h.update({
                "base_moic_sec":  _fv(cells.get((r, base_s))),
                "base_moic_fund": _fv(cells.get((r, base_f))),
                "bear_moic_sec":  _fv(cells.get((r, bear_s))),
                "bear_moic_fund": _fv(cells.get((r, bear_f))),
                "mgr_moic_sec":   _fv(cells.get((r, mgr_s))),
                "mgr_moic_fund":  _fv(cells.get((r, mgr_f))),
            })
        elif "DURATION" in lu or "WAL" in lu:
            h.update({
                "base_dur_sec":  _fv(cells.get((r, base_s))),
                "base_dur_fund": _fv(cells.get((r, base_f))),
                "bear_dur_sec":  _fv(cells.get((r, bear_s))),
                "bear_dur_fund": _fv(cells.get((r, bear_f))),
                "mgr_dur_sec":   _fv(cells.get((r, mgr_s))),
                "mgr_dur_fund":  _fv(cells.get((r, mgr_f))),
            })

    # Read cashflow data rows; omit trailing rows where all six values are near-zero
    THRESHOLD = 0.01
    value_cols = [base_s, base_f, bear_s, bear_f, mgr_s, mgr_f]
    keys = ["base_s", "base_f", "bear_s", "bear_f", "mgr_s", "mgr_f"]

    cf_rows = []
    for r, date_str in date_rows:
        vals = [_fv(cells.get((r, c))) for c in value_cols]
        if any(v is not None and abs(v) >= THRESHOLD for v in vals):
            entry = {"date": date_str}
            for k, v in zip(keys, vals):
                entry[k] = v if (v is not None and abs(v) >= THRESHOLD) else None
            cf_rows.append(entry)

    return {"header": h, "rows": cf_rows}


def read_fund_returns(path: str) -> list:
    """
    Extract Base/Bear/Manager Net IRR + Net MOIC (Secondary and Fund columns)
    for the combined position and every fund block on the Model tab, in the
    same left-to-right order as the sheet's Include?/fund-name rows.

    Handles any number of fund blocks (not hardcoded to a fixed count) by
    locating blocks from the repeating BASE/BEAR/MANAGER CASE headers rather
    than fixed columns.

    Returns an ordered list, combined entry first:
      [{"name": None, "included": True, "base_irr_sec": .., "base_irr_fund": ..,
        "bear_irr_sec": .., "bear_irr_fund": .., "mgr_irr_sec": .., "mgr_irr_fund": ..,
        "base_moic_sec": .., ... (same pattern)}, ...]
    "name" is None for the combined entry (caller fills in the project name).
    Funds with Include? == 0 are still returned, with included=False, so the
    caller decides whether to render them.
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
                    cells[(cell.row, cell.column)] = cell.value
        finally:
            wb.close()
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass

    # --- Fund order + Include? flags (row with "Include?" label, names one row below) ---
    include_row: Optional[int] = None
    for (r, c), v in cells.items():
        if isinstance(v, str) and v.strip().rstrip("?").strip().upper() == "INCLUDE":
            include_row = r
            break

    fund_list: list = []  # [(name, included), ...] left-to-right, combined excluded
    if include_row is not None:
        name_row = include_row + 1
        name_cols = sorted(
            c for (r, c), v in cells.items()
            if r == name_row and isinstance(v, str) and v.strip()
        )
        if len(name_cols) > 1:
            for c in name_cols[1:]:  # first col is the combined label, skip
                name = str(cells[(name_row, c)]).strip()
                flag = cells.get((include_row, c))
                fund_list.append((name, bool(flag)))

    # --- Locate Base/Bear/Mgr x Sec/Fund column blocks (repeating across the sheet) ---
    sec_by_row: dict = {}
    for (r, c), v in cells.items():
        if isinstance(v, str) and v.strip() == "Sec":
            sec_by_row.setdefault(r, []).append(c)
    if not sec_by_row:
        return []
    sec_row = max(sec_by_row, key=lambda r: len(sec_by_row[r]))
    case_row = sec_row - 1

    def case_at(col: int) -> Optional[str]:
        for c in range(col, max(1, col - 10) - 1, -1):
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

    def fund_col_after(sec_col: int) -> Optional[int]:
        for offset in range(1, 4):
            v = cells.get((sec_row, sec_col + offset))
            if isinstance(v, str) and v.strip() == "Fund":
                return sec_col + offset
        return None

    # A new block starts whenever a case repeats (works for any number of blocks).
    blocks: list = []
    current: dict = {}
    for sc in sorted(sec_by_row[sec_row]):
        case = case_at(sc)
        if not case:
            continue
        if case in current:
            blocks.append(current)
            current = {}
        current[case] = (sc, fund_col_after(sc))
    if current:
        blocks.append(current)
    if not blocks:
        return []

    # --- Net IRR / Net MOIC rows (shared across all blocks) ---
    irr_row = moic_row = None
    for r in range(sec_row + 1, sec_row + 15):
        for c in range(1, 60):
            v = cells.get((r, c))
            if isinstance(v, str):
                u = v.strip().upper()
                if u == "NET IRR" and irr_row is None:
                    irr_row = r
                elif u == "NET MOIC" and moic_row is None:
                    moic_row = r
        if irr_row and moic_row:
            break

    def metrics_for(block: dict) -> dict:
        out = {}
        for case, (sc, fc) in block.items():
            if irr_row is not None:
                v = cells.get((irr_row, sc))
                out[f"{case}_irr_sec"] = float(v) if isinstance(v, (int, float)) else None
                if fc is not None:
                    v = cells.get((irr_row, fc))
                    out[f"{case}_irr_fund"] = float(v) if isinstance(v, (int, float)) else None
            if moic_row is not None:
                v = cells.get((moic_row, sc))
                out[f"{case}_moic_sec"] = float(v) if isinstance(v, (int, float)) else None
                if fc is not None:
                    v = cells.get((moic_row, fc))
                    out[f"{case}_moic_fund"] = float(v) if isinstance(v, (int, float)) else None
        return out

    result: list = [{"name": None, "included": True, **metrics_for(blocks[0])}]
    for (name, included), block in zip(fund_list, blocks[1:]):
        result.append({"name": name, "included": included, **metrics_for(block)})

    return result


# Fixed row order of the Effective Bid Price bridge block, starting at the
# "Effective Bid Price" anchor row.
_BID_BRIDGE_ROW_KEYS = [
    "eff_bid", "gross_bid", "rd_nav", "sec_px", "eff_bid_nof",
    "known_dist", "known_contrib", "rd_nav_nof", "sec_px_nof",
    "eff_bid_nofp", "mtm_since_rd", "rd_nav_nofp", "sec_px_nofp",
]


def read_bid_bridge(path: str) -> list:
    """
    Read the Effective Bid Price bridge block from the Model sheet:
    Effective Bid Price, Gross Bid Price, RD NAV, Sec. Px Price, Eff. Bid
    Price Net of Flows, Known Dist / Known Contribs btwn RD and TD, RD NAV /
    Sec. Px net of flows, Eff. Bid Price Net of Flows + PNL, MTM Since RD,
    RD NAV / Sec. Px net of Flows + PNL — 13 rows in fixed order.

    Layout, located relative to the "Effective Bid Price" label cell (anchor):
      anchor_row - 2: "Include?"   0/1 flag per fund column
      anchor_row - 1: fund names, one per fund column
      anchor_row:     "Effective Bid Price" (label col; combined value is the
                      next column to the right, fund values follow after that)
      anchor_row .. anchor_row+12: the 13 metrics, one row each

    Returns an ordered list, combined entry first:
      [{"name": None, "included": True, "eff_bid": .., "gross_bid": .., ...}, ...]
    "name" holds the model's own column header text (row anchor_row-1), so the
    caller can use it to (re)label the memo table's header cells. Funds with
    Include? == 0 are still returned, with included=False, so the caller
    decides whether to render them.
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
                    cells[(cell.row, cell.column)] = cell.value
        finally:
            wb.close()
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass

    anchor_row: Optional[int] = None
    anchor_col: Optional[int] = None
    for (r, c), v in cells.items():
        if isinstance(v, str) and v.strip().lower() == "effective bid price":
            anchor_row, anchor_col = r, c
            break
    if anchor_row is None or anchor_col is None:
        return []

    name_row = anchor_row - 1
    include_row = anchor_row - 2
    combined_col = anchor_col + 1

    fund_cols = []
    c = combined_col + 1
    while isinstance(cells.get((name_row, c)), str) and cells[(name_row, c)].strip():
        fund_cols.append(c)
        c += 1

    def col_values(col: int) -> list:
        vals = []
        for r in range(anchor_row, anchor_row + len(_BID_BRIDGE_ROW_KEYS)):
            v = cells.get((r, col))
            vals.append(float(v) if isinstance(v, (int, float)) else None)
        return vals

    combined_name = str(cells.get((name_row, combined_col)) or "").strip()
    result: list = [dict(zip(_BID_BRIDGE_ROW_KEYS, col_values(combined_col)),
                          name=combined_name or None, included=True)]
    for c in fund_cols:
        name = str(cells.get((name_row, c)) or "").strip()
        flag = cells.get((include_row, c))
        included = flag == 1 or flag is True
        result.append(dict(zip(_BID_BRIDGE_ROW_KEYS, col_values(c)),
                            name=name, included=included))
    return result


def read_cash_flow_profile(path: str) -> dict:
    """
    Read the Cash Flow Profile chart source table from the 'Memo' sheet.

    This table (dates + a 'Return of Capital' column + a 'MOIC' column) is
    formula-driven off the Model sheet's Base Case Combined Cashflows and the
    model's own computed Net MOIC, so it is read as-is rather than recomputed.

    Located by finding the 'Return of Capital' column header; the date column
    sits one column to its left, the MOIC column one column to its right.

    Returns {"rows": [{"date": datetime|None, "roc": float|None,
                       "moic": float|None}, ...]}
    Trailing rows with no MOIC value (the line series ran out) are dropped.
    """
    if not _HAS_OPENPYXL:
        raise RuntimeError("openpyxl is required to read Excel models.")

    from datetime import datetime as _dt

    tmp = Path(tempfile.mktemp(suffix=".xlsx"))
    shutil.copy2(path, str(tmp))
    try:
        wb = openpyxl.load_workbook(str(tmp), data_only=True)
        if "Memo" not in wb.sheetnames:
            return {"rows": []}
        ws = wb["Memo"]
        cells = {}
        for row in ws.iter_rows():
            for cell in row:
                cells[(cell.row, cell.column)] = cell.value
        wb.close()
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass

    label_row = label_col = None
    for (r, c), v in cells.items():
        if isinstance(v, str) and "return of capital" in v.strip().lower():
            label_row, label_col = r, c
            break
    if label_row is None:
        return {"rows": []}

    date_col = label_col - 1
    moic_col = label_col + 1

    rows = []
    r = label_row + 1
    while True:
        d = cells.get((r, date_col))
        roc = cells.get((r, label_col))
        moic = cells.get((r, moic_col))
        if roc is None and moic is None:
            break
        rows.append({
            "date": d if isinstance(d, _dt) else None,
            "roc": float(roc) if isinstance(roc, (int, float)) else None,
            "moic": float(moic) if isinstance(moic, (int, float)) else None,
        })
        r += 1

    while rows and rows[-1]["moic"] is None:
        rows.pop()

    return {"rows": rows}
