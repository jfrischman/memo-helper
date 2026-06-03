"""
Fund information extraction and storage.
Parses quarterly letters, AFS, and LPAs using OpenAI GPT to populate the
fund summary table (IRR, TVPI, RVPI, DPI, leverage, NAV, dates, extensions).
Manual overrides always take precedence over parsed values.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

FIELDS = ["lp_nav", "unfunded", "commits", "invest_end", "term_end",
          "extensions", "irr", "tvpi", "rvpi", "dpi", "leverage"]


def _f(value=None, source=None, override=None, **extra) -> Dict:
    d = {"value": value, "source": source, "override": override}
    d.update(extra)
    return d


def default_fund_info() -> Dict:
    fi = {
        "scale_pct": 100.0,
        "quarterly_letter_path": None,
        "afs_path": None,
        "lpa_path": None,
        "fields": {
            "lp_nav":     _f(),
            "unfunded":   _f(),
            "commits":    _f(),
            "invest_end": _f(),
            "term_end":   _f(),
            "extensions": _f(perpetuity=False, perpetuity_note=""),
            "irr":        _f(),
            "tvpi":       _f(),
            "rvpi":       _f(),
            "dpi":        _f(),
            "leverage":   _f(),
        },
    }
    return fi


def effective(field: Any) -> Any:
    """Return the effective value: override if set, else parsed value."""
    if not isinstance(field, dict):
        return field
    ov = field.get("override")
    return ov if ov is not None else field.get("value")


def fmt_money(field: Any, scale_pct: float = 100.0) -> str:
    """Format LP NAV / unfunded / commits as '$X.Xm', applying scale."""
    val = effective(field)
    if val is None:
        return ""
    try:
        num = float(str(val).replace("$", "").replace("m", "").replace(",", ""))
        return f"${num * scale_pct / 100.0:.1f}m"
    except (ValueError, TypeError):
        return str(val) if val is not None else ""


def fmt_ext(field: Any) -> str:
    """Format extensions value, appending * if perpetuity."""
    if not isinstance(field, dict):
        return str(field) if field else ""
    val = effective(field)
    if not val:
        return ""
    star = "*" if field.get("perpetuity") else ""
    return f"{val}{star}"


# ---------------------------------------------------------------------------
# PDF text extraction (pdfplumber → pypdf fallback)
# ---------------------------------------------------------------------------

def extract_pdf_pages(path: str) -> List[Tuple[int, str]]:
    """Extract (page_num, text) from a PDF. Uses pdfplumber; falls back to pypdf."""
    try:
        import pdfplumber
        pages = []
        with pdfplumber.open(str(path)) as pdf:
            for i, page in enumerate(pdf.pages):
                pages.append((i + 1, page.extract_text() or ""))
        return pages
    except ImportError:
        pass

    try:
        from pypdf import PdfReader
        reader = PdfReader(str(path))
        return [(i + 1, p.extract_text() or "") for i, p in enumerate(reader.pages)]
    except ImportError:
        raise ImportError("No PDF parser available. Run: pip install pdfplumber")


def _filter_pages(pages: List[Tuple[int, str]], keywords: List[str],
                  fallback_n: int = 25) -> List[Tuple[int, str]]:
    kw = [k.lower() for k in keywords]
    rel = [(p, t) for p, t in pages if any(k in t.lower() for k in kw)]
    return rel[:30] if rel else pages[:fallback_n]


def _ctx(pages: List[Tuple[int, str]], max_chars: int = 80_000) -> str:
    return "\n\n".join(f"[Page {p}]\n{t}" for p, t in pages)[:max_chars]


# ---------------------------------------------------------------------------
# OpenAI GPT helper
# ---------------------------------------------------------------------------

_SYS = ("You are a financial document analyst. Extract specific data accurately. "
        "Return ONLY valid JSON with the exact keys requested. "
        "Never fabricate values — use null if a value cannot be found.")


def _gpt(api_key: str, prompt: str, model: str = "gpt-4o") -> Dict:
    from openai import OpenAI
    client = OpenAI(api_key=api_key)
    r = client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": _SYS},
                  {"role": "user", "content": prompt}],
        temperature=0,
        response_format={"type": "json_object"},
    )
    return json.loads(r.choices[0].message.content)


# ---------------------------------------------------------------------------
# Extraction functions
# ---------------------------------------------------------------------------

def parse_quarterly_letter(path: str, api_key: str, fund_name: str = "") -> Dict:
    """Extract IRR, TVPI, RVPI, DPI from a quarterly letter PDF."""
    pages = extract_pdf_pages(path)
    rel = _filter_pages(pages, ["net irr", "tvpi", "rvpi", "dpi", "performance",
                                 "return", "moic", "multiple"])
    ctx = _ctx(rel)
    r = _gpt(api_key, f"""Extract performance metrics for {fund_name or 'the fund'} from this quarterly letter.

{ctx}

Return JSON with exactly these keys:
- "irr": since-inception net IRR as "X.X%" (e.g. "15.0%")
- "tvpi": Total Value to Paid-In as "X.XXx" (e.g. "1.20x")
- "rvpi": Remaining Value to Paid-In as "X.XXx"
- "dpi": Distributions to Paid-In as "X.XXx"
- "irr_source": exact quoted text and page number where IRR was found
- "tvpi_source": exact quoted text and page number where TVPI was found
- "as_of": period/date these metrics are as of (e.g. "Q4 2025" or "December 31, 2025")
Use null for any value not found.""")

    note = r.get("as_of", "")
    return {
        "irr":  _f(r.get("irr"),  f"{r.get('irr_source','')}{(' | '+note) if note else ''}"),
        "tvpi": _f(r.get("tvpi"), f"{r.get('tvpi_source','')}{(' | '+note) if note else ''}"),
        "rvpi": _f(r.get("rvpi"), note),
        "dpi":  _f(r.get("dpi"),  note),
    }


def parse_afs(path: str, api_key: str, fund_name: str = "") -> Dict:
    """Extract financial leverage from Audited Financial Statements."""
    pages = extract_pdf_pages(path)
    rel = _filter_pages(pages, ["balance sheet", "statement of assets", "statement of net assets",
                                 "credit facility", "revolving credit", "borrowings",
                                 "partners' capital", "shareholders' equity", "net assets"])
    ctx = _ctx(rel)
    r = _gpt(api_key, f"""From this Audited Financial Statement for {fund_name or 'the fund'}, calculate financial leverage.

{ctx}

Leverage = Interest-bearing liabilities (credit facility / revolving credit / fund borrowings ONLY) divided by Net Assets / Partners' Capital.
EXCLUDE: tax liabilities, management fee payables, redemptions payable, trade payables. Only include actual fund borrowings (credit lines, revolving facilities).
If no borrowings exist, leverage = 0.0%.

Return JSON:
- "leverage": as "X.X%" (e.g. "83.0%")
- "borrowings": borrowings amount used (e.g. "$50.2m credit facility")
- "net_assets": net assets / partners' capital amount used (e.g. "$60.5m")
- "source": page number and exact line items used
- "as_of": balance sheet date
Use null for leverage if cannot be determined.""")

    src = " | ".join(filter(None, [r.get("source", ""), r.get("as_of", "")]))
    return {"leverage": _f(r.get("leverage"), src)}


def parse_lpa(path: str, api_key: str, fund_name: str = "") -> Dict:
    """Extract investment period end, fund term, and extensions from LPA."""
    pages = extract_pdf_pages(path)
    rel = _filter_pages(pages, ["investment period", "term of the fund", "fund term",
                                 "extension", "termination", "dissolution",
                                 "lpac", "advisory committee", "limited partner advisory"])
    ctx = _ctx(rel)
    r = _gpt(api_key, f"""From this Limited Partnership Agreement for {fund_name or 'the fund'}, extract fund term information.

{ctx}

Return JSON:
- "invest_end": investment period end as "Mmm-YY" (e.g. "Feb-27")
- "term_end": fund base termination date EXCLUDING any extensions as "Mmm-YY"
- "extensions": extension provisions as string. Format: count + type where type is "GP" (General Partner sole discretion) or "LPAC" (LP Advisory Committee approval required). Examples: "3 GP", "2 LPAC", "1 GP, 2 LPAC"
- "perpetuity": true if the fund can be extended indefinitely / in perpetuity beyond stated extensions, false otherwise
- "perpetuity_note": if perpetuity=true, brief description of the perpetuity provision (e.g. "after stated extensions, LPAC may approve additional one-year extensions")
- "invest_end_source": exact quoted text showing investment period end
- "term_end_source": exact quoted text showing fund termination date
- "extensions_source": exact quoted text showing extension provisions
Use null for any value not found.""")

    return {
        "invest_end": _f(r.get("invest_end"), r.get("invest_end_source", "")),
        "term_end":   _f(r.get("term_end"),   r.get("term_end_source", "")),
        "extensions": _f(r.get("extensions"), r.get("extensions_source", ""),
                         perpetuity=bool(r.get("perpetuity", False)),
                         perpetuity_note=r.get("perpetuity_note") or ""),
    }


def expand_paths(paths: List[str], extensions: tuple = (".pdf",),
                 max_files: int = 20) -> List[str]:
    """Expand any directory paths to individual files matching extensions.
    Scans the directory itself and one level of subdirectories."""
    result = []
    for p in paths:
        path = Path(p)
        if path.is_file() and path.suffix.lower() in extensions:
            result.append(str(path))
        elif path.is_dir():
            # Top-level files
            for f in sorted(path.iterdir()):
                if f.is_file() and f.suffix.lower() in extensions:
                    result.append(str(f))
                    if len(result) >= max_files:
                        return result
            # One level of subdirectories
            for sub in sorted(path.iterdir()):
                if sub.is_dir():
                    for f in sorted(sub.iterdir()):
                        if f.is_file() and f.suffix.lower() in extensions:
                            result.append(str(f))
                            if len(result) >= max_files:
                                return result
    return result


def parse_project_files(paths: List[str], api_key: str,
                        fund_names: List[str]) -> Dict[str, Dict]:
    """Extract LP NAV, Unfunded, Commits for all funds from project-level overview files.
    Accepts individual PDF paths or directory paths (scanned for PDFs up to 1 level deep)."""
    # Expand any directory paths to individual PDFs
    resolved = expand_paths(paths)
    if not resolved:
        raise ValueError(
            f"No PDF files found. Searched: {', '.join(paths)}.\n"
            "Upload PDF files directly or point to a folder containing PDFs.")

    all_ctx = ""
    files_used = []
    for p in resolved:
        try:
            pages = extract_pdf_pages(str(p))
            rel = _filter_pages(pages, ["nav", "unfunded", "commitment", "net asset",
                                         "capital account", "balance"], fallback_n=8)
            chunk = _ctx(rel, 25_000)
            if chunk.strip():
                all_ctx += f"[File: {Path(p).name}]\n{chunk}\n\n"
                files_used.append(Path(p).name)
        except Exception:
            continue
    if not all_ctx:
        raise ValueError(
            f"Found {len(resolved)} PDF(s) but could not extract text from any of them. "
            "They may be scanned images — try uploading the specific overview PDF directly.")
    all_ctx = all_ctx[:80_000]

    fund_list = ", ".join(fund_names)
    r = _gpt(api_key, f"""From these deal overview documents, extract LP NAV, Unfunded Commitments, and Total Commitments for each of these funds: {fund_list}.

{all_ctx}

Return JSON where each top-level key is exactly one of the fund names listed above, and each value is an object with:
- "lp_nav": LP NAV in millions as a number (null if not found)
- "unfunded": unfunded commitments in millions as a number (null if not found)
- "commits": total commitments in millions as a number (null if not found)
- "source": brief quote showing where these figures appear
- "as_of": date the figures are as of
If a fund is not found in the documents, still include its key with all null values.""")

    result = {"_files_used": files_used}
    for fname in fund_names:
        fd = r.get(fname) or {}
        src = " | ".join(filter(None, [fd.get("source", ""), fd.get("as_of", "")]))
        result[fname] = {
            "lp_nav":   _f(fd.get("lp_nav"),   src),
            "unfunded": _f(fd.get("unfunded"), src),
            "commits":  _f(fd.get("commits"),  src),
        }
    return result


# ---------------------------------------------------------------------------
# Merge helpers
# ---------------------------------------------------------------------------

def merge_into(base: Dict, updates: Dict) -> Dict:
    """Merge extracted fields into a fund_info dict, preserving overrides."""
    fields = base.get("fields") or {}
    for key, new_field in updates.items():
        if key not in FIELDS:
            continue
        existing = fields.get(key) or _f()
        # Preserve override; update parsed value + source if we got new data
        if isinstance(new_field, dict) and new_field.get("value") is not None:
            existing["value"] = new_field["value"]
            existing["source"] = new_field.get("source") or existing.get("source")
            # Preserve extra keys (perpetuity etc.)
            for k, v in new_field.items():
                if k not in ("value", "source", "override"):
                    existing[k] = v
        fields[key] = existing
    base["fields"] = fields
    return base
