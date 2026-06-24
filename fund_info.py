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
        "cas_path": None,
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


def _gpt_vision(api_key: str, images_b64: List[str], prompt: str,
                model: str = "gpt-4o") -> Dict:
    """Send page images to GPT-4o vision and return parsed JSON."""
    from openai import OpenAI
    client = OpenAI(api_key=api_key)
    content: list = [{"type": "text", "text": prompt}]
    for b64 in images_b64:
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{b64}", "detail": "high"},
        })
    r = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": _SYS},
            {"role": "user",   "content": content},
        ],
        temperature=0,
        response_format={"type": "json_object"},
    )
    return json.loads(r.choices[0].message.content)


def _pdf_to_images_b64(path: str, page_indices: List[int],
                       scale: float = 1.5) -> List[str]:
    """Render specific PDF pages to base64 PNG strings using pymupdf."""
    import base64
    import fitz
    result = []
    with fitz.open(str(path)) as doc:
        matrix = fitz.Matrix(scale, scale)
        for i in page_indices:
            if 0 <= i < len(doc):
                try:
                    pix = doc[i].get_pixmap(matrix=matrix, alpha=False)
                    result.append(base64.b64encode(pix.tobytes("png")).decode())
                except Exception:
                    pass
    return result


def _select_vision_pages(path: str, max_pages: int = 20) -> List[int]:
    """Pick the most data-rich pages for vision parsing."""
    import fitz
    kw = ["irr", "tvpi", "dpi", "rvpi", "moic", "net irr", "invested", "committed",
          "track record", "investment period", "fund summary", "recap", "snapshot",
          "distribution", "nav", "unfunded", "term", "extension"]
    with fitz.open(str(path)) as doc:
        n = len(doc)
        scored: List[Tuple[int, int]] = []
        for i in range(n):
            try:
                text = doc[i].get_text().lower()
            except Exception:
                text = ""
            score = sum(text.count(k) for k in kw)
            scored.append((i, score))

    priority = list(range(min(2, n)))
    rest = sorted(
        [(i, sc) for i, sc in scored if i not in priority],
        key=lambda x: -x[1]
    )
    pages = priority + [i for i, _ in rest[: max_pages - len(priority)]]
    return sorted(set(pages))[:max_pages]


def parse_with_vision(path: str, api_key: str, fund_name: str = "",
                      include_nav: bool = True) -> Dict:
    """Parse a fund PDF using GPT-4o vision.

    include_nav=False: skip lp_nav/unfunded/commits (for quarterly letters).
    """
    pages = _select_vision_pages(path, max_pages=20)
    images = _pdf_to_images_b64(path, pages, scale=1.5)
    if not images:
        raise ValueError("Could not render any pages from the PDF.")
    prompt_template = _COMPREHENSIVE_PROMPT if include_nav else _PERFORMANCE_PROMPT
    prompt = prompt_template.format(
        fund=fund_name or "the fund",
        ctx="[See attached page images — extract data from the tables and text visible in the images]",
    )
    r = _gpt_vision(api_key, images, prompt)
    note = r.get("as_of", "")
    srcs = r.get("sources") or {}

    def s(k):
        base = srcs.get(k, "") if isinstance(srcs, dict) else ""
        return " | ".join(filter(None, [base, note]))

    result = {
        "irr":        _f(r.get("irr"),        s("irr")),
        "tvpi":       _f(r.get("tvpi"),       s("tvpi")),
        "rvpi":       _f(r.get("rvpi"),       s("rvpi")),
        "dpi":        _f(r.get("dpi"),        s("dpi")),
        "leverage":   _f(r.get("leverage"),   s("leverage")),
        "invest_end": _f(r.get("invest_end"), s("invest_end")),
        "term_end":   _f(r.get("term_end"),   s("term_end")),
        "extensions": _f(r.get("extensions"), s("extensions"),
                         perpetuity=bool(r.get("perpetuity", False)),
                         perpetuity_note=r.get("perpetuity_note") or ""),
    }
    if include_nav:
        result["lp_nav"]   = _f(r.get("lp_nav"),   s("lp_nav"))
        result["unfunded"] = _f(r.get("unfunded"), s("unfunded"))
        result["commits"]  = _f(r.get("commits"),  s("commits"))
    return result


# ---------------------------------------------------------------------------
# Extraction functions
# ---------------------------------------------------------------------------

# Used for project-files parsing (deal overview, capital account statements)
_COMPREHENSIVE_PROMPT = """Extract all available fund information for {fund} from this document.

{ctx}

Return JSON with these exact keys (null for any not found in this document):
- "irr": since-inception net IRR as "X.X%" — look for "Net IRR", "Net Return", or equivalent
- "tvpi": Total Value to Paid-In (TVPI) as "X.XXx" — SAME as MOIC (Multiple on Invested Capital); look for "TVPI", "MOIC", "Total Multiple", "Cash-on-Cash"
- "rvpi": Remaining Value to Paid-In (RVPI) as "X.XXx" — NAV divided by paid-in capital
- "dpi": Distributed to Paid-In (DPI) as "X.XXx" — cumulative distributions divided by paid-in capital (e.g. 0.776x if 77.6%)
- "leverage": fund-level debt-to-equity ratio as "X.X%" — interest-bearing debt / net assets; exclude tax liabilities; "0.0%" if no borrowings
- "lp_nav": GCM's LP stake NAV in millions as a number — this is GCM's specific share of NAV from a capital account or deal summary, NOT total fund NAV
- "unfunded": GCM's remaining unfunded commitment in millions as a number
- "commits": GCM's total commitment to the fund in millions as a number
- "invest_end": investment period end date as "Mmm-YY" (e.g. "Feb-27")
- "term_end": fund base termination date EXCLUDING extensions as "Mmm-YY"
- "extensions": extension provisions as "N GP" or "N LPAC" or "N GP, N LPAC"
- "perpetuity": true if further extensions beyond stated ones are possible, false otherwise
- "perpetuity_note": one-sentence description if perpetuity=true, else null
- "as_of": date the performance / NAV data is as of
- "sources": object mapping each populated key to a brief note (file section + quote)"""

# Used for quarterly letters and AFS — performance metrics only, NO NAV/commitment fields
_PERFORMANCE_PROMPT = """Extract performance and lifecycle data for {fund} from this document.

{ctx}

Return JSON with these exact keys (null for any not found):
- "irr": since-inception net IRR as "X.X%" — look for "Net IRR", "Net Return", or equivalent
- "tvpi": Total Value to Paid-In (TVPI) as "X.XXx" — SAME as MOIC (Multiple on Invested Capital); look for "TVPI", "MOIC", "Total Multiple", "Cash-on-Cash", "Gross MOIC", "Net MOIC"
- "rvpi": Remaining Value to Paid-In (RVPI) as "X.XXx" — NAV / paid-in capital
- "dpi": Distributed to Paid-In (DPI) as "X.XXx" — cumulative distributions / paid-in capital (e.g. 0.776x for 77.6% DPI)
- "leverage": fund-level debt-to-equity ratio as "X.X%" — interest-bearing debt / net assets; "0.0%" if no borrowings
- "invest_end": investment period end date as "Mmm-YY"
- "term_end": fund base termination date EXCLUDING extensions as "Mmm-YY"
- "extensions": extension provisions as "N GP" or "N LPAC" or "N GP, N LPAC"
- "perpetuity": true if further extensions beyond stated ones are possible, false otherwise
- "perpetuity_note": one-sentence description if perpetuity=true, else null
- "as_of": date the performance data is as of
- "sources": object mapping each populated key to a brief note (file section + quote)

DO NOT extract lp_nav, unfunded, or commits — those come from capital account statements only.
The document may show total fund NAV or total committed capital; ignore those."""


def _extract(prompt_template: str, ctx: str, api_key: str, fund_name: str,
             include_nav: bool = True) -> Dict:
    """Run an extraction prompt and return a structured field dict.

    include_nav=False omits lp_nav/unfunded/commits from the result so that
    quarterly-letter parses never overwrite deal-overview NAV data.
    """
    r = _gpt(api_key, prompt_template.format(fund=fund_name or "the fund", ctx=ctx))
    note = r.get("as_of", "")
    srcs = r.get("sources") or {}

    def s(k):
        base = srcs.get(k, "") if isinstance(srcs, dict) else ""
        return " | ".join(filter(None, [base, note]))

    result = {
        "irr":        _f(r.get("irr"),        s("irr")),
        "tvpi":       _f(r.get("tvpi"),       s("tvpi")),
        "rvpi":       _f(r.get("rvpi"),       s("rvpi")),
        "dpi":        _f(r.get("dpi"),        s("dpi")),
        "leverage":   _f(r.get("leverage"),   s("leverage")),
        "invest_end": _f(r.get("invest_end"), s("invest_end")),
        "term_end":   _f(r.get("term_end"),   s("term_end")),
        "extensions": _f(r.get("extensions"), s("extensions"),
                         perpetuity=bool(r.get("perpetuity", False)),
                         perpetuity_note=r.get("perpetuity_note") or ""),
    }
    if include_nav:
        result["lp_nav"]   = _f(r.get("lp_nav"),   s("lp_nav"))
        result["unfunded"] = _f(r.get("unfunded"), s("unfunded"))
        result["commits"]  = _f(r.get("commits"),  s("commits"))
    return result


def _comprehensive_extract(ctx: str, api_key: str, fund_name: str) -> Dict:
    return _extract(_COMPREHENSIVE_PROMPT, ctx, api_key, fund_name, include_nav=True)


def parse_quarterly_letter(path: str, api_key: str, fund_name: str = "") -> Dict:
    """Extract performance metrics from a quarterly/semi-annual letter.

    Uses vision API when pymupdf is available. Never overwrites lp_nav,
    unfunded, or commits — those come from deal-overview docs only.
    """
    try:
        import fitz  # noqa: F401
        return parse_with_vision(path, api_key, fund_name, include_nav=False)
    except ImportError:
        pass
    pages = extract_pdf_pages(path)
    rel = _filter_pages(pages, ["net irr", "tvpi", "rvpi", "dpi", "performance",
                                 "return", "moic", "investment period", "term", "extension"])
    return _extract(_PERFORMANCE_PROMPT, _ctx(rel), api_key, fund_name, include_nav=False)


def parse_afs(path: str, api_key: str, fund_name: str = "") -> Dict:
    """Extract performance and leverage data from Audited Financial Statements.

    Uses vision API when pymupdf is available. Never overwrites lp_nav,
    unfunded, or commits — those come from deal-overview docs only.
    """
    try:
        import fitz  # noqa: F401
        return parse_with_vision(path, api_key, fund_name, include_nav=False)
    except ImportError:
        pass
    pages = extract_pdf_pages(path)
    rel = _filter_pages(pages, ["balance sheet", "net assets", "credit facility",
                                 "borrowings", "partners' capital", "irr", "tvpi",
                                 "moic", "performance", "investment period", "term", "extension"])
    return _extract(_PERFORMANCE_PROMPT, _ctx(rel), api_key, fund_name, include_nav=False)


_CAS_PROMPT = """Extract GCM's capital account data for {fund} from this Capital Account Statement.

{ctx}

Return JSON with these exact keys (null for any not found):
- "lp_nav": GCM's LP ending capital balance in millions as a number — look for "Ending Balance", "Ending Capital Balance", "Net Asset Value", "Partners' Capital", "Ending Capital Account Balance"
- "unfunded": GCM's remaining unfunded commitment in millions as a number — look for "Unfunded Commitment", "Remaining Commitment", "Undrawn Commitment"
- "commits": GCM's total commitment to the fund in millions as a number — look for "Total Commitment", "Capital Commitment", "Committed Capital"
- "as_of": the statement period end date as "Mmm-YY"
- "sources": object mapping each populated key to a brief note (section + quoted label)

Focus on GCM's specific LP account balance, not total fund NAV. Values should be in millions (e.g. 12.5 for $12.5m)."""


def parse_cas(path: str, api_key: str, fund_name: str = "") -> Dict:
    """Extract NAV (ending capital balance), unfunded, and total commitment from a CAS.
    Uses vision API when pymupdf is available since CAS docs are often scanned images."""
    def _build_result(r: Dict) -> Dict:
        note = r.get("as_of", "")
        srcs = r.get("sources") or {}
        def s(k):
            base = srcs.get(k, "") if isinstance(srcs, dict) else ""
            return " | ".join(filter(None, [base, note]))
        return {
            "lp_nav":   _f(r.get("lp_nav"),   s("lp_nav")),
            "unfunded": _f(r.get("unfunded"), s("unfunded")),
            "commits":  _f(r.get("commits"),  s("commits")),
        }

    try:
        import fitz  # noqa: F401
        pages = _select_vision_pages(path, max_pages=10)
        images = _pdf_to_images_b64(path, pages, scale=1.5)
        if images:
            r = _gpt_vision(api_key, images, _CAS_PROMPT.format(
                fund=fund_name or "the fund",
                ctx="[See attached page images — extract data visible in tables and text]",
            ))
            return _build_result(r)
    except ImportError:
        pass

    pages = extract_pdf_pages(path)
    rel = _filter_pages(pages, [
        "ending balance", "ending capital", "net asset value", "partners' capital",
        "unfunded", "commitment", "capital account", "nav",
    ], fallback_n=10)
    r = _gpt(api_key, _CAS_PROMPT.format(fund=fund_name or "the fund", ctx=_ctx(rel, 30_000)))
    return _build_result(r)


_MONTH = r"(?:January|February|March|April|May|June|July|August|September|October|November|December)"
_DATE  = rf"(?:{_MONTH}\s+\d{{1,2}},?\s+20\d{{2}}|20[2-9]\d)"

def _extract_lpa_clauses(pages: List[Tuple[int, str]]) -> str:
    """Regex-based extraction of the specific sentences in an LPA that mention
    investment period dates, fund term dates, and extension provisions.
    Returns a small focused snippet (typically < 3k chars) ready for GPT."""
    all_text = " ".join(t for _, t in pages)
    # Normalise whitespace
    all_text = re.sub(r"\s+", " ", all_text)

    sentence_patterns = [
        # Investment period + date
        rf"[^.]*[Ii]nvestment\s+[Pp]eriod[^.]*?(?:expire|terminat|end|shall\s+(?:expire|end|terminat))[^.]*?{_DATE}[^.]*\.",
        rf"[^.]*{_DATE}[^.]*[Ii]nvestment\s+[Pp]eriod[^.]*(?:expire|terminat|end)[^.]*\.",
        # Fund term + date
        rf"[^.]*(?:[Tt]erm\s+of\s+the|[Ff]und\s+[Tt]erm|[Pp]artnership\s+shall\s+terminat)[^.]*?{_DATE}[^.]*\.",
        rf"[^.]*(?:dissolv|terminat|wind\s+up)[^.]*?[Pp]artnership[^.]*?{_DATE}[^.]*\.",
        # Extensions
        r"[^.]*(?:one[-\s]year|[1-5][-\s]year|\b[Oo]ne\b|\b[Tt]wo\b|\b[Tt]hree\b)[^.]*(?:extension|extend)[^.]*(?:GP|LPAC|[Gg]eneral\s+[Pp]artner|[Aa]dvisory\s+[Cc]ommittee)[^.]*\.",
        r"[^.]*(?:LPAC|[Aa]dvisory\s+[Cc]ommittee)[^.]*(?:approv|consent|elect)[^.]*(?:extend|extension)[^.]*\.",
        r"[^.]*[Gg]eneral\s+[Pp]artner[^.]*(?:elect|option|discret)[^.]*(?:extend|extension)[^.]*\.",
        r"[^.]*[Pp]erpetuity[^.]*\.",
    ]

    seen, clauses = set(), []
    for pat in sentence_patterns:
        for m in re.finditer(pat, all_text, re.S):
            s = re.sub(r"\s+", " ", m.group()).strip()
            if len(s) > 30 and s not in seen:
                seen.add(s)
                clauses.append(s)
            if len(clauses) >= 20:
                break

    return "\n\n".join(clauses)[:8_000]


def parse_lpa(path: str, api_key: str, fund_name: str = "") -> Dict:
    """Extract investment period end, fund term, and extensions from LPA.
    Uses regex pre-extraction to pull only the relevant clauses (typically < 3k chars)
    before calling GPT — much faster than sending entire pages for long documents."""
    pages = extract_pdf_pages(path)

    # Stage 1: regex clause extraction (fast, no API cost)
    ctx = _extract_lpa_clauses(pages)

    # Stage 2: fall back to page filter if regex found nothing
    if len(ctx) < 100:
        rel = _filter_pages(pages, ["investment period", "term of the fund", "fund term",
                                     "extension", "termination", "lpac",
                                     "advisory committee", "limited partner advisory"],
                            fallback_n=10)
        ctx = _ctx(rel, 40_000)

    return _comprehensive_extract(ctx, api_key, fund_name)


def expand_paths(paths: List[str], extensions: tuple = (".pdf",),
                 max_files: int = 20) -> List[str]:
    """Expand any directory paths to individual files matching extensions.
    Searches recursively through all subdirectories."""
    result = []
    for p in paths:
        path = Path(p)
        if path.is_file() and path.suffix.lower() in extensions:
            result.append(str(path))
        elif path.is_dir():
            for f in sorted(path.rglob("*")):
                if not f.is_dir() and f.suffix.lower() in extensions:
                    result.append(str(f))
                    if len(result) >= max_files:
                        return result
    return result


def _build_context_from_paths(paths: List[str], max_chars: int = 80_000) -> tuple:
    """Build combined text context from a list of PDF paths or a directory.
    Returns (context_text, files_used_list)."""
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
            rel = _filter_pages(pages, [
                "nav", "unfunded", "commitment", "net asset", "capital account",
                "net irr", "tvpi", "rvpi", "dpi", "performance", "return", "moic",
                "balance sheet", "credit facility", "borrowing", "partners' capital",
                "investment period", "term", "extension", "lpac",
            ], fallback_n=15)
            chunk = _ctx(rel, 20_000)
            if chunk.strip():
                all_ctx += f"\n\n[File: {Path(p).name}]\n{chunk}"
                files_used.append(Path(p).name)
        except Exception:
            continue
    if not all_ctx:
        raise ValueError(
            f"Found {len(resolved)} PDF(s) but could not extract text. "
            "They may be scanned images — try uploading the specific PDFs directly.")
    return all_ctx[:max_chars], files_used


def parse_all_fields_for_fund(paths: List[str], api_key: str, fund_name: str) -> Dict:
    """Extract ALL fund information fields from any collection of PDFs/folder."""
    ctx, files_used = _build_context_from_paths(paths)
    result = _comprehensive_extract(ctx, api_key, fund_name)
    result["_files_used"] = files_used
    return result


def parse_project_files(paths: List[str], api_key: str,
                        fund_names: List[str]) -> Dict[str, Dict]:
    """Extract ALL fund information fields for every fund from a shared folder/file list.
    Runs one GPT call per fund so fund-specific data (IRR, TVPI etc.) is correctly attributed."""
    result = {}
    all_files_used = set()
    for fname in fund_names:
        extracted = parse_all_fields_for_fund(paths, api_key, fname)
        files_used = extracted.pop("_files_used", [])
        all_files_used.update(files_used)
        result[fname] = extracted
    result["_files_used"] = list(all_files_used)
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
