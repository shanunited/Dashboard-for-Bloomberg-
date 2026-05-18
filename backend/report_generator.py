"""Daily Stock Scoring PDF report generator."""
from __future__ import annotations

import io
from typing import Any

import pandas as pd
from openpyxl import load_workbook as xl_load
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT

PAGE_W, PAGE_H = A4          # 595.28 x 841.89 pts
MARGIN = 14 * mm
FOOTER_H = 13 * mm
USABLE_W = PAGE_W - 2 * MARGIN

# ── Color palette ─────────────────────────────────────────────────────────────
CN    = colors.HexColor("#0a1628")   # deep navy (headers)
CN2   = colors.HexColor("#0e2038")   # lighter navy (sub-headers)
CGOLD = colors.HexColor("#f0a500")   # gold
CGND  = colors.HexColor("#0c2a10")   # green dark bg (top cards)
CGNA  = colors.HexColor("#27ae60")   # green accent
CRDD  = colors.HexColor("#2a0c0c")   # red dark bg (bottom cards)
CRDA  = colors.HexColor("#c0392b")   # red accent
CTXT  = colors.HexColor("#dbe7f2")   # main text
CMUT  = colors.HexColor("#8ea3b8")   # muted text
CBG   = colors.HexColor("#09111b")   # card body bg
CBDR  = colors.HexColor("#243447")   # border
CTICK = colors.HexColor("#4ade80")   # tick green
CCRSS = colors.HexColor("#f87171")   # cross red
CW    = colors.white


# ── Paragraph styles ──────────────────────────────────────────────────────────
def _ps(name: str, **kw) -> ParagraphStyle:
    p = ParagraphStyle(name)
    for k, v in kw.items():
        setattr(p, k, v)
    return p


ST_CVRTITLE = _ps("cvrtitle", fontName="Helvetica-Bold", fontSize=24, textColor=CW,    alignment=TA_CENTER, spaceAfter=6)
ST_CVRSUB   = _ps("cvrsub",   fontName="Helvetica",      fontSize=13, textColor=CGOLD, alignment=TA_CENTER, spaceAfter=4)
ST_IDXHDR   = _ps("idxhdr",   fontName="Helvetica-Bold", fontSize=13, textColor=CW,    alignment=TA_CENTER)
ST_SECHDR   = _ps("sechdr",   fontName="Helvetica-Bold", fontSize=9,  textColor=CGOLD, alignment=TA_LEFT)
ST_COMPANY  = _ps("company",  fontName="Helvetica-Bold", fontSize=10, textColor=CW,    alignment=TA_LEFT)
ST_CMP      = _ps("cmp_s",    fontName="Helvetica-Bold", fontSize=10, textColor=CGOLD, alignment=TA_RIGHT)
ST_SCLBL    = _ps("sclbl",    fontName="Helvetica",      fontSize=7,  textColor=CMUT,  alignment=TA_CENTER)
ST_SCPOS    = _ps("scpos",    fontName="Helvetica-Bold", fontSize=13, textColor=CTICK, alignment=TA_CENTER)
ST_SCNEG    = _ps("scneg",    fontName="Helvetica-Bold", fontSize=13, textColor=CCRSS, alignment=TA_CENTER)
ST_SCNEU    = _ps("scneu",    fontName="Helvetica-Bold", fontSize=13, textColor=CGOLD, alignment=TA_CENTER)
ST_INDHDR   = _ps("indhdr",   fontName="Helvetica-Bold", fontSize=8,  textColor=CMUT,  alignment=TA_LEFT)
ST_TICK     = _ps("tick_s",   fontName="Helvetica",      fontSize=8,  textColor=CTICK, alignment=TA_LEFT, leading=11)
ST_CROSS    = _ps("cross_s",  fontName="Helvetica",      fontSize=8,  textColor=CCRSS, alignment=TA_LEFT, leading=11)
ST_MVAL     = _ps("mval",     fontName="Helvetica",      fontSize=8,  textColor=CTXT,  alignment=TA_LEFT, leading=12)
ST_NARR     = _ps("narr",     fontName="Helvetica",      fontSize=8,  textColor=CTXT,  alignment=TA_JUSTIFY, leading=12)
ST_LEGEND   = _ps("legend",   fontName="Helvetica",      fontSize=9,  textColor=CTXT,  alignment=TA_LEFT, leading=13)
ST_LEGKEY   = _ps("legkey",   fontName="Helvetica-Bold", fontSize=9,  textColor=CGOLD, alignment=TA_LEFT, leading=13)
ST_COVBODY  = _ps("covbody",  fontName="Helvetica",      fontSize=10, textColor=CTXT,  alignment=TA_LEFT, leading=14)


# ── Column alias map for the report ──────────────────────────────────────────
REPORT_COLS: dict[str, list[str]] = {
    # Core
    "company_name":   ["Company Name"],
    "index_clean":    ["Index_Clean"],
    "cmp":            ["CMP"],
    "fs":             ["FS", "FUNDAMENTAL SCORE", "Fundamental Score"],
    "ts":             ["TS", "TECHNICAL SCORE", "Technical Score"],
    "total":          ["Total"],
    # Technical score indicators (binary)
    "rsir":           ["RSIR"],
    "wma30r":         ["30war"],
    "weekly_rs":      ["52 WR"],
    "daily_rs":       ["52 DR"],
    # Fundamental score indicators (binary)
    "sales_yoy":      ["SALES-R(YOY)"],
    "sales_qoq":      ["SALES-R(QOQ)"],
    "rocer":          ["ROCER"],
    "roer":           ["ROER"],
    "de_ratio":       ["D/E Ratio"],
    "fcf_r":          ["FCF-R"],
    "eps_s":          ["EPS(S)"],
    "roce5y_score":   ["ROCE5Y CAGR"],
    "pat_margin":     ["PAT Margin"],
    "ebitda_margin":  ["EBITA Margin"],
    # Raw metric values
    "rsi_val":        ["RSI.1"],
    "yoy_pct":        ["YOY%"],
    "qoq_pct":        ["QOQ%"],
    "de_raw":         ["D/E"],
    "ebitda_pct":     ["Ebitda%"],
    "roce5y_raw":     ["ROCE 5Y CAGR"],
    "eps5y_cagr":     ["EPS 5Y CAGR"],
}

FUND_INDICATORS = [
    ("sales_yoy",    "Sales Growth YoY"),
    ("sales_qoq",    "Sales Growth QoQ"),
    ("rocer",        "ROCE"),
    ("roer",         "ROE"),
    ("de_ratio",     "D/E Ratio"),
    ("fcf_r",        "FCF"),
    ("eps_s",        "EPS"),
    ("roce5y_score", "ROCE 5Y CAGR"),
    ("pat_margin",   "PAT Margin"),
    ("ebitda_margin","EBITDA Margin"),
]

TECH_INDICATORS = [
    ("rsir",      "RSI"),
    ("wma30r",    "30 WMA"),
    ("weekly_rs", "Weekly RS (52W)"),
    ("daily_rs",  "Daily RS (52W)"),
]


# ── Excel reading utilities (self-contained, no import from main.py) ──────────
def _norm(v: Any) -> str:
    if v is None:
        return ""
    return " ".join(str(v).strip().split())


def _ckey(v: Any) -> str:
    t = _norm(v).lower()
    return "".join(c for c in t if c.isalnum())


def _dedup(headers: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out: list[str] = []
    for h in headers:
        n = seen.get(h, 0)
        seen[h] = n + 1
        out.append(h if n == 0 else f"{h}.{n}")
    return out


def _find_header_row(ws) -> int | None:
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=25, values_only=True), start=1):
        vals = {_ckey(c) for c in row if _norm(c)}
        if "companyname" in vals and "indexclean" in vals:
            return i
    return None


def _build_headers(ws, hi: int) -> list[str]:
    hrow = list(next(ws.iter_rows(min_row=hi, max_row=hi, values_only=True)))
    grow: list[Any] | None = None
    if hi > 1:
        grow = list(next(ws.iter_rows(min_row=hi - 1, max_row=hi - 1, values_only=True)))

    headers: list[str] = []
    active_g = ""
    for ci, hv in enumerate(hrow):
        gv = grow[ci] if grow and ci < len(grow) else None
        gt = _norm(gv)
        ht = _norm(hv)
        if gt:
            active_g = gt
        if ht.startswith("FY") and active_g:
            headers.append(f"{active_g} {ht}")
        elif ht:
            headers.append(ht)
        elif gt:
            headers.append(gt)
        else:
            headers.append(f"Unnamed_{ci}")
    return _dedup(headers)


def _read_raw(file_bytes: bytes) -> pd.DataFrame:
    wb = xl_load(io.BytesIO(file_bytes), read_only=True, data_only=False)
    sheet_name = hi = headers = None
    for ws in wb.worksheets:
        found = _find_header_row(ws)
        if found is not None:
            sheet_name = ws.title
            hi = found
            headers = _build_headers(ws, found)
            break
    wb.close()
    if sheet_name is None:
        raise ValueError("Could not find a sheet containing Company Name and Index_Clean headers.")
    df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=None,
        skiprows=hi,
        names=headers,
    )
    return df.dropna(how="all").reset_index(drop=True)


def _resolve(df: pd.DataFrame, aliases: list[str]) -> str | None:
    """Return the actual DataFrame column matching the first alias found, or None."""
    norm_lu = {_norm(c).lower(): c for c in df.columns}
    can_lu: dict[str, list[str]] = {}
    for c in df.columns:
        can_lu.setdefault(_ckey(c), []).append(c)

    for alias in aliases:
        key = _norm(alias).lower()
        if key in norm_lu:
            return norm_lu[key]
    for alias in aliases:
        key = _ckey(alias)
        if key in can_lu:
            return can_lu[key][0]
    return None


def read_report_dataframe(file_bytes: bytes) -> pd.DataFrame:
    """Read Excel and return a clean DataFrame with all report columns."""
    raw = _read_raw(file_bytes)

    out_cols: dict[str, Any] = {}
    for field, aliases in REPORT_COLS.items():
        col = _resolve(raw, aliases)
        out_cols[field] = raw[col] if col is not None else pd.Series([None] * len(raw))

    df = pd.DataFrame(out_cols)
    df["company_name"] = df["company_name"].map(lambda x: _norm(x) if x is not None else "")
    df["index_clean"]  = df["index_clean"].map(lambda x: _norm(x) if x is not None else "")

    numeric_fields = [f for f in REPORT_COLS if f not in ("company_name", "index_clean")]
    for f in numeric_fields:
        df[f] = pd.to_numeric(df[f], errors="coerce")

    df = df.dropna(subset=["company_name", "index_clean", "total"]).copy()
    df = df[(df["company_name"] != "") & (df["index_clean"] != "")].copy()

    # Add ranks
    df = df.sort_values(
        ["index_clean", "total", "company_name"], ascending=[True, False, True]
    ).reset_index(drop=True)
    df["top_rank"] = df.groupby("index_clean").cumcount() + 1

    df_bot = df.sort_values(
        ["index_clean", "total", "company_name"], ascending=[True, True, True]
    ).copy()
    df_bot["bot_rank"] = df_bot.groupby("index_clean").cumcount() + 1

    df = df.merge(
        df_bot[["index_clean", "company_name", "bot_rank"]],
        on=["index_clean", "company_name"],
        how="left",
    )
    return df


# ── Helpers ───────────────────────────────────────────────────────────────────
def _isna(val: Any) -> bool:
    if val is None:
        return True
    try:
        return bool(pd.isna(val))
    except Exception:
        return False


def _fmt(val: Any, decimals: int = 1, suffix: str = "", mult: float = 1.0) -> str:
    if _isna(val):
        return "—"
    try:
        return f"{float(val) * mult:.{decimals}f}{suffix}"
    except (TypeError, ValueError):
        return "—"


def _is_positive(val: Any) -> bool:
    if _isna(val):
        return False
    try:
        return float(val) > 0
    except (TypeError, ValueError):
        return False


def _tc(val: Any, label: str) -> Paragraph:
    """Return a tick or cross paragraph for an indicator."""
    pos = _is_positive(val)
    return Paragraph(f"{'✓' if pos else '✗'}  {label}", ST_TICK if pos else ST_CROSS)


def _score_cell(val: Any, label: str, col_w: float) -> Table:
    """Small 1×2 table: label row + coloured score row."""
    if not _isna(val):
        try:
            v = float(val)
            st = ST_SCPOS if v > 0 else ST_SCNEG
            vs = f"{v:.1f}"
        except (TypeError, ValueError):
            st = ST_SCNEU
            vs = "—"
    else:
        st = ST_SCNEU
        vs = "—"

    t = Table([[Paragraph(label, ST_SCLBL)], [Paragraph(vs, st)]], colWidths=[col_w])
    t.setStyle(TableStyle([
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


# ── Narrative ─────────────────────────────────────────────────────────────────
def _build_narrative(row: dict, rank: int, is_top: bool, index_name: str, total_in_index: int) -> str:
    name   = row.get("company_name") or "This stock"
    total  = _fmt(row.get("total"), 1)
    fs     = _fmt(row.get("fs"),    1)
    ts     = _fmt(row.get("ts"),    1)
    rsi    = _fmt(row.get("rsi_val"), 1)

    if is_top:
        s1 = (
            f"{name} ranks #{rank} out of {total_in_index} stocks in {index_name} "
            f"with a Total Score of {total} (FS: {fs}, TS: {ts}), placing it among the top performers."
        )
    else:
        s1 = (
            f"{name} ranks #{rank} from the bottom out of {total_in_index} stocks in {index_name} "
            f"with a Total Score of {total} (FS: {fs}, TS: {ts}), flagging it as a laggard."
        )

    pos_fund = [lbl for key, lbl in FUND_INDICATORS if _is_positive(row.get(key))]
    neg_fund = [lbl for key, lbl in FUND_INDICATORS if not _is_positive(row.get(key))]

    if is_top:
        s2 = (
            f"Key fundamental strengths: {', '.join(pos_fund[:4])}."
            if pos_fund
            else "Fundamental indicators are predominantly weak despite a strong technical score."
        )
    else:
        s2 = (
            f"Fundamental concerns include: {', '.join(neg_fund[:4])}."
            if neg_fund
            else "Fundamental indicators are mostly healthy, suggesting the weakness is technical."
        )

    pos_tech = [lbl for key, lbl in TECH_INDICATORS if _is_positive(row.get(key))]
    n_pos = len(pos_tech)
    qual = "strong" if n_pos >= 3 else ("moderate" if n_pos >= 2 else "weak")
    s3 = f"Technical setup is {qual} with {n_pos}/4 positive signals and RSI at {rsi}."

    return f"{s1}  {s2}  {s3}"


# ── Stock card ────────────────────────────────────────────────────────────────
def _stock_card(row: dict, rank: int, is_top: bool, index_name: str, total_in_index: int) -> KeepTogether:
    hdr_bg = CN
    body_bg = CGND if is_top else CRDD
    acc = CGNA if is_top else CRDA
    W = USABLE_W
    HW = W / 2
    TW = W / 3

    cmp_str = "₹" + _fmt(row.get("cmp"), 2) if not _isna(row.get("cmp")) else "—"
    tag = f"#{rank}  {'▲ TOP' if is_top else '▼ BOTTOM'}"

    # ── Header: Company name + CMP ────────────────────────────────────────────
    hdr = Table(
        [[Paragraph(f"<b>{tag}</b>   {row.get('company_name', '—')}", ST_COMPANY),
          Paragraph(f"CMP: {cmp_str}", ST_CMP)]],
        colWidths=[W * 0.72, W * 0.28],
    )
    hdr.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), hdr_bg),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING",   (0, 0), (0,  -1), 8),
        ("RIGHTPADDING",  (-1, 0), (-1, -1), 8),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW",     (0, 0), (-1, -1), 1.5, acc),
    ]))

    # ── Scores row ────────────────────────────────────────────────────────────
    scores = Table(
        [[_score_cell(row.get("total"), "TOTAL SCORE", TW - 2),
          _score_cell(row.get("fs"),    "FUNDAMENTAL",  TW - 2),
          _score_cell(row.get("ts"),    "TECHNICAL",    TW - 2)]],
        colWidths=[TW, TW, TW],
    )
    scores.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), body_bg),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("INNERGRID",     (0, 0), (-1, -1), 0.5, CBDR),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
    ]))

    # ── Scorecards: Fundamentals + Technicals ─────────────────────────────────
    def _ind_table(items: list, col_w: float) -> Table:
        data = [[Paragraph("", ST_INDHDR)]] + [[_tc(row.get(k), lbl)] for k, lbl in items]
        t = Table(data, colWidths=[col_w - 6])
        t.setStyle(TableStyle([
            ("TOPPADDING",    (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ]))
        return t

    fund_block = Table(
        [[Paragraph("FUNDAMENTALS", ST_INDHDR)], [_ind_table(FUND_INDICATORS, HW)]],
        colWidths=[HW - 1],
    )
    tech_block = Table(
        [[Paragraph("TECHNICALS", ST_INDHDR)], [_ind_table(TECH_INDICATORS, HW)]],
        colWidths=[HW - 1],
    )
    for blk in (fund_block, tech_block):
        blk.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (0, 0), CN2),
            ("BACKGROUND",    (0, 1), (-1, -1), CBG),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 6),
            ("LINEBELOW",     (0, 0), (0, 0), 0.5, acc),
        ]))

    scorecard_row = Table([[fund_block, tech_block]], colWidths=[HW, HW])
    scorecard_row.setStyle(TableStyle([
        ("VALIGN",    (0, 0), (-1, -1), "TOP"),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, CBDR),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
    ]))

    # ── Raw metrics row ───────────────────────────────────────────────────────
    rsi    = _fmt(row.get("rsi_val"),    1)
    yoy    = _fmt(row.get("yoy_pct"),    1, "%")
    qoq    = _fmt(row.get("qoq_pct"),    1, "%")
    de     = _fmt(row.get("de_raw"),     2)
    ebitda = _fmt(row.get("ebitda_pct"), 1, "%")
    roce5y = _fmt(row.get("roce5y_raw"), 1, "%", mult=100)
    eps5y  = _fmt(row.get("eps5y_cagr"), 1, "%")

    metrics_text = (
        f"<b>RSI:</b> {rsi}  ·  <b>Revenue YoY:</b> {yoy}  ·  <b>Revenue QoQ:</b> {qoq}  ·  "
        f"<b>D/E:</b> {de}  ·  <b>EBITDA:</b> {ebitda}  ·  "
        f"<b>ROCE 5Y CAGR:</b> {roce5y}  ·  <b>EPS 5Y CAGR:</b> {eps5y}"
    )
    met_row = Table([[Paragraph(metrics_text, ST_MVAL)]], colWidths=[W])
    met_row.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), CN2),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
    ]))

    # ── Narrative row ─────────────────────────────────────────────────────────
    narrative = _build_narrative(row, rank, is_top, index_name, total_in_index)
    narr_row = Table([[Paragraph(narrative, ST_NARR)]], colWidths=[W])
    narr_row.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), CBG),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
        ("LINEABOVE",     (0, 0), (-1, 0),  0.5, CBDR),
    ]))

    # ── Outer card wrapper ────────────────────────────────────────────────────
    card = Table(
        [[hdr], [scores], [scorecard_row], [met_row], [narr_row]],
        colWidths=[W],
    )
    card.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("BOX",           (0, 0), (-1, -1), 1, acc),
    ]))

    return KeepTogether([card, Spacer(1, 4 * mm)])


# ── Page chrome ───────────────────────────────────────────────────────────────
def _add_footer(canvas, doc, date_str: str) -> None:
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(CMUT)
    canvas.drawString(MARGIN, 7 * mm, "Proprietary & Confidential")
    canvas.drawRightString(PAGE_W - MARGIN, 7 * mm, f"Page {doc.page}  |  {date_str}")
    canvas.setStrokeColor(CBDR)
    canvas.setLineWidth(0.5)
    canvas.line(MARGIN, 11 * mm, PAGE_W - MARGIN, 11 * mm)
    canvas.restoreState()


def _index_banner(index_name: str) -> Table:
    t = Table([[Paragraph(index_name.upper(), ST_IDXHDR)]], colWidths=[USABLE_W])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), CN),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LINEBELOW",     (0, 0), (-1, -1), 2, CGOLD),
    ]))
    return t


def _section_banner(text: str, acc: colors.Color) -> Table:
    t = Table([[Paragraph(text, ST_SECHDR)]], colWidths=[USABLE_W])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), CN2),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("LINEABOVE",     (0, 0), (-1, 0),  1, acc),
        ("LINEBELOW",     (0, 0), (-1, -1), 0.5, CBDR),
    ]))
    return t


# ── Cover page ────────────────────────────────────────────────────────────────
def _cover_page(date_str: str, indices: list[str]) -> list:
    W = USABLE_W
    elems = []
    elems.append(Spacer(1, 25 * mm))

    title_tbl = Table(
        [[Paragraph("Daily Stock Scoring Report", ST_CVRTITLE)],
         [Paragraph(f"Reporting Date: {date_str}", ST_CVRSUB)]],
        colWidths=[W],
    )
    title_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), CN),
        ("TOPPADDING",    (0, 0), (0, 0),   22),
        ("BOTTOMPADDING", (0, 0), (0, 0),   8),
        ("TOPPADDING",    (0, 1), (0, 1),   4),
        ("BOTTOMPADDING", (0, 1), (0, 1),   22),
        ("BOX",           (0, 0), (-1, -1), 1.5, CGOLD),
    ]))
    elems.append(title_tbl)
    elems.append(Spacer(1, 12 * mm))

    # Indices covered
    idx_data = [[Paragraph(f"•  {idx}", ST_COVBODY)] for idx in sorted(indices)]
    idx_data.insert(0, [Paragraph("Indices Covered", ST_SECHDR)])
    idx_tbl = Table(idx_data, colWidths=[W])
    idx_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (0, 0),   CN2),
        ("BACKGROUND",    (0, 1), (-1, -1), CBG),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("LINEBELOW",     (0, 0), (0, 0),   0.5, CGOLD),
        ("BOX",           (0, 0), (-1, -1), 0.5, CBDR),
    ]))
    elems.append(idx_tbl)
    elems.append(Spacer(1, 10 * mm))

    # Scoring legend
    legend_rows = [
        ("FS — Fundamental Score", "Binary sum of: Sales YoY, Sales QoQ, ROCE, ROE, D/E Ratio, FCF, EPS, ROCE 5Y CAGR, PAT Margin, EBITDA Margin."),
        ("TS — Technical Score",   "Binary sum of: RSI, 30 WMA, Weekly RS (52W), Daily RS (52W)."),
        ("Total Score",            "FS + TS  —  primary ranking column. Higher = stronger stock."),
        ("✓  Tick",                "Indicator score > 0  (positive / bullish signal)."),
        ("✗  Cross",               "Indicator score ≤ 0  (negative / bearish or neutral signal)."),
        ("Top 3",                  "Three highest Total Score stocks in the index  (green cards)."),
        ("Bottom 3",               "Three lowest Total Score stocks in the index  (red cards)."),
    ]

    leg_data = [[Paragraph(k, ST_LEGKEY), Paragraph(v, ST_LEGEND)] for k, v in legend_rows]
    leg_data.insert(0, [Paragraph("Scoring Legend & Indicator Definitions", ST_SECHDR), Paragraph("", ST_LEGEND)])
    leg_tbl = Table(leg_data, colWidths=[W * 0.34, W * 0.66])
    leg_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  CN2),
        ("BACKGROUND",    (0, 1), (0, -1),  colors.HexColor("#0d1e30")),
        ("BACKGROUND",    (1, 1), (1, -1),  CBG),
        ("SPAN",          (0, 0), (1, 0)),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
        ("INNERGRID",     (0, 0), (-1, -1), 0.3, CBDR),
        ("BOX",           (0, 0), (-1, -1), 0.5, CBDR),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LINEBELOW",     (0, 0), (-1, 0),  0.5, CGOLD),
    ]))
    elems.append(leg_tbl)
    elems.append(PageBreak())
    return elems


# ── Main entry point ──────────────────────────────────────────────────────────
def generate_pdf(file_bytes: bytes, date_str: str) -> bytes:
    """Parse the Excel file and return a styled PDF report as bytes."""
    df = read_report_dataframe(file_bytes)
    indices = sorted(df["index_clean"].unique())

    buf = io.BytesIO()

    doc = BaseDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN,
        bottomMargin=FOOTER_H + 4 * mm,
    )

    frame = Frame(
        MARGIN,
        FOOTER_H + 3 * mm,
        USABLE_W,
        PAGE_H - MARGIN - FOOTER_H - 5 * mm,
        id="main",
    )

    _date = date_str
    template = PageTemplate(
        id="main_template",
        frames=[frame],
        onPage=lambda c, d: _add_footer(c, d, _date),
    )
    doc.addPageTemplates([template])

    story: list = []
    story.extend(_cover_page(date_str, list(indices)))

    for idx_name in indices:
        idx_df = df[df["index_clean"] == idx_name].copy()
        total_in_idx = len(idx_df)

        top3 = idx_df.nsmallest(3, "top_rank")
        bot3 = idx_df.nsmallest(3, "bot_rank")

        story.append(_index_banner(idx_name))
        story.append(Spacer(1, 4 * mm))

        story.append(_section_banner("▲   TOP 3 STOCKS", CGNA))
        story.append(Spacer(1, 3 * mm))

        for _, row in top3.iterrows():
            story.append(_stock_card(row.to_dict(), int(row["top_rank"]), True, idx_name, total_in_idx))

        story.append(Spacer(1, 5 * mm))
        story.append(_section_banner("▼   BOTTOM 3 STOCKS", CRDA))
        story.append(Spacer(1, 3 * mm))

        for _, row in bot3.iterrows():
            story.append(_stock_card(row.to_dict(), int(row["bot_rank"]), False, idx_name, total_in_idx))

        story.append(PageBreak())

    doc.build(story)
    return buf.getvalue()
