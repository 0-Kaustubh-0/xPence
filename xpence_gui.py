#!/usr/bin/env python3
"""
xPence GUI — Report Generator + Bank Scrubber
==============================================
Single-file application.  Drop alongside xpence_analyzer.py and run.

Bank scrubber engine is fully inlined (no xpence_scrubber.py needed).
Worker threads call scrubber functions directly — no subprocess, no tkinter
import-at-module-level crash.

Supported banks (auto-detected by filename or column fingerprint):
  TD Bank       headerless 5-col: date, desc, debit, credit, balance
  CIBC          headerless 5-col: date, desc, debit, credit, card_no
  Wealthsimple  named headers:    transaction_date ... net_cash_amount
  Generic       best-effort named-header matching

Output format matches xPence filtered_data exactly:
  col 0  Excel date serial (int)
  col 1  Transaction description (str)
  col 2  Amount: positive=expense/debit, negative=credit/payment
  col 3  Empty  (analyzer expects 4 cols; col 3 is the credit col, unused here)
"""

# ==============================================================================
# PART 1  BANK SCRUBBER ENGINE   (no tkinter dependency)
# ==============================================================================
import os
import re
import subprocess
from datetime import datetime, date
from pathlib import Path

import pandas as pd

# -- Excel date serial ---------------------------------------------------------
_EXCEL_EPOCH = datetime(1899, 12, 30)


def _to_serial(d) -> int:
    """Convert a date-like value to an Excel date serial integer."""
    if isinstance(d, (int, float)):
        return int(d)
    if isinstance(d, datetime):
        return (d - _EXCEL_EPOCH).days
    if isinstance(d, date):
        return (datetime(d.year, d.month, d.day) - _EXCEL_EPOCH).days
    if isinstance(d, str):
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d",
                    "%d-%m-%Y", "%m-%d-%Y", "%d %b %Y", "%b %d, %Y"):
            try:
                return (datetime.strptime(d.strip(), fmt) - _EXCEL_EPOCH).days
            except ValueError:
                pass
        try:
            dt = pd.to_datetime(d, dayfirst=False, errors="raise")
            return (dt.to_pydatetime() - _EXCEL_EPOCH).days
        except Exception:
            pass
    return 0


# -- PII scrubber patterns -----------------------------------------------------
_PII_COL_RE = re.compile(
    r"\b(card[\s_-]?(no|num|number)?|account[\s_-]?(no|id|number|holder)?|"
    r"holder|client|customer|member|sin|ssn|transit|sort[\s_-]?code|"
    r"branch|institution|routing|pan|iban|bic|swift)\b",
    re.IGNORECASE,
)
_CARD_RE    = re.compile(r"\b\d{4}[\s*\-]?\*{4,}[\s*\-]?\d{4}\b|\b\d{13,19}\b")
_ACCOUNT_RE = re.compile(
    r"\b[A-Z]{2}\d{2}[A-Z0-9]{4}\d{7}([A-Z0-9]?){0,16}\b"
    r"|\b\d{5,6}[\- ]\d{7,12}\b"
)
_PHONE_RE = re.compile(r"\(?\d{3}\)?[\s.\-]\d{3}[\s.\-]\d{4}|\b\d{10,11}\b")


def _scrub(val: str) -> str:
    val = _CARD_RE.sub("", val)
    val = _ACCOUNT_RE.sub("", val)
    val = _PHONE_RE.sub("", val)
    return re.sub(r"  +", " ", val).strip().strip(",").strip()


def _is_pii_col(name: str) -> bool:
    return bool(_PII_COL_RE.search(str(name)))


def _f(v) -> float:
    """Safe string-to-float; returns 0.0 on NaN / empty."""
    s = str(v).strip()
    if s.lower() in ("", "nan", "none"):
        return 0.0
    try:
        return float(s.replace(",", ""))
    except (ValueError, TypeError):
        return 0.0


# -- XLS reader (no LibreOffice dependency) ------------------------------------
def _read_xls(filepath: str) -> pd.DataFrame:
    """Read legacy .xls files using xlrd via pandas, no external tools needed."""
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == ".xls":
            # xlrd handles legacy .xls natively; install with: pip install xlrd
            return pd.read_excel(filepath, header=None, dtype=str, engine="xlrd")
        else:
            # .xlsm / .xlsx fallback (should rarely reach here)
            return pd.read_excel(filepath, header=None, dtype=str, engine="openpyxl")
    except ImportError as e:
        pkg = "xlrd" if "xlrd" in str(e) else "openpyxl"
        raise RuntimeError(
            f"Missing package '{pkg}'. Install it with:  pip install {pkg}") from e
    except Exception as e:
        raise RuntimeError(f"Could not read {os.path.basename(filepath)}: {e}") from e


# -- Bank parsers --------------------------------------------------------------

def _parse_td(df: pd.DataFrame, filepath: str = "") -> pd.DataFrame:
    """
    TD Bank — no header, 5 cols:
      0=date(MM/DD/YYYY)  1=description  2=debit  3=credit  4=balance(drop)

    Debit  (col 2) -> positive amount  (expense / purchase)
    Credit (col 3) -> negative amount  (payment / refund)
    Balance(col 4) -> dropped

    Account type isn't exposed by this export format, so it's inferred from
    the filename (falls back to "Credit" — see _infer_account_type()).
    """
    acct_type = _infer_account_type(filepath)
    rows = []
    for _, r in df.iterrows():
        serial = _to_serial(str(r.iloc[0]).strip())
        if serial <= 0:
            continue
        desc   = _scrub(str(r.iloc[1]).strip())
        debit  = _f(r.iloc[2]) if df.shape[1] > 2 else 0.0
        credit = _f(r.iloc[3]) if df.shape[1] > 3 else 0.0
        # col 4 = running balance, always dropped
        if debit:
            amount = debit       # purchase, already positive
        elif credit:
            amount = -credit     # payment/refund, make negative
        else:
            amount = 0.0
        rows.append((serial, desc, round(amount, 2), acct_type))
    return pd.DataFrame(rows, columns=["date_serial", "description", "amount", "account_type"])


def _parse_cibc(df: pd.DataFrame, filepath: str = "") -> pd.DataFrame:
    """
    CIBC — no header, 5 cols:
      0=date(MM/DD/YYYY)  1=description  2=debit  3=credit  4=card_no(PII,drop)

    Account type isn't exposed by this export format, so it's inferred from
    the filename (falls back to "Credit" — see _infer_account_type()).
    """
    acct_type = _infer_account_type(filepath)
    rows = []
    for _, r in df.iterrows():
        serial = _to_serial(str(r.iloc[0]).strip())
        if serial <= 0:
            continue
        desc   = _scrub(str(r.iloc[1]).strip())
        debit  = _f(r.iloc[2]) if df.shape[1] > 2 else 0.0
        credit = _f(r.iloc[3]) if df.shape[1] > 3 else 0.0
        # col 4 = masked card number, dropped
        if debit:
            amount = debit
        elif credit:
            amount = -credit
        else:
            amount = 0.0
        rows.append((serial, desc, round(amount, 2), acct_type))
    return pd.DataFrame(rows, columns=["date_serial", "description", "amount", "account_type"])


def _parse_wealthsimple(df: pd.DataFrame) -> pd.DataFrame:
    """
    Wealthsimple chequing export — named headers.

    Columns kept:
      transaction_date    -> date
      activity_type       -> description fallback  (e.g. "MoneyMovement", "Interest")
      activity_sub_type   -> description primary   (e.g. "EFT", "E_TRFOUT")
      net_cash_amount     -> amount
      account_type        -> account type (e.g. "Chequing", "Savings") — kept
                             as-is rather than dropped, so the report can
                             show/filter by account type.

    All other columns dropped (PII or irrelevant):
      settlement_date, account_id, direction, symbol,
      name, currency, quantity, unit_price, commission

    Sign: Wealthsimple net_cash_amount:
      negative = money left account (withdrawal/expense) -> store as POSITIVE
      positive = money entered account (deposit/interest) -> store as NEGATIVE
    This already matches the xPence filtered_data convention.

    Rows where 'name' is NaN are ordinary bank movements (withdrawals, deposits,
    interest) — they are described by activity_sub_type / activity_type instead.
    """
    _LABELS = {
        "E_TRFOUT": "Withdrawal",
        "E_TRFIN":  "Deposit",
        "EFT":      "EFT Transfer",
        "E_PAD":    "Pre-auth Debit",
        "INTEREST": "Interest",
        "FEE":      "Fee",
        "BUY":      "Purchase",
        "SELL":     "Sale",
        "DIV":      "Dividend",
    }

    cols_lc = {c.lower().strip(): c for c in df.columns}

    def gc(*hints):
        for h in hints:
            if h in cols_lc:
                return cols_lc[h]
        return None

    date_col    = gc("transaction_date", "date")
    amount_col  = gc("net_cash_amount", "amount")
    subtype_col = gc("activity_sub_type")
    type_col    = gc("activity_type")
    accttype_col = gc("account_type")

    def label(sub, typ) -> str:
        for v in (sub, typ):
            s = str(v).strip() if v is not None else ""
            if s and s.lower() not in ("nan", "none", ""):
                return _LABELS.get(s.upper(), s)
        return "Transfer"

    rows = []
    for _, r in df.iterrows():
        raw_date = str(r[date_col]).strip() if date_col else ""
        serial   = _to_serial(raw_date)
        if serial <= 0:
            continue

        # Skip rows with no amount (e.g. footer row "As of …")
        amt_raw = r[amount_col] if amount_col else None
        if amt_raw is None or str(amt_raw).strip().lower() in ("nan", "none", ""):
            continue
        amt = _f(amt_raw)

        sub  = r.get(subtype_col) if subtype_col else None
        typ  = r.get(type_col)    if type_col    else None
        desc = label(sub, typ)

        acct_raw = str(r.get(accttype_col)).strip() if accttype_col else ""
        acct_type = acct_raw if acct_raw and acct_raw.lower() not in ("nan", "none", "") else _DEFAULT_ACCOUNT_TYPE

        # Negate: WS negative -> positive (expense); WS positive -> negative (credit)
        rows.append((serial, desc, round(-amt, 2), acct_type))

    return pd.DataFrame(rows, columns=["date_serial", "description", "amount", "account_type"])


def _parse_generic(df: pd.DataFrame, filepath: str = "") -> pd.DataFrame:
    """Best-effort parser for unknown banks with named headers."""
    _DATE_H   = ["date", "transaction date", "trans date", "posted date", "value date"]
    _NAME_H   = ["description", "transaction", "merchant", "narration",
                 "particulars", "details", "memo", "payee", "reference"]
    _DEBIT_H  = ["debit", "spent", "amount", "withdrawal", "charge", "dr", "expense"]
    _CREDIT_H = ["credit", "deposit", "payment in", "cr", "received", "refund"]
    _ACCTYPE_H = ["account type", "account_type", "acct type", "card type", "account"]

    def best(candidates, hints):
        lmap = {c.lower().strip(): c for c in candidates if not _is_pii_col(c)}
        for h in hints:
            if h in lmap:
                return lmap[h]
        for h in hints:
            for lc, orig in lmap.items():
                if h in lc or lc in h:
                    return orig
        return None

    safe = [c for c in df.columns if not _is_pii_col(str(c))]
    df   = df[safe]
    cols = list(df.columns)

    date_col    = best(cols, _DATE_H)   or (cols[0] if cols else None)
    name_col    = best(cols, _NAME_H)   or (cols[1] if len(cols) > 1 else None)
    debit_col   = best(cols, _DEBIT_H)  or (cols[2] if len(cols) > 2 else None)
    credit_col  = best(cols, _CREDIT_H) or (cols[3] if len(cols) > 3 else None)
    accttype_col = best(cols, _ACCTYPE_H)
    # If nothing in the file itself identifies the account type, fall back to
    # a filename-based guess (see _infer_account_type()).
    fallback_acct_type = _infer_account_type(filepath)

    rows = []
    for _, r in df.iterrows():
        serial = _to_serial(str(r[date_col]).strip() if date_col else "")
        if serial <= 0:
            continue
        desc = _scrub(str(r[name_col]).strip() if name_col else "")
        if not desc or desc.lower() in ("nan", "none"):
            continue
        dv = _f(r[debit_col])  if debit_col  else 0.0
        cv = _f(r[credit_col]) if (credit_col and credit_col != debit_col) else 0.0
        if dv and cv:
            amount = dv - cv
        elif dv:
            amount = dv
        elif cv:
            amount = -cv
        else:
            amount = 0.0
        acct_raw = str(r.get(accttype_col)).strip() if accttype_col else ""
        acct_type = acct_raw if acct_raw and acct_raw.lower() not in ("nan", "none", "") else fallback_acct_type
        rows.append((serial, desc, round(amount, 2), acct_type))
    return pd.DataFrame(rows, columns=["date_serial", "description", "amount", "account_type"])


# -- Bank auto-detector --------------------------------------------------------
_BANK_FP = {
    "td":           re.compile(r"td[\s_-]?bank|tdbank|td[\s_-]?(chequing|savings|credit)", re.I),
    "cibc":         re.compile(r"cibc", re.I),
    "wealthsimple": re.compile(r"wealthsimple|wealth[\s_-]?simple", re.I),
    "rbc":          re.compile(r"\brbc\b|royal[\s_-]?bank", re.I),
    "bmo":          re.compile(r"\bbmo\b|bank[\s_-]?of[\s_-]?montreal", re.I),
    "scotiabank":   re.compile(r"scotiabank|scotia", re.I),
    "tangerine":    re.compile(r"tangerine", re.I),
    "eqbank":       re.compile(r"eq[\s_-]?bank", re.I),
}
_WS_COLS = {"transaction_date", "account_id", "net_cash_amount", "activity_type"}

# Default account type assumed for bank formats that don't expose one
# explicitly (TD / CIBC / generic exports are almost always credit-card
# statements downloaded from online banking).
_DEFAULT_ACCOUNT_TYPE = "Credit"
_ACCTYPE_FP = {
    "Chequing": re.compile(r"chequ|checking", re.I),
    "Savings":  re.compile(r"saving", re.I),
    "Credit":   re.compile(r"credit", re.I),
}


def _infer_account_type(filepath: str) -> str:
    """Best-effort account type guess from the filename, falling back to the
    safe default ('Credit') when nothing in the name gives it away."""
    name = os.path.basename(filepath)
    for acct_type, pat in _ACCTYPE_FP.items():
        if pat.search(name):
            return acct_type
    return _DEFAULT_ACCOUNT_TYPE


def _detect_bank(filepath: str, df: pd.DataFrame) -> str:
    name = os.path.basename(filepath).lower()
    for bank, pat in _BANK_FP.items():
        if pat.search(name):
            return bank
    if df is not None:
        lc = {c.lower().strip() for c in df.columns}
        if _WS_COLS.issubset(lc):
            return "wealthsimple"
        if len(df.columns) == 5 and df.columns[0] == 0:
            sample = df.iloc[:, 4].dropna().astype(str)
            if any("****" in v or re.match(r"\d{4}\*+\d{4}", v) for v in sample.head(10)):
                return "cibc"
            return "td"
    return "generic"


# -- Raw-file reader -----------------------------------------------------------
def _read_raw(filepath: str) -> pd.DataFrame:
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".xls":
        raw = _read_xls(filepath)
    elif ext in (".xlsx", ".xlsm"):
        try:
            raw = pd.read_excel(filepath, header=0, dtype=str)
            if all(isinstance(c, int) or str(c).isdigit() for c in raw.columns):
                raw = pd.read_excel(filepath, header=None, dtype=str)
        except Exception:
            raw = pd.read_excel(filepath, header=None, dtype=str)
    elif ext == ".ods":
        raw = pd.read_excel(filepath, engine="odf", header=0, dtype=str)
    elif ext == ".csv":
        raw = None
        for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
            for sep in (",", ";", "\t"):
                try:
                    tmp = pd.read_csv(filepath, sep=sep, encoding=enc,
                                      dtype=str, skipinitialspace=True)
                    if tmp.shape[1] >= 2:
                        raw = tmp; break
                except Exception:
                    pass
            if raw is not None:
                break
        if raw is None:
            raw = pd.read_csv(filepath, header=None, dtype=str)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    raw = raw[~raw.iloc[:, 0].astype(str).str.startswith("As of")].copy()
    raw = raw.dropna(how="all").reset_index(drop=True)
    return raw


def _parse(filepath: str, raw: pd.DataFrame) -> pd.DataFrame:
    bank = _detect_bank(filepath, raw)
    if bank == "td":           return _parse_td(raw, filepath)
    if bank == "cibc":         return _parse_cibc(raw, filepath)
    if bank == "wealthsimple": return _parse_wealthsimple(raw)
    return _parse_generic(raw, filepath)


def _write_xpence_xlsx(clean: pd.DataFrame, output_path: str) -> str:
    """Write a cleaned DataFrame to the xPence filtered_data .xlsx format.

    Columns (no header row, positional):
      0=date_serial  1=description  2=amount(signed)  3=""(reserved)
      4=account_type  — identified by the parser for this bank format,
                        or the safe "Credit" default when it couldn't be.
    """
    clean = clean[clean["date_serial"] > 0].copy()
    acct_col = clean["account_type"].astype(str) if "account_type" in clean.columns else _DEFAULT_ACCOUNT_TYPE
    out = pd.DataFrame({
        0: clean["date_serial"].astype(int),
        1: clean["description"].astype(str),
        2: clean["amount"].round(2),
        3: "",
        4: acct_col,
    })
    out.to_excel(output_path, index=False, header=False)
    return output_path


def scrub_file(filepath: str, output_path: str = None) -> str:
    """Scrub a single bank export and write xPence .xlsx. Returns output path."""
    raw   = _read_raw(filepath)
    clean = _parse(filepath, raw)
    if output_path is None:
        output_path = os.path.splitext(filepath)[0] + "_scrubbed.xlsx"
    return _write_xpence_xlsx(clean, output_path)


def scrub_and_merge(filepaths: list, output_path: str) -> tuple:
    """Scrub multiple files and merge into one xPence .xlsx.
    Returns (output_path, {filepath: row_count | error_str})."""
    frames, stats = [], {}
    for fp in filepaths:
        try:
            raw   = _read_raw(fp)
            clean = _parse(fp, raw)
            clean = clean[clean["date_serial"] > 0]
            frames.append(clean)
            stats[fp] = len(clean)
        except Exception as exc:
            stats[fp] = f"error: {exc}"

    if not frames:
        raise ValueError("No valid transactions found in any provided file.")

    merged = (pd.concat(frames, ignore_index=True)
                .sort_values("date_serial", ascending=False)
                .reset_index(drop=True))
    _write_xpence_xlsx(merged, output_path)
    return output_path, stats


# ==============================================================================
# ==============================================================================
# ==============================================================================
# PART 2  GUI
# ==============================================================================
import sys
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import shutil

SCRIPT_DIR    = Path(__file__).parent.resolve()
ANALYZER      = SCRIPT_DIR / "xpence_analyzer.py"
MERCHANT_JSON = SCRIPT_DIR / "merchant_categories.json"

# ── Palette ───────────────────────────────────────────────────────────────────
BG        = "#0d1117"
CARD      = "#161b22"
CARD2     = "#1c2230"
BORDER    = "#30363d"
TEXT      = "#e6edf3"
MUTED     = "#7d8590"
ACCENT    = "#58a6ff"
GREEN     = "#3fb950"
RED       = "#f85149"
GOLD      = "#c9900a"
GOLD_LT   = "#e6a812"
GOLD_FG   = "#fff8e7"
PURPLE    = "#8b5cf6"
TEAL      = "#22d3ee"
WARN_YEL  = "#fbbf24"   # "What gets scrubbed" heading colour


def _lt(h, d=22):
    try:
        r,g,b = int(h[1:3],16),int(h[3:5],16),int(h[5:7],16)
        return f"#{min(255,r+d):02x}{min(255,g+d):02x}{min(255,b+d):02x}"
    except Exception: return h

def _dk(h, d=18):
    try:
        r,g,b = int(h[1:3],16),int(h[3:5],16),int(h[5:7],16)
        return f"#{max(0,r-d):02x}{max(0,g-d):02x}{max(0,b-d):02x}"
    except Exception: return h


# ── Round-cornered canvas button ──────────────────────────────────────────────
class RoundButton(tk.Canvas):
    """
    A canvas-drawn button with genuine rounded corners.
    Supports: text, bg/fg colours, hover, disabled state, variable height/width.
    """
    RADIUS = 10

    def __init__(self, parent, text, command, bg=ACCENT, fg="#ffffff",
                 font=("Segoe UI", 10, "bold"), padx=22, pady=9,
                 width=None, state="normal", **kw):
        self._bg      = bg
        self._fg      = fg
        self._hov     = _lt(bg, 28)
        self._dis_bg  = CARD2
        self._dis_fg  = MUTED
        self._font    = font
        self._padx    = padx
        self._pady    = pady
        self._text    = text
        self._cmd     = command
        self._state   = state
        self._fix_w   = width

        # Measure text to get canvas size
        tmp = tk.Label(parent, text=text, font=font)
        tmp.update_idletasks()
        tw = tmp.winfo_reqwidth()
        th = tmp.winfo_reqheight()
        tmp.destroy()

        cw = (width * 7 if width else tw + padx * 2)
        ch = th + pady * 2

        super().__init__(parent, width=cw, height=ch,
                         bg=parent["bg"], highlightthickness=0,
                         cursor="hand2" if state == "normal" else "", **kw)

        self._cw, self._ch = cw, ch
        self._draw(bg if state == "normal" else self._dis_bg,
                   fg if state == "normal" else self._dis_fg)

        self.bind("<Configure>", self._on_configure)
        self.bind("<Enter>",    self._on_enter)
        self.bind("<Leave>",    self._on_leave)
        self.bind("<Button-1>", self._on_click)

    # ── drawing ───────────────────────────────────────────────────────────────
    def _rr(self, canvas, x1, y1, x2, y2, r, **kw):
        """Draw a filled rounded rectangle."""
        canvas.create_arc(x1,    y1,    x1+2*r, y1+2*r, start= 90, extent=90, **kw)
        canvas.create_arc(x2-2*r,y1,    x2,     y1+2*r, start=  0, extent=90, **kw)
        canvas.create_arc(x2-2*r,y2-2*r,x2,     y2,     start=270, extent=90, **kw)
        canvas.create_arc(x1,    y2-2*r,x1+2*r, y2,     start=180, extent=90, **kw)
        canvas.create_rectangle(x1+r, y1,   x2-r, y2,   **kw)
        canvas.create_rectangle(x1,   y1+r, x2,   y2-r, **kw)

    def _draw(self, bg, fg):
        self.delete("all")
        r = self.RADIUS
        self._rr(self, 0, 0, self._cw, self._ch, r,
                 fill=bg, outline=bg)
        self.create_text(self._cw//2, self._ch//2,
                         text=self._text, fill=fg,
                         font=self._font, anchor="center")

    # ── events ────────────────────────────────────────────────────────────────
    def _on_configure(self, event):
        if event.width > 1 and event.height > 1:
            self._cw = event.width
            self._ch = event.height
            col = self._bg if self._state == "normal" else self._dis_bg
            fgc = self._fg if self._state == "normal" else self._dis_fg
            self._draw(col, fgc)

    def _on_enter(self, _):
        if self._state == "normal":
            self._draw(self._hov, self._fg)

    def _on_leave(self, _):
        if self._state == "normal":
            self._draw(self._bg, self._fg)

    def _on_click(self, _):
        if self._state == "normal" and self._cmd:
            self._cmd()

    # ── public API ────────────────────────────────────────────────────────────
    def config(self, **kw):
        changed = False
        if "state" in kw:
            self._state = kw.pop("state")
            self.configure(cursor="hand2" if self._state == "normal" else "")
            changed = True
        if "text" in kw:
            self._text = kw.pop("text")
            changed = True
        if "bg_color" in kw:
            self._bg  = kw.pop("bg_color")
            self._hov = _lt(self._bg, 28)
            changed = True
        if changed:
            col = self._bg if self._state == "normal" else self._dis_bg
            fgc = self._fg if self._state == "normal" else self._dis_fg
            self._draw(col, fgc)
        super().config(**kw)

    configure = config


# ── Shared widget helpers ─────────────────────────────────────────────────────
class _W:
    def _lbl(self, p, t, fg=TEXT, font=("Segoe UI", 10, "bold"), anchor="w"):
        lbl = tk.Label(p, text=t, font=font, fg=fg, bg=p["bg"], anchor=anchor)
        lbl.pack(fill="x", pady=(0, 2))
        return lbl

    def _sublbl(self, p, t, fg=MUTED):
        lbl = tk.Label(p, text=t, font=("Segoe UI", 8), fg=fg, bg=p["bg"],
                       anchor="w", wraplength=640, justify="left")
        lbl.pack(fill="x")
        return lbl

    def _entry(self, p, v):
        return tk.Entry(p, textvariable=v, font=("Segoe UI", 10),
                        bg=CARD2, fg=TEXT, insertbackground=TEXT,
                        relief="flat", bd=0, highlightthickness=1,
                        highlightbackground=BORDER, highlightcolor=ACCENT)

    def _card(self, p, pady=(0, 10)):
        f = tk.Frame(p, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        f.pack(fill="x", pady=pady)
        return f

    def _status_frame(self, p, var):
        sf = tk.Frame(p, bg=CARD2, highlightthickness=1, highlightbackground=BORDER)
        sf.pack(fill="x", pady=(0, 6))
        lbl = tk.Label(sf, textvariable=var, font=("Segoe UI", 9),
                       fg=MUTED, bg=CARD2, anchor="w",
                       padx=12, pady=6, wraplength=640, justify="left")
        lbl.pack(fill="x")
        return lbl

    def _progressbar(self, p):
        sty = ttk.Style()
        sty.configure("xp.Horizontal.TProgressbar",
                      troughcolor=CARD2, background=ACCENT,
                      borderwidth=0, thickness=3)
        pb = ttk.Progressbar(p, style="xp.Horizontal.TProgressbar",
                             mode="indeterminate")
        pb.pack(fill="x", pady=(0, 8))
        return pb

    def _divider(self, p, pady=(6, 10)):
        tk.Frame(p, bg=BORDER, height=1).pack(fill="x", pady=pady)

    def _round_btn(self, p, text, cmd, bg=ACCENT, fg="#fff",
                   font=("Segoe UI", 10, "bold"), padx=22, pady=9,
                   width=None, state="normal"):
        b = RoundButton(p, text, cmd, bg=bg, fg=fg, font=font,
                        padx=padx, pady=pady, width=width, state=state)
        return b


# ── Nav pill (replaces ttk.Notebook) ─────────────────────────────────────────
class NavPill(tk.Frame):
    def __init__(self, parent, labels, frames):
        super().__init__(parent, bg=CARD)
        self._frames = frames
        self._btns   = []
        for i, lbl in enumerate(labels):
            b = RoundButton(self, lbl, lambda i=i: self._select(i),
                            bg=CARD, fg=MUTED,
                            font=("Segoe UI", 10, "bold"),
                            padx=22, pady=7)
            b.pack(side="left", padx=(0, 8))
            self._btns.append(b)
        self._select(0)

    def _select(self, idx):
        for i, (b, f) in enumerate(zip(self._btns, self._frames)):
            if i == idx:
                b._bg  = ACCENT
                b._fg  = "#fff"
                b._hov = _lt(ACCENT, 28)
                b._draw(ACCENT, "#fff")
                f.pack(fill="both", expand=True)
            else:
                b._bg  = CARD
                b._fg  = MUTED
                b._hov = CARD2
                b._draw(CARD, MUTED)
                f.pack_forget()


# ── Report Generator Tab ──────────────────────────────────────────────────────
class ReportTab(tk.Frame, _W):
    def __init__(self, master):
        super().__init__(master, bg=BG)
        self.csv_path    = tk.StringVar()
        self.out_path    = tk.StringVar()
        self.status_var  = tk.StringVar(value="Select a bank file or xPence data file to begin.")
        self._report_path   = None
        self._build()
        self._check_analyzer()

    def _build(self):
        body = tk.Frame(self, bg=BG, padx=24, pady=14)
        body.pack(fill="both", expand=True)

        # ── Input file ────────────────────────────────────────────────────────
        self._lbl(body, "Input file", font=("Segoe UI", 9, "bold"), fg=MUTED)
        c1 = self._card(body, pady=(4, 10))
        c1i = tk.Frame(c1, bg=CARD, padx=12, pady=10); c1i.pack(fill="x")
        row1 = tk.Frame(c1i, bg=CARD); row1.pack(fill="x")
        self._entry(row1, self.csv_path).pack(
            side="left", fill="x", expand=True, ipady=6, ipadx=6)
        self._round_btn(row1, "Browse", self._browse_csv,
                        bg=CARD2, fg=TEXT, font=("Segoe UI", 9),
                        padx=16, pady=6).pack(side="left", padx=(8, 0))
        self._sublbl(c1i,
            "Accepts raw bank exports (TD, CIBC, Wealthsimple, generic CSV/XLS/XLSX) "
            "or a prepared xPence filtered_data.xlsx — bank files are scrubbed automatically.")

        # ── Report output path ────────────────────────────────────────────────
        self._lbl(body, "Report output path", font=("Segoe UI", 9, "bold"), fg=MUTED)
        c2 = self._card(body, pady=(4, 10))
        c2i = tk.Frame(c2, bg=CARD, padx=12, pady=10); c2i.pack(fill="x")
        row2 = tk.Frame(c2i, bg=CARD); row2.pack(fill="x")
        self._entry(row2, self.out_path).pack(
            side="left", fill="x", expand=True, ipady=6, ipadx=6)
        self._round_btn(row2, "Save as", self._browse_out,
                        bg=CARD2, fg=TEXT, font=("Segoe UI", 9),
                        padx=16, pady=6).pack(side="left", padx=(8, 0))

        # ── Status bar ────────────────────────────────────────────────────────
        self.slbl     = self._status_frame(body, self.status_var)
        self.progress = self._progressbar(body)

        # ── Action area ───────────────────────────────────────────────────────
        BTN_PADY = 14
        BTN_FONT = ("Segoe UI", 11, "bold")

        # Export Scrubbed Data checkbox — always visible, unchecked by default
        self._chk_var = tk.BooleanVar(value=False)
        self.dl_chk = tk.Checkbutton(
            body,
            text="Export scrubbed data after generating report",
            variable=self._chk_var,
            font=("Segoe UI", 9), fg=MUTED, bg=BG,
            activebackground=BG, activeforeground=TEAL,
            selectcolor=BG, cursor="hand2",
            anchor="w",
        )
        self.dl_chk.pack(fill="x", pady=(4, 2))

        self.run_btn = RoundButton(
            body, "Generate Report  →", self._run,
            bg=GOLD, fg=GOLD_FG, font=BTN_FONT,
            padx=0, pady=BTN_PADY)
        self.run_btn.pack(fill="x", pady=(4, 0))

        self.view_btn = RoundButton(
            body, "Open Report  ↗", self._open,
            bg=GREEN, fg="#0d1117", font=BTN_FONT,
            padx=0, pady=BTN_PADY, state="disabled")
        # Not packed yet — packed in _ok() once report exists

    # ── helpers ───────────────────────────────────────────────────────────────
    def _st(self, msg, colour=MUTED):
        self.status_var.set(msg)
        self.slbl.config(fg=colour)

    def _check_analyzer(self):
        if not ANALYZER.exists():
            self._st(f"xpence_analyzer.py not found in {SCRIPT_DIR}", RED)
            self.run_btn.config(state="disabled")
        elif MERCHANT_JSON.exists():
            self._st(f"Ready  —  merchant list loaded ({MERCHANT_JSON.name})")
        else:
            self._st("Ready  —  no merchant_categories.json (will be created on first run)", GOLD)

    def _browse_csv(self):
        p = filedialog.askopenfilename(
            title="Select bank export or xPence data file",
            filetypes=[("Supported files", "*.csv *.xlsx *.xlsm *.xls *.ods"),
                       ("All files", "*.*")])
        if p:
            self.csv_path.set(p)
            if not self.out_path.get():
                self.out_path.set(str(Path(p).with_suffix("")) + "_report.html")
            self._st(f"Loaded: {Path(p).name}")
            # Hide View Report when a new file is selected
            self.view_btn.pack_forget()

    def _browse_out(self):
        p = filedialog.asksaveasfilename(title="Save report as",
                                         defaultextension=".html",
                                         filetypes=[("HTML", "*.html")])
        if p: self.out_path.set(p)

    @staticmethod
    def _is_already_xpence(filepath: str) -> bool:
        try:
            ext = os.path.splitext(filepath)[1].lower()
            if ext in (".xlsx", ".xlsm"):
                df = pd.read_excel(filepath, header=None, nrows=5)
            elif ext == ".csv":
                df = pd.read_csv(filepath, header=None, nrows=5)
            else:
                return False
            if df.shape[1] < 2:
                return False
            col0 = pd.to_numeric(df.iloc[:, 0], errors="coerce").dropna()
            return bool(not col0.empty and (col0 > 40000).all() and (col0 < 60000).all())
        except Exception:
            return False

    def _run(self):
        csv = self.csv_path.get().strip()
        out = self.out_path.get().strip()
        if not csv:
            messagebox.showwarning("No file", "Please select an input file first.")
            return
        if not Path(csv).exists():
            messagebox.showerror("Not found", f"Cannot find:\n{csv}"); return
        if not out:
            out = str(Path(csv).parent / "xpence_report.html")
            self.out_path.set(out)
        self.run_btn.config(state="disabled", text="Working...")
        self.view_btn.pack_forget()
        self.progress.start(10)
        self._st("Detecting file format...", ACCENT)
        want_scrubbed = self._chk_var.get()   # snapshot checkbox NOW before thread runs
        threading.Thread(target=self._worker, args=(csv, out, want_scrubbed), daemon=True).start()

    def _worker(self, input_path, out, want_scrubbed):
        tmp_scrubbed = None
        try:
            if self._is_already_xpence(input_path):
                self.after(0, self._st, "Already in xPence format — skipping scrub...", ACCENT)
                analysis_file = input_path
            else:
                self.after(0, self._st, "Scrubbing bank data...", PURPLE)
                base = os.path.splitext(out)[0]
                tmp_scrubbed = base + "_scrubbed_data.xlsx"
                scrub_file(input_path, tmp_scrubbed)
                n = len(pd.read_excel(tmp_scrubbed, header=None))
                self.after(0, self._st,
                           f"Scrubbed {n} transactions — generating report...", ACCENT)
                analysis_file = tmp_scrubbed

            # Single stable overrides file beside the analyzer — not per-report
            overrides_path = str(SCRIPT_DIR / "user_overrides.json")
            cmd = [sys.executable, str(ANALYZER),
                   "--file", analysis_file, "--output", out,
                   "--merchants", str(MERCHANT_JSON),
                   "--overrides", overrides_path]
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
            report_ok = Path(out).exists() and Path(out).stat().st_size > 1000
            if report_ok:
                self.after(0, self._ok, out, tmp_scrubbed, want_scrubbed)
            else:
                err = r.stderr.strip() or r.stdout.strip() or "Analyzer produced no output."
                self.after(0, self._err, err)
        except Exception:
            import traceback
            self.after(0, self._err, traceback.format_exc())

    def _ok(self, out, scrubbed_path, want_scrubbed):
        self.progress.stop()
        self.run_btn.config(state="normal", text="Generate Report  →")
        self._report_path = out

        if want_scrubbed and scrubbed_path and Path(scrubbed_path).exists():
            # Checkbox was checked — ask where to save the scrubbed file
            dest = filedialog.asksaveasfilename(
                title="Save scrubbed data",
                initialfile=Path(scrubbed_path).name,
                initialdir=str(Path(scrubbed_path).parent),
                defaultextension=".xlsx",
                filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")])
            if dest:
                shutil.copy2(scrubbed_path, dest)
            # Clean up temp file regardless
            try: Path(scrubbed_path).unlink()
            except OSError: pass
        elif scrubbed_path and Path(scrubbed_path).exists():
            # Checkbox was NOT checked — silently delete the temp scrubbed file
            try: Path(scrubbed_path).unlink()
            except OSError: pass

        self._chk_var.set(False)   # reset checkbox for next run
        self.view_btn.config(state="normal")
        self.view_btn.pack(fill="x", pady=(6, 0))
        self._st("Report ready.", GREEN)

    def _err(self, msg):
        self.progress.stop()
        self.run_btn.config(state="normal", text="Generate Report  →")
        self._st(f"Error: {msg.splitlines()[-1]}", RED)
        messagebox.showerror("Failed", msg)

    def _open(self):
        p = self._report_path
        if p and Path(p).exists():
            webbrowser.open(Path(p).as_uri())
        else:
            messagebox.showwarning("Not found", "Generate the report first.")



# ── Bank Scrubber Tab ─────────────────────────────────────────────────────────
class ScrubberTab(tk.Frame, _W):
    def __init__(self, master):
        super().__init__(master, bg=BG)
        self._files    = []
        self._results  = []      # [(disp_label, fname_stem, df, row_count)]
        self.fname_var = tk.StringVar(value="filtered_data")
        self.merge_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(
            value="Add one or more transaction files to begin.")
        self._build()

    def _build(self):
        # Scrollable container so all content is reachable on any screen size
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True)

        _vsb = tk.Scrollbar(outer, orient="vertical",
                            bg=CARD, troughcolor=CARD2, relief="flat",
                            width=0, bd=0)
        _vsb.pack(side="right", fill="y")

        _canvas = tk.Canvas(outer, bg=BG, highlightthickness=0,
                            yscrollcommand=_vsb.set)
        _canvas.pack(side="left", fill="both", expand=True)
        _vsb.config(command=_canvas.yview)

        body = tk.Frame(_canvas, bg=BG, padx=24, pady=12)
        _canvas_win = _canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_resize(event):
            _canvas.itemconfig(_canvas_win, width=event.width)
        _canvas.bind("<Configure>", _on_resize)

        def _on_body_resize(event):
            _canvas.configure(scrollregion=_canvas.bbox("all"))
        body.bind("<Configure>", _on_body_resize)

        # Mousewheel scroll (cross-platform)
        def _on_wheel(event):
            if event.num == 4 or event.delta > 0:
                _canvas.yview_scroll(-1, "units")
            else:
                _canvas.yview_scroll(1, "units")
        _canvas.bind_all("<MouseWheel>", _on_wheel)
        _canvas.bind_all("<Button-4>",   _on_wheel)
        _canvas.bind_all("<Button-5>",   _on_wheel)

        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(body, bg=BG); hdr.pack(fill="x", pady=(0, 2))
        tk.Label(hdr, text="Personal Data Scrubber and File Merger",
                 font=("Segoe UI", 13, "bold"), fg=TEXT, bg=BG).pack(side="left")
        self._round_btn(hdr, "? Manual Formatting Guide", self._show_hint,
                        bg=CARD2, fg=TEAL, font=("Segoe UI", 9),
                        padx=12, pady=5).pack(side="right")
        self._sublbl(body,
            "Strips personal data from raw bank exports and merges files into "
            "a single xPence-ready .xlsx.")
        self._divider(body, pady=(4, 6))

        # ── 1. What gets scrubbed ─────────────────────────────────────────────
        tk.Label(body, text="What gets scrubbed",
                 font=("Segoe UI", 11, "bold"), fg=WARN_YEL,
                 bg=BG, anchor="w").pack(fill="x", pady=(2, 1))
        tk.Label(body,
                 text="The Report Generator tab does this automatically when given a raw bank file.",
                 font=("Segoe UI", 8), fg=MUTED, bg=BG, anchor="w").pack(fill="x", pady=(0, 4))

        pii_card = tk.Frame(body, bg=CARD, highlightthickness=1,
                            highlightbackground=BORDER)
        pii_card.pack(fill="x", pady=(0, 8))
        pii_inner = tk.Frame(pii_card, bg=CARD, padx=12, pady=5)
        pii_inner.pack(fill="x")
        pii_items = [
            ("Card & account numbers",      "masked PANs, 13-19 digit sequences"),
            ("Account holder name",         "full name / holder columns dropped"),
            ("Account / institution IDs",   "account_id, branch, transit numbers"),
            ("IBAN / routing / sort codes", "international and domestic formats"),
            ("Phone numbers",              "10/11-digit patterns in descriptions"),
            ("PII-labelled columns",        "any column header matching PII keywords"),
        ]
        for title, detail in pii_items:
            row = tk.Frame(pii_inner, bg=CARD); row.pack(fill="x", pady=1)
            tk.Label(row, text="•", font=("Segoe UI", 8), fg=ACCENT,
                     bg=CARD, width=2).pack(side="left", anchor="n", pady=(2, 0))
            tk.Label(row, text=title,
                     font=("Segoe UI", 8, "bold"), fg=TEXT,
                     bg=CARD).pack(side="left")
            tk.Label(row, text=f"  —  {detail}",
                     font=("Segoe UI", 8), fg=MUTED,
                     bg=CARD).pack(side="left")

        # ── 2. Transaction files ──────────────────────────────────────────────
        hdr2 = tk.Frame(body, bg=BG); hdr2.pack(fill="x", pady=(0, 2))
        tk.Label(hdr2, text="Transaction files",
                 font=("Segoe UI", 9, "bold"), fg=MUTED, bg=BG,
                 anchor="w").pack(side="left")
        tk.Label(hdr2, text="  —  TD, CIBC, Wealthsimple, generic CSV / XLS / XLSX",
                 font=("Segoe UI", 8), fg=MUTED, bg=BG,
                 anchor="w").pack(side="left")
        lf = tk.Frame(body, bg=CARD, highlightthickness=1,
                      highlightbackground=BORDER)
        lf.pack(fill="x", pady=(3, 3))
        sb = tk.Scrollbar(lf, bg=CARD, troughcolor=CARD2, relief="flat", bd=0)
        self.file_list = tk.Listbox(
            lf, font=("Segoe UI", 9), bg=CARD2, fg=TEXT,
            selectbackground=ACCENT, selectforeground="#fff",
            relief="flat", bd=0, highlightthickness=0,
            activestyle="none", height=3, yscrollcommand=sb.set)
        sb.config(command=self.file_list.yview)
        self.file_list.pack(side="left", fill="both", expand=True,
                            padx=(8, 0), pady=5)
        sb.pack(side="right", fill="y")

        br = tk.Frame(body, bg=BG); br.pack(fill="x", pady=(3, 6))
        for label, cmd, fg in [
            ("+  Add Files",   self._add_files,       TEXT),
            ("Remove",         self._remove_selected, RED),
            ("Clear All",      self._clear_all,       MUTED),
        ]:
            self._round_btn(br, label, cmd, bg=CARD2, fg=fg,
                            font=("Segoe UI", 9), padx=12, pady=5
                            ).pack(side="left", padx=(0, 6))

        # ── 3. Output card — filename row + Scrub button inline ─────────────
        self._lbl(body, "Output", font=("Segoe UI", 9, "bold"), fg=MUTED)
        out_card = tk.Frame(body, bg=CARD, highlightthickness=1,
                            highlightbackground=BORDER)
        out_card.pack(fill="x", pady=(3, 6))
        out_inner = tk.Frame(out_card, bg=CARD, padx=12, pady=8)
        out_inner.pack(fill="x")

        tk.Checkbutton(
            out_inner,
            text="Merge all files into one combined file",
            variable=self.merge_var,
            font=("Segoe UI", 9), fg=TEXT, bg=CARD,
            activebackground=CARD, activeforeground=TEXT,
            selectcolor=CARD2, cursor="hand2",
        ).pack(anchor="w")

        # Filename row — entry left, Scrub button right (inline in output card)
        fn_row = tk.Frame(out_inner, bg=CARD); fn_row.pack(fill="x", pady=(8, 0))
        tk.Label(fn_row, text="Filename:", font=("Segoe UI", 9),
                 fg=MUTED, bg=CARD, width=9, anchor="w").pack(side="left")
        fn_e = tk.Entry(fn_row, textvariable=self.fname_var,
                        font=("Segoe UI", 10), bg=CARD2, fg=TEXT,
                        insertbackground=TEXT, relief="flat", bd=0,
                        highlightthickness=1,
                        highlightbackground=BORDER, highlightcolor=ACCENT)
        fn_e.pack(side="left", fill="x", expand=True, ipady=5, ipadx=6)
        tk.Label(fn_row, text=".xlsx", font=("Segoe UI", 9),
                 fg=MUTED, bg=CARD, anchor="w").pack(side="left", padx=(4, 8))

        # Scrub button lives here — right of filename entry
        self.scrub_btn = self._round_btn(
            fn_row, "Scrub & Convert  →", self._run,
            bg=GOLD, fg=GOLD_FG, font=("Segoe UI", 10, "bold"),
            padx=18, pady=6)
        self.scrub_btn.pack(side="left")

        # ── Status + progress ─────────────────────────────────────────────────
        self.slbl     = self._status_frame(body, self.status_var)
        self.progress = self._progressbar(body)

        # ── Results card (hidden until scrub completes) ───────────────────────
        self.results_card = tk.Frame(body, bg=CARD, highlightthickness=1,
                                     highlightbackground=BORDER)
        self.results_inner = tk.Frame(self.results_card, bg=CARD)
        self.results_inner.pack(fill="x", padx=1, pady=1)

        # Bottom padding only — Export button removed (Scrub & Convert handles download)
        tk.Frame(body, bg=BG, height=8).pack(fill="x")

    # ── helpers ───────────────────────────────────────────────────────────────
    def _st(self, msg, colour=MUTED):
        self.status_var.set(msg); self.slbl.config(fg=colour)

    # ── Format guide modal ────────────────────────────────────────────────────
    def _show_hint(self):
        win = tk.Toplevel(self)
        win.title("xPence — Manual Formatting Guide")
        win.configure(bg=BG); win.resizable(True, True); win.grab_set()
        self.update_idletasks()
        W, H = 660, 680
        px = self.winfo_rootx() + self.winfo_width()  // 2
        py = self.winfo_rooty() + self.winfo_height() // 2
        win.geometry(f"{W}x{H}+{px - W//2}+{py - H//2}")
        win.minsize(500, 500)

        # ── Fixed header ──────────────────────────────────────────────────────
        hf = tk.Frame(win, bg=CARD2, padx=24, pady=14); hf.pack(fill="x")
        tk.Label(hf, text="Manual Formatting Guide",
                 font=("Segoe UI", 13, "bold"), fg=TEXT, bg=CARD2).pack(anchor="w")
        tk.Label(hf,
                 text="The Transaction Data Merger tab handles this automatically. "
                      "Use this guide only if you want to prepare a filtered_data.xlsx by hand.",
                 font=("Segoe UI", 9), fg=MUTED, bg=CARD2,
                 wraplength=580, justify="left").pack(anchor="w", pady=(4, 0))

        # ── Scrollable body ───────────────────────────────────────────────────
        scroll_outer = tk.Frame(win, bg=BG); scroll_outer.pack(fill="both", expand=True)
        vsb = tk.Scrollbar(scroll_outer, orient="vertical",
                           bg=CARD, troughcolor=CARD2, relief="flat", bd=0)
        vsb.pack(side="right", fill="y")
        canvas = tk.Canvas(scroll_outer, bg=BG, highlightthickness=0,
                           yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.config(command=canvas.yview)

        scr = tk.Frame(canvas, bg=BG, padx=24, pady=12)
        _win_id = canvas.create_window((0, 0), window=scr, anchor="nw")

        canvas.bind("<Configure>", lambda e: canvas.itemconfig(_win_id, width=e.width))
        scr.bind("<Configure>",    lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        def _on_wheel(event):
            canvas.yview_scroll(-1 if (event.num == 4 or event.delta > 0) else 1, "units")
        canvas.bind_all("<MouseWheel>", _on_wheel)
        canvas.bind_all("<Button-4>",   _on_wheel)
        canvas.bind_all("<Button-5>",   _on_wheel)
        win.bind("<Destroy>", lambda e: (canvas.unbind_all("<MouseWheel>"),
                                         canvas.unbind_all("<Button-4>"),
                                         canvas.unbind_all("<Button-5>")))

        # ── Helpers ───────────────────────────────────────────────────────────
        def section(t, colour=TEAL):
            tk.Frame(scr, bg=BORDER, height=1).pack(fill="x", pady=(12, 5))
            tk.Label(scr, text=t, font=("Segoe UI", 10, "bold"),
                     fg=colour, bg=BG, anchor="w").pack(fill="x")

        def para(t, fg=TEXT, font=("Segoe UI", 9)):
            tk.Label(scr, text=t, font=font, fg=fg, bg=BG,
                     anchor="w", wraplength=560, justify="left").pack(fill="x", pady=(1, 0))

        def bullet(title, body_txt):
            row = tk.Frame(scr, bg=BG); row.pack(fill="x", pady=1)
            tk.Label(row, text="•", font=("Segoe UI", 9), fg=ACCENT,
                     bg=BG, width=2).pack(side="left", anchor="n")
            tf = tk.Frame(row, bg=BG); tf.pack(side="left")
            tk.Label(tf, text=title, font=("Segoe UI", 9, "bold"),
                     fg=TEXT, bg=BG, anchor="w").pack(fill="x")
            tk.Label(tf, text=body_txt, font=("Segoe UI", 8),
                     fg=MUTED, bg=BG, anchor="w").pack(fill="x")

        # ── File requirements ─────────────────────────────────────────────────
        section("File requirements")
        bullet("Format",  "Excel Workbook (.xlsx) — save from Excel or LibreOffice")
        bullet("Headers", "None — row 1 is the first data row")
        bullet("Columns", "Exactly 4 columns: Date, Merchant / Tx Name, Credit, Debit")

        # ── Column layout — table ─────────────────────────────────────────────
        section("Column layout")
        para("4 columns, no header row.  Leave Credit or Debit blank if not applicable.",
             fg=MUTED, font=("Segoe UI", 8))

        HDR_BG  = CARD
        ROW_BGS = [CARD2, CARD]
        CR_FG   = "#58a6ff"   # blue  — credit / payment
        DB_FG   = "#3fb950"   # green — debit / expense

        tbl = tk.Frame(scr, bg=BORDER); tbl.pack(fill="x", pady=(6, 0))

        # Header row: Date | Merchant / Tx Name | Credit | Debit
        hdr_row = tk.Frame(tbl, bg=HDR_BG); hdr_row.pack(fill="x")
        for txt, expand, anchor, w in [
            ("Date",                 False, "w", 14),
            ("Merchant / Tx Name",  True,  "w", 0),
            ("Credit",              False, "e", 10),
            ("Debit",               False, "e", 10),
        ]:
            tk.Label(hdr_row, text=txt,
                     font=("Segoe UI", 8, "bold"), fg=ACCENT, bg=HDR_BG,
                     anchor=anchor, padx=10, pady=6,
                     width=w if not expand else 0,
                     ).pack(side="left", fill="x", expand=expand)

        tk.Frame(tbl, bg=BORDER, height=1).pack(fill="x")

        # 2 sample rows: (date, merchant, credit_str, debit_str)
        # Tim Hortons = credit (money in), PAYMENT = debit (money out)
        sample_rows = [
            ("2025-12-10", "Tim Hortons #42 — TORONTO ON",  "6.60",   "",       CR_FG, DB_FG),
            ("2025-12-09", "PAYMENT – THANK YOU",           "",       "250.00", CR_FG, DB_FG),
        ]
        for ri, (date_s, desc_s, cr_s, db_s, cr_fg, db_fg) in enumerate(sample_rows):
            rbg = ROW_BGS[ri % 2]
            dr = tk.Frame(tbl, bg=rbg); dr.pack(fill="x")
            tk.Label(dr, text=date_s,
                     font=("Courier New", 8), fg=MUTED, bg=rbg,
                     anchor="w", padx=10, pady=5, width=14).pack(side="left")
            tk.Label(dr, text=desc_s,
                     font=("Segoe UI", 8), fg=TEXT, bg=rbg,
                     anchor="w", padx=10, pady=5).pack(side="left", fill="x", expand=True)
            tk.Label(dr, text=cr_s,
                     font=("Courier New", 8, "bold"), fg=cr_fg, bg=rbg,
                     anchor="e", padx=10, pady=5, width=10).pack(side="left")
            tk.Label(dr, text=db_s,
                     font=("Courier New", 8, "bold"), fg=db_fg, bg=rbg,
                     anchor="e", padx=10, pady=5, width=10).pack(side="left")

        para("Date format: yyyy-mm-dd.  Credit = money in (e.g. purchases charged to card).  Debit = money out (e.g. payments made).",
             fg=MUTED, font=("Segoe UI", 7))

        # ── Do NOT include ────────────────────────────────────────────────────
        section("Do NOT include", colour=RED)
        for item in ["Header rows", "Card or account numbers", "Account holder name",
                     "Balance column", "Currency column", "Bank-internal reference codes"]:
            tk.Label(scr, text=f"  ✗  {item}", font=("Segoe UI", 9),
                     fg=RED, bg=BG, anchor="w").pack(fill="x")

        tk.Frame(scr, bg=BG, height=8).pack()   # bottom breathing room

        # ── Fixed footer ──────────────────────────────────────────────────────
        tk.Frame(win, bg=BORDER, height=1).pack(fill="x")
        foot = tk.Frame(win, bg=CARD2, padx=24, pady=10); foot.pack(fill="x")
        self._round_btn(foot, "Close", win.destroy,
                        bg=ACCENT, fg="#fff",
                        font=("Segoe UI", 10, "bold"),
                        padx=24, pady=7).pack(side="right")

    # ── File management ───────────────────────────────────────────────────────
    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select transaction file(s)",
            filetypes=[("Transaction files", "*.csv *.xls *.xlsx *.xlsm *.ods"),
                       ("All files", "*.*")])
        for p in paths:
            pp = Path(p)
            if pp not in self._files:
                self._files.append(pp)
                self.file_list.insert(
                    tk.END, f"  {pp.name}   [{self._bank_label(pp)}]")
        if paths:
            self._st(f"{len(self._files)} file(s) ready.")
            self.results_card.pack_forget()
            self._results = []

    def _bank_label(self, p: Path) -> str:
        n = p.name.lower()
        for k, lbl in {"td": "TD Bank", "cibc": "CIBC",
                       "wealthsimple": "Wealthsimple", "rbc": "RBC",
                       "bmo": "BMO", "scotiabank": "Scotiabank",
                       "tangerine": "Tangerine", "eqbank": "EQ Bank"}.items():
            if k in n: return lbl
        return "auto-detect"

    def _remove_selected(self):
        for idx in reversed(list(self.file_list.curselection())):
            self.file_list.delete(idx); self._files.pop(idx)
        self._st(f"{len(self._files)} file(s) ready." if self._files
                 else "Add one or more bank export files to begin.")
        self._results = []

    def _clear_all(self):
        self.file_list.delete(0, tk.END); self._files.clear()
        self._results = []; self.results_card.pack_forget()
        self._st("Add one or more bank export files to begin.")

    # ── Scrub ─────────────────────────────────────────────────────────────────
    def _run(self):
        if not self._files:
            messagebox.showwarning("No files", "Add at least one bank export file.")
            return
        missing = [f for f in self._files if not f.exists()]
        if missing:
            messagebox.showerror("File not found",
                                 "Cannot find:\n" + "\n".join(str(m) for m in missing))
            return
        fname = self.fname_var.get().strip() or "filtered_data"
        self.scrub_btn.config(state="disabled", text="Scrubbing...")
        self.results_card.pack_forget()
        self.progress.start(10)
        self._st("Scrubbing personal data...", PURPLE)
        threading.Thread(target=self._worker, args=(fname,), daemon=True).start()

    def _worker(self, fname: str):
        """Scrub files into memory — nothing written to disk until Download."""
        try:
            frames, stats = [], {}
            for fp in self._files:
                try:
                    raw   = _read_raw(str(fp))
                    clean = _parse(str(fp), raw)
                    clean = clean[clean["date_serial"] > 0]
                    frames.append((fp, clean))
                    stats[str(fp)] = len(clean)
                except Exception as exc:
                    stats[str(fp)] = f"error: {exc}"

            if not frames:
                self.after(0, self._err, "No valid transactions found in any file.")
                return

            if self.merge_var.get():
                merged = (pd.concat([f for _, f in frames], ignore_index=True)
                            .sort_values("date_serial", ascending=False)
                            .reset_index(drop=True))
                total = sum(v for v in stats.values() if isinstance(v, int))
                results = [("Merged output", fname, merged, total)]
            else:
                results = []
                for fp, clean in frames:
                    n = stats.get(str(fp), 0)
                    stem = fp.stem + "_scrubbed"
                    results.append((fp.name, stem, clean,
                                    n if isinstance(n, int) else 0))

            self.after(0, self._ok, results, stats)
        except Exception:
            import traceback
            self.after(0, self._err, traceback.format_exc())

    def _ok(self, results, stats):
        self.progress.stop()
        self.scrub_btn.config(state="normal", text="Scrub & Convert  →")
        self._results = results

        for w in self.results_inner.winfo_children():
            w.destroy()

        total = sum(r[3] for r in results)
        errors = [fp for fp, v in stats.items() if not isinstance(v, int)]
        self._st(
            f"{'Done with errors — ' if errors else ''}"
            f"{total} transactions ready across {len(results)} output(s).",
            GOLD if errors else GREEN)

        for disp_label, fname_stem, df, row_count in results:
            row = tk.Frame(self.results_inner, bg=CARD2); row.pack(fill="x", pady=1)
            inner = tk.Frame(row, bg=CARD2, padx=12, pady=7); inner.pack(fill="x")
            tk.Label(inner, text="●", font=("Segoe UI", 9),
                     fg=GREEN if row_count > 0 else RED,
                     bg=CARD2).pack(side="left", padx=(0, 8))
            info = tk.Frame(inner, bg=CARD2); info.pack(side="left", fill="x", expand=True)
            tk.Label(info, text=disp_label,
                     font=("Segoe UI", 9, "bold"), fg=TEXT,
                     bg=CARD2, anchor="w").pack(fill="x")
            tk.Label(info,
                     text=f"{row_count} transactions  ·  saves as  {fname_stem}.xlsx",
                     font=("Segoe UI", 8), fg=MUTED,
                     bg=CARD2, anchor="w").pack(fill="x")

        self.results_card.pack(fill="x", pady=(6, 0))
        self._download()

    def _err(self, msg):
        self.progress.stop()
        self.scrub_btn.config(state="normal", text="Scrub & Convert  →")
        self._st(f"Error: {msg.splitlines()[-1]}", RED)
        messagebox.showerror("Scrubber failed", msg)

    # ── Download ──────────────────────────────────────────────────────────────
    def _download(self):
        if not self._results:
            messagebox.showwarning("Nothing to download", "Scrub some files first.")
            return
        if self.merge_var.get():
            _, fname_stem, df, _ = self._results[0]
            dest = filedialog.asksaveasfilename(
                title="Save scrubbed data as",
                initialfile=f"{fname_stem}.xlsx",
                defaultextension=".xlsx",
                filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")])
            if not dest: return
            _write_xpence_xlsx(df, dest)
            self._st(f"Saved: {Path(dest).name}", GREEN)
        else:
            folder = filedialog.askdirectory(
                title="Choose folder to save scrubbed files")
            if not folder: return
            saved = 0
            for _, fname_stem, df, _ in self._results:
                _write_xpence_xlsx(df, os.path.join(folder, f"{fname_stem}.xlsx"))
                saved += 1
            self._st(f"Saved {saved} file(s) to {Path(folder).name}/", GREEN)


# ── Main window ───────────────────────────────────────────────────────────────
class XPenceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("xPence")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._build()
        self._centre(740, 660)

    def _build(self):
        # Header bar
        hdr = tk.Frame(self, bg=CARD2, padx=24, pady=12); hdr.pack(fill="x")
        tk.Label(hdr, text="xPence", font=("Georgia", 20, "italic"),
                 fg=ACCENT, bg=CARD2).pack(side="left")
        tk.Label(hdr, text="  Personal Finance",
                 font=("Segoe UI", 11), fg=MUTED, bg=CARD2).pack(side="left", pady=2)

        # Nav strip
        nav_strip = tk.Frame(self, bg=CARD, padx=16, pady=8); nav_strip.pack(fill="x")

        # Content frames
        self._report_frame   = tk.Frame(self, bg=BG)
        self._scrubber_frame = tk.Frame(self, bg=BG)
        self.report_tab   = ReportTab(self._report_frame)
        self.scrubber_tab = ScrubberTab(self._scrubber_frame)
        self.report_tab.pack(fill="both", expand=True)
        self.scrubber_tab.pack(fill="both", expand=True)

        NavPill(
            nav_strip,
            ["  Report Generator  ", "  Transaction Data Merger  "],
            [self._report_frame, self._scrubber_frame],
        ).pack(side="left")

    def _centre(self, w, h):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if sys.platform == "darwin":
        try:
            subprocess.run(
                ["osascript", "-e",
                 'tell app "Finder" to set frontmost of process "Python" to true'],
                capture_output=True)
        except Exception:
            pass
    XPenceApp().mainloop()
