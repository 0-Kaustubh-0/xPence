"""
xPence Financial Analyzer — v5
================================
Generates a self-contained HTML expense report with three views:
  • Review         — flag & interactively reclassify questionable transactions
  • All Spending   — category bar chart + full transaction list
  • Monthly        — month picker, pivot table, pie chart, filtered transactions

All charts and totals update live when Review decisions change.

Three-file architecture
-----------------------
  merchant_categories.json   Master reference list (internet-downloadable).
                             Contains category_keywords, category_colors,
                             category_renames, and the community merchants table.

  user_overrides.json        Personal overrides file (never shared).
                             Contains overrides, user-specific merchants,
                             custom_categories, and subscriptions.

  xpence_analyzer.py         This script — no hardcoded keyword data.

Classification priority:
  user overrides > user merchants > master merchants
  > category_keywords > custom_categories > Other / Uncategorised
"""

import argparse
import glob
import json
import os
import statistics
import sys
from collections import defaultdict
from datetime import datetime

# Ensure UTF-8 output on Windows (cp1252 terminals choke on arrow/special chars)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import pandas as pd

# ---------------------------------------------------------------------------
# Runtime state — populated by _load_master() at startup
# ---------------------------------------------------------------------------
CATEGORY_KEYWORDS: dict = {}   # loaded from merchant_categories.json
CATEGORIES:        list = []   # ordered list of built-in category names
_KW_FLAT:          list = []   # [(keyword, category), ...] pre-flattened
_CAT_INDEX:        dict = {}   # {category: index} for O(1) colour lookup
CAT_COLORS:        list = []   # one hex colour per built-in category
CATEGORY_RENAMES:  dict = {}   # stale-name migration map

# Fallback values used when the master file is absent
_DEFAULT_CAT_COLORS = [
    "#2563eb","#0d9488","#ea580c","#7c3aed","#0ea5e9","#db2777",
    "#16a34a","#f59e0b","#6366f1","#64748b","#c026d3",
]
_DEFAULT_CATEGORY_RENAMES = {
    "Restaurants":                       "Restaurants, Pubs & Cafes",
    "Hotel, Entertainment & Recreation": "Entertainment & Recreation",
    "Home & Office Improvement":         "Electronics, Home & Office Improvement",
}

def _load_master(path: str | None) -> dict:
    """
    Load merchant_categories.json (master reference list).

    Populates the module-level CATEGORY_KEYWORDS, CATEGORIES, _KW_FLAT,
    _CAT_INDEX, CAT_COLORS, and CATEGORY_RENAMES globals so the rest of the
    module works identically whether the master file is present or not.

    Returns the full parsed dict so callers can also access the 'merchants'
    community lookup table.
    """
    global CATEGORY_KEYWORDS, CATEGORIES, _KW_FLAT, _CAT_INDEX
    global CAT_COLORS, CATEGORY_RENAMES

    _empty_master: dict = {
        "category_keywords":  {},
        "category_colors":    {},
        "category_renames":   _DEFAULT_CATEGORY_RENAMES,
        "merchants":          {},
    }

    if not path or not os.path.exists(path):
        print(f"  [master] File not found: {path!r} — using keyword classification only")
        CATEGORY_KEYWORDS = {}
        CATEGORIES        = []
        _KW_FLAT          = []
        _CAT_INDEX        = {}
        CAT_COLORS        = list(_DEFAULT_CAT_COLORS)
        CATEGORY_RENAMES  = dict(_DEFAULT_CATEGORY_RENAMES)
        return _empty_master

    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception as exc:
        print(f"  [master] WARNING: could not read {path}: {exc}")
        CATEGORY_KEYWORDS = {}
        CATEGORIES        = []
        _KW_FLAT          = []
        _CAT_INDEX        = {}
        CAT_COLORS        = list(_DEFAULT_CAT_COLORS)
        CATEGORY_RENAMES  = dict(_DEFAULT_CATEGORY_RENAMES)
        return _empty_master

    CATEGORY_KEYWORDS = data.get("category_keywords", {})
    CATEGORIES        = list(CATEGORY_KEYWORDS.keys())
    _KW_FLAT          = [
        (kw, cat)
        for cat, kws in CATEGORY_KEYWORDS.items()
        for kw in kws
    ]
    _CAT_INDEX        = {cat: i for i, cat in enumerate(CATEGORIES)}
    # Merge file colours with defaults so missing entries fall back gracefully
    file_colors       = data.get("category_colors", {})
    CAT_COLORS        = [
        file_colors.get(cat, _DEFAULT_CAT_COLORS[i % len(_DEFAULT_CAT_COLORS)])
        for i, cat in enumerate(CATEGORIES)
    ]
    CATEGORY_RENAMES  = data.get("category_renames", dict(_DEFAULT_CATEGORY_RENAMES))

    schema = data.get("_schema", "")
    n_kw   = sum(len(v) for v in CATEGORY_KEYWORDS.values())
    n_m    = len(data.get("merchants", {}))
    print(f"  [master] Loaded {path!r}  "
          f"({len(CATEGORIES)} categories, {n_kw} keywords, {n_m} merchant entries)")
    return data

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
# CSV/Excel column header hints — defined once, shared by both loaders
_DATE_H   = ["date","transaction date","trans date","posted date","value date"]
_NAME_H   = ["description","transaction","merchant","name","narration","particulars","details","memo","payee"]
_DEBIT_H  = ["debit","spent","amount","withdrawal","charge","payment out","dr","expense","paid"]
_CREDIT_H = ["credit","deposit","payment in","cr","received","refund","inflow"]
_ACCTYPE_H = ["account type","account_type","acct type","acct_type","card type",
              "card_type","account","account name","account category"]

# Default account type assumed when no account-type column is present in the
# source data (see build_report()). Most personal-finance exports fed into
# xPence originate from a credit-card statement, so "Credit" is the safest
# assumption when nothing else is known.
DEFAULT_ACCOUNT_TYPE = "Credit"

def _best_match(candidates: list, hints: list) -> str | None:
    """Return the best matching candidate for any hint, or None."""
    lmap = {c.lower().strip(): c for c in candidates}
    for h in hints:
        if h in lmap:
            return lmap[h]
    for h in hints:
        for lc, orig in lmap.items():
            if h in lc or lc in h:
                return orig
    return None

def classify(
    name: str,
    user_map:   dict | None = None,
    master_map: dict | None = None,
    _valid_cats: set | None = None,
) -> str:
    """Classify a merchant name.

    Priority:
      1. user_map['overrides']           (user-confirmed decisions)
      2. user_map['merchants']           (user-specific auto-learned)
      3. master_map['merchants']         (community/downloaded lookup)
      4. _KW_FLAT                        (category_keywords from master file)
      5. user_map['custom_categories']   (user-defined keyword categories)
      6. "Other / Uncategorised"

    Pass ``_valid_cats`` (a pre-built set) when calling in a tight loop to avoid
    rebuilding it on every invocation.
    """
    if _valid_cats is None:
        _valid_cats = (
            set(CATEGORIES)
            | {"Other / Uncategorised"}
            | {cc["name"] for cc in (user_map or {}).get("custom_categories", [])}
        )

    # 1. User overrides (highest priority)
    if user_map:
        ov = user_map.get("overrides", {})
        if name in ov and (t := ov[name]) in _valid_cats:
            return t
        # 2. User-specific merchants
        me = user_map.get("merchants", {})
        if name in me and (t := me[name]) in _valid_cats:
            return t

    # 3. Master merchant lookup
    if master_map:
        mm = master_map.get("merchants", {})
        if name in mm and (t := mm[name]) in _valid_cats:
            return t

    # 4. Built-in keyword matching (loaded from master file)
    low = name.lower() if isinstance(name, str) else ""
    for kw, cat in _KW_FLAT:
        if kw in low:
            return cat

    # 5. User custom categories
    if user_map:
        for cc in user_map.get("custom_categories", []):
            for kw in cc.get("keywords", []):
                if kw.lower() in low:
                    return cc["name"]

    return "Other / Uncategorised"


def matches_expense_keyword(name: str) -> tuple[bool, str | None]:
    """Return (True, category) if *name* matches any known-expense keyword, else (False, None)."""
    low = name.lower() if isinstance(name, str) else ""
    # Exclude catch-all financial services keywords that are clearly credits
    skip_cats = {"Professional & Financial Services"}
    for cat, kws in CATEGORY_KEYWORDS.items():
        if cat in skip_cats:
            continue
        for kw in kws:
            if kw in low:
                return True, cat
    return False, None

# ---------------------------------------------------------------------------
# User overrides JSON  –  load / save
# ---------------------------------------------------------------------------
USER_OVERRIDES_SCHEMA = {
    "_schema":  "xpence-user-overrides-v2",
    "_note": (
        "User-specific overrides for xPence. Takes priority over merchant_categories.json.\n"
        "\n"
        "  overrides          Explicit user category decisions (highest priority).\n"
        "                     Written by category re-assignments in the HTML report.\n"
        "\n"
        "  merchants          User-specific auto-learned per-name mappings.\n"
        "                     Updated on every --file run. Wins over master merchants.\n"
        "\n"
        "  debit_credit_flips Debit/Credit flip decisions keyed by transaction signature\n"
        "                     'YYYY-MM-DD|merchant name|amount'. Applied across all reports\n"
        "                     containing the same transaction.\n"
        "\n"
        "  flagged            Transactions manually flagged for re-assignment, keyed by\n"
        "                     transaction signature 'YYYY-MM-DD|merchant name|amount'.\n"
        "\n"
        "  custom_categories  User-defined category definitions [{name, keywords}].\n"
        "                     Matched after category_keywords, before 'Other'.\n"
        "\n"
        "  subscriptions      [{name, pinned, dismissed}]\n"
        "                     pinned=true   -> always shown in Subscriptions tab.\n"
        "                     dismissed=true -> never shown.\n"
        "\n"
        "Priority: overrides > merchants (user) > merchants (master)"
        " > category_keywords > custom_categories > Other"
    ),
    "overrides":           {},
    "merchants":           {},
    "debit_credit_flips":  {},
    "flagged":             {},
    "custom_categories":   [],
    "subscriptions":       [],
}

# Colour assigned to ALL custom categories — always dark grey
CUSTOM_CAT_COLOR = "#3d3d3d"


def load_user_overrides(path: str | None) -> dict:
    """Load user_overrides.json; migrate any stale category names."""
    _empty: dict = {
        "overrides": {}, "merchants": {}, "debit_credit_flips": {}, "flagged": {},
        "custom_categories": [], "subscriptions": []
    }
    if not path or not os.path.exists(path):
        return _empty
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)

        # Legacy migration: old single-file format
        if "overrides" not in data and "merchants" not in data:
            print("  [user overrides] Migrating legacy flat format -> new schema")
            return {"overrides": dict(data), "merchants": {}, "debit_credit_flips": {},
                    "flagged": {}, "custom_categories": [], "subscriptions": []}

        result = {
            "overrides":          data.get("overrides", {}),
            "merchants":          data.get("merchants", {}),
            "debit_credit_flips": data.get("debit_credit_flips", {}),
            "flagged":            data.get("flagged", {}),
            "custom_categories":  data.get("custom_categories", []),
            "subscriptions":      data.get("subscriptions", []),
        }
        # Migrate stale category names
        migrated = 0
        for section in ("overrides", "merchants"):
            for merchant, cat in list(result[section].items()):
                if cat in CATEGORY_RENAMES:
                    result[section][merchant] = CATEGORY_RENAMES[cat]
                    migrated += 1
        if migrated:
            print(f"  [user overrides] Migrated {migrated} stale category name(s)")
        return result
    except Exception as exc:
        print(f"  [user overrides] WARNING: could not read {path}: {exc}")
        return _empty


def save_user_overrides(path: str, user_map: dict) -> None:
    """Write user_overrides.json with schema header, sorted keys."""
    out = dict(USER_OVERRIDES_SCHEMA)
    out["overrides"]          = dict(sorted(user_map["overrides"].items()))
    out["merchants"]          = dict(sorted(user_map["merchants"].items()))
    out["debit_credit_flips"] = dict(sorted(user_map.get("debit_credit_flips", {}).items()))
    out["flagged"]            = dict(sorted(user_map.get("flagged", {}).items()))
    out["custom_categories"]  = user_map.get("custom_categories", [])
    out["subscriptions"]      = user_map.get("subscriptions", [])
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Legacy alias — kept so any external tooling using load_merchant_json()
# continues to work. Calls load_user_overrides() under the hood.
# ---------------------------------------------------------------------------
def load_merchant_json(path: str | None) -> dict:
    """Deprecated alias for load_user_overrides(). Use that instead."""
    return load_user_overrides(path)


def save_merchant_json(path: str, merchant_map: dict) -> None:
    """Deprecated alias for save_user_overrides(). Use that instead."""
    save_user_overrides(path, merchant_map)


def update_merchant_json_from_xpr(xpr_path, overrides_path):
    """
    LEGACY / FALLBACK path only.

    The primary way to update user_overrides.json is now the
    'Save Overrides' button in the HTML report, which downloads a
    ready-to-use user_overrides.json directly — no --xpr needed.

    This function is kept for backwards-compatibility with older .xpr files
    that pre-date the direct-download approach.  It reads catOverrides and
    customCategories out of a .xpr and writes them into user_overrides.json.

    Conflicts: newer value from the xpr wins (most-recent review wins).
    """
    try:
        with open(xpr_path, "r", encoding="utf-8") as fh:
            xpr = json.load(fh)
    except Exception as e:
        print(f"  ERROR reading xpr file: {e}")
        return

    # If the xpr was produced by 'Save Overrides' (direct-JSON era), it will
    # not have a review block — nothing to do.
    review = xpr.get("review", {})
    if not review:
        print("  .xpr has no review block — nothing to merge.")
        return

    cat_overrides = review.get("catOverrides", {})
    rows          = xpr.get("raw", {}).get("rows", [])
    user_map      = load_user_overrides(overrides_path)
    pushed        = 0
    conflicts     = 0

    if cat_overrides and rows:
        idx_to_name = {str(r["idx"]): r["name"] for r in rows}
        for idx_str, new_cat in cat_overrides.items():
            name = idx_to_name.get(str(idx_str))
            if not name or not new_cat:
                continue
            existing = user_map["overrides"].get(name)
            if existing and existing != new_cat:
                print(f"  [override conflict] '{name}': '{existing}' -> '{new_cat}' (xpr wins)")
                conflicts += 1
            user_map["overrides"][name] = new_cat
            pushed += 1
    else:
        print("  No category overrides found in xpr.")

    xpr_custom = review.get("customCategories", [])
    if xpr_custom:
        existing_names = {c["name"] for c in user_map.get("custom_categories", [])}
        for xc in xpr_custom:
            if not isinstance(xc, dict) or not xc.get("name"):
                continue
            if xc["name"] not in existing_names:
                user_map.setdefault("custom_categories", []).append(xc)
                existing_names.add(xc["name"])
            else:
                for cc in user_map["custom_categories"]:
                    if cc["name"] == xc["name"]:
                        cc["keywords"] = xc.get("keywords", cc.get("keywords", []))
                        break

    save_user_overrides(overrides_path, user_map)
    if pushed:
        print(f"  Pushed {pushed} override(s) to {overrides_path}"
              + (f"  ({conflicts} conflict(s) resolved)" if conflicts else ""))


def find_xpence_file(directory: str = ".") -> str | None:
    """Discover an xPence CSV/Excel file in *directory*."""
    for pat in ("xpence*", "xPence*", "Xpence*", "XPENCE*"):
        matches = glob.glob(os.path.join(directory, pat))
        if matches:
            return matches[0]
    for pat in ("*.xlsx", "*.xls", "*.xlsm"):
        matches = glob.glob(os.path.join(directory, pat))
        if matches:
            return matches[0]
    return None

def load_csv(filepath: str) -> tuple[pd.DataFrame, dict]:
    """Load CSV, XLSX, XLS, or ODS — returns (df, col_map)."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls", ".ods"):
        return _load_excel(filepath)
    return _load_csv_inner(filepath)


def _load_excel(filepath: str) -> tuple[pd.DataFrame, dict]:
    ext = os.path.splitext(filepath)[1].lower()
    engine = "xlrd" if ext == ".xls" else ("odf" if ext == ".ods" else "openpyxl")
    try:
        df = pd.read_excel(filepath, engine=engine, header=0)
        if df.shape[1] >= 2:
            cols = list(df.columns)
            col = {
                "date":         _best_match(cols, _DATE_H),
                "name":         _best_match(cols, _NAME_H),
                "debit":        _best_match(cols, _DEBIT_H),
                "credit":       _best_match(cols, _CREDIT_H),
                "account_type": _best_match(cols, _ACCTYPE_H),
            }
            if col["date"] and col["name"]:
                return df, col
    except Exception:
        pass
    try:
        df = pd.read_excel(filepath, engine=engine, header=None)
        n = df.shape[1]
        df.columns = range(n)
        col = {
            "date":         0,
            "name":         1,
            "debit":        2 if n > 2 else None,
            "credit":       3 if n > 3 else None,
            # The data-scrubbing tool (xpence_gui.py) writes account type as
            # the 5th positional column when it can determine one.
            "account_type": 4 if n > 4 else None,
        }
        return df, col
    except Exception as exc:
        raise ValueError(f"Could not parse Excel file '{filepath}': {exc}") from exc


def _load_csv_inner(filepath: str) -> tuple[pd.DataFrame, dict]:
    for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
        for sep in (",", ";", "\t"):
            try:
                df = pd.read_csv(
                    filepath, sep=sep, encoding=enc, header=0,
                    thousands=",", skipinitialspace=True,
                )
                if df.shape[1] >= 3:
                    cols = list(df.columns)
                    col = {
                        "date":         _best_match(cols, _DATE_H),
                        "name":         _best_match(cols, _NAME_H),
                        "debit":        _best_match(cols, _DEBIT_H),
                        "credit":       _best_match(cols, _CREDIT_H),
                        "account_type": _best_match(cols, _ACCTYPE_H),
                    }
                    if col["date"] and col["name"]:
                        return df, col
            except Exception:
                pass
    for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
        try:
            df = pd.read_csv(
                filepath, header=None,
                names=["date", "description", "debit", "credit"],
                encoding=enc, skipinitialspace=True,
            )
            return df, {"date": "date", "name": "description", "debit": "debit", "credit": "credit"}
        except Exception:
            pass
    raise ValueError(f"Could not parse '{filepath}'")

def clean_amount(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .str.replace(r"[$,\xa3\u20ac\u20b9()\s]", "", regex=True)
        .str.replace(r"\((.+)\)", r"-\1", regex=True)
        .replace({"": "0", "nan": "0"})
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0.0)
    )


def fmt(v: float, sym: str = "$") -> str:
    return f"{sym}{v:,.2f}" if v >= 0 else f"-{sym}{abs(v):,.2f}"

# ---------------------------------------------------------------------------
# Build report
# ---------------------------------------------------------------------------
def build_report(
    df_raw:        pd.DataFrame,
    col:           dict,
    user_map:      dict | None = None,
    master_map:    dict | None = None,
    overrides_path: str | None = None,
    account_type_filter: str | None = None,
) -> str:
    df = df_raw.copy()
    date_col = df[col["date"]]
    if pd.api.types.is_numeric_dtype(date_col):
        df["_date"] = pd.to_datetime(
            pd.to_numeric(date_col, errors="coerce") - 25569,
            unit="D", origin="1970-01-01", errors="coerce"
        )
    else:
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            df["_date"] = pd.to_datetime(date_col, errors="coerce", dayfirst=False)
    df["_name"] = df[col["name"]].astype(str).str.strip()

    # ── Debit / Credit ─────────────────────────────────────────────────────
    # A transaction can arrive in one of two shapes:
    #   (a) separate debit & credit columns (both already non-negative), or
    #   (b) a single signed "amount" column, where a NEGATIVE value means the
    #       transaction was a payment/refund credited to the account (this is
    #       exactly what the xPence data-scrubbing tool emits for banks such
    #       as TD/CIBC when no separate credit column exists).
    # Feature: any negative amount that would otherwise be discarded is
    # instead routed into "_credit" so it shows up under Credits / Payments
    # rather than silently vanishing.
    raw_debit = clean_amount(df[col["debit"]]) if col.get("debit") else pd.Series(0.0, index=df.index)
    if col.get("credit"):
        df["_credit"] = clean_amount(df[col["credit"]]).clip(lower=0)
        df["_debit"]  = raw_debit.clip(lower=0)
        # Even when a dedicated credit column exists, a stray negative value
        # in the debit/amount column still represents money credited back —
        # fold it into credits instead of dropping it.
        df["_credit"] = df["_credit"] + (-raw_debit).clip(lower=0)
    else:
        df["_debit"]  = raw_debit.clip(lower=0)
        df["_credit"] = (-raw_debit).clip(lower=0)

    # ── Account type ───────────────────────────────────────────────────────
    # Identified from the column the data-scrubbing tool produced. If that
    # column isn't present in the source data at all, every transaction is
    # assumed to be on a Credit account (the common case for xPence users).
    account_type_detected = bool(col.get("account_type"))
    if account_type_detected:
        df["_account_type"] = (
            df[col["account_type"]].astype(str).str.strip()
            .replace({"": DEFAULT_ACCOUNT_TYPE, "nan": DEFAULT_ACCOUNT_TYPE, "None": DEFAULT_ACCOUNT_TYPE})
        )
    else:
        df["_account_type"] = DEFAULT_ACCOUNT_TYPE

    df = df.dropna(subset=["_date"]).reset_index(drop=True)

    # Optionally scope the whole report down to a single account type. Only
    # meaningful when the source data actually carried an account-type
    # column — filtering against an assumed/default value would be
    # misleading since it wasn't verified from the data itself.
    if account_type_filter and account_type_detected:
        df = df[df["_account_type"].str.casefold() == account_type_filter.casefold()].reset_index(drop=True)
    df["_month"]    = df["_date"].dt.to_period("M").astype(str)
    df["_date_str"] = df["_date"].dt.strftime("%b %d, %Y")
    df["_date_iso"] = df["_date"].dt.strftime("%Y-%m-%d")

    # Build valid_cats once; reuse across every classify() call
    _valid_cats = (
        set(CATEGORIES)
        | {"Other / Uncategorised"}
        | {cc["name"] for cc in (user_map or {}).get("custom_categories", [])}
    )

    df["_category"] = df["_name"].apply(
        lambda n: classify(n, user_map, master_map, _valid_cats)
    )

    # ── Questionable transactions ─────────────────────────────────────────
    # Vectorised: find credit-only rows, then check keyword match
    cred_mask = (df["_credit"] > 0) & (df["_debit"] == 0)
    questionable = []
    if cred_mask.any():
        for idx, row in df[cred_mask].iterrows():
            is_expense, matched_cat = matches_expense_keyword(row["_name"])
            if is_expense:
                questionable.append({
                    "qid":      idx,
                    "date":     row["_date_str"],
                    "date_iso": row["_date_iso"],
                    "name":     row["_name"],
                    "amount":   round(float(row["_credit"]), 2),
                    "category": matched_cat,
                    "month":    row["_month"],
                })
    if questionable:
        print(f"  Found {len(questionable)} questionable transaction(s) — see Review tab in report")
        for q in questionable:
            print(f"    {q['date']}  {q['name'][:50]:<50}  ${q['amount']:.2f}  [{q['category']}]")

    # ── Serialise rows — to_dict("records") is faster than iterrows() ────
    rows = [
        {
            "idx":          int(i),
            "date":         r["_date_str"],
            "date_iso":     r["_date_iso"],
            "name":         r["_name"],
            "account_type": r["_account_type"],
            "category":     r["_category"],
            "debit":        round(float(r["_debit"]),  2),
            "credit":       round(float(r["_credit"]), 2),
            "month":        r["_month"],
        }
        for i, r in enumerate(df.to_dict("records"))
    ]

    all_months     = sorted(df["_month"].unique())
    account_types  = sorted(df["_account_type"].unique().tolist())

    # ── Active categories — preserve defined order ────────────────────────
    _custom_cat_names = {
        cc["name"] for cc in (user_map or {}).get("custom_categories", [])
    }
    exp_mask   = df["_debit"] > 0
    cats_seen  = set(df.loc[exp_mask, "_category"].unique())
    _builtin_active = [c for c in CATEGORIES       if c in cats_seen and c not in _custom_cat_names]
    _custom_active  = [c for c in _custom_cat_names if c in cats_seen]
    _other_active   = ["Other / Uncategorised"] if "Other / Uncategorised" in cats_seen else []
    active_cats = _builtin_active + _custom_active + _other_active

    # O(1) colour lookup using precomputed index map
    def cat_color(cat):
        if cat in _custom_cat_names:   return CUSTOM_CAT_COLOR
        if cat == "Other / Uncategorised": return "#111111"
        idx = _CAT_INDEX.get(cat)
        if idx is not None:            return CAT_COLORS[idx % len(CAT_COLORS)]
        return CAT_COLORS[(len(CATEGORIES) + _builtin_active.index(cat)) % len(CAT_COLORS)]
    active_color_map = {cat: cat_color(cat) for cat in active_cats}

    # ── Pivot and totals — single groupby, then unstack ───────────────────
    exp_df       = df[exp_mask].copy()
    grp          = exp_df.groupby(["_month", "_category"])["_debit"].sum()
    cat_totals_all = {
        cat: round(float(grp.xs(cat, level="_category").sum()
                         if cat in grp.index.get_level_values("_category") else 0.0), 2)
        for cat in active_cats
    }
    pivot = {
        m: {cat: round(float(grp.get((m, cat), 0.0)), 2) for cat in active_cats}
        for m in all_months
    }

    total_spent = float(df["_debit"].sum())
    total_cred  = float(df["_credit"].sum())
    net         = total_spent - total_cred
    period_str  = f"{all_months[0]}  to  {all_months[-1]}" if all_months else "N/A"

    # Per-report merchant snapshot (flat dict, JS row data only)
    _row_merchant_map = (
        exp_df.groupby("_name")["_category"]
        .agg(lambda s: s.mode().iat[0])
        .to_dict()
    )

    # ── Overview statistics (pre-computed Python-side) ────────────────────
    exp = df[df["_debit"] > 0].copy()

    # 1. Top categories by total spend
    cat_rank = (
        exp.groupby("_category")["_debit"].sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    top_cats = [
        {"cat": r["_category"], "total": round(float(r["_debit"]), 2),
         "pct": round(float(r["_debit"] / exp["_debit"].sum() * 100), 1)}
        for _, r in cat_rank.head(5).iterrows()
    ]

    # 2. Busiest spending month
    month_totals = exp.groupby("_month")["_debit"].sum()
    peak_month   = month_totals.idxmax() if not month_totals.empty else ""
    peak_month_amt = round(float(month_totals.max()), 2) if not month_totals.empty else 0
    avg_month_amt  = round(float(month_totals.mean()), 2) if not month_totals.empty else 0

    # 3. Quietest month
    quiet_month     = month_totals.idxmin() if not month_totals.empty else ""
    quiet_month_amt = round(float(month_totals.min()), 2) if not month_totals.empty else 0

    # 4. Week-of-month breakdown (weeks 1-4) across all data
    exp = exp.copy()
    exp["_week_of_month"] = ((exp["_date"].dt.day - 1) // 7 + 1).clip(upper=4)
    week_totals = exp.groupby("_week_of_month")["_debit"].sum()
    busiest_week = int(week_totals.idxmax()) if not week_totals.empty else 1
    week_labels  = {1: "first", 2: "second", 3: "third", 4: "fourth/last"}

    # 5. Top merchants per category (top 3 per cat, top 5 cats)
    merch_by_cat = {}
    for cat in [r["_category"] for _, r in cat_rank.head(5).iterrows()]:
        top_m = (
            exp[exp["_category"] == cat]
            .groupby("_name")["_debit"]
            .agg(total="sum", visits="count")
            .sort_values("total", ascending=False)
            .reset_index()
            .head(3)
        )
        merch_by_cat[cat] = [
            {"name": r["_name"], "total": round(float(r["total"]), 2), "visits": int(r["visits"])}
            for _, r in top_m.iterrows()
        ]

    # 6. Unusual / outlier transactions (amount > mean + 2*std per category)
    unusual = []
    for cat, grp in exp.groupby("_category"):
        if len(grp) < 3:
            continue
        mu, sigma = grp["_debit"].mean(), grp["_debit"].std()
        if sigma == 0:
            continue
        for _, r in grp[grp["_debit"] > mu + 2 * sigma].iterrows():
            unusual.append({
                "name":   r["_name"],
                "cat":    cat,
                "amount": round(float(r["_debit"]), 2),
                "date":   r["_date_str"],
                "z":      round(float((r["_debit"] - mu) / sigma), 1),
            })
    unusual.sort(key=lambda x: x["z"], reverse=True)
    unusual = unusual[:6]

    # 7. Most frequent day-of-week for spending
    exp["_dow"] = exp["_date"].dt.day_name()
    dow_counts  = exp.groupby("_dow")["_debit"].sum()
    busiest_dow = dow_counts.idxmax() if not dow_counts.empty else ""

    # 8. Average transaction size
    avg_tx = round(float(exp["_debit"].mean()), 2) if not exp.empty else 0
    max_tx = exp.loc[exp["_debit"].idxmax()] if not exp.empty else None
    max_tx_info = {
        "name":   str(max_tx["_name"]),
        "amount": round(float(max_tx["_debit"]), 2),
        "date":   str(max_tx["_date_str"]),
        "cat":    str(max_tx["_category"]),
    } if max_tx is not None else {}

    # 9. Month-over-month trend (last 2 months vs prior 2 months)
    # Reindex against the full all_months list so sparse months (e.g. Wealthsimple
    # months where every transaction is a deposit, leaving no expense rows) get 0
    # instead of raising a KeyError when indexed via string label.
    mom_trend = None
    if len(all_months) >= 4:
        mt_full = month_totals.reindex(all_months, fill_value=0.0)
        recent  = mt_full.iloc[-2:].mean()
        prior   = mt_full.iloc[-4:-2].mean()
        if prior > 0:
            mom_pct = round(float((recent - prior) / prior * 100), 1)
            mom_trend = {"pct": mom_pct, "dir": "up" if mom_pct > 0 else "down"}

    # 10. Spending consistency — CV per category (lower = more consistent)
    consistency = {}
    for cat in [r["_category"] for _, r in cat_rank.head(5).iterrows()]:
        mvals = [v for v in (pivot[m].get(cat, 0) for m in all_months) if v > 0]
        if len(mvals) >= 2:
            mean_v = statistics.mean(mvals)
            std_v  = statistics.stdev(mvals)
            cv     = std_v / (mean_v or 1)
            consistency[cat] = {
                "cv":    round(cv, 2),
                "mean":  round(mean_v, 2),
                "std":   round(std_v, 2),
                "min":   round(min(mvals), 2),
                "max":   round(max(mvals), 2),
                "n":     len(mvals),
            }

    # 11. Short spend-habit tip — rule-based, data-driven, ≤2 sentences
    # Exclude "Other / Uncategorised" — those transactions are unknown/unclassified
    # and cannot reliably inform a spend-reduction recommendation.
    EXCLUDE_TIP = {"Other / Uncategorised"}
    tip_top_cats = [c for c in top_cats if c["cat"] not in EXCLUDE_TIP]
    tip_cat    = tip_top_cats[0]["cat"] if tip_top_cats else ""
    tip_pct    = tip_top_cats[0]["pct"] if tip_top_cats else 0
    n_months   = len(all_months) or 1

    spend_tip = ""
    if tip_cat:
        # ── Find best habitual merchant in the top category ──────────────────
        # A merchant is "habitual" only if they appear in ≥50% of months AND
        # have ≥3 total visits. This prevents one-off large purchases from
        # being cited as a regular habit worth cutting.
        cat_merchants = merch_by_cat.get(tip_cat, [])
        habitual_merch = None
        for m in cat_merchants:
            mname    = m["name"]
            # Count how many distinct months this merchant appears in
            months_present = exp[
                (exp["_category"] == tip_cat) & (exp["_name"] == mname)
            ]["_month"].nunique()
            if months_present >= max(2, n_months // 2) and m["visits"] >= 3:
                habitual_merch = {**m, "months_present": months_present}
                break  # already sorted by total desc — take the top qualifying one

        if habitual_merch:
            avg_visit_spend = round(habitual_merch["total"] / habitual_merch["visits"], 2)
            monthly_visits  = round(habitual_merch["visits"] / n_months, 1)
            months_pct      = round(habitual_merch["months_present"] / n_months * 100)
            spend_tip = (
                f"Your top category is <strong>{tip_cat}</strong> ({tip_pct}% of all spending). "
                f"<strong>{habitual_merch['name']}</strong> shows up in {months_pct}% of months "
                f"— about {monthly_visits}× per month at ~${avg_visit_spend:.0f} each visit. "
                f"Skipping one visit per month would free up ~${avg_visit_spend:.0f} consistently."
            )
        else:
            # Fallback: category-level tip, no specific merchant called out
            # (avoids citing a one-off big purchase as a savings lever)
            tip_cat_avg = round(exp[exp["_category"] == tip_cat]["_debit"].sum() / n_months, 2)
            spend_tip = (
                f"<strong>{tip_cat}</strong> is your largest spending category at {tip_pct}% of total spend "
                f"(~${tip_cat_avg:.0f}/month on average). "
                f"Setting a monthly budget for this category is the highest-impact change you can make."
            )
    else:
        spend_tip = "Keep tracking your spending consistently to unlock personalised saving tips."

    overview = {
        "top_cats":       top_cats,
        "peak_month":     peak_month,
        "peak_month_amt": peak_month_amt,
        "avg_month_amt":  avg_month_amt,
        "quiet_month":    quiet_month,
        "quiet_month_amt":quiet_month_amt,
        "busiest_week":   busiest_week,
        "busiest_dow":    busiest_dow,
        "merch_by_cat":   merch_by_cat,
        "unusual":        unusual,
        "avg_tx":         avg_tx,
        "max_tx":         max_tx_info,
        "mom_trend":      mom_trend,
        "consistency":    consistency,
        "spend_tip":      spend_tip,
        "n_months":       n_months,
    }

    data_json = json.dumps({
        "rows":              rows,
        "months":            all_months,
        "categories":        active_cats,
        "cat_colors":        [active_color_map[c] for c in active_cats],
        "cat_color_map":     active_color_map,
        "cat_totals_all":    cat_totals_all,
        "pivot":             pivot,
        "questionable":      questionable,
        "all_categories":    CATEGORIES + ["Other / Uncategorised"],
        # User overrides baked in so the HTML report can reconstruct
        # an updated user_overrides.json via 'Overrides → Download' — no --xpr needed.
        # Also used as fallback when user_overrides.json cannot be fetched (file:// protocol).
        "merchant_json": {
            "merchants":          _row_merchant_map,
            "overrides":          (user_map or {}).get("overrides", {}),
            "debit_credit_flips": (user_map or {}).get("debit_credit_flips", {}),
            "flagged":            (user_map or {}).get("flagged", {}),
            "custom_categories":  (user_map or {}).get("custom_categories", []),
            "subscriptions":      (user_map or {}).get("subscriptions", []),
        },
        # Convenience flat copies used by existing JS (cat colours, custom cats list)
        "custom_categories": (user_map or {}).get("custom_categories", []),
        "custom_cat_color":  CUSTOM_CAT_COLOR,
        "merchant_map":      _row_merchant_map,
        "overview":          overview,
        "summary": {
            "total_spent": round(total_spent, 2),
            "total_cred":  round(total_cred,  2),
            "net":         round(net, 2),
            "tx_count":    len(df),
            "period":      period_str,
            "generated":   datetime.now().strftime("%Y-%m-%d %H:%M"),
        },
        "overrides_path": overrides_path or "",
        # ── Account type metadata ──────────────────────────────────────────
        "account_types":          account_types,
        "account_type_detected":  account_type_detected,
        # Only baked in when the source data genuinely had an account-type
        # column — an assumed/default type is never persisted as if it were
        # a verified fact, since the same dates/merchants could belong to a
        # different account type in other data.
        "selected_account_type": (
            account_type_filter if (account_type_filter and account_type_detected) else None
        ),
    }, ensure_ascii=False)

    return HTML.replace("__DATA_JSON__", data_json)

# ---------------------------------------------------------------------------
# HTML template
# ---------------------------------------------------------------------------
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>xPence Report</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap');

:root {
  --ink:#0f172a; --ink2:#334155; --muted:#94a3b8;
  --surface:#fff; --border:#e2e8f0;
  --blue:#2563eb; --blue-lt:#eff6ff;
  --red:#dc2626;  --red-lt:#fef2f2;
  --green:#16a34a;--green-lt:#f0fdf4;
  --amber:#d97706;--amber-lt:#fffbeb;
  --radius:12px;  --radius-sm:7px;
  --shadow:0 1px 3px rgba(0,0,0,.08),0 4px 16px rgba(0,0,0,.04);
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'DM Sans',sans-serif;background:#f1f5f9;color:var(--ink);min-height:100vh}

/* header */
.site-header{background:var(--ink);padding:2rem 2.5rem 1.6rem;display:flex;flex-wrap:wrap;gap:1.2rem;align-items:center;justify-content:space-between}
.site-header h1{font-family:'DM Serif Display',serif;font-size:2rem;color:#fff;letter-spacing:-.02em;line-height:1.1}
.site-header h1 span{color:#93c5fd;font-style:italic}
.header-actions{display:flex;align-items:center;gap:.75rem;flex-shrink:0}
.hdr-drop{position:relative}
.hdr-btn{font-family:inherit;font-size:.92rem;font-weight:700;padding:.65rem 1.3rem;
  border-radius:var(--radius-sm);cursor:pointer;display:flex;align-items:center;gap:.5rem;
  transition:opacity .15s,background .15s;white-space:nowrap;border:none;letter-spacing:.01em}
.hdr-btn.report-btn{background:#93c5fd;color:#1e3a5f}
.hdr-btn.report-btn:hover{background:#bfdbfe}
.hdr-btn.report-btn.has-changes{background:#f59e0b;color:#1c1107;animation:pulse-chg 2s ease-in-out infinite}
.hdr-btn.report-btn.has-changes:hover{background:#fbbf24}
@keyframes pulse-chg{0%,100%{box-shadow:0 0 0 0 rgba(245,158,11,.55)}60%{box-shadow:0 0 0 6px rgba(245,158,11,0)}}
.hdr-btn.overrides-btn{background:#22c55e;color:#fff;position:relative}
.hdr-btn.overrides-btn:hover{background:#16a34a}
.hdr-btn.overrides-btn.has-changes{background:#f59e0b;color:#1c1107;animation:pulse-chg 2s ease-in-out infinite}
.hdr-btn.overrides-btn.has-changes:hover{background:#fbbf24}
.hdr-btn.overrides-btn.has-changes .upd-badge{display:none}
.hdr-drop-menu{display:none;position:absolute;top:calc(100% + 6px);right:0;background:var(--surface);
  border:1.5px solid var(--border);border-radius:var(--radius-sm);min-width:230px;
  box-shadow:0 4px 20px rgba(0,0,0,.12);z-index:200;overflow:hidden}
.hdr-drop-menu.open{display:block}
.hdr-drop-sep{height:1px;background:var(--border);margin:2px 0}
.hdr-drop-item{display:flex;align-items:center;gap:.6rem;padding:.7rem 1rem;
  font-size:.82rem;font-weight:500;color:var(--ink2);cursor:pointer;transition:background .12s;
  width:100%;box-sizing:border-box;background:none;border:none;font-family:inherit;text-align:left}
.hdr-drop-item:hover{background:var(--blue-lt);color:var(--blue)}
.hdr-drop-item .item-sub{font-size:.72rem;color:var(--muted);display:block;font-weight:400}
.hdr-drop-item.danger:hover{background:#fff1f2;color:var(--red)}
.meta{font-size:.8rem;color:#94a3b8}

/* summary strip */
.summary-strip{display:flex;flex-wrap:wrap;background:var(--ink);border-top:1px solid #1e293b;padding:0 2.5rem 1.6rem}
.stat{flex:1;min-width:140px;padding:.9rem 1.2rem}
.stat .lbl{font-size:.7rem;text-transform:uppercase;letter-spacing:.08em;color:#64748b;margin-bottom:.3rem}
.stat .val{font-size:1.35rem;font-weight:600;color:#fff}
.stat .val.red{color:#fca5a5}.stat .val.grn{color:#86efac}

/* tab bar */
.tab-bar{display:flex;justify-content:center;gap:.5rem;background:var(--surface);border-bottom:2px solid var(--border);padding:0 2.5rem;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.06)}
.tab-btn{background:none;border:none;cursor:pointer;font-family:inherit;font-size:.9rem;font-weight:500;color:var(--muted);padding:1rem 1.4rem;border-bottom:3px solid transparent;margin-bottom:-2px;transition:color .2s,border-color .2s;display:flex;align-items:center;gap:.45rem}
.tab-btn.active{color:var(--blue);border-bottom-color:var(--blue)}
.tab-btn:hover:not(.active){color:var(--ink2)}
.tab-btn .badge{font-size:.75rem;font-weight:900;background:#f59e0b;color:#fff;border-radius:999px;width:1.1rem;height:1.1rem;display:inline-flex;align-items:center;justify-content:center;margin-left:.2rem;line-height:1}

/* main */
.main{padding:2rem 2.5rem;max-width:1400px;margin:0 auto}
.view{display:none}.view.active{display:block}

/* section cards */
.section-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);margin-bottom:1.8rem;overflow:hidden}
.section-head{padding:1.1rem 1.5rem;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:1rem;flex-wrap:wrap}
.section-head h2{font-family:'DM Serif Display',serif;font-size:1.4rem;color:var(--ink);font-weight:400}
.section-body{padding:1.5rem}

/* ── Review tab styles ── */
.review-intro{background:var(--amber-lt);border:1.5px solid #fcd34d;border-radius:var(--radius-sm);padding:1rem 1.2rem;margin-bottom:1.5rem;font-size:.88rem;color:#92400e;line-height:1.6}
.review-intro strong{color:#78350f}
.q-month-group{margin-bottom:1.8rem}
.q-month-header{display:flex;align-items:baseline;justify-content:space-between;gap:.8rem;margin-bottom:.75rem;padding-bottom:.5rem;border-bottom:2px solid var(--border)}
.q-month-label{font-family:'DM Serif Display',serif;font-size:1rem;color:var(--ink);font-weight:400}
.q-month-meta{font-size:.78rem;color:var(--muted)}
.review-empty{text-align:center;padding:3rem;color:var(--muted);font-size:.9rem}
/* ── Unified review cards (uncat, reclassify, questionable) ── */
.rev-card{border:1.5px solid var(--border);border-radius:var(--radius-sm);padding:1rem 1.2rem;margin-bottom:.65rem;background:var(--surface);transition:border-color .2s}
.rev-card.is-debit{border-color:#bfdbfe;background:#f8faff}
.rev-card.is-credit{border-color:#bbf7d0;background:#f0fdf8}
.rev-card.is-uncat{border-color:#fed7aa;background:#fff7ed}
.rev-card.is-flag{border-color:#e2e8f0;background:#f8fafc}
/* Row 1: meta line */
.rev-meta{display:flex;align-items:center;gap:.75rem;margin-bottom:.65rem}
.rev-date{font-size:.72rem;color:var(--muted);white-space:nowrap;flex-shrink:0;min-width:82px}
.rev-name{font-size:.86rem;font-weight:600;color:var(--ink2);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.rev-cat-badge{font-size:.7rem;background:var(--blue-lt);color:var(--blue);border-radius:4px;padding:.2rem .55rem;white-space:nowrap;flex-shrink:0}
.rev-amt{font-size:.95rem;font-weight:700;white-space:nowrap;flex-shrink:0;min-width:72px;text-align:right}
.rev-amt.spend{color:var(--red)}
.rev-amt.credit{color:var(--green)}
/* Row 2: controls line — consistent height and gaps across all three sections */
.rev-controls{display:flex;align-items:center;gap:.6rem;flex-wrap:nowrap}
.rev-controls .cat-select{flex:1;min-width:0;max-width:100%}
.rev-controls .apply-btn{white-space:nowrap;flex-shrink:0;min-width:88px;text-align:center}
.rev-controls .toggle-wrap{flex-shrink:0}
.rev-controls .remove-btn{flex-shrink:0;background:none;border:1.5px solid var(--border);border-radius:var(--radius-sm);font-size:.72rem;color:var(--muted);padding:.38rem .65rem;cursor:pointer;transition:border-color .15s,color .15s;line-height:1}
.rev-controls .remove-btn:hover{border-color:var(--red);color:var(--red)}
.rev-applied{margin-top:.45rem;font-size:.72rem;font-weight:600;color:#16a34a}
/* Keep legacy selectors pointing to new classes */
.uncat-card{display:contents}
.q-card{display:contents}
/* Shared dropdown — consistent height across all sections */
.cat-select{font-family:inherit;font-size:.82rem;font-weight:500;color:var(--ink);background:var(--surface);border:1.5px solid var(--border);border-radius:var(--radius-sm);padding:.38rem 2rem .38rem .75rem;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='11' height='11' viewBox='0 0 24 24' fill='none' stroke='%2364748b' stroke-width='2.5'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right .6rem center;cursor:pointer;width:100%;height:2.1rem;box-sizing:border-box}
.cat-select:focus{outline:none;border-color:var(--blue)}
/* Shared action buttons — same height as dropdown for optical alignment */
.apply-btn{font-family:inherit;font-size:.8rem;font-weight:600;padding:0 1rem;border-radius:var(--radius-sm);border:none;cursor:pointer;transition:opacity .15s;white-space:nowrap;height:2.1rem;display:inline-flex;align-items:center;justify-content:center}
.apply-btn.apply-cat{background:var(--blue);color:#fff}
.apply-btn.apply-cat:hover{opacity:.85}
.apply-btn.apply-q{background:var(--ink);color:#fff;margin-top:1.2rem;height:auto;padding:.5rem 1.4rem;font-size:.85rem}
.apply-btn.apply-q:hover{opacity:.8}
.applied-tag{font-size:.72rem;font-weight:600;background:#dcfce7;color:#16a34a;border-radius:4px;padding:.2rem .55rem;white-space:nowrap}
/* Toggle — same height as dropdown and buttons */
.toggle-wrap{display:flex;border:1.5px solid var(--border);border-radius:6px;overflow:hidden;flex-shrink:0;height:2.1rem}
.toggle-btn{background:none;border:none;cursor:pointer;font-family:inherit;font-size:.78rem;font-weight:600;padding:0 .9rem;transition:background .15s,color .15s;color:var(--muted)}
.toggle-btn.active-debit{background:var(--red-lt);color:var(--red)}
.toggle-btn.active-credit{background:var(--green-lt);color:var(--green)}
.toggle-sep{width:1px;background:var(--border)}
.section-head.collapsible{cursor:pointer;user-select:none}
.section-head.collapsible:hover{background:#f8fafc}
.collapse-icon{display:inline-flex;align-items:center;margin-left:.5rem;color:var(--muted);transition:transform .2s;flex-shrink:0}
.section-card.collapsed .collapse-icon{transform:rotate(-90deg)}
.section-card.collapsed .section-body{display:none}

/* custom over-time chart legend buttons */
.chart-legend-btn{display:inline-flex;align-items:center;gap:.42rem;font-family:inherit;font-size:.78rem;font-weight:500;padding:.28rem .72rem;border-radius:999px;border:2px solid var(--leg-color,#e2e8f0);color:var(--leg-color,#94a3b8);background:color-mix(in srgb,var(--leg-color,transparent) 8%,transparent);cursor:pointer;transition:opacity .15s;white-space:nowrap}
.chart-legend-btn .leg-dot{width:9px;height:9px;border-radius:50%;flex-shrink:0;background:var(--leg-color,#cbd5e1)}
.chart-legend-btn .leg-dot-dashed{width:18px;height:3px;border-radius:0;background:none;border-top:2.5px dashed var(--leg-color,#000);flex-shrink:0}
.chart-legend-btn .leg-label{}
.chart-legend-btn.leg-off{--leg-color:#e2e8f0;color:#94a3b8;background:#f8fafc}
.chart-legend-btn.leg-off .leg-dot{background:#cbd5e1}
.chart-legend-btn.leg-off .leg-dot-dashed{border-color:#cbd5e1}
.chart-legend-btn.leg-off .leg-label{text-decoration:line-through}
.chart-legend-total{border-style:dashed}
.month-picker-row{display:flex;align-items:center;gap:1rem;flex-wrap:wrap;margin-bottom:1.8rem}
.month-picker-row label{font-size:.85rem;font-weight:500;color:var(--ink2)}
select{font-family:inherit;font-size:.9rem;font-weight:500;color:var(--ink);background:var(--surface);border:1.5px solid var(--border);border-radius:var(--radius-sm);padding:.5rem 2rem .5rem .9rem;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2364748b' stroke-width='2.5'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right .7rem center;cursor:pointer}
select:focus{outline:none;border-color:var(--blue)}
.month-stat-chips{display:flex;gap:.6rem;flex-wrap:wrap}
.chip{font-size:.78rem;font-weight:500;padding:.3rem .75rem;border-radius:999px;border:1.5px solid}
.chip.spend{color:var(--red);border-color:#fecaca;background:var(--red-lt)}
.chip.credit{color:var(--green);border-color:#bbf7d0;background:var(--green-lt)}
.chip.tx{color:var(--blue);border-color:#bfdbfe;background:var(--blue-lt)}

/* bar chart */
.cat-bar-list{display:flex;flex-direction:column;gap:.8rem}
.cat-bar-item{display:flex;align-items:center;gap:.8rem;font-size:.93rem;cursor:pointer;border-radius:var(--radius-sm);padding:.35rem .4rem;margin:0 -.4rem;transition:background .15s}
.cat-bar-item:hover{background:#f1f5f9}
.cat-bar-item.active-filter{background:var(--blue-lt);outline:1.5px solid var(--blue)}
.cat-bar-item.active-filter .cat-name{color:var(--blue);font-weight:600}
.cat-bar-item .cat-name{flex-shrink:0;font-weight:500;color:var(--ink2);white-space:nowrap}
.cat-bar-item .bar-track{flex:1;min-width:40px;background:#f1f5f9;border-radius:4px;height:12px;overflow:hidden}
.cat-bar-item .bar-fill{height:12px;border-radius:4px;transition:width .5s cubic-bezier(.4,0,.2,1)}
.cat-bar-item .bar-amt{min-width:90px;text-align:right;font-weight:600;color:var(--ink)}
.cat-bar-item .bar-pct{min-width:46px;text-align:right;color:var(--muted);font-size:.83rem}

/* pie */
.pie-wrap{display:flex;gap:2rem;align-items:flex-start;flex-wrap:wrap}
.pie-canvas-wrap{flex-shrink:0}
.pie-canvas-wrap canvas{display:block;max-width:300px;max-height:300px}
.pie-legend{flex:1;min-width:200px}
.legend-item{display:flex;align-items:center;gap:.7rem;font-size:.93rem;padding:.35rem 0;border-bottom:1px solid var(--border)}
.legend-item:last-child{border-bottom:none}
.legend-dot{width:13px;height:13px;border-radius:50%;flex-shrink:0}
.legend-name{flex:1;color:var(--ink2)}
.legend-amt{font-weight:600;color:var(--ink)}
.legend-pct{color:var(--muted);font-size:.84rem;min-width:42px;text-align:right}
.legend-item{cursor:pointer;border-radius:var(--radius-sm);margin:0 -.4rem;padding:.35rem .4rem;transition:background .15s}
.legend-item:hover{background:#f1f5f9}
.legend-item.active-filter{background:var(--blue-lt);outline:1.5px solid var(--blue)}
.legend-item.active-filter .legend-name{color:var(--blue);font-weight:600}
.clear-filter-btn{display:none;align-items:center;gap:.35rem;background:none;border:1.5px solid var(--border);
  border-radius:var(--radius-sm);padding:.35rem .75rem;font-family:inherit;font-size:.78rem;
  font-weight:500;color:var(--ink2);cursor:pointer;white-space:nowrap;transition:border-color .15s,color .15s}
.clear-filter-btn:hover{border-color:var(--blue);color:var(--blue)}
.clear-filter-btn.visible{display:flex}

/* pivot */
.pivot-wrap{overflow:auto;position:relative;max-height:calc(45px + 10 * 38px)}
/* border-separate (not collapse) is required for position:sticky on td/th */
.pivot-wrap table{width:auto;min-width:100%;border-collapse:separate;border-spacing:0;font-size:.82rem;table-layout:auto}
/* Frozen header row */
.pivot-wrap thead th{background:var(--ink);color:#fff;padding:.6rem .4rem;text-align:left;font-weight:500;
  position:sticky;top:0;z-index:3;
  overflow:visible;vertical-align:middle;font-size:1rem;line-height:1;
  border-bottom:2px solid #334155}
.pivot-wrap thead th.amt{text-align:right}
/* Frozen col 1 (Month) — header cell */
.pivot-wrap thead th:nth-child(1){left:0;z-index:4}
/* Frozen col 2 (Total) — header cell */
.pivot-wrap thead th:nth-child(2){left:90px;z-index:4}
/* Body cells */
.pivot-wrap tbody td{padding:.55rem .4rem;border-bottom:1px solid var(--border);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;background:var(--surface)}
.pivot-wrap tbody tr:last-child td{border-bottom:none}
.pivot-wrap tbody tr:hover td{background:var(--blue-lt)}
.pivot-wrap td.amt{text-align:right;font-variant-numeric:tabular-nums}
.pivot-wrap td.spend{color:var(--red);font-weight:600}
.pivot-wrap td.credit{color:var(--green);font-weight:600}
.pivot-wrap .total-row td{background:#f8fafc;font-weight:700;border-top:2px solid var(--border)}
#pivot-table tbody tr:not(.total-row){cursor:pointer}
/* Frozen col 1 (Month) — body cells */
.pivot-wrap tbody td:nth-child(1){position:sticky;left:0;z-index:1;border-right:1px solid var(--border)}
/* Frozen col 2 (Total) — body cells */
.pivot-wrap tbody td:nth-child(2){position:sticky;left:90px;z-index:1;border-right:1px solid #cbd5e1}
/* Total footer frozen cols */
.pivot-wrap .total-row td:nth-child(1),.pivot-wrap .total-row td:nth-child(2){background:#f8fafc}
/* Tooltip for category emoji headers — cursor-following via JS */
.pivot-wrap thead th[data-tip]{position:sticky;top:0}
#xp-cursor-tip{
  position:fixed;pointer-events:none;z-index:99999;
  background:#1e293b;color:#fff;font-size:.72rem;font-weight:400;white-space:nowrap;
  padding:.3rem .6rem;border-radius:5px;
  box-shadow:0 2px 8px rgba(0,0,0,.25);
  opacity:0;transition:opacity .08s;
}
/* general table baseline (tx-tables, etc.) */
table{width:100%;border-collapse:collapse;font-size:.82rem}
thead th{background:var(--ink);color:#fff;padding:.6rem .85rem;text-align:left;font-weight:500;white-space:nowrap}
thead th.amt{text-align:right}
tbody td{padding:.55rem .85rem;border-bottom:1px solid var(--border);white-space:nowrap}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover td{background:var(--blue-lt)}
td.amt{text-align:right;font-variant-numeric:tabular-nums}
td.spend{color:var(--red);font-weight:600}
td.credit{color:var(--green);font-weight:600}
.total-row td{background:#f8fafc;font-weight:700;border-top:2px solid var(--border)}

/* tx table */
.tx-table thead th{background:#1e3a5f}
.tx-table thead th.sortable{cursor:pointer;user-select:none;white-space:nowrap}
.tx-table thead th.sortable:hover{background:#2d5285}
.sort-icon{display:inline-block;margin-left:.35rem;opacity:.45;font-size:.75em;vertical-align:middle}
.tx-table thead th.sort-asc .sort-icon,.tx-table thead th.sort-desc .sort-icon{opacity:1;color:#93c5fd}
.tx-count-badge{font-size:.75rem;font-weight:500;padding:.2rem .6rem;border-radius:999px;background:var(--blue-lt);color:var(--blue)}
.cat-tag{display:inline-block;font-size:.7rem;font-weight:500;padding:.15rem .5rem;border-radius:4px;white-space:nowrap;color:#fff}
.q-flag{display:inline-block;font-size:.65rem;font-weight:700;padding:.1rem .4rem;border-radius:4px;background:#fef3c7;color:#b45309;margin-left:.3rem;white-space:nowrap}
.flag-btn{background:none;border:1px solid transparent;border-radius:4px;
  color:transparent;font-size:.75rem;font-weight:600;padding:.15rem .35rem;
  cursor:pointer;font-family:inherit;transition:border-color .15s,color .15s,background .15s;
  white-space:nowrap;display:block;margin:0 auto}
tr:hover .flag-btn{color:var(--muted);border-color:var(--border)}
.flag-btn:hover{border-color:var(--amber) !important;color:var(--amber) !important}
.flag-btn.flagged{background:rgba(245,158,11,.12);border-color:rgba(245,158,11,.3) !important;color:var(--amber) !important}
.search-box{font-family:inherit;font-size:.85rem;border:1.5px solid var(--border);border-radius:var(--radius-sm);padding:.45rem .85rem;min-width:220px}
.search-box:focus{outline:none;border-color:var(--blue)}

/* ── Overview tab ── */
.ov-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(310px,1fr));gap:1.4rem;margin-bottom:1.8rem}
.ov-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);padding:1.4rem 1.6rem;position:relative;overflow:hidden}
.ov-card-title{font-size:.95rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-bottom:.9rem}
.ov-card-title svg{display:none}
.ov-headline{font-family:'DM Serif Display',serif;font-size:1.5rem;color:var(--ink);line-height:1.2;margin-bottom:.35rem}
.ov-sub{font-size:.84rem;color:var(--ink2);line-height:1.55}
.ov-sub strong{color:var(--ink);font-weight:600}
/* hero layout inside ov-card — mirrors sub-hero style */
.ov-hero-row{display:flex;gap:1.2rem;flex-wrap:wrap}
.ov-hero-cell{flex:1;display:flex;flex-direction:column;gap:.18rem}
.ov-hero-label{font-size:.9rem;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:var(--muted);white-space:nowrap}
.ov-hero-value{font-size:1.9rem;font-weight:800;color:var(--ink);font-family:'DM Serif Display',serif;line-height:1.1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ov-hero-sub{font-size:.88rem;color:var(--muted);margin-top:.05rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ov-cat-row{display:flex;align-items:center;gap:.7rem;padding:.38rem 0;border-bottom:1px solid var(--border)}
.ov-cat-row:last-child{border-bottom:none}
.ov-cat-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
.ov-cat-name{flex:1;font-size:.85rem;font-weight:500;color:var(--ink2)}
.ov-cat-bar-wrap{width:90px;background:#f1f5f9;border-radius:4px;height:7px;overflow:hidden;flex-shrink:0}
.ov-cat-bar{height:7px;border-radius:4px}
.ov-cat-pct{font-size:.75rem;font-weight:600;color:var(--ink);min-width:34px;text-align:right}
.ov-merch-row{display:flex;align-items:baseline;justify-content:space-between;gap:.5rem;padding:.32rem 0;border-bottom:1px solid var(--border);font-size:.83rem}
.ov-merch-row:last-child{border-bottom:none}
.ov-merch-name{color:var(--ink2);font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1}
.ov-merch-meta{color:var(--muted);font-size:.75rem;white-space:nowrap;flex-shrink:0}
.ov-merch-amt{font-weight:600;color:var(--ink);white-space:nowrap;flex-shrink:0}
.ov-unusual-row{padding:.45rem 0;border-bottom:1px solid var(--border);font-size:.83rem}
.ov-unusual-row:last-child{border-bottom:none}
.ov-unusual-name{font-weight:500;color:var(--ink2)}
.ov-unusual-meta{font-size:.75rem;color:var(--muted);margin-top:.1rem}
.ov-badge{display:inline-block;font-size:.67rem;font-weight:700;padding:.1rem .45rem;border-radius:4px;background:var(--red-lt);color:var(--red);margin-left:.3rem;vertical-align:middle}
.ov-trend-up{color:var(--red)}
.ov-trend-dn{color:var(--green)}
.ov-tip-card{background:linear-gradient(135deg,#1e3a5f 0%,#1e293b 100%);border-radius:var(--radius);padding:1.8rem 2rem;color:#e2e8f0;margin-bottom:1.8rem;position:relative;overflow:hidden}
.ov-tip-card::after{content:"💡";position:absolute;right:1.5rem;top:50%;transform:translateY(-50%);font-size:3rem;opacity:.15}
.ov-tip-label{font-size:.68rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:#64748b;margin-bottom:.6rem}
.ov-tip-text{font-size:.97rem;line-height:1.65;color:#cbd5e1}
.ov-tip-text strong{color:#93c5fd}
.ov-week-row{display:flex;align-items:center;gap:.7rem;padding:.3rem 0}
.ov-week-label{font-size:.82rem;color:var(--ink2);min-width:70px}
.ov-week-bar-wrap{flex:1;background:#f1f5f9;border-radius:4px;height:8px;overflow:hidden}
.ov-week-bar{height:8px;border-radius:4px;background:var(--blue)}
.ov-week-amt{font-size:.8rem;font-weight:600;color:var(--ink);min-width:72px;text-align:right}
.ov-consist-row{display:flex;align-items:center;gap:.6rem;padding:.32rem 0;border-bottom:1px solid var(--border);font-size:.83rem}
.ov-consist-row:last-child{border-bottom:none}
.ov-consist-name{flex:1;color:var(--ink2)}
.ov-consist-label{font-size:.72rem;font-weight:600;padding:.1rem .45rem;border-radius:4px}
.consist-low{background:#f0fdf4;color:#16a34a}
.consist-med{background:#fffbeb;color:#d97706}
.consist-high{background:#fef2f2;color:#dc2626}

/* ── Customize panel ── */
.cust-section{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);margin-bottom:1.8rem;overflow:hidden}
.cust-head{padding:1.1rem 1.5rem;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:.8rem;flex-wrap:wrap}
.cust-head h2{font-family:'DM Serif Display',serif;font-size:1.4rem;color:var(--ink);font-weight:400;flex:1}
.cust-body{padding:1.5rem}
.cust-intro{font-size:.86rem;color:var(--ink2);line-height:1.6;margin-bottom:1.4rem}
.cust-intro strong{color:var(--ink)}
.cust-cat-list{display:flex;flex-direction:column;gap:.75rem;margin-bottom:1.6rem}
.cust-cat-item{border:1.5px solid #cbd5e1;border-left:4px solid #3d3d3d;border-radius:var(--radius-sm);padding:.85rem 1rem;background:#f8fafc;display:flex;align-items:flex-start;gap:.75rem}
.cust-cat-dot{width:11px;height:11px;border-radius:50%;background:#3d3d3d;flex-shrink:0;margin-top:.3rem}
.cust-cat-info{flex:1;min-width:0}
.cust-cat-name{font-size:.9rem;font-weight:600;color:var(--ink);margin-bottom:.4rem}
.cust-kw-chips{display:flex;flex-wrap:wrap;gap:.35rem}
.cust-kw-chip{font-size:.72rem;font-weight:500;background:#e2e8f0;color:#475569;border-radius:4px;padding:.18rem .52rem}
.cust-cat-del{background:none;border:1.5px solid var(--border);border-radius:var(--radius-sm);padding:.32rem .65rem;font-size:.75rem;color:var(--muted);cursor:pointer;flex-shrink:0;transition:border-color .15s,color .15s;line-height:1;margin-left:auto}
.cust-cat-del:hover{border-color:var(--red);color:var(--red)}
.cust-add-form{border:1.5px solid var(--border);border-radius:var(--radius-sm);padding:1.2rem 1.3rem;background:#fafbfc;margin-top:.2rem}
.cust-add-title{font-size:.73rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-bottom:1rem}
.cust-field{margin-bottom:.85rem}
.cust-field label{display:block;font-size:.8rem;font-weight:600;color:var(--ink2);margin-bottom:.35rem}
.cust-field input{font-family:inherit;font-size:.88rem;width:100%;border:1.5px solid var(--border);border-radius:var(--radius-sm);padding:.5rem .85rem;background:var(--surface);color:var(--ink);box-sizing:border-box}
.cust-field input:focus{outline:none;border-color:var(--blue)}
.cust-field .field-hint{font-size:.73rem;color:var(--muted);margin-top:.3rem}
.cust-add-btn{font-family:inherit;font-size:.85rem;font-weight:600;padding:.52rem 1.3rem;border-radius:var(--radius-sm);border:none;cursor:pointer;background:var(--blue);color:#fff;transition:opacity .15s}
.cust-add-btn:hover{opacity:.85}
.cust-save-bar{margin-top:1.5rem;padding-top:1.25rem;border-top:1.5px solid var(--border);display:flex;align-items:center;gap:1rem;flex-wrap:wrap}
.cust-save-note{font-size:.8rem;color:var(--muted);flex:1;line-height:1.55}
.cust-save-btn{font-family:inherit;font-size:.85rem;font-weight:600;padding:.52rem 1.4rem;border-radius:var(--radius-sm);border:none;cursor:pointer;background:#1e3a5f;color:#fff;transition:opacity .15s;display:flex;align-items:center;gap:.45rem;white-space:nowrap}
.cust-save-btn:hover{opacity:.85}
.cust-empty{font-size:.85rem;color:var(--muted);font-style:italic;padding:.35rem 0}
.cust-feedback{font-size:.82rem;font-weight:600;color:#16a34a;margin-left:.5rem;transition:opacity .4s}
/* ── Category color picker rows ── */
.cust-color-list{display:flex;flex-direction:column;gap:.5rem}
.cust-color-row{display:flex;align-items:center;gap:.9rem;padding:.6rem .8rem;border-radius:var(--radius-sm);border:1px solid var(--border);background:#fafbfc;transition:background .15s}
.cust-color-row:hover{background:var(--blue-lt)}
.cust-color-swatch{width:22px;height:22px;border-radius:50%;flex-shrink:0;border:2px solid rgba(0,0,0,.08)}
.cust-color-name{flex:1;font-size:.87rem;font-weight:500;color:var(--ink2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cust-color-picker{width:34px;height:28px;border:1.5px solid var(--border);border-radius:6px;padding:2px;cursor:pointer;background:none;flex-shrink:0}
.cust-color-picker:hover{border-color:var(--blue)}
.cust-color-reset{font-family:inherit;font-size:.73rem;font-weight:600;padding:.28rem .65rem;border-radius:var(--radius-sm);border:1.5px solid var(--border);background:none;color:var(--muted);cursor:pointer;white-space:nowrap;flex-shrink:0;transition:border-color .15s,color .15s}
.cust-color-reset:hover{border-color:var(--blue);color:var(--blue)}
/* color compact row */
.cust-color-compact{display:flex;align-items:center;gap:.9rem;padding:1rem 0;flex-wrap:wrap}
.cust-color-compact select{font-family:inherit;font-size:.87rem;padding:.42rem .75rem;
  border:1.5px solid var(--border);border-radius:var(--radius-sm);background:var(--surface);color:var(--ink);flex:1;min-width:160px}
.cust-color-compact select:focus{outline:none;border-color:var(--blue)}
/* override cards */
.ovr-list{display:flex;flex-direction:column;gap:.6rem;margin-bottom:1rem}
.ovr-card{display:flex;align-items:center;gap:.8rem;padding:.7rem 1rem;
  border:1.5px solid var(--border);border-radius:var(--radius-sm);background:#fafbfc}
.ovr-card-merchant{flex:1;font-size:.87rem;font-weight:600;color:var(--ink);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ovr-card-arrow{font-size:.8rem;color:var(--muted);flex-shrink:0}
.ovr-card-cat{font-size:.8rem;font-weight:600;padding:.18rem .55rem;border-radius:4px;
  background:var(--blue-lt);color:var(--blue);flex-shrink:0;white-space:nowrap}
.ovr-card-source{font-size:.7rem;color:var(--muted);flex-shrink:0;white-space:nowrap}
.ovr-card-del{background:none;border:1.5px solid var(--border);border-radius:var(--radius-sm);
  padding:.25rem .55rem;font-size:.75rem;color:var(--muted);cursor:pointer;flex-shrink:0;
  transition:border-color .15s,color .15s;line-height:1}
.ovr-card-del:hover{border-color:var(--red);color:var(--red)}
.ovr-empty{font-size:.85rem;color:var(--muted);font-style:italic;padding:.5rem 0}

/* ── Subscriptions tab ── */
.sub-hero{display:flex;gap:1.4rem;flex-wrap:wrap;margin-bottom:1.8rem}
.sub-hero-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);padding:1.6rem 2rem;flex:1;min-width:200px;display:flex;flex-direction:column;gap:.3rem}
.sub-hero-label{font-size:.95rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:var(--muted)}
.sub-hero-value{font-size:1.9rem;font-weight:800;color:var(--ink);font-family:'DM Serif Display',serif;line-height:1.1}
.sub-hero-sub{font-size:.95rem;color:var(--muted);margin-top:.1rem}
.sub-hero-pct{color:var(--blue)}
.sub-note{font-size:.8rem;color:var(--muted);background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-sm);padding:.65rem 1rem;margin-bottom:1.4rem;line-height:1.6}
/* list table */
.sub-list-wrap{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;margin-bottom:1.8rem}
.sub-list-head{display:grid;grid-template-columns:1fr 100px 110px 110px 110px 100px 130px;gap:0;padding:.6rem 1.2rem;background:#f8fafc;border-bottom:1px solid var(--border);font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--muted)}
.sub-list-row{display:grid;grid-template-columns:1fr 100px 110px 110px 110px 100px 130px;gap:0;padding:.75rem 1.2rem;border-bottom:1px solid var(--border);align-items:start;transition:background .12s}
.sub-list-row:last-child{border-bottom:none}
.sub-list-row:hover{background:#f8fafc}
.sub-row-name{font-size:.88rem;font-weight:600;color:var(--ink);display:flex;align-items:flex-start;gap:.5rem;min-width:0}
.sub-row-name-info{min-width:0;flex:1}
.sub-row-name-text{white-space:normal;overflow-wrap:break-word;word-break:break-word;line-height:1.35}
.sub-row-cat{font-size:.75rem;color:var(--muted);white-space:normal;overflow-wrap:break-word;display:flex;align-items:center;gap:.3rem;margin-top:.15rem}
.sub-row-amt{font-size:.9rem;font-weight:700;color:var(--ink);text-align:right;padding-right:1rem}
.sub-col-amt-head{text-align:right;padding-right:1rem}
.sub-row-meta{font-size:.78rem;color:var(--ink2)}
.sub-row-muted{font-size:.78rem;color:var(--muted)}
.sub-pinned-dot{width:8px;height:8px;border-radius:50%;background:#16a34a;flex-shrink:0;display:inline-block}
.sub-manual-dot{width:8px;height:8px;border-radius:50%;background:#2563eb;flex-shrink:0;display:inline-block}
.sub-row-actions{display:flex;gap:.4rem;justify-content:flex-end}
.sub-btn{font-family:inherit;font-size:.73rem;font-weight:600;padding:.22rem .65rem;border-radius:4px;border:1.5px solid;cursor:pointer;background:none;line-height:1.4;white-space:nowrap;transition:all .15s}
.sub-btn-confirm{color:#16a34a;border-color:#16a34a}
.sub-btn-confirm:hover{background:#f0fdf4}
.sub-btn-delete{color:var(--red);border-color:var(--red)}
.sub-btn-delete:hover{background:#fef2f2}
/* manual add */
.sub-empty{text-align:center;padding:2.5rem 1rem;color:var(--muted)}
.sub-empty-icon{font-size:2rem;margin-bottom:.5rem}
.sub-empty-msg{font-size:.9rem;font-weight:500;color:var(--ink2);margin-bottom:.3rem}
.sub-empty-hint{font-size:.8rem;color:var(--muted)}
.sub-manual-section{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:1.3rem 1.4rem}
.sub-manual-title{font-size:.73rem;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:var(--muted);margin-bottom:.9rem;display:flex;align-items:center;gap:.4rem}
.sub-manual-row{display:flex;gap:.7rem;flex-wrap:wrap;align-items:flex-end}
.sub-manual-field{display:flex;flex-direction:column;gap:.28rem;flex:1;min-width:130px}
.sub-manual-field label{font-size:.78rem;font-weight:600;color:var(--ink2)}
.sub-manual-field input,.sub-manual-field select{font-family:inherit;font-size:.85rem;border:1.5px solid var(--border);border-radius:var(--radius-sm);padding:.42rem .72rem;background:var(--surface);color:var(--ink)}
.sub-manual-field input:focus,.sub-manual-field select:focus{outline:none;border-color:var(--blue)}
.sub-manual-add-btn{font-family:inherit;font-size:.85rem;font-weight:600;padding:.46rem 1.2rem;border-radius:var(--radius-sm);border:none;cursor:pointer;background:var(--blue);color:#fff;white-space:nowrap;align-self:flex-end}
@media(max-width:860px){
  .sub-list-head,.sub-list-row{grid-template-columns:1fr 90px 90px 120px}
  .sub-list-head .sub-col-day,.sub-list-row .sub-col-day,
  .sub-list-head .sub-col-first,.sub-list-row .sub-col-first,
  .sub-list-head .sub-col-total,.sub-list-row .sub-col-total{display:none}
}

@media(max-width:720px){
  .site-header,.summary-strip,.main{padding-left:1rem;padding-right:1rem}
  .tab-bar{padding:0 1rem}
}
</style>
</head>
<body>

<header class="site-header">
  <div>
    <h1>x<span>Pence</span></h1>
    <div class="meta" id="meta-period"></div>
    <span id="overrides-source-note" style="display:none;font-size:11px;color:#f59e0b;margin-left:8px;font-weight:500;"></span>
  </div>

  <div class="header-actions">
    <!-- ── Report dropdown ── -->
    <div class="hdr-drop" id="report-drop">
      <button class="hdr-btn report-btn" onclick="toggleDrop('report-drop')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
        Report
        <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
      </button>
      <div class="hdr-drop-menu" id="report-drop-menu">
        <button class="hdr-drop-item" onclick="saveReport();closeDrop('report-drop')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
          <span>Download<span class="item-sub">Save a .xpr snapshot with all current decisions</span></span>
        </button>
        <div class="hdr-drop-sep"></div>
        <label class="hdr-drop-item" for="upload-fresh">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 9h6M9 12h6M9 15h4"/></svg>
          <span>Upload<span class="item-sub">Load a .xpr or .html report file</span></span>
          <input type="file" id="upload-fresh" accept=".xpr,.html" style="display:none" onchange="loadReport(this);closeDrop('report-drop')">
        </label>
      </div>
    </div>

    <!-- ── Overrides dropdown ── -->
    <div class="hdr-drop" id="overrides-drop">
      <button class="hdr-btn overrides-btn" id="update-merchants-btn" onclick="toggleDrop('overrides-drop')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M20 7H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="1"/></svg>
        Overrides
        <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
        <span class="upd-badge"></span>
      </button>
      <div class="hdr-drop-menu" id="overrides-drop-menu">
        <button class="hdr-drop-item" onclick="updateMerchants();closeDrop('overrides-drop')" title="Download user_overrides.json with all current changes">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
          <span>Download<span class="item-sub">Save user_overrides.json</span></span>
        </button>
        <div class="hdr-drop-sep"></div>
        <label class="hdr-drop-item" for="upload-overrides">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
          <span>Upload<span class="item-sub">Load a user_overrides.json file</span></span>
          <input type="file" id="upload-overrides" accept=".json" style="display:none" onchange="loadOverridesFile(this);closeDrop('overrides-drop')">
        </label>
        <div class="hdr-drop-sep"></div>
        <div id="overrides-path-row" style="padding:.45rem 1rem .5rem;font-size:.68rem;color:var(--muted);line-height:1.4;word-break:break-all;max-width:280px">
          <span style="opacity:.6">Loaded from:</span><br>
          <span id="overrides-path-text" style="font-family:monospace;font-size:.65rem;opacity:.85"></span>
        </div>
      </div>
    </div>
  </div>
</header>
<div id="unsaved-banner" style="display:none;background:#fef3c7;border-bottom:1.5px solid #fcd34d;padding:.6rem 2.5rem;display:none;align-items:center;gap:.75rem;font-size:.82rem;color:#92400e;font-weight:500">
  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="flex-shrink:0;color:#d97706"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
  <span>You have unsaved changes — download the updated <strong>Report</strong> and <strong>Overrides</strong> before closing to preserve all decisions.</span>
  <button onclick="document.getElementById('unsaved-banner').style.display='none'" style="margin-left:auto;background:none;border:none;cursor:pointer;color:#92400e;font-size:1rem;line-height:1;padding:0 .25rem">✕</button>
</div>
<div class="summary-strip">
  <div class="stat"><div class="lbl">Total Spent</div><div class="val red" id="s-spent"></div></div>
  <div class="stat"><div class="lbl">Credits / Payments</div><div class="val grn" id="s-cred"></div></div>
  <div class="stat"><div class="lbl">Net Outflow</div><div class="val red" id="s-net"></div></div>
  <div class="stat"><div class="lbl">Transactions</div><div class="val" id="s-tx"></div></div>
  <div class="stat"><div class="lbl">Avg / Month</div><div class="val red" id="s-avg-month"></div></div>
  <div class="stat"><div class="lbl">Generated</div><div class="val" style="font-size:.95rem" id="s-gen"></div></div>
</div>

<nav class="tab-bar">
  <button class="tab-btn" data-view="overview" onclick="switchView('overview')">
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="13" width="4" height="8" rx="1"/><rect x="10" y="8" width="4" height="13" rx="1"/><rect x="17" y="3" width="4" height="18" rx="1"/></svg>
    Overview
  </button>
  <button class="tab-btn active" data-view="all" onclick="switchView('all')">
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 3v9l5 3"/></svg>
    All Spending
  </button>
  <button class="tab-btn" data-view="monthly" onclick="switchView('monthly')">
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>
    Monthly Spending
  </button>
  <button class="tab-btn" data-view="subscriptions" onclick="switchView('subscriptions')">
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 10c0 7-9 13-9 13S3 17 3 10a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
    Subscriptions
    <span class="badge" id="sub-badge" style="display:none"></span>
  </button>
  <button class="tab-btn" data-view="review" onclick="switchView('review')">
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
    Review
    <span class="badge" id="review-badge"></span>
  </button>
  <button class="tab-btn" data-view="customize" onclick="switchView('customize')">
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
    Customize
  </button>
</nav>

<main class="main">

  <!-- ═══════════════ OVERVIEW ═══════════════ -->
  <div class="view" id="view-overview"></div>

  <!-- ═══════════════ REVIEW ═══════════════ -->
  <div class="view" id="view-review">

    <!-- Uncategorised section -->
    <div class="section-card" id="uncat-section" style="display:none">
      <div class="section-head collapsible" onclick="toggleCollapse('uncat-section')">
        <h2>Uncategorised Transactions
          <svg class="collapse-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
        </h2>
        <span id="uncat-chip" class="chip" style="color:var(--amber);border-color:#fcd34d;background:#fffbeb"></span>
      </div>
      <div class="section-body">
        <div class="review-intro" style="background:#fff7ed;border-color:#fed7aa;color:#92400e">
          <strong>These transactions did not match any known category.</strong>
          Select the correct category for each and click <strong>Apply</strong> to update the charts live.
        </div>
        <div id="uncat-list"></div>
      </div>
    </div>

    <!-- Reclassify flagged categories -->
    <div class="section-card" id="reclassify-panel" style="display:none">
      <div class="section-head">
        <h2>Flagged for Re-classification</h2>
        <span class="chip" style="color:var(--amber);border:1.5px solid #fcd34d;background:#fffbeb">⛑ From chart flags</span>
      </div>
      <div class="section-body">
        <div class="review-intro" style="background:#fffbeb;border-color:#fcd34d;color:#92400e">
          <strong>These merchants were flagged from the category bars or pie chart legend.</strong>
          Select a new category and click Reassign. Charts update live.
        </div>
        <div id="reclassify-list"></div>
      </div>
    </div>

    <!-- Questionable transactions section -->
    <div class="section-card" id="q-section">
      <div class="section-head collapsible" onclick="toggleCollapse('q-section')">
        <h2>Questionable Transactions
          <svg class="collapse-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
        </h2>
        <span id="q-summary-chip" class="chip tx"></span>
      </div>
      <div class="section-body">
        <div class="review-intro">
          <strong>These transactions appear in the credit column but match known expense merchants.</strong>
          They could be legitimate refunds, or expenses that were recorded in the wrong column by your bank.
          Use the toggle on each row to classify them as <strong>Debit (expense)</strong> or <strong>Credit (refund/payment)</strong>.
          All charts and totals update live.
        </div>
        <div id="q-list"></div>
        <div id="q-apply-wrap" style="display:none;margin-top:1.2rem;padding-top:1rem;border-top:1px solid var(--border);text-align:right">
          <span style="font-size:.82rem;color:var(--muted);margin-right:1rem">Decisions are applied live to all charts. Click to lock and clear the badge.</span>
          <button class="apply-btn apply-q" onclick="applyQuestionable()">&#10003; Mark as Reviewed</button>
        </div>
      </div>
    </div>

  </div>

  <!-- ═══════════════ ALL SPENDING ═══════════════ -->
  <div class="view active" id="view-all">
    <div class="section-card">
      <div class="section-head">
        <h2>Spending by Category — All Time</h2>
        <select id="all-acct-filter" onchange="setAcctFilter(this.value)"
                style="font-family:inherit;font-size:.82rem;font-weight:600;color:var(--ink2);
                       padding:.4rem .7rem;border-radius:var(--radius-sm);border:1.5px solid var(--border);
                       background:#fff;cursor:pointer"></select>
      </div>
      <div class="section-body">
        <div class="cat-bar-list" id="all-cat-bars"></div>
      </div>
    </div>
    <div class="section-card">
      <div class="section-head">
        <h2>Spending by Category — Over Time</h2>
        <span id="over-time-filter-label" style="font-size:.8rem;color:var(--muted)"></span>
      </div>
      <div class="section-body" style="padding-bottom:.75rem">
        <div style="position:relative;height:340px;width:100%">
          <canvas id="over-time-chart"></canvas>
        </div>
        <div id="over-time-legend" style="display:flex;flex-wrap:wrap;gap:.5rem;margin-top:1rem;padding-top:.75rem;border-top:1px solid var(--border)"></div>
      </div>
    </div>
    <div class="section-card">
      <div class="section-head">
        <h2>All Transactions</h2>
        <div style="display:flex;gap:.7rem;align-items:center;flex-wrap:wrap">
          <button class="clear-filter-btn" id="clear-all-cat-btn" onclick="clearAllCategoryFilter()">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            Clear filter
          </button>
          <span class="tx-count-badge" id="all-tx-count"></span>
          <input class="search-box" id="all-search-box" type="text" placeholder="Search transactions..." oninput="filterAllTx(this.value)"/>
        </div>
      </div>
      <div style="padding:0;height:60vh;overflow-y:auto;overflow-x:auto">
          <table class="tx-table" style="min-width:600px">
            <thead><tr><th class="sortable" style="position:sticky;top:0;z-index:2;background:#1e3a5f" onclick="sortAllTx('date')">Date<span class="sort-icon" id="all-sort-icon-date">⇅</span></th><th style="position:sticky;top:0;z-index:2;background:#1e3a5f">Account</th><th style="position:sticky;top:0;z-index:2;background:#1e3a5f">Transaction</th><th style="position:sticky;top:0;z-index:2;background:#1e3a5f;width:36px;padding:0 .3rem"></th><th style="position:sticky;top:0;z-index:2;background:#1e3a5f">Category</th><th class="amt sortable" style="position:sticky;top:0;z-index:2;background:#1e3a5f" onclick="sortAllTx('debit')">Spent<span class="sort-icon" id="all-sort-icon-debit">⇅</span></th><th class="amt sortable" style="position:sticky;top:0;z-index:2;background:#1e3a5f" onclick="sortAllTx('credit')">Credit<span class="sort-icon" id="all-sort-icon-credit">⇅</span></th></tr></thead>
            <tbody id="all-tx-body"></tbody>
          </table>
      </div>
    </div>
  </div>

  <!-- ═══════════════ MONTHLY ═══════════════ -->
  <div class="view" id="view-monthly">
    <div class="month-picker-row">
      <label for="month-select">Select month:</label>
      <select id="month-select" onchange="renderMonthly(this.value)"></select>
      <div class="month-stat-chips" id="month-chips"></div>
    </div>
    <div class="section-card" id="pivot-section-card">
      <div class="section-head"><h2>Category Breakdown by Month</h2></div>
      <div class="section-body" style="padding:0">
        <div class="pivot-wrap"><table id="pivot-table"></table></div>
      </div>
    </div>
    <div class="section-card">
      <div class="section-head">
        <h2>Spending Distribution &nbsp;<span id="pie-month-label" style="font-family:'DM Sans',sans-serif;font-size:.85rem;font-weight:500;color:var(--muted)"></span></h2>
      </div>
      <div class="section-body" style="overflow:visible">
        <div class="pie-wrap">
          <div class="pie-canvas-wrap">
            <canvas id="pie-chart"></canvas>
          </div>
          <div class="pie-legend" id="pie-legend"></div>
        </div>
      </div>
    </div>
    <div class="section-card">
      <div class="section-head">
        <h2 id="month-tx-title">Transactions</h2>
        <div style="display:flex;gap:.7rem;align-items:center;flex-wrap:wrap">
          <button class="clear-filter-btn" id="clear-cat-btn" onclick="clearCategoryFilter()">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            Clear filter
          </button>
          <span class="tx-count-badge" id="month-tx-count"></span>
          <input class="search-box" id="month-search-box" type="text" placeholder="Search..." oninput="filterMonthTx(this.value)"/>
        </div>
      </div>
      <div style="padding:0;height:60vh;overflow-y:auto;overflow-x:auto">
          <table class="tx-table" style="min-width:600px">
            <thead><tr><th class="sortable" style="position:sticky;top:0;z-index:2;background:#1e3a5f" onclick="sortMonthTx('date')">Date<span class="sort-icon" id="month-sort-icon-date">⇅</span></th><th style="position:sticky;top:0;z-index:2;background:#1e3a5f">Account</th><th style="position:sticky;top:0;z-index:2;background:#1e3a5f">Transaction</th><th style="position:sticky;top:0;z-index:2;background:#1e3a5f;width:36px;padding:0 .3rem"></th><th style="position:sticky;top:0;z-index:2;background:#1e3a5f">Category</th><th class="amt sortable" style="position:sticky;top:0;z-index:2;background:#1e3a5f" onclick="sortMonthTx('debit')">Spent<span class="sort-icon" id="month-sort-icon-debit">⇅</span></th><th class="amt sortable" style="position:sticky;top:0;z-index:2;background:#1e3a5f" onclick="sortMonthTx('credit')">Credit<span class="sort-icon" id="month-sort-icon-credit">⇅</span></th></tr></thead>
            <tbody id="month-tx-body"></tbody>
          </table>
      </div>
    </div>
  </div>

  <!-- ═══════════════ SUBSCRIPTIONS ═══════════════ -->
  <div class="view" id="view-subscriptions">
    <div id="sub-content"></div>
  </div>

  <!-- ═══════════════ CUSTOMIZE ═══════════════ -->
  <div class="view" id="view-customize">
    <div id="customize-section"></div>
  </div>

</main>

<!-- Month PDF picker modal -->
<script>
const RAW = __DATA_JSON__;

// ── Overrides path display ────────────────────────────────────────────────
(function() {
  var pathEl = document.getElementById("overrides-path-text");
  var rowEl  = document.getElementById("overrides-path-row");
  if (pathEl && RAW.overrides_path) {
    pathEl.textContent = RAW.overrides_path;
  } else if (rowEl) {
    rowEl.style.display = "none";
  }
})();

// ── Utilities ─────────────────────────────────────────────────────────────
function fmtAmt(v) {
  if (!v) return "-";
  return "$" + Math.abs(v).toLocaleString("en-CA",{minimumFractionDigits:2,maximumFractionDigits:2});
}

// Sets the Avg/Month stat in the summary strip.
// Must receive the current effective rows so it stays in sync with
// the displayed Total Spent (which is also computed from effective rows).
function _setAvgMonth(rows) {
  const el = document.getElementById("s-avg-month");
  if (!el) return;
  // Use effective rows when provided, fall back to all rows
  const r  = rows || getEffectiveRows();
  const mc = (RAW.months || []).length || 1;
  const totalDebit = r.reduce((s, row) => s + (row.debit || 0), 0);
  const avg = totalDebit / mc;
  el.textContent = "$" + avg.toLocaleString("en-CA", {minimumFractionDigits:2, maximumFractionDigits:2});
}

// ── State ─────────────────────────────────────────────────────────────────
var pieChartInstance = null;
var overTimeChartInstance = null;
var currentMonthRows = [];
// overrides: REMOVED — debit/credit flips now stored by transaction signature in _liveFlips
// flaggedIdxs: REMOVED — flagged state now stored by transaction signature in _liveFlagged

// ── Transaction signature helper ─────────────────────────────────────────
// Identifies a real-world transaction stably across report regenerations.
// Key: "YYYY-MM-DD|merchant name|amount"  (amount is the raw CSV credit/debit value)
function _txSig(r) {
  return (r.date_iso || r.date || "") + "|" + (r.name || "") + "|" + (r.debit || r.credit || 0);
}
// Whether the Total Spent dashed line on the over-time chart is hidden
var _totalLineHidden = false;
// Set of row indices flagged as questionable — initialised before init() runs
var qidSet = new Set(RAW.questionable.map(q => q.qid));
var activeCategoryFilter = null;  // monthly tab category filter
var activeAllCategoryFilter = null;  // all-spending tab category filter
var activeAcctFilter = "All";  // all-spending tab account-type filter ("All" = no filter)

function getAccountFilteredRows(rows) {
  if (activeAcctFilter === "All") return rows;
  return rows.filter(r => (r.account_type || "Credit") === activeAcctFilter);
}

function setAcctFilter(val) {
  activeAcctFilter = val;
  const rows = getEffectiveRows();
  rebuildAllCatBars(rows);
  buildOverTimeChart(rows);
  const sb = document.getElementById("all-search-box");
  filterAllTx(sb ? sb.value : "");
}
var catOverrides = {};  // {rowIdx: newCategoryString} for in-session reclassifications
var qReviewed = false;  // true after Apply is clicked on the Review tab
// flaggedIdxs replaced by _liveFlagged (signature-based, in user_overrides.json)
var allSort   = {col: "date", dir: "desc"};    // current sort for All Spending table
var monthSort = {col: "date", dir: "desc"};    // current sort for Monthly table
var hiddenChartCats = new Set();  // categories toggled off in the over-time chart

// ── Live overrides — fetched from user_overrides.json at page load ────────
// _liveOverrides  : {merchantName: categoryString} — read from disk each open.
// _liveMerchants  : {merchantName: categoryString} — user-specific learned map.
// _liveSubBase    : [{name, pinned, dismissed, ...}] — subscription state from disk.
//
// These are separate from catOverrides (in-session changes not yet saved).
// getEffectiveRows() applies _liveOverrides then catOverrides in priority order,
// so in-session changes always win over what's on disk.
//
// "Save Overrides" merges _liveOverrides + catOverrides → new file on disk,
// then resets catOverrides and updates _liveOverrides to reflect the saved state.
var _liveOverrides = {};   // from user_overrides.json "overrides"
var _liveMerchants = {};   // from user_overrides.json "merchants"
var _liveFlips     = {};   // from user_overrides.json "debit_credit_flips" — sig→"debit"|"credit"
var _liveFlagged   = {};   // from user_overrides.json "flagged" — sig→true
var _overridesFileFound = false;  // false = fell back to baked snapshot

// ── Custom categories — populated from fetched user_overrides.json ────────
// Starts from baked snapshot; overwritten by fetch result in _applyLiveOverrides().
var customCategories = (RAW.custom_categories || []).map(cc => ({
  name:     cc.name,
  keywords: [...(cc.keywords || [])],
}));

// ── Subscription session state ─────────────────────────────────────────────
// _subBase / subPinned / subDismissed / subManual are set by _applyLiveOverrides()
// once the fetch resolves. These initial values are immediately overwritten.
var _liveSubBase = [];
var _subBase     = [];
var subPinned    = new Set();
var subDismissed = new Set();
var subManual    = [];

// Builds the merged subscriptions array for the JSON download
function buildMergedSubscriptions(base) {
  const byName = {};
  base.forEach(s => { byName[s.name] = Object.assign({}, s); });
  // Apply session pinned
  subPinned.forEach(name => {
    if (!byName[name]) byName[name] = { name };
    byName[name].pinned    = true;
    byName[name].dismissed = false;
    const m = subManual.find(s => s.name === name);
    if (m) { byName[name].monthlyAmt = m.monthlyAmt; byName[name].category = m.category; }
  });
  // Apply session dismissed
  subDismissed.forEach(name => {
    if (!byName[name]) byName[name] = { name };
    byName[name].dismissed = true;
    byName[name].pinned    = false;
  });
  // Remove entries that are neither pinned nor dismissed (clean slate)
  return Object.values(byName).filter(s => s.pinned || s.dismissed)
    .sort((a, b) => a.name.localeCompare(b.name));
}

// Whether subscription state has diverged from what was baked in
function _subsDiverged() {
  const origPinned    = new Set(_subBase.filter(s => s.pinned).map(s => s.name));
  const origDismissed = new Set(_subBase.filter(s => s.dismissed).map(s => s.name));
  if (subPinned.size !== origPinned.size || subDismissed.size !== origDismissed.size) return true;
  for (const n of subPinned)    { if (!origPinned.has(n)) return true; }
  for (const n of subDismissed) { if (!origDismissed.has(n)) return true; }
  return false;
}

// ── Baseline snapshot — set from fetched file, not baked snapshot ─────────
// _applyLiveOverrides() sets these after the fetch resolves.
// Used by _merchantsDiverged() to detect unsaved session changes.
var _origCustomCategoriesJson = JSON.stringify(customCategories);
var _origOverridesJson        = "{}";  // overwritten by _applyLiveOverrides()
var _origFlipsJson            = "{}";  // overwritten by _applyLiveOverrides()
var _origFlaggedJson          = "{}";  // overwritten by _applyLiveOverrides()

// Mirror of the Python CATEGORY_RENAMES — applied when "Save Overrides" writes
// the JSON so that any stale old-name overrides are migrated before saving.
var CATEGORY_RENAMES = {
  "Restaurants":                       "Restaurants, Pubs & Cafes",
  "Hotel, Entertainment & Recreation": "Entertainment & Recreation",
  "Home & Office Improvement":         "Electronics, Home & Office Improvement",
};

function _merchantsDiverged() {
  if (JSON.stringify(customCategories) !== _origCustomCategoriesJson) return true;
  if (Object.keys(catOverrides).length > 0) return true;
  if (JSON.stringify(_liveFlips)   !== _origFlipsJson)   return true;
  if (JSON.stringify(_liveFlagged) !== _origFlaggedJson) return true;
  if (_subsDiverged()) return true;
  return false;
}

// ── Effective row view (applies overrides to base data) ───────────────────
// Priority: debit/credit overrides → _liveOverrides (from file) → catOverrides
// (in-session). _liveOverrides gives the saved merchant→category map; catOverrides
// are the user's current-session changes not yet written to disk.
var _rowCache    = null;
var _cacheKey    = "";

// Version counter — incremented whenever _liveOverrides or _liveMerchants changes.
// Used in the row cache key instead of JSON.stringify(_liveOverrides) which would
// be O(n) on every single getEffectiveRows() call.
var _liveOverridesVersion = 0;

// Fast name→category lookup rebuilt whenever _liveOverrides/_liveMerchants changes
var _liveOverridesByName = {};
function _rebuildLiveByName() {
  _liveOverridesByName = Object.assign({}, _liveMerchants, _liveOverrides);
  _liveOverridesVersion++;
}

function _overridesKey() {
  return JSON.stringify(catOverrides) + "|" + _liveOverridesVersion;
}

function getEffectiveRows() {
  const key = _overridesKey();
  if (_rowCache && key === _cacheKey) return _rowCache;
  _rowCache = RAW.rows.map(r => {
    let row = r;

    // 1. Debit/Credit flip — keyed by transaction signature, persistent in user_overrides
    const sig  = _txSig(r);
    const flip = _liveFlips[sig];
    if (flip) {
      const amt = r.debit > 0 ? r.debit : r.credit;
      row = Object.assign({}, r, {
        debit:  flip === "debit"  ? amt : 0,
        credit: flip === "credit" ? amt : 0,
      });
    }

    // 2. In-session reclassification (highest category priority)
    if (catOverrides[r.idx]) {
      row = Object.assign({}, row, { category: catOverrides[r.idx] });
    } else {
      // 3. Persistent category override from user_overrides.json
      const liveCat = _liveOverridesByName[row.name];
      if (liveCat && liveCat !== row.category) {
        row = Object.assign({}, row, { category: liveCat });
      }
    }
    return row;
  });
  _cacheKey = key;
  return _rowCache;
}

function _invalidateRowCache() { _rowCache = null; _cacheKey = ""; }

function computeSummary(rows) {
  // Single-pass: accumulate both totals together
  let total_spent = 0, total_cred = 0;
  for (let i = 0; i < rows.length; i++) {
    total_spent += rows[i].debit;
    total_cred  += rows[i].credit;
  }
  return { total_spent, total_cred, net: total_spent - total_cred, tx_count: rows.length };
}

function computeCatTotals(rows) {
  const t = {};
  getActiveCategories(rows).forEach(c => t[c] = 0);
  rows.forEach(r => { if (r.debit > 0) t[r.category] = (t[r.category]||0) + r.debit; });
  return t;
}

function computePivot(rows) {
  const p = {};
  RAW.months.forEach(m => { p[m] = {}; });
  rows.forEach(r => {
    if (r.debit > 0 && p[r.month]) p[r.month][r.category] = (p[r.month][r.category]||0) + r.debit;
  });
  return p;
}

// Returns only categories that have at least one debit tx in the current effective rows,
// preserving the original defined order: built-ins first, then custom cats (dark grey), then Other.
function getActiveCategories(rows) {
  const withSpend = new Set(rows.filter(r => r.debit > 0).map(r => r.category));
  const customNames = new Set((RAW.custom_categories || []).map(cc => cc.name));
  const builtinActive = RAW.categories.filter(c =>
    withSpend.has(c) && !customNames.has(c) && c !== "Other / Uncategorised");
  const customActive = [...customNames].filter(c => withSpend.has(c));
  const otherActive = withSpend.has("Other / Uncategorised") ? ["Other / Uncategorised"] : [];
  const allKnown = new Set([...RAW.categories, ...customNames, "Other / Uncategorised"]);
  const novelActive = [...withSpend].filter(c => !allKnown.has(c));
  return [...builtinActive, ...novelActive, ...customActive, ...otherActive];
}

// ── Master refresh — called after any toggle change ───────────────────────
function refreshAll() {
  _invalidateRowCache();
  const rows = getEffectiveRows();
  const s    = computeSummary(rows);

  document.getElementById("s-spent").textContent = "$" + s.total_spent.toLocaleString("en-CA",{minimumFractionDigits:2});
  document.getElementById("s-cred").textContent  = "$" + s.total_cred .toLocaleString("en-CA",{minimumFractionDigits:2});
  document.getElementById("s-net").textContent   = "$" + Math.abs(s.net).toLocaleString("en-CA",{minimumFractionDigits:2});
  _setAvgMonth(rows);

  const activeView = document.querySelector(".view.active").id;
  if (activeView === "view-all") {
    rebuildAllCatBars(rows);
    buildOverTimeChart(rows);
    buildAllTxTable(rows);
  } else if (activeView === "view-monthly") {
    const month = document.getElementById("month-select").value;
    renderMonthlyFromRows(rows, month);
  } else if (activeView === "view-subscriptions") {
    buildSubscriptionsTab();
  }
  // When not on the subscriptions tab, only refresh the badge count
  // (avoid full rebuild on every keypress/toggle)
  if (activeView !== "view-subscriptions") {
    _refreshSubBadgeOnly(rows);
  }
}

// Lightweight badge-only update — avoids full buildSubscriptionsTab() cost
function _refreshSubBadgeOnly(rows) {
  const badge = document.getElementById("sub-badge");
  if (!badge) return;
  const n = detectSubscriptions(rows || getEffectiveRows()).length;
  badge.textContent   = n;
  badge.style.display = n > 0 ? "inline-flex" : "none";
}


// ── _applyLiveOverrides — called after user_overrides.json is fetched ────
// Populates _liveOverrides, _liveMerchants, customCategories, sub state,
// and baseline snapshots. Called whether the source is the file or the baked fallback.
function _applyLiveOverrides(data, fromFile, onFileProtocol) {
  _liveOverrides      = data.overrides          || {};
  _liveMerchants      = data.merchants          || {};
  _liveFlips          = data.debit_credit_flips || {};
  _liveFlagged        = data.flagged            || {};
  _liveSubBase        = data.subscriptions      || [];
  _overridesFileFound = fromFile;

  // Custom categories from fetched file override the baked snapshot
  if (data.custom_categories && data.custom_categories.length > 0) {
    customCategories = data.custom_categories.map(cc => ({
      name:     cc.name,
      keywords: [...(cc.keywords || [])],
    }));
  }

  // Rebuild subscription state from fetched file
  _subBase     = _liveSubBase;
  subPinned    = new Set(_liveSubBase.filter(s => s.pinned).map(s => s.name));
  subDismissed = new Set(_liveSubBase.filter(s => s.dismissed).map(s => s.name));
  subManual    = _liveSubBase
    .filter(s => s.pinned && s.monthlyAmt)
    .map(s => ({ name: s.name, monthlyAmt: s.monthlyAmt, category: s.category || "" }));

  // Rebuild the fast name→category lookup used by getEffectiveRows()
  _rebuildLiveByName();
  _invalidateRowCache();

  // Set baselines from the fetched file — divergence badge compares against these
  _origCustomCategoriesJson = JSON.stringify(customCategories);
  _origOverridesJson        = JSON.stringify(_liveOverrides);
  _origFlipsJson            = JSON.stringify(_liveFlips);
  _origFlaggedJson          = JSON.stringify(_liveFlagged);

  // Only show a note when served over HTTP and user_overrides.json is genuinely absent.
  // On file:// protocol, fetch() always fails — that's expected, not an error.
  const note = document.getElementById("overrides-source-note");
  if (note) {
    if (fromFile || onFileProtocol) {
      note.textContent = "";
      note.style.display = "none";
    } else {
      // HTTP serve but file missing — soft informational note, not a warning
      note.textContent = "ℹ user_overrides.json not found — using baked snapshot";
      note.style.display = "inline";
    }
  }
  // Update review badge to reflect how many items the new overrides have resolved
  _refreshReviewBadge();
}

// ── Boot ──────────────────────────────────────────────────────────────────
// init() is async so it can fetch user_overrides.json before first render.
// The fetch uses a relative URL — works when the HTML and user_overrides.json
// are served from the same folder (e.g. python3 -m http.server).
// On file:// protocol most browsers block fetch(); falls back to baked snapshot.
(async function init() {
  // ── Static header content (no data dependency) ───────────────────────
  const s = RAW.summary;
  document.getElementById("meta-period").textContent = s.period;
  document.getElementById("s-spent").textContent = "$" + s.total_spent.toLocaleString("en-CA",{minimumFractionDigits:2});
  document.getElementById("s-cred").textContent  = "$" + s.total_cred .toLocaleString("en-CA",{minimumFractionDigits:2});
  document.getElementById("s-net").textContent   = "$" + Math.abs(s.net).toLocaleString("en-CA",{minimumFractionDigits:2});
  document.getElementById("s-tx").textContent    = s.tx_count;
  document.getElementById("s-gen").textContent   = s.generated;

  // ── Fetch user_overrides.json — shared across all reports in the folder ─
  // cache: "no-store" ensures we always read the latest version from disk,
  // not a browser-cached copy from a previous report session.
  let fetchedData = null;
  const onFileProtocol = location.protocol === "file:";
  try {
    const resp = await fetch("user_overrides.json", { cache: "no-store" });
    if (resp.ok) fetchedData = await resp.json();
  } catch (_) { /* file:// CORS block or network error — fall through */ }

  if (fetchedData) {
    _applyLiveOverrides(fetchedData, true, onFileProtocol);
  } else {
    const baked = RAW.merchant_json || {};
    _applyLiveOverrides(baked, false, onFileProtocol);
    if (!onFileProtocol) {
      console.warn(
        "xPence: could not fetch user_overrides.json — using baked snapshot.\n" +
        "To get live shared overrides, serve your reports folder over HTTP:\n" +
        "  cd /path/to/reports && python3 -m http.server 8080"
      );
    }
  }

  // ── Build all UI (uses live override state loaded above) ─────────────
  _setAvgMonth(getEffectiveRows());
  buildOverviewTab();
  buildUncatSection();
  buildFlaggedReclassifyPanel();
  buildReviewTab();
  buildSubscriptionsTab();
  const initRows = getEffectiveRows();
  rebuildAllCatBars(initRows);
  buildOverTimeChart(initRows);
  buildAllTxTable(initRows);

  const sel = document.getElementById("month-select");
  RAW.months.forEach(m => {
    const o = document.createElement("option"); o.value = m; o.textContent = m; sel.appendChild(o);
  });
  sel.value = RAW.months[RAW.months.length - 1];
  renderMonthly(sel.value);

  // Account-type filter for the All Spending tab
  const acctSel = document.getElementById("all-acct-filter");
  if (acctSel) {
    const allOpt = document.createElement("option");
    allOpt.value = "All"; allOpt.textContent = "All accounts";
    acctSel.appendChild(allOpt);
    (RAW.account_types || []).forEach(t => {
      const o = document.createElement("option"); o.value = t; o.textContent = t;
      acctSel.appendChild(o);
    });
    acctSel.value = "All";
    acctSel.style.display = (RAW.account_types || []).length > 1 ? "" : "none";
  }

  // Show review badge only for unresolved questionable rows
  const badge = document.getElementById("review-badge");
  _refreshReviewBadge();
  _refreshUpdateBadge();
})();

// ── View switching ────────────────────────────────────────────────────────
function switchView(v) {
  document.querySelectorAll(".view").forEach(el => el.classList.remove("active"));
  document.querySelectorAll(".tab-btn").forEach(el => el.classList.remove("active"));
  document.getElementById("view-" + v).classList.add("active");
  document.querySelector(`[data-view="${v}"]`).classList.add("active");
  window.scrollTo({top: 0, behavior: "instant"});  // each tab starts at the top

  // Refresh the view we just switched into with current override state
  const rows = getEffectiveRows();
  if (v === "all") {
    rebuildAllCatBars(rows);
    buildOverTimeChart(rows);
    buildAllTxTable(rows);
  } else if (v === "monthly") {
    const month = document.getElementById("month-select").value;
    renderMonthlyFromRows(rows, month);
  } else if (v === "subscriptions") {
    buildSubscriptionsTab();
  } else if (v === "customize") {
    buildCustomizePanel();
  }
}

// ── OVERVIEW TAB ─────────────────────────────────────────────────────────
function buildOverviewTab() {
  const ov  = RAW.overview;
  const s   = RAW.summary;
  const fmt = v => "$" + Math.abs(v).toLocaleString("en-CA",{minimumFractionDigits:2,maximumFractionDigits:2});
  const prd = s.period;

  // ── helpers ──
  function card(accentColor, titleIcon, titleText, bodyHtml) {
    return `<div class="ov-card">
      <div class="ov-card-title">${titleIcon}${titleText}</div>
      ${bodyHtml}
    </div>`;
  }
  function icon(path, vb="0 0 24 24") {
    return `<svg width="13" height="13" viewBox="${vb}" fill="none" stroke="currentColor" stroke-width="2">${path}</svg>`;
  }

  // ── 1. Spending snapshot ──
  const trendDir   = ov.mom_trend ? ov.mom_trend.dir : null;
  const trendPct   = ov.mom_trend ? Math.abs(ov.mom_trend.pct) : 0;
  const trendHtml  = trendDir
    ? `<span class="ov-trend-${trendDir === "up" ? "up" : "dn"}" style="font-weight:800;font-family:'DM Serif Display',serif">`
      + `${trendDir === "up" ? "↑" : "↓"} ${trendPct}%</span>`
    : `<span style="font-weight:800;font-family:'DM Serif Display',serif;color:var(--muted)">—</span>`;

  const snapshotCard = card("#2563eb",
    icon('<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 9h6M9 12h6M9 15h4"/>'),
    "Spending Snapshot",
    `<div class="ov-hero-row">
       <div class="ov-hero-cell">
         <div class="ov-hero-label">Total spent</div>
         <div class="ov-hero-value">${fmt(s.total_spent)}</div>
         <div class="ov-hero-sub">${ov.n_months} month${ov.n_months!==1?"s":""} · ${prd}</div>
       </div>
       <div class="ov-hero-cell">
         <div class="ov-hero-label">Avg / month</div>
         <div class="ov-hero-value">${fmt(ov.avg_month_amt)}</div>
         <div class="ov-hero-sub">over ${ov.n_months} month${ov.n_months!==1?"s":""}</div>
       </div>
       <div class="ov-hero-cell">
         <div class="ov-hero-label">MoM trend</div>
         <div class="ov-hero-value">${trendHtml}</div>
         <div class="ov-hero-sub">${trendDir ? `vs. prior 2 months` : "not enough data"}</div>
       </div>
     </div>`
  );

  // ── 2. Top categories ──
  const maxCatTotal = ov.top_cats.length ? ov.top_cats[0].pct : 1;
  const catsHtml = ov.top_cats.map(tc => {
    const color = RAW.cat_color_map[tc.cat] || "#64748b";
    return `<div class="ov-cat-row">
      <div class="ov-cat-dot" style="background:${color}"></div>
      <div class="ov-cat-name">${tc.cat}</div>
      <div class="ov-cat-bar-wrap"><div class="ov-cat-bar" style="width:${(tc.pct/maxCatTotal*100).toFixed(0)}%;background:${color}"></div></div>
      <div class="ov-cat-pct">${tc.pct}%</div>
    </div>`;
  }).join("");
  const topCatCard = card("#ea580c",
    icon('<path d="M18 20V10M12 20V4M6 20v-6"/>'),
    "Top Spending Categories",
    catsHtml
  );

  // ── 3. Timing patterns ──
  const weekLabel = ["first","second","third","fourth"][ov.busiest_week-1] || "last";
  const timingCard = card("#7c3aed",
    icon('<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>'),
    "When You Spend Most",
    `<div class="ov-hero-row">
       <div class="ov-hero-cell">
         <div class="ov-hero-label">Busiest month</div>
         <div class="ov-hero-value">${ov.peak_month}</div>
         <div class="ov-hero-sub">${fmt(ov.peak_month_amt)} spent</div>
       </div>
       <div class="ov-hero-cell">
         <div class="ov-hero-label">Quietest month</div>
         <div class="ov-hero-value">${ov.quiet_month}</div>
         <div class="ov-hero-sub">${fmt(ov.quiet_month_amt)} spent</div>
       </div>
       <div class="ov-hero-cell">
         <div class="ov-hero-label">Peak pattern</div>
         <div class="ov-hero-value">${ov.busiest_dow}s</div>
         <div class="ov-hero-sub">${weekLabel} week of month</div>
       </div>
     </div>`
  );

  // ── 4. Favourite merchants — top merchant per each of the top 5 spend categories ──
  let merchantsHtml = "";
  const top5cats = ov.top_cats.slice(0, 5);
  top5cats.forEach(tc => {
    const color = RAW.cat_color_map[tc.cat] || "#64748b";
    const top   = (ov.merch_by_cat[tc.cat] || [])[0];
    if (!top) return;
    const nameStr = top.name.length > 36 ? top.name.slice(0, 36) + "…" : top.name;
    merchantsHtml += `<div class="ov-merch-row" style="display:flex;align-items:center;gap:.55rem;padding:.42rem 0;border-bottom:1px solid var(--border)">
      <span style="width:8px;height:8px;border-radius:50%;background:${color};flex-shrink:0;display:inline-block"></span>
      <div style="flex:1;min-width:0">
        <div style="font-size:.84rem;font-weight:600;color:var(--ink2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${nameStr}</div>
        <div style="font-size:.7rem;color:var(--muted)">${tc.cat}</div>
      </div>
      <div style="font-size:.72rem;color:var(--muted);flex-shrink:0">${top.visits}×</div>
      <div class="ov-merch-amt" style="flex-shrink:0">${fmt(top.total)}</div>
    </div>`;
  });
  const merchantCard = card("#0d9488",
    icon('<path d="M3 3h18l-2 13H5L3 3z"/><path d="M8 21h8M12 17v4"/>'),
    "Top Merchant per Category",
    merchantsHtml || '<div class="ov-sub" style="color:var(--muted)">Not enough data.</div>'
  );

  // ── 5. Unusual transactions ──
  let unusualHtml = "";
  if (ov.unusual.length) {
    ov.unusual.forEach(u => {
      unusualHtml += `<div class="ov-unusual-row">
        <div class="ov-unusual-name">${u.name.length>50?u.name.slice(0,50)+"…":u.name}
          <span class="ov-badge">${u.z}σ above avg</span>
        </div>
        <div class="ov-unusual-meta">${u.date} · ${u.cat} · <strong>${fmt(u.amount)}</strong></div>
      </div>`;
    });
  } else {
    unusualHtml = '<div class="ov-sub" style="color:var(--muted)">No statistically unusual transactions detected — your spending is very consistent.</div>';
  }
  const unusualCard = card("#dc2626",
    icon('<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>'),
    "Unusual Spends",
    unusualHtml
  );

  // ── 6. Transaction profile ──
  const maxTx = ov.max_tx;
  const profileCard = card("#ca8a04",
    icon('<circle cx="12" cy="12" r="10"/><path d="M12 8v4l3 3"/>'),
    "Transaction Profile",
    `<div class="ov-sub" style="margin-bottom:.6rem">Average transaction size: <strong>${fmt(ov.avg_tx)}</strong></div>
     ${maxTx && maxTx.name ? `<div class="ov-sub">Largest single purchase: <strong>${fmt(maxTx.amount)}</strong> at 
     <strong>${maxTx.name.length>40?maxTx.name.slice(0,40)+"…":maxTx.name}</strong> 
     (${maxTx.date}, ${maxTx.cat})</div>` : ""}
     <div class="ov-sub" style="margin-top:.6rem">Total transactions recorded: <strong>${s.tx_count}</strong> over ${ov.n_months} month${ov.n_months!==1?"s":""} 
     — roughly <strong>${Math.round(s.tx_count/ov.n_months)}</strong> per month.</div>`
  );

  // ── 7. Spending consistency ──
  let consistHtml = "";
  const consistEntries = Object.entries(ov.consistency);
  if (consistEntries.length) {
    // Sort stable first (lowest CV)
    const sorted = [...consistEntries].sort((a, b) => a[1].cv - b[1].cv);
    const stableCats = sorted.filter(([, d]) => d.cv < 0.4).map(([c]) => c);

    if (stableCats.length) {
      consistHtml += `<div style="font-size:.78rem;color:#16a34a;font-weight:600;margin-bottom:.7rem;display:flex;align-items:center;gap:.4rem">
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg>
        Stable: ${stableCats.map(c => `<span style="font-weight:400">${c}</span>`).join(" · ")}
      </div>`;
    }

    // ── Shared axis: all bars use the same scale so means land on the same
    //    vertical line and variance widths are directly comparable.
    // The shared domain is [0, globalMax] where globalMax = max(mean + std) across all cats.
    // We pin the "mean" column at 50% of the track width — everything to the
    // left is the lower half of the domain, everything to the right is upper.
    // This way bars with wider std visually extend further from center.
    const globalMax = Math.max(...sorted.map(([, d]) => d.mean + d.std + 10));
    // MEAN_COL: fraction of track where the shared mean axis sits (50% = center)
    const MEAN_COL = 0.5;

    // ── Compute label column width from longest category name ──────────────
    // We measure at render time using a canvas 2D context (zero-layout cost).
    // This gives an exact pixel width so both the header row and all data rows
    // share one identical grid-template-columns value — no RAF / reflow needed.
    // Measure label column width. Canvas rem values are ignored (treated as px),
    // so we resolve 0.83rem manually: base font is typically 16px → 13.3px.
    // We use a char-width fallback of 8.5px per char which is safe for DM Sans 500
    // at 13.3px, then add 8px padding so no label is ever clipped.
    const _ctx = (function(){
      try {
        const c = document.createElement("canvas");
        const ctx = c.getContext("2d");
        ctx.font = "500 13.3px 'DM Sans', sans-serif";
        return ctx;
      } catch(e) { return null; }
    })();
    const labelColPx = sorted.reduce((max, [cat]) => {
      const w = _ctx ? Math.ceil(_ctx.measureText(cat).width) : cat.length * 8.5;
      return Math.max(max, w);
    }, 0) + 8; // +8px breathing room so no label is clipped

    const gridCols = `${labelColPx}px 1fr 120px`;

    sorted.forEach(([cat, d]) => {
      const label = d.cv < 0.4 ? "Stable" : d.cv < 0.8 ? "Variable" : "Erratic";
      const cls   = d.cv < 0.4 ? "consist-low" : d.cv < 0.8 ? "consist-med" : "consist-high";
      const color = RAW.cat_color_map[cat] || "#64748b";

      const pxPerDollar = MEAN_COL / (d.mean || 1);
      const stdHalfW    = Math.min(MEAN_COL, pxPerDollar * d.std);
      const stdLeft     = (MEAN_COL - stdHalfW) * 100;
      const stdWidth    = stdHalfW * 2 * 100;
      const minLeft     = Math.max(0,   (MEAN_COL - pxPerDollar * (d.mean - d.min)) * 100);
      const maxRight    = Math.min(100, (MEAN_COL + pxPerDollar * (d.max - d.mean)) * 100);

      consistHtml += `
      <div style="display:grid;grid-template-columns:${gridCols};align-items:center;gap:.75rem;padding:.52rem 0;border-bottom:1px solid var(--border)">
        <div style="white-space:nowrap;overflow:visible">
          <div style="font-size:.83rem;font-weight:500;color:var(--ink2)">${cat}</div>
          <span class="ov-consist-label ${cls}" style="margin-top:.18rem;display:inline-block">${label}</span>
        </div>
        <div style="position:relative;height:10px;background:#f1f5f9;border-radius:4px;overflow:hidden">
          <div style="position:absolute;top:2px;height:6px;border-radius:3px;background:${color}22;left:${minLeft}%;width:${maxRight-minLeft}%"></div>
          <div style="position:absolute;top:1px;height:8px;border-radius:3px;background:${color}55;left:${stdLeft}%;width:${stdWidth}%"></div>
          <div style="position:absolute;top:0;height:100%;width:2.5px;background:${color};border-radius:1px;left:calc(${MEAN_COL*100}% - 1.5px)"></div>
          <div style="position:absolute;top:0;height:100%;width:1px;background:#cbd5e1;left:${MEAN_COL*100}%"></div>
        </div>
        <div style="font-size:.74rem;color:var(--muted);text-align:right;white-space:nowrap">
          <strong style="color:var(--ink);font-size:.8rem">${fmtAmt(d.mean)}</strong>
          <span style="display:block">±${fmtAmt(d.std)}</span>
        </div>
      </div>`;
    });

    // Header row uses same gridCols so first column is perfectly aligned
    consistHtml = `
    <div style="display:grid;grid-template-columns:${gridCols};gap:.75rem;margin-bottom:.3rem">
      <div></div>
      <div style="position:relative;font-size:.68rem;color:var(--muted);text-align:center">
        ← less &nbsp;&nbsp; <strong>mean</strong> &nbsp;&nbsp; more →
      </div>
      <div style="font-size:.68rem;color:var(--muted);text-align:right">avg ± std</div>
    </div>` + consistHtml;
  }
  const consistCard = card("#64748b",
    icon('<path d="M22 12h-4l-3 9L9 3l-3 9H2"/>'),
    "Month-to-Month Consistency",
    consistHtml || '<div class="ov-sub" style="color:var(--muted)">Not enough months to assess.</div>'
  );

  // ── Assemble grid ──
  const grid1 = `<div class="ov-grid">${snapshotCard}${timingCard}${profileCard}</div>`;
  const grid2 = `<div class="ov-grid">${topCatCard}${merchantCard}${unusualCard}</div>`;
  const grid3 = `<div class="ov-grid" style="grid-template-columns:1fr">${consistCard}</div>`;

  // ── Tip card ──
  const tipCard = `<div class="ov-tip-card">
    <div class="ov-tip-label">💡 Spending Habit Tip</div>
    <div class="ov-tip-text">${ov.spend_tip}</div>
  </div>`;

  document.getElementById("view-overview").innerHTML =
    `<div style="margin-bottom:.5rem;font-family:'DM Serif Display',serif;font-size:1.05rem;color:var(--ink2)">
       Your spending summary for <strong>${prd}</strong>
     </div>` +
    grid1 + grid2 + grid3 + tipCard;

  // Align hero cells within each hero-row to the widest label in that row,
  // and fix the consistency card label column to the widest category name.
  requestAnimationFrame(() => {
    document.querySelectorAll("#view-overview .ov-hero-row").forEach(heroRow => {
      const cells = heroRow.querySelectorAll(".ov-hero-cell");
      let maxLabelW = 0;
      cells.forEach(cell => {
        const lbl = cell.querySelector(".ov-hero-label");
        if (lbl) maxLabelW = Math.max(maxLabelW, lbl.scrollWidth);
      });
      if (maxLabelW > 0) {
        cells.forEach(cell => { cell.style.minWidth = maxLabelW + "px"; });
      }
    });

  });
}

// ── Collapsible sections ──────────────────────────────────────────────────
function toggleCollapse(cardId) {
  const card = document.getElementById(cardId);
  if (card) card.classList.toggle("collapsed");
}

// ── CUSTOMIZE PANEL ───────────────────────────────────────────────────────

function buildCustomizePanel() {
  const el = document.getElementById("customize-section");
  if (!el) return;

  // ── 1. Category Colours — compact dropdown row ─────────────────────────
  const rows = getEffectiveRows();
  const activeCats = getActiveCategories(rows);

  const colorOpts = activeCats.map(cat => {
    const safeCat = escHtml(cat);
    return `<option value="${safeCat}">${safeCat}</option>`;
  }).join("");

  // Pick initial color from first active cat
  const firstCat  = activeCats[0] || "";
  const firstColor = RAW.cat_color_map[firstCat] || "#64748b";

  const colorsSection = `
  <div class="cust-section" style="margin-bottom:1.8rem">
    <div class="cust-head">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--muted)"><circle cx="13.5" cy="6.5" r="2.5"/><circle cx="19" cy="13" r="2.5"/><circle cx="6.5" cy="13" r="2.5"/><circle cx="11" cy="19.5" r="2.5"/></svg>
      <h2>Category Colours</h2>
    </div>
    <div class="cust-body">
      <div class="cust-color-compact">
        <div id="cust-color-swatch" style="width:26px;height:26px;border-radius:50%;background:${firstColor};border:2px solid rgba(0,0,0,.08);flex-shrink:0"></div>
        <select id="cust-color-cat-select" onchange="onColorCatChange(this.value)" style="flex:1;min-width:160px;font-family:inherit;font-size:.87rem;padding:.42rem .75rem;border:1.5px solid var(--border);border-radius:var(--radius-sm);background:var(--surface);color:var(--ink)">
          ${colorOpts}
        </select>
        <input type="color" id="cust-color-picker-input" class="cust-color-picker" value="${firstColor}"
          oninput="previewCatColor(document.getElementById('cust-color-cat-select').value, this.value)"
          onchange="applyCatColor(document.getElementById('cust-color-cat-select').value, this.value); syncColorSwatch()"/>
        <button class="cust-color-reset" onclick="resetCatColor(document.getElementById('cust-color-cat-select').value); syncColorSwatch()">↺ Reset</button>
      </div>
    </div>
  </div>`;

  // ── 2. User Overrides section ─────────────────────────────────────────────
  // Shows all persistent overrides from _liveOverrides + current session catOverrides
  // Each card has a delete button; deletion removes from _liveOverrides / catOverrides.
  const allOverrides = {};  // merchant → { cat, source }
  // Layer 1: live overrides from file
  Object.entries(_liveOverrides).forEach(([m, c]) => { allOverrides[m] = { cat: c, source: "saved" }; });
  // Layer 2: session catOverrides (by idx → name)
  const idxToName = {};
  (RAW.rows || []).forEach(r => { idxToName[r.idx] = r.name; });
  Object.entries(catOverrides).forEach(([idx, cat]) => {
    const name = idxToName[idx];
    if (name) allOverrides[name] = { cat, source: "session" };
  });

  const ovrEntries = Object.entries(allOverrides).sort(([a],[b]) => a.localeCompare(b));
  let ovrHtml = "";
  if (ovrEntries.length === 0) {
    ovrHtml = '<div class="ovr-empty">No overrides yet — re-assign a transaction category to create one.</div>';
  } else {
    ovrHtml = '<div class="ovr-list">' + ovrEntries.map(([merchant, info]) => {
      const catColor = RAW.cat_color_map[info.cat] || "#64748b";
      const badge = info.source === "session"
        ? '<span class="ovr-card-source">unsaved</span>'
        : '<span class="ovr-card-source">saved</span>';
      const safeMerchant = merchant.replace(/\\/g,"\\\\").replace(/'/g,"\\'");
      return `
      <div class="ovr-card" id="ovr-card-${escHtml(merchant)}">
        <div class="ovr-card-merchant" title="${escHtml(merchant)}">${escHtml(merchant)}</div>
        <div class="ovr-card-arrow">→</div>
        <div class="ovr-card-cat" style="background:${catColor}22;color:${catColor}">${escHtml(info.cat)}</div>
        ${badge}
        <button class="ovr-card-del" onclick="deleteOverride('${safeMerchant}')" title="Remove this override">✕</button>
      </div>`;
    }).join("") + "</div>";
  }

  const overridesSection = `
  <div class="cust-section" style="margin-bottom:1.8rem">
    <div class="cust-head">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--muted)"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
      <h2>Category Overrides</h2>
      <span style="font-size:.78rem;color:var(--muted)">${ovrEntries.length} assignment${ovrEntries.length !== 1 ? "s" : ""}</span>
    </div>
    <div class="cust-body">
      <div class="cust-intro" style="margin-bottom:1rem">
        These are your merchant→category assignments. <strong>Saved</strong> entries come from
        <code>user_overrides.json</code>; <strong>unsaved</strong> entries are from this session only
        and will be included when you download via <strong>Overrides → Download</strong>.
      </div>
      ${ovrHtml}
    </div>
  </div>`;

  // ── 3. Custom Categories section ─────────────────────────────────────────
  let listHtml = "";
  if (customCategories.length === 0) {
    listHtml = '<div class="cust-empty">No custom categories yet — add one below.</div>';
  } else {
    listHtml = '<div class="cust-cat-list">' +
      customCategories.map((cc, i) => {
        const kwChips = cc.keywords.map(k =>
          `<span class="cust-kw-chip">${escHtml(k)}</span>`
        ).join("");
        const color = RAW.cat_color_map[cc.name] || "#3d3d3d";
        return `
        <div class="cust-cat-item" id="custcat-${i}" style="border-left-color:${color}">
          <div class="cust-cat-dot" style="background:${color}"></div>
          <div class="cust-cat-info">
            <div class="cust-cat-name">${escHtml(cc.name)}</div>
            <div class="cust-kw-chips">${kwChips || '<span style="font-size:.75rem;color:var(--muted)">no keywords</span>'}</div>
          </div>
          <button class="cust-cat-del" onclick="deleteCustomCategory(${i})" title="Delete this category">✕ Remove</button>
        </div>`;
      }).join("") +
    '</div>';
  }

  const formHtml = `
  <div class="cust-add-form">
    <div class="cust-add-title">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="vertical-align:middle;margin-right:.35rem"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
      Add New Category
    </div>
    <div class="cust-field">
      <label for="cust-name-input">Category name</label>
      <input type="text" id="cust-name-input" placeholder="e.g. Pet Expenses" maxlength="60"/>
    </div>
    <div class="cust-field">
      <label for="cust-kw-input">Keywords <span style="font-weight:400;color:var(--muted)">(comma-separated)</span></label>
      <input type="text" id="cust-kw-input" placeholder="e.g. petco, banfield, petsmart, veterinar"/>
      <div class="field-hint">Transactions whose description contains any of these keywords will be assigned to this category. Case-insensitive.</div>
    </div>
    <button class="cust-add-btn" onclick="addCustomCategory()">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="vertical-align:middle;margin-right:.3rem"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
      Add Category
    </button>
  </div>`;

  const saveBar = `
  <div class="cust-save-bar">
    <div class="cust-save-note">
      <strong>To persist changes</strong>, use <strong>Overrides → Download</strong> in the header.
      Replace the existing <code>user_overrides.json</code> beside <code>xpence_analyzer.py</code>.
    </div>
    <span class="cust-feedback" id="cust-feedback" style="opacity:0"></span>
  </div>`;

  const customCatsSection = `
  <div class="cust-section">
    <div class="cust-head">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--muted)"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
      <h2>Custom Categories</h2>
      <span style="font-size:.78rem;color:var(--muted)">Define your own spending categories</span>
    </div>
    <div class="cust-body">
      <div class="cust-intro">
        Custom categories let you track spending that doesn't fit the built-in groups.
        They match by <strong>keyword</strong> and are saved to <code>user_overrides.json</code> when you use <strong>Overrides → Download</strong>.
      </div>
      ${listHtml}
      ${formHtml}
      ${saveBar}
    </div>
  </div>`;

  el.innerHTML = colorsSection + customCatsSection + overridesSection;
}

// ── Category color helpers ────────────────────────────────────────────────
var _defaultCatColors = null;

function _getDefaultCatColors() {
  if (_defaultCatColors) return _defaultCatColors;
  _defaultCatColors = Object.assign({}, RAW.cat_color_map);
  return _defaultCatColors;
}

// Called when the category <select> changes — syncs swatch + picker to selected cat
function onColorCatChange(cat) {
  const color = RAW.cat_color_map[cat] || "#64748b";
  const swatch = document.getElementById("cust-color-swatch");
  const picker = document.getElementById("cust-color-picker-input");
  if (swatch) swatch.style.background = color;
  if (picker) picker.value = color;
}

function syncColorSwatch() {
  const sel = document.getElementById("cust-color-cat-select");
  if (sel) onColorCatChange(sel.value);
}

function previewCatColor(cat, color) {
  const swatch = document.getElementById("cust-color-swatch");
  if (swatch) swatch.style.background = color;
}

function applyCatColor(cat, color) {
  RAW.cat_color_map[cat] = color;
  const idx = (RAW.categories || []).indexOf(cat);
  if (idx >= 0 && RAW.cat_colors) RAW.cat_colors[idx] = color;
  const rows = getEffectiveRows();
  rebuildAllCatBars(rows);
  buildOverTimeChart(rows);
  const month = document.getElementById("month-select") && document.getElementById("month-select").value;
  if (month) renderMonthlyFromRows(rows, month);
  // Rebuild panel but preserve select position
  const sel = document.getElementById("cust-color-cat-select");
  const prevCat = sel ? sel.value : null;
  buildCustomizePanel();
  if (prevCat) {
    const newSel = document.getElementById("cust-color-cat-select");
    if (newSel) { newSel.value = prevCat; onColorCatChange(prevCat); }
  }
}

function resetCatColor(cat) {
  const defaults = _getDefaultCatColors();
  if (defaults[cat]) applyCatColor(cat, defaults[cat]);
  syncColorSwatch();
}

// ── Override card management ──────────────────────────────────────────────
function deleteOverride(merchant) {
  let changed = false;
  if (_liveOverrides[merchant] !== undefined) {
    delete _liveOverrides[merchant];
    _rebuildLiveByName();
    changed = true;
  }
  // Also clear any in-session catOverride rows for this merchant
  (RAW.rows || []).forEach(r => {
    if (r.name === merchant && catOverrides[r.idx] !== undefined) {
      delete catOverrides[r.idx];
      changed = true;
    }
  });
  if (changed) {
    _invalidateRowCache();
    refreshAll();
    buildCustomizePanel();
    _refreshUpdateBadge();
  }
}

function escHtml(s) {
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
function addCustomCategory() {
  const nameEl = document.getElementById("cust-name-input");
  const kwEl   = document.getElementById("cust-kw-input");
  const name   = (nameEl.value || "").trim();
  const kwRaw  = (kwEl.value  || "").trim();

  if (!name) { nameEl.focus(); nameEl.style.borderColor = "var(--red)"; return; }
  nameEl.style.borderColor = "";

  // Reject if name clashes with a built-in or existing custom category
  const builtins = new Set(RAW.all_categories.map(c => c.toLowerCase()));
  if (builtins.has(name.toLowerCase())) {
    alert(`"${name}" is already a built-in category. Choose a different name.`); return;
  }
  if (customCategories.some(cc => cc.name.toLowerCase() === name.toLowerCase())) {
    alert(`A custom category named "${name}" already exists.`); return;
  }

  const keywords = kwRaw
    ? kwRaw.split(",").map(k => k.trim().toLowerCase()).filter(Boolean)
    : [];

  customCategories.push({ name, keywords });

  // Push into RAW so getActiveCategories / colour maps include it immediately
  RAW.custom_categories = customCategories;
  // Add to all_categories dropdown list so Review tab picks it up
  if (!RAW.all_categories.includes(name)) {
    RAW.all_categories.push(name);
  }
  // Assign dark-grey colour
  RAW.cat_color_map[name] = RAW.custom_cat_color || "#3d3d3d";

  // Re-classify every row that would match the new keywords
  if (keywords.length > 0) {
    RAW.rows.forEach(r => {
      // Only re-classify if currently Other/Uncategorised or a built-in (not another custom)
      const low = r.name.toLowerCase();
      const isCustomAlready = customCategories.slice(0,-1).some(cc => cc.name === r.category);
      if (!isCustomAlready && keywords.some(kw => low.includes(kw))) {
        if (r.category === "Other / Uncategorised") {
          catOverrides[r.idx] = name;
        }
      }
    });
  }

  // Clear inputs
  nameEl.value = "";
  kwEl.value   = "";

  buildCustomizePanel();
  // Rebuild uncategorised section since some rows may now be assigned
  buildUncatSection();
  _refreshReviewBadge();
  // Refresh all spending views
  const rows = getEffectiveRows();
  rebuildAllCatBars(rows);
  buildOverTimeChart(rows);
  buildAllTxTable(rows);
  const month = document.getElementById("month-select").value;
  renderMonthlyFromRows(rows, month);
  _showCustFeedback(`"${name}" added`);
  _refreshUpdateBadge();
}

function deleteCustomCategory(index) {
  const cc = customCategories[index];
  if (!cc) return;
  if (!confirm(`Remove custom category "${cc.name}"?\n\nTransactions assigned to it will revert to their original classification.`)) return;

  // Undo any catOverrides that pointed to this category
  Object.keys(catOverrides).forEach(idx => {
    if (catOverrides[idx] === cc.name) delete catOverrides[idx];
  });

  // Remove from runtime lists
  customCategories.splice(index, 1);
  RAW.custom_categories = customCategories;
  const aci = RAW.all_categories.indexOf(cc.name);
  if (aci > -1) RAW.all_categories.splice(aci, 1);
  delete RAW.cat_color_map[cc.name];

  buildCustomizePanel();
  buildUncatSection();
  _refreshReviewBadge();
  const rows = getEffectiveRows();
  rebuildAllCatBars(rows);
  buildOverTimeChart(rows);
  buildAllTxTable(rows);
  const month = document.getElementById("month-select").value;
  renderMonthlyFromRows(rows, month);
  _showCustFeedback(`"${cc.name}" removed`);
  _refreshUpdateBadge();
}

function _showCustFeedback(msg) {
  // Rebuild may have recreated the element — re-query after DOM update
  requestAnimationFrame(() => {
    const el = document.getElementById("cust-feedback");
    if (!el) return;
    el.textContent = "✓ " + msg;
    el.style.opacity = "1";
    setTimeout(() => { el.style.opacity = "0"; }, 2800);
  });
}

// ── Category options builder — built-ins first, custom cats grouped at bottom ──
function buildCategoryOptions() {
  const customNames = new Set((RAW.custom_categories || []).map(cc => cc.name));
  const builtinOpts = RAW.all_categories
    .filter(c => c !== "Other / Uncategorised" && !customNames.has(c))
    .map(c => `<option value="${c}">${c}</option>`)
    .join("");
  const customOpts = [...customNames]
    .map(c => `<option value="${c}">⬛ ${c} (custom)</option>`)
    .join("");
  if (!customOpts) return builtinOpts;
  return builtinOpts +
    `<optgroup label="─── Custom Categories ───" style="color:var(--muted)">` +
    customOpts + `</optgroup>`;
}

// ── UNCATEGORISED SECTION ────────────────────────────────────────────────
function buildUncatSection() {
  // Use effective rows so overrides-file assignments are reflected immediately
  const uncatRows = getEffectiveRows().filter(r => r.category === "Other / Uncategorised" && r.debit > 0);
  const section   = document.getElementById("uncat-section");
  const chip      = document.getElementById("uncat-chip");
  const list      = document.getElementById("uncat-list");
  if (uncatRows.length === 0) { section.style.display = "none"; return; }

  // Group by merchant — one card per unique merchant name
  const byMerchant = {};
  uncatRows.forEach(r => {
    const key = r.name.trim().toLowerCase();
    if (!byMerchant[key]) byMerchant[key] = { name: r.name, rows: [], total: 0 };
    byMerchant[key].rows.push(r);
    byMerchant[key].total += r.debit;
  });
  const merchants = Object.values(byMerchant);

  // Any merchant still in uncatRows after getEffectiveRows() is genuinely unresolved.
  // Only additionally exclude those with an in-session catOverride applied this session.
  const pending = merchants.filter(m => !m.rows.every(r => catOverrides[r.idx]));

  if (pending.length === 0) {
    // All have been categorised — hide the section entirely
    section.style.display = "none";
    return;
  }

  section.style.display = "";
  chip.textContent = pending.length + " merchant" + (pending.length > 1 ? "s" : "") +
    " (" + pending.reduce((s,m) => s + m.rows.length, 0) + " transaction" +
    (pending.reduce((s,m) => s + m.rows.length, 0) > 1 ? "s" : "") + ") need categorisation";

  const opts = buildCategoryOptions();

  // Only render pending merchants — already-categorised ones are hidden
  list.innerHTML = pending.map(m => {
    const rep = m.rows[0];
    const countBadge = m.rows.length > 1
      ? `<span style="font-size:.68rem;color:var(--muted);margin-left:.35rem">${m.rows.length}×</span>`
      : "";
    const idxJson = JSON.stringify(m.rows.map(r => r.idx));
    return `
    <div class="rev-card is-uncat" id="uccard-${rep.idx}">
      <div class="rev-meta">
        <span class="rev-date">${rep.date}</span>
        <span class="rev-name" title="${m.name}">${m.name.length > 52 ? m.name.slice(0,52) + "\u2026" : m.name}${countBadge}</span>
        <span class="rev-amt spend">${fmtAmt(m.total)}</span>
      </div>
      <div class="rev-controls">
        <select class="cat-select" id="ucsel-${rep.idx}">
          <option value="">— select category —</option>
          ${opts}
        </select>
        <button class="apply-btn apply-cat" id="ucbtn-${rep.idx}" onclick="applyMerchantOverride(${idxJson}, ${rep.idx}, null, null, null)">Apply to all</button>
        <div id="uctag-${rep.idx}"></div>
      </div>
    </div>`;
  }).join("");
}

// Unified handler: applies a category to one or more rows (by idx list).
// Used by both the Uncategorised panel (single/multiple rows per merchant)
// and the Reclassify panel. Accepts optional element ids for reclassify cards.
function applyMerchantOverride(idxList, repIdx, cardId, selId, btnId) {
  const selElemId = selId  || ("ucsel-" + repIdx);
  const btnElemId = btnId  || ("ucbtn-" + repIdx);
  const tagElemId = selId  ? ("retag-" + repIdx) : ("uctag-" + repIdx);
  const sel = document.getElementById(selElemId);
  if (!sel || !sel.value) return;
  idxList.forEach(i => { catOverrides[i] = sel.value; });

  // For the reclassify panel: show the tag and disable controls (those cards persist)
  if (selId) {
    const tag = document.getElementById(tagElemId);
    if (tag) tag.innerHTML = '<span class="applied-tag">✓ ' + sel.value + '</span>';
    sel.disabled = true;
    const btn = document.getElementById(btnElemId);
    if (btn) btn.disabled = true;
  } else {
    // For the uncat panel: remove the card entirely once applied
    const card = document.getElementById("uccard-" + repIdx);
    if (card) card.remove();
    // Hide section if no cards remain
    const remaining = RAW.rows.filter(r => r.category === "Other / Uncategorised" && r.debit > 0 && !catOverrides[r.idx]);
    if (remaining.length === 0) {
      document.getElementById("uncat-section").style.display = "none";
    } else {
      // Update chip count
      const chip = document.getElementById("uncat-chip");
      const remainingTx = remaining.length;
      const remainingMerchants = new Set(remaining.map(r => r.name.trim().toLowerCase())).size;
      chip.textContent = remainingMerchants + " merchant" + (remainingMerchants > 1 ? "s" : "") +
        " (" + remainingTx + " transaction" + (remainingTx > 1 ? "s" : "") + ") need categorisation";
    }
  }

  // Refresh all views
  const rows = getEffectiveRows();
  rebuildAllCatBars(rows);
  buildOverTimeChart(rows);
  buildAllTxTable(rows);
  const month = document.getElementById("month-select").value;
  renderMonthlyFromRows(rows, month);
  const s = computeSummary(rows);
  document.getElementById("s-spent").textContent = "$" + s.total_spent.toLocaleString("en-CA",{minimumFractionDigits:2});
  document.getElementById("s-cred").textContent  = "$" + s.total_cred .toLocaleString("en-CA",{minimumFractionDigits:2});
  document.getElementById("s-net").textContent   = "$" + Math.abs(s.net).toLocaleString("en-CA",{minimumFractionDigits:2});
  _setAvgMonth(rows);
  // Rebuild reclassify panel to reflect new assignments
  buildFlaggedReclassifyPanel();
  _refreshUpdateBadge();
}

// ── REVIEW TAB ────────────────────────────────────────────────────────────
function buildReviewTab() {
  const list = document.getElementById("q-list");
  const chip = document.getElementById("q-summary-chip");

  // Filter out items already resolved by a flip in the loaded overrides
  const pending = RAW.questionable.filter(q => !_liveFlips[_txSig(q)]);
  const allResolvedByFile = RAW.questionable.length > 0 && pending.length === 0;

  // No questionable transactions at all (baked data had none)
  if (RAW.questionable.length === 0) {
    list.innerHTML = '<div class="review-empty">No questionable transactions found. All credits look like genuine payments or refunds.</div>';
    chip.textContent = "0 flagged";
    chip.style.cssText = "";
    document.getElementById("q-apply-wrap").style.display = "none";
    return;
  }

  // All resolved — either by overrides file or by user clicking Apply this session
  if (allResolvedByFile || qReviewed) {
    const label = allResolvedByFile ? "\u2713 Accepted via overrides" : "\u2713 Reviewed";
    chip.textContent = label;
    chip.style.cssText = "color:#16a34a;border-color:#bbf7d0;background:#f0fdf4;border:1.5px solid;border-radius:999px;padding:.3rem .75rem;font-size:.78rem;font-weight:500";
    document.getElementById("q-apply-wrap").style.display = "none";
    list.innerHTML = '<div class="review-empty" style="color:#16a34a">' + label + ' — all debit/credit directions confirmed in loaded overrides.</div>';
    return;
  }

  // Some items still pending
  chip.textContent = pending.length + " flagged";
  chip.style.cssText = "";
  document.getElementById("q-apply-wrap").style.display = "block";

  // Group by month, preserving chronological order
  const byMonth = {};
  const monthOrder = [];
  pending.forEach(q => {
    if (!byMonth[q.month]) { byMonth[q.month] = []; monthOrder.push(q.month); }
    byMonth[q.month].push(q);
  });

  function qCardHtml(q) {
    const sig          = _txSig(q);
    const dir          = _liveFlips[sig] || "credit";
    const disabledAttr = qReviewed ? "disabled" : "";
    const appliedTag   = qReviewed
      ? `<span class="applied-tag">\u2713 ${dir === "debit" ? "Debit" : "Credit"}</span>`
      : "";
    return `
      <div class="rev-card is-${dir}" id="qcard-${q.qid}">
        <div class="rev-meta">
          <span class="rev-date">${q.date}</span>
          <span class="rev-name">${q.name.length > 52 ? q.name.slice(0,52) + "\u2026" : q.name}</span>
          <span class="rev-cat-badge">${q.category}</span>
          <span class="rev-amt${dir === "debit" ? " spend" : " credit"}" id="qamt-${q.qid}">${fmtAmt(q.amount)}</span>
        </div>
        <div class="rev-controls">
          <div class="toggle-wrap">
            <button class="toggle-btn${dir === "debit" ? " active-debit" : ""}" id="qbtn-debit-${q.qid}"
                    onclick="setOverride('${sig.replace(/'/g,"\\\\'")}', 'debit')" ${disabledAttr}>Debit</button>
            <div class="toggle-sep"></div>
            <button class="toggle-btn${dir === "credit" ? " active-credit" : ""}" id="qbtn-credit-${q.qid}"
                    onclick="setOverride('${sig.replace(/'/g,"\\\\'")}', 'credit')" ${disabledAttr}>Credit</button>
          </div>
          ${appliedTag}
        </div>
      </div>`;
  }

  list.innerHTML = monthOrder.map(month => {
    const qs    = byMonth[month];
    const total = qs.reduce((s, q) => s + q.amount, 0);
    const cards = qs.map(qCardHtml).join("");
    return `
      <div class="q-month-group">
        <div class="q-month-header">
          <span class="q-month-label">${month}</span>
          <span class="q-month-meta">${qs.length} transaction${qs.length > 1 ? "s" : ""} &nbsp;·&nbsp; ${fmtAmt(total)}</span>
        </div>
        ${cards}
      </div>`;
  }).join("");
}

function setOverride(sig, direction) {
  _liveFlips[sig] = direction;
  // Find all rows with this signature to update UI
  RAW.questionable.forEach(q => {
    if (_txSig(q) !== sig) return;
    const qid = q.qid;
    const card = document.getElementById("qcard-" + qid);
    if (card) card.className = "rev-card is-" + direction;
    const btnD = document.getElementById("qbtn-debit-"  + qid);
    const btnC = document.getElementById("qbtn-credit-" + qid);
    if (btnD) btnD.className = "toggle-btn" + (direction === "debit"  ? " active-debit"  : "");
    if (btnC) btnC.className = "toggle-btn" + (direction === "credit" ? " active-credit" : "");
    const amt = document.getElementById("qamt-" + qid);
    if (amt) amt.className = "rev-amt" + (direction === "debit" ? " spend" : " credit");
  });
  _invalidateRowCache();
  refreshAll();
  _refreshUpdateBadge();
}

function applyQuestionable() {
  qReviewed = true;
  // Rebuild the entire section — it now renders in fully-reviewed state
  buildReviewTab();
  // Clear the tab badge
  const badge = document.getElementById("review-badge");
  badge.textContent = "";
  badge.style.display = "none";
}


// ── Table sorting ─────────────────────────────────────────────────────────
function sortRows(rows, col, dir) {
  const sorted = [...rows].sort((a, b) => {
    let av = a[col], bv = b[col];
    if (col === "date") {
      // date strings are "Jan 01, 2026" — compare as Date objects
      av = new Date(av); bv = new Date(bv);
    }
    if (av < bv) return dir === "asc" ? -1 : 1;
    if (av > bv) return dir === "asc" ?  1 : -1;
    return 0;
  });
  return sorted;
}

function updateSortIcons(prefix, col, dir) {
  ["date","debit","credit"].forEach(c => {
    const el = document.getElementById(prefix + "-sort-icon-" + c);
    const th = el ? el.closest("th") : null;
    if (!el || !th) return;
    if (c === col) {
      el.textContent = dir === "asc" ? "↑" : "↓";
      th.classList.remove("sort-asc", "sort-desc");
      th.classList.add(dir === "asc" ? "sort-asc" : "sort-desc");
    } else {
      el.textContent = "⇅";
      th.classList.remove("sort-asc", "sort-desc");
    }
  });
}

function sortAllTx(col) {
  if (allSort.col === col) {
    allSort.dir = allSort.dir === "asc" ? "desc" : "asc";
  } else {
    allSort.col = col;
    allSort.dir = col === "date" ? "desc" : "desc";
  }
  updateSortIcons("all", allSort.col, allSort.dir);
  const sb = document.getElementById("all-search-box");
  filterAllTx(sb ? sb.value : "");
}

function sortMonthTx(col) {
  if (monthSort.col === col) {
    monthSort.dir = monthSort.dir === "asc" ? "desc" : "asc";
  } else {
    monthSort.col = col;
    monthSort.dir = col === "date" ? "desc" : "desc";
  }
  updateSortIcons("month", monthSort.col, monthSort.dir);
  const sb = document.getElementById("month-search-box");
  filterMonthTx(sb ? sb.value : "");
}


// ── Flag category for re-review ──────────────────────────────────────────
function flagTransaction(idx) {
  const r   = (RAW.rows || []).find(r => r.idx === idx);
  const sig = r ? _txSig(r) : null;
  if (!sig) return;
  const wasFlagged = !!_liveFlagged[sig];
  if (wasFlagged) { delete _liveFlagged[sig]; }
  else            { _liveFlagged[sig] = true; }
  const rows = getEffectiveRows();
  buildAllTxTable(rows);
  const month = document.getElementById("month-select").value;
  renderMonthlyFromRows(rows, month);
  buildFlaggedReclassifyPanel();
  if (!wasFlagged) {
    const uncatSection = document.getElementById("uncat-section");
    if (uncatSection && uncatSection.style.display !== "none") {
      uncatSection.classList.add("collapsed");
    }
    switchView("review");
  }
  _refreshUpdateBadge();
}

function applyTxOverride(idx) {
  const sel = document.getElementById("resel-" + idx);
  if (!sel || !sel.value) return;
  catOverrides[idx] = sel.value;
  sel.disabled = true;
  const btn = document.getElementById("rebtn-" + idx);
  if (btn) btn.disabled = true;
  const tag = document.getElementById("retag-" + idx);
  if (tag) tag.innerHTML = '<span class="applied-tag">\u2192 ' + sel.value + '</span>';
  const rows = getEffectiveRows();
  rebuildAllCatBars(rows);
  buildOverTimeChart(rows);
  buildAllTxTable(rows);
  const month = document.getElementById("month-select").value;
  renderMonthlyFromRows(rows, month);
  const sm = computeSummary(rows);
  document.getElementById("s-spent").textContent = "$" + sm.total_spent.toLocaleString("en-CA",{minimumFractionDigits:2});
  document.getElementById("s-cred").textContent  = "$" + sm.total_cred .toLocaleString("en-CA",{minimumFractionDigits:2});
  document.getElementById("s-net").textContent   = "$" + Math.abs(sm.net).toLocaleString("en-CA",{minimumFractionDigits:2});
  _setAvgMonth(rows);
  _refreshUpdateBadge();
}

function buildFlaggedReclassifyPanel() {
  const panel = document.getElementById("reclassify-panel");
  if (!panel) return;
  const flaggedSigs = Object.keys(_liveFlagged);
  if (flaggedSigs.length === 0) { panel.style.display = "none"; return; }
  panel.style.display = "";

  const rowMap = {};
  getEffectiveRows().forEach(r => { rowMap[_txSig(r)] = r; });

  let html = "";
  flaggedSigs.forEach(sig => {
    const r = rowMap[sig];
    const idx = r ? r.idx : null;
    if (!r || r.debit <= 0) return;
    const already = catOverrides[idx];
    const color   = RAW.cat_color_map[r.category] || "#64748b";
    const opts    = buildCategoryOptions();
    html += `<div class="rev-card is-flag" id="recard-${idx}">
      <div class="rev-meta">
        <span class="rev-date">${r.date}</span>
        <span class="rev-name" title="${r.name}">${r.name.length > 52 ? r.name.slice(0,52) + "\u2026" : r.name}</span>
        <span class="rev-cat-badge" style="background:${color}22;color:${color}">${r.category}</span>
        <span class="rev-amt spend">${fmtAmt(r.debit)}</span>
      </div>
      <div class="rev-controls">
        <select class="cat-select" id="resel-${idx}" ${already ? "disabled" : ""}>
          <option value="${r.category}" selected>${r.category} (current)</option>
          ${opts}
        </select>
        <button class="apply-btn apply-cat" id="rebtn-${idx}"
          onclick="applyTxOverride(${idx})" ${already ? "disabled" : ""}>Reassign</button>
        <button class="remove-btn" onclick="flagTransaction(${idx})" title="Remove flag">\u2715</button>
        <div id="retag-${idx}">${already ? '<span class="applied-tag">\u2192 ' + catOverrides[idx] + '</span>' : ""}</div>
      </div>
    </div>`;
  });

  document.getElementById("reclassify-list").innerHTML = html ||
    '<p style="color:var(--muted);font-size:.82rem">No transactions flagged.</p>';
}

// ── ALL SPENDING ──────────────────────────────────────────────────────────
function rebuildAllCatBars(rows) {
  rows = getAccountFilteredRows(rows);
  const totals = computeCatTotals(rows);
  const grand  = Object.values(totals).reduce((a,b)=>a+b,0) || 1;
  const max_v  = Math.max(...Object.values(totals)) || 1;
  const sorted = [...getActiveCategories(rows)].sort((a,b)=>totals[b]-totals[a]);
  const container = document.getElementById("all-cat-bars");
  container.innerHTML = "";
  sorted.forEach((cat,i) => {
    const v      = totals[cat] || 0;
    const pct    = (v/grand*100).toFixed(1);
    const w      = (v/max_v*100).toFixed(1);
    const color  = RAW.cat_color_map[cat] || RAW.cat_colors[i % RAW.cat_colors.length];
    const active = activeAllCategoryFilter === cat ? " active-filter" : "";
    const safecat = cat.replace(/'/g, "\\'");
    container.innerHTML += `
      <div class="cat-bar-item${active}" data-cat="${cat}" onclick="toggleAllCategoryFilter('${safecat}')">
        <div class="cat-name" title="${cat}">${cat}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${w}%;background:${color}"></div></div>
        <div class="bar-amt">${fmtAmt(v)}</div>
        <div class="bar-pct">${pct}%</div>
      </div>`;
  });
  // Align all cat-name labels to the width of the widest one
  requestAnimationFrame(() => {
    const labels = container.querySelectorAll(".cat-name");
    let maxW = 0;
    labels.forEach(el => { maxW = Math.max(maxW, el.scrollWidth); });
    labels.forEach(el => { el.style.width = maxW + "px"; });
  });
}

function buildAllTxTable(rows) {
  rows = getAccountFilteredRows(rows);
  const sorted = sortRows(rows, allSort.col, allSort.dir);
  document.getElementById("all-tx-count").textContent = sorted.length + " transactions";
  document.getElementById("all-tx-body").innerHTML = sorted.map(r => txRow(r)).join("");
}

function filterAllTx(q) {
  const lq = q ? q.toLowerCase() : "";
  let rows = getEffectiveRows();
  if (activeAllCategoryFilter) rows = rows.filter(r => r.category === activeAllCategoryFilter);
  if (lq) rows = rows.filter(r =>
    r.name.toLowerCase().includes(lq) || r.category.toLowerCase().includes(lq) || r.date.toLowerCase().includes(lq));
  buildAllTxTable(rows);
}

function toggleAllCategoryFilter(cat) {
  activeAllCategoryFilter = (activeAllCategoryFilter === cat) ? null : cat;
  updateAllClearBtn();
  const rows = getEffectiveRows();
  rebuildAllCatBars(rows);
  buildOverTimeChart(rows);
  const sb = document.getElementById("all-search-box");
  filterAllTx(sb ? sb.value : "");
}

function clearAllCategoryFilter() {
  activeAllCategoryFilter = null;
  updateAllClearBtn();
  const rows = getEffectiveRows();
  rebuildAllCatBars(rows);
  buildOverTimeChart(rows);
  const sb = document.getElementById("all-search-box");
  filterAllTx(sb ? sb.value : "");
}

function updateAllClearBtn() {
  const btn = document.getElementById("clear-all-cat-btn");
  if (btn) btn.classList.toggle("visible", activeAllCategoryFilter !== null);
}


// ── Over-time line chart ──────────────────────────────────────────────────
function buildOverTimeChart(rows) {
  rows = getAccountFilteredRows(rows);
  const months = RAW.months;
  const cats   = getActiveCategories(rows);
  const filter = activeAllCategoryFilter;

  // Per-month totals per category
  const pivot = {};
  months.forEach(m => { pivot[m] = {}; cats.forEach(c => pivot[m][c] = 0); });
  rows.forEach(r => { if (r.debit > 0 && pivot[r.month]) pivot[r.month][r.category] += r.debit; });

  // Per-month grand total (sum across all categories)
  const monthTotals = months.map(m =>
    cats.reduce((s, c) => s + (pivot[m][c] || 0), 0)
  );

  // Filter label
  const lbl = document.getElementById("over-time-filter-label");
  lbl.textContent = filter ? `Filtered: ${filter}` : "";

  // The sentinel label for the total line
  const TOTAL_LABEL = "Total Spent";

  // Datasets — bake hidden state directly into each dataset
  let datasets;
  if (!filter) {
    datasets = cats.map(cat => ({
      label:           cat,
      data:            months.map(m => pivot[m][cat] || 0),
      borderColor:     RAW.cat_color_map[cat] || "#111111",
      backgroundColor: (RAW.cat_color_map[cat] || "#111111") + "22",
      tension:         0.35,
      pointRadius:     4,
      pointHoverRadius:6,
      borderWidth:     2,
      fill:            false,
      hidden:          hiddenChartCats.has(cat),
    }));

    // Add dashed total line — always rendered; visibility driven by _totalLineHidden
    datasets.push({
      label:           TOTAL_LABEL,
      data:            monthTotals,
      borderColor:     "#000000",
      backgroundColor: "transparent",
      borderDash:      [6, 4],
      tension:         0.35,
      pointRadius:     3,
      pointHoverRadius:5,
      borderWidth:     2,
      fill:            false,
      hidden:          _totalLineHidden,
    });
  } else {
    const color = RAW.cat_color_map[filter] || "#111111";
    datasets = [{
      label:           filter,
      data:            months.map(m => pivot[m][filter] || 0),
      borderColor:     color,
      backgroundColor: color + "22",
      tension:         0.35,
      pointRadius:     5,
      pointHoverRadius:7,
      borderWidth:     2.5,
      fill:            true,
      hidden:          false,
    }];
  }

  // Delta plugin (only active when filtered to a single category)
  const deltaPlugin = {
    id: "deltaLabels",
    afterDatasetsDraw(chart) {
      if (!filter) return;
      const ds   = chart.data.datasets[0];
      const meta = chart.getDatasetMeta(0);
      const ctx  = chart.ctx;
      ctx.save();
      ctx.font = "bold 11px 'DM Sans', sans-serif";
      ctx.textAlign = "center";
      for (let i = 1; i < meta.data.length; i++) {
        const prev = ds.data[i - 1];
        const curr = ds.data[i];
        if (!prev) continue;
        const pct   = (curr - prev) / prev * 100;
        const label = (pct >= 0 ? "+" : "") + pct.toFixed(0) + "%";
        const col   = pct >= 0 ? "#dc2626" : "#16a34a";
        const px    = meta.data[i].x;
        const py    = meta.data[i].y - 14;
        const tw    = ctx.measureText(label).width;
        ctx.fillStyle = col + "22";
        ctx.beginPath(); ctx.roundRect(px - tw/2 - 4, py - 11, tw + 8, 16, 3); ctx.fill();
        ctx.fillStyle = col;
        ctx.fillText(label, px, py);
      }
      ctx.restore();
    }
  };

  // Destroy previous instance then create fresh
  if (overTimeChartInstance) { overTimeChartInstance.destroy(); overTimeChartInstance = null; }
  overTimeChartInstance = new Chart(
    document.getElementById("over-time-chart").getContext("2d"), {
      type: "line",
      data: { labels: months, datasets },
      options: {
        responsive:          true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        scales: {
          x: { grid: { color: "#e2e8f0" },
               ticks: { font: { family: "'DM Sans', sans-serif", size: 12 } } },
          y: { beginAtZero: true,
               grid: { color: "#e2e8f0" },
               ticks: { font: { family: "'DM Sans', sans-serif", size: 12 },
                        callback: v => "$" + v.toLocaleString("en-CA") } }
        },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: {
            label: c => ` ${c.dataset.label}: $${c.parsed.y.toLocaleString("en-CA",{minimumFractionDigits:2})}`
          }}
        }
      },
      plugins: [deltaPlugin],
    }
  );

  // Render decoupled legend buttons below the chart
  _renderOverTimeLegend(cats, filter);
}

// Renders pill buttons that look like legend items but are plain HTML —
// fully decoupled from the Chart.js canvas so no resize-loop can occur.
function _renderOverTimeLegend(cats, filter) {
  const container = document.getElementById("over-time-legend");
  if (!container) return;

  if (filter) { container.style.display = "none"; return; }
  container.style.display = "flex";

  const TOTAL_LABEL = "Total Spent";
  const allHidden   = cats.every(c => hiddenChartCats.has(c));

  // Category buttons
  const catBtns = cats.map(cat => {
    const color  = RAW.cat_color_map[cat] || "#111111";
    const off    = hiddenChartCats.has(cat);
    const safeCat = cat.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    return `<button
        class="chart-legend-btn${off ? " leg-off" : ""}"
        style="${off ? "" : `--leg-color:${color}`}"
        onclick="toggleChartCat('${safeCat}')">
      <span class="leg-dot"></span>
      <span class="leg-label">${cat}</span>
    </button>`;
  }).join("");

  // Total Spent button — independent toggle; off when _totalLineHidden, on otherwise.
  // Clicking it ONLY toggles the dashed line and does NOT filter the Tx list.
  const totalOff = _totalLineHidden;
  const totalBtn = `<button
      class="chart-legend-btn chart-legend-total${totalOff ? " leg-off" : ""}"
      onclick="toggleChartCat('${TOTAL_LABEL}')"
      title="Toggle the total spending line (does not filter transactions)"
      style="${totalOff ? "" : "--leg-color:#000000"}">
    <span class="leg-dot leg-dot-dashed"></span>
    <span class="leg-label">${TOTAL_LABEL}</span>
  </button>`;

  container.innerHTML = catBtns + totalBtn;
}

function toggleChartCat(cat) {
  const TOTAL_LABEL = "Total Spent";
  const cats = getActiveCategories(getEffectiveRows());

  if (cat === TOTAL_LABEL) {
    // Toggle the Total Spent line visibility independently of category filters.
    // This does NOT touch hiddenChartCats (which controls category lines)
    // and does NOT affect the Tx list at all.
    _totalLineHidden = !_totalLineHidden;
    // Rebuild chart in-place: update the hidden flag on the total dataset
    if (overTimeChartInstance) {
      const ds = overTimeChartInstance.data.datasets;
      const totalDs = ds.find(d => d.label === TOTAL_LABEL);
      if (totalDs) {
        totalDs.hidden = _totalLineHidden;
        overTimeChartInstance.update("none");
      }
    }
    // Re-render legend buttons only (no full chart rebuild needed)
    _renderOverTimeLegend(cats, activeAllCategoryFilter);
  } else {
    hiddenChartCats.has(cat) ? hiddenChartCats.delete(cat) : hiddenChartCats.add(cat);
    buildOverTimeChart(getEffectiveRows());
  }
}

// ── MONTHLY ───────────────────────────────────────────────────────────────
function renderMonthly(month) {
  renderMonthlyFromRows(getEffectiveRows(), month);
}

function renderMonthlyFromRows(rows, month) {
  const pivot      = computePivot(rows);
  const activeCats = getActiveCategories(rows);
  renderPivotTableFromPivot(pivot, month, activeCats);
  renderPieChartFromRows(rows, month, activeCats);
  renderMonthTxTable(rows, month);
}

function renderPivotTableFromPivot(pivot, selectedMonth, activeCats) {
  // ── Category emoji logos — shown as column headers; full name via CSS tooltip ──
  const CAT_LOGO = {
    "Restaurants, Pubs & Cafes":              "🍽️",
    "Travel":                                 "✈️",
    "Transportation":                         "🚌",
    "Entertainment & Recreation":             "🎭",
    "Retail & Grocery":                       "🛒",
    "Personal & Household Expenses":          "🏠",
    "Health":                                 "💊",
    "Gifts":                                  "🎁",
    "Electronics, Home & Office Improvement": "💻",
    "Professional & Financial Services":      "🏦",
    "Foreign Currency Transactions":          "💱",
    "Other / Uncategorised":                  "❓",
  };

  // ── yyyy-mmm formatter: "2024-03" → "2024-Mar" ──────────────────────────
  const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  function fmtMonthLabel(m) {
    if (!m || m.length < 7) return m;
    const [y, mo] = m.split("-");
    const idx = parseInt(mo, 10) - 1;
    return `${y}-${MONTH_NAMES[idx] || mo}`;
  }

  const rawCats = activeCats || RAW.categories;

  // ── Sort categories by their all-time total (desc) ────────────────────────
  // Use RAW.cat_totals_all when available (pre-computed Python-side),
  // fall back to summing the pivot for the current effective rows.
  const allTimeTotals = RAW.cat_totals_all || {};
  const cats = [...rawCats].sort((a, b) => {
    const ta = allTimeTotals[a] ?? Object.values(pivot).reduce((s, m) => s + (m[a] || 0), 0);
    const tb = allTimeTotals[b] ?? Object.values(pivot).reduce((s, m) => s + (m[b] || 0), 0);
    return tb - ta;
  });

  // ── Max column width cap via inline style ─────────────────────────────────
  const COL_MAX  = "72px";
  const COL_TOT  = "88px";
  const COL_DATE = "90px";

  // Build header: Date | Total | ...categories sorted by spend...
  let html = `<thead><tr>
    <th style="min-width:${COL_DATE};width:${COL_DATE};left:0">Month</th>
    <th class="amt" style="min-width:${COL_TOT};width:${COL_TOT};left:90px;border-right:1px solid #334155">Total</th>`;
  cats.forEach(c => {
    const logo = CAT_LOGO[c] || "📦";
    html += `<th class="amt" data-tip="${c}" style="max-width:${COL_MAX};white-space:normal;word-break:break-word;padding:.6rem .4rem">${logo}</th>`;
  });
  html += `</tr></thead><tbody>`;

  const colTots = {}; cats.forEach(c => colTots[c] = 0);
  let grandTot = 0;

  RAW.months.forEach(m => {
    const rowTot = cats.reduce((s, c) => s + (pivot[m]?.[c] || 0), 0);
    grandTot += rowTot;
    cats.forEach(c => { colTots[c] += (pivot[m]?.[c] || 0); });
    const sel = m === selectedMonth;
    const safeM = m.replace(/'/g, "\\'");
    html += `<tr onclick="selectMonthFromPivot('${safeM}')" style="${sel ? 'background:#eff6ff;' : ''}">`;
    html += `<td style="width:${COL_DATE};left:0"><strong>${fmtMonthLabel(m)}${sel ? ' ◀' : ''}</strong></td>`;
    html += `<td class="amt spend" style="width:${COL_TOT};left:90px">${fmtAmt(rowTot)}</td>`;
    cats.forEach(c => {
      const v = pivot[m]?.[c] || 0;
      html += `<td class="amt" style="max-width:${COL_MAX}">${v > 0 ? fmtAmt(v) : '-'}</td>`;
    });
    html += `</tr>`;
  });

  // Total footer row
  html += `<tr class="total-row">`;
  html += `<td style="width:${COL_DATE};left:0">TOTAL</td>`;
  html += `<td class="amt spend" style="width:${COL_TOT};left:90px">${fmtAmt(grandTot)}</td>`;
  cats.forEach(c => html += `<td class="amt" style="max-width:${COL_MAX}">${fmtAmt(colTots[c])}</td>`);
  html += `</tr></tbody>`;

  document.getElementById("pivot-table").innerHTML = html;

  // maxHeight/overflow handled entirely in CSS (.pivot-wrap)
}

function renderPieChartFromRows(rows, month, activeCats) {
  const cats   = activeCats || getActiveCategories(rows);
  const mrows  = rows.filter(r => r.month === month && r.debit > 0);
  const totals = {};
  cats.forEach(c => totals[c] = 0);
  mrows.forEach(r => totals[r.category] = (totals[r.category]||0) + r.debit);

  const labels=[], values=[], colors=[], borderColors=[];
  cats.forEach((cat,i) => {
    const v = totals[cat]||0;
    if (v > 0) {
      labels.push(cat); values.push(v);
      // Use cat_color_map first (covers custom cats + built-ins); fall back to palette
      const base = RAW.cat_color_map[cat] || RAW.cat_colors[i % RAW.cat_colors.length];
      colors.push(base + "cc");
      borderColors.push(base);
    }
  });

  document.getElementById("pie-month-label").textContent = month;
  if (pieChartInstance) { pieChartInstance.destroy(); pieChartInstance = null; }
  const old = document.getElementById("pie-chart");
  const neu = document.createElement("canvas"); neu.id = "pie-chart";
  old.parentNode.replaceChild(neu, old);
  pieChartInstance = new Chart(neu.getContext("2d"), {
    type: "doughnut",
    data: { labels, datasets:[{ data:values, backgroundColor:colors, borderColor:borderColors, borderWidth:2, hoverOffset:16 }] },
    options: {
      responsive:true, cutout:"58%",
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:ctx=>{
          const tot=ctx.dataset.data.reduce((a,b)=>a+b,0);
          return ` ${fmtAmt(ctx.parsed)}  (${(ctx.parsed/tot*100).toFixed(1)}%)`;
        }}}
      }
    }
  });
  const total = values.reduce((a,b)=>a+b,0)||1;
  const legendEl = document.getElementById("pie-legend");
  legendEl.innerHTML = labels.map((l,i)=>{
    const safeL  = l.replace(/'/g,"\\'");
    return `
    <div class="legend-item${activeCategoryFilter===l?' active-filter':''}" 
         onclick="toggleCategoryFilter('${safeL}')" id="legend-item-${i}">
      <div class="legend-dot" style="background:${borderColors[i]}"></div>
      <div class="legend-name">${l}</div>
      <div class="legend-amt">${fmtAmt(values[i])}</div>
      <div class="legend-pct">${(values[i]/total*100).toFixed(1)}%</div>
    </div>`;}).join("");;
}

function renderMonthTxTable(rows, month) {
  currentMonthRows = rows.filter(r => r.month === month);
  monthSort = {col: "date", dir: "desc"};  // reset sort for new month
  updateSortIcons("month", "date", "desc");
  const sb = document.getElementById("month-search-box");
  if (sb) sb.value = "";  // clear search box, but keep category filter
  document.getElementById("month-tx-title").textContent = `Transactions — ${month}`;
  filterMonthTx("");  // apply active category filter (if any) to the new month
  const spent = currentMonthRows.reduce((s,r)=>s+r.debit, 0);
  const cred  = currentMonthRows.reduce((s,r)=>s+r.credit,0);
  document.getElementById("month-chips").innerHTML = `
    <span class="chip spend">Spent ${fmtAmt(spent)}</span>
    <span class="chip credit">Credits ${fmtAmt(cred)}</span>
    <span class="chip tx">${currentMonthRows.length} transactions</span>`;
}

function buildMonthTxTable(rows) {
  const sorted = sortRows(rows, monthSort.col, monthSort.dir);
  document.getElementById("month-tx-count").textContent = sorted.length + " transactions";
  document.getElementById("month-tx-body").innerHTML = sorted.map(r => txRow(r)).join("");
}

function filterMonthTx(q) {
  const lq = q ? q.toLowerCase() : "";
  let rows = currentMonthRows;
  if (activeCategoryFilter) rows = rows.filter(r => r.category === activeCategoryFilter);
  if (lq) rows = rows.filter(r => r.name.toLowerCase().includes(lq) || r.category.toLowerCase().includes(lq));
  buildMonthTxTable(rows);
}

function toggleCategoryFilter(cat) {
  activeCategoryFilter = (activeCategoryFilter === cat) ? null : cat;
  updateClearBtn();
  // Re-highlight legend items
  document.querySelectorAll("#pie-legend .legend-item").forEach(el => {
    const name = el.querySelector(".legend-name").textContent;
    el.classList.toggle("active-filter", name === activeCategoryFilter);
  });
  const sb = document.getElementById("month-search-box");
  filterMonthTx(sb ? sb.value : "");
}

function clearCategoryFilter() {
  activeCategoryFilter = null;
  updateClearBtn();
  document.querySelectorAll("#pie-legend .legend-item").forEach(el => el.classList.remove("active-filter"));
  const sb = document.getElementById("month-search-box");
  filterMonthTx(sb ? sb.value : "");
}

function updateClearBtn() {
  const btn = document.getElementById("clear-cat-btn");
  if (btn) btn.classList.toggle("visible", activeCategoryFilter !== null);
}

// ── Pivot month selector ─────────────────────────────────────────────────
function selectMonthFromPivot(month) {
  const sel = document.getElementById("month-select");
  sel.value = month;
  renderMonthly(month);
}

// ── Shared row renderer ───────────────────────────────────────────────────
function txRow(r) {
  const isQ         = qidSet.has(r.idx);
  const reviewBadge = isQ ? '<span class="q-flag">? Review</span>' : '';
  const isFlagged   = !!_liveFlagged[_txSig(r)];
  const flagBtn     = `<button class="flag-btn${isFlagged ? ' flagged' : ''}" title="Flag for re-assignment"
    onclick="event.stopPropagation();flagTransaction(${r.idx})">⛑</button>`;
  const dbCell = r.debit  > 0 ? `<td class="amt spend">${fmtAmt(r.debit)}</td>`   : `<td class="amt" style="color:var(--muted)">-</td>`;
  const crCell = r.credit > 0 ? `<td class="amt credit">${fmtAmt(r.credit)}</td>` : `<td class="amt" style="color:var(--muted)">-</td>`;
  const name   = r.name.length > 55 ? r.name.slice(0,55) + "\u2026" : r.name;
  return `<tr>
    <td style="white-space:nowrap;color:var(--muted);font-size:.8rem">${r.date}</td>
    <td style="white-space:nowrap;font-size:.78rem;color:var(--ink2)"><span class="cat-tag" style="background:#334155">${r.account_type||'Credit'}</span></td>
    <td style="max-width:320px">${name}${reviewBadge}</td>
    <td style="width:36px;padding:0 .3rem;text-align:center">${flagBtn}</td>
    <td><span class="cat-tag" style="background:${RAW.cat_color_map[r.category]||'#64748b'}">${r.category}</span></td>
    ${dbCell}${crCell}
  </tr>`;
}

// ── Save / Load report ────────────────────────────────────────────────────

// ── Dropdown open/close helpers ───────────────────────────────────────────
function toggleDrop(dropId) {
  const menu = document.getElementById(dropId + "-menu");
  const isOpen = menu.classList.contains("open");
  // Close all dropdowns first
  document.querySelectorAll(".hdr-drop-menu").forEach(m => m.classList.remove("open"));
  if (!isOpen) menu.classList.add("open");
}
function closeDrop(dropId) {
  const menu = document.getElementById(dropId + "-menu");
  if (menu) menu.classList.remove("open");
}
// Close all dropdowns when clicking outside
document.addEventListener("click", function(e) {
  if (!e.target.closest(".hdr-drop")) {
    document.querySelectorAll(".hdr-drop-menu").forEach(m => m.classList.remove("open"));
  }
});

// ── Upload user_overrides.json from disk ──────────────────────────────────
function loadOverridesFile(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = JSON.parse(e.target.result);
      // Validate it looks like a user_overrides file
      if (typeof data !== "object" || Array.isArray(data)) throw new Error("Not a valid overrides file");
      qReviewed = false;  // let the overrides file determine reviewed state
      _applyLiveOverrides(data, true, location.protocol === "file:");
      _invalidateRowCache();
      refreshAll();
      buildCustomizePanel();
      buildUncatSection();
      buildFlaggedReclassifyPanel();
      buildReviewTab();
      _refreshReviewBadge();
      _refreshUpdateBadge();
      // Flash feedback
      const btn = document.getElementById("update-merchants-btn");
      if (btn) {
        const orig = btn.innerHTML;
        btn.innerHTML = btn.innerHTML.replace("Overrides", "✓ Loaded");
        setTimeout(() => { btn.innerHTML = orig; _refreshUpdateBadge(); }, 1800);
      }
    } catch(err) {
      alert("Could not load overrides file: " + err.message);
    }
  };
  reader.readAsText(file);
  input.value = "";
}

// ── Unsaved-changes banner + beforeunload guard ───────────────────────────
function _checkUnsavedWarning() {
  const diverged   = _merchantsDiverged();
  const banner     = document.getElementById("unsaved-banner");
  const reportBtn  = document.querySelector(".hdr-btn.report-btn");
  if (banner)    banner.style.display = diverged ? "flex" : "none";
  if (reportBtn) reportBtn.classList.toggle("has-changes", diverged);
}

window.addEventListener("beforeunload", function(e) {
  if (_merchantsDiverged()) {
    e.preventDefault();
    e.returnValue = "You have unsaved changes. Download your Report and Overrides before closing.";
  }
});

// ── Update Merchants badge ────────────────────────────────────────────────
function _refreshReviewBadge() {
  const badge = document.getElementById("review-badge");
  if (!badge) return;

  // Unresolved debit/credit flips
  const pendingFlips = RAW.questionable.filter(q => !_liveFlips[_txSig(q)]).length;

  // Unresolved uncategorised merchants (not in file overrides and not session-assigned)
  const uncatRows = RAW.rows.filter(r => r.category === "Other / Uncategorised" && r.debit > 0);
  const byMerchant = {};
  uncatRows.forEach(r => { const k = r.name.trim().toLowerCase(); if (!byMerchant[k]) byMerchant[k] = r; });
  const pendingUncat = Object.values(byMerchant)
    .filter(r => !_liveOverridesByName[r.name.trim()] && !catOverrides[r.idx]).length;

  // Flagged transactions awaiting reclassification
  const pendingFlagged = Object.keys(_liveFlagged).length;

  const total = pendingFlips + pendingUncat + pendingFlagged;
  if (total > 0) {
    badge.textContent = "!";
    badge.style.display = "inline-flex";
  } else {
    badge.textContent = "";
    badge.style.display = "none";
  }
}

function _refreshUpdateBadge() {
  const btn = document.getElementById("update-merchants-btn");
  if (!btn) return;
  const diverged = _merchantsDiverged();
  btn.classList.toggle("has-changes", diverged);

  if (diverged) {
    const nOverrides = Object.keys(catOverrides).length;
    const nCustom    = customCategories.length;
    const origCustom = JSON.parse(_origCustomCategoriesJson);
    const added   = customCategories.filter(c => !origCustom.some(o => o.name === c.name)).length;
    const removed = origCustom.filter(o => !customCategories.some(c => c.name === o.name)).length;
    const parts = [];
    if (nOverrides)  parts.push(`${nOverrides} category re-assignment(s)`);
    if (added)       parts.push(`${added} custom category addition(s)`);
    if (removed)     parts.push(`${removed} custom category deletion(s)`);
    btn.title = parts.join(", ") + " — click to download updated user_overrides.json";
  } else {
    btn.title = "No override changes since this report was generated.";
  }
  _checkUnsavedWarning();
}

// updateMerchants — builds and downloads a ready-to-use merchant_categories.json.
//
// The downloaded file is the authoritative replacement for the file on disk:
//   merchants         — keyword-classified per-name snapshot from this report run
//   overrides         — base overrides loaded from disk + any catOverrides made this
//                       session (manual re-assignments). Session wins on conflict.
//   custom_categories — the current customCategories list, treated as authoritative.
//                       Deletions ARE reflected — whatever is in the list right now
//                       is what gets written. Additions and deletions both work.
//
// The user drops this file next to xpence_analyzer.py — no flags required.
function updateMerchants() {
  if (!_merchantsDiverged()) {
    alert(
      "No changes to save.\n\n" +
      "user_overrides.json will be updated when you:\n" +
      "  • Re-assign a transaction's category (Review tab or ⛑ flag button)\n" +
      "  • Add or remove a custom category (Customize tab)\n" +
      "  • Confirm or delete a subscription"
    );
    return;
  }

  // ── Build the merged overrides section ───────────────────────────────────
  // Start from _liveOverrides (what was fetched from disk at page load),
  // then apply session catOverrides on top — session changes always win.
  const baseOverrides = Object.assign({}, _liveOverrides);
  const mergedOverrides = Object.assign({}, baseOverrides);

  // catOverrides maps rowIdx → categoryName. Resolve each idx to a merchant name
  // via RAW.rows, then write merchant name → category into overrides.
  const idxToName = {};
  (RAW.rows || []).forEach(r => { idxToName[r.idx] = r.name; });
  Object.entries(catOverrides).forEach(([idx, cat]) => {
    const name = idxToName[idx];
    if (name && cat) mergedOverrides[name] = cat;
  });

  // ── Migrate stale category names ─────────────────────────────────────────
  // Any overrides entry pointing at an old category name gets updated to the
  // new name before the file is written.
  let renamed = 0;
  Object.keys(mergedOverrides).forEach(merchant => {
    const old = mergedOverrides[merchant];
    if (CATEGORY_RENAMES[old]) { mergedOverrides[merchant] = CATEGORY_RENAMES[old]; renamed++; }
  });
  if (renamed > 0) console.log(`[updateMerchants] Migrated ${renamed} stale category name(s)`);

  // ── Purge stale overrides that point at deleted custom categories ──────────
  // If a custom category was deleted this session, any overrides entry whose
  // target matches that deleted name must be removed — otherwise the ghost
  // category name survives in the JSON and classify() will use it next run.
  // We also strip overrides pointing at names that were never valid built-ins
  // and are no longer in the current custom_categories list.
  const validCatNames = new Set([
    ...RAW.all_categories,                          // built-ins + Other
    ...customCategories.map(cc => cc.name),         // surviving custom cats
  ]);
  let purged = 0;
  Object.keys(mergedOverrides).forEach(merchant => {
    if (!validCatNames.has(mergedOverrides[merchant])) {
      delete mergedOverrides[merchant];
      purged++;
    }
  });
  if (purged > 0) {
    console.log(`[updateMerchants] Purged ${purged} stale override(s) pointing at deleted categories.`);
  }

  // ── Build the full JSON structure ─────────────────────────────────────────
  // Build the merged subscriptions list: base from _liveSubBase (fetched file),
  // then apply session pinned/dismissed state on top.
  const mergedSubscriptions = buildMergedSubscriptions(_liveSubBase);

  const userOverridesJson = {
    "_schema": "xpence-user-overrides-v2",
    "_note": (
      "User-specific overrides for xPence. Place beside xpence_analyzer.py.\n\n" +
      "  overrides          Explicit user category decisions (highest priority).\n" +
      "  merchants          User-specific auto-learned per-name mappings.\n" +
      "  debit_credit_flips Debit/Credit flip decisions keyed by 'date|merchant|amount'.\n" +
      "  flagged            Flagged transactions keyed by 'date|merchant|amount'.\n" +
      "  custom_categories  User-defined categories with keyword lists.\n" +
      "  subscriptions      [{name, pinned, dismissed}]\n\n" +
      "Priority: overrides > merchants (user) > merchants (master) > category_keywords > custom_categories > Other"
    ),
    "merchants":          Object.assign({}, _liveMerchants),
    "overrides":          mergedOverrides,
    "debit_credit_flips": Object.assign({}, _liveFlips),
    "flagged":            Object.assign({}, _liveFlagged),
    "custom_categories":  customCategories,
    "subscriptions":      mergedSubscriptions,
  };

  // Sort merchants and overrides alphabetically for readability
  userOverridesJson.merchants = Object.fromEntries(
    Object.entries(userOverridesJson.merchants).sort(([a],[b]) => a.localeCompare(b))
  );
  userOverridesJson.overrides = Object.fromEntries(
    Object.entries(userOverridesJson.overrides).sort(([a],[b]) => a.localeCompare(b))
  );

  const json = JSON.stringify(userOverridesJson, null, 2);
  const blob = new Blob([json], {type: "application/json"});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = "user_overrides.json";
  a.click();
  URL.revokeObjectURL(url);

  // ── Update in-memory live state to match what was just saved ────────────
  // This means the badge immediately reflects "no unsaved changes" and any
  // subsequent edits are correctly tracked as new deltas against the saved file.
  _liveOverrides = Object.assign({}, mergedOverrides);
  _liveMerchants = Object.assign({}, userOverridesJson.merchants);
  _liveFlips     = Object.assign({}, userOverridesJson.debit_credit_flips);
  _liveFlagged   = Object.assign({}, userOverridesJson.flagged);
  _liveSubBase   = mergedSubscriptions;
  _subBase       = mergedSubscriptions;
  catOverrides   = {};
  _rebuildLiveByName();
  _invalidateRowCache();
  _origOverridesJson        = JSON.stringify(_liveOverrides);
  _origCustomCategoriesJson = JSON.stringify(customCategories);
  _origFlipsJson            = JSON.stringify(_liveFlips);
  _origFlaggedJson          = JSON.stringify(_liveFlagged);

  // Flash confirm on the button
  const btn = document.getElementById("update-merchants-btn");
  if (btn) {
    btn.textContent = "✓ Saved";
    setTimeout(() => {
      btn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M20 7H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="1"/></svg> Overrides <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg><span class="upd-badge"></span>`;
      _refreshUpdateBadge();
    }, 1600);
  }
}

function saveReport() {
  RAW.custom_categories = customCategories;  // ensure in sync
  // Bake current _liveFlips and _liveFlagged into merchant_json so the downloaded
  // report is self-contained — reopening it restores all decisions without needing
  // the user_overrides.json file to be present.
  const bakedMerchantJson = Object.assign({}, RAW.merchant_json, {
    overrides:          Object.assign({}, _liveOverrides),
    debit_credit_flips: Object.assign({}, _liveFlips),
    flagged:            Object.assign({}, _liveFlagged),
    custom_categories:  customCategories,
    subscriptions:      buildMergedSubscriptions(_liveSubBase),
  });
  const payload = {
    version:      3,
    saved_at:     new Date().toISOString(),
    raw:          Object.assign({}, RAW, { merchant_json: bakedMerchantJson }),
  };
  const json  = JSON.stringify(payload);
  const blob  = new Blob([json], {type: "application/json"});
  const url   = URL.createObjectURL(blob);
  const a     = document.createElement("a");
  const date  = new Date().toISOString().slice(0,10);
  a.href      = url;
  a.download  = `xpence_${date}.xpr`;
  a.click();
  URL.revokeObjectURL(url);
}

async function loadReport(input) {
  const file = input.files[0];
  if (!file) return;
  closeDrop("report-drop");
  const reader = new FileReader();
  reader.onload = async function(e) {
    try {
      let payload;
      const text = e.target.result;
      if (file.name.endsWith(".html") || file.name.endsWith(".htm")) {
        // Extract the embedded RAW data from the HTML report
        const marker = "const RAW = ";
        const markerIdx = text.indexOf(marker);
        if (markerIdx === -1) throw new Error("Could not find embedded report data in HTML file.");
        const jsonStart = markerIdx + marker.length;
        // Find the matching closing } by counting braces
        let depth = 0, i = jsonStart, inStr = false, esc = false;
        for (; i < text.length; i++) {
          const ch = text[i];
          if (esc) { esc = false; continue; }
          if (ch === "\\" && inStr) { esc = true; continue; }
          if (ch === '"') { inStr = !inStr; continue; }
          if (inStr) continue;
          if (ch === "{") depth++;
          else if (ch === "}") { depth--; if (depth === 0) { i++; break; } }
        }
        payload = { version: 3, raw: JSON.parse(text.slice(jsonStart, i)) };
      } else {
        payload = JSON.parse(text);
      }
      if (!payload.raw || !payload.raw.rows) {
        alert("Invalid xPence report file.");
        return;
      }

      // Load transaction data
      Object.assign(RAW, payload.raw);
      qidSet = new Set(RAW.questionable.map(q => q.qid));

      // Reset session-only state
      catOverrides = {};
      qReviewed    = false;
      customCategories = (RAW.custom_categories || []).map(cc => ({
        name: cc.name, keywords: [...(cc.keywords || [])]
      }));

      // Sync custom categories into RAW colour map
      RAW.custom_categories = customCategories;
      const customColor = RAW.custom_cat_color || "#3d3d3d";
      customCategories.forEach(cc => {
        RAW.cat_color_map[cc.name] = customColor;
        if (!RAW.all_categories.includes(cc.name)) RAW.all_categories.push(cc.name);
      });

      // ── Fetch live user_overrides.json — shared source of truth for all reports ──
      // Falls back to the overrides baked into this .xpr (which includes flips + flagged).
      let fetchedDataOnLoad = null;
      const onFileProtocol = location.protocol === "file:";
      try {
        const resp = await fetch("user_overrides.json", { cache: "no-store" });
        if (resp.ok) {
          const text = await resp.text();
          try { fetchedDataOnLoad = JSON.parse(text); } catch (_) {}
        }
      } catch (_) {}
      if (fetchedDataOnLoad) {
        _applyLiveOverrides(fetchedDataOnLoad, true, onFileProtocol);
      } else {
        // Baked snapshot includes flips + flagged saved at download time
        const baked = RAW.merchant_json || {};
        _applyLiveOverrides(baked, false, onFileProtocol);
      }

      _invalidateRowCache();

      // Reset filter/sort state
      activeCategoryFilter    = null;
      activeAllCategoryFilter = null;
      allSort   = {col: "date", dir: "desc"};
      monthSort = {col: "date", dir: "desc"};

      // Full re-render
      const s = RAW.summary;
      document.getElementById("meta-period").textContent = s.period;
      document.getElementById("s-spent").textContent = "$" + s.total_spent.toLocaleString("en-CA",{minimumFractionDigits:2});
      document.getElementById("s-cred").textContent  = "$" + s.total_cred .toLocaleString("en-CA",{minimumFractionDigits:2});
      document.getElementById("s-net").textContent   = "$" + Math.abs(s.net).toLocaleString("en-CA",{minimumFractionDigits:2});
      document.getElementById("s-tx").textContent    = s.tx_count;
      document.getElementById("s-gen").textContent   = s.generated;
      _setAvgMonth(getEffectiveRows());

      buildUncatSection();
      buildFlaggedReclassifyPanel();
      buildReviewTab();
      buildOverviewTab();
      buildSubscriptionsTab();
      buildCustomizePanel();

      const rows = getEffectiveRows();
      rebuildAllCatBars(rows);
      buildOverTimeChart(rows);
      buildAllTxTable(rows);

      const sel = document.getElementById("month-select");
      sel.innerHTML = "";
      RAW.months.forEach(m => {
        const o = document.createElement("option"); o.value = m; o.textContent = m; sel.appendChild(o);
      });
      sel.value = RAW.months[RAW.months.length - 1];
      renderMonthly(sel.value);

      _refreshReviewBadge();

      switchView("all");
      _refreshUpdateBadge();
      _checkUnsavedWarning();

    } catch(err) {
      alert("Could not load report: " + err.message);
    }
  };
  reader.readAsText(file);
  input.value = "";
}

// ── SUBSCRIPTIONS TAB ────────────────────────────────────────────────────
//
// Detection rules:
//   • Debits only
//   • Merchant must appear in MORE THAN 2 distinct calendar months (≥ 3)
//   • Must NOT appear more than twice in any single month
//   • All charges within ±5% of the median (tight consistency)
//   • Dismissed merchants (subDismissed) are excluded from auto-detection
//   • Pinned merchants (subPinned) are always included even if not auto-detected
//
function detectSubscriptions(rows) {
  const debitRows = rows.filter(r => r.debit > 0);

  const byMerchant = {};
  debitRows.forEach(r => {
    if (!byMerchant[r.name]) byMerchant[r.name] = [];
    byMerchant[r.name].push(r);
  });

  const detected = [];

  Object.entries(byMerchant).forEach(([name, txs]) => {
    // Skip dismissed merchants
    if (subDismissed.has(name)) return;

    const byMonth = {};
    txs.forEach(r => {
      if (!byMonth[r.month]) byMonth[r.month] = [];
      byMonth[r.month].push(r);
    });
    const months = Object.keys(byMonth).sort();

    // Must appear in MORE THAN 2 months (i.e. ≥ 3)
    if (months.length <= 2) return;

    // Months must be CONSECUTIVE — no gaps allowed.
    // Convert each "YYYY-MM" to a single integer (year*12 + month) and
    // check that each step is exactly 1.
    const monthNums = months.map(m => {
      const [y, mo] = m.split("-").map(Number);
      return y * 12 + mo;
    });
    const isConsecutive = monthNums.every((n, i) =>
      i === 0 || n === monthNums[i-1] + 1
    );
    if (!isConsecutive) return;

    // Must NOT appear more than twice in any single month
    if (months.some(m => byMonth[m].length > 2)) return;

    const amounts = txs.map(r => r.debit).sort((a, b) => a - b);
    const median  = amounts[Math.floor(amounts.length / 2)];

    // All charges within ±5% of the median
    if (!amounts.every(a => Math.abs(a - median) / median <= 0.05)) return;

    detected.push(_buildSubEntry(name, txs, months, amounts, median, false));
  });

  // Add manually pinned entries that weren't auto-detected
  subManual.forEach(m => {
    if (!detected.some(d => d.name === m.name)) {
      // Find actual transactions for this merchant in effective rows
      const txs = debitRows.filter(r => r.name === m.name);
      if (txs.length > 0) {
        const byMonth = {};
        txs.forEach(r => { if (!byMonth[r.month]) byMonth[r.month] = []; byMonth[r.month].push(r); });
        const months  = Object.keys(byMonth).sort();
        const amounts = txs.map(r => r.debit).sort((a, b) => a - b);
        const median  = m.monthlyAmt || amounts[Math.floor(amounts.length / 2)];
        detected.push(_buildSubEntry(m.name, txs, months, amounts, median, true));
      } else {
        // No transactions found — show as manually-added with stored amount
        detected.push({
          name:       m.name,
          category:   m.category || "—",
          monthlyAmt: m.monthlyAmt || 0,
          typicalDay: null,
          firstDate:  null,
          lastDate:   null,
          totalPaid:  0,
          monthCount: 0,
          txCount:    0,
          isPinned:   true,
          isManual:   true,
        });
      }
    }
  });

  detected.sort((a, b) => b.monthlyAmt - a.monthlyAmt);
  return detected;
}

function _buildSubEntry(name, txs, months, amounts, median, isPinned) {
  // Use date_iso for reliable day-of-month parsing
  const isoDates = txs.map(r => r.date_iso).filter(Boolean).sort();

  const days = isoDates
    .map(d => parseInt(d.slice(8), 10))   // characters 8-9 of "YYYY-MM-DD" = day
    .filter(d => !isNaN(d))
    .sort((a, b) => a - b);
  const typicalDay = days.length ? days[Math.floor(days.length / 2)] : null;

  return {
    name,
    category:   txs[0] ? txs[0].category : "—",
    monthlyAmt: median,
    typicalDay,
    firstDate:  isoDates[0]  || null,
    lastDate:   isoDates[isoDates.length - 1] || null,
    totalPaid:  amounts.reduce((s, a) => s + a, 0),
    monthCount: months.length,
    txCount:    txs.length,
    isPinned,
    isManual:   false,
  };
}

function subConfirm(name, monthlyAmt, category) {
  subPinned.add(name);
  subDismissed.delete(name);
  // Store in subManual so the amount/category survives in the JSON
  if (!subManual.some(m => m.name === name)) {
    subManual.push({ name, monthlyAmt, category });
  }
  _refreshUpdateBadge();
  buildSubscriptionsTab();
}

function subDismiss(name) {
  subDismissed.add(name);
  subPinned.delete(name);
  subManual = subManual.filter(m => m.name !== name);
  _refreshUpdateBadge();
  buildSubscriptionsTab();
}

function subAddManual() {
  const nameEl = document.getElementById("sub-manual-name");
  const amtEl  = document.getElementById("sub-manual-amt");
  const catEl  = document.getElementById("sub-manual-cat");
  const name   = (nameEl.value || "").trim();
  const amt    = parseFloat(amtEl.value) || 0;
  const cat    = catEl.value || "";

  if (!name) { nameEl.style.borderColor = "var(--red)"; nameEl.focus(); return; }
  nameEl.style.borderColor = "";

  if (subManual.some(m => m.name === name)) {
    alert(`"${name}" is already in your subscription list.`); return;
  }

  subManual.push({ name, monthlyAmt: amt, category: cat });
  subPinned.add(name);
  subDismissed.delete(name);

  nameEl.value = "";
  amtEl.value  = "";
  _refreshUpdateBadge();
  buildSubscriptionsTab();
}

function buildSubscriptionsTab() {
  const el = document.getElementById("sub-content");
  if (!el) return;

  const rows = getEffectiveRows();
  const subs = detectSubscriptions(rows);

  // Build a set of merchant names that actually have transactions in this dataset
  const namesInData = new Set(rows.filter(r => r.debit > 0).map(r => r.name));

  // Update badge on tab
  const badge = document.getElementById("sub-badge");
  if (badge) {
    badge.textContent   = subs.length;
    badge.style.display = subs.length > 0 ? "inline-flex" : "none";
  }

  const fmt = v => "$" + Number(v).toLocaleString("en-CA",
    {minimumFractionDigits:2, maximumFractionDigits:2});

  // Safe ISO date formatter — expects "YYYY-MM-DD"
  const fmtDate = d => {
    if (!d || typeof d !== "string" || d.length < 10) return "—";
    const mo  = parseInt(d.slice(5,7),  10);
    const day = parseInt(d.slice(8,10), 10);
    const yr  = d.slice(0,4);
    if (isNaN(mo) || isNaN(day)) return "—";
    const mn = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    return `${mn[mo-1]} ${day}, ${yr}`;
  };

  const ordinal = n => {
    if (!n || isNaN(n)) return "—";
    const s = n % 100;
    if (s >= 11 && s <= 13) return `${n}th`;
    switch (n % 10) {
      case 1: return `${n}st`;
      case 2: return `${n}nd`;
      case 3: return `${n}rd`;
      default: return `${n}th`;
    }
  };

  const duration = sub => {
    if (!sub.firstDate || !sub.lastDate)
      return sub.monthCount > 0 ? `${sub.monthCount} mo` : "—";
    const fy = parseInt(sub.firstDate.slice(0,4), 10);
    const fm = parseInt(sub.firstDate.slice(5,7), 10);
    const ly = parseInt(sub.lastDate.slice(0,4),  10);
    const lm = parseInt(sub.lastDate.slice(5,7),  10);
    if (isNaN(fy+fm+ly+lm)) return "—";
    const n = (ly - fy) * 12 + (lm - fm) + 1;
    return n === 1 ? "1 month" : `${n} months`;
  };

  // ── Hero stat cards ───────────────────────────────────────────────────────
  // Sum of each subscription's median monthly charge = what you pay per month
  const totalMonthly = subs.reduce((s, sub) => s + (sub.monthlyAmt || 0), 0);

  // Average monthly expenditure computed from effective rows
  const monthCount  = (RAW.months || []).length || 1;
  const totalDebit  = rows.reduce((s, r) => s + (r.debit || 0), 0);
  const avgMonthly  = totalDebit / monthCount;

  // Spent on subscriptions in the most recent (last/current) month
  // = sum of all debit rows whose merchant name matches a detected subscription
  // and whose month == the latest month in the data.
  const latestMonth  = (RAW.months || []).slice(-1)[0] || "";
  const subNames     = new Set(subs.map(s => s.name));
  const lastMonthSubSpend = rows
    .filter(r => r.debit > 0 && r.month === latestMonth && subNames.has(r.name))
    .reduce((s, r) => s + r.debit, 0);
  const lastMonthLabel = (() => {
    if (!latestMonth || latestMonth.length < 7) return "Last month";
    const [y, m] = latestMonth.split("-");
    const mn = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    return mn[parseInt(m,10)-1] + " " + y;
  })();

  // % = subscription cost per month ÷ average monthly expenditure × 100
  const pct = avgMonthly > 0
    ? (totalMonthly / avgMonthly * 100).toFixed(1)
    : "—";

  const hero = `
  <div class="sub-hero">
    <div class="sub-hero-card">
      <div class="sub-hero-label">Subscriptions detected</div>
      <div class="sub-hero-value">${subs.length}</div>
      <div class="sub-hero-sub">${fmt(totalMonthly)}&thinsp;/&thinsp;month · ${fmt(totalMonthly * 12)}&thinsp;/&thinsp;year</div>
    </div>
    <div class="sub-hero-card">
      <div class="sub-hero-label">% of monthly expenses</div>
      <div class="sub-hero-value sub-hero-pct">${pct}%</div>
      <div class="sub-hero-sub">${fmt(totalMonthly)}&thinsp;/&thinsp;mo subscriptions ÷ ${fmt(avgMonthly)}&thinsp;/&thinsp;mo avg spend</div>
    </div>
    <div class="sub-hero-card">
      <div class="sub-hero-label">Spent on subs — ${lastMonthLabel}</div>
      <div class="sub-hero-value">${fmt(lastMonthSubSpend)}</div>
      <div class="sub-hero-sub">${lastMonthSubSpend > 0 ? "actual charges from subscription merchants" : "no subscription charges found this month"}</div>
    </div>
  </div>`;

  const note = `
  <div class="sub-note">
    <strong>Detection:</strong> consistent amount (±5%), 1–2 times per month, across
    <strong>more than 2 months</strong>. Merchants appearing 3+ times/month are excluded.
    <strong>Confirm</strong> pins a subscription so it persists across reports even when the
    dataset changes. <strong>Delete</strong> hides it permanently.
    Save changes via <strong>Save Overrides</strong> in the header.
  </div>`;

  // ── List ──────────────────────────────────────────────────────────────────
  let listHtml = "";
  if (subs.length === 0) {
    listHtml = `
    <div class="sub-empty">
      <div class="sub-empty-icon">📭</div>
      <div class="sub-empty-msg">No subscriptions detected yet</div>
      <div class="sub-empty-hint">
        Subscriptions need a consistent charge (±5%) once or twice per month
        across more than 2 months. You can also add one manually below.
      </div>
    </div>`;
  } else {
    const rows_html = subs.map((sub, i) => {
      const isPinned  = subPinned.has(sub.name);
      const isManual  = sub.isManual;
      const inData    = namesInData.has(sub.name);

      const cat   = sub.category && sub.category !== "—" ? sub.category : "";
      const color = RAW.cat_color_map[sub.category] || "#64748b";
      const catPip = cat
        ? `<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${color};margin-right:3px;vertical-align:middle"></span>`
        : "";

      const dayStr = sub.typicalDay
        ? `~${ordinal(sub.typicalDay)}`
        : "Varies";

      // Button logic:
      // • If pinned AND not in current dataset → Delete only (no Confirm)
      // • If pinned AND in current dataset → Delete only (already confirmed)
      // • If NOT pinned AND in current dataset → Confirm + Delete
      // • If NOT pinned AND not in current dataset → shouldn't appear (detectSubscriptions handles this)
      let actionBtns = "";
      if (!isPinned && inData) {
        actionBtns =
          `<button class="sub-btn sub-btn-confirm" data-action="confirm" data-idx="${i}">Confirm</button>` +
          `<button class="sub-btn sub-btn-delete"  data-action="delete"  data-idx="${i}">Delete</button>`;
      } else {
        // Pinned or manual — only Delete
        actionBtns =
          `<button class="sub-btn sub-btn-delete" data-action="delete" data-idx="${i}">Delete</button>`;
      }

      return `
      <div class="sub-list-row" data-idx="${i}">
        <div class="sub-row-name">
          <div class="sub-row-name-info">
            <div class="sub-row-name-text">${sub.name}</div>
            <div class="sub-row-cat">${catPip}${cat}</div>
          </div>
        </div>
        <div class="sub-row-amt">${fmt(sub.monthlyAmt)}<br><span class="sub-row-muted" style="font-size:.7rem;font-weight:400">/month</span></div>
        <div class="sub-col-day sub-row-meta">${dayStr}</div>
        <div class="sub-col-first sub-row-muted">${fmtDate(sub.firstDate)}</div>
        <div class="sub-row-muted">${fmtDate(sub.lastDate)}</div>
        <div class="sub-col-total sub-row-meta">${fmt(sub.totalPaid)}</div>
        <div class="sub-row-actions">${actionBtns}</div>
      </div>`;
    }).join("");

    listHtml = `
    <div class="sub-list-wrap" id="sub-list-container">
      <div class="sub-list-head">
        <div>Merchant</div>
        <div class="sub-col-amt-head">Monthly</div>
        <div class="sub-col-day">Charge date</div>
        <div class="sub-col-first">First charged</div>
        <div>Last charged</div>
        <div class="sub-col-total">Total paid</div>
        <div></div>
      </div>
      ${rows_html}
    </div>`;
  }

  // ── Manual add form ───────────────────────────────────────────────────────
  const catOptions = RAW.all_categories.map(c =>
    `<option value="${c}">${c}</option>`).join("");
  const manualForm = `
  <div class="sub-manual-section">
    <div class="sub-manual-title">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
        <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
      </svg>
      Add Subscription Manually
    </div>
    <div class="sub-manual-row">
      <div class="sub-manual-field" style="flex:2">
        <label for="sub-manual-name">Merchant name</label>
        <input type="text" id="sub-manual-name" placeholder="e.g. Netflix"/>
      </div>
      <div class="sub-manual-field" style="flex:1">
        <label for="sub-manual-amt">Monthly amount ($)</label>
        <input type="number" id="sub-manual-amt" placeholder="17.99" min="0" step="0.01"/>
      </div>
      <div class="sub-manual-field" style="flex:1.5">
        <label for="sub-manual-cat">Category</label>
        <select id="sub-manual-cat">
          <option value="">— select —</option>${catOptions}
        </select>
      </div>
      <button class="sub-manual-add-btn" onclick="subAddManual()">Add</button>
    </div>
  </div>`;

  el.innerHTML = hero + note + listHtml + manualForm;

  // ── Event delegation — avoids inline onclick with escaped strings ──────────
  // Store the current subs array on the element so the handler can look up entries
  el._subs = subs;
  const container = document.getElementById("sub-list-container");
  if (container) {
    container.addEventListener("click", function(e) {
      const btn = e.target.closest("[data-action]");
      if (!btn) return;
      const idx    = parseInt(btn.dataset.idx, 10);
      const action = btn.dataset.action;
      const sub    = el._subs[idx];
      if (!sub) return;
      if (action === "confirm") {
        subConfirm(sub.name, sub.monthlyAmt, sub.category);
      } else if (action === "delete") {
        subDismiss(sub.name);
      }
    });
  }
}

// ── Cursor-following tooltip for monthly pivot table headers ────────────────
(function(){
  const tip = document.createElement('div');
  tip.id = 'xp-cursor-tip';
  document.body.appendChild(tip);

  document.addEventListener('mousemove', function(e){
    if (tip.style.opacity === '1') {
      tip.style.left = (e.clientX + 14) + 'px';
      tip.style.top  = (e.clientY + 14) + 'px';
    }
  });

  document.addEventListener('mouseover', function(e){
    const th = e.target.closest('th[data-tip]');
    if (th && th.closest('.pivot-wrap')) {
      tip.textContent = th.dataset.tip;
      tip.style.left  = (e.clientX + 14) + 'px';
      tip.style.top   = (e.clientY + 14) + 'px';
      tip.style.opacity = '1';
    }
  });

  document.addEventListener('mouseout', function(e){
    const th = e.target.closest('th[data-tip]');
    if (th && th.closest('.pivot-wrap')) {
      tip.style.opacity = '0';
    }
  });
})();

</script>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def _update_user_overrides_from_report(
    df:       pd.DataFrame,
    col:      dict,
    path:     str,
    user_map: dict,
) -> None:
    """
    Merge newly-seen merchants (keyword-classified) into the 'merchants'
    section of user_overrides.json.

    Only the user file is written — the master merchant_categories.json is
    never modified by a local run.

    Also purges any 'overrides' entries whose target category no longer exists
    (e.g. a custom category that was deleted via the Customize panel and then
    re-saved via 'Save Overrides').
    """
    df2 = df.copy()
    df2["_name"]  = df2[col["name"]].astype(str).str.strip()
    df2["_debit"] = clean_amount(df2[col["debit"]]).clip(lower=0) if col.get("debit") else 0.0
    # Classify using keywords only (no map) to get the baseline category per merchant
    _vc = set(CATEGORIES) | {"Other / Uncategorised"}
    df2["_category"] = df2["_name"].apply(lambda n: classify(n, None, None, _vc))
    new_entries: dict = (
        df2[df2["_debit"] > 0]
        .groupby("_name")["_category"]
        .agg(lambda s: s.mode().iat[0])
        .to_dict()
    )
    # Existing user merchants win over newly-derived entries.
    # Never store "Other / Uncategorised" — it permanently blocks keyword re-matching.
    new_entries_filtered = {
        k: v for k, v in new_entries.items()
        if v != "Other / Uncategorised"
    }
    user_map["merchants"] = {**new_entries_filtered, **user_map["merchants"]}

    # ── Purge stale overrides pointing at categories that no longer exist ────
    valid_cats = (
        set(CATEGORIES)
        | {"Other / Uncategorised"}
        | {cc["name"] for cc in user_map.get("custom_categories", [])}
    )
    stale = [k for k, v in user_map["overrides"].items() if v not in valid_cats]
    stale_summary = {k: user_map["overrides"][k] for k in stale[:5]}
    for k in stale:
        del user_map["overrides"][k]
    if stale:
        examples = ", ".join(f'"{k}"->"{v}"' for k, v in stale_summary.items())
        print(f"  Purged {len(stale)} stale override(s) pointing at deleted categories: "
              + examples + (" …" if len(stale) > 5 else ""))

    save_user_overrides(path, user_map)
    total = len(user_map["merchants"]) + len(user_map["overrides"])
    print(f"  User overrides -> {path}  "
          f"({len(user_map['merchants'])} user merchants, "
          f"{len(user_map['overrides'])} overrides"
          + (f", {len(stale)} stale purged" if stale else "")
          + f",  {total} total)")


def main():
    parser = argparse.ArgumentParser(description="xPence — monthly expense report generator")
    parser.add_argument("--file",      "-f", help="CSV/Excel path (default: auto-discover xPence*)")
    parser.add_argument("--output",    "-o", help="Output HTML path (default: xpence_report.html)")
    parser.add_argument("--csv",       "-c", help="Also write a category summary CSV")
    parser.add_argument("--merchants", "-m", help=(
        "Path to master merchant_categories.json "
        "(default: merchant_categories.json beside this script). "
        "Downloaded from the internet; contains category_keywords and community merchants."))
    parser.add_argument("--overrides", "-u", help=(
        "Path to user_overrides.json "
        "(default: user_overrides.json beside this script). "
        "Contains your personal overrides, custom categories, and subscriptions. "
        "Created automatically on first run."))
    parser.add_argument("--account-type", "-a", help=(
        "Only include transactions on this account type in the report "
        "(e.g. Credit, Chequing, Savings). Only takes effect if the source "
        "data has an identifiable account-type column — see README."))
    parser.add_argument("--xpr",       "-x", help=(
        "Path to a previously-saved .xpr report. Category overrides inside "
        "the xpr are pushed into user_overrides.json before generating the "
        "new report, so reclassifications carry forward."))
    args = parser.parse_args()

    _here = os.path.dirname(os.path.abspath(__file__))

    # ── Resolve file paths ────────────────────────────────────────────────
    master_path    = args.merchants or os.path.join(_here, "merchant_categories.json")
    overrides_path = args.overrides or os.path.join(_here, "user_overrides.json")

    # ── Step 1: load master reference list (keywords + community merchants) ──
    master_map = _load_master(master_path)

    # ── Step 2: ingest any xpr overrides into user_overrides BEFORE classifying ──
    if args.xpr:
        if not os.path.exists(args.xpr):
            print(f"ERROR: xpr file not found: {args.xpr}")
            sys.exit(1)
        print(f"Reading overrides from: {args.xpr}")
        update_merchant_json_from_xpr(args.xpr, overrides_path)

    # ── Step 3: load user overrides (reflects any just-pushed xpr data) ──
    user_map = load_user_overrides(overrides_path)
    n_o = len(user_map["overrides"])
    n_u = len(user_map["merchants"])
    if n_o or n_u:
        print(f"  User overrides loaded: {n_u} user merchants, {n_o} overrides")
    else:
        print("  No user overrides found — using master list + keywords only")

    # ── Step 4: find and load CSV/Excel ──────────────────────────────────
    filepath = args.file
    if not filepath:
        filepath = find_xpence_file(".")
        if not filepath:
            for d in ["data","downloads","Downloads","input"]:
                filepath = find_xpence_file(d)
                if filepath: break
    if not filepath:
        print("ERROR: Could not find an 'xPence' CSV file.")
        print("       Use: --file path/to/xPence.csv")
        sys.exit(1)

    print(f"Loading: {filepath}")
    df, col = load_csv(filepath)
    print(f"  {len(df)} rows x {len(df.columns)} columns")
    print(f"  Column mapping: { {k:v for k,v in col.items()} }")

    # ── Step 5: build report (classification uses both maps) ─────────────
    print("Classifying transactions...")
    html = build_report(df, col, user_map=user_map, master_map=master_map,
                        overrides_path=overrides_path,
                        account_type_filter=args.account_type)

    out_path = args.output or "xpence_report.html"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Report written -> {out_path}")

    # ── Step 6: merge new user-merchants into user_overrides.json ────────
    _update_user_overrides_from_report(df, col, overrides_path, user_map)

    # ── Step 7: optional CSV summary ─────────────────────────────────────
    if args.csv:
        df2 = df.copy()
        df2["_date"]  = pd.to_datetime(df2[col["date"]], errors="coerce")
        df2["_name"]  = df2[col["name"]].astype(str).str.strip()
        df2["_debit"] = clean_amount(df2[col["debit"]]).clip(lower=0) if col.get("debit") else 0.0
        df2 = df2.dropna(subset=["_date"])
        df2["_category"] = df2["_name"].apply(
            lambda n: classify(n, user_map, master_map)
        )
        df2["_month"] = df2["_date"].dt.to_period("M").astype(str)
        summary = (df2.groupby(["_month","_category"])["_debit"].sum().reset_index()
                      .rename(columns={"_month":"Month","_category":"Category","_debit":"Total Spent"}))
        summary.to_csv(args.csv, index=False)
        print(f"Summary CSV -> {args.csv}")

    print("Done! Open the HTML report in your browser.")


if __name__ == "__main__":
    main()
