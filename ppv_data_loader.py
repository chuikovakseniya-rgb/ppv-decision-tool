"""
Load and merge PPV export files (spending, active listers, price per day) by category_id.
Does not depend on decision_engine or Streamlit UI.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Dict, Mapping, Optional, Tuple, Union

import pandas as pd

logger = logging.getLogger(__name__)

PathLike = Union[str, Path]

# spending measure label (normalized) -> snake_case key in merged before/after buckets
SPENDING_MEASURE_ALIASES: Mapping[str, str] = {
    "paid users": "paid_users",
    "new paid listers": "paid_users",
    "new paid lister": "paid_users",
    "paid listers": "paid_users",
    "paid lister": "paid_users",
    "npl": "paid_users",
    "spending": "spending",
    "spendings": "spending",
    "campaign per user": "campaign_per_user",
    "new campaign cnt": "new_campaign_cnt",
    "new campaign count": "new_campaign_cnt",
    "price per day": "price_per_day",
    "arp pcampaign": "arp_p_campaign",
    "arp ppcampaign": "arp_p_campaign",
    "arp pcampaing": "arp_p_campaign",  # typo in some exports
    "refund": "refund",
    "plan imp per campaign": "plan_imp_per_campaign",
    "fact imp per campaign": "fact_imp_per_campaign",
    "%campaign with refund": "pct_campaign_with_refund",
    "campaign with refund": "pct_campaign_with_refund",
    "%execution inventory": "pct_execution_inventory",
    "execution inventory": "pct_execution_inventory",
    "arppcampaing": "arp_p_campaign",
    "arppcampaign": "arp_p_campaign",
}

_CSV_ENCODINGS = (
    "utf-8",
    "utf-8-sig",
    "utf-16",
    "utf-16-le",
    "utf-16-be",
    "cp1251",
    "cp1252",
    "latin-1",
)


def _snake_case(name: str) -> str:
    s = str(name).strip().lower()
    s = re.sub(r"[^\w\s]", " ", s, flags=re.UNICODE)
    s = re.sub(r"\s+", "_", s)
    return s.strip("_")


def _normalize_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Return df with snake_case columns and mapping snake -> original."""
    mapping: Dict[str, str] = {}
    new_cols = []
    for c in df.columns:
        sc = _snake_case(c)
        if sc in mapping:
            sc = f"{sc}_{len([k for k in new_cols if k.startswith(sc)])}"
        mapping[sc] = str(c)
        new_cols.append(sc)
    out = df.copy()
    out.columns = new_cols
    return out, mapping


def _peek_file_start(path: Path, n: int = 8) -> bytes:
    with open(path, "rb") as f:
        return f.read(n)


def _read_table(path: PathLike) -> pd.DataFrame:
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(path)
    suf = path.suffix.lower()
    head = _peek_file_start(path, 8)

    def _try_excel() -> pd.DataFrame:
        return pd.read_excel(path)

    # .xlsx is ZIP; .xls often starts with OLE header (even if extension is wrong)
    if head[:2] == b"PK":
        return _try_excel()
    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return _try_excel()

    if suf in (".xlsx", ".xls"):
        try:
            return _try_excel()
        except Exception:
            pass

    # CSV / text exports: Excel often saves "CSV UTF-16" → BOM 0xFF 0xFE (decode fails as utf-8)
    last_err: Optional[Exception] = None
    for enc in _CSV_ENCODINGS:
        try:
            return pd.read_csv(path, encoding=enc)
        except UnicodeDecodeError as e:
            last_err = e
            continue
        except Exception as e:
            last_err = e
            continue

    try:
        return _try_excel()
    except Exception as e:
        last_err = e

    raise ValueError(f"Could not read {path.name!r}: {last_err}") from last_err


def _read_table_no_header(path: PathLike) -> pd.DataFrame:
    """
    Read sheet/CSV without treating row 0 as column names.
    Use when Excel has multi-row headers so pandas' default header=0 breaks alignment.
    """
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(path)
    suf = path.suffix.lower()
    if suf in (".xlsx", ".xls"):
        return pd.read_excel(path, header=None)
    if suf == ".csv":
        last_err: Optional[Exception] = None
        for enc in _CSV_ENCODINGS:
            try:
                return pd.read_csv(path, encoding=enc, header=None)
            except UnicodeDecodeError as e:
                last_err = e
                continue
            except Exception as e:
                last_err = e
                continue
        raise ValueError(f"Could not read CSV without header: {path.name!r}: {last_err}") from last_err
    raise ValueError(f"Unsupported file type for no-header read: {suf}")


def _retry_if_missing_category_id(path: Path, df: pd.DataFrame, parse_fn):
    """Re-parse from raw grid if category_id is missing (multi-row Excel headers)."""
    path = Path(path)
    try:
        return parse_fn(df)
    except ValueError as e:
        if "category_id" not in str(e).lower():
            raise
        df2 = _read_table_no_header(path)
        return parse_fn(df2)


def _parse_number(val: Any) -> Optional[float]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return float(val)
    s = str(val).strip().replace("\u00a0", " ")
    if not s or s.lower() == "nan":
        return None
    if s.endswith("%"):
        s = s[:-1].strip()
    # European: "2 959" thousands, "11,93" decimal
    if re.match(r"^-?[\d\s.,]+%?$", s.replace("-", ""), re.UNICODE):
        if "," in s and "." in s:
            s = s.replace(" ", "").replace(".", "").replace(",", ".")
        elif "," in s and re.search(r",\d{1,2}$", s):
            s = s.replace(" ", "").replace(",", ".")
        else:
            s = s.replace(" ", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def pct_change_relative(new: Optional[Any], old: Optional[Any]) -> float:
    """
    Relative change in percent: (new − old) / old × 100.

    Used in the UI summary table as:

    - **beforeDiff** («До %»): ``new`` = absolute «До», ``old`` = baseline (опорный период из выгрузки).
    - **afterDiff** («После %»): ``new`` = «После», ``old`` = «До».

    Returns ``float('nan')`` if the ratio is undefined (e.g. ``old`` is 0 or missing).
    """
    if new is None or old is None:
        return float("nan")
    try:
        o = float(old)
        n = float(new)
    except (TypeError, ValueError):
        return float("nan")
    if o == 0.0:
        return float("nan")
    return (n - o) / o * 100.0


def _norm_measure_token(text: str) -> str:
    t = text.lower().strip()
    t = re.sub(r"\s+", " ", t)
    t = t.replace("ppv (spending)total ", "").replace("ppv (spending)", "")
    t = re.sub(r"^total\s+", "", t)
    return t.strip()


def _map_spending_measure(measure_raw: str) -> Optional[str]:
    if not measure_raw or str(measure_raw).lower() in ("nan", "none"):
        return None
    raw = str(measure_raw).strip().lower().replace("_", " ")
    key = _norm_measure_token(raw)
    key = re.sub(r"^\d+\s*", "", key)  # drop leading id fragments if glued
    if raw in SPENDING_MEASURE_ALIASES:
        return SPENDING_MEASURE_ALIASES[raw]
    if key in SPENDING_MEASURE_ALIASES:
        return SPENDING_MEASURE_ALIASES[key]
    for alias, snake in SPENDING_MEASURE_ALIASES.items():
        if alias in key or key in alias:
            return snake
    if re.search(r"new\s+paid\s+lister", key):
        return "paid_users"
    if re.search(r"\bnpl\b", key):
        return "paid_users"
    if "spending" in key:
        return "spending"
    if "fact imp" in key and "campaign" in key:
        return "fact_imp_per_campaign"
    if "plan imp" in key and "campaign" in key:
        return "plan_imp_per_campaign"
    if "campaign with refund" in key or key.startswith("%campaign with refund"):
        return "pct_campaign_with_refund"
    if "execution inventory" in key or key.startswith("%execution inventory"):
        return "pct_execution_inventory"
    return None


def _infer_measure_column_from_values(
    df: pd.DataFrame,
    id_col: str,
    default_col: str,
    target_col: str,
) -> Optional[str]:
    """
    Long-format PPV exports often put metric labels (Paid users, Spending, …) in a column
    with an empty / nonstandard header (becomes unnamed_* after normalization). Pick that
    column by how many cells map to known spending measures.
    """
    skip = {id_col, default_col, target_col}
    best_col: Optional[str] = None
    best_mapped = 0
    for c in df.columns:
        if c in skip:
            continue
        sample = df[c].dropna()
        if len(sample) == 0:
            continue
        sample = sample.head(400)
        n = len(sample)
        mapped = 0
        numeric_like = 0
        for v in sample:
            s = str(v).strip()
            if _map_spending_measure(s):
                mapped += 1
            if _parse_number(v) is not None:
                numeric_like += 1
        if mapped == 0:
            continue
        if numeric_like / n > 0.65 and mapped / n < 0.35:
            continue
        if mapped > best_mapped:
            best_mapped = mapped
            best_col = str(c)
    if best_col is None:
        return None
    if best_mapped < 2 and len(df) > 3:
        return None
    return best_col


def _find_column(columns: Any, *candidates: str) -> Optional[str]:
    colset = set(columns)
    for c in candidates:
        if c in colset:
            return c
    return None


def _category_id_column(df_norm: pd.DataFrame) -> str:
    for cand in ("category_id", "categoryid", "id"):
        if cand in df_norm.columns:
            return cand
    for c in df_norm.columns:
        if c == "category_id" or c.endswith("category_id"):
            return c
    raise ValueError("Could not find category_id column after normalization")


def _coerce_category_id(val: Any) -> Optional[int]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return int(val)
    s = str(val).strip()
    m = re.search(r"(\d{3,})", s)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            pass
    try:
        return int(float(s.replace(",", ".")))
    except ValueError:
        return None


def _guess_category_id_column_from_values(df_norm: pd.DataFrame) -> Optional[str]:
    """
    When headers are merged/odd, pick the column whose values look like category_id integers.
    """
    if df_norm.empty:
        return None

    skip_name_substr = (
        "spending",
        "paid_user",
        "active_lister",
        "measure",
        "metric",
        "default",
        "target",
        "before",
        "after",
        "period_group",
        "price_per",
        "catalog_target",
        "catalog_default",
    )

    best_col: Optional[str] = None
    best_ids = -1
    best_ratio = 0.0

    for c in df_norm.columns:
        cl = str(c).lower()
        if any(s in cl for s in skip_name_substr):
            continue
        col = df_norm[c]
        sample = col.dropna().head(400)
        if len(sample) < 2:
            continue
        n = len(sample)
        ids = 0
        tiny = 0
        for v in sample:
            if _coerce_category_id(v) is not None:
                ids += 1
            try:
                fv = float(v)
                if fv != int(fv) and 0 < fv < 1:
                    tiny += 1
            except (TypeError, ValueError):
                pass
        if tiny / n > 0.35:
            continue
        ratio = ids / n
        if ids < 2 or ratio < 0.2:
            continue
        if ids > best_ids or (ids == best_ids and ratio > best_ratio):
            best_ids = ids
            best_ratio = ratio
            best_col = str(c)

    if best_col is not None:
        return best_col

    if len(df_norm.columns) > 0:
        c0 = df_norm.columns[0]
        sample = df_norm[c0].dropna().head(400)
        if len(sample) >= 2:
            ids = sum(1 for v in sample if _coerce_category_id(v) is not None)
            if ids >= max(2, int(0.15 * len(sample))):
                return str(c0)
    return None


def _resolve_category_id_column(df: pd.DataFrame) -> str:
    try:
        return _category_id_column(df)
    except ValueError as exc:
        guess = _guess_category_id_column_from_values(df)
        if guess is None:
            raise ValueError(
                "Could not find category_id column after normalization"
            ) from exc
        logger.info("Using inferred category_id column %r (header detection fallback)", guess)
        return guess


def _unpack_active_triplet(
    t: Tuple[Optional[float], ...],
) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    """Active export: (default, target) or (baseline, default, target)."""
    if t is None or len(t) == 0:
        return (None, None, None)
    if len(t) == 2:
        return (None, t[0], t[1])
    return (t[0], t[1], t[2])


def _parse_spending_dataframe(df_raw: pd.DataFrame) -> Dict[int, Dict[str, Dict[str, Optional[float]]]]:
    df_raw = _maybe_promote_inline_header_row_spending_active(df_raw)
    df, _ = _normalize_columns(df_raw)
    id_col = _resolve_category_id_column(df)

    # Long format: measure + default + target
    measure_col = _find_column(
        df.columns,
        "measure",
        "measure_name",
        "metric",
        "measure_group",
        "ppv_spending_total",
    )
    if measure_col is None:
        for c in df.columns:
            if "measure" in c:
                measure_col = c
                break

    default_col = _find_column(
        df.columns,
        "default",
        "catalog_default",
        "before",
    )
    target_col = _find_column(
        df.columns,
        "target",
        "after",
        "catalog_target",
    )
    baseline_col = _find_column(
        df.columns,
        "baseline",
        "catalog_baseline",
        "reference",
    )

    if measure_col is None and default_col and target_col:
        measure_col = _infer_measure_column_from_values(
            df, id_col, default_col, target_col
        )

    out: Dict[int, Dict[str, Dict[str, Optional[float]]]] = {}

    def ensure(cid: int) -> Dict[str, Dict[str, Optional[float]]]:
        if cid not in out:
            emp = {k: None for k in SPENDING_MEASURE_ALIASES.values()}
            out[cid] = {"baseline": dict(emp), "before": dict(emp), "after": dict(emp)}
        return out[cid]

    if measure_col and default_col and target_col:
        for _, row in df.iterrows():
            cid = _coerce_category_id(row[id_col])
            if cid is None:
                continue
            mkey = _map_spending_measure(str(row.get(measure_col, "")))
            if not mkey:
                continue
            bucket = ensure(cid)
            bucket["before"][mkey] = _parse_number(row[default_col])
            bucket["after"][mkey] = _parse_number(row[target_col])
            if baseline_col:
                bucket["baseline"][mkey] = _parse_number(row[baseline_col])
        return out

    # Wide format: suffix _default/_before and _target/_after
    wide_before: Dict[int, Dict[str, Optional[float]]] = {}
    wide_after: Dict[int, Dict[str, Optional[float]]] = {}
    wide_baseline: Dict[int, Dict[str, Optional[float]]] = {}

    # Longer suffixes first so e.g. *_catalog_default is not parsed as *_default.
    suffix_map = (
        ("_catalog_default", "_catalog_target"),
        ("_default", "_target"),
        ("_before", "_after"),
    )
    baseline_suffixes = ("_catalog_baseline", "_baseline", "_reference")

    for _, row in df.iterrows():
        cid = _coerce_category_id(row[id_col])
        if cid is None:
            continue
        wb = wide_before.setdefault(cid, {})
        wa = wide_after.setdefault(cid, {})
        wbl = wide_baseline.setdefault(cid, {})
        for col in df.columns:
            if col == id_col:
                continue
            hit_bl = False
            for bs in baseline_suffixes:
                if col.endswith(bs):
                    base = col[: -len(bs)]
                    mkey = _map_spending_measure(base.replace("_", " "))
                    if mkey:
                        wbl[mkey] = _parse_number(row[col])
                    hit_bl = True
                    break
            if hit_bl:
                continue
            base = None
            side = None
            for sb, sa in suffix_map:
                if col.endswith(sb):
                    base = col[: -len(sb)]
                    side = "before"
                    break
                if col.endswith(sa):
                    base = col[: -len(sa)]
                    side = "after"
                    break
            if base is None or side is None:
                continue
            mkey = _map_spending_measure(base.replace("_", " "))
            if not mkey:
                continue
            val = _parse_number(row[col])
            if side == "before":
                wb[mkey] = val
            else:
                wa[mkey] = val

    for cid in set(wide_before) | set(wide_after) | set(wide_baseline):
        bucket = ensure(cid)
        for mk in SPENDING_MEASURE_ALIASES.values():
            bucket["before"][mk] = wide_before.get(cid, {}).get(mk)
            bucket["after"][mk] = wide_after.get(cid, {}).get(mk)
            bucket["baseline"][mk] = wide_baseline.get(cid, {}).get(mk)

    return out


def _parse_active_dataframe(
    df_raw: pd.DataFrame,
) -> Dict[int, Tuple[Optional[float], Optional[float], Optional[float]]]:
    df_raw = _maybe_promote_inline_header_row_spending_active(df_raw)
    df, _ = _normalize_columns(df_raw)
    id_col = _resolve_category_id_column(df)

    default_col = _find_column(
        df.columns,
        "default",
        "catalog_default",
        "before",
    )
    target_col = _find_column(
        df.columns,
        "target",
        "after",
        "catalog_target",
    )
    baseline_col = _find_column(
        df.columns,
        "baseline",
        "catalog_baseline",
        "reference",
    )

    out: Dict[int, Tuple[Optional[float], Optional[float], Optional[float]]] = {}

    if default_col and target_col:
        for _, row in df.iterrows():
            cid = _coerce_category_id(row[id_col])
            if cid is None:
                continue
            bl = _parse_number(row[baseline_col]) if baseline_col else None
            out[cid] = (
                bl,
                _parse_number(row[default_col]),
                _parse_number(row[target_col]),
            )
        return out

    # Try columns active_listers_* wide layout
    for _, row in df.iterrows():
        cid = _coerce_category_id(row[id_col])
        if cid is None:
            continue
        bl: Optional[float] = None
        b: Optional[float] = None
        a: Optional[float] = None
        for col in df.columns:
            if col == id_col:
                continue
            lc = col.lower()
            if "active" not in lc:
                continue
            if col.endswith("_catalog_baseline") or col.endswith("_baseline") or col.endswith("_reference"):
                bl = _parse_number(row[col])
                continue
            if col.endswith("_catalog_default") or col.endswith("_default") or col.endswith("_before"):
                b = _parse_number(row[col])
            elif col.endswith("_catalog_target") or col.endswith("_target") or col.endswith("_after"):
                a = _parse_number(row[col])
        out[cid] = (bl, b, a)

    return out


def _dedupe_raw_column_names(names: list) -> list:
    """Make column names unique for pandas (empty -> unnamed_N)."""
    out: list = []
    seen: Dict[str, int] = {}
    for i, raw in enumerate(names):
        base = str(raw).strip() if raw is not None and str(raw).strip() else f"unnamed_{i}"
        n = seen.get(base, 0)
        if n == 0:
            out.append(base)
            seen[base] = 1
        else:
            out.append(f"{base}_{n}")
            seen[base] = n + 1
    return out


def _first_df_row_looks_like_spending_active_headers(row: pd.Series) -> bool:
    """
    Excel often puts a merged 'Period Group' row above real column names.
    The real titles may be the first *or second* row of the DataFrame.

    Expected markers in the same row: category_id (or similar) + default/before + target/after.
    """
    tokens: list = []
    for x in row:
        if pd.isna(x):
            continue
        s = str(x).strip().lower().replace("\ufeff", "")
        if s:
            tokens.append(s)

    if not tokens:
        return False

    def _sn(t: str) -> str:
        return re.sub(r"\s+", "_", t.strip())

    cat_ok = any(
        _sn(t) in ("category_id", "categoryid")
        or ("category" in t and "id" in t)
        or t == "id"
        or t == "cat_id"
        for t in tokens
    )

    def_ok = any(
        t == "default" or _sn(t) in ("catalog_default", "before") for t in tokens
    )
    tgt_ok = any(
        t == "target" or t == "after" or _sn(t) == "catalog_target"
        for t in tokens
    )

    if not def_ok or not tgt_ok:
        for x in row:
            if pd.isna(x):
                continue
            cell = str(x).strip().lower().replace("\ufeff", "")
            if not cell:
                continue
            cs = _sn(cell)
            if not def_ok and (
                cell.endswith("_default")
                or cell.endswith("_before")
                or cs.endswith("_default")
                or cs.endswith("_before")
                or "catalog_default" in cs
            ):
                def_ok = True
            if not tgt_ok and (
                cell.endswith("_target")
                or cell.endswith("_after")
                or cs.endswith("_target")
                or cs.endswith("_after")
                or "catalog_target" in cs
            ):
                tgt_ok = True

    return bool(cat_ok and def_ok and tgt_ok)


def _maybe_promote_inline_header_row_spending_active(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Find the first row (within the first few lines) that looks like column titles,
    promote it to header, and drop all rows above including it.
    """
    if df_raw is None or df_raw.empty:
        return df_raw
    max_scan = min(6, len(df_raw))
    for i in range(max_scan):
        if i + 1 >= len(df_raw):
            break
        row = df_raw.iloc[i]
        if not _first_df_row_looks_like_spending_active_headers(row):
            continue
        raw_headers = [
            ("" if pd.isna(x) else str(x).strip().replace("\ufeff", "")) for x in row.tolist()
        ]
        out = df_raw.iloc[i + 1 :].copy()
        out.columns = _dedupe_raw_column_names(raw_headers)
        return out.reset_index(drop=True)
    return df_raw


def _price_header_row_has_category_and_period(row: pd.Series) -> bool:
    """
    True if the first data row looks like real headers: category_id + period group,
    before normalization (cell text as exported from Excel).
    """
    tokens: list = []
    for x in row:
        if pd.isna(x):
            continue
        s = str(x).strip().lower().replace("\ufeff", "")
        if s:
            tokens.append(s)

    if not tokens:
        return False

    def _spaced_to_snake(t: str) -> str:
        return re.sub(r"\s+", "_", t.strip())

    cat_ok = any(
        _spaced_to_snake(t) in ("category_id", "categoryid") for t in tokens
    ) or any("category" in t and "id" in t for t in tokens) or any(
        t in ("id", "cat_id") for t in tokens
    )
    pg_ok = any("period" in t and "group" in t for t in tokens) or any(
        _spaced_to_snake(t) == "period_group" for t in tokens
    )
    return bool(cat_ok and pg_ok)


def _parse_price_dataframe(df_raw: pd.DataFrame) -> Dict[int, Dict[str, Dict[str, Any]]]:
    df_work = df_raw.copy()

    def _normalize_and_find_id(d_in: pd.DataFrame) -> Tuple[pd.DataFrame, Optional[str]]:
        d_norm, _ = _normalize_columns(d_in)
        try:
            return d_norm, _category_id_column(d_norm)
        except ValueError:
            pass
        try:
            return d_norm, _resolve_category_id_column(d_norm)
        except ValueError:
            return d_norm, None

    df, id_col = _normalize_and_find_id(df_work)

    if id_col is None and len(df_work) > 0:
        _max_scan = min(8, len(df_work))
        for _i in range(_max_scan):
            if _i + 1 >= len(df_work):
                break
            top = df_work.iloc[_i]
            _row_looks_like_header = (
                _price_header_row_has_category_and_period(top)
                or _first_df_row_looks_like_spending_active_headers(top)
            )
            if not _row_looks_like_header:
                continue
            raw_headers = [
                ("" if pd.isna(x) else str(x).strip().replace("\ufeff", ""))
                for x in top.tolist()
            ]
            df_promoted = df_work.iloc[_i + 1 :].copy()
            df_promoted.columns = _dedupe_raw_column_names(raw_headers)
            df_promoted = df_promoted.reset_index(drop=True)
            df, id_col = _normalize_and_find_id(df_promoted)
            if id_col is not None:
                break

    if id_col is None:
        df_sp = _maybe_promote_inline_header_row_spending_active(df_raw.copy())
        df, id_col = _normalize_and_find_id(df_sp)

    if id_col is None:
        raise ValueError("Could not find category_id column in price export (check header row).")

    # period_group: Default vs Target
    pg_col = _find_column(
        df.columns,
        "period_group",
        "period_group_3",
        "catalog",
    )
    for c in df.columns:
        if "period" in c and "group" in c.lower():
            pg_col = c
            break

    price_data: Dict[int, Dict[str, Dict[str, Any]]] = {}

    if pg_col:
        for _, row in df.iterrows():
            cid = _coerce_category_id(row[id_col])
            if cid is None:
                continue
            pg = str(row.get(pg_col, "")).strip().lower()
            bucket_key = "default" if "default" in pg else "target" if "target" in pg else "default"
            payload: Dict[str, Any] = {}
            for c in df.columns:
                if c in (id_col, pg_col):
                    continue
                val = row[c]
                num = _parse_number(val)
                payload[str(c)] = num if num is not None else val
            if cid not in price_data:
                price_data[cid] = {"default": {}, "target": {}}
            price_data[cid][bucket_key].update(payload)
        return price_data

    # No period column: put all non-id columns under "default"
    for _, row in df.iterrows():
        cid = _coerce_category_id(row[id_col])
        if cid is None:
            continue
        payload = {}
        for c in df.columns:
            if c == id_col:
                continue
            val = row[c]
            num = _parse_number(val)
            payload[str(c)] = num if num is not None else val
        price_data[cid] = {"default": dict(payload), "target": {}}

    return price_data


def _warn_set_diff(name_a: str, name_b: str, set_a: set, set_b: set) -> None:
    only_a = set_a - set_b
    only_b = set_b - set_a
    if only_a:
        logger.warning("%s has %s category_id(s) not in %s: %s", name_a, len(only_a), name_b, sorted(list(only_a))[:50])
    if only_b:
        logger.warning("%s has %s category_id(s) not in %s: %s", name_b, len(only_b), name_a, sorted(list(only_b))[:50])


def load_and_merge_data(
    spending_file: PathLike,
    active_file: PathLike,
    price_file: PathLike,
) -> Tuple[Dict[int, Dict[str, Any]], Dict[int, Dict[str, Dict[str, Any]]]]:
    """
    Load three exports and merge spending + active listers by category_id.

    Returns:
        merged_data: category_id -> {"baseline": {...}, "before": {...}, "after": {...}}
            Core keys per side: paid_users, spending, active_listers (+ other spending metrics).
            ``baseline`` is filled when the export includes Baseline/Reference columns (else zeros in UI).
        price_data: category_id -> {"default": {...}, "target": {...}}  (not merged into merged_data)
    """
    spending_path, active_path, price_path = Path(spending_file), Path(active_file), Path(price_file)

    spend_df = _read_table(spending_path)
    active_df = _read_table(active_path)
    price_df = _read_table(price_path)

    spending_by_id = _retry_if_missing_category_id(
        spending_path, spend_df, _parse_spending_dataframe
    )
    active_by_id = _retry_if_missing_category_id(
        active_path, active_df, _parse_active_dataframe
    )
    price_data = _retry_if_missing_category_id(
        price_path, price_df, _parse_price_dataframe
    )

    ids_s = set(spending_by_id.keys())
    ids_a = set(active_by_id.keys())
    ids_p = set(price_data.keys())

    _warn_set_diff("spending", "active listers", ids_s, ids_a)
    _warn_set_diff("spending", "price per day", ids_s, ids_p)
    _warn_set_diff("active listers", "price per day", ids_a, ids_p)

    merged: Dict[int, Dict[str, Any]] = {}
    for cid in sorted(ids_s & ids_a):
        sb = dict(spending_by_id[cid]["before"])
        sa = dict(spending_by_id[cid]["after"])
        sb_line = dict(spending_by_id[cid].get("baseline") or {})
        abl, abefore, aafter = _unpack_active_triplet(active_by_id[cid])
        _kid = int(cid)
        merged[_kid] = {
            "baseline": {**sb_line, "active_listers": abl},
            "before": {**sb, "active_listers": abefore},
            "after": {**sa, "active_listers": aafter},
        }

    price_data = {int(k): v for k, v in price_data.items()}
    return merged, price_data


def load_and_merge_spending_active(
    spending_file: PathLike,
    active_file: PathLike,
) -> Dict[int, Dict[str, Any]]:
    """
    Same merged shape as load_and_merge_data[0], but only spending + active listers (no price file).
    """
    spending_path, active_path = Path(spending_file), Path(active_file)

    spend_df = _read_table(spending_path)
    active_df = _read_table(active_path)

    spending_by_id = _retry_if_missing_category_id(
        spending_path, spend_df, _parse_spending_dataframe
    )
    active_by_id = _retry_if_missing_category_id(
        active_path, active_df, _parse_active_dataframe
    )

    ids_s = set(spending_by_id.keys())
    ids_a = set(active_by_id.keys())
    _warn_set_diff("previous year spending", "previous year active listers", ids_s, ids_a)

    merged: Dict[int, Dict[str, Any]] = {}
    for cid in sorted(ids_s & ids_a):
        sb = dict(spending_by_id[cid]["before"])
        sa = dict(spending_by_id[cid]["after"])
        sb_line = dict(spending_by_id[cid].get("baseline") or {})
        abl, abefore, aafter = _unpack_active_triplet(active_by_id[cid])
        _kid = int(cid)
        merged[_kid] = {
            "baseline": {**sb_line, "active_listers": abl},
            "before": {**sb, "active_listers": abefore},
            "after": {**sa, "active_listers": aafter},
        }

    return merged
