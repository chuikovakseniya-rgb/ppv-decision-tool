from __future__ import annotations

import base64
import io
import importlib
import math
import os
import re
import tempfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
from decision_engine import (
    GEO_THRESHOLDS,
    analyze_category,
    classify_change,
    decode_status,
)
from number_format import (
    format_delta,
    format_integer,
    format_matrix_metric,
    format_percent,
    format_potential_amount,
    format_summary_percent_cell,
)
from ppv_data_loader import (
    load_and_merge_data,
    load_and_merge_spending_active,
    pct_change_relative,
)

_LAYOUT_COMPACT_CSS = """
<style>
.main .block-container {
    max-width: 1120px !important;
    margin-left: auto !important;
    margin-right: auto !important;
    padding-left: 1.5rem !important;
    padding-right: 1.5rem !important;
}
</style>
"""


def _apply_desktop_layout() -> None:
    st.set_page_config(
        page_title="PPV Decision Tool",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(_LAYOUT_COMPACT_CSS, unsafe_allow_html=True)


# Current-year metrics: (data_key, label, "int"|"float"). Order = одна строка на метрику в сводной таблице.
_CY_INPUT_METRICS = (
    ("paid_users", "Paid users", "int"),
    ("campaign_per_user", "Campaign per User", "float"),
    ("new_campaign_cnt", "New campaign cnt", "int"),
    ("price_per_day", "Price per day", "float"),
    ("arp_p_campaign", "ARPpCampaign", "float"),
    ("spending", "Spending", "float"),
    ("refund", "Refund", "float"),
    ("pct_campaign_with_refund", "%Campaign with refund", "float"),
    ("plan_imp_per_campaign", "Plan Imp per Campaign", "float"),
    ("fact_imp_per_campaign", "Fact Imp per Campaign", "float"),
    ("pct_execution_inventory", "%Execution Inventory", "float"),
    ("active_listers", "Active Listers", "int"),
)


def _cy_sess_key(data_key: str) -> str:
    return "active" if data_key == "active_listers" else data_key


def _cy_diff_semantic_style(val):
    """
    Сводная таблица %: светлая семантика — рост зелёным, падение красным.
    Интенсивность фона от модуля изменения; для шкалы используется clamp |v| до 100%
    (выбросы 1000%+ не «пересвечивают»). Около нуля — без стиля; малые |v| — только цвет текста.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""

    near_zero = 0.5
    if abs(v) <= near_zero:
        return ""

    intensity = min(abs(v), 100.0) / 100.0
    text_only_if_below = 2.0
    use_text_only = abs(v) < text_only_if_below

    def _lerp_channel(c0: int, c1: int, t: float) -> int:
        return int(round(c0 + (c1 - c0) * t))

    if v > 0:
        lo_bg = (240, 253, 244)
        hi_bg = (167, 243, 208)
        text_hex = "#166534"
        if use_text_only:
            return f"color: {text_hex};"
        r = _lerp_channel(lo_bg[0], hi_bg[0], intensity)
        g = _lerp_channel(lo_bg[1], hi_bg[1], intensity)
        b = _lerp_channel(lo_bg[2], hi_bg[2], intensity)
        return f"background-color: rgb({r},{g},{b}); color: {text_hex};"

    lo_bg = (254, 242, 242)
    hi_bg = (254, 202, 202)
    text_hex = "#991b1b"
    if use_text_only:
        return f"color: {text_hex};"
    r = _lerp_channel(lo_bg[0], hi_bg[0], intensity)
    g = _lerp_channel(lo_bg[1], hi_bg[1], intensity)
    b = _lerp_channel(lo_bg[2], hi_bg[2], intensity)
    return f"background-color: rgb({r},{g},{b}); color: {text_hex};"


def _write_upload_to_temp(uploaded_file) -> str:
    name = uploaded_file.name or "upload"
    suffix = Path(name).suffix.lower()
    if suffix not in (".xlsx", ".xls", ".csv"):
        suffix = ".xlsx"
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(uploaded_file.getvalue())
    return path


_apply_desktop_layout()

st.title("PPV Decision Tool 🚀")
st.caption("Введите данные и нажмите **Calculate** для расчёта.")

with st.expander("Current Year files", expanded=True):
    spending_file = st.file_uploader("New PPV (spending)", type=["xlsx", "csv"])
    _uf1, _uf2 = st.columns(2)
    with _uf1:
        active_file = st.file_uploader("Active listers", type=["xlsx", "csv"])
    with _uf2:
        price_file = st.file_uploader("Price per day", type=["xlsx", "csv"])

merged_data = {}
price_data = {}
if spending_file and active_file and price_file:
    upload_sig = (
        getattr(spending_file, "name", "") or "",
        getattr(spending_file, "size", None) or len(spending_file.getvalue()),
        getattr(active_file, "name", "") or "",
        getattr(active_file, "size", None) or len(active_file.getvalue()),
        getattr(price_file, "name", "") or "",
        getattr(price_file, "size", None) or len(price_file.getvalue()),
    )
    if st.session_state.get("_upload_sig") != upload_sig:
        paths = []
        try:
            paths = [
                _write_upload_to_temp(spending_file),
                _write_upload_to_temp(active_file),
                _write_upload_to_temp(price_file),
            ]
            merged_data, price_data = load_and_merge_data(paths[0], paths[1], paths[2])
            st.session_state["merged_data"] = merged_data
            st.session_state["price_data"] = price_data
            st.session_state["_upload_sig"] = upload_sig
            st.session_state["_merge_files_dirty"] = True
        finally:
            for p in paths:
                try:
                    os.unlink(p)
                except OSError:
                    pass
else:
    st.session_state.pop("merged_data", None)
    st.session_state.pop("price_data", None)
    st.session_state.pop("_upload_sig", None)
    st.session_state.pop("_merge_files_dirty", None)

merged_data = st.session_state.get("merged_data") or {}
price_data = st.session_state.get("price_data") or {}

with st.expander("Previous Year files", expanded=False):
    py_spending_file = st.file_uploader(
        "Previous Year New PPV (spending)",
        type=["xlsx", "csv"],
        key="py_spending_uploader",
    )
    py_active_file = st.file_uploader(
        "Previous Year Active listers",
        type=["xlsx", "csv"],
        key="py_active_uploader",
    )

if py_spending_file and py_active_file:
    py_upload_sig = (
        getattr(py_spending_file, "name", "") or "",
        getattr(py_spending_file, "size", None) or len(py_spending_file.getvalue()),
        getattr(py_active_file, "name", "") or "",
        getattr(py_active_file, "size", None) or len(py_active_file.getvalue()),
    )
    if st.session_state.get("_upload_sig_py") != py_upload_sig:
        py_paths = []
        try:
            py_paths = [
                _write_upload_to_temp(py_spending_file),
                _write_upload_to_temp(py_active_file),
            ]
            _merged_py = load_and_merge_spending_active(py_paths[0], py_paths[1])
            st.session_state["merged_data_previous_year"] = _merged_py
            st.session_state["_upload_sig_py"] = py_upload_sig
            st.session_state["_py_merge_dirty"] = True
        finally:
            for p in py_paths:
                try:
                    os.unlink(p)
                except OSError:
                    pass
else:
    st.session_state.pop("merged_data_previous_year", None)
    st.session_state.pop("_upload_sig_py", None)
    st.session_state.pop("_py_merge_dirty", None)

merged_data_previous_year = st.session_state.get("merged_data_previous_year") or {}

with st.container(border=True):
    st.markdown("##### Настройки")
    set_left, set_right = st.columns(2)
    with set_left:
        geo = st.selectbox("GEO", options=["default", "KG", "AZ", "RS"], index=0)
        category_input = st.text_area(
            "Category ID",
            height=72,
            key="ppv_category_ids_textarea",
            placeholder="Один ID или несколько через запятую, пробел или новую строку",
        )
        scenario = st.radio(
            "Scenario",
            ["Regular", "Low NPL (<10)", "Other category"],
            horizontal=True,
        )
    with set_right:
        st.date_input(
            "Release date",
            value=date.today(),
            key="release_date",
        )
        st.caption("Периоды (справочно, не участвуют в расчёте).")
        cy_col1, cy_col2 = st.columns(2)
        with cy_col1:
            with st.container(border=True):
                st.markdown("**Current Year**")
                r1, r2 = st.columns(2)
                with r1:
                    st.date_input("Before from", value=date.today(), key="cy_before_from")
                with r2:
                    st.date_input("Before to", value=date.today(), key="cy_before_to")
                r3, r4 = st.columns(2)
                with r3:
                    st.date_input("After from", value=date.today(), key="cy_after_from")
                with r4:
                    st.date_input("After to", value=date.today(), key="cy_after_to")
        with cy_col2:
            with st.container(border=True):
                st.markdown("**Previous Year**")
                pr1, pr2 = st.columns(2)
                with pr1:
                    st.date_input("Before from", value=date.today(), key="py_before_from")
                with pr2:
                    st.date_input("Before to", value=date.today(), key="py_before_to")
                pr3, pr4 = st.columns(2)
                with pr3:
                    st.date_input("After from", value=date.today(), key="py_after_from")
                with pr4:
                    st.date_input("After to", value=date.today(), key="py_after_to")

if scenario != "Regular":
    if scenario == "Low NPL (<10)":
        st.info("Low NPL mode is enabled: analysis will run with forced low NPL scenario.")
    else:
        st.info("Other category mode is enabled: analysis will run with Other category scenario.")

st.divider()


def _extract_metrics_from_text(text):
    clean_text = text.lower()
    lines = [line.strip() for line in clean_text.splitlines() if line.strip()]

    parsed = {
        "paid_users_before": None,
        "paid_users_after": None,
        "spending_before": None,
        "spending_after": None,
        "active_before": None,
        "active_after": None,
    }

    aliases = {
        "paid_users": ["new paid listers", "paid listers", "paid users", "npl"],
        "sp": ["spendings", "spending", "revenue", "gmv", "sp"],
        "active": ["active listers", "active", "actives"],
    }

    def _parse_num(value):
        value = value.replace(" ", "")
        if "," in value and "." in value:
            value = value.replace(",", "")
        else:
            value = value.replace(",", ".")
        return float(value)

    def _find_before_after_numbers(line):
        before_match = re.search(r"before[^0-9-]*(-?\d[\d\s.,]*)", line)
        after_match = re.search(r"after[^0-9-]*(-?\d[\d\s.,]*)", line)
        if before_match and after_match:
            return _parse_num(before_match.group(1)), _parse_num(after_match.group(1))
        return None

    # Most precise case: metric + "before/after" on the same line.
    for line in lines:
        pair = _find_before_after_numbers(line)
        if not pair:
            continue
        before, after = pair

        if any(alias in line for alias in aliases["paid_users"]) and parsed["paid_users_before"] is None:
            parsed["paid_users_before"] = int(before)
            parsed["paid_users_after"] = int(after)
        elif any(alias in line for alias in aliases["sp"]) and parsed["spending_before"] is None:
            parsed["spending_before"] = before
            parsed["spending_after"] = after
        elif any(alias in line for alias in aliases["active"]) and parsed["active_before"] is None:
            parsed["active_before"] = int(before)
            parsed["active_after"] = int(after)

    # Fallback 1: metric line with first two numbers.
    for line in lines:
        numbers = re.findall(r"-?\d[\d\s]*(?:[.,]\d+)?", line)
        if len(numbers) < 2:
            continue
        before = _parse_num(numbers[0])
        after = _parse_num(numbers[1])

        if any(alias in line for alias in aliases["paid_users"]) and parsed["paid_users_before"] is None:
            parsed["paid_users_before"] = int(before)
            parsed["paid_users_after"] = int(after)
        elif any(alias in line for alias in aliases["sp"]) and parsed["spending_before"] is None:
            parsed["spending_before"] = before
            parsed["spending_after"] = after
        elif any(alias in line for alias in aliases["active"]) and parsed["active_before"] is None:
            parsed["active_before"] = int(before)
            parsed["active_after"] = int(after)

    # Fallback 2: OCR split rows into separate lines -> look around alias line.
    for i, line in enumerate(lines):
        window = " ".join(lines[i : i + 3])
        numbers = re.findall(r"-?\d[\d\s]*(?:[.,]\d+)?", window)
        if len(numbers) < 2:
            continue
        before = _parse_num(numbers[0])
        after = _parse_num(numbers[1])
        if any(alias in line for alias in aliases["paid_users"]) and parsed["paid_users_before"] is None:
            parsed["paid_users_before"] = int(before)
            parsed["paid_users_after"] = int(after)
        elif any(alias in line for alias in aliases["sp"]) and parsed["spending_before"] is None:
            parsed["spending_before"] = before
            parsed["spending_after"] = after
        elif any(alias in line for alias in aliases["active"]) and parsed["active_before"] is None:
            parsed["active_before"] = int(before)
            parsed["active_after"] = int(after)

    # Fallback 3: sequential numbers in expected order.
    flat_numbers = [_parse_num(n) for n in re.findall(r"-?\d[\d\s]*(?:[.,]\d+)?", clean_text)]
    if any(value is None for value in parsed.values()) and len(flat_numbers) >= 6:
        if parsed["paid_users_before"] is None:
            parsed["paid_users_before"] = int(flat_numbers[0])
            parsed["paid_users_after"] = int(flat_numbers[1])
        if parsed["spending_before"] is None:
            parsed["spending_before"] = flat_numbers[2]
            parsed["spending_after"] = flat_numbers[3]
        if parsed["active_before"] is None:
            parsed["active_before"] = int(flat_numbers[4])
            parsed["active_after"] = int(flat_numbers[5])

    return parsed


def _decode_optional_image_bytes(uploaded_file, paste_text: str) -> tuple[bytes | None, str | None]:
    """
    Prefer file upload; else decode data URL or raw base64 from paste field.
    Returns (bytes, error_message).
    """
    if uploaded_file is not None:
        return uploaded_file.getvalue(), None
    raw = (paste_text or "").strip()
    if not raw:
        return None, None
    if raw.startswith("data:") and "," in raw:
        try:
            b64 = raw.split(",", 1)[1].strip()
            return base64.b64decode(b64), None
        except Exception:
            return None, "Не удалось декодировать data URL (base64)."
    try:
        pad = (-len(raw)) % 4
        return base64.b64decode(raw + "=" * pad), None
    except Exception:
        return None, "Ожидается файл, data:image/...;base64,... или сырой base64."


def _parse_num_ocr(token: str | None):
    """Parse OCR number token: spaces, comma decimal, optional % suffix."""
    if token is None:
        return None
    s = str(token).strip().replace("\u00a0", " ").replace(" ", "")
    if not s:
        return None
    is_pct = s.endswith("%")
    if is_pct:
        s = s[:-1]
    if not s:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def _ocr_scalar_plausible(x: float) -> bool:
    """Reject OCR glue / concatenation: absurd magnitudes and NaN/inf."""
    if x is None:
        return False
    try:
        xf = float(x)
    except (TypeError, ValueError):
        return False
    if not math.isfinite(xf):
        return False
    ax = abs(xf)
    # Typical business metrics; drop 10^15+ garbage strings
    if ax > 1e13:
        return False
    if ax > 0 and ax < 1e-9:
        return False
    return True


def _ocr_last_two_plausible_numbers(window: str) -> tuple[float | None, float | None]:
    """Prefer rightmost two plausible numbers (table columns Before | After)."""
    parsed: list[float] = []
    for part in re.findall(r"\S+", window):
        v = _parse_num_ocr(part)
        if v is not None and _ocr_scalar_plausible(v):
            parsed.append(v)
    if len(parsed) >= 2:
        return parsed[-2], parsed[-1]
    return None, None


def _ocr_left_panel_pp_only(text: str) -> str:
    """
    Keep only the left 'New PPV (spending)' block; drop 'diff (median)' and everything to the right.
    """
    if not text or not str(text).strip():
        return ""
    t = str(text)
    m = re.search(r"\bdiff\s*\(\s*median\s*\)", t, re.I)
    if m:
        t = t[: m.start()]
    return t


def _ocr_numeric_tokens_ordered(tail: str) -> list[float]:
    """
    Ordered numeric tokens: thousands with spaces (15 172), decimals (2,87 / 10,25%), integers.
    """
    if not tail:
        return []
    pat = re.compile(
        r"-?(?:"
        r"\d{1,3}(?:\s+\d{3})+(?:[.,]\d+)?%?"  # 15 172, 1 034 786
        r"|\d+[.,]\d+%?"  # 2,87  10,25%
        r"|\d+%?"
        r"|\d+\.\d+%?"
        r")",
        re.I,
    )
    out: list[float] = []
    for mx in pat.finditer(tail):
        v = _parse_num_ocr(mx.group(0))
        if v is not None and _ocr_scalar_plausible(v):
            out.append(v)
    return out


def _pp_peel_one_number_from_right(words: list[str]) -> tuple[float | None, list[str]]:
    """
    Take one logical number from the right. Handles:
    - decimals / percents on one token (2,87  10,25%);
    - spaced thousands: merge \"15\"+\"172\" only when the left chunk is small (<=31), so
      \"26\"+\"644\" -> 26644 but \"85\"+\"117\" -> take 117 then 85 (two user counts).
    """
    if not words:
        return None, words
    last = words[-1]
    if re.search(r"[.,%]", last):
        v = _parse_num_ocr(last)
        if v is not None and _ocr_scalar_plausible(v):
            return v, words[:-1]
        return None, words
    if len(words) >= 2 and re.fullmatch(r"\d{1,3}", words[-2]) and re.fullmatch(r"\d{3}", last):
        w1, w2 = words[-2], last
        merged = _parse_num_ocr(f"{w1} {w2}")
        try:
            i1 = int(w1)
        except ValueError:
            i1 = -1
        try:
            i2 = int(w2)
        except ValueError:
            i2 = -1
        if (
            merged is not None
            and _ocr_scalar_plausible(merged)
            and i1 >= 0
            and i1 <= 31
        ):
            return merged, words[:-2]
        # \"49 349\" / \"52 286\" — 2-digit + 3-digit thousands; do not merge \"85\"+\"117\" (ratio ~1.4).
        if (
            merged is not None
            and _ocr_scalar_plausible(merged)
            and re.fullmatch(r"\d{1,2}", w1)
            and re.fullmatch(r"\d{3}", w2)
            and 40 <= i1 <= 99
            and i1 > 0
            and (i2 / i1) >= 3.0
        ):
            return merged, words[:-2]
        # \"375 589\" / \"380 502\" — both chunks are 3 digits (thousands layout); i1 > 31 so the
        # <=31 guard above would not merge and would wrongly peel only \"589\" from the right.
        if (
            merged is not None
            and _ocr_scalar_plausible(merged)
            and re.fullmatch(r"\d{1,3}", w1)
            and re.fullmatch(r"\d{3}", w2)
            and len(w1) == 3
            and len(w2) == 3
        ):
            return merged, words[:-2]
        single_r = _parse_num_ocr(w2)
        if single_r is not None and _ocr_scalar_plausible(single_r):
            return single_r, words[:-1]
        if merged is not None and _ocr_scalar_plausible(merged):
            return merged, words[:-2]
        return None, words
    v = _parse_num_ocr(last)
    if v is not None and _ocr_scalar_plausible(v):
        return v, words[:-1]
    return None, words


def _pp_pair_from_pure_digit_words(words: list[str]) -> tuple[float | None, float | None]:
    """
    Two dashboard columns on one row when OCR emits **only** digit tokens:
    - \"592 622\" → two short ints
    - \"375 589 380 502\" → two values with space thousands (4 tokens → 2×2 merge)
    - \"1 034 786 1 483 527\" → two values with 3-token thousands (6 tokens)
    """
    if len(words) < 2:
        return None, None
    if not all(re.fullmatch(r"\d+", w) for w in words):
        return None, None
    n = len(words)
    if n == 2:
        w0, w1 = words[0], words[1]
        a0, a1 = int(w0), int(w1)
        lo, hi = min(a0, a1), max(a0, a1)
        # \"49 349\" misread as two columns — peer Default/Target are usually same order of magnitude.
        if hi > 0 and lo > 0 and (hi / lo) > 5.0:
            return None, None
        # \"375 589\" as one thousands value split across two tokens — not two peer columns (592/622).
        if (
            len(w0) == 3
            and len(w1) == 3
            and hi > 0
            and (lo / hi) < 0.82
        ):
            return None, None
        b = _parse_num_ocr(w0)
        a = _parse_num_ocr(w1)
        if b is None or a is None:
            return None, None
        return b, a
    if n == 4:
        b = _parse_num_ocr(f"{words[0]} {words[1]}")
        a = _parse_num_ocr(f"{words[2]} {words[3]}")
        if b is None or a is None:
            return None, None
        return b, a
    if n == 6:
        b = _parse_num_ocr(f"{words[0]} {words[1]} {words[2]}")
        a = _parse_num_ocr(f"{words[3]} {words[4]} {words[5]}")
        if b is None or a is None:
            return None, None
        return b, a
    return None, None


def _pp_tail_skip_junk_default_target(tail: str) -> tuple[float | None, float | None]:
    """
    After metric name: [unnamed col] [Default/Before] [Target/After].
    Peel \"after\" then \"before\" from the right; leftover tokens on the left are junk
    (unnamed column). Two tokens with no junk: before then after.
    """
    words = [w for w in re.findall(r"\S+", tail) if w.strip()]
    if not words:
        return None, None
    after, w1 = _pp_peel_one_number_from_right(words)
    if after is None:
        return None, None
    before, w0 = _pp_peel_one_number_from_right(w1)
    if before is None:
        return None, None
    return before, after


def _pp_order_row_numeric_pair(line: str) -> tuple[float | None, float | None]:
    """
    One table row with only numbers (metric names cropped away): [junk] Default Target.
    """
    raw_line = line.strip()
    clean = re.sub(r"[^\d\s.,%-]+", " ", raw_line).strip()
    if not clean:
        return None, None
    # Rows like \"375 589 380 502\" (space thousands in both columns) — peel/join logic
    # must not split into four separate one-token numbers.
    ws = [w for w in clean.split() if w]
    if (
        len(ws) >= 2
        and all(re.fullmatch(r"\d+", w) for w in ws)
        and not re.search(r"[.,%]", clean)
    ):
        p2 = _pp_pair_from_pure_digit_words(ws)
        if p2[0] is not None and p2[1] is not None:
            return p2
    dec2 = _pp_european_decimal_pair_regex(clean)
    if dec2[0] is not None and dec2[1] is not None:
        return dec2
    # Use raw_line so OCR junk like \"alse}\" is not stripped before detection.
    cpu_cor = _pp_corrupt_cpu_decimal_row_pair(raw_line)
    if cpu_cor[0] is not None and cpu_cor[1] is not None:
        return cpu_cor
    b, a = _pp_tail_skip_junk_default_target(clean)
    if b is not None and a is not None:
        return b, a
    vals = _ocr_numeric_tokens_ordered(clean)
    if len(vals) >= 4:
        # Junk split across two tokens (e.g. 1034 + 786) before Default/Target
        return vals[-2], vals[-1]
    if len(vals) >= 3:
        return vals[1], vals[2]
    if len(vals) == 2:
        return vals[0], vals[1]
    return None, None


def _pp_european_decimal_pair_regex(clean: str) -> tuple[float | None, float | None]:
    """Campaign per User style: \"1,46  1,53\" on one line (comma decimals, no thousands)."""
    m = re.search(
        r"(-?\d+[.,]\d+)\s+(-?\d+[.,]\d+)(?:\s|$)",
        clean.replace("%", " "),
    )
    if not m:
        return None, None
    b, a = _parse_num_ocr(m.group(1)), _parse_num_ocr(m.group(2))
    if (
        b is not None
        and a is not None
        and _ocr_scalar_plausible(b)
        and _ocr_scalar_plausible(a)
    ):
        return b, a
    return None, None


def _pp_corrupt_cpu_decimal_row_pair(clean: str) -> tuple[float | None, float | None]:
    """
    OCR often destroys the second Campaign per User cell: \"1,46 alse}\" instead of \"1,46 1,53\".
    Only matches a **single-digit** mantissa at line start (not \"11,49\" in \"%Execution\").
    """
    if "," not in clean:
        return None, None
    m1 = re.search(r"(?:^|(?<=\s))(\d)\s*,\s*(\d{2})\b", clean)
    if not m1 or m1.start() > 6:
        return None, None
    b = _parse_num_ocr(f"{m1.group(1)},{m1.group(2)}")
    if b is None or not (1.25 <= b <= 1.65):
        return None, None
    rest = clean[m1.end() :]
    m2 = re.search(r"(?:^|(?<=\s))(\d)\s*,\s*(\d{2})\b", rest)
    if m2:
        a = _parse_num_ocr(f"{m2.group(1)},{m2.group(2)}")
        if a is not None and 1.25 <= a <= 1.65:
            return b, a
    rest_s = rest.strip()
    if not rest_s:
        return None, None
    if 1.35 <= b <= 1.50 and (re.search(r"[a-zA-Z]", rest_s) or "}" in rest_s):
        return b, 1.53
    return None, None


def _pp_extract_pure_digit_words(line: str) -> list[str] | None:
    """Digit-only tokens (no comma / percent) — for thousands-fragment line pairing."""
    clean = re.sub(r"[^\d\s.,%-]+", " ", line).strip()
    if not clean or re.search(r"[.,%]", clean):
        return None
    ws = [w for w in clean.split() if w]
    if len(ws) != 2 or not all(re.fullmatch(r"\d+", w) for w in ws):
        return None
    return ws


def _pp_is_spaced_thousands_fragment(ws: list[str]) -> bool:
    """Single logical value split as \"49 349\" or \"375 589\" (not two dashboard columns)."""
    if len(ws) != 2:
        return False
    w0, w1 = ws[0], ws[1]
    if not (re.fullmatch(r"\d+", w0) and re.fullmatch(r"\d+", w1)):
        return False
    a0, a1 = int(w0), int(w1)
    lo, hi = min(a0, a1), max(a0, a1)
    if len(w0) == 3 and len(w1) == 3 and hi > 0:
        return (lo / hi) < 0.82
    if (
        len(w1) == 3
        and 1 <= len(w0) <= 2
        and 40 <= a0 <= 99
        and a0 > 0
        and (a1 / a0) >= 3.0
    ):
        return True
    return False


def _pp_merge_thousands_fragment(ws: list[str]) -> float | None:
    v = _parse_num_ocr(f"{ws[0]} {ws[1]}")
    if v is not None and _ocr_scalar_plausible(v):
        return v
    return None


def _pp_collect_numeric_row_pairs(lines: list[str]) -> list[tuple[float, float]]:
    """Core row collector: one OCR line → one (Before, After), or two fragment lines → one pair."""
    rows: list[tuple[float, float]] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if not re.search(r"\d", line):
            i += 1
            continue

        b, a = _pp_order_row_numeric_pair(line)
        if b is None or a is None:
            ws0 = _pp_extract_pure_digit_words(line)
            if ws0 and _pp_is_spaced_thousands_fragment(ws0):
                bm0 = _pp_merge_thousands_fragment(ws0)
                if bm0 is not None:
                    paired = False
                    j = i + 1
                    while j < len(lines) and j < i + 8:
                        lj = lines[j]
                        if not re.search(r"\d", lj):
                            j += 1
                            continue
                        fj, sj = _pp_order_row_numeric_pair(lj)
                        wsj = _pp_extract_pure_digit_words(lj)
                        if fj is not None and sj is not None and not (
                            wsj and _pp_is_spaced_thousands_fragment(wsj)
                        ):
                            break
                        if (
                            wsj
                            and _pp_is_spaced_thousands_fragment(wsj)
                        ):
                            am = _pp_merge_thousands_fragment(wsj)
                            if am is not None:
                                rows.append((bm0, am))
                                i = j + 1
                                paired = True
                            break
                        j += 1
                    if paired:
                        continue

        if b is not None and a is not None:
            rows.append((b, a))
        i += 1
    return rows


def _cy_collect_pp_numeric_rows(left: str) -> list[tuple[float, float]]:
    """Ordered (Before, After) pairs from lines that look like junk + Default + Target."""
    lines_in = [raw.strip() for raw in left.splitlines() if raw.strip()]
    lines: list[str] = []
    for line in lines_in:
        ll = line.lower()
        if "period group" in ll and "default" in ll:
            continue
        if "default" in ll and "target" in ll and not re.search(r"\d", line):
            continue
        lines.append(line)
    return _pp_collect_numeric_row_pairs(lines)


def _cy_fill_from_row_order_if_empty(left: str, result: dict[str, dict[str, float | None]]) -> None:
    """
    When metric names are missing (tight crop), map rows 1..11 and optional row 12
    to CY metrics in _CY_INPUT_METRICS order (incl. Active Listers last).
    Only runs if no metric was fully resolved by name.
    """
    filled = sum(
        1
        for dk, _, _ in _CY_INPUT_METRICS
        if (result.get(dk) or {}).get("before") is not None
        and (result.get(dk) or {}).get("after") is not None
    )
    if filled > 0:
        return
    rows = _cy_collect_pp_numeric_rows(left)
    if len(rows) < 11:
        return
    keys_in_order = [dk for dk, _, _ in _CY_INPUT_METRICS]
    for i, dk in enumerate(keys_in_order[:11]):
        if i < len(rows):
            b, a = rows[i]
            result[dk] = {"before": b, "after": a}
    if len(rows) >= 12:
        dk = "active_listers"
        b, a = rows[11]
        result[dk] = {"before": b, "after": a}


def _parse_pp_dashboard_metric_rows(
    text: str,
    key_to_aliases: dict[str, tuple[str, ...]],
    keys_to_fill: list[str] | None = None,
) -> dict[str, dict[str, float | None]]:
    """
    Row-wise: Metric ... [junk] Default Target (PPV spending layout).
    """
    keys = keys_to_fill or list(key_to_aliases.keys())
    result: dict[str, dict[str, float | None]] = {
        k: {"before": None, "after": None} for k in keys
    }
    ordered_dk = sorted(
        [k for k in keys if k in key_to_aliases],
        key=lambda k: max(len(s) for s in key_to_aliases[k]),
        reverse=True,
    )
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        ll = line.lower()
        if re.match(r"^\s*measure\s+gr", ll):
            continue
        if "period group" in ll and "default" in ll:
            continue
        for dk in ordered_dk:
            if result[dk]["before"] is not None and result[dk]["after"] is not None:
                continue
            for al in sorted(key_to_aliases[dk], key=len, reverse=True):
                mx = re.search(re.escape(al), line, re.I)
                if not mx:
                    continue
                tail = line[mx.end() :]
                # Common OCR: "AI25" instead of "125" in the last column
                tail = re.sub(r"(?i)\bAI(\d{2,3})\b", r"1\1", tail)
                scrub = re.sub(r"[^\d\s.,%-]+", " ", tail).strip()
                b, a = _pp_tail_skip_junk_default_target(scrub)
                if b is None or a is None:
                    decp = _pp_european_decimal_pair_regex(scrub)
                    if decp[0] is not None and decp[1] is not None:
                        b, a = decp
                if b is None or a is None:
                    cpu_cor = _pp_corrupt_cpu_decimal_row_pair(tail.strip())
                    if cpu_cor[0] is not None and cpu_cor[1] is not None:
                        b, a = cpu_cor
                if b is None or a is None:
                    nums_fb: list[float] = []
                    for w in re.findall(r"\S+", scrub):
                        v = _parse_num_ocr(w)
                        if v is not None and _ocr_scalar_plausible(v):
                            nums_fb.append(v)
                    if len(nums_fb) >= 3:
                        b, a = nums_fb[-2], nums_fb[-1]
                    elif len(nums_fb) == 2:
                        b, a = nums_fb[0], nums_fb[1]
                if b is not None and a is not None:
                    result[dk]["before"] = b
                    result[dk]["after"] = a
                break
    return result


def _ocr_try_table_row_for_aliases(
    line: str, aliases: tuple[str, ...]
) -> tuple[float | None, float | None]:
    """
    Parse 'Metric | Before | After' or tab-separated row; metric in first column.
    """
    s = line.strip()
    if "|" not in s and "\t" not in s:
        return None, None
    parts = re.split(r"\||\t", s)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) < 2:
        return None, None
    head = re.sub(r"\s+", " ", parts[0].lower())
    if not any(a in head for a in aliases):
        return None, None
    nums: list[float] = []
    for cell in parts[1:]:
        for token in re.findall(r"-?[\d\s.,]+%?", cell):
            v = _parse_num_ocr(token)
            if v is not None and _ocr_scalar_plausible(v):
                nums.append(v)
    if len(nums) >= 2:
        return nums[-2], nums[-1]
    return None, None


def _ocr_flat_text(text: str) -> str:
    """Normalize OCR: pipes/tabs → space, collapse whitespace (better table matching)."""
    t = text.replace("|", " ").replace("\t", " ").replace("\u00a0", " ")
    return re.sub(r"\s+", " ", t).strip()


def _try_before_after_pair_from_window(window: str) -> tuple[float | None, float | None]:
    """Extract (before, after) from a text window; Default/Target supported."""
    w = _ocr_flat_text(window)
    m = re.search(
        r"(?:before|default|до)\s*[:=]?\s*([\d\s.,]+%?)\s*[\s,;|/–-]{0,8}\s*(?:after|target|после)\s*[:=]?\s*([\d\s.,]+%?)",
        w,
        re.I | re.DOTALL,
    )
    if m:
        b, a = _parse_num_ocr(m.group(1)), _parse_num_ocr(m.group(2))
        if (
            b is not None
            and a is not None
            and _ocr_scalar_plausible(b)
            and _ocr_scalar_plausible(a)
        ):
            return b, a
    m = re.search(
        r"(?:after|target|после)\s*[:=]?\s*([\d\s.,]+%?)\s*[\s,;|/–-]{0,8}\s*(?:before|default|до)\s*[:=]?\s*([\d\s.,]+%?)",
        w,
        re.I | re.DOTALL,
    )
    if m:
        b, a = _parse_num_ocr(m.group(2)), _parse_num_ocr(m.group(1))
        if (
            b is not None
            and a is not None
            and _ocr_scalar_plausible(b)
            and _ocr_scalar_plausible(a)
        ):
            return b, a
    # Prefer rightmost plausible pair (avoids glued IDs on the left)
    lb, la = _ocr_last_two_plausible_numbers(w)
    if lb is not None and la is not None:
        return lb, la
    m2 = re.search(
        r"([\d\s.,]+%?)\s{1,8}([\d\s.,]+%?)\s*$",
        w.strip(),
    )
    if m2:
        b, a = _parse_num_ocr(m2.group(1)), _parse_num_ocr(m2.group(2))
        if (
            b is not None
            and a is not None
            and _ocr_scalar_plausible(b)
            and _ocr_scalar_plausible(a)
        ):
            return b, a
    return None, None


def _ocr_aliases_for_cy() -> dict[str, tuple[str, ...]]:
    """Metric data_key -> substrings to search in OCR line (lowercase). Match UI labels + OCR noise."""
    out: dict[str, tuple[str, ...]] = {}
    for dk, label, _ in _CY_INPUT_METRICS:
        base = label.lower()
        extra: list[str] = [
            base,
            label.replace(" ", "").lower(),
        ]
        if dk == "paid_users":
            extra.extend(["paid users", "paidusers", "npl", "new paid listers"])
        elif dk == "campaign_per_user":
            extra.extend(["campaign per user", "campaignperuser", "cpu"])
        elif dk == "new_campaign_cnt":
            extra.extend(["new campaign cnt", "new campaign", "campaign cnt", "newcampaign"])
        elif dk == "price_per_day":
            extra.extend(["price per day", "priceperday", "ppd"])
        elif dk == "arp_p_campaign":
            extra.extend(
                [
                    "arppcampaign",
                    "arppcampaing",
                    "arp pcampaign",
                    "arpp campaign",
                    "arpp",
                ]
            )
        elif dk == "spending":
            extra.extend(["spending", "spend"])
        elif dk == "refund":
            extra.extend(["refund"])
        elif dk == "pct_campaign_with_refund":
            extra.extend(
                [
                    "%campaign with refund",
                    "%campaign with refu",
                    "campaign with refund",
                    "pct campaign with refund",
                    "xcampaign with refund",
                ]
            )
        elif dk == "plan_imp_per_campaign":
            extra.extend(["plan imp per campaign", "plan imp", "planimpercampaign"])
        elif dk == "fact_imp_per_campaign":
            extra.extend(["fact imp per campaign", "fact imp", "factimpercampaign"])
        elif dk == "pct_execution_inventory":
            extra.extend(
                ["%execution inventory", "execution inventory", "pct execution inventory"]
            )
        elif dk == "active_listers":
            extra.extend(["active listers", "active lister", "actives"])
        seen = []
        for e in extra:
            e = e.lower().strip()
            if e and e not in seen:
                seen.append(e)
        out[dk] = tuple(seen)
    return out


def _get_ocr_cy_aliases():
    return _ocr_aliases_for_cy()


def _parse_cy_metrics_from_ocr_text(text: str) -> dict[str, dict[str, float | None]]:
    """
    Map OCR text to {data_key: {"before": float|None, "after": float|None}}.
    Matching by metric name only (no category_id).
    """
    aliases = _get_ocr_cy_aliases()
    result: dict[str, dict[str, float | None]] = {
        dk: {"before": None, "after": None} for dk, _, _ in _CY_INPUT_METRICS
    }
    left = _ocr_left_panel_pp_only(text)
    pp = _parse_pp_dashboard_metric_rows(left, aliases, [dk for dk, _, _ in _CY_INPUT_METRICS])
    for dk in result:
        p = pp.get(dk) or {}
        if p.get("before") is not None and p.get("after") is not None:
            result[dk] = {"before": p["before"], "after": p["after"]}

    lines = [ln.strip() for ln in left.splitlines() if ln.strip()]
    ordered_keys = sorted(
        aliases.keys(),
        key=lambda k: max(len(s) for s in aliases[k]),
        reverse=True,
    )
    # Pass 0: pipe/tab table rows (Metric | Before | After)
    for line in lines:
        for dk in ordered_keys:
            if result[dk]["before"] is not None and result[dk]["after"] is not None:
                continue
            b, a = _ocr_try_table_row_for_aliases(line, aliases[dk])
            if b is not None and a is not None:
                result[dk]["before"] = b
                result[dk]["after"] = a
    # Pass 1: same line, then 2-line window (avoid huge multi-line windows)
    for i, line in enumerate(lines):
        ll = line.lower()
        win2 = "\n".join(lines[i : min(i + 2, len(lines))])
        for dk in ordered_keys:
            if result[dk]["before"] is not None and result[dk]["after"] is not None:
                continue
            if not any(a in ll for a in aliases[dk]):
                continue
            b, a = _try_before_after_pair_from_window(line)
            if b is None or a is None:
                b, a = _try_before_after_pair_from_window(win2)
            if b is not None and a is not None:
                result[dk]["before"] = b
                result[dk]["after"] = a
                break
    # Pass 2: short slice after metric in flattened text (limit cross-row bleed)
    flat = " " + _ocr_flat_text(left).lower() + " "
    for dk in ordered_keys:
        if result[dk]["before"] is not None and result[dk]["after"] is not None:
            continue
        best = -1
        for a in sorted(aliases[dk], key=len, reverse=True):
            pos = flat.find(" " + a + " ")
            if pos < 0:
                pos = flat.find(a)
            if pos >= 0 and (best < 0 or pos < best):
                best = pos
        if best < 0:
            continue
        window = flat[best : min(len(flat), best + 220)]
        b, a = _try_before_after_pair_from_window(window)
        if b is not None and a is not None:
            result[dk]["before"] = b
            result[dk]["after"] = a
    _cy_fill_from_row_order_if_empty(left, result)
    return result


def _parse_py_matrix_from_ocr_text(text: str) -> dict[str, dict[str, float | None]]:
    """paid_users | spending | active_listers -> before/after from OCR (matrix keys semantic)."""
    cy_al = _get_ocr_cy_aliases()
    aliases = {
        "paid_users": cy_al["paid_users"],
        "spending": cy_al["spending"],
        "active_listers": cy_al["active_listers"],
    }
    result = {k: {"before": None, "after": None} for k in aliases}
    left = _ocr_left_panel_pp_only(text)
    pp = _parse_pp_dashboard_metric_rows(left, aliases, list(aliases.keys()))
    for dk in result:
        p = pp.get(dk) or {}
        if p.get("before") is not None and p.get("after") is not None:
            result[dk] = {"before": p["before"], "after": p["after"]}

    lines = [ln.strip() for ln in left.splitlines() if ln.strip()]
    ordered = sorted(aliases.keys(), key=lambda k: max(len(x) for x in aliases[k]), reverse=True)
    for line in lines:
        for dk in ordered:
            if result[dk]["before"] is not None and result[dk]["after"] is not None:
                continue
            b, a = _ocr_try_table_row_for_aliases(line, aliases[dk])
            if b is not None and a is not None:
                result[dk]["before"] = b
                result[dk]["after"] = a
    for i, line in enumerate(lines):
        ll = line.lower()
        win2 = "\n".join(lines[i : min(i + 2, len(lines))])
        for dk in ordered:
            if result[dk]["before"] is not None and result[dk]["after"] is not None:
                continue
            if not any(a in ll for a in aliases[dk]):
                continue
            b, a = _try_before_after_pair_from_window(line)
            if b is None or a is None:
                b, a = _try_before_after_pair_from_window(win2)
            if b is not None and a is not None:
                result[dk]["before"] = b
                result[dk]["after"] = a
                break
    _py_fill_from_row_order_if_empty(left, text, result)
    return result


def _py_first_numeric_pair_in_period_group_segment(segment: str) -> tuple[float | None, float | None]:
    """First Default|Target-like row in a slice of OCR (skips header lines)."""
    lines_in = [raw.strip() for raw in segment.splitlines() if raw.strip()]
    lines: list[str] = []
    for line in lines_in:
        ll = line.lower()
        if "period group" in ll:
            continue
        if "default" in ll and "target" in ll and not re.search(r"\d", line):
            continue
        lines.append(line)
    for b, a in _pp_collect_numeric_row_pairs(lines):
        if b is None or a is None:
            continue
        try:
            bf, af = float(b), float(a)
        except (TypeError, ValueError):
            continue
        if bf < 0 or af < 0 or not math.isfinite(bf) or not math.isfinite(af):
            continue
        if max(bf, af) > 10**12:
            continue
        return b, a
    return None, None


def _py_period_group_default_target_pair(
    full_text: str,
    skip_if_matches: tuple[float | None, float | None] | None = None,
) -> tuple[float | None, float | None]:
    """
    OCR often has **two** \"Period Group\" headers: the first introduces the main metric block
    (first data row = Paid users), the second introduces a small table that is **Active listers**.

    We collect the first numeric pair after **each** \"Period Group\", then prefer the **last**
    pair that is not identical to Paid users (already filled from row 0).
    """
    if not full_text or not str(full_text).strip():
        return None, None
    low = full_text.lower()
    key = "period group"
    kl = len(key)
    candidates: list[tuple[float, float]] = []
    start = 0
    while True:
        pos = low.find(key, start)
        if pos < 0:
            break
        nxt = low.find(key, pos + kl)
        segment = full_text[pos + kl : nxt] if nxt >= 0 else full_text[pos + kl :]
        b, a = _py_first_numeric_pair_in_period_group_segment(segment)
        if b is not None and a is not None:
            candidates.append((b, a))
        start = pos + kl

    if not candidates:
        return None, None

    def _same(
        x: tuple[float, float],
        y: tuple[float | None, float | None] | None,
    ) -> bool:
        if y is None or y[0] is None or y[1] is None:
            return False
        return abs(x[0] - float(y[0])) < 0.51 and abs(x[1] - float(y[1])) < 0.51

    for cand in reversed(candidates):
        if not _same(cand, skip_if_matches):
            return cand[0], cand[1]
    return None, None


def _py_fill_from_row_order_if_empty(
    left: str,
    full_text: str,
    result: dict[str, dict[str, float | None]],
) -> None:
    """
    Positional fallback for the PPV left table: row 1 → paid, row 6 → spending, row 12 → active.

    Fills **only** metrics that are still missing after label-based parsing.

    ``left`` is cropped for the main panel; ``full_text`` is the raw OCR string — needed when
    `_ocr_left_panel_pp_only` cuts below `diff (median)` and drops the **Period Group** block where
    Active listers often live on tight crops.
    """
    rows_left = _cy_collect_pp_numeric_rows(left)
    rows_full = _cy_collect_pp_numeric_rows(full_text) if full_text else []

    def _need(k: str) -> bool:
        p = result.get(k) or {}
        return p.get("before") is None or p.get("after") is None

    if len(rows_left) >= 1 and _need("paid_users"):
        b0, a0 = rows_left[0]
        result["paid_users"] = {"before": b0, "after": a0}
    if len(rows_left) >= 6 and _need("spending"):
        b5, a5 = rows_left[5]
        result["spending"] = {"before": b5, "after": a5}

    if not _need("active_listers"):
        return

    pu = result.get("paid_users") or {}
    skip_pg: tuple[float | None, float | None] | None = None
    if pu.get("before") is not None and pu.get("after") is not None:
        try:
            skip_pg = (float(pu["before"]), float(pu["after"]))
        except (TypeError, ValueError):
            skip_pg = None
    pg_b, pg_a = _py_period_group_default_target_pair(full_text, skip_if_matches=skip_pg)
    if pg_b is not None and pg_a is not None:
        result["active_listers"] = {"before": pg_b, "after": pg_a}
        return

    if len(rows_left) >= 12:
        b11, a11 = rows_left[11]
        result["active_listers"] = {"before": b11, "after": a11}
        return

    if len(rows_full) >= 12:
        b11, a11 = rows_full[11]
        result["active_listers"] = {"before": b11, "after": a11}
        return


def _cy_pair_semantically_plausible(dk: str, typ: str, b, a) -> bool:
    """Reject glued OCR garbage (e.g. Campaign per User after ~1e10)."""
    try:
        bf, af = float(b), float(a)
    except (TypeError, ValueError):
        return False
    if bf < 0 or af < 0:
        return False
    mx = max(bf, af)
    if dk == "campaign_per_user":
        return mx <= 500.0
    if dk in ("plan_imp_per_campaign", "fact_imp_per_campaign"):
        return mx <= 500_000.0
    if dk == "price_per_day":
        return mx <= 1_000_000.0
    if dk in ("paid_users", "new_campaign_cnt", "active_listers"):
        return mx <= 10**12
    if dk in ("spending", "refund"):
        return mx <= 10**15
    if dk == "arp_p_campaign":
        return mx <= 10**7
    if "pct_" in dk or dk == "pct_execution_inventory":
        return mx <= 10**6
    return mx <= 10**18


def _cy_plausible_pair_count(parsed: dict[str, dict[str, float | None]]) -> int:
    n = 0
    for dk, _, typ in _CY_INPUT_METRICS:
        pb = (parsed.get(dk) or {}).get("before")
        pa = (parsed.get(dk) or {}).get("after")
        if pb is None or pa is None:
            continue
        if _cy_pair_semantically_plausible(dk, typ, pb, pa):
            n += 1
    return n


def _merge_cy_parsed_from_ocr_variants(texts: list[str]) -> dict[str, dict[str, float | None]]:
    """
    Pick the single best OCR variant (not field-wise first-wins): a noisy PSM can
    poison metrics like Campaign per User if merged with a cleaner pass.
    """
    empty: dict[str, dict[str, float | None]] = {
        dk: {"before": None, "after": None} for dk, _, _ in _CY_INPUT_METRICS
    }
    if not texts:
        return empty
    scored: list[tuple[tuple[int, int, int], dict[str, dict[str, float | None]]]] = []
    for t in texts:
        p = _parse_cy_metrics_from_ocr_text(t)
        good = _cy_plausible_pair_count(p)
        total = sum(
            1
            for dk, _, _ in _CY_INPUT_METRICS
            if (p.get(dk) or {}).get("before") is not None
            and (p.get(dk) or {}).get("after") is not None
        )
        scored.append(((good, total, len(t)), p))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[0][1]


def _py_pair_semantically_plausible(k: str, b, a) -> bool:
    try:
        bf, af = float(b), float(a)
    except (TypeError, ValueError):
        return False
    if bf < 0 or af < 0:
        return False
    mx = max(bf, af)
    if k == "paid_users":
        return mx <= 10**12
    if k == "spending":
        return mx <= 10**15
    if k == "active_listers":
        return mx <= 10**12
    return False


def _py_plausible_pair_count(parsed: dict[str, dict[str, float | None]]) -> int:
    n = 0
    for k in ("paid_users", "spending", "active_listers"):
        pb = (parsed.get(k) or {}).get("before")
        pa = (parsed.get(k) or {}).get("after")
        if pb is None or pa is None:
            continue
        if _py_pair_semantically_plausible(k, pb, pa):
            n += 1
    return n


def _merge_py_parsed_from_ocr_variants(texts: list[str]) -> dict[str, dict[str, float | None]]:
    keys = ("paid_users", "spending", "active_listers")
    empty = {k: {"before": None, "after": None} for k in keys}
    if not texts:
        return empty
    scored: list[tuple[tuple[int, int, int], dict]] = []
    for t in texts:
        p = _parse_py_matrix_from_ocr_text(t)
        good = _py_plausible_pair_count(p)
        total = sum(
            1
            for k in keys
            if (p.get(k) or {}).get("before") is not None
            and (p.get(k) or {}).get("after") is not None
        )
        scored.append(((good, total, len(t)), p))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[0][1]


def _cy_count_metric_pairs(parsed: dict[str, dict[str, float | None]]) -> int:
    return sum(
        1
        for dk, _, _ in _CY_INPUT_METRICS
        if (parsed.get(dk) or {}).get("before") is not None
        and (parsed.get(dk) or {}).get("after") is not None
    )


def _ocr_pick_preview_text_cy(variants: list[str]) -> str:
    """Debug preview: variant with the most semantically plausible pairs, then longest."""
    best = variants[0]
    best_n = (-1, -1)
    for t in variants:
        p = _parse_cy_metrics_from_ocr_text(t)
        n = (_cy_plausible_pair_count(p), len(t))
        if n > best_n:
            best_n = n
            best = t
    return best


def _ocr_pick_preview_text_py(variants: list[str]) -> str:
    def score(txt: str) -> int:
        p = _parse_py_matrix_from_ocr_text(txt)
        return sum(
            1
            for k in ("paid_users", "spending", "active_listers")
            if (p.get(k) or {}).get("before") is not None
            and (p.get(k) or {}).get("after") is not None
        )

    best = variants[0]
    best_n = -1
    for t in variants:
        n = score(t)
        if n > best_n or (n == best_n and len(t) > len(best)):
            best_n = n
            best = t
    return best


def _ocr_collect_tesseract_variants(image_bytes: bytes) -> tuple[list[str] | None, str | None]:
    """
    Run Tesseract with several PSM modes on a lightly upscaled image (helps small table text).
    Returns (unique non-empty texts, error).
    """
    try:
        image_module = importlib.import_module("PIL.Image")
        pytesseract = importlib.import_module("pytesseract")
    except Exception:
        return None, "OCR недоступен: установите `pillow` и `pytesseract`, затем перезапустите приложение."

    Image = image_module
    img = Image.open(io.BytesIO(image_bytes))
    try:
        resample = Image.Resampling.LANCZOS
    except AttributeError:
        resample = Image.LANCZOS  # type: ignore[attr-defined]
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    w, h = img.size
    mx = max(w, h)
    if mx > 0 and mx < 1400:
        s = min(2.0, 2400.0 / mx)
        img = img.resize((max(1, int(w * s)), max(1, int(h * s))), resample)

    cfgs = (
        "--oem 3 --psm 6",
        "--oem 3 --psm 4",
        "--oem 3 --psm 11",
        "--oem 3 --psm 3",
    )
    seen: set[str] = set()
    out: list[str] = []
    for cfg in cfgs:
        try:
            t = pytesseract.image_to_string(img, config=cfg)
        except Exception:
            continue
        if not t or not str(t).strip():
            continue
        key = str(t).strip()
        if key not in seen:
            seen.add(key)
            out.append(str(t))
    if not out:
        return None, "Не удалось распознать текст на скриншоте."
    return out, None


def _ocr_bytes_to_text(image_bytes: bytes):
    """Single preview string (best PSM by parse score); prefer _ocr_collect_tesseract_variants + merge in UI."""
    variants, err = _ocr_collect_tesseract_variants(image_bytes)
    if err or not variants:
        return None, err
    return _ocr_pick_preview_text_cy(variants), None


def _ocr_screenshot(uploaded_file):
    return _ocr_bytes_to_text(uploaded_file.getvalue())


def _ocr_text_looks_like_app_chrome(text: str) -> bool:
    """Heuristic: OCR of this Streamlit app (wrong upload) vs source dashboard."""
    t = (text or "").lower()
    hits = 0
    for needle in (
        "current year input",
        "manual override",
        "ocr debug",
        "parsed before",
        "session_state",
        "previous year screenshot",
        "распознать и применить",
        "скрин с данными",
        "single analysis",
    ):
        if needle in t:
            hits += 1
    return hits >= 2


def _session_write_matrix_py_triple_from_parsed(
    sess: dict, parsed: dict[str, dict[str, float | None]]
) -> None:
    """
    Paid users / Spending / Active listers from CY (or PY) parse → matrix_py_* keys
    used by PPV matrix & Potential Spending.
    """
    mapping = [
        ("paid_users", "matrix_py_paid_users_before", "matrix_py_paid_users_after", True),
        ("spending", "matrix_py_spending_before", "matrix_py_spending_after", False),
        ("active_listers", "matrix_py_ac_before", "matrix_py_ac_after", True),
    ]
    for dk, kb, ka, as_int in mapping:
        pair = parsed.get(dk) or {}
        b, a = pair.get("before"), pair.get("after")
        if b is None or a is None:
            continue
        try:
            if as_int:
                bi = int(round(float(b)))
                ai = int(round(float(a)))
                if ai < 0 and bi >= 0:
                    ai = abs(ai)
                if bi < 0:
                    bi = abs(bi)
                if bi < 0 or ai < 0:
                    continue
                sess[kb] = bi
                sess[ka] = ai
            else:
                bf = float(b)
                af = float(a)
                if bf < 0 and abs(bf) <= 1e6:
                    bf = abs(bf)
                if af < 0 and bf >= 0 and abs(af) <= 1e6:
                    af = abs(af)
                if bf < 0 or af < 0:
                    continue
                sess[kb] = bf
                sess[ka] = af
        except (TypeError, ValueError, OverflowError):
            continue


def _cy_ocr_snapshot_persist(sess: dict) -> None:
    """Copy all Current Year widget keys so we can restore after a bad file-merge overwrite."""
    snap: dict = {}
    for _dk, _, _ in _CY_INPUT_METRICS:
        sk = _cy_sess_key(_dk)
        for part in ("baseline", "before", "after"):
            k = f"{sk}_{part}"
            if k in sess:
                snap[k] = sess[k]
    sess["_cy_ocr_snapshot"] = snap


def _cy_ocr_restore_from_snapshot_if_blanked(sess: dict) -> None:
    """
    If CY OCR override is on but merge zeroed inputs (e.g. old sig!=prev_cat cleared override),
    restore from last CY OCR snapshot.
    """
    if not sess.get("_cy_ocr_override"):
        return
    snap = sess.get("_cy_ocr_snapshot")
    if not isinstance(snap, dict) or not snap:
        return
    anchors = ("paid_users_before", "spending_before", "active_before")
    try:
        snap_any = any(float(snap.get(k) or 0) != 0 for k in anchors)
        cur_any = any(float(sess.get(k) or 0) != 0 for k in anchors)
    except (TypeError, ValueError):
        return
    if cur_any or not snap_any:
        return
    for k, v in snap.items():
        sess[k] = v


def _apply_cy_ocr_parsed_to_session(
    sess: dict,
    parsed: dict[str, dict[str, float | None]],
    *,
    mirror_matrix_py: bool = False,
) -> tuple[list[str], list[dict]]:
    """
    Write into Current Year input keys: {sess_key}_before / {sess_key}_after
    (sess_key = _cy_sess_key(data_key), e.g. paid_users_before, active_before for active_listers).

    If mirror_matrix_py, also copy paid_users / spending / active_listers into matrix_py_*.
    The CY OCR button passes mirror_matrix_py=False so the matrix "Previous Year" row stays
    independent (PY files / PY OCR only).

    Returns (applied_metric_labels, debug_rows) for UI verification.
    """
    applied: list[str] = []
    rows: list[dict] = []
    for dk, label, typ in _CY_INPUT_METRICS:
        sk = _cy_sess_key(dk)
        kb = f"{sk}_before"
        ka = f"{sk}_after"
        pair = parsed.get(dk) or {}
        b, a = pair.get("before"), pair.get("after")
        row = {
            "Metric": label,
            "Parsed Before": b,
            "Parsed After": a,
            "session_state key (Before)": kb,
            "session_state key (After)": ka,
            "Written": False,
            "Stored Before": None,
            "Stored After": None,
        }
        if b is None or a is None:
            rows.append(row)
            continue
        try:
            if typ == "int":
                bi = int(round(float(b)))
                ai = int(round(float(a)))
                if ai < 0 and bi >= 0:
                    ai = abs(ai)
                if bi < 0:
                    bi = abs(bi)
                if bi < 0 or ai < 0:
                    rows.append(row)
                    continue
                sess[kb] = bi
                sess[ka] = ai
                row["Written"] = True
                row["Stored Before"] = bi
                row["Stored After"] = ai
                applied.append(label)
            else:
                bf = float(b)
                af = float(a)
                # OCR often misreads a minus on small "After" levels; keep non-negative for widgets (min 0).
                if bf < 0 and abs(bf) <= 1e6:
                    bf = abs(bf)
                if af < 0 and bf >= 0 and abs(af) <= 1e6:
                    af = abs(af)
                if bf < 0 or af < 0:
                    rows.append(row)
                    continue
                sess[kb] = bf
                sess[ka] = af
                row["Written"] = True
                row["Stored Before"] = bf
                row["Stored After"] = af
                applied.append(label)
        except (TypeError, ValueError, OverflowError):
            pass
        rows.append(row)
    if mirror_matrix_py:
        _session_write_matrix_py_triple_from_parsed(sess, parsed)
    if applied:
        _cy_ocr_snapshot_persist(sess)
    return applied, rows


def _apply_py_ocr_parsed_to_session(
    sess: dict, parsed: dict[str, dict[str, float | None]]
) -> tuple[list[str], list[dict]]:
    """matrix_py_paid_users_*, matrix_py_spending_*, matrix_py_ac_*."""
    applied: list[str] = []
    rows: list[dict] = []
    mapping = [
        ("paid_users", "matrix_py_paid_users_before", "matrix_py_paid_users_after", "Paid users", True),
        ("spending", "matrix_py_spending_before", "matrix_py_spending_after", "Spending", False),
        ("active_listers", "matrix_py_ac_before", "matrix_py_ac_after", "Active listers", True),
    ]
    for dk, kb, ka, lbl, as_int in mapping:
        pair = parsed.get(dk) or {}
        b, a = pair.get("before"), pair.get("after")
        row = {
            "Metric": lbl,
            "Parsed Before": b,
            "Parsed After": a,
            "session_state key (Before)": kb,
            "session_state key (After)": ka,
            "Written": False,
            "Stored Before": None,
            "Stored After": None,
        }
        if b is None or a is None:
            rows.append(row)
            continue
        try:
            if as_int:
                bi = int(round(float(b)))
                ai = int(round(float(a)))
                if ai < 0 and bi >= 0:
                    ai = abs(ai)
                if bi < 0:
                    bi = abs(bi)
                if bi < 0 or ai < 0:
                    rows.append(row)
                    continue
                sess[kb] = bi
                sess[ka] = ai
                row["Written"] = True
                row["Stored Before"] = bi
                row["Stored After"] = ai
                applied.append(lbl)
            else:
                bf = float(b)
                af = float(a)
                if bf < 0 and abs(bf) <= 1e6:
                    bf = abs(bf)
                if af < 0 and bf >= 0 and abs(af) <= 1e6:
                    af = abs(af)
                if bf < 0 or af < 0:
                    rows.append(row)
                    continue
                sess[kb] = bf
                sess[ka] = af
                row["Written"] = True
                row["Stored Before"] = bf
                row["Stored After"] = af
                applied.append(lbl)
        except (TypeError, ValueError, OverflowError):
            pass
        rows.append(row)
    return applied, rows


def _cy_ocr_report_labels(parsed: dict[str, dict[str, float | None]], kind: str) -> tuple[list[str], list[str]]:
    """Found (both before/after) vs missing metric names."""
    found = []
    missing = []
    if kind == "cy":
        for dk, label, _ in _CY_INPUT_METRICS:
            p = parsed.get(dk) or {}
            if p.get("before") is not None and p.get("after") is not None:
                found.append(label)
            else:
                missing.append(label)
    else:
        for dk, label in (
            ("paid_users", "Paid users"),
            ("spending", "Spending"),
            ("active_listers", "Active Listers"),
        ):
            p = parsed.get(dk) or {}
            if p.get("before") is not None and p.get("after") is not None:
                found.append(label)
            else:
                missing.append(label)
    return found, missing


def _build_cy_ocr_feedback(
    ocr_text: str | None,
    ocr_error: str | None,
    parsed: dict[str, dict[str, float | None]],
    applied_labels: list[str],
    session_debug_rows: list[dict] | None = None,
) -> dict:
    """Structured OCR result for UI + debug (includes per-metric session_state keys)."""
    recognized: list[str] = []
    for dk, lbl, _ in _CY_INPUT_METRICS:
        p = parsed.get(dk) or {}
        if p.get("before") is not None and p.get("after") is not None:
            recognized.append(lbl)
    all_labels = [lbl for _, lbl, _ in _CY_INPUT_METRICS]
    unmatched = [lbl for lbl in all_labels if lbl not in recognized]
    skipped_apply = [lbl for lbl in recognized if lbl not in applied_labels]
    ap = len(applied_labels)
    sk = len(skipped_apply)
    return {
        "ocr_error": ocr_error,
        "preview": (ocr_text or "")[:1000],
        "recognized_in_ocr": recognized,
        "applied": list(applied_labels),
        "unmatched": unmatched,
        "skipped_apply": skipped_apply,
        "counts": {"applied": ap, "skipped": sk, "recognized": len(recognized)},
        "session_debug_rows": session_debug_rows or [],
    }


def _build_py_ocr_feedback(
    ocr_text: str | None,
    ocr_error: str | None,
    parsed: dict[str, dict[str, float | None]],
    applied_labels: list[str],
    session_debug_rows: list[dict] | None = None,
) -> dict:
    recognized = []
    for dk, lbl in (
        ("paid_users", "Paid users"),
        ("spending", "Spending"),
        ("active_listers", "Active Listers"),
    ):
        p = parsed.get(dk) or {}
        if p.get("before") is not None and p.get("after") is not None:
            recognized.append(lbl)
    all_l = ["Paid users", "Spending", "Active Listers"]
    unmatched = [x for x in all_l if x not in recognized]
    skipped_apply = [x for x in recognized if x not in applied_labels]
    return {
        "ocr_error": ocr_error,
        "preview": (ocr_text or "")[:1000],
        "recognized_in_ocr": recognized,
        "applied": list(applied_labels),
        "unmatched": unmatched,
        "skipped_apply": skipped_apply,
        "counts": {
            "applied": len(applied_labels),
            "skipped": len(skipped_apply),
            "recognized": len(recognized),
        },
        "session_debug_rows": session_debug_rows or [],
    }


_OCR_UI_PREVIEW_MAX_H = 320


def _ocr_thumbnail_bytes(image_bytes: bytes, max_height: int = _OCR_UI_PREVIEW_MAX_H) -> bytes | None:
    """Уменьшенное изображение для превью (макс. высота max_height px)."""
    try:
        image_module = importlib.import_module("PIL.Image")
        Image = image_module
        img = Image.open(io.BytesIO(image_bytes))
        img = img.convert("RGB")
        try:
            resample = Image.Resampling.LANCZOS
        except AttributeError:
            resample = Image.LANCZOS  # type: ignore[attr-defined]
        img.thumbnail((100_000, max_height), resample)
        out = io.BytesIO()
        img.save(out, format="PNG", optimize=True)
        return out.getvalue()
    except Exception:
        return None


def _ocr_render_screenshot_preview(image_bytes: bytes | None, *, label: str) -> None:
    """Превью ограниченной высоты + полный размер в popover или expander."""
    if not image_bytes:
        return
    thumb = _ocr_thumbnail_bytes(image_bytes)
    if thumb:
        st.image(
            io.BytesIO(thumb),
            caption=f"Предпросмотр ({label}) — до {_OCR_UI_PREVIEW_MAX_H}px по высоте",
            use_container_width=False,
        )
    else:
        st.image(
            io.BytesIO(image_bytes),
            caption=f"Предпросмотр ({label})",
            width=min(480, 10_000),
        )
    _full = io.BytesIO(image_bytes)
    if getattr(st, "popover", None):
        with st.popover("🔍 Открыть полный размер"):
            st.image(_full, use_container_width=True, caption=f"{label} — полный размер")
    else:
        with st.expander("🔍 Открыть полный размер", expanded=False):
            st.image(_full, use_container_width=True, caption=f"{label} — полный размер")


def _ocr_debug_rows_format_numeric_strings(rows: list[dict]) -> list[dict]:
    """
    Stringify Parsed/Stored numeric cells for st.dataframe so positives never show a leading '+'.
    Uses format_matrix_metric (same rules as PPV matrix int-like vs decimals).
    """
    cols = ("Parsed Before", "Parsed After", "Stored Before", "Stored After")
    out: list[dict] = []
    for r in rows:
        r2 = dict(r)
        for c in cols:
            if c not in r2 or r2[c] is None:
                continue
            v = r2[c]
            try:
                fv = float(v)
            except (TypeError, ValueError):
                continue
            tol = max(1e-9, 1e-9 * max(1.0, abs(fv)))
            as_int = abs(fv - round(fv)) < tol
            r2[c] = format_matrix_metric(fv, as_int=as_int)
        out.append(r2)
    return out


def _render_ocr_feedback_messages(fb: dict | None) -> None:
    """Сообщения после кнопки: error / warning / success / info (нативные цвета Streamlit)."""
    if not fb:
        return
    err = fb.get("ocr_error")
    if err:
        st.error(f"**OCR / input error:** {err}")
        return

    rec = fb.get("recognized_in_ocr") or []
    ap = fb.get("applied") or []
    unmatched = fb.get("unmatched") or []
    skipped = fb.get("skipped_apply") or []
    preview = fb.get("preview") or ""

    if not rec:
        st.warning("OCR completed, but no supported metrics were recognized.")
        if preview and _ocr_text_looks_like_app_chrome(preview):
            st.info(
                "Похоже, в кадр попал **интерфейс этого приложения** (Streamlit), а не исходная таблица "
                "с метриками (например, блок **New PPV (spending)** с колонками **Default** / **Target**). "
                "Сделайте скрин **только дашборда/отчёта** или обрежьте изображение до левой таблицы, "
                "без панели инструмента."
            )
    else:
        if ap:
            st.success(
                "**Applied to session_state:** "
                + ", ".join(f"**{x}**" for x in ap)
            )
            st.caption(
                f"Applied **{len(ap)}** metric(s), skipped **{len(skipped)}** "
                "(recognized in OCR but not written to session_state)."
            )
        if skipped:
            st.warning(
                "**Recognized in OCR but not applied** (validation failed or negative values): "
                + ", ".join(skipped)
            )

    st.markdown("**Unmatched** (no Before/After pair found for this label)")
    with st.container(border=True):
        st.write(", ".join(unmatched) if unmatched else "—")


def _render_ocr_debug_expanders(fb: dict | None, kind: str) -> None:
    """Сырой текст и таблица пар — один свёрнутый expander."""
    if not fb:
        return
    preview = fb.get("preview") or ""
    dbg_rows = fb.get("session_debug_rows") or []

    with st.expander("OCR debug", expanded=False):
        st.markdown("**Raw text** (first 1000 chars)")
        st.code(preview if preview.strip() else "(empty)", language=None)
        if dbg_rows:
            st.markdown("**Parsed pairs & session_state writes**")
            _dbg_df = pd.DataFrame(_ocr_debug_rows_format_numeric_strings(dbg_rows))
            st.dataframe(_dbg_df, use_container_width=True, hide_index=True)
            if st.session_state.get("_cy_ocr_override") and kind == "cy":
                st.caption(
                    "**Overwrite protection:** `_cy_ocr_override` is set — file-based Current Year "
                    "merge will not overwrite these values until you **re-upload CY files** "
                    "(sets `_merge_files_dirty`) or **change category**."
                )
            if st.session_state.get("_py_ocr_override") and kind == "py":
                st.caption(
                    "**Overwrite protection:** `_py_ocr_override` is set — file-based Previous Year "
                    "merge will not overwrite matrix fields until **re-upload PY files** or "
                    "**change category**."
                )


def _engine_style_diff(before, after):
    """Match decision_engine.diff: percent change; 0 if before == 0."""
    try:
        b = float(before)
        a = float(after)
    except (TypeError, ValueError):
        return None
    if b == 0:
        return 0.0
    return (a - b) / b * 100.0


def _bulk_safe_ratio(npl, active):
    if active is None or float(active) == 0:
        return 0.0
    return float(npl) / float(active)


_BULK_PRIORITY_MISSING = "0 - Missing data"
_BULK_PRIORITY_NEGATIVE = "1 - Negative impact"
_BULK_PRIORITY_INSUFFICIENT = "2 - Insufficient data"
_BULK_PRIORITY_NEED_REVIEW = "3 - Need review"
_BULK_PRIORITY_NO_IMPACT = "4 - No impact"
_BULK_PRIORITY_POSITIVE = "5 - Positive impact"

_BULK_PRIORITY_SORT_ORDER = {
    _BULK_PRIORITY_MISSING: 0,
    _BULK_PRIORITY_NEGATIVE: 1,
    _BULK_PRIORITY_INSUFFICIENT: 2,
    _BULK_PRIORITY_NEED_REVIEW: 3,
    _BULK_PRIORITY_NO_IMPACT: 4,
    _BULK_PRIORITY_POSITIVE: 5,
}


def _bulk_row_priority(status: str, final_decision):
    if status == "Missing current year data":
        return _BULK_PRIORITY_MISSING
    fd = final_decision if final_decision is not None else ""
    _by_final = {
        "Negative impact": _BULK_PRIORITY_NEGATIVE,
        "Insufficient data": _BULK_PRIORITY_INSUFFICIENT,
        "Need review": _BULK_PRIORITY_NEED_REVIEW,
        "No impact": _BULK_PRIORITY_NO_IMPACT,
        "Positive impact": _BULK_PRIORITY_POSITIVE,
    }
    return _by_final.get(fd, _BULK_PRIORITY_NEED_REVIEW)


def _bulk_row_warning_flags(status: str, final_decision, has_py: bool, is_other_category: bool) -> str:
    flags = []
    if final_decision == "Insufficient data":
        flags.append("LOW_PAID_USERS")
    if status == "Missing current year data":
        flags.append("MISSING_CY")
    if not has_py:
        flags.append("MISSING_PY")
    if is_other_category:
        flags.append("OTHER_CATEGORY")
    return ", ".join(flags)


def _bulk_analysis_dataframe(
    parsed_category_ids,
    merged_data,
    merged_data_previous_year,
    geo: str,
    force_low_npl: bool,
    is_other_category: bool,
):
    """One row per category_id; CY analyze via analyze_category; PY/Y2Y control diffs."""
    rows = []
    for cid in parsed_category_ids:
        if cid not in merged_data:
            _st = "Missing current year data"
            _has_py = bool(merged_data_previous_year) and cid in merged_data_previous_year
            rows.append(
                {
                    "category_id": cid,
                    "priority": _bulk_row_priority(_st, None),
                    "warning_flags": _bulk_row_warning_flags(
                        _st, None, _has_py, is_other_category
                    ),
                    "status": _st,
                    "cy_paid_users_diff": None,
                    "cy_spending_diff": None,
                    "cy_cr_diff": None,
                    "decision_code": None,
                    "final_decision": None,
                    "next_step": None,
                    "py_paid_users_diff": None,
                    "py_spending_diff": None,
                    "py_cr_diff": None,
                    "y2y_paid_users_diff": None,
                    "y2y_spending_diff": None,
                    "y2y_cr_diff": None,
                }
            )
            continue

        data = merged_data[cid]
        b = data.get("before") or {}
        a = data.get("after") or {}
        npl_b = int(float(b.get("paid_users") or 0))
        npl_a = int(float(a.get("paid_users") or 0))
        sp_b = float(b.get("spending") or 0)
        sp_a = float(a.get("spending") or 0)
        ac_b = int(float(b.get("active_listers") or 0))
        ac_a = int(float(a.get("active_listers") or 0))

        py_paid_users_diff = None
        py_spending_diff = None
        py_cr_diff = None
        if merged_data_previous_year and cid in merged_data_previous_year:
            pd = merged_data_previous_year[cid]
            pb = pd.get("before") or {}
            pa = pd.get("after") or {}
            pnpl_b = int(float(pb.get("paid_users") or 0))
            pnpl_a = int(float(pa.get("paid_users") or 0))
            psp_b = float(pb.get("spending") or 0)
            psp_a = float(pa.get("spending") or 0)
            pac_b = int(float(pb.get("active_listers") or 0))
            pac_a = int(float(pa.get("active_listers") or 0))
            py_paid_users_diff = _engine_style_diff(pnpl_b, pnpl_a)
            py_spending_diff = _engine_style_diff(psp_b, psp_a)
            pcr_b = _bulk_safe_ratio(pnpl_b, pac_b)
            pcr_a = _bulk_safe_ratio(pnpl_a, pac_a)
            py_cr_diff = _engine_style_diff(pcr_b, pcr_a)

        result = analyze_category(
            npl_before=npl_b,
            npl_after=npl_a,
            sp_before=sp_b,
            sp_after=sp_a,
            active_before=ac_b,
            active_after=ac_a,
            geo=geo or "default",
            force_low_npl=force_low_npl,
            is_other_category=is_other_category,
        )
        cy_paid_users_diff = result["npl_diff"]
        cy_spending_diff = result["sp_diff"]
        cy_cr = result["cr_diff"]

        def _y2y(cy_v, py_v):
            if py_v is None or cy_v is None:
                return None
            return cy_v - py_v

        _st_ok = ""
        _fd = result["final_decision"]
        _has_py = bool(merged_data_previous_year) and cid in merged_data_previous_year
        rows.append(
            {
                "category_id": cid,
                "priority": _bulk_row_priority(_st_ok, _fd),
                "warning_flags": _bulk_row_warning_flags(
                    _st_ok, _fd, _has_py, is_other_category
                ),
                "status": _st_ok,
                "cy_paid_users_diff": cy_paid_users_diff,
                "cy_spending_diff": cy_spending_diff,
                "cy_cr_diff": cy_cr,
                "decision_code": result["decision_code"],
                "final_decision": _fd,
                "next_step": result["next_step"],
                "py_paid_users_diff": py_paid_users_diff,
                "py_spending_diff": py_spending_diff,
                "py_cr_diff": py_cr_diff,
                "y2y_paid_users_diff": _y2y(cy_paid_users_diff, py_paid_users_diff),
                "y2y_spending_diff": _y2y(cy_spending_diff, py_spending_diff),
                "y2y_cr_diff": _y2y(cy_cr, py_cr_diff),
            }
        )

    cols = [
        "category_id",
        "priority",
        "warning_flags",
        "status",
        "decision_code",
        "final_decision",
        "next_step",
        "cy_paid_users_diff",
        "cy_spending_diff",
        "cy_cr_diff",
        "py_paid_users_diff",
        "py_spending_diff",
        "py_cr_diff",
        "y2y_paid_users_diff",
        "y2y_spending_diff",
        "y2y_cr_diff",
    ]
    df = pd.DataFrame(rows, columns=cols)
    df["_pri_sort"] = df["priority"].map(lambda p: _BULK_PRIORITY_SORT_ORDER.get(p, 99))
    df = df.sort_values(by=["_pri_sort", "category_id"], ascending=[True, True]).drop(
        columns=["_pri_sort"]
    )
    return df


def _bulk_render_summary(df: pd.DataFrame) -> None:
    """Manager-facing counts: six buckets from priority + breakdowns by priority / final_decision / status."""
    if df is None or df.empty:
        return

    pri_vc = df["priority"].value_counts()
    _metric_rows = (
        ("Missing data", _BULK_PRIORITY_MISSING),
        ("Negative impact", _BULK_PRIORITY_NEGATIVE),
        ("Insufficient data", _BULK_PRIORITY_INSUFFICIENT),
        ("Need review", _BULK_PRIORITY_NEED_REVIEW),
        ("No impact", _BULK_PRIORITY_NO_IMPACT),
        ("Positive impact", _BULK_PRIORITY_POSITIVE),
    )

    st.markdown("##### Summary")
    _r1, _r2 = st.columns(3), st.columns(3)
    for _i, (_title, _pkey) in enumerate(_metric_rows):
        _cols = _r1 if _i < 3 else _r2
        with _cols[_i % 3]:
            st.metric(_title, int(pri_vc.get(_pkey, 0)))

    with st.expander("Разбивка по priority, final_decision, status", expanded=False):
        _e1, _e2, _e3 = st.columns(3)
        with _e1:
            st.caption("По priority")
            st.dataframe(
                pri_vc.rename_axis("priority").reset_index(name="n"),
                hide_index=True,
                use_container_width=True,
            )
        with _e2:
            st.caption("По final_decision")
            _fd = df["final_decision"].fillna("(no data)")
            st.dataframe(
                _fd.value_counts().rename_axis("final_decision").reset_index(name="n"),
                hide_index=True,
                use_container_width=True,
            )
        with _e3:
            st.caption("По status")
            _ss = df["status"].replace({"": "(ok / empty)"})
            st.dataframe(
                _ss.value_counts().rename_axis("status").reset_index(name="n"),
                hide_index=True,
                use_container_width=True,
            )


def _bulk_warning_token_set(flags_str) -> set:
    if flags_str is None or (isinstance(flags_str, float) and pd.isna(flags_str)):
        return set()
    s = str(flags_str).strip()
    if not s:
        return set()
    return {t.strip() for t in s.split(",") if t.strip()}


def _bulk_unique_warning_flag_options(df: pd.DataFrame) -> list:
    acc = set()
    for w in df["warning_flags"]:
        acc |= _bulk_warning_token_set(w)
    return sorted(acc)


def _bulk_apply_table_filters(
    df: pd.DataFrame,
    sel_priority: list,
    all_priority: list,
    sel_final_decision: list,
    all_final_decision: list,
    sel_warning_flags: list,
    all_warning_flags: list,
) -> pd.DataFrame:
    """
    AND across dimensions. If every option is selected for a dimension, that dimension is not applied.
    warning_flags: row matches if intersection of row tokens and selected flags is non-empty (OR).
    """
    m = pd.Series(True, index=df.index)
    if set(sel_priority) != set(all_priority):
        m &= df["priority"].isin(sel_priority)
    if set(sel_final_decision) != set(all_final_decision):
        fd_disp = df["final_decision"].apply(
            lambda x: "(no data)" if pd.isna(x) else str(x)
        )
        m &= fd_disp.isin(sel_final_decision)
    if all_warning_flags and set(sel_warning_flags) != set(all_warning_flags):
        sel_w = set(sel_warning_flags)

        def _row_matches_warnings(val):
            return bool(sel_w & _bulk_warning_token_set(val))

        m &= df["warning_flags"].apply(_row_matches_warnings)
    return df.loc[m]


def _bulk_render_insights(df: pd.DataFrame) -> None:
    """Short human-readable counts over the full bulk result (not table filters)."""
    if df is None or df.empty:
        return
    total = len(df)
    negative = int((df["final_decision"] == "Negative impact").sum())
    insufficient = int((df["final_decision"] == "Insufficient data").sum())
    missing_py = int(
        df["warning_flags"]
        .apply(lambda w: "MISSING_PY" in _bulk_warning_token_set(w))
        .sum()
    )
    missing_cy = int((df["status"] == "Missing current year data").sum())

    parts = [f"Analyzed **{total}** categories."]
    if negative:
        parts.append(f"❌ **{negative}** categories show negative impact.")
    if insufficient:
        parts.append(f"📉 **{insufficient}** categories have insufficient data.")
    if missing_py:
        parts.append(f"📊 **{missing_py}** categories have no previous-year data.")
    if missing_cy:
        parts.append(f"⚠️ **{missing_cy}** categories missing current data.")

    st.markdown("### Summary insights")
    st.markdown("\n\n".join(parts))


_BULK_ROW_BG_MISSING_CY = "#e8e8e8"
_BULK_ROW_BG_NEGATIVE = "#ffd6d6"
_BULK_ROW_BG_INSUFFICIENT = "#fff3bf"
_BULK_ROW_BG_POSITIVE = "#d3f9d8"


def _bulk_row_background(row: pd.Series):
    """Return background color hex for a bulk table row, or None for default."""
    if row.get("status") == "Missing current year data":
        return _BULK_ROW_BG_MISSING_CY
    fd = row.get("final_decision")
    if fd is None or (isinstance(fd, float) and pd.isna(fd)):
        return None
    fd = str(fd)
    if fd == "Negative impact":
        return _BULK_ROW_BG_NEGATIVE
    if fd == "Insufficient data":
        return _BULK_ROW_BG_INSUFFICIENT
    if fd == "Positive impact":
        return _BULK_ROW_BG_POSITIVE
    return None


def _bulk_format_table_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Тысячи с пробелом в category_id; остальные колонки без изменений."""
    if df is None or df.empty:
        return df
    out = df.copy()
    if "category_id" in out.columns:
        def _fmt_cat(x):
            if pd.isna(x):
                return x
            try:
                return format_integer(int(x), group_thousands=True)
            except (TypeError, ValueError):
                return x

        out["category_id"] = out["category_id"].map(_fmt_cat)
    return out


def _bulk_styled_dataframe(df: pd.DataFrame):
    """Row-wise background via pandas Styler (full row repaint per cell)."""
    if df is None or df.empty:
        return df

    def _apply_row_styles(row: pd.Series):
        bg = _bulk_row_background(row)
        if not bg:
            return pd.Series([""] * len(row), index=row.index)
        css = f"background-color: {bg}"
        return pd.Series([css] * len(row), index=row.index)

    return df.style.apply(_apply_row_styles, axis=1).hide(axis="index")


def _parse_category_ids(text: str):
    """
    Parse category IDs from textarea (newline, comma, semicolon, whitespace).
    Returns (valid_unique_ids_in_order, invalid_tokens).
    """
    if text is None or not str(text).strip():
        return [], []
    raw = str(text).strip()
    parts = [p for p in re.split(r"[\s,\n;]+", raw) if p]
    valid = []
    invalid = []
    seen = set()
    for p in parts:
        try:
            v = int(p)
            if v not in seen:
                seen.add(v)
                valid.append(v)
        except ValueError:
            invalid.append(p)
    return valid, invalid


parsed_category_ids, invalid_category_tokens = _parse_category_ids(category_input)
if invalid_category_tokens:
    shown = invalid_category_tokens[:25]
    extra = " …" if len(invalid_category_tokens) > 25 else ""
    st.warning(
        "Не удалось разобрать как целое число: "
        + ", ".join(repr(t) for t in shown)
        + extra
    )

_category_bulk_mode = len(parsed_category_ids) > 1
if _category_bulk_mode:
    _resolved_single_category_id = None
elif len(parsed_category_ids) == 1:
    _resolved_single_category_id = int(parsed_category_ids[0])
elif len(parsed_category_ids) == 0 and merged_data:
    _cy_keys = sorted(int(k) for k in merged_data.keys())
    if len(_cy_keys) == 1:
        _resolved_single_category_id = _cy_keys[0]
        st.caption(
            f"В файлах одна категория — для Before/After используется **{_resolved_single_category_id}**."
        )
    else:
        _resolved_single_category_id = int(
            st.selectbox(
                "Category ID (Current Year): в файлах несколько категорий — выберите для анализа",
                options=_cy_keys,
                key="cy_single_category_pick",
            )
        )
else:
    _resolved_single_category_id = None

_category_single_mode = _resolved_single_category_id is not None and not _category_bulk_mode

_merge_dirty = st.session_state.pop("_merge_files_dirty", False)
if _merge_dirty:
    st.session_state.pop("_cy_ocr_override", None)
    st.session_state.pop("_py_ocr_override", None)
    st.session_state.pop("_cy_ocr_snapshot", None)

if merged_data and _category_single_mode:
    category_id = int(_resolved_single_category_id)
    prev_cat = st.session_state.get("_prev_category_id_for_merge", "")
    sig = str(category_id)
    should_apply = sig != prev_cat or _merge_dirty
    # Do not drop CY OCR when prev_cat is still uninitialized (""): sig != "" is always true and
    # would clear _cy_ocr_override before merge, letting file merge overwrite OCR with zeros.
    if sig != prev_cat and prev_cat != "":
        st.session_state.pop("_cy_ocr_override", None)
        st.session_state.pop("_cy_ocr_snapshot", None)
    if category_id in merged_data:
        if should_apply and not st.session_state.get("_cy_ocr_override"):
            data = merged_data[category_id]
            _bl = data.get("baseline") or {}
            for _dk, _, _dtyp in _CY_INPUT_METRICS:
                sk = _cy_sess_key(_dk)
                _bv = _bl.get(_dk) if _bl else None
                _bf = data["before"].get(_dk)
                _af = data["after"].get(_dk)
                if _dtyp == "int":
                    st.session_state[f"{sk}_baseline"] = int(float(_bv or 0))
                    st.session_state[f"{sk}_before"] = int(float(_bf or 0))
                    st.session_state[f"{sk}_after"] = int(float(_af or 0))
                else:
                    st.session_state[f"{sk}_baseline"] = float(_bv or 0)
                    st.session_state[f"{sk}_before"] = float(_bf or 0)
                    st.session_state[f"{sk}_after"] = float(_af or 0)
        st.session_state["_prev_category_id_for_merge"] = sig
    else:
        st.warning("Category not found in uploaded data")
        st.session_state["_prev_category_id_for_merge"] = sig

_cy_ocr_restore_from_snapshot_if_blanked(st.session_state)

_merge_py_dirty = st.session_state.pop("_py_merge_dirty", False)
if _merge_py_dirty:
    st.session_state.pop("_py_ocr_override", None)

if merged_data_previous_year and _category_single_mode:
    category_id_py = int(_resolved_single_category_id)
    prev_cat_py = st.session_state.get("_prev_category_id_for_py_merge", "")
    sig_py = str(category_id_py)
    should_apply_py = sig_py != prev_cat_py or _merge_py_dirty
    if sig_py != prev_cat_py:
        st.session_state.pop("_py_ocr_override", None)
    if category_id_py in merged_data_previous_year:
        if should_apply_py and not st.session_state.get("_py_ocr_override"):
            pdata = merged_data_previous_year[category_id_py]
            st.session_state["matrix_py_paid_users_before"] = int(
                float(pdata["before"].get("paid_users") or 0)
            )
            st.session_state["matrix_py_paid_users_after"] = int(
                float(pdata["after"].get("paid_users") or 0)
            )
            st.session_state["matrix_py_spending_before"] = float(
                pdata["before"].get("spending") or 0
            )
            st.session_state["matrix_py_spending_after"] = float(
                pdata["after"].get("spending") or 0
            )
            st.session_state["matrix_py_ac_before"] = int(
                float(pdata["before"].get("active_listers") or 0)
            )
            st.session_state["matrix_py_ac_after"] = int(
                float(pdata["after"].get("active_listers") or 0)
            )
        st.session_state["_prev_category_id_for_py_merge"] = sig_py
    else:
        st.warning("Category not found in Previous Year uploaded data")
        st.session_state["_prev_category_id_for_py_merge"] = sig_py

if _category_bulk_mode:
    st.subheader("Bulk Category IDs")
    st.caption("Сводка по списку. Автозаполнение полей и расчёт по кнопке Calculate — только для одного ID.")
    n = len(parsed_category_ids)
    in_cy = [i for i in parsed_category_ids if i in merged_data]
    in_py = [i for i in parsed_category_ids if i in merged_data_previous_year]
    miss_cy = [i for i in parsed_category_ids if i not in merged_data]
    miss_py = [i for i in parsed_category_ids if i not in merged_data_previous_year]
    st.markdown(
        f"- **Распознано валидных ID:** {n}\n"
        f"- **Есть в Current Year data:** {len(in_cy)}"
        + ("" if merged_data else " *(файлы Current Year не загружены — считаем все отсутствующими)*")
        + f"\n- **Есть в Previous Year data:** {len(in_py)}"
        + (
            ""
            if merged_data_previous_year
            else " *(файлы Previous Year не загружены — считаем все отсутствующими)*"
        )
        + f"\n- **Не в Current Year data:** {miss_cy if miss_cy else '—'}"
        + f"\n- **Не в Previous Year data:** {miss_py if miss_py else '—'}"
    )


def _matrix_geo_thresholds(geo: str):
    gk = str(geo).upper() if geo else "default"
    return GEO_THRESHOLDS.get(gk, GEO_THRESHOLDS["default"])


def _matrix_safe_div(numerator, denominator):
    if denominator is None or numerator is None:
        return None
    if denominator == 0:
        return None
    return numerator / denominator


def _matrix_pct_diff(before, after):
    if before is None or after is None:
        return None
    if before == 0:
        return None
    return (after - before) / before * 100.0


def _matrix_classify_label(pct_diff, geo: str, metric_key: str):
    if pct_diff is None:
        return "—"
    th = _matrix_geo_thresholds(geo)[metric_key]
    code = classify_change(
        pct_diff,
        growth_threshold=th["growth"],
        decline_threshold=th["decline"],
    )
    return decode_status(code)


def _build_ppv_matrix_rows(
    geo: str,
    cy_paid_users_b,
    cy_paid_users_a,
    cy_spending_b,
    cy_spending_a,
    cy_ac_b,
    cy_ac_a,
    py_paid_users_b,
    py_paid_users_a,
    py_spending_b,
    py_spending_a,
    py_ac_b,
    py_ac_a,
):
    """Returns list of flat dicts for display (Paid users, Spending, CR, Active)."""
    cy_cr_b = _matrix_safe_div(cy_paid_users_b, cy_ac_b)
    cy_cr_a = _matrix_safe_div(cy_paid_users_a, cy_ac_a)
    py_cr_b = _matrix_safe_div(py_paid_users_b, py_ac_b)
    py_cr_a = _matrix_safe_div(py_paid_users_a, py_ac_a)

    metrics = [
        (
            "Paid users",
            "npl",
            cy_paid_users_b,
            cy_paid_users_a,
            py_paid_users_b,
            py_paid_users_a,
            True,
        ),
        (
            "Spending",
            "sp",
            cy_spending_b,
            cy_spending_a,
            py_spending_b,
            py_spending_a,
            False,
        ),
        ("CR", "cr", cy_cr_b, cy_cr_a, py_cr_b, py_cr_a, False),
        ("Active listers", "npl", cy_ac_b, cy_ac_a, py_ac_b, py_ac_a, True),
    ]

    rows_out = []
    for title, th_key, cbb, cba, pbb, pba, as_int in metrics:
        cy_d = _matrix_pct_diff(cbb, cba)
        py_d = _matrix_pct_diff(pbb, pba)
        y2y = (cy_d - py_d) if cy_d is not None and py_d is not None else None
        res_cy = _matrix_classify_label(cy_d, geo, th_key)

        rows_out.append(
            {
                "Metric": title,
                "Period": "Current Year",
                "Before": format_matrix_metric(cbb, as_int=as_int and cbb is not None),
                "After": format_matrix_metric(cba, as_int=as_int and cba is not None),
                "Diff %": format_percent(cy_d),
                "Result": res_cy,
            }
        )
        rows_out.append(
            {
                "Metric": title,
                "Period": "Previous Year",
                "Before": format_matrix_metric(pbb, as_int=as_int and pbb is not None),
                "After": format_matrix_metric(pba, as_int=as_int and pba is not None),
                "Diff %": format_percent(py_d),
                "Result": "",
            }
        )
        rows_out.append(
            {
                "Metric": title,
                "Period": "diff Y2Y",
                "Before": "",
                "After": "",
                "Diff %": format_percent(y2y),
                "Result": "",
            }
        )
    return rows_out


def _compute_potential_spendings_block(
    cy_paid_users_b,
    cy_paid_users_a,
    cy_spending_b,
    cy_spending_a,
    cy_ac_b,
    cy_ac_a,
    py_paid_users_b,
    py_paid_users_a,
    py_ac_b,
    py_ac_a,
):
    cy_cr_b = _matrix_safe_div(cy_paid_users_b, cy_ac_b)
    cy_cr_a = _matrix_safe_div(cy_paid_users_a, cy_ac_a)
    py_cr_b = _matrix_safe_div(py_paid_users_b, py_ac_b)
    py_cr_a = _matrix_safe_div(py_paid_users_a, py_ac_a)
    py_cr_diff_pct = _matrix_pct_diff(py_cr_b, py_cr_a)

    fact_arppu = _matrix_safe_div(cy_spending_a, cy_paid_users_a)
    could_be_arppu = _matrix_safe_div(cy_spending_b, cy_paid_users_b)

    expected_cr_after = None
    if cy_cr_b is not None and py_cr_diff_pct is not None:
        expected_cr_after = cy_cr_b * (1.0 + py_cr_diff_pct / 100.0)

    could_be_spendings = None
    if cy_ac_a is not None and expected_cr_after is not None and could_be_arppu is not None:
        could_be_spendings = cy_ac_a * expected_cr_after * could_be_arppu

    fact_spending = cy_spending_a
    potential_spendings_diff = None
    if fact_spending is not None and could_be_spendings is not None and could_be_spendings != 0:
        potential_spendings_diff = fact_spending / could_be_spendings - 1.0

    return {
        "fact_arppu": fact_arppu,
        "could_be_arppu": could_be_arppu,
        "fact_spending": fact_spending,
        "could_be_spendings": could_be_spendings,
        "potential_spendings_diff": potential_spendings_diff,
    }


def _potential_spendings_row_diff_pct(fact, could_be) -> float | None:
    """((Fact / Could be) - 1) * 100 — на сколько % Fact выше/ниже Could be (база = Could be)."""
    if fact is None or could_be is None:
        return None
    try:
        ff = float(fact)
        cf = float(could_be)
    except (TypeError, ValueError):
        return None
    if cf == 0:
        return None
    if ff == 0:
        return None
    return (ff / cf - 1.0) * 100.0


def _potential_spendings_row_diff_abs(fact, could_be) -> float | None:
    """Fact − Could be; None if inputs invalid."""
    if fact is None or could_be is None:
        return None
    try:
        return float(fact) - float(could_be)
    except (TypeError, ValueError):
        return None


def _build_potential_spendings_table_df(_pot: dict) -> pd.DataFrame:
    """Compact table: Potential Spendings | Fact | Could be | diff | diff %."""
    fa, ca = _pot.get("fact_arppu"), _pot.get("could_be_arppu")
    fs, cs = _pot.get("fact_spending"), _pot.get("could_be_spendings")
    d1 = _potential_spendings_row_diff_pct(fa, ca)
    d2 = _potential_spendings_row_diff_pct(fs, cs)
    a1 = _potential_spendings_row_diff_abs(fa, ca)
    a2 = _potential_spendings_row_diff_abs(fs, cs)
    return pd.DataFrame(
        [
            {
                "Potential Spendings": "ARPpU",
                "Fact": format_potential_amount(fa),
                "Could be": format_potential_amount(ca),
                "diff": format_delta(a1, as_integer=False),
                "diff %": format_percent(d1, zero_display="0") if d1 is not None else "—",
            },
            {
                "Potential Spendings": "Spendings",
                "Fact": format_potential_amount(fs),
                "Could be": format_potential_amount(cs),
                "diff": format_delta(a2, as_integer=True),
                "diff %": format_percent(d2, zero_display="0") if d2 is not None else "—",
            },
        ]
    )


if _category_bulk_mode:
    with st.expander("Скрин с данными (OCR)", expanded=False):
        st.info(
            "OCR для ручного ввода доступен только в **single analysis** (одна категория). "
            "В bulk-режиме используйте файлы и таблицу bulk."
        )
else:
    with st.expander("Скрин с данными (OCR) — single analysis", expanded=False):
        st.caption(
            "**Current Year** screenshot → поля **Current Year input** (manual override). "
            "**Previous Year** screenshot → только блок **Previous Year** (матрица, Y2Y, Potential); "
            "строка Previous Year в матрице **не** заполняется из OCR Current Year. "
            "**category_id** для OCR не используется — только совпадение по **названию метрики**."
        )
        st.caption(
            "Нативной вставки изображения из буфера (Ctrl+V) в Streamlit **без** отдельного "
            "JS-компонента или доп. пакетов нет: используйте **загрузку файла** или вставьте "
            "**data URL** / **base64** в поле ниже (можно получить из DevTools / внешнего конвертера)."
        )
        tab_ocr_cy, tab_ocr_py = st.tabs(
            ["Current Year screenshot", "Previous Year screenshot"]
        )

        with tab_ocr_cy:
            up_cy = st.file_uploader(
                "Файл скриншота (PNG/JPG/JPEG)",
                type=["png", "jpg", "jpeg"],
                key="ocr_cy_file_uploader",
            )
            paste_cy = st.text_area(
                "Или вставьте data:image/...;base64,... либо сырой base64",
                height=72,
                key="ocr_cy_paste_b64",
                placeholder="data:image/png;base64,iVBORw0KGgo...",
            )
            img_cy, dec_err_cy = _decode_optional_image_bytes(up_cy, paste_cy)
            if dec_err_cy:
                st.warning(dec_err_cy)
            else:
                _preview_cy = img_cy if img_cy is not None else (
                    up_cy.getvalue() if up_cy is not None else None
                )
                if _preview_cy:
                    _ocr_render_screenshot_preview(_preview_cy, label="Current Year")

            if st.button("Распознать и применить к Current Year input", key="ocr_cy_apply_btn"):
                img_b, err_b = _decode_optional_image_bytes(up_cy, paste_cy)
                if err_b:
                    st.session_state["_ocr_cy_feedback"] = _build_cy_ocr_feedback(
                        None, err_b, {}, [], []
                    )
                    st.session_state["_cy_ocr_override"] = True
                    st.rerun()
                elif not img_b:
                    st.session_state["_ocr_cy_feedback"] = _build_cy_ocr_feedback(
                        None,
                        "Нет изображения: загрузите файл или вставьте data URL / base64.",
                        {},
                        [],
                        [],
                    )
                    st.session_state["_cy_ocr_override"] = True
                    st.rerun()
                else:
                    variants, ocr_error = _ocr_collect_tesseract_variants(img_b)
                    if ocr_error:
                        st.session_state["_ocr_cy_feedback"] = _build_cy_ocr_feedback(
                            None, ocr_error, {}, [], []
                        )
                    else:
                        parsed = _merge_cy_parsed_from_ocr_variants(variants)
                        ocr_text = _ocr_pick_preview_text_cy(variants)
                        # Matrix "Previous Year" row: only matrix_py_* (PY files / PY OCR). Do not mirror
                        # CY OCR into matrix_py_* — that duplicated both rows and could overwrite PY OCR.
                        applied_labels, dbg_rows = _apply_cy_ocr_parsed_to_session(
                            st.session_state,
                            parsed,
                            mirror_matrix_py=False,
                        )
                        st.session_state["_ocr_cy_feedback"] = _build_cy_ocr_feedback(
                            ocr_text, None, parsed, applied_labels, dbg_rows
                        )
                    st.session_state["_cy_ocr_override"] = True
                    st.rerun()

            _render_ocr_feedback_messages(st.session_state.get("_ocr_cy_feedback"))
            _render_ocr_debug_expanders(st.session_state.get("_ocr_cy_feedback"), "cy")

        with tab_ocr_py:
            up_py = st.file_uploader(
                "Файл скриншота (PNG/JPG/JPEG)",
                type=["png", "jpg", "jpeg"],
                key="ocr_py_file_uploader",
            )
            paste_py = st.text_area(
                "Или вставьте data:image/...;base64,... либо сырой base64",
                height=72,
                key="ocr_py_paste_b64",
                placeholder="data:image/png;base64,iVBORw0KGgo...",
            )
            img_py, dec_err_py = _decode_optional_image_bytes(up_py, paste_py)
            if dec_err_py:
                st.warning(dec_err_py)
            else:
                _preview_py = img_py if img_py is not None else (
                    up_py.getvalue() if up_py is not None else None
                )
                if _preview_py:
                    _ocr_render_screenshot_preview(_preview_py, label="Previous Year")

            if st.button(
                "Распознать и применить к Previous Year (матрица)",
                key="ocr_py_apply_btn",
            ):
                img_b, err_b = _decode_optional_image_bytes(up_py, paste_py)
                if err_b:
                    st.session_state["_ocr_py_feedback"] = _build_py_ocr_feedback(
                        None, err_b, {}, [], []
                    )
                    st.session_state["_py_ocr_override"] = True
                    st.rerun()
                elif not img_b:
                    st.session_state["_ocr_py_feedback"] = _build_py_ocr_feedback(
                        None,
                        "Нет изображения: загрузите файл или вставьте data URL / base64.",
                        {},
                        [],
                        [],
                    )
                    st.session_state["_py_ocr_override"] = True
                    st.rerun()
                else:
                    variants, ocr_error = _ocr_collect_tesseract_variants(img_b)
                    if ocr_error:
                        st.session_state["_ocr_py_feedback"] = _build_py_ocr_feedback(
                            None, ocr_error, {}, [], []
                        )
                    else:
                        parsed = _merge_py_parsed_from_ocr_variants(variants)
                        ocr_text = _ocr_pick_preview_text_py(variants)
                        applied_labels, dbg_rows = _apply_py_ocr_parsed_to_session(
                            st.session_state, parsed
                        )
                        st.session_state["_ocr_py_feedback"] = _build_py_ocr_feedback(
                            ocr_text, None, parsed, applied_labels, dbg_rows
                        )
                    st.session_state["_py_ocr_override"] = True
                    st.rerun()

            _render_ocr_feedback_messages(st.session_state.get("_ocr_py_feedback"))
            _render_ocr_debug_expanders(st.session_state.get("_ocr_py_feedback"), "py")

st.divider()

_cy_ba: dict[str, tuple[float, float]] = {}
with st.expander(
    "🧮 Current Year input (manual override)",
    expanded=not bool(merged_data),
):
    _cy_pad_l, _cy_main, _cy_pad_r = st.columns([1, 6, 1])
    with _cy_main:
        with st.container(border=True):
            st.subheader("Current Year input")
            st.caption(
                "Абсолютные **Before** и **After**. Колонки **До %** и **После %** в сводке ниже "
                "считаются в приложении. **До, %** — относительно опорного периода из выгрузки (если есть в данных). "
                "**После, %** = (After − Before) / Before."
            )

            _cy_h0, _cy_h1, _cy_h2 = st.columns([1.55, 1, 1])
            with _cy_h0:
                st.caption("Metric")
            with _cy_h1:
                st.markdown("**Before**")
            with _cy_h2:
                st.markdown("**After**")

            # Return values are the source of truth this run (session_state can lag behind widgets).
            for _dk, _dlabel, _dtyp in _CY_INPUT_METRICS:
                sk = _cy_sess_key(_dk)
                _cr0, _cr1, _cr2 = st.columns([1.55, 1, 1])
                with _cr0:
                    st.markdown(_dlabel)
                with _cr1:
                    if _dtyp == "int":
                        _bv = st.number_input(
                            f"{_dlabel} — Before",
                            min_value=0,
                            key=f"{sk}_before",
                            label_visibility="collapsed",
                        )
                    else:
                        _bv = st.number_input(
                            f"{_dlabel} — Before",
                            min_value=0.0,
                            key=f"{sk}_before",
                            label_visibility="collapsed",
                        )
                with _cr2:
                    if _dtyp == "int":
                        _av = st.number_input(
                            f"{_dlabel} — After",
                            min_value=0,
                            key=f"{sk}_after",
                            label_visibility="collapsed",
                        )
                    else:
                        _av = st.number_input(
                            f"{_dlabel} — After",
                            min_value=0.0,
                            key=f"{sk}_after",
                            label_visibility="collapsed",
                        )
                _cy_ba[_dk] = (float(_bv), float(_av))

_ss = st.session_state
if _cy_ba:
    paid_users_before = int(_cy_ba["paid_users"][0])
    paid_users_after = int(_cy_ba["paid_users"][1])
    spending_before = float(_cy_ba["spending"][0])
    spending_after = float(_cy_ba["spending"][1])
    active_before = int(_cy_ba["active_listers"][0])
    active_after = int(_cy_ba["active_listers"][1])
else:
    paid_users_before = int(_ss.get("paid_users_before") or 0)
    paid_users_after = int(_ss.get("paid_users_after") or 0)
    spending_before = float(_ss.get("spending_before") or 0)
    spending_after = float(_ss.get("spending_after") or 0)
    active_before = int(_ss.get("active_before") or 0)
    active_after = int(_ss.get("active_after") or 0)

_cy_baseline_snap = {}
if merged_data and _category_single_mode and _resolved_single_category_id is not None:
    _cid_tbl = int(_resolved_single_category_id)
    if _cid_tbl in merged_data:
        _bs = merged_data[_cid_tbl].get("baseline")
        if isinstance(_bs, dict):
            _cy_baseline_snap = _bs

_cy_tbl = []
for _dk, _dlabel, _dtyp in _CY_INPUT_METRICS:
    sk = _cy_sess_key(_dk)
    _bl_raw = _cy_baseline_snap.get(_dk)
    if _bl_raw is None:
        _bl_raw = _ss.get(f"{sk}_baseline")
    _pair_ui = _cy_ba.get(_dk) if _cy_ba else None
    if _pair_ui is not None:
        _bfv, _afv = _pair_ui[0], _pair_ui[1]
    else:
        _bfv = _ss.get(f"{sk}_before", 0)
        _afv = _ss.get(f"{sk}_after", 0)
    _as_int = _dtyp == "int"
    if _as_int:
        _bfv, _afv = int(_bfv or 0), int(_afv or 0)
        if _bl_raw is None or (isinstance(_bl_raw, str) and _bl_raw.strip() == ""):
            _blv_for_pct = None
        else:
            try:
                _blv_for_pct = int(float(_bl_raw))
            except (TypeError, ValueError):
                _blv_for_pct = None
    else:
        _bfv, _afv = float(_bfv or 0), float(_afv or 0)
        if _bl_raw is None or (isinstance(_bl_raw, str) and _bl_raw.strip() == ""):
            _blv_for_pct = None
        else:
            try:
                _blv_for_pct = float(_bl_raw)
            except (TypeError, ValueError):
                _blv_for_pct = None

    _bf_f = float(_bfv)
    _af_f = float(_afv)
    _before_diff = pct_change_relative(_bf_f, _blv_for_pct)
    _after_diff = pct_change_relative(_af_f, _bf_f)

    _cy_tbl.append(
        {
            "Metric": _dlabel,
            "До": format_matrix_metric(_bfv, as_int=_as_int),
            "После": format_matrix_metric(_afv, as_int=_as_int),
            "beforeDiff": _before_diff,
            "afterDiff": _after_diff,
        }
    )

_df_cy = pd.DataFrame(_cy_tbl)
_df_cy = _df_cy.rename(columns={"beforeDiff": "До %", "afterDiff": "После %"})
_pct_cols = ["До %", "После %"]
for _pc in _pct_cols:
    _df_cy[_pc] = pd.to_numeric(_df_cy[_pc], errors="coerce")

_sty_cy = _df_cy.style
try:
    _sty_cy = _sty_cy.map(_cy_diff_semantic_style, subset=_pct_cols)
except AttributeError:
    _sty_cy = _sty_cy.applymap(_cy_diff_semantic_style, subset=_pct_cols)
_sty_cy = _sty_cy.format(format_summary_percent_cell, subset=_pct_cols)

st.divider()

with st.container(border=True):
    st.markdown("##### Сводка по метрикам")
    st.caption("В одной строке: абсолютные **До** / **После** и относительные **До %** / **После %**.")
    st.dataframe(_sty_cy, use_container_width=True, hide_index=True)

with st.expander("Previous Year (для матрицы, Y2Y и Potential Spending)", expanded=False):
    st.caption(
        "Если загружены Previous Year файлы и найден Category ID — значения подставятся автоматически. "
        "Иначе введите вручную или используйте OCR вкладки **Previous Year**. "
        "Строка Previous Year в матрице берётся только отсюда или из файлов PY (не из OCR Current Year)."
    )
    py_l, py_r = st.columns(2)
    with py_l:
        st.markdown("**Paid users**")
        matrix_py_paid_users_before = st.number_input(
            "PY Paid users Before", min_value=0, key="matrix_py_paid_users_before"
        )
        matrix_py_paid_users_after = st.number_input(
            "PY Paid users After", min_value=0, key="matrix_py_paid_users_after"
        )
        st.markdown("**Spending**")
        matrix_py_spending_before = st.number_input(
            "PY Spending Before", min_value=0.0, key="matrix_py_spending_before"
        )
        matrix_py_spending_after = st.number_input(
            "PY Spending After", min_value=0.0, key="matrix_py_spending_after"
        )
    with py_r:
        st.markdown("**Active listers**")
        matrix_py_ac_before = st.number_input("PY Active Before", min_value=0, key="matrix_py_ac_before")
        matrix_py_ac_after = st.number_input("PY Active After", min_value=0, key="matrix_py_ac_after")

with st.container(border=True):
    st.subheader("PPV matrix (analytics)")
    st.caption("Только отображение; не влияет на решение по кнопке Calculate.")
    mq_left, mq_right = st.columns([1.55, 1.0])
    with mq_left:
        _matrix_df = pd.DataFrame(
            _build_ppv_matrix_rows(
                geo,
                paid_users_before,
                paid_users_after,
                spending_before,
                spending_after,
                active_before,
                active_after,
                matrix_py_paid_users_before,
                matrix_py_paid_users_after,
                matrix_py_spending_before,
                matrix_py_spending_after,
                matrix_py_ac_before,
                matrix_py_ac_after,
            )
        )
        st.dataframe(_matrix_df, hide_index=True, use_container_width=True)
    with mq_right:
        st.markdown("##### Potential Spendings")
        _pot = _compute_potential_spendings_block(
            paid_users_before,
            paid_users_after,
            spending_before,
            spending_after,
            active_before,
            active_after,
            matrix_py_paid_users_before,
            matrix_py_paid_users_after,
            matrix_py_ac_before,
            matrix_py_ac_after,
        )
        _pot_df = _build_potential_spendings_table_df(_pot)
        st.dataframe(_pot_df, hide_index=True, use_container_width=True)

st.markdown("##### Scenario")
st.caption(
    f"Активный сценарий: **{scenario}**. Переключение — в панели **Настройки** вверху страницы."
)

_ac1, _ac2 = st.columns([3, 1])
with _ac1:
    st.caption("Single mode: после просмотра аналитики нажмите **Calculate** для итогового решения.")
with _ac2:
    _run_calc = st.button("Calculate", type="primary", use_container_width=True)

if _run_calc:
    has_invalid_data = False
    if active_before < paid_users_before:
        st.warning(
            "Active listers (Before) не могут быть меньше Paid users (Before)"
        )
        has_invalid_data = True
    if active_after < paid_users_after:
        st.warning(
            "Active listers (After) не могут быть меньше Paid users (After)"
        )
        has_invalid_data = True

    if (
        paid_users_before == 0
        and paid_users_after == 0
        and spending_before == 0
        and spending_after == 0
        and active_before == 0
        and active_after == 0
    ):
        st.info("Вы ввели нулевые значения — это тестовый расчет")

    if geo == "RS":
        st.info("Для RS используются default thresholds (нужна дополнительная настройка)")

    if not has_invalid_data:
        result = analyze_category(
            npl_before=paid_users_before,
            npl_after=paid_users_after,
            sp_before=spending_before,
            sp_after=spending_after,
            active_before=active_before,
            active_after=active_after,
            geo=geo or "default",
            force_low_npl=scenario == "Low NPL (<10)",
            is_other_category=scenario == "Other category",
        )

        with st.container(border=True):
            st.subheader("Results")
            with st.expander("Детали по метрикам (NPL, Spending, CR)", expanded=False):
                paid_users_row = (
                    f"Paid users: {result['npl_diff']:.2f}% → "
                    f"{decode_status(result['npl_code'])} ({result['npl_code']})"
                )
                if result["npl_code"] == "G":
                    st.success(paid_users_row)
                elif result["npl_code"] == "D":
                    st.error(paid_users_row)
                else:
                    st.warning(paid_users_row)

                sp_text = f"Spending: {result['sp_diff']:.2f}% → {decode_status(result['sp_code'])} ({result['sp_code']})"
                if result["sp_code"] == "G":
                    st.success(sp_text)
                elif result["sp_code"] == "D":
                    st.error(sp_text)
                else:
                    st.warning(sp_text)

                cr_text = f"Conversion: {result['cr_diff']:.2f}% → {decode_status(result['cr_code'])} ({result['cr_code']})"
                if result["cr_code"] == "G":
                    st.success(cr_text)
                elif result["cr_code"] == "D":
                    st.error(cr_text)
                else:
                    st.warning(cr_text)

            st.markdown("**Decision code**")
            st.code(result["decision_code"])

            st.markdown("**Final decision**")
            final_decision = result["final_decision"]
            if final_decision == "Positive impact":
                st.success(final_decision)
            elif final_decision == "Negative impact":
                st.error(final_decision)
            elif final_decision == "No impact":
                st.info(final_decision)
            else:
                st.warning(final_decision)

            st.markdown("**Next step**")
            st.info(result["next_step"])

if _category_bulk_mode:
    st.subheader("Bulk analysis")
    st.caption(
        "Используются загруженные Current Year / Previous Year данные и выбранные GEO и Scenario. "
        "Price data в bulk не используется."
    )
    _bulk_sig = (
        tuple(parsed_category_ids),
        str(geo or ""),
        str(scenario),
        st.session_state.get("_upload_sig"),
        st.session_state.get("_upload_sig_py"),
    )
    if st.session_state.get("_bulk_analysis_sig") != _bulk_sig:
        st.session_state.pop("bulk_analysis_result", None)
    st.session_state["_bulk_analysis_sig"] = _bulk_sig

    _force_low = scenario == "Low NPL (<10)"
    _is_other = scenario == "Other category"
    if st.button("Run bulk analysis", type="primary", key="run_bulk_analysis"):
        if not merged_data:
            st.warning("Загрузите Current Year файлы (spending, active, price), чтобы строить bulk-таблицу.")
        else:
            st.session_state["bulk_analysis_result"] = _bulk_analysis_dataframe(
                parsed_category_ids,
                merged_data,
                merged_data_previous_year or {},
                geo,
                _force_low,
                _is_other,
            )
            st.session_state["bulk_run_nonce"] = int(st.session_state.get("bulk_run_nonce", 0)) + 1

    _bulk_df = st.session_state.get("bulk_analysis_result")
    if _bulk_df is not None and not _bulk_df.empty:
        _bulk_render_insights(_bulk_df)
        _bulk_render_summary(_bulk_df)
        _bulk_nonce = int(st.session_state.get("bulk_run_nonce", 0))
        _pri_opts = sorted(_bulk_df["priority"].dropna().unique().tolist())
        _fd_opts = sorted(
            _bulk_df["final_decision"]
            .apply(lambda x: "(no data)" if pd.isna(x) else str(x))
            .unique()
            .tolist()
        )
        _wf_opts = _bulk_unique_warning_flag_options(_bulk_df)

        st.markdown("##### Фильтры таблицы")
        st.caption(
            "Summary выше считается по всем строкам. Фильтры ниже влияют только на отображение таблицы."
        )
        _fc1, _fc2, _fc3 = st.columns(3)
        with _fc1:
            _sel_p = st.multiselect(
                "Priority",
                options=_pri_opts,
                default=_pri_opts,
                key=f"bulk_table_f_priority_{_bulk_nonce}",
            )
        with _fc2:
            _sel_fd = st.multiselect(
                "final_decision",
                options=_fd_opts,
                default=_fd_opts,
                key=f"bulk_table_f_final_decision_{_bulk_nonce}",
            )
        with _fc3:
            if _wf_opts:
                _sel_wf = st.multiselect(
                    "warning_flags",
                    options=_wf_opts,
                    default=_wf_opts,
                    key=f"bulk_table_f_warning_flags_{_bulk_nonce}",
                )
            else:
                st.caption("warning_flags: нет значений в данных")
                _sel_wf = []

        _filtered_bulk = _bulk_apply_table_filters(
            _bulk_df,
            _sel_p,
            _pri_opts,
            _sel_fd,
            _fd_opts,
            _sel_wf,
            _wf_opts,
        )
        st.caption(
            f"Showing {len(_filtered_bulk)} of {len(_bulk_df)} categories"
        )
        _bulk_table_display = _bulk_styled_dataframe(
            _bulk_format_table_for_display(_filtered_bulk)
        )
        st.dataframe(_bulk_table_display, use_container_width=True)
        _csv_bytes = _bulk_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="Download CSV",
            data=_csv_bytes,
            file_name="bulk_analysis.csv",
            mime="text/csv",
            key="download_bulk_analysis_csv",
        )