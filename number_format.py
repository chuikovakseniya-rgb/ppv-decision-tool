"""
Единое отображение чисел для Streamlit UI (таблицы, heatmap-ячейки, OCR debug).

Меняйте формат здесь — не дублируйте toFixed/locale-логику в компонентах.

Основные функции (аналоги общих UI-имён):

- ``format_integer`` — целые; по умолчанию пробелы в тысячах.
- ``format_number`` — дробные / смешанные; параметры группировки и хвоста ``.00``.
- ``format_percent`` — проценты с суффиксом ``%``.
- ``format_delta`` — абсолютные отклонения (Potential Spendings ``diff``).

Дополнительно: ``format_matrix_metric``, ``format_potential_amount``,
``format_summary_percent_cell``, ``group_thousands_digits``, константа ``EMPTY``.

Виджеты ``st.number_input`` в Streamlit не поддерживают произвольное форматирование тысяч
в поле ввода — там остаётся «сырое» число; форматирование применяется ко всем таблицам
и текстовым представлениям, собранным через эти функции.
"""

from __future__ import annotations

import math
from typing import Any

import pandas as pd

# Отсутствие значения в таблицах
EMPTY = "—"

__all__ = [
    "EMPTY",
    "group_thousands_digits",
    "format_integer",
    "format_number",
    "format_percent",
    "format_delta",
    "format_potential_amount",
    "format_matrix_metric",
    "format_summary_percent_cell",
]


def _missing(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and math.isnan(value):
        return True
    return False


def _to_float(value: Any) -> float | None:
    if _missing(value):
        return None
    try:
        v = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(v):
        return None
    return v


def group_thousands_digits(digits: str) -> str:
    """Группы по 3 справа налево: ``26644`` → ``26 644`` (только цифры, без знака)."""
    if not digits:
        return "0"
    x = digits.lstrip("0") or "0"
    parts: list[str] = []
    while len(x) > 3:
        parts.append(x[-3:])
        x = x[:-3]
    parts.append(x)
    return " ".join(reversed(parts))


def format_integer(value: Any, *, group_thousands: bool = True) -> str:
    """
    Целые для UI. По умолчанию — пробел как разделитель тысяч (``26 644``).
    С ``group_thousands=False`` — компактная строка без пробелов (редкие случаи).
    """
    v = _to_float(value)
    if v is None:
        return EMPTY
    n = int(round(v))
    if group_thousands:
        neg = n < 0
        s = group_thousands_digits(str(abs(n)))
        return f"-{s}" if neg else s
    return str(n)


def format_number(
    value: Any,
    *,
    max_decimals: int = 2,
    group_thousands: bool = False,
    strip_trailing_zeros: bool = False,
    prefer_int_if_whole: bool = False,
) -> str:
    """
    Дробные / смешанные значения.

    - ``prefer_int_if_whole``: если число близко к целому — показать без ``.00`` (PPV matrix).
    - ``group_thousands`` + ``strip_trailing_zeros``: блок Fact/Could be в Potential Spendings.
    """
    if value is None:
        return EMPTY
    try:
        fv = float(value)
    except (TypeError, ValueError):
        return EMPTY
    if math.isnan(fv):
        return EMPTY

    if prefer_int_if_whole:
        tol = max(1e-4, 1e-9 * max(1.0, abs(fv)))
        if abs(fv - round(fv)) < tol:
            return format_integer(fv, group_thousands=group_thousands)

    if not group_thousands:
        return f"{fv:.{max_decimals}f}"

    if abs(fv) < 1e-15:
        return "0"

    neg = fv < 0
    av = abs(round(fv + 1e-12, max_decimals))
    raw = f"{av:.{max_decimals}f}"
    if strip_trailing_zeros:
        raw = raw.rstrip("0").rstrip(".")
    if "." in raw:
        ip, fp = raw.split(".", 1)
        body = group_thousands_digits(ip) + "." + fp
    else:
        body = group_thousands_digits(raw)
    return ("-" if neg else "") + body


def format_percent(value: Any, *, zero_display: str = "0.00%") -> str:
    """
    Проценты с суффиксом ``%``; положительные без ``+``, отрицательные с ``-``.
    ``zero_display``: для колонки diff % в Potential Spendings можно передать ``"0"``.
    """
    if _missing(value):
        return EMPTY
    try:
        v = float(value)
    except (TypeError, ValueError):
        return EMPTY
    if abs(v) < 1e-15:
        return zero_display
    return f"{v:.2f}%"


def format_delta(value: Any, *, as_integer: bool = False) -> str:
    """
    Абсолютная разница (Fact − Could be и т.п.): минус только у отрицательных;
    при ``as_integer`` — целое с группировкой тысяч; иначе до 2 знаков с группировкой.
    """
    v = _to_float(value)
    if v is None:
        return EMPTY
    if abs(v) < 1e-15:
        return "0"
    neg = v < 0
    av = abs(v)
    if as_integer:
        n = int(round(av))
        body = group_thousands_digits(str(n))
        return ("-" if neg else "") + body
    av = round(av + 1e-12, 2)
    raw = f"{av:.2f}".rstrip("0").rstrip(".")
    if "." in raw:
        ip, fp = raw.split(".", 1)
        body = group_thousands_digits(ip) + "." + fp
    else:
        body = group_thousands_digits(raw)
    return ("-" if neg else "") + body


def format_potential_amount(value: Any) -> str:
    """Fact / Could be в Potential Spendings: пробелы в тысячах, до 2 дробных, без хвоста ``.00``."""
    return format_number(
        value,
        max_decimals=2,
        group_thousands=True,
        strip_trailing_zeros=True,
        prefer_int_if_whole=False,
    )


def format_matrix_metric(value: Any, *, as_int: bool) -> str:
    """
    Абсолютные метрики в PPV matrix, сводке «До»/«После», OCR debug:
    те же правила, что Fact/Could be в Potential Spendings (пробелы в тысяцах).
    """
    if value is None:
        return EMPTY
    if as_int:
        return format_integer(value, group_thousands=True)
    return format_number(
        value,
        max_decimals=2,
        group_thousands=True,
        strip_trailing_zeros=True,
        prefer_int_if_whole=True,
    )


def format_summary_percent_cell(value: Any) -> str:
    """Колонки «До %» / «После %» в сводке (Styler + heatmap)."""
    return format_percent(value, zero_display="0.00%")
