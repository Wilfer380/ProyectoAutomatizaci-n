from __future__ import annotations

import re


_TRAILING_DECIMAL_ZERO_RE = re.compile(r"^(?P<integer>[+-]?\d+)\.(?P<zeros>0+)$")


def normalize_excel_scalar(value) -> str:
    if value is None:
        return ""

    if isinstance(value, bool):
        return str(value)

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float):
        return _normalize_float(value)

    text = " ".join(str(value).replace("\xa0", " ").split()).strip()
    if not text:
        return ""

    match = _TRAILING_DECIMAL_ZERO_RE.match(text)
    if match:
        return match.group("integer")

    return text


def _normalize_float(value: float) -> str:
    if value.is_integer():
        return str(int(value))

    text = repr(value)
    if "e" in text.lower():
        text = f"{value:.15f}"

    text = text.rstrip("0").rstrip(".")
    return text
