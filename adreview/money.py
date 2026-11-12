"""Aritmetica decimale, parsing e formattazione di importi/percentuali in formato italiano."""
from __future__ import annotations

import re
from decimal import (ROUND_DOWN, ROUND_HALF_EVEN, ROUND_HALF_UP, ROUND_UP, Decimal,
                     InvalidOperation)

ROUNDING_MODES = {
    "half_up": ROUND_HALF_UP,
    "half_even": ROUND_HALF_EVEN,
    "down": ROUND_DOWN,
    "up": ROUND_UP,
}

NUM = r"\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+(?:,\d+)?|\d+\.\d+"
NUM_RE = re.compile(NUM)
CURRENCY_RE = re.compile(
    r"(?:€|EUR|euro)\s*(?P<a>" + NUM + r")|(?P<b>" + NUM + r")\s*(?:€|EUR\b|euro\b)", re.I)

D0 = Decimal(0)


def D(x) -> Decimal:
    if isinstance(x, Decimal):
        return x
    return Decimal(str(x))


def parse_number(text: str | None) -> Decimal | None:
    """'1.234,56' -> 1234.56 ; '€ 100' -> 100 ; '12,5 %' -> 12.5. None se non interpretabile."""
    if text is None:
        return None
    t = re.sub(r"[€%\s]|EUR|euro|eur", "", str(text), flags=re.I)
    t = t.replace("−", "-").replace(" ", "")
    if not t or not re.fullmatch(r"-?[\d.,]+", t):
        return None
    neg = t.startswith("-")
    t = t.lstrip("-")
    if "," in t:
        ip, _, fp = t.rpartition(",")
        ip = ip.replace(".", "")
        if not ip.isdigit() and ip != "" or not fp.isdigit():
            return None
        t = f"{ip or '0'}.{fp}"
    elif "." in t:
        if re.fullmatch(r"\d{1,3}(\.\d{3})+", t):
            t = t.replace(".", "")
        elif t.count(".") > 1:
            return None
    try:
        v = Decimal(t)
    except InvalidOperation:
        return None
    return -v if neg else v


def parse_percent(text: str | None) -> Decimal | None:
    if text is None:
        return None
    m = re.search(r"(-?\d+(?:[.,]\d+)?)\s*%", str(text))
    if m:
        return parse_number(m.group(1))
    return None


def quantize(d: Decimal, decimals: int = 2, mode: str = "half_up") -> Decimal:
    q = Decimal(1).scaleb(-decimals)
    return d.quantize(q, rounding=ROUNDING_MODES.get(mode, ROUND_HALF_UP))


def fmt_num(d: Decimal | None, decimals: int = 2) -> str:
    """Formato italiano senza simbolo: 1.234,56"""
    if d is None:
        return "—"
    d = quantize(d, decimals) if decimals is not None else d
    s = f"{abs(d):,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if d < 0 else "") + s


def fmt_eur(d: Decimal | None, decimals: int = 2) -> str:
    return "—" if d is None else f"{fmt_num(d, decimals)} €"


def fmt_pct(d: Decimal | None, decimals: int = 2) -> str:
    return "n.a." if d is None else f"{fmt_num(d, decimals)}%"


def find_numbers(text: str) -> list[tuple[str, Decimal]]:
    out = []
    for m in NUM_RE.finditer(text or ""):
        v = parse_number(m.group(0))
        if v is not None:
            out.append((m.group(0), v))
    return out


def find_amounts(text: str) -> list[tuple[str, Decimal]]:
    """Importi con simbolo di valuta (€, EUR, euro) in un testo."""
    out = []
    for m in CURRENCY_RE.finditer(text or ""):
        raw = m.group("a") or m.group("b")
        v = parse_number(raw)
        if v is not None:
            out.append((raw, v))
    return out


def replace_number_token(cell_text: str, new_value: Decimal, decimals: int | None = None) -> tuple[str, str] | None:
    """Sostituisce il primo token numerico di una cella mantenendo prefisso/suffisso (€, %, ecc.).
    Ritorna (token_originale, nuovo_token) oppure None."""
    m = NUM_RE.search(cell_text or "")
    if not m:
        return None
    tok = m.group(0)
    if decimals is None:
        decimals = len(tok.split(",")[1]) if "," in tok else (0 if new_value == new_value.to_integral_value() else 2)
    return tok, fmt_num(new_value, decimals)
