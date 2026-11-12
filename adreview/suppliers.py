"""Estrazione delle righe di costo dalle quotazioni fornitore (XLSX, PDF testuale, PDF scansionato con OCR opzionale)."""
from __future__ import annotations

import io
import re
import shutil
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal

from .docx_utils import norm
from .econ import detect_periodicity, header_roles
from .extract import DATE_RE, find_dates, parse_date
from .money import D, find_numbers, parse_number


@dataclass
class CostLine:
    id: str
    source: str
    source_file: str = ""
    category: str = "fornitore"      # fornitore | interno
    sku: str = ""
    desc: str = ""
    qty: Decimal | None = None
    unit_cost: Decimal | None = None
    total: Decimal | None = None
    currency: str = "EUR"
    periodicity: str | None = None    # una_tantum|mensile|trimestrale|semestrale|annuale|periodo_quotazione
    period_months: int | None = None
    vat_included: bool = False
    vat_pct: Decimal | None = None
    base_confirmed: bool = False
    uncertain: bool = False
    note: str = ""
    manual: bool = False

    @property
    def total_effective(self) -> Decimal | None:
        if self.total is not None:
            return self.total
        if self.qty is not None and self.unit_cost is not None:
            return self.qty * self.unit_cost
        return None


@dataclass
class QuoteResult:
    filename: str
    lines: list[CostLine] = field(default_factory=list)
    valid_until: date | None = None
    currency: str | None = None
    vat_included: bool = False
    duration_months: int | None = None
    warnings: list[str] = field(default_factory=list)
    needs_manual: bool = False


def _dec(v) -> Decimal | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float, Decimal)):
        return Decimal(str(v))
    n = find_numbers(str(v))
    return parse_number(n[0][0]) if n else None


def _period(text: str) -> tuple[str | None, int | None]:
    p = detect_periodicity(text)
    if p == "contratto":
        m = re.search(r"(\d{1,3})\s*(mesi|anni)", norm(text))
        months = (int(m.group(1)) * (12 if m.group(2) == "anni" else 1)) if m else None
        return "periodo_quotazione", months
    return p, None


def parse_rows(rows: list[list], prefix: str, fname: str, res: QuoteResult, uncertain: bool = False) -> None:
    """rows: lista di righe (liste di valori). Cerca intestazione e righe voce."""
    hdr_i, roles = None, {}
    for i, row in enumerate(rows[:60]):
        txts = [("" if v is None else str(v)) for v in row]
        rl = header_roles(txts)
        for j, t in enumerate(txts):
            if re.search(r"valuta|currency|divisa", norm(t)) and j not in rl:
                rl[j] = "currency"
        if len(rl) >= 3 and ({"amount", "unit"} & set(rl.values())):
            hdr_i, roles = i, rl
            break
    # metadati nelle righe sopra l'intestazione (o in tutto il foglio se intestazione non trovata)
    meta_rows = rows[: hdr_i] if hdr_i is not None else rows
    for row in meta_rows:
        cells = [v for v in row if v not in (None, "")]
        line = " ".join(c.strftime("%d/%m/%Y") if isinstance(c, (datetime, date)) else str(c) for c in cells)
        if not line:
            continue
        nl = norm(line)
        if re.search(r"validit|valid until|scadenza", nl):
            for c in cells:
                if isinstance(c, (datetime, date)):
                    res.valid_until = c.date() if isinstance(c, datetime) else c
                    break
            else:
                ds = find_dates(line)
                if ds and ds[0][1]:
                    res.valid_until = ds[0][1]
        if re.search(r"valuta|currency|divisa", nl):
            m = re.search(r"\b(EUR|USD|GBP|CHF)\b", line.upper())
            if m:
                res.currency = m.group(1)
        if re.search(r"iva (inclusa|compresa)|ivato", nl):
            res.vat_included = True
        if re.search(r"durata|periodo di fornitura", nl):
            m = re.search(r"(\d{1,3})\s*(mesi|anni)", nl)
            if m:
                res.duration_months = int(m.group(1)) * (12 if m.group(2) == "anni" else 1)
    if hdr_i is None:
        res.warnings.append(f"{prefix}: intestazione della tabella dei costi non riconosciuta; usare la griglia manuale.")
        res.needs_manual = True
        return
    by_role = {v: k for k, v in roles.items()}
    blank = 0
    for i in range(hdr_i + 1, len(rows)):
        row = rows[i]
        def g(role):
            j = by_role.get(role)
            return row[j] if j is not None and j < len(row) else None
        texts = [("" if v is None else str(v)) for v in row]
        if not any(t.strip() for t in texts):
            blank += 1
            if blank >= 3:
                break
            continue
        blank = 0
        first = next((t for t in texts if t.strip()), "")
        if re.match(r"(?i)^\s*(sub)?totale|^\s*iva\b|^\s*imponibile", first):
            continue
        qty, unit, tot = _dec(g("qty")), _dec(g("unit")), _dec(g("amount"))
        if qty is None and unit is None and tot is None:
            continue
        cl = CostLine(f"C{len(res.lines) + 1}:{fname}", f"{fname} › {prefix} › riga {i + 1}", fname,
                      sku=str(g("sku") or "").strip(), desc=str(g("desc") or "").strip(), qty=qty, unit_cost=unit, total=tot,
                      uncertain=uncertain)
        cur = g("currency")
        cl.currency = (str(cur).strip().upper()[:3] if cur else (res.currency or "EUR"))
        pc = str(g("period") or "")
        cl.periodicity, cl.period_months = _period(pc) if pc else (None, None)
        if cl.periodicity is None:
            cl.periodicity, cl.period_months = _period(cl.desc)
        if cl.periodicity is None and res.duration_months and False:
            pass
        cl.vat_included = res.vat_included
        if qty is not None and unit is not None and tot is not None and abs(qty * unit - tot) > Decimal("0.01"):
            cl.note = f"Totale riportato {tot} ≠ quantità × costo unitario {qty * unit}"
            cl.uncertain = True
        if uncertain:
            cl.note = (cl.note + "; " if cl.note else "") + "estrazione da testo/OCR: verificare"
        res.lines.append(cl)
    if not res.lines:
        res.warnings.append(f"{prefix}: nessuna riga di costo riconosciuta; usare la griglia manuale.")
        res.needs_manual = True


def parse_xlsx(data: bytes, filename: str) -> QuoteResult:
    from openpyxl import load_workbook
    res = QuoteResult(filename)
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception as exc:
        res.warnings.append(f"File XLSX non leggibile: {exc}")
        res.needs_manual = True
        return res
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if not any(any(c is not None for c in r) for r in rows):
            continue
        n_before = len(res.lines)
        parse_rows(rows, ws.title, filename, res)
        if len(res.lines) > n_before:
            res.needs_manual = False
    try:
        wbf = load_workbook(io.BytesIO(data), data_only=False)
        for ws in wbf.worksheets:
            if any(isinstance(c.value, str) and c.value.startswith("=") for r in ws.iter_rows() for c in r):
                if any(cl.total is None for cl in res.lines):
                    res.warnings.append("Alcune celle contengono formule senza valore memorizzato: aprire e salvare il file in Excel o inserire i valori a mano.")
                break
    except Exception:
        pass
    return res


LINE_RE = re.compile(r"^(?P<sku>[A-Z0-9][\w\-./]{2,})\s+(?P<desc>.+?)\s+(?P<qty>\d+(?:[.,]\d+)?)\s+(?P<unit>\d[\d.,]*)\s+(?P<tot>\d[\d.,]*)(?:\s+(?P<per>[A-Za-zà-ù ]+))?$")


def _ocr_available() -> bool:
    try:
        import pytesseract  # noqa
        return shutil.which("tesseract") is not None
    except Exception:
        return False


def parse_pdf(data: bytes, filename: str) -> QuoteResult:
    import pdfplumber
    res = QuoteResult(filename)
    all_text = ""
    try:
        pdf = pdfplumber.open(io.BytesIO(data))
    except Exception as exc:
        res.warnings.append(f"PDF non leggibile: {exc}")
        res.needs_manual = True
        return res
    with pdf:
        for pn, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            all_text += "\n" + text
            tables = page.extract_tables() or []
            for ti, tb in enumerate(tables, 1):
                parse_rows([[c for c in r] for r in tb], f"pag. {pn} tab. {ti}", filename, res)
        if not all_text.strip():
            if _ocr_available():
                import pytesseract
                for pn, page in enumerate(pdf.pages, 1):
                    img = page.to_image(resolution=200).original
                    all_text += "\n" + pytesseract.image_to_string(img, lang="ita+eng")
                res.warnings.append("PDF scansionato: testo ottenuto con OCR, tutte le righe sono da verificare.")
                _text_lines(all_text, filename, res, uncertain=True)
            else:
                res.warnings.append("PDF scansionato o senza testo: OCR non disponibile (installare tesseract e pytesseract). Inserire i costi con la griglia manuale.")
                res.needs_manual = True
            _meta_from_text(all_text, res)
            return res
    _meta_from_text(all_text, res)
    if not res.lines:
        _text_lines(all_text, filename, res, uncertain=True)
    if not res.lines:
        res.warnings.append("Nessuna riga di costo riconosciuta nel PDF: usare la griglia manuale.")
        res.needs_manual = True
    return res


def _meta_from_text(text: str, res: QuoteResult) -> None:
    m = re.search(r"(?i)validit[àa][^\n\d]{0,40}(" + DATE_RE + ")", text)
    if m and not res.valid_until:
        res.valid_until = parse_date(m.group(1))
    m = re.search(r"\b(EUR|USD|GBP|CHF)\b", text)
    if m and not res.currency:
        res.currency = m.group(1)
    if re.search(r"(?i)iva\s+(inclusa|compresa)", text):
        res.vat_included = True
        for l in res.lines:
            l.vat_included = True


def _text_lines(text: str, fname: str, res: QuoteResult, uncertain: bool) -> None:
    for i, line in enumerate(text.splitlines(), 1):
        m = LINE_RE.match(line.strip())
        if not m:
            continue
        per, pm = _period(m.group("per") or "")
        cl = CostLine(f"C{len(res.lines) + 1}:{fname}", f"{fname} › testo › riga {i}", fname, sku=m.group("sku"), desc=m.group("desc"),
                      qty=parse_number(m.group("qty")), unit_cost=parse_number(m.group("unit")), total=parse_number(m.group("tot")),
                      currency=res.currency or "EUR", periodicity=per, period_months=pm, uncertain=uncertain,
                      note="estrazione da testo libero: verificare")
        res.lines.append(cl)


def parse_quotation(filename: str, data: bytes) -> QuoteResult:
    low = filename.lower()
    if low.endswith((".xlsx", ".xlsm")):
        return parse_xlsx(data, filename)
    if low.endswith(".pdf"):
        return parse_pdf(data, filename)
    r = QuoteResult(filename)
    r.warnings.append("Formato non supportato (usare XLSX o PDF): inserire i costi con la griglia manuale.")
    r.needs_manual = True
    return r
