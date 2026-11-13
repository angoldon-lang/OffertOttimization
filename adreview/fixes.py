"""Mancanze con correzione proposta: celle senza valore, destinatario incompleto, date ripetute, punti elenco vuoti.
Ogni rilievo con un punto di applicazione nel Word porta una modifica approvabile (eventualmente con valore da inserire)."""
from __future__ import annotations

import calendar
import re

from .docx_utils import Change, DocView, Location
from .extract import DATE_RE, FieldResult
from .findings import Finding


def loc_change(loc: Location, text: str, before: str, after: str, label: str, category: str = "testo", reason: str = "",
               needs_value: bool = False) -> Change | None:
    if loc.kind == "par" and loc.paragraph is not None:
        return Change("para", before, after, label, category, para=loc.paragraph, ctx_before=text, reason=reason, editable=True, needs_value=needs_value)
    if loc.kind == "cell" and loc.table is not None:
        return Change("cell", before, after, label, category, table=loc.table, row=loc.row, col=loc.col, ctx_before=text, reason=reason,
                      editable=True, needs_value=needs_value)
    if loc.kind == "box":
        return Change("box", before, after, label, category, ctx_before=text, reason=reason, editable=True, needs_value=needs_value)
    return None


def suggested_duration(fields: dict, econ) -> str:
    m = (fields["durata"].data.get("months") if fields.get("durata") else None) or getattr(econ, "span_months", None)
    if not m:
        return ""
    return f"{m // 12} anni" if m % 12 == 0 else f"{m} mesi"


def check_completeness(view: DocView, fields: dict[str, FieldResult], econ) -> list[Finding]:
    out: list[Finding] = []
    # 1) righe «ETICHETTA | (vuoto)»
    for t in view.tables:
        for r, cells in enumerate(t.rows):
            if len(cells) != 2:
                continue
            label, value = cells[0].text.strip(), cells[1].text.strip()
            if not label or value or len(label) > 70 or re.search(r"\d", label):
                continue
            if not (label.isupper() or label.endswith(":") or len(label.split()) <= 4):
                continue
            sug = suggested_duration(fields, econ) if re.search(r"(?i)validit[àa]\s+(del\s+)?contratto|durata", label) else ""
            loc = Location(t.section, table=t.idx, row=r, col=cells[1].col, kind="cell")
            ch = loc_change(loc, "", "", sug, f"Valore di «{label}»", "clausola", "completamento di un campo vuoto", needs_value=True)
            out.append(Finding("campi", f"Campo «{label}» senza valore", "Errore", loc,
                               f"La riga «{label}» non ha alcun valore.",
                               "Inserire il valore" + (f" (suggerito dal documento: {sug})" if sug else ""), ch, kind="sostanziale", needs_input=True))
    # 2) destinatario «C.A. di» senza nominativo; 3) date ripetute
    seen = set()
    for loc, text in list(view.scan()) + list(view.scan_extra()):
        if (loc.label, text) in seen:
            continue
        seen.add((loc.label, text))
        if re.search(r"\bC\.?\s?A\.?\s+di\s*$", text):
            ch = loc_change(loc, text, text.rstrip(), text.rstrip() + " ", "Destinatario (C.A.)", "clausola", "inserire il nominativo", needs_value=True)
            out.append(Finding("campi", "Destinatario «C.A.» senza nominativo", "Errore", loc, f"«{text}» termina senza il nome del referente.",
                               "Completare con il nominativo (scrivilo nel campo)", ch, kind="sostanziale", needs_input=True))
        m = re.search(r"(?P<d>" + DATE_RE + r")(?P<rep>(?:\s+(?P=d))+)", text)
        if m:
            seq = m.group(0).strip()
            ch = loc_change(loc, text, seq, m.group("d"), "Data ripetuta", "data", "la stessa data compare due volte di seguito")
            out.append(Finding("campi", "Data ripetuta due volte", "Errore", loc, f"«{seq}»", f"Lasciare una sola volta «{m.group('d')}»", ch))
    # 4) punti elenco vuoti «…»
    for p in view.paras:
        if p.text.strip() in ("…", "..."):
            loc = Location(p.section, paragraph=p.idx, kind="par")
            ch = loc_change(loc, p.text, p.text.strip(), "", "Punto elenco vuoto", "clausola", "sostituire il segnaposto con il contenuto", needs_value=True)
            out.append(Finding("segnaposto", "Punto elenco vuoto («…»)", "Da verificare", loc, "Il punto elenco contiene solo «…».",
                               "Scrivere il contenuto nel campo (o lasciare vuoto per svuotare il punto)", ch, kind="sostanziale", needs_input=True))
    # 5) rinnovo automatico senza durata
    r = fields.get("rinnovo")
    if r and r.present and "automatico" in r.data.get("types", []) and not r.data.get("renew_months") and not fields["durata"].present:
        out.append(Finding("campi", "Rinnovo automatico senza durata del contratto né del rinnovo", "Da verificare", r.loc or Location(kind="doc"),
                           "Il rinnovo è automatico ma non è indicata la durata del contratto né il periodo di rinnovo.",
                           "Indicare durata del contratto e durata del rinnovo", kind="sostanziale", needs_input=True))
    return out


def invalid_date_fix(raw: str, others: list[str]) -> str:
    """Suggerisce una data valida: un'altra data di validità del documento, oppure l'ultimo giorno del mese indicato."""
    if others:
        return others[0]
    m = re.fullmatch(r"(\d{1,2})(\s*[/\-.]\s*)(\d{1,2})(\s*[/\-.]\s*)(\d{2,4})", raw.strip())
    if m:
        try:
            last = calendar.monthrange(int(m.group(5)) if len(m.group(5)) == 4 else 2000 + int(m.group(5)), int(m.group(3)))[1]
            return f"{last}{m.group(2)}{m.group(3)}{m.group(4)}{m.group(5)}"
        except (ValueError, calendar.IllegalMonthError):
            return raw
    return raw
