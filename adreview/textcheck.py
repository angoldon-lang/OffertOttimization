"""Controlli di testo: refusi, grammatica, segnaposto, coerenza di denominazioni/riferimenti e importi nel testo."""
from __future__ import annotations

import re
from datetime import date
from decimal import Decimal

from .config import Profile
from .docx_utils import Change, DocView, Location, norm
from .findings import Finding
from .money import fmt_num, find_amounts, parse_number

BUILTIN_TYPOS = {
    "sopratutto": "soprattutto", "perchè": "perché", "perche": "perché", "poichè": "poiché", "finchè": "finché",
    "affinchè": "affinché", "anzichè": "anziché", "benchè": "benché", "nonchè": "nonché", "dopodichè": "dopodiché",
    "propio": "proprio", "effetuare": "effettuare", "un'altro": "un altro", "qual'è": "qual è", "pò": "po'",
    "accellerare": "accelerare", "eccezzionale": "eccezionale", "obbiettivo": "obiettivo", "piu": "più", "gia": "già",
    "puo": "può", "cosi": "così", "quantita": "quantità", "attivita": "attività", "e'": "è", "sè stesso": "sé stesso",
    "avvolte": "a volte", "daccordo": "d'accordo", "pultroppo": "purtroppo", "sopratutto,": "soprattutto,",
}

PLACEHOLDER_PATTERNS = [
    (r"\[[^\]\n]*[A-Za-zÀ-ÿ][^\]\n]*\]", "segnaposto tra parentesi quadre"),
    (r"\{\{[^}\n]+\}\}", "campo di unione non compilato"),
    (r"<[A-Za-z][^<>\n]{1,40}>", "segnaposto tra parentesi angolari"),
    (r"\b[Xx]{3,}\b", "valore «XXX» da completare"),
    (r"(?i)lorem ipsum", "testo di esempio (lorem ipsum)"),
    (r"\b(?:TBD|TODO|DA DEFINIRE|DA COMPLETARE)\b", "elemento da definire"),
    (r"(?i)\bda\s+(?:inserire|compilare|completare|definire|confermare)\b", "indicazione «da compilare»"),
    (r"_{4,}", "campo vuoto da compilare"),
    (r"\.{6,}", "campo puntinato da compilare"),
    (r"€\s*[Xx?]{2,}", "importo da completare"),
]


def units(view: DocView):
    """(testo, Location, parametri-change) per ogni paragrafo e cella (i paragrafi di cella sono uniti)."""
    for p in view.paras:
        if p.text.strip():
            yield p.text, Location(p.section, paragraph=p.idx, kind="par"), {"kind": "para", "para": p.idx}
    for t in view.tables:
        for r, row in enumerate(t.rows):
            for c in row:
                for ptxt in c.text.split("\n"):
                    if ptxt.strip():
                        yield ptxt, Location(t.section, table=t.idx, row=r, col=c.col, kind="cell"), \
                            {"kind": "cell", "table": t.idx, "row": r, "col": c.col}


def _case(src: str, rep: str) -> str:
    if src[:1].isupper() and rep[:1].islower():
        return rep[0].upper() + rep[1:]
    return rep


def _snippet(text: str, s: int, e: int, replacement: str, ctx: int = 18) -> tuple[str, str]:
    l, r = max(0, s - ctx), min(len(text), e + ctx)
    while l > 0 and not text[l - 1].isspace():
        l -= 1
    while r < len(text) and not text[r].isspace():
        r += 1
    return text[l:r], text[l:s] + replacement + text[e:r]


def _mk(params, before, after, label, category="testo", reason="", ctx=""):
    return Change(params["kind"], before, after, label, category, table=params.get("table"), row=params.get("row"),
                  col=params.get("col"), para=params.get("para"), ctx_before=ctx, reason=reason)


def check_typos(view: DocView, profile: Profile) -> list[Finding]:
    typos = {**BUILTIN_TYPOS, **{k.lower(): v for k, v in (profile.typos or {}).items()}}
    rx = re.compile(r"(?<![\w'’])(" + "|".join(re.escape(k) for k in sorted(typos, key=len, reverse=True)) + r")(?![\w])", re.I)
    out, seen = [], set()
    for text, loc, params in units(view):
        for m in rx.finditer(text):
            rep = _case(m.group(1), typos[m.group(1).lower()])
            b, a = _snippet(text, m.start(1), m.end(1), rep)
            if (loc.label, b) in seen:
                continue
            seen.add((loc.label, b))
            out.append(Finding("testo", "Refuso", "Errore", loc, f"«{m.group(1)}» in «…{b}…»",
                               f"Correggere in «{rep}»", _mk(params, b, a, "Refuso", reason=f"{m.group(1)} → {rep}", ctx=text)))
        for m in re.finditer(r"(?<=\S)  +(?=\S)", text):
            b, a = _snippet(text, m.start(), m.end(), " ", 12)
            if (loc.label, b) not in seen:
                seen.add((loc.label, b))
                out.append(Finding("testo", "Spazi doppi", "Errore", loc, f"«…{b}…»", "Ridurre a uno spazio",
                                   _mk(params, b, a, "Spazi doppi", reason="spazio doppio", ctx=text), kind="grafica"))
        for m in re.finditer(r"\b([A-Za-zÀ-ÿ]{2,})\s+\1\b", text, re.I):
            b, a = _snippet(text, m.start(), m.end(), m.group(1), 12)
            if (loc.label, b) not in seen:
                seen.add((loc.label, b))
                out.append(Finding("testo", "Parola ripetuta", "Errore", loc, f"«…{b}…»", f"Eliminare la ripetizione di «{m.group(1)}»",
                                   _mk(params, b, a, "Parola ripetuta", reason="ripetizione", ctx=text)))
        for m in re.finditer(r"(?<=[A-Za-zÀ-ÿ0-9])\s+([,;:!?])(?=\s|$)", text):
            b, a = _snippet(text, m.start(), m.end(), m.group(1), 12)
            if (loc.label, b) not in seen:
                seen.add((loc.label, b))
                out.append(Finding("testo", "Spazio prima della punteggiatura", "Errore", loc, f"«…{b}…»", "Rimuovere lo spazio",
                                   _mk(params, b, a, "Punteggiatura", reason="spazio prima di punteggiatura", ctx=text)))
        for m in re.finditer(r"(?<=[a-zà-ù]),(?=[A-Za-zÀ-ÿ])|(?<=[a-zà-ù])\.(?=[A-ZÀ-Ý][a-zà-ù])", text):
            tok_s = text.rfind(" ", 0, m.start()) + 1
            tok_e = text.find(" ", m.end())
            tok = text[tok_s: tok_e if tok_e > 0 else len(text)]
            if re.search(r"[@/]|www|\.[a-z]{2,4}\b", tok):
                continue
            b, a = _snippet(text, m.start(), m.end(), m.group(0) + " ", 12)
            if (loc.label, b) not in seen:
                seen.add((loc.label, b))
                out.append(Finding("testo", "Spazio mancante dopo la punteggiatura", "Errore", loc, f"«…{b}…»", "Inserire uno spazio",
                                   _mk(params, b, a, "Punteggiatura", reason="spazio mancante", ctx=text)))
    return out


def check_placeholders(view: DocView, extra_texts: list[tuple[str, str]] | None = None) -> list[Finding]:
    out = []
    for loc, text in view.scan():
        for pat, what in PLACEHOLDER_PATTERNS:
            for m in re.finditer(pat, text):
                out.append(Finding("segnaposto", "Segnaposto rimasto", "Errore", loc, f"«{m.group(0)}» – {what}",
                                   "Compilare o rimuovere (nessuna modifica automatica: il contenuto va deciso dall'utente)", kind="sostanziale"))
    for where, text in (extra_texts or []):
        for pat, what in PLACEHOLDER_PATTERNS:
            for m in re.finditer(pat, text):
                out.append(Finding("segnaposto", "Segnaposto rimasto in " + where, "Errore", Location(kind=where),
                                   f"«{m.group(0)}» – {what}", "Compilare o rimuovere", kind="sostanziale"))
    return out


LEGAL = r"(?:S\.?\s?p\.?\s?A\.?|S\.?\s?r\.?\s?l\.?|S\.?\s?n\.?\s?c\.?|S\.?\s?a\.?\s?s\.?)"
ENTITY_RE = re.compile(r"\b((?:[A-ZÀ-Ý][\w&'’\-]*\.?)(?:\s+(?:[A-ZÀ-Ý][\w&'’\-]*\.?|&|e|di|de|del))*?)\s+(" + LEGAL + r")(?![\w])")
STRIP_LEAD = re.compile(r"^(?:Spett\.?le|Spettabile|Cliente|Oggetto|Gentile|Egregio|Alla|Alle|Della|Dalla|Società|Ditta|Il|La|Lo|Le|Gli|Un|Una|Per|Con|Tra|Da|Di|Fra|Presso)\.?\s+", re.I)


def _ent_key(name: str, legal: str) -> str:
    return re.sub(r"[^a-z0-9]", "", norm(name)) + re.sub(r"[^a-z]", "", legal.lower())


def _split_entity(m) -> tuple[str, str, int]:
    name = m.group(1)
    start = m.start(1)
    while True:
        s = STRIP_LEAD.match(name)
        if not s:
            break
        name = name[s.end():]
        start += s.end()
    return name, m.group(2), start


def check_consistency(view: DocView, profile: Profile, fields: dict, econ, today: date, corrections: dict | None = None) -> list[Finding]:
    out: list[Finding] = []
    # --- denominazione AD Consulting
    own = profile.company_name
    legal_m = re.search(LEGAL + r"\s*$", own)
    core = own[: legal_m.start()].strip() if legal_m else own
    words = core.split()
    pat_words = []
    for w in words:
        if w.isupper() and 2 <= len(w) <= 3:
            pat_words.append(r"\.?".join(re.escape(ch) for ch in w) + r"\.?")
        else:
            pat_words.append(re.escape(w))
    own_rx = re.compile(r"\b" + r"\s?".join(pat_words[:1]) + r"\s+" + r"\s+".join(pat_words[1:]) + (r"(?:\s+" + LEGAL + r")?" if legal_m else ""), re.I) if len(pat_words) > 1 else None
    seen = set()
    if own_rx:
        for text, loc, params in units(view):
            for m in own_rx.finditer(text):
                occ = m.group(0)
                if occ == own or not re.search(LEGAL + r"$", occ):
                    continue
                b, a = _snippet(text, m.start(), m.end(), own)
                if (loc.label, b) in seen:
                    continue
                seen.add((loc.label, b))
                out.append(Finding("coerenza", "Denominazione aziendale non uniforme", "Da verificare", loc,
                                   f"«{occ}» invece di «{own}»", f"Uniformare a «{own}»",
                                   _mk(params, b, a, "Denominazione", "clausola", f"{occ} → {own}", text)))
    # --- denominazione cliente / altre società
    client = fields["cliente"]
    ckey, cname = None, None
    if client.occurrences:
        cm = ENTITY_RE.search(client.occurrences[0].value)
        if cm:
            n, lg, _ = _split_entity(cm)
            ckey, cname = _ent_key(n, lg), f"{n} {lg}"
    known = [norm(k) for k in profile.known_entities]
    own_key = re.sub(r"[^a-z0-9]", "", norm(core))
    for text, loc, params in units(view):
        for m in ENTITY_RE.finditer(text):
            n, lg, start = _split_entity(m)
            key = _ent_key(n, lg)
            raw = text[start: m.end()]
            nk = re.sub(r"[^a-z0-9]", "", norm(n))
            if nk.startswith(own_key) or any(k and k in norm(n) for k in known):
                continue
            if ckey and key == ckey:
                if raw != client.occurrences[0].value and raw != cname and norm(raw) != norm(client.occurrences[0].value):
                    b, a = _snippet(text, start, m.end(), client.occurrences[0].value)
                    out.append(Finding("coerenza", "Denominazione del cliente non uniforme", "Da verificare", loc,
                                       f"«{raw}» invece di «{client.occurrences[0].value}»", "Uniformare alla denominazione del campo Cliente",
                                       _mk(params, b, a, "Denominazione cliente", "clausola", "denominazione cliente", text)))
            elif ckey:
                if re.sub(r"[^a-z0-9]", "", norm(client.occurrences[0].value)).find(nk) >= 0 or nk.find(ckey[: max(4, len(ckey) - 4)]) >= 0:
                    continue
                if (loc.label, raw) in seen:
                    continue
                seen.add((loc.label, raw))
                out.append(Finding("coerenza", "Riferimento a una società diversa dal cliente", "Da verificare", loc,
                                   f"«{raw}» (cliente dell'offerta: «{client.occurrences[0].value}»)",
                                   "Verificare che il riferimento sia corretto; non modificato automaticamente"))
    # --- aliquote IVA citate nel testo
    rates = {}
    for text, loc, params in units(view):
        for m in re.finditer(r"(?i)\biva\b[^.\n%]{0,25}?(\d{1,2}(?:,\d+)?)\s*%", text):
            rates.setdefault(m.group(1), loc)
    if econ.vat_rate is not None:
        rates.setdefault(fmt_num(econ.vat_rate, 0 if econ.vat_rate == econ.vat_rate.to_integral() else 2), econ.vat_source or Location())
    if len(rates) > 1:
        loc = next(iter(rates.values()))
        out.append(Finding("coerenza", "Aliquote IVA discordanti", "Da verificare", loc,
                           "Aliquote citate: " + ", ".join(f"{r}%" for r in rates), "Verificare l'aliquota applicabile"))
    # --- importi citati nel testo ("totale ... € X") vs tabelle
    reported_to_calc, known_vals = {}, set()
    for l in econ.lines:
        for v in (l.reported, l.calc):
            if v is not None:
                known_vals.add(v)
    for t in econ.totals:
        for v in (t.reported, t.calc):
            if v is not None:
                known_vals.add(v)
        if t.reported is not None and t.calc is not None and t.status == "Errore":
            reported_to_calc.setdefault(t.reported, t)
    for p in view.paras:
        if not re.search(r"(?i)totale|complessiv|importo|canone|corrispettivo", p.text):
            continue
        for raw, v in find_amounts(p.text):
            loc = Location(p.section, paragraph=p.idx, kind="par")
            if v in reported_to_calc:
                t = reported_to_calc[v]
                tok = raw
                new_tok = fmt_num(t.calc, 2)
                ch = Change("para", tok, new_tok, "Importo citato nel testo", "totale", para=p.idx, ctx_before=p.text,
                            reason=f"allineato al totale ricalcolato di {t.label}")
                out.append(Finding("calcoli", "Importo nel testo non coerente con il ricalcolo", "Errore", loc,
                                   f"Testo: {fmt_num(v)} € · totale ricalcolato «{t.label}»: {fmt_num(t.calc)} €",
                                   f"Sostituire {fmt_num(v)} con {fmt_num(t.calc)}", ch))
            elif v not in known_vals and corrections and v in corrections:
                new_tok = fmt_num(corrections[v], 2)
                ch = Change("para", raw, new_tok, "Importo citato nel testo", "totale", para=p.idx, ctx_before=p.text,
                            reason="allineato a una correzione già approvata nelle tabelle")
                out.append(Finding("calcoli", "Importo nel testo non coerente con il ricalcolo", "Errore", loc,
                                   f"Testo: {fmt_num(v)} € · valore corretto e approvato in tabella: {fmt_num(corrections[v])} €",
                                   f"Sostituire {fmt_num(v)} con {fmt_num(corrections[v])}", ch))
            elif v not in known_vals and econ.lines:
                out.append(Finding("calcoli", "Importo nel testo senza riscontro nelle tabelle", "Da verificare", loc,
                                   f"Importo {fmt_num(v)} € citato nel testo ma assente dalle tabelle economiche",
                                   "Verificare la coerenza con le tabelle (non modificato automaticamente)"))
    return out
