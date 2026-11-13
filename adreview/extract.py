"""Estrazione dei campi richiesti dal profilo (cliente, numero, data, oggetto, perimetro, durata, rinnovo, validità, pagamento)."""
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass, field
from datetime import date, timedelta

from .config import Profile
from .docx_utils import DocView, Location, norm

MONTHS = {m: i + 1 for i, m in enumerate(
    ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto",
     "settembre", "ottobre", "novembre", "dicembre"])}
WORDS = {"un": 1, "uno": 1, "una": 1, "due": 2, "tre": 3, "quattro": 4, "cinque": 5, "sei": 6, "dodici": 12,
         "diciotto": 18, "ventiquattro": 24, "trentasei": 36, "quarantotto": 48, "sessanta": 60}
WORD_RE = "|".join(WORDS)
EN_MONTHS = {m: i + 1 for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}
DATE_RE = (r"\d{1,2}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{2,4}|\d{1,2}\s*°?\s*(?:" + "|".join(MONTHS) +
           r"|january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sept?|oct|nov|dec)\.?,?\s+\d{4}")
FIELD_LABELS = {
    "cliente": "Cliente", "numero_offerta": "Numero offerta", "data_offerta": "Data offerta",
    "oggetto": "Oggetto", "perimetro": "Perimetro", "prezzi": "Prezzi", "durata": "Durata",
    "rinnovo": "Rinnovo", "validita": "Validità offerta", "pagamento": "Condizioni di pagamento",
}
NOT_INDICATED = "Non indicato"


def parse_date(text: str | None) -> date | None:
    if not text:
        return None
    t = text.strip().lower()
    m = re.fullmatch(r"(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{2,4})", t)
    try:
        if m:
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if y < 100:
                y += 2000
            return date(y, mo, d)
        m = re.fullmatch(r"(\d{1,2})\s*°?\s*([a-zà-ù]+)\.?,?\s+(\d{4})", t)
        if m:
            mon = MONTHS.get(m.group(2)) or EN_MONTHS.get(m.group(2)[:3])
            if mon:
                return date(int(m.group(3)), mon, int(m.group(1)))
    except ValueError:
        return None
    return None


def fmt_date(d: date | None) -> str:
    return d.strftime("%d/%m/%Y") if d else "—"


def add_months(d: date, n: int) -> date:
    y, m = divmod(d.month - 1 + n, 12)
    y += d.year
    m += 1
    return date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def find_dates(text: str) -> list[tuple[str, date | None]]:
    return [(m.group(0), parse_date(m.group(0))) for m in re.finditer(DATE_RE, text or "", re.I)]


@dataclass
class Occ:
    value: str
    loc: Location
    source: str = ""
    data: dict = field(default_factory=dict)


@dataclass
class FieldResult:
    key: str
    label: str
    required: bool = True
    status: str = "Da verificare"
    value: str = NOT_INDICATED
    occurrences: list[Occ] = field(default_factory=list)
    note: str = ""
    data: dict = field(default_factory=dict)

    @property
    def present(self) -> bool:
        return bool(self.occurrences)

    @property
    def loc(self) -> Location | None:
        return self.occurrences[0].loc if self.occurrences else None


def _num(tok: str) -> int | None:
    tok = tok.lower()
    return int(tok) if tok.isdigit() else WORDS.get(tok)


def _sentences(text: str) -> list[str]:
    return [s.strip() for s in re.split(r"(?<=[.;!?])\s+|\n", text) if s.strip()]


def _clean(v: str) -> str:
    return v.strip().strip("-–:;,").strip()


def _distinct(occs: list[Occ], keyf=lambda o: norm(o.value)) -> list[str]:
    seen = []
    for o in occs:
        k = keyf(o)
        if k not in seen:
            seen.append(k)
    return seen


def extract_fields(view: DocView, profile: Profile, has_prices: bool = False,
                   months_override: int | None = None) -> dict[str, FieldResult]:
    items = list(view.scan()) + list(view.scan_extra())
    F = {k: FieldResult(k, lab, required=k in profile.required_fields) for k, lab in FIELD_LABELS.items()}

    def label_value(pattern: str, key: str, nextline: bool = True, transform=_clean):
        rx = re.compile(pattern, re.I)
        for i, (loc, text) in enumerate(items):
            m = rx.match(text)
            if not m:
                continue
            v = transform(m.group("v"))
            if not v and nextline and i + 1 < len(items):
                v = transform(items[i + 1][1])
                loc = items[i + 1][0]
            if v:
                F[key].occurrences.append(Occ(v, loc, text))

    # --- cliente, oggetto
    label_value(r"^\s*(?:spett\.?le|spettabile|cliente|destinatario|committente)\b\.?\s*[:\-–]?\s*(?P<v>.*)$", "cliente")
    label_value(r"^\s*oggetto\s*[:\-–]\s*(?P<v>.*)$", "oggetto")

    if not F["oggetto"].occurrences:
        label_value(r"^\s*(?:offerta|preventivo|proposta)\s+(?:commerciale\s+)?per\s+(?P<v>.+)$", "oggetto", nextline=False)

    # --- numero offerta
    tok = r"(?P<v>[A-Za-z0-9][A-Za-z0-9/_\-.]*)"
    n_pats = [
        r"\b(?:offerta|preventivo|proposta(?:\s+commerciale)?)\s*(?:n\.?|n°|nº|nr\.?|numero|num\.?|prot\.?|rif\.?)\s*[:\-]?\s*" + tok,
        r"\b(?:numero|n\.|n°|nr\.?|codice|riferimento|rif\.?)\s*(?:dell['’]\s*|di\s+)?(?:offerta|preventivo|proposta)\s*[:\-]?\s*" + tok,
        r"^\s*(?:offerta|preventivo|rif\.?\s*offerta)\s*:\s*" + tok,
        r"^\s*rif(?:erimento)?\.?\s*[:\-]?\s*" + tok,
    ]
    for loc, text in items:
        for p in n_pats:
            for m in re.finditer(p, text, re.I):
                v = m.group("v").rstrip(".,;:")
                if re.search(r"\d", v):
                    F["numero_offerta"].occurrences.append(Occ(v, loc, text))
            if F["numero_offerta"].occurrences and F["numero_offerta"].occurrences[-1].source == text:
                break

    # --- data offerta
    unreadable_date = None
    for loc, text in items:
        m = re.search(r"\bdata(?:\s+(?:dell['’]\s*)?offerta|\s+di\s+emissione|\s+emissione)?\s*[:\-]\s*(?P<v>\S.*)$", text, re.I)
        if m:
            ds = find_dates(m.group("v"))
            if ds:
                F["data_offerta"].occurrences.append(Occ(ds[0][0], loc, text, {"date": ds[0][1]}))
            else:
                unreadable_date = unreadable_date or Occ(m.group("v")[:30], loc, text)
            continue
        m = re.match(r"^\s*[A-ZÀ-Ý][\w' ]{2,25}\s*[,–—\-]\s*(?P<v>" + DATE_RE + r")(?:\s+(?:" + DATE_RE + r"))*\s*$", text, re.I)
        if m:
            F["data_offerta"].occurrences.append(Occ(m.group("v"), loc, text, {"date": parse_date(m.group("v"))}))
            continue
        m = re.search(r"\b(?:offerta|preventivo)\s*(?:n\.?|n°)\s*\S+\s+del\s+(?P<v>" + DATE_RE + ")", text, re.I)
        if m:
            F["data_offerta"].occurrences.append(Occ(m.group("v"), loc, text, {"date": parse_date(m.group("v"))}))
    if not F["data_offerta"].occurrences and unreadable_date:
        F["data_offerta"].occurrences.append(unreadable_date)
        F["data_offerta"].data["unreadable"] = True

    # --- perimetro
    p_alias = [norm(a) for a in ["perimetro", "ambito", "ambito della fornitura", "descrizione della fornitura",
                                  "oggetto della fornitura", "scope", "servizi inclusi", "cosa include"]]
    for sec in profile.required_sections:
        if "perimetro" in norm(sec.get("title", "")):
            p_alias += [norm(sec["title"])] + [norm(a) for a in sec.get("aliases", [])]
    for h in view.headings:
        if any(norm(h.text) == a or a in norm(h.text) for a in p_alias):
            body, ntab, started = [], 0, False
            for kind, it in view.items:
                if kind == "p" and it is h:
                    started = True
                    continue
                if not started:
                    continue
                if kind == "p":
                    if it.heading_level:
                        break
                    if it.text.strip():
                        body.append(it.text.strip())
                else:
                    ntab += 1
            val = (" ".join(body)[:160] + ("…" if len(" ".join(body)) > 160 else "")) or ("(descritto in tabella)" if ntab else "")
            if val:
                F["perimetro"].occurrences.append(Occ(val, Location(h.text, paragraph=h.idx, kind="par"), h.text))
            break
    if not F["perimetro"].occurrences:
        label_value(r"^\s*(?:perimetro|ambito)(?:\s+della\s+fornitura)?\s*[:\-–]\s*(?P<v>.*)$", "perimetro", nextline=False)

    # --- durata
    dur = F["durata"]
    for loc, text in items:
        found = False
        for p in [r"durata[^.\n:]{0,40}?[:\-–]?\s*(?:di\s*)?(?P<n>\d{1,3}|" + WORD_RE + r")\s*(?P<u>mesi|mese|anni|anno)\b",
                  r"validit[àa]\s+(?:del\s+)?contratto\s*[:\-–]?\s*(?P<n>\d{1,3}|" + WORD_RE + r")\s*(?P<u>mesi|anni|anno)\b",
                  r"\b(?:contratto|fornitura|servizio|abbonamento|canone)\s+(?:di\s+durata\s+|della\s+durata\s+di\s+|di\s+)(?P<n>\d{1,3}|" + WORD_RE + r")\s*(?P<u>mesi|anni)\b"]:
            for m in re.finditer(p, text, re.I):
                n = _num(m.group("n"))
                if n:
                    months = n * 12 if m.group("u").lower().startswith("ann") else n
                    dur.occurrences.append(Occ(f"{months} mesi", loc, text, {"months": months}))
                    found = True
            if found:
                break
        if not found:
            m = re.search(r"\b(?P<w>triennale|biennale|quadriennale|quinquennale)\b", text, re.I)
            if m and re.search(r"durata|contratt|rapporto|fornitura|servizio|canone", text, re.I):
                months = {"biennale": 24, "triennale": 36, "quadriennale": 48, "quinquennale": 60}[m.group("w").lower()]
                dur.occurrences.append(Occ(f"{months} mesi", loc, text, {"months": months}))
        m = re.search(r"decorrenz\w*[^.\n]{0,40}?(?P<d>" + DATE_RE + ")", text, re.I) or \
            re.search(r"\bdal\s+(?P<d>" + DATE_RE + r")\s+al\s+(?P<e>" + DATE_RE + ")", text, re.I)
        if m:
            dur.data["decorrenza"] = (parse_date(m.group("d")), loc)
            if "e" in m.groupdict() and m.group("e"):
                dur.data["scadenza"] = (parse_date(m.group("e")), loc)
        m = re.search(r"scadenza[^.\n]{0,30}?(?P<d>" + DATE_RE + ")", text, re.I)
        if m and "scadenza" not in dur.data:
            dur.data["scadenza"] = (parse_date(m.group("d")), loc)
    if months_override:
        dur.data["months"] = months_override
        dur.data["confirmed"] = True
    elif dur.occurrences:
        dur.data["months"] = dur.occurrences[0].data["months"]

    # --- rinnovo
    ren = F["rinnovo"]
    types, renew_months, disdetta = [], None, None
    for loc, text in items:
        for s in _sentences(text):
            sl = s.lower()
            if re.search(r"rinnov|tacit", sl):
                if re.search(r"non\s+(?:è\s+previsto|si\s+rinnova|prevede|sono\s+previsti)|nessun\s+rinnovo|senza\s+(?:tacito\s+)?rinnovo|nessuna\s+rinnov", sl):
                    t = "assente"
                elif re.search(r"tacit|automatic", sl):
                    t = "automatico"
                elif re.search(r"da\s+(?:concordare|definire|valutare|stabilire)|\bTBD\b", sl):
                    t = "da_definire"
                elif re.search(r"espress|previo\s+accordo|su\s+richiesta|per\s+iscritto|previa\s+accettazione", sl):
                    t = "espresso"
                else:
                    t = None
                m = re.search(r"(?:per|di|ulterior\w*)\s+(?:periodi\s+di\s+)?(?P<n>\d{1,3}|" + WORD_RE + r")\s*(?P<u>mesi|anni|anno|mese)\b", sl)
                if m and _num(m.group("n")):
                    renew_months = _num(m.group("n")) * (12 if m.group("u").startswith("ann") else 1)
                ren.occurrences.append(Occ(s[:140], loc, text, {"type": t, "months": renew_months}))
                if t:
                    types.append(t)
            if re.search(r"disdett|preavviso", sl):
                m = re.search(r"(?P<n>\d{1,3})\s*(?P<u>giorni|gg|mesi)", sl)
                if m:
                    disdetta = f"{m.group('n')} {'giorni' if m.group('u') in ('giorni', 'gg') else 'mesi'}"
                    if not ren.occurrences or ren.occurrences[-1].source != text:
                        ren.occurrences.append(Occ(s[:140], loc, text, {"type": None}))
    ren.data.update(types=types, renew_months=renew_months, disdetta=disdetta)

    # --- validità
    val = F["validita"]
    for loc, text in items:
        m = re.search(r"(?:validit[àa]|valid[ao])[^.\n]{0,60}?(?P<n>\d{1,3})\s*(?P<u>giorni|gg|giorno|mesi)\b", text, re.I)
        if m:
            clear = bool(re.search(r"dalla\s+data|dalla\s+presente|dall['’]\s*emissione|dalla\s+emissione|data\s+(?:di\s+)?(?:emissione|invio|dell['’]\s*offerta)", text, re.I))
            val.occurrences.append(Occ(f"{m.group('n')} {'giorni' if m.group('u') in ('giorni', 'gg', 'giorno') else 'mesi'}", loc, text,
                                       {"days": int(m.group("n")) if m.group("u") in ("giorni", "gg", "giorno") else None,
                                        "months": int(m.group("n")) if m.group("u") == "mesi" else None, "clear": clear}))
            continue
        m = re.search(r"valid\w*[^.\n]{0,40}?(?:fino\s+al|entro\s+il|sino\s+al)\s*:?\s*(?P<d>" + DATE_RE + ")", text, re.I) or \
            re.search(r"validit[àa][^.\n]{0,50}?[:]\s*(?P<d>" + DATE_RE + ")", text, re.I)
        if m:
            d = parse_date(m.group("d"))
            val.occurrences.append(Occ(m.group("d"), loc, text, {"date": d, "clear": True, "raw": m.group("d")}))

    # --- pagamento
    pay = F["pagamento"]
    for loc, text in items:
        for s in _sentences(text):
            if re.search(r"pagament|bonifico|d\.?\s?f\.?\s?f\.?\s?m|giorni\s+data\s+fattura|\bgg\.?\s*d\.?f|\d\s*gg", s, re.I) and not re.match(r"^\s*(?:condizioni\s+di\s+pagamento|riepilogo\b.*)\s*$", s, re.I):
                m = re.search(r"(\d{1,3})\s*(?:gg|giorni)\s*(d\.?\s?f\.?\s?f\.?\s?m\.?|d\.?\s?f\.?|data\s+fattura|fine\s+mese)?", s, re.I)
                pay.occurrences.append(Occ(s[:160], loc, text, {"term": norm(m.group(0)) if m else None}))

    # --- prezzi
    F["prezzi"].data["has"] = has_prices
    _finalize(F, has_prices)
    return F


def _finalize(F: dict[str, FieldResult], has_prices: bool) -> None:
    for k, fr in F.items():
        if k == "prezzi":
            if has_prices:
                fr.status, fr.value = "Conforme", "Righe economiche estratte"
            else:
                fr.status, fr.value = (("Da verificare" if fr.required else "Non applicabile"), NOT_INDICATED)
                fr.note = "Nessuna tabella economica riconoscibile." if fr.required else ""
            continue
        if not fr.occurrences:
            fr.value = NOT_INDICATED
            fr.status = "Da verificare" if fr.required else "Non applicabile"
            fr.note = "Dato richiesto dal profilo non presente nel documento." if fr.required else "Non richiesto dal profilo."
            continue
        if k == "data_offerta":
            if fr.data.get("unreadable"):
                fr.status, fr.value = "Non verificabile", fr.occurrences[0].value
                fr.note = "Data presente ma non leggibile in un formato riconosciuto."
                continue
            ds = [o.data.get("date") for o in fr.occurrences]
            if any(d is None for d in ds):
                fr.status, fr.value = "Non verificabile", fr.occurrences[0].value
                fr.note = "Data non valida o non leggibile."
                continue
            distinct = sorted(set(ds))
            fr.value = fmt_date(distinct[0])
            if len(distinct) > 1:
                fr.status = "Errore"
                fr.note = "Date offerta discordanti: " + ", ".join(fmt_date(d) for d in distinct)
            else:
                fr.status = "Conforme"
            continue
        if k == "durata":
            ms = sorted({o.data["months"] for o in fr.occurrences})
            fr.value = f"{fr.data.get('months', ms[0])} mesi"
            if len(ms) > 1 and not fr.data.get("confirmed"):
                fr.status = "Errore"
                fr.note = "Durate discordanti: " + ", ".join(f"{m} mesi" for m in ms)
            else:
                fr.status = "Conforme"
            continue
        if k == "rinnovo":
            ts = set(fr.data.get("types", []))
            parts = []
            if len(ts) > 1:
                fr.status = "Errore"
                fr.note = "Indicazioni di rinnovo contrastanti: " + ", ".join(sorted(ts))
            elif ts == {"da_definire"}:
                fr.status = "Da verificare"
                fr.note = ("Rinnovo «da concordare»: non sono definiti modalità (tacito/espresso), durata del rinnovo né termine di disdetta"
                           + (f" (il contratto dura {fr.data.get('months') or ''} mesi)" if False else "") + ".")
            elif not ts:
                fr.status = "Da verificare"
                fr.note = "Si parla di rinnovo/disdetta ma la modalità (automatico/espresso/assente) non è chiara."
            else:
                fr.status = "Conforme"
            if ts:
                parts.append("Rinnovo " + (", ".join(t.replace("da_definire", "da concordare (non definito)") for t in sorted(ts))))
            if fr.data.get("renew_months"):
                parts.append(f"durata rinnovo {fr.data['renew_months']} mesi")
            if fr.data.get("disdetta"):
                parts.append(f"disdetta {fr.data['disdetta']}")
            fr.value = "; ".join(parts) if parts else fr.occurrences[0].value
            continue
        if k == "validita":
            invalid = [o for o in fr.occurrences if "raw" in o.data and o.data.get("date") is None]
            good = [o for o in fr.occurrences if not ("raw" in o.data and o.data.get("date") is None)]
            fr.value = (good or fr.occurrences)[0].value
            fr.status = "Conforme"
            vals = {(o.data.get("days"), o.data.get("months"), o.data.get("date")) for o in good}
            notes = []
            if invalid:
                fr.status = "Errore"
                notes.append("Data inesistente: " + ", ".join(f"{o.value} ({o.loc.label})" for o in invalid))
            if len(vals) > 1:
                fr.status = "Errore"
                notes.append("Validità discordante tra sezioni: " + " / ".join(o.value for o in good))
            elif invalid and good:
                notes.append(f"Altra indicazione di validità nel documento: {good[0].value}")
            fr.note = ". ".join(notes)
            if invalid:
                fr.occurrences = good + invalid
            continue
        if k == "pagamento":
            terms = {o.data.get("term") for o in fr.occurrences if o.data.get("term")}
            fr.occurrences = [o for o in fr.occurrences if o.data.get("term")] + [o for o in fr.occurrences if not o.data.get("term")]
            fr.value = fr.occurrences[0].value
            if len(terms) > 1:
                fr.status, fr.note = "Errore", "Termini di pagamento discordanti: " + " / ".join(sorted(terms))
            else:
                fr.status = "Conforme"
            continue
        # cliente, numero_offerta, oggetto, perimetro
        distinct = _distinct(fr.occurrences, keyf=lambda o: re.sub(r"[^a-z0-9]", "", norm(o.value)))
        fr.value = fr.occurrences[0].value
        if k in ("cliente", "numero_offerta", "oggetto") and len(distinct) > 1:
            fr.status = "Errore" if k != "oggetto" else "Da verificare"
            fr.note = "Valori discordanti: " + " / ".join(dict.fromkeys(o.value for o in fr.occurrences))
        else:
            fr.status = "Conforme"
        if re.fullmatch(r"[\[<{(]?\s*(?:x+|\.{3,}|_+|\?+|da\s+(?:inserire|compilare))\s*[\]>})]?", fr.value, re.I) or \
                re.search(r"\[[^\]]*\]|\{\{.*?\}\}", fr.value):
            fr.status, fr.note = "Errore", "Valore segnaposto non compilato."
