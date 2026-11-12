"""Estrazione righe economiche dalle tabelle del DOCX e ricalcolo (quantità × prezzo, sconti, subtotali, totali, IVA)."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal

from datetime import timedelta

from .docx_utils import Change, DocView, Location, TInfo, norm
from .extract import DATE_RE, find_dates
from .money import D, find_numbers, fmt_num, parse_number, parse_percent, quantize

PERIODS = {"una_tantum": "Una tantum", "mensile": "Mensile", "trimestrale": "Trimestrale",
           "semestrale": "Semestrale", "annuale": "Annuale", "contratto": "Totale contratto"}
ANNUAL_FACTOR = {"mensile": Decimal(12), "trimestrale": Decimal(4), "semestrale": Decimal(2), "annuale": Decimal(1)}
KINDS = {"base": "Base", "opzione": "Opzione", "alternativa": "Alternativa"}


def detect_periodicity(text: str) -> str | None:
    t = norm(text)
    if re.search(r"una tantum|setup|set up|attivazione|installazione|implementazione|configurazione|one off|avviamento|onboarding", t):
        return "una_tantum"
    if re.search(r"mensil|\bmese\b|al mese|per mese", t):
        return "mensile"
    if re.search(r"trimestr", t):
        return "trimestrale"
    if re.search(r"semestr", t):
        return "semestrale"
    if re.search(r"annu|\banno\b|per anno|all anno", t):
        return "annuale"
    if re.search(r"triennal|biennal|pluriennal|intero contratto|totale contratto|durata contrattuale|per tutta la durata|\b\d+ (anni|mesi)\b", t):
        return "contratto"
    return None


def detect_kind(text: str) -> str:
    t = norm(text)
    if re.search(r"opzion|facoltativ|\bopt\b", t):
        return "opzione"
    if re.search(r"alternativ|variante", t):
        return "alternativa"
    return "base"


ROLE_PATTERNS = [
    ("sku", r"\b(sku|codice|cod|part number|p n|cod articolo)\b"),
    ("qty", r"\b(q ?ta|qta|quantita|qty|pezzi|n pezzi|unita|utenti|licenze|giornate|gg)\b"),
    ("unit", r"prezzo unit|unitario|\bp u\b|listino|costo unit|canone unit|prezzo singolo|prezzo cad"),
    ("disc", r"sconto|\bsc\b"),
    ("period", r"periodic|frequenza|tipologia|cadenza|periodo|\btipo\b"),
    ("gross", r"netto iva|iva incl|lordo|ivato|con iva|totale iva|\+ ?iva"),
    ("net", r"\bnetto\b|prezzo netto|importo netto|scontato"),
    ("amount", r"importo|totale|netto|valore|corrispettivo|canone|prezzo"),
    ("desc", r"descrizion|\bvoce\b|articolo|prodotto|servizio|dettaglio|oggetto|\bpos\b"),
]


def header_roles(cells: list[str]) -> dict[int, str]:
    roles: dict[int, str] = {}
    used = set()
    for i, txt in enumerate(cells):
        n = norm(txt)
        if not n:
            continue
        for role, pat in ROLE_PATTERNS:
            if re.search(pat, n):
                if role == "amount" and re.search(r"prezzo", n) and not re.search(r"importo|totale|netto|valore", n):
                    role = "unit"
                if role in used and role not in ("desc",):
                    continue
                roles[i] = role
                used.add(role)
                break
    if "net" in used and "amount" not in used:          # tabella semplice con la sola colonna «Netto»
        for i, r in roles.items():
            if r == "net":
                roles[i] = "amount"
    return roles


TOTAL_START = re.compile(r"^(sub ?totale|sotto ?totale|totale|imponibile|importo totale|iva|i v a|sconto|parziale|canone totale|netto a pagare|importo complessivo)\b")


def classify_total(label: str) -> str | None:
    n = norm(label)
    if not n:
        return None
    if n.startswith("sconto"):
        return "sconto"
    if re.search(r"iva inclusa|iva compresa|lordo|con iva|ivato|totale fattura|da pagare|importo ivato", n):
        return "lordo"
    if re.match(r"^(iva|i v a)\b", n):
        return "iva"
    if not TOTAL_START.match(n):
        return None
    if re.search(r"opzion", n):
        return "opzioni"
    if re.search(r"sub ?totale|sotto ?totale|parziale", n):
        return "subtotale"
    if re.search(r"una tantum", n):
        return "una_tantum"
    if re.search(r"ricorrent|canoni\b|solo canoni", n):
        return "ricorrenti"
    if re.search(r"contratto|triennal|biennal|pluriennal|complessiv|\b\d+ (anni|mesi)\b|totale offerta|periodo|quinquennal", n):
        return "contratto"
    if re.search(r"annu|\banno\b|per anno|primo anno|anno \d", n):
        return "anno"
    return "generico"


@dataclass
class EconLine:
    id: str
    table: int
    row: int
    section: str
    desc: str = ""
    sku: str = ""
    qty: Decimal | None = None
    unit: Decimal | None = None
    disc: Decimal | None = None
    reported: Decimal | None = None
    periodicity: str | None = None
    period_src: str = "ambigua"      # documento | confermata | ambigua
    kind: str = "base"
    col_amount: int | None = None
    col_qty: int | None = None
    col_unit: int | None = None
    col_disc: int | None = None
    cell_text: dict = field(default_factory=dict)
    net: Decimal | None = None          # colonna «Netto» (prezzo scontato, IVA esclusa)
    gross: Decimal | None = None        # colonna «Netto + IVA»
    col_net: int | None = None
    col_gross: int | None = None
    gross_calc: Decimal | None = None
    gross_status: str = ""
    gross_note: str = ""
    period: tuple | None = None
    calc: Decimal | None = None
    formula: str = ""
    diff: Decimal | None = None
    status: str = "Conforme"
    note: str = ""
    verifiable: bool = True

    @property
    def location(self) -> Location:
        return Location(self.section, table=self.table, row=self.row, col=self.col_amount, kind="cell")

    @property
    def amount(self) -> Decimal | None:
        return self.calc if self.calc is not None else self.reported

    @property
    def source(self) -> str:
        return f"Tabella {self.table + 1}, riga {self.row + 1}"


@dataclass
class TotalRow:
    id: str
    table: int
    row: int
    section: str
    label: str
    ctype: str
    col_amount: int
    reported: Decimal | None
    calc: Decimal | None = None
    formula: str = ""
    diff: Decimal | None = None
    status: str = "Conforme"
    note: str = ""
    cell_text: str = ""
    data_rate: Decimal | None = None
    role: str = "amount"          # amount (Listino/Importo) | net | gross
    header: str = ""

    @property
    def location(self) -> Location:
        return Location(self.section, table=self.table, row=self.row, col=self.col_amount, kind="cell")


@dataclass
class Econ:
    lines: list[EconLine] = field(default_factory=list)
    totals: list[TotalRow] = field(default_factory=list)
    months: int | None = None
    years: Decimal | None = None
    changes: list[Change] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    # aggregati calcolati (solo righe base)
    once: Decimal = Decimal(0)
    annual: Decimal | None = Decimal(0)
    contract: Decimal | None = Decimal(0)
    options: Decimal = Decimal(0)
    ambiguous: list[str] = field(default_factory=list)
    declared_vat: Decimal | None = None
    declared_gross: Decimal | None = None
    declared_net: Decimal | None = None
    vat_rate: Decimal | None = None
    vat_source: Location | None = None
    periods: list = field(default_factory=list)          # (inizio, fine, Location) da righe «dd mmm yyyy - dd mmm yyyy»
    period_totals: list = field(default_factory=list)    # (inizio, fine, importo)
    span_months: int | None = None                       # durata ricavata dai periodi in tabella
    period_rows: list = field(default_factory=list)      # dict(start, end, amount, net, gross)
    list_total: Decimal | None = None
    net_total: Decimal | None = None
    gross_total: Decimal | None = None
    inferred_vat: Decimal | None = None                  # aliquota dedotta da gross/net delle righe
    has_net: bool = False

    def base_lines(self):
        return [l for l in self.lines if l.kind == "base"]

    def per_year(self) -> list[dict] | None:
        if not self.years or self.annual is None or self.years != int(self.years):
            return None
        out = []
        for y in range(1, int(self.years) + 1):
            amt = self.annual + (self.once if y == 1 else Decimal(0))
            out.append({"anno": y, "importo": amt})
        return out


def _line_contract(l: EconLine, years: Decimal | None) -> Decimal | None:
    a = l.amount
    if a is None or l.periodicity is None:
        return None
    if l.periodicity == "una_tantum":
        return a
    if l.periodicity == "contratto":
        return a
    if years is None:
        return None
    return a * ANNUAL_FACTOR[l.periodicity] * years


def _line_annual(l: EconLine, years: Decimal | None) -> Decimal | None:
    a = l.amount
    if a is None or l.periodicity in (None, "una_tantum"):
        return None
    if l.periodicity == "contratto":
        return a / years if years else None
    return a * ANNUAL_FACTOR[l.periodicity]


def line_contract(l: EconLine, years):
    return _line_contract(l, years)


def line_annual(l: EconLine, years):
    return _line_annual(l, years)


def extract_and_calc(view: DocView, months: int | None, confirmations: dict | None = None,
                     decimals: int = 2, rounding: str = "half_up") -> Econ:
    conf = (confirmations or {}).get("lines", {})
    e = Econ(months=months, years=(Decimal(months) / 12 if months else None))
    q = lambda d: quantize(d, decimals, rounding)

    for t in view.tables:
        header_idx, roles = None, {}
        for r in range(min(3, len(t.rows))):
            rl = header_roles([c.text for c in t.rows[r]])
            if len(rl) >= 2 and ({"amount", "unit"} & set(rl.values())):
                header_idx, roles = r, rl
                break
        if header_idx is None:
            continue
        by_role = {v: t.rows[header_idx][k].grid for k, v in roles.items()}   # ruolo -> colonna di griglia
        hdr_texts = {v: t.rows[header_idx][k].text for k, v in roles.items()}
        hdr_hint = {role: detect_periodicity(txt) for role, txt in hdr_texts.items()}
        cur_period = None
        window: list[EconLine] = []
        subtotals: list[Decimal] = []
        discounts = Decimal(0)
        tlines: list[EconLine] = []
        last_net: Decimal | None = None
        last_iva: Decimal | None = None
        for r in range(header_idx + 1, len(t.rows)):
            cells = t.rows[r]
            texts = [c.text for c in cells]
            if not any(texts):
                continue

            def cell(role):
                g = by_role.get(role)
                return next((c for c in cells if c.grid == g), None) if g is not None else None

            joined = " ".join(x for x in texts if x)
            ds = find_dates(joined)
            if len(ds) == 2 and all(d for _, d in ds) and re.fullmatch(r"\s*(?:dal\s+)?(?:" + DATE_RE + r")\s*(?:-|–|—|al|to)\s*(?:" + DATE_RE + r")\s*", joined, re.I):
                cur_period = (ds[0][1], ds[1][1])
                e.periods.append((ds[0][1], ds[1][1], Location(t.section, table=t.idx, row=r, kind="cell")))
                continue

            desc_c = cell("desc")
            label = (desc_c.text if desc_c and desc_c.text else next((x for x in texts if x), ""))
            ctype = classify_total(label)
            amt_c = cell("amount")
            if ctype and amt_c is None:
                nums = [i for i, c in enumerate(cells) if find_numbers(c.text) and c.text != label]
                amt_c = cells[nums[-1]] if nums else None
            if ctype:
                if amt_c is None:
                    continue
                rep = parse_number(find_numbers(amt_c.text)[0][0]) if find_numbers(amt_c.text) else None
                tr = TotalRow(f"T{t.idx + 1}R{r + 1}", t.idx, r, t.section, label, ctype, amt_c.col, rep, cell_text=amt_c.text,
                              header=hdr_texts.get("amount", ""))
                win_snapshot = list(window)
                if ctype in ("generico", "subtotale", "contratto"):
                    for role in ("net", "gross"):
                        c2 = cell(role)
                        if c2 is None or not find_numbers(c2.text):
                            continue
                        srcs = [l for l in (tlines if ctype == "contratto" else win_snapshot) if l.kind == "base"]
                        vals = [getattr(l, role) for l in srcs]
                        tr2 = TotalRow(f"T{t.idx + 1}R{r + 1}{role[0]}", t.idx, r, t.section, label, ctype, c2.col,
                                       parse_number(find_numbers(c2.text)[0][0]), cell_text=c2.text, role=role, header=hdr_texts.get(role, ""))
                        if not srcs or any(v is None for v in vals):
                            tr2.status, tr2.note = "Non verificabile", "Righe senza valore in questa colonna."
                        else:
                            tr2.calc = q(sum(vals, Decimal(0)))
                            tr2.formula = f"somma colonna «{tr2.header}»: " + " + ".join(fmt_num(v) for v in vals[:6]) + (" + …" if len(vals) > 6 else "") + f" = {fmt_num(tr2.calc)}"
                            tr2.diff = q(tr2.reported - tr2.calc) if tr2.reported is not None else None
                            tr2.status = "Conforme" if tr2.diff == 0 else "Errore"
                            if role == "gross" and e.inferred_vat is not None:
                                pass
                        e.totals.append(tr2)
                _calc_total(e, tr, tlines, window, subtotals, discounts, last_net, last_iva, q, label, cells, texts)
                if ctype == "sconto":
                    discounts += tr.calc if tr.calc is not None else (tr.reported or Decimal(0))
                elif ctype == "subtotale":
                    subtotals.append(tr.calc if tr.calc is not None else (tr.reported or Decimal(0)))
                    window = []
                    discounts = Decimal(0)
                elif ctype == "generico":
                    window, subtotals, discounts = [], [], Decimal(0)
                if ctype in ("generico", "subtotale", "contratto", "anno", "ricorrenti", "una_tantum"):
                    last_net = tr.calc if tr.calc is not None else tr.reported
                if ctype == "iva":
                    last_iva = tr.calc if tr.calc is not None else tr.reported
                e.totals.append(tr)
                continue
            # riga voce
            qty_c, unit_c, disc_c, per_c, sku_c = cell("qty"), cell("unit"), cell("disc"), cell("period"), cell("sku")
            qty = (parse_number(find_numbers(qty_c.text)[0][0]) if qty_c and find_numbers(qty_c.text) else None)
            unit = (parse_number(find_numbers(unit_c.text)[0][0]) if unit_c and find_numbers(unit_c.text) else None)
            disc = None
            if disc_c and disc_c.text:
                disc = parse_percent(disc_c.text)
                if disc is None:
                    disc = parse_number(disc_c.text)
                if disc is not None:
                    disc = abs(disc)
            rep = (parse_number(find_numbers(amt_c.text)[0][0]) if amt_c and find_numbers(amt_c.text) else None)
            net_c, gross_c = cell("net"), cell("gross")
            net_v = parse_number(find_numbers(net_c.text)[0][0]) if net_c and find_numbers(net_c.text) else None
            gross_v = parse_number(find_numbers(gross_c.text)[0][0]) if gross_c and find_numbers(gross_c.text) else None
            if qty is None and unit is None and rep is None and net_v is None and gross_v is None:
                continue
            ln = EconLine(f"T{t.idx + 1}R{r + 1}", t.idx, r, t.section, desc=(desc_c.text if desc_c else label),
                          sku=(sku_c.text if sku_c else ""), qty=qty, unit=unit, disc=disc, reported=rep,
                          col_amount=amt_c.col if amt_c else None, col_qty=qty_c.col if qty_c else None,
                          col_unit=unit_c.col if unit_c else None, col_disc=disc_c.col if disc_c else None,
                          cell_text={"amount": amt_c.text if amt_c else "", "qty": qty_c.text if qty_c else "",
                                     "unit": unit_c.text if unit_c else "", "disc": disc_c.text if disc_c else "",
                                     "net": net_c.text if net_c else "", "gross": gross_c.text if gross_c else ""})
            ln.net, ln.gross = net_v, gross_v
            ln.col_net, ln.col_gross = (net_c.col if net_c else None), (gross_c.col if gross_c else None)
            ln.kind = conf.get(ln.id, {}).get("kind") or detect_kind(" ".join(texts))
            pcell = per_c.text if per_c else ""
            p = conf.get(ln.id, {}).get("periodicita")
            ln.period = cur_period
            if p:
                ln.periodicity, ln.period_src = p, "confermata"
            elif not detect_periodicity(pcell) and cur_period:
                ln.periodicity, ln.period_src = "contratto", "documento"     # importo riferito al periodo indicato nella riga di intestazione
            else:
                for src in (pcell, ln.desc, hdr_hint.get("unit") and hdr_texts.get("unit"), hdr_hint.get("amount") and hdr_texts.get("amount")):
                    pp = detect_periodicity(src) if src else None
                    if pp:
                        ln.periodicity, ln.period_src = pp, "documento"
                        break
            # ricalcolo riga
            if qty is not None and unit is not None:
                d = disc or Decimal(0)
                ln.calc = q(qty * unit * (1 - d / 100))
                ln.formula = (f"{fmt_num(qty, 0 if qty == qty.to_integral() else 2)} × {fmt_num(unit)}"
                              + (f" × (1 − {fmt_num(d, 0 if d == d.to_integral() else 2)}%)" if d else "")
                              + f" = {fmt_num(ln.calc)}")
                if rep is None:
                    ln.status, ln.note = "Da verificare", "Importo di riga non indicato."
                else:
                    ln.diff = q(rep - ln.calc)
                    if ln.diff != 0:
                        ln.status = "Errore"
                        gross = q(qty * unit)
                        if rep == gross and d:
                            ln.note = "L'importo riportato coincide con quantità × prezzo senza sconto: sconto non applicato?"
                    else:
                        ln.status = "Conforme"
            else:
                ln.verifiable = False
                ln.status = "Non verificabile"
                ln.note = "Quantità o prezzo unitario mancanti: ricalcolo non possibile."
                ln.formula = "importo riportato (non ricalcolabile)"
            if ln.amount is None and (ln.net is not None or ln.gross is not None):
                ln.reported = ln.reported if ln.reported is not None else (ln.net if ln.net is not None else ln.gross)
            if ln.periodicity is None and ln.kind == "base":
                e.ambiguous.append(f"{ln.source} «{ln.desc[:40]}»: periodicità non determinabile")
            tlines.append(ln)
            if ln.kind == "base":
                window.append(ln)
            e.lines.append(ln)
    both = [l for l in e.lines if l.net and l.gross]
    e.has_net = any(l.net is not None for l in e.lines)
    if both:
        from collections import Counter
        cnt = Counter(quantize((l.gross / l.net - 1) * 100, 2) for l in both)
        rate, n = cnt.most_common(1)[0]
        if n >= max(1, int(len(both) * 0.6 + 0.99)):
            e.inferred_vat = rate
            for l in both:
                l.gross_calc = q(l.net * (1 + rate / 100))
                if l.gross_calc != l.gross:
                    l.gross_status = "Errore"
                    l.gross_note = f"Netto {fmt_num(l.net)} + IVA {fmt_num(rate, 0 if rate == rate.to_integral_value() else 2)}% = {fmt_num(l.gross_calc)}, riportato {fmt_num(l.gross)}"
                else:
                    l.gross_status = "Conforme"
    if e.periods:
        e.span_months = _span_months(min(p[0] for p in e.periods), max(p[1] for p in e.periods))
        tot: dict = {}
        for l in e.lines:
            if l.kind == "base" and l.period:
                d = tot.setdefault(l.period, {"amount": Decimal(0), "net": Decimal(0), "gross": Decimal(0), "n": 0})
                d["amount"] += l.amount or Decimal(0)
                d["net"] += l.net or Decimal(0)
                d["gross"] += l.gross or Decimal(0)
                d["n"] += 1
        e.period_rows = [dict(start=a, end=b, **v) for (a, b), v in sorted(tot.items())]
        e.period_totals = [(r["start"], r["end"], r["amount"]) for r in e.period_rows]
    _aggregates(e)
    return e


def _span_months(start, end) -> int | None:
    end = end + timedelta(days=1)
    m = (end.year - start.year) * 12 + end.month - start.month + (1 if end.day > start.day + 3 else 0) - (1 if end.day < start.day - 3 else 0)
    return m if m > 0 else None


def _sum(ls, f):
    vals = [f(l) for l in ls]
    return None if any(v is None for v in vals) else sum(vals, Decimal(0))


def _aggregates(e: Econ) -> None:
    base = e.base_lines()
    e.once = sum((l.amount or Decimal(0) for l in base if l.periodicity == "una_tantum"), Decimal(0))
    amb = [l for l in base if l.periodicity is None]
    rec = [l for l in base if l.periodicity not in (None, "una_tantum", "contratto")]
    e.annual = None if (amb or any(_line_annual(l, e.years) is None for l in rec)) else \
        sum((_line_annual(l, e.years) for l in rec), Decimal(0))
    e.contract = None if (amb or any(_line_contract(l, e.years) is None for l in base)) else \
        sum((_line_contract(l, e.years) for l in base), Decimal(0))
    e.options = sum((l.amount or Decimal(0) for l in e.lines if l.kind == "opzione"), Decimal(0))
    if not rec and any(l.periodicity == "contratto" for l in base):
        e.annual = None
    if base:
        e.list_total = sum((l.amount for l in base if l.amount is not None), Decimal(0)) if all(l.amount is not None for l in base) else None
        e.net_total = sum((l.net for l in base), Decimal(0)) if all(l.net is not None for l in base) else None
        e.gross_total = sum((l.gross for l in base), Decimal(0)) if all(l.gross is not None for l in base) else None
    for t in e.totals:
        if t.ctype in ("contratto", "generico") and t.reported is not None:
            if t.ctype == "contratto" or e.declared_net is None:
                e.declared_net = t.reported
        if t.ctype == "iva":
            e.declared_vat = t.reported
        if t.ctype == "lordo":
            e.declared_gross = t.reported


def _calc_total(e: Econ, tr: TotalRow, tlines, window, subtotals, discounts, last_net, last_iva, q, label, cells, texts):
    base = [l for l in tlines if l.kind == "base"]
    opts = [l for l in tlines if l.kind == "opzione"]
    yrs = e.years
    ct = tr.ctype
    expected = None
    formula = ""
    alt: list[tuple[str, Decimal]] = []

    def cs(ls):
        return " + ".join(fmt_num(l.amount) for l in ls) if ls else "0"

    if ct == "sconto":
        pct = parse_percent(label) or next((parse_percent(x) for x in texts if parse_percent(x) is not None), None)
        basis = sum((l.amount or Decimal(0) for l in window), Decimal(0)) + sum(subtotals, Decimal(0))
        if pct is not None:
            expected = q(basis * pct / 100)
            formula = f"({fmt_num(basis)}) × {fmt_num(pct, 0 if pct == pct.to_integral() else 2)}% = {fmt_num(expected)}"
        else:
            tr.status, tr.note = "Non verificabile", "Percentuale di sconto non leggibile."
    elif ct == "subtotale":
        if any(l.amount is None for l in window):
            tr.status, tr.note = "Non verificabile", "Righe senza importo."
        else:
            expected = q(sum((l.amount for l in window), Decimal(0)))
            formula = f"{cs(window)} = {fmt_num(expected)}"
    elif ct == "generico":
        n = sum((l.amount or Decimal(0) for l in window), Decimal(0)) + sum(subtotals, Decimal(0)) - discounts
        expected = q(n)
        parts = ([cs(window)] if window else []) + ([f"subtotali {' + '.join(fmt_num(s) for s in subtotals)}"] if subtotals else [])
        formula = " + ".join(parts) + (f" − sconti {fmt_num(discounts)}" if discounts else "") + f" = {fmt_num(expected)}"
        if yrs and not subtotals and any(l.periodicity not in ("una_tantum", None, "contratto") for l in base):
            c = _sum(base, lambda l: _line_contract(l, yrs))
            if c is not None:
                alt.append((f"totale contratto ({int(e.months)} mesi)", q(c - discounts)))
    elif ct in ("contratto", "anno", "ricorrenti", "una_tantum"):
        if ct == "una_tantum":
            expected = q(sum((l.amount or Decimal(0) for l in base if l.periodicity == "una_tantum"), Decimal(0)))
            formula = f"somma voci una tantum = {fmt_num(expected)}"
        elif any(l.periodicity is None for l in base):
            tr.status, tr.note = "Non verificabile", "Periodicità di alcune righe da confermare: " + "; ".join(l.source for l in base if l.periodicity is None)
        elif ct in ("anno", "ricorrenti"):
            rec = [l for l in base if l.periodicity not in ("una_tantum",)]
            vals = [_line_annual(l, yrs) for l in rec]
            if any(v is None for v in vals):
                tr.status, tr.note = "Non verificabile", "Durata contratto non indicata o non confermata."
            else:
                s = sum(vals, Decimal(0))
                incl_once = bool(re.search(r"primo anno|anno 1\b|1° anno", norm(label))) and ct == "anno"
                once = sum((l.amount or Decimal(0) for l in base if l.periodicity == "una_tantum"), Decimal(0)) if incl_once else Decimal(0)
                expected = q(s + once - discounts)
                formula = "canoni annualizzati " + " + ".join(fmt_num(v) for v in vals) + (f" + una tantum {fmt_num(once)}" if once else "") + (f" − sconti {fmt_num(discounts)}" if discounts else "") + f" = {fmt_num(expected)}"
                if not incl_once and any(l.periodicity == "una_tantum" for l in base):
                    alt.append(("canoni + una tantum (primo anno)", q(s + sum((l.amount or Decimal(0) for l in base if l.periodicity == "una_tantum"), Decimal(0)) - discounts)))
        else:
            vals = [_line_contract(l, yrs) for l in base]
            if any(v is None for v in vals):
                tr.status, tr.note = "Non verificabile", "Durata contratto non indicata o non confermata: totale contratto non ricalcolabile."
            else:
                expected = q(sum(vals, Decimal(0)) - discounts)
                formula = ("una tantum + canoni × durata: " + " + ".join(fmt_num(v) for v in vals)) + (f" − sconti {fmt_num(discounts)}" if discounts else "") + f" = {fmt_num(expected)}"
    elif ct == "opzioni":
        expected = q(sum((l.amount or Decimal(0) for l in opts), Decimal(0)))
        formula = f"somma opzioni {cs(opts)} = {fmt_num(expected)}"
    elif ct == "iva":
        rate = parse_percent(label) or next((parse_percent(x) for x in texts if parse_percent(x) is not None), None)
        if rate is None:
            m = re.search(r"(\d{1,2}(?:[.,]\d+)?)\s*%?", norm(label).replace("iva", ""))
            rate = parse_number(m.group(1)) if m else None
        if rate is not None:
            tr.data_rate = rate
            e.vat_rate = rate
            e.vat_source = tr.location
        net = last_net
        if rate is None or net is None:
            tr.status, tr.note = "Non verificabile", "Aliquota IVA o imponibile non leggibili."
        else:
            expected = q(net * rate / 100)
            formula = f"{fmt_num(net)} × {fmt_num(rate, 0 if rate == rate.to_integral() else 2)}% = {fmt_num(expected)}"
    elif ct == "lordo":
        if last_net is None or last_iva is None:
            tr.status, tr.note = "Non applicabile", "IVA non presente nella tabella: totale lordo non verificabile."
        else:
            expected = q(last_net + last_iva)
            formula = f"{fmt_num(last_net)} + {fmt_num(last_iva)} = {fmt_num(expected)}"
    tr.formula = formula
    if expected is not None:
        tr.calc = expected
        if tr.reported is None:
            tr.status, tr.note = "Da verificare", "Importo non leggibile."
        else:
            tr.diff = q(tr.reported - expected)
            if tr.diff == 0:
                tr.status = "Conforme"
            else:
                match = next((n for n, v in alt if v == tr.reported), None)
                if match:
                    tr.status = "Conforme"
                    tr.note = f"Il valore riportato coincide con il {match}."
                    tr.calc, tr.diff = tr.reported, Decimal(0)
                else:
                    tr.status = "Errore"
                    sum_rep = sum((l.reported or Decimal(0) for l in window), Decimal(0))
                    if ct == "generico" and q(sum_rep) == tr.reported:
                        tr.note = "Coerente con gli importi di riga riportati, ma questi non risultano corretti al ricalcolo."
    elif tr.status == "Conforme" and not formula:
        tr.status = "Non verificabile"


def change_for_cell(table: int, row: int, col: int, cell_text: str, new_value: Decimal, label: str, category: str,
                    decimals: int | None = None, origin: str = "auto", reason: str = "") -> Change | None:
    from .money import replace_number_token
    r = replace_number_token(cell_text, new_value, decimals)
    if not r:
        return None
    tok, new = r
    return Change("cell", tok, new, label, category, table=table, row=row, col=col, origin=origin,
                  ctx_before=cell_text, reason=reason)
