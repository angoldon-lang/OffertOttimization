"""Abbinamento ricavi/costi, normalizzazione dei costi e calcolo del margine (aritmetica decimale)."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from difflib import SequenceMatcher

from .docx_utils import norm
from .econ import Econ, EconLine, line_annual, line_contract
from .money import quantize
from .suppliers import CostLine

ZERO = Decimal(0)


@dataclass
class Alloc:
    mode: str = "none"                  # none | single | equal | proportional | manual
    targets: list[str] = field(default_factory=list)
    pcts: dict[str, Decimal] = field(default_factory=dict)
    suggested: bool = False             # proposto automaticamente e non ancora confermato
    score: float = 0.0

    @property
    def criterion(self) -> str:
        return {"none": "non abbinato", "single": "100% a una voce", "equal": "ripartizione in parti uguali",
                "proportional": "ripartizione proporzionale al ricavo di contratto", "manual": "percentuali manuali"}[self.mode]


@dataclass
class SaleRow:
    id: str
    desc: str
    sku: str
    kind: str
    qty: Decimal | None
    periodicity: str | None
    revenue: Decimal | None          # contratto, netto IVA
    annual: Decimal | None
    once: Decimal | None
    source: str


def sales_from_econ(econ: Econ) -> list[SaleRow]:
    out = []
    for l in econ.lines:
        if l.kind == "alternativa":
            continue
        a = l.amount
        rev = line_contract(l, econ.years)
        ann = line_annual(l, econ.years)
        once = a if l.periodicity == "una_tantum" else None
        out.append(SaleRow(l.id, l.desc, l.sku, l.kind, l.qty, l.periodicity, rev, ann, once, l.source))
    return out


def _tok(s: str) -> set[str]:
    return {t for t in norm(s).split() if len(t) > 2}


def suggest_allocs(costs: list[CostLine], sales: list[SaleRow]) -> dict[str, Alloc]:
    out: dict[str, Alloc] = {}
    base = [s for s in sales if s.kind == "base"] or sales
    for c in costs:
        scored = []
        for s in base:
            sc = 0.0
            if c.sku and s.sku and norm(c.sku) == norm(s.sku):
                sc = 1.0
            else:
                a, b = _tok(c.desc), _tok(s.desc)
                jac = len(a & b) / len(a | b) if a | b else 0
                sc = 0.7 * max(SequenceMatcher(None, norm(c.desc), norm(s.desc)).ratio() * 0.8, jac)
                if c.qty is not None and s.qty is not None and c.qty == s.qty:
                    sc += 0.15
                if c.periodicity and s.periodicity and c.periodicity == s.periodicity:
                    sc += 0.1
            scored.append((sc, s))
        scored.sort(key=lambda x: -x[0])
        if scored and scored[0][0] >= 0.6 and (len(scored) == 1 or scored[0][0] - scored[1][0] >= 0.1):
            out[c.id] = Alloc("single", [scored[0][1].id], suggested=True, score=round(scored[0][0], 2))
        else:
            out[c.id] = Alloc("none", score=round(scored[0][0], 2) if scored else 0.0)
    return out


def resolve_shares(a: Alloc, sales: dict[str, SaleRow]) -> tuple[dict[str, Decimal] | None, str | None]:
    tg = [t for t in a.targets if t in sales]
    if a.mode == "none" or not tg:
        return None, None
    if a.mode == "single":
        return {tg[0]: Decimal(1)}, None
    if a.mode == "equal":
        n = len(tg)
        sh = {t: Decimal(1) / n for t in tg[:-1]}
        sh[tg[-1]] = Decimal(1) - sum(sh.values(), ZERO)
        return sh, None
    if a.mode == "proportional":
        w = {t: (sales[t].revenue or ZERO) for t in tg}
        tot = sum(w.values(), ZERO)
        if tot == 0:
            return None, "ripartizione proporzionale impossibile: ricavi nulli o non calcolabili"
        sh = {t: w[t] / tot for t in tg[:-1]}
        sh[tg[-1]] = Decimal(1) - sum(sh.values(), ZERO)
        return sh, None
    if a.mode == "manual":
        p = {t: a.pcts.get(t, ZERO) for t in tg}
        if sum(p.values(), ZERO) != 100:
            return None, f"le percentuali di ripartizione sommano {sum(p.values(), ZERO)}% invece di 100%"
        return {t: v / 100 for t, v in p.items()}, None
    return None, "modalità di ripartizione sconosciuta"


@dataclass
class CostCalc:
    cost: CostLine
    contract: Decimal | None = None      # EUR netto IVA, intero contratto
    annual: Decimal | None = None
    once: Decimal | None = None
    mult: Decimal | None = None          # moltiplicatore al contratto
    fx: Decimal | None = None
    issues: list[str] = field(default_factory=list)     # bloccanti
    warnings: list[str] = field(default_factory=list)


def normalize_cost(c: CostLine, months: int | None, fx_rates: dict[str, Decimal]) -> CostCalc:
    cc = CostCalc(c)
    total = c.total_effective
    if total is None:
        cc.issues.append("importo di costo mancante o non leggibile")
        return cc
    cur = (c.currency or "EUR").upper()
    if cur != "EUR":
        rate = fx_rates.get(cur)
        if not rate:
            cc.issues.append(f"valuta {cur}: indicare il cambio verso EUR")
            return cc
        cc.fx = rate
        total = total * rate
    else:
        cc.fx = Decimal(1)
    if c.vat_included:
        if c.vat_pct is None:
            cc.issues.append("importo IVA inclusa: indicare l'aliquota per ricavare il netto")
            return cc
        total = total / (1 + c.vat_pct / 100)
    years = Decimal(months) / 12 if months else None
    p = c.periodicity
    if p is None:
        cc.issues.append("periodicità del costo da confermare")
    elif p == "una_tantum":
        cc.contract, cc.once, cc.mult = total, total, Decimal(1)
    elif p in ("annuale", "mensile", "trimestrale", "semestrale"):
        f = {"annuale": 1, "mensile": 12, "trimestrale": 4, "semestrale": 2}[p]
        cc.annual = total * f
        if years is None:
            cc.issues.append("durata dell'offerta non indicata: costo ricorrente non riportabile al contratto")
        else:
            cc.contract, cc.mult = total * f * years, Decimal(f) * years
    elif p == "periodo_quotazione":
        m = c.period_months
        if not m:
            cc.issues.append("durata coperta dalla quotazione da indicare")
        elif months is None:
            cc.issues.append("durata dell'offerta non indicata")
        elif m != months and not c.base_confirmed:
            cc.issues.append(f"la quotazione copre {m} mesi, l'offerta {months}: confermare la base di confronto (riproporzione lineare)")
        else:
            cc.mult = Decimal(months) / m
            cc.contract = total * cc.mult
            cc.annual = total * 12 / m
            if m != months:
                cc.warnings.append(f"costo riproporzionato da {m} a {months} mesi (base confermata dall'utente)")
    return cc


@dataclass
class MarginLine:
    sale: SaleRow
    revenue: Decimal | None = None
    supplier_cost: Decimal = ZERO
    internal_cost: Decimal = ZERO
    status: str = "completa"           # completa | parziale | mancante
    issues: list[str] = field(default_factory=list)       # bloccanti: la voce non entra nei totali
    warnings: list[str] = field(default_factory=list)     # non bloccanti (es. abbinamento suggerito)
    cost_ids: list[str] = field(default_factory=list)
    annual_rev: Decimal | None = None
    annual_cost: Decimal | None = None
    once_rev: Decimal | None = None
    once_cost: Decimal = ZERO
    annual_ok: bool = True

    def margin_supplier(self):
        return None if self.revenue is None else self.revenue - self.supplier_cost

    def margin_total(self):
        return None if self.revenue is None else self.revenue - self.supplier_cost - self.internal_cost

    @staticmethod
    def pct(m, rev):
        return None if m is None or not rev else m / rev * 100


@dataclass
class MarginResult:
    lines: list[MarginLine] = field(default_factory=list)
    options: list[MarginLine] = field(default_factory=list)
    costs: list[CostCalc] = field(default_factory=list)
    shares: dict[str, dict[str, Decimal]] = field(default_factory=dict)
    unmatched: list[CostCalc] = field(default_factory=list)
    status: str = "non calcolabile"    # completo | parziale | non calcolabile
    included: list[str] = field(default_factory=list)
    excluded: list[tuple[str, str]] = field(default_factory=list)
    revenue: Decimal = ZERO
    supplier_cost: Decimal = ZERO
    internal_cost: Decimal = ZERO
    has_internal: bool = False
    years: list[dict] | None = None
    years_note: str = ""
    months: int | None = None
    threshold: Decimal = Decimal(20)
    messages: list[str] = field(default_factory=list)
    cost_total_all: Decimal = ZERO       # somma dei costi normalizzati (controllo anti doppio conteggio)
    cost_total_allocated: Decimal = ZERO

    @property
    def margin_supplier(self):
        return self.revenue - self.supplier_cost

    @property
    def margin_total(self):
        return self.revenue - self.supplier_cost - self.internal_cost

    @property
    def pct_supplier(self):
        return MarginLine.pct(self.margin_supplier, self.revenue)

    @property
    def pct_total(self):
        return MarginLine.pct(self.margin_total, self.revenue)


def compute_margin(sales: list[SaleRow], costs: list[CostLine], allocs: dict[str, Alloc], months: int | None,
                   fx_rates: dict[str, Decimal], threshold: Decimal, no_cost_ok: set[str] | None = None,
                   qty_confirmed: set[str] | None = None) -> MarginResult:
    no_cost_ok, qty_confirmed = no_cost_ok or set(), qty_confirmed or set()
    res = MarginResult(months=months, threshold=threshold)
    sales_by = {s.id: s for s in sales}
    mlines = {s.id: MarginLine(s, s.revenue, annual_rev=s.annual, once_rev=s.once) for s in sales}
    for ml in mlines.values():
        if ml.revenue is None:
            ml.issues.append("ricavo non calcolabile (periodicità o durata da confermare)")
            ml.status = "mancante"
        if ml.sale.periodicity == "contratto":
            ml.annual_ok = False
    for c in costs:
        cc = normalize_cost(c, months, fx_rates)
        res.costs.append(cc)
        a = allocs.get(c.id, Alloc())
        shares, err = resolve_shares(a, sales_by)
        if err:
            cc.issues.append(err)
        if cc.contract is not None:
            res.cost_total_all += cc.contract
        if shares is None:
            res.unmatched.append(cc)
            if not err:
                cc.issues.append("costo non abbinato a nessuna voce di vendita")
            continue
        res.shares[c.id] = shares
        for sid, sh in shares.items():
            ml = mlines[sid]
            ml.cost_ids.append(c.id)
            if cc.issues:
                ml.issues += [f"costo «{c.desc or c.sku}»: {i}" for i in cc.issues]
                ml.status = "parziale" if ml.status != "mancante" else ml.status
                continue
            amt = cc.contract * sh
            res.cost_total_allocated += amt
            if c.category == "interno":
                ml.internal_cost += amt
            else:
                ml.supplier_cost += amt
            if cc.annual is not None:
                ml.annual_cost = (ml.annual_cost or ZERO) + cc.annual * sh
            if cc.once:
                ml.once_cost += cc.once * sh
            if (a.mode == "single" and c.qty is not None and sales_by[sid].qty is not None and c.qty != sales_by[sid].qty
                    and c.id not in qty_confirmed and c.category == "fornitore"):
                ml.issues.append(f"costo «{c.desc or c.sku}»: quantità {c.qty} diversa da quella dell'offerta ({sales_by[sid].qty}): confermare il confronto a pari quantità")
                ml.status = "parziale"
            if a.suggested:
                ml.warnings.append(f"abbinamento al costo «{c.desc or c.sku}» suggerito automaticamente: da confermare")
    for ml in mlines.values():
        if not ml.cost_ids and ml.sale.id not in no_cost_ok:
            ml.issues.append("nessun costo abbinato")
            ml.status = "mancante"
    res.lines = [m for m in mlines.values() if m.sale.kind == "base"]
    res.options = [m for m in mlines.values() if m.sale.kind == "opzione"]
    for ml in res.lines:
        if ml.status == "completa" and ml.revenue is not None:
            res.included.append(ml.sale.id)
            res.revenue += ml.revenue
            res.supplier_cost += ml.supplier_cost
            res.internal_cost += ml.internal_cost
        else:
            res.excluded.append((ml.sale.id, "; ".join(ml.issues) or "dati insufficienti"))
    res.has_internal = any(c.category == "interno" and (c.total_effective or ZERO) > 0 for c in costs)
    if not res.lines:
        res.status = "non calcolabile"
        res.messages.append("Nessuna riga di vendita base estratta dall'offerta.")
    elif not res.included:
        res.status = "non calcolabile"
    elif res.excluded or res.unmatched:
        res.status = "parziale"
    else:
        res.status = "completo"
    if res.unmatched:
        res.messages.append("Costi non abbinati (margine incompleto): " + "; ".join(f"«{u.cost.desc or u.cost.sku}» ({u.cost.source})" for u in res.unmatched))
    # --- anno per anno (se ricostruibile)
    inc = [m for m in res.lines if m.sale.id in res.included]
    yrs = Decimal(months) / 12 if months else None
    if inc and yrs and yrs == int(yrs) and all(m.annual_ok and (m.annual_rev is not None or m.once_rev is not None) for m in inc):
        rows = []
        for y in range(1, int(yrs) + 1):
            rev = sum(((m.annual_rev or ZERO) + (m.once_rev or ZERO if y == 1 else ZERO) for m in inc), ZERO)
            rows.append({"anno": y, "ricavo": rev})
        # costi per anno: ricalcolati per categoria
        for r in rows:
            r["costo_fornitore"] = r["costo_interno"] = ZERO
        for cc in res.costs:
            if cc.cost.id not in res.shares or cc.issues:
                continue
            for sid, sh in res.shares[cc.cost.id].items():
                if sid not in res.included:
                    continue
                key = "costo_interno" if cc.cost.category == "interno" else "costo_fornitore"
                for r in rows:
                    amt = (cc.annual or ZERO) if cc.cost.periodicity != "una_tantum" else ZERO
                    if cc.cost.periodicity == "una_tantum" and r["anno"] == 1:
                        amt = cc.once or ZERO
                    r[key] += amt * sh
        for r in rows:
            r["margine"] = r["ricavo"] - r["costo_fornitore"] - r["costo_interno"]
            r["margine_pct"] = MarginLine.pct(r["margine"], r["ricavo"])
        res.years = rows
    else:
        res.years_note = "Ripartizione per anno non ricostruibile (durata non intera, periodicità non determinabile o ricavi riferiti all'intero contratto)."
    return res


def check_quote_validity(offer_expiry: date | None, quote_dates: list[tuple[str, date | None]]) -> tuple[str, str]:
    if not quote_dates:
        return "Non applicabile", "Nessuna quotazione fornitore caricata."
    if offer_expiry is None:
        return "Non verificabile", "Scadenza dell'offerta cliente non determinabile (data di partenza o criterio non chiari)."
    bad, unk = [], []
    for name, d in quote_dates:
        if d is None:
            unk.append(name)
        elif d < offer_expiry:
            bad.append(f"{name}: valida fino al {d.strftime('%d/%m/%Y')}")
    if bad:
        return "Errore", f"La quotazione scade prima dell'offerta ({offer_expiry.strftime('%d/%m/%Y')}): " + "; ".join(bad)
    if unk:
        return "Non verificabile", "Validità non indicata in: " + ", ".join(unk)
    return "Conforme", f"Le quotazioni coprono la validità dell'offerta (fino al {offer_expiry.strftime('%d/%m/%Y')})."
