"""Riepilogo commerciale interno: ogni informazione riporta il punto dell'offerta da cui proviene."""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal

from .checks import offer_expiry
from .docx_utils import Location
from .econ import Econ
from .extract import fmt_date
from .money import fmt_eur, fmt_num, fmt_pct


@dataclass
class SumRow:
    label: str
    value: str
    source: str = ""
    status: str = ""     # Conforme | Da verificare | Errore | Non verificabile | Non indicato
    note: str = ""


def _src(loc: Location | None) -> str:
    return loc.label if loc else "—"


def build_summary(fields: dict, econ: Econ, margin, margin_base: str, quote_validity: tuple[str, str] | None) -> dict[str, list[SumRow]]:
    S: dict[str, list[SumRow]] = {"Valori": [], "Durata": [], "Rinnovo": [], "Validità offerta": [], "Margine": []}
    V = S["Valori"]
    tot = {t.ctype: t for t in econ.totals}
    contract_rows = [t for t in econ.totals if t.ctype == "contratto"] or [t for t in econ.totals if t.ctype == "generico"]
    net_t = contract_rows[-1] if contract_rows else None
    if econ.contract is not None:
        note = ""
        st = "Conforme"
        if net_t and net_t.reported is not None and net_t.reported != econ.contract:
            st, note = "Errore", f"Dichiarato in offerta: {fmt_eur(net_t.reported)}"
        V.append(SumRow("Totale offerto (netto IVA, ricalcolato)", fmt_eur(econ.contract), _src(net_t.location if net_t else None), st, note))
    elif net_t and net_t.reported is not None:
        V.append(SumRow("Totale offerto (netto IVA, dichiarato)", fmt_eur(net_t.reported), _src(net_t.location), "Non verificabile", "Ricalcolo non possibile (periodicità/durata da confermare)"))
    else:
        V.append(SumRow("Totale offerto (netto IVA)", "Non indicato", "—", "Non indicato"))
    base = net_t.calc if net_t and net_t.calc is not None else econ.contract
    iva = tot.get("iva")
    lordo = tot.get("lordo")
    if iva:
        V.append(SumRow(f"IVA{(' ' + fmt_num(iva.data_rate, 0) + '%') if iva.data_rate is not None else ''}",
                        fmt_eur(iva.calc if iva.calc is not None else iva.reported), _src(iva.location), iva.status, iva.note))
    else:
        V.append(SumRow("IVA", "Non indicata", "—", "Non indicato", "IVA assente: nessuna verifica applicabile"))
    if lordo:
        V.append(SumRow("Totale lordo (IVA inclusa)", fmt_eur(lordo.calc if lordo.calc is not None else lordo.reported), _src(lordo.location), lordo.status, lordo.note))
    else:
        V.append(SumRow("Totale lordo (IVA inclusa)", "Non indicato", "—", "Non indicato"))
    py = econ.per_year()
    if py:
        for r in py:
            V.append(SumRow(f"Anno {r['anno']}", fmt_eur(r["importo"]), "canoni annuali" + (" + una tantum" if r["anno"] == 1 and econ.once else ""), "Conforme"))
    elif econ.lines:
        V.append(SumRow("Dettaglio per anno", "Non ricostruibile", "—", "Non verificabile", "Durata non intera o periodicità da confermare"))
    V.append(SumRow("Una tantum", fmt_eur(econ.once), ", ".join(l.source for l in econ.base_lines() if l.periodicity == "una_tantum") or "—", "Conforme" if not econ.ambiguous else "Da verificare"))
    V.append(SumRow("Ricorrenti (per anno)", fmt_eur(econ.annual) if econ.annual is not None else "Non calcolabile",
                    ", ".join(l.source for l in econ.base_lines() if l.periodicity not in (None, "una_tantum")) or "—",
                    "Conforme" if econ.annual is not None else "Non verificabile"))
    V.append(SumRow("Opzioni (escluse dai totali)", fmt_eur(econ.options) if econ.options else "Nessuna",
                    ", ".join(l.source for l in econ.lines if l.kind == "opzione") or "—", "Conforme"))
    # Durata
    d = fields["durata"]
    D_ = S["Durata"]
    D_.append(SumRow("Durata contratto", f"{d.data['months']} mesi" if d.data.get("months") else "Non indicata", _src(d.loc), d.status if d.present else "Non indicato",
                     d.note + (" (confermata dall'utente)" if d.data.get("confirmed") else "")))
    dec, sc = d.data.get("decorrenza"), d.data.get("scadenza")
    D_.append(SumRow("Decorrenza", fmt_date(dec[0]) if dec else "Non indicata", _src(dec[1]) if dec else "—", "Conforme" if dec else "Non indicato"))
    D_.append(SumRow("Scadenza", fmt_date(sc[0]) if sc else "Non indicata", _src(sc[1]) if sc else "—", "Conforme" if sc else "Non indicato"))
    # Rinnovo
    r = fields["rinnovo"]
    ts = set(r.data.get("types", []))
    S["Rinnovo"].append(SumRow("Rinnovo", ("Automatico" if "automatico" in ts else "Espresso" if "espresso" in ts else "Assente (dichiarato)" if "assente" in ts else r.value) if r.present else "Non indicato",
                               _src(r.loc), r.status if r.present else "Non indicato", r.note))
    S["Rinnovo"].append(SumRow("Durata del rinnovo", f"{r.data['renew_months']} mesi" if r.data.get("renew_months") else "Non indicata", _src(r.loc), "Conforme" if r.data.get("renew_months") else "Non indicato"))
    S["Rinnovo"].append(SumRow("Termine di disdetta", r.data.get("disdetta") or "Non indicato", _src(r.loc), "Conforme" if r.data.get("disdetta") else "Non indicato"))
    # Validità
    v = fields["validita"]
    exp, note, eloc = offer_expiry(fields)
    S["Validità offerta"].append(SumRow("Validità", v.value if v.present else "Non indicata", _src(v.loc), v.status if v.present else "Non indicato", v.note))
    S["Validità offerta"].append(SumRow("Data di scadenza", fmt_date(exp) if exp else ("Non calcolabile" if v.present else "Non indicata"), _src(eloc),
                                        "Conforme" if exp else ("Non verificabile" if v.present else "Non indicato"), note if v.present else ""))
    if quote_validity:
        S["Validità offerta"].append(SumRow("Copertura validità quotazione fornitore", quote_validity[1], "quotazioni caricate", quote_validity[0]))
    # Margine
    M = S["Margine"]
    if margin is None:
        M.append(SumRow("Margine", "Non calcolabile", "—", "Non verificabile", "Nessuna quotazione fornitore o costo caricato: la revisione Word e i calcoli dell'offerta restano disponibili."))
    else:
        label = {"completo": "COMPLETO", "parziale": "PARZIALE", "non calcolabile": "NON CALCOLABILE"}[margin.status]
        M.append(SumRow("Stato del calcolo", label, "report margine", "Conforme" if margin.status == "completo" else "Da verificare",
                        "Voci incluse: " + (", ".join(margin.included) or "nessuna") + ("; escluse: " + ", ".join(i for i, _ in margin.excluded) if margin.excluded else "")))
        if margin.included:
            use_int = margin_base == "con_interni" and margin.has_internal
            m_eur, m_pct = (margin.margin_total, margin.pct_total) if use_int else (margin.margin_supplier, margin.pct_supplier)
            st = "Errore" if m_eur < 0 else ("Da verificare" if m_pct is not None and m_pct < margin.threshold else "Conforme")
            M.append(SumRow("Margine €" + (" (parziale)" if margin.status != "completo" else ""), fmt_eur(m_eur), "voci incluse", st,
                            "sotto soglia" if st == "Da verificare" else ("negativo" if st == "Errore" else "")))
            M.append(SumRow("Margine %", fmt_pct(m_pct), "margine ÷ ricavo netto", st))
            M.append(SumRow("Base di costo", "costi fornitore + costi diretti interni" if use_int else "solo costi fornitore", "", "Conforme"))
    return S
