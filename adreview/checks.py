"""Conversione di campi estratti e risultati di calcolo in rilievi (Finding)."""
from __future__ import annotations

from datetime import date, timedelta

from .docx_utils import Location
from .econ import Econ, change_for_cell
from .extract import FieldResult, add_months, fmt_date
from .findings import Finding
from .money import fmt_eur, fmt_num


def offer_expiry(fields: dict[str, FieldResult]) -> tuple[date | None, str, Location | None]:
    v = fields["validita"]
    if not v.present:
        return None, "Validità non indicata", None
    o = v.occurrences[0]
    if o.data.get("date"):
        return o.data["date"], "data indicata nel documento", o.loc
    od = fields["data_offerta"]
    start = od.occurrences[0].data.get("date") if od.occurrences else None
    if not o.data.get("clear"):
        return None, "criterio di decorrenza della validità non chiaro (es. «dalla data dell'offerta»)", o.loc
    if start is None:
        return None, "data di partenza (data offerta) non leggibile", o.loc
    if o.data.get("days"):
        return start + timedelta(days=o.data["days"]), f"{o.data['days']} giorni dalla data offerta {fmt_date(start)}", o.loc
    if o.data.get("months"):
        return add_months(start, o.data["months"]), f"{o.data['months']} mesi dalla data offerta {fmt_date(start)}", o.loc
    return None, "", o.loc


def check_fields(fields: dict[str, FieldResult], today: date) -> list[Finding]:
    out = []
    for k, f in fields.items():
        loc = f.loc or Location(kind="doc")
        ev = f.value if f.value else "Non indicato"
        if f.status == "Errore" and len(f.occurrences) > 1:
            ev += " — " + "; ".join(f"{o.value[:40]} ({o.loc.label})" for o in f.occurrences[:4])
        if f.note:
            ev += f" — {f.note}"
        prop = ""
        if f.status == "Da verificare" and not f.present:
            prop = "Inserire il dato (richiesto dal profilo): non modificato automaticamente"
        elif f.status == "Errore":
            prop = "Allineare i valori tra le sezioni: la scelta del valore corretto spetta all'utente"
        elif f.status == "Non verificabile":
            prop = "Verificare manualmente il valore"
        out.append(Finding("campi", f"{f.label}", f.status, loc, ev, prop, kind="sostanziale" if f.status != "Conforme" else "",
                           needs_input=f.status in ("Da verificare", "Non verificabile") and not f.present))
    exp, note, eloc = offer_expiry(fields)
    v = fields["validita"]
    if v.present:
        if exp:
            if exp < today:
                out.append(Finding("campi", "Validità dell'offerta scaduta", "Da verificare", eloc or Location(kind="doc"),
                                   f"Scadenza {fmt_date(exp)} ({note}) precedente alla data odierna {fmt_date(today)}", "Verificare se aggiornare la validità"))
            od = fields["data_offerta"].occurrences[0].data.get("date") if fields["data_offerta"].occurrences else None
            if od and exp < od:
                out.append(Finding("campi", "Scadenza precedente alla data dell'offerta", "Errore", eloc or Location(kind="doc"),
                                   f"Scadenza {fmt_date(exp)} < data offerta {fmt_date(od)}", "Correggere la data di validità"))
        else:
            out.append(Finding("campi", "Scadenza dell'offerta non calcolabile", "Non verificabile", eloc or Location(kind="doc"),
                               note, "Esplicitare data di partenza e criterio (es. «30 giorni dalla data dell'offerta»)"))
    d = fields["durata"]
    if d.present and d.data.get("decorrenza") and d.data.get("scadenza"):
        (dec, dl), (sc, _) = d.data["decorrenza"], d.data["scadenza"]
        m = d.data.get("months")
        if dec and sc and m:
            exp_end = add_months(dec, m)
            if abs((exp_end - sc).days) > 1:
                out.append(Finding("campi", "Decorrenza/scadenza non coerenti con la durata", "Errore", dl,
                                   f"Decorrenza {fmt_date(dec)} + {m} mesi = {fmt_date(exp_end)}, indicata scadenza {fmt_date(sc)}", "Allineare le date alla durata"))
    return out


def check_econ(econ: Econ, months_known: bool, has_tables: bool) -> list[Finding]:
    out = []
    if not econ.lines and not econ.totals:
        return [Finding("calcoli", "Righe economiche", "Non verificabile" if has_tables else "Da verificare", Location(kind="doc"),
                        "Nessuna tabella economica riconoscibile (servono colonne tipo Descrizione / Quantità / Prezzo / Importo).",
                        "Verificare la struttura della tabella oppure inserire le righe a mano nella schermata Calcoli", needs_input=True)]
    ok_lines = ok_tot = 0
    for l in econ.lines:
        if l.status == "Errore" and l.calc is not None:
            ch = change_for_cell(l.table, l.row, l.col_amount, l.cell_text["amount"], l.calc, f"Importo riga {l.source}", "prezzo",
                                 decimals=2, reason=l.formula) if l.col_amount is not None else None
            out.append(Finding("calcoli", "Importo di riga non coerente con quantità × prezzo", "Errore", l.location,
                               f"{l.desc[:40]}: formula {l.formula}; riportato {fmt_eur(l.reported)}, calcolato {fmt_eur(l.calc)}, differenza {fmt_eur(l.diff)}. {l.note}".strip(),
                               f"Sostituire {fmt_num(l.reported)} con {fmt_num(l.calc)}", ch))
        elif l.status in ("Non verificabile", "Da verificare"):
            out.append(Finding("calcoli", "Riga non ricalcolabile" if l.status == "Non verificabile" else "Importo di riga non indicato", l.status,
                               l.location, f"{l.desc[:40]}: {l.note}", "Completare i dati nella tabella di revisione", needs_input=True))
        else:
            ok_lines += 1
        if l.periodicity is None and l.kind == "base":
            out.append(Finding("calcoli", "Periodicità/unità da confermare", "Da verificare", l.location,
                               f"{l.desc[:40]}: non è chiaro se l'importo sia una tantum, mensile o annuale.",
                               "Confermare la periodicità nella tabella di revisione (schermata Calcoli)", needs_input=True))
    for t in econ.totals:
        if t.status == "Errore" and t.calc is not None:
            ch = change_for_cell(t.table, t.row, t.col_amount, t.cell_text, t.calc, t.label, "totale", decimals=2, reason=t.formula)
            out.append(Finding("calcoli", f"Totale errato: «{t.label}»", "Errore", t.location,
                               f"Dichiarato {fmt_eur(t.reported)} · ricalcolato {fmt_eur(t.calc)} · differenza {fmt_eur(t.diff)}. Formula: {t.formula}. {t.note}".strip(),
                               f"Sostituire {fmt_num(t.reported)} con {fmt_num(t.calc)} (il Word resta invariato finché non approvi)", ch))
        elif t.status in ("Non verificabile", "Da verificare"):
            out.append(Finding("calcoli", f"Totale non verificabile: «{t.label}»", t.status, t.location, t.note or "Dati insufficienti", "Completare/confermare i dati", needs_input=True))
        elif t.status == "Conforme":
            ok_tot += 1
    if econ.months is None and any(l.periodicity not in (None, "una_tantum", "contratto") for l in econ.base_lines()):
        out.append(Finding("calcoli", "Durata del contratto non indicata", "Da verificare", Location(kind="doc"),
                           "Ci sono voci ricorrenti ma la durata non è nota: totali per contratto non ricalcolabili.",
                           "Indicare/confermare la durata (mesi) nella schermata Calcoli", needs_input=True))
    if ok_lines or ok_tot:
        out.append(Finding("calcoli", "Righe e totali coerenti con il ricalcolo", "Conforme", Location(kind="doc"),
                           f"{ok_lines} righe e {ok_tot} totali verificati (arrotondamento configurato: vedi Impostazioni)."))
    return out
