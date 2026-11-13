"""Conversione di campi estratti e risultati di calcolo in rilievi (Finding)."""
from __future__ import annotations

from datetime import date, timedelta

import re

from .docx_utils import Change, DocView, Location, norm
from .econ import Econ, change_for_cell
from .extract import FieldResult, add_months, fmt_date
from .findings import Finding
from .fixes import invalid_date_fix, loc_change
from .money import D, find_numbers, fmt_eur, fmt_num, parse_number, parse_percent, quantize


def offer_expiry(fields: dict[str, FieldResult]) -> tuple[date | None, str, Location | None]:
    v = fields["validita"]
    if not v.present:
        return None, "Validità non indicata", None
    o = next((x for x in v.occurrences if x.data.get("date") or "raw" not in x.data), v.occurrences[0])
    if "raw" in o.data and not o.data.get("date"):
        return None, f"data «{o.value}» non esistente", o.loc
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
        elif f.status == "Da verificare" and f.present:
            prop = "Chiarire la clausola (es. per il rinnovo: modalità, durata e termine di disdetta)" if k == "rinnovo" else "Verificare il dato"
        out.append(Finding("campi", f"{f.label}", f.status, loc, ev, prop, kind="sostanziale" if f.status != "Conforme" else "",
                           needs_input=f.status in ("Da verificare", "Non verificabile") and not f.present))
    # allineamento di valori discordanti e correzione di date inesistenti (proposte approvabili)
    for k in ("numero_offerta", "cliente", "data_offerta"):
        f = fields[k]
        if f.status == "Errore" and len(f.occurrences) > 1:
            primary = f.occurrences[0]
            for o in f.occurrences[1:]:
                if o.value.strip() != primary.value.strip() and o.value in o.source:
                    ch = loc_change(o.loc, o.source, o.value, primary.value, f"{f.label} discordante", "clausola", f"allineato a «{primary.value}» ({primary.loc.label})")
                    out.append(Finding("campi", f"{f.label} discordante", "Errore", o.loc, f"«{o.value}» invece di «{primary.value}» ({primary.loc.label}).",
                                       f"Allineare a «{primary.value}» (modifica il valore se è quello sbagliato)", ch, kind="sostanziale"))
    vf = fields["validita"]
    good = [o.value for o in vf.occurrences if "raw" not in o.data or o.data.get("date")]
    for o in vf.occurrences:
        if "raw" in o.data and not o.data.get("date") and o.value in o.source:
            sug = invalid_date_fix(o.value, good)
            ch = loc_change(o.loc, o.source, o.value, sug, "Data di validità inesistente", "data", "la data indicata non esiste in calendario") if sug != o.value else None
            out.append(Finding("campi", "Data di validità inesistente", "Errore", o.loc,
                               f"«{o.value}» non è una data valida" + (f"; altrove nel documento: {good[0]}" if good else ""),
                               f"Sostituire con «{sug}» (modificabile)", ch, kind="sostanziale"))
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
        elif l.status in ("Non verificabile", "Da verificare") and not (l.net is not None or l.gross is not None):
            out.append(Finding("calcoli", "Riga non ricalcolabile" if l.status == "Non verificabile" else "Importo di riga non indicato", l.status,
                               l.location, f"{l.desc[:40]}: {l.note}", "Completare i dati nella tabella di revisione", needs_input=True))
        else:
            ok_lines += 1
        if l.gross_status == "Errore" and l.col_gross is not None:
            ch = change_for_cell(l.table, l.row, l.col_gross, l.cell_text["gross"], l.gross_calc, f"Netto + IVA {l.source}", "prezzo", decimals=2,
                                 reason=l.gross_note)
            out.append(Finding("calcoli", "«Netto + IVA» di riga non coerente", "Errore", Location(l.section, table=l.table, row=l.row, col=l.col_gross, kind="cell"),
                               f"{l.desc[:40]}: {l.gross_note}", f"Sostituire {fmt_num(l.gross)} con {fmt_num(l.gross_calc)}", ch))
        if l.periodicity is None and l.kind == "base":
            out.append(Finding("calcoli", "Periodicità/unità da confermare", "Da verificare", l.location,
                               f"{l.desc[:40]}: non è chiaro se l'importo sia una tantum, mensile o annuale.",
                               "Confermare la periodicità nella tabella di revisione (schermata Calcoli)", needs_input=True))
    for t in econ.totals:
        if t.status == "Errore" and t.calc is not None:
            ch = change_for_cell(t.table, t.row, t.col_amount, t.cell_text, t.calc, t.label, "totale", decimals=2, reason=t.formula)
            col = f" – colonna «{t.header}»" if t.role != "amount" and t.header else (f" – colonna «{t.header}»" if t.header else "")
            out.append(Finding("calcoli", f"Totale errato: «{t.label}»{col}", "Errore", t.location,
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
    if econ.assumed:
        out.append(Finding("calcoli", "Periodicità non indicata: assunta «una tantum»", "Da verificare", Location(kind="doc"),
                           f"Nessuna indicazione di durata o periodicità nel documento ({len(econ.assumed)} righe): gli importi sono trattati come una tantum.",
                           "Se sono canoni, indicare periodicità e durata nella scheda Calcoli", needs_input=True))
    if econ.inferred_vat is not None:
        out.append(Finding("calcoli", f"Aliquota IVA dedotta dalle righe: {fmt_num(econ.inferred_vat, 0 if econ.inferred_vat == econ.inferred_vat.to_integral_value() else 2)}%", "Conforme",
                           Location(kind="doc"), "Rapporto «Netto + IVA» / «Netto» uniforme sulle righe; l'aliquota non è dichiarata in modo esplicito."))
    if ok_lines or ok_tot:
        out.append(Finding("calcoli", "Righe e totali coerenti con il ricalcolo", "Conforme", Location(kind="doc"),
                           f"{ok_lines} righe e {ok_tot} totali verificati (arrotondamento configurato: vedi Impostazioni)."))
    return out


def _units(view: DocView):
    from .textcheck import units
    return units(view)


def check_discount_text(view: DocView, econ: Econ) -> list[Finding]:
    """«ulteriore sconto del 3% per un importo complessivo di X»: verifica X rispetto ai totali ricalcolati (netto o IVA inclusa)."""
    out = []
    rx = re.compile(r"(?i)sconto\s+(?:ulteriore\s+)?(?:del\s+|di\s+)?(?P<p>\d+(?:[.,]\d+)?)\s*%[^€\n]*?(?:importo|totale|pari|ammontare)[^€\d\n]*?(?P<a>\d[\d.]*(?:,\d+)?)\s*(?:€|euro|EUR)")
    for text, loc, params in _units(view):
        m = rx.search(text)
        if not m:
            continue
        pct, stated = parse_number(m.group("p")), parse_number(m.group("a"))
        if pct is None or stated is None:
            continue
        cands = []
        if econ.net_total is not None:
            cands.append(("netto IVA esclusa", econ.net_total))
        if econ.gross_total is not None:
            cands.append(("IVA inclusa", econ.gross_total))
        if not cands and econ.contract is not None:
            cands.append(("netto IVA esclusa", econ.contract))
        if not cands:
            out.append(Finding("calcoli", "Sconto per pagamento anticipato non verificabile", "Non verificabile", loc, f"Totali non ricalcolabili: «{m.group(0)[:80]}»", "Verificare a mano", needs_input=True))
            continue
        f_ = 1 - pct / 100
        vals = [(lab, quantize(v * f_), v) for lab, v in cands]
        if any(v == stated for _, v, _ in vals):
            out.append(Finding("calcoli", "Sconto per pagamento anticipato coerente", "Conforme", loc, f"{fmt_num(stated)} € = totale ricalcolato × {fmt_num(f_ * 100, 0)}%"))
            continue
        best = min(vals, key=lambda x: abs(x[1] - stated))
        ev = "; ".join(f"{lab}: {fmt_num(base)} × {fmt_num(f_ * 100, 0)}% = {fmt_num(v)}" for lab, v, base in vals)
        hint = ""
        for t in econ.totals:
            if t.role == "net" and t.reported is not None and t.calc is not None and t.reported != t.calc and econ.inferred_vat is not None:
                if abs(quantize(t.reported * (1 + econ.inferred_vat / 100) * f_) - stated) <= D("1"):
                    hint = f" L'importo dichiarato deriva dal totale netto errato ({fmt_num(t.reported)})."
        tok = m.group("a") + ("" if False else "")
        ch = Change(params["kind"], tok, fmt_num(best[1], 2), "Importo sconto anticipato", "sconto", table=params.get("table"), row=params.get("row"),
                    col=params.get("col"), para=params.get("para"), ctx_before=text, reason=f"{best[0]}: {fmt_num(best[2])} × {fmt_num(f_ * 100, 0)}%")
        out.append(Finding("calcoli", "Importo dello sconto per pagamento anticipato non coerente", "Errore", loc,
                           f"Dichiarato {fmt_num(stated)} €. Ricalcolo ({fmt_num(pct, 0)}% di sconto): {ev}.{hint}",
                           f"Sostituire {fmt_num(stated)} con {fmt_num(best[1])} (base: {best[0]}; verificare quale base voluta)", ch))
    return out


def check_year_summary(view: DocView, econ: Econ) -> list[Finding]:
    """Riepilogo «Totale AAAA | importo»: confronto con i totali per periodo (netto o IVA inclusa) e coerenza di etichette."""
    out = []
    if not econ.period_rows:
        return out
    prev_text = ""
    for kind, it in view.items:
        if kind == "p":
            if it.text.strip():
                prev_text = it.text.strip()
            continue
        rows = []
        for r, cells in enumerate(it.rows):
            m = re.match(r"^\s*totale\s+(?:anno\s+)?(\d{4})\s*$", cells[0].text, re.I) if cells else None
            if m and len(cells) >= 2 and find_numbers(cells[-1].text):
                rows.append((int(m.group(1)), r, cells[-1], parse_number(find_numbers(cells[-1].text)[0][0])))
        if not rows:
            continue
        header = " ".join(c.text for c in it.rows[0]) if it.rows else ""
        want = "gross" if re.search(r"(?i)con\s+iva|iva\s+incl|lordo", prev_text + " " + header) else ("net" if re.search(r"(?i)netto|iva\s+escl", header) else None)
        caption_gross = bool(re.search(r"(?i)con\s+iva|iva\s+incl", prev_text))
        header_net = bool(re.search(r"(?i)netto", header)) and not re.search(r"(?i)iva\s+incl|lordo", header)
        by_year = {p["start"].year: p for p in econ.period_rows}
        mism, basis_hits = [], []
        for y, r, c, v in rows:
            p = by_year.get(y)
            if not p or v is None:
                continue
            hit = "gross" if v == quantize(p["gross"]) else ("net" if v == quantize(p["net"]) else None)
            basis_hits.append(hit)
            exp_key = "gross" if caption_gross else ("net" if header_net else None)
            if exp_key and hit != exp_key and not (hit and exp_key == "gross" and caption_gross and header_net):
                mism.append((y, r, c, v, quantize(p[exp_key]), exp_key))
        loc0 = Location(it.section, table=it.idx, row=0, kind="cell")
        if caption_gross and header_net and basis_hits and all(h == "gross" for h in basis_hits):
            hc = it.rows[0][-1]
            ch = Change("cell", "Netto", "IVA inclusa", "Intestazione riepilogo annuale", "testo", table=it.idx, row=0, col=hc.col, ctx_before=hc.text,
                        reason="gli importi coincidono con «Netto + IVA» e il titolo dice «con IVA»")
            out.append(Finding("calcoli", "Riepilogo annuale: intestazione «Netto» ma importi IVA inclusa", "Errore", loc0,
                               f"Il titolo («{prev_text[:50]}») e gli importi (= Netto + IVA per anno) indicano IVA inclusa, la colonna è intestata «{hc.text}».",
                               "Rinominare l'intestazione in «Totale (IVA inclusa)»", ch if hc.text.count("Netto") else None))
        for y, r, c, v, exp, key in mism:
            ch = change_for_cell(it.idx, r, c.col, c.text, exp, f"Totale {y}", "totale", decimals=2, reason=f"somma {'Netto + IVA' if key == 'gross' else 'Netto'} del periodo {y}")
            out.append(Finding("calcoli", f"Riepilogo annuale non coerente: Totale {y}", "Errore", Location(it.section, table=it.idx, row=r, col=c.col, kind="cell"),
                               f"Dichiarato {fmt_num(v)} €, somma {'Netto + IVA' if key == 'gross' else 'Netto'} del periodo {fmt_num(exp)} €.", f"Sostituire con {fmt_num(exp)}", ch))
        if basis_hits and all(h for h in basis_hits) and not mism and not (caption_gross and header_net):
            out.append(Finding("calcoli", "Riepilogo annuale coerente con i totali per periodo", "Conforme", loc0, "Importi verificati per anno."))
    return out


def check_references(view: DocView) -> list[Finding]:
    """«si veda il paragrafo X» con X non presente tra i titoli del documento."""
    out = []
    heads = [norm(re.sub(r"^\d+(?:\.\d+)*[.)]?\s*", "", h.text)) for h in view.headings]
    rx = re.compile(r"(?i)(?:si\s+veda|vedi|cfr\.?|rimando\s+a)\s+(?:il\s+|la\s+|l['’]\s*)?(?:paragrafo|sezione|capitolo)\s+[«\"“]?(?P<t>[A-Za-zÀ-ÿ][^.»\"”\n,;()]{2,60})")
    for p in view.paras:
        for m in rx.finditer(p.text):
            t = norm(m.group("t"))
            if not t or any(t == h or t in h or (h and h in t) for h in heads):
                continue
            out.append(Finding("coerenza", "Riferimento a un paragrafo non presente nel documento", "Da verificare", Location(p.section, paragraph=p.idx, kind="par"),
                               f"«{m.group(0).strip()[:90]}»: nel documento non c'è un titolo «{m.group('t').strip()}».",
                               "Aggiungere il paragrafo richiamato o correggere il riferimento", kind="sostanziale"))
    return out
