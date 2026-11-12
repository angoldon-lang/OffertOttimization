"""Report interno dei controlli (DOCX). RISERVATO: contiene margini e costi quando disponibili; mai destinato al cliente."""
from __future__ import annotations

import io
from datetime import datetime

from docx import Document
from docx.shared import Pt, RGBColor

from .config import Profile
from .econ import PERIODS
from .money import fmt_eur, fmt_num, fmt_pct, quantize

STATUS_COLOR = {"Errore": "C8102E", "Da verificare": "B26A00", "Non verificabile": "6B6B6B", "Conforme": "1B7F3B", "Non applicabile": "6B6B6B"}


def _table(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    for i, h in enumerate(headers):
        t.rows[0].cells[i].text = h
        for r in t.rows[0].cells[i].paragraphs[0].runs:
            r.bold = True
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
            for p in cells[i].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(8.5)
                    if v in STATUS_COLOR:
                        r.font.color.rgb = RGBColor.from_string(STATUS_COLOR[v])
                        r.bold = True
    return t


def build_report_docx(s, a, profile: Profile) -> bytes:
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)
    p = doc.add_paragraph()
    r = p.add_run("RISERVATO – USO INTERNO AD Consulting. Contiene dati di costo e marginalità: non inviare al cliente.")
    r.bold = True
    r.font.color.rgb = RGBColor(0xC8, 0x10, 0x2E)
    doc.add_heading("Report interno di revisione offerta", 0)
    F = a.fields
    doc.add_paragraph(f"Cliente: {F['cliente'].value} · Offerta n. {F['numero_offerta'].value} · Data {F['data_offerta'].value}\n"
                      f"File: {s.offer_name} · Profilo: {profile.name} · Template: {s.template_name or 'nessuno (solo profilo)'}\n"
                      f"Generato il {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_heading("1. Riepilogo commerciale", 1)
    for sec, rows in a.summary.items():
        doc.add_heading(sec, 2)
        _table(doc, ["Voce", "Valore", "Fonte nell'offerta", "Stato", "Note"], [(r.label, r.value, r.source, r.status, r.note) for r in rows])
    doc.add_heading("2. Esito dei controlli", 1)
    _table(doc, ["Controllo", "Esito", "Rilievi aperti"], [(c["name"], c["status"], c["n"]) for c in a.controls])
    doc.add_heading("3. Anomalie ancora aperte", 1)
    if a.open_findings:
        _table(doc, ["Esito", "Controllo", "Rilievo", "Posizione", "Evidenza", "Correzione proposta", "Stato"],
               [(f.status, f.check, f.title, f.location.label, f.evidence[:300], f.proposal[:200], "") for f in a.open_findings])
    else:
        doc.add_paragraph("Nessuna anomalia aperta.")
    doc.add_heading("4. Tutti i rilievi", 1)
    _table(doc, ["Esito", "Rilievo", "Posizione", "Evidenza", "Stato revisione"],
           [(f.status, f.title, f.location.label, f.evidence[:260], a.states.get(i, "")) for i, f in enumerate(a.findings)])
    doc.add_heading("5. Righe economiche e totali (ricalcolo)", 1)
    if a.econ.lines:
        _table(doc, ["Fonte", "Descrizione", "Periodicità", "Formula", "Riportato", "Calcolato", "Diff.", "Esito"],
               [(l.source, l.desc[:40], PERIODS.get(l.periodicity or "", "da confermare"), l.formula, fmt_num(l.reported), fmt_num(l.calc), fmt_num(l.diff), l.status) for l in a.econ.lines])
    if a.econ.totals:
        doc.add_paragraph()
        _table(doc, ["Fonte", "Totale", "Formula", "Dichiarato", "Ricalcolato", "Diff.", "Esito"],
               [(f"Tabella {t.table + 1}, riga {t.row + 1}", t.label, t.formula, fmt_num(t.reported), fmt_num(t.calc), fmt_num(t.diff), t.status) for t in a.econ.totals])
    doc.add_paragraph(f"Arrotondamento: {a.rounding}, {a.decimals} decimali.")
    doc.add_heading("6. Modifiche approvate e applicate al Word", 1)
    if a.history:
        _table(doc, ["Tipo", "Posizione", "Prima", "Dopo", "Origine"],
               [(c.label, f"par. {c.para + 1}" if c.para is not None else f"Tab. {c.table + 1}, riga {c.row + 1}", c.ctx_before or c.before, c.ctx_after, c.origin) for c in a.history])
    else:
        doc.add_paragraph("Nessuna modifica approvata: il DOCX cliente mantiene i contenuti commerciali dell'originale.")
    doc.add_heading("7. Margine (riservato)", 1)
    m = a.margin
    if m is None:
        doc.add_paragraph("Non calcolato: nessuna quotazione fornitore o costo caricato.")
    else:
        doc.add_paragraph(f"Stato: {m.status.upper()}. Voci incluse: {', '.join(m.included) or 'nessuna'}. "
                          + (f"Escluse: {'; '.join(f'{i} ({r})' for i, r in m.excluded)}. " if m.excluded else ""))
        if m.included:
            doc.add_paragraph(f"Ricavo {fmt_eur(quantize(m.revenue))} · costo fornitore {fmt_eur(quantize(m.supplier_cost))}"
                              + (f" · costi interni {fmt_eur(quantize(m.internal_cost))}" if m.has_internal else "")
                              + f" · margine solo fornitore {fmt_eur(quantize(m.margin_supplier))} ({fmt_pct(quantize(m.pct_supplier) if m.pct_supplier is not None else None)})"
                              + (f" · margine con costi interni {fmt_eur(quantize(m.margin_total))} ({fmt_pct(quantize(m.pct_total) if m.pct_total is not None else None)})" if m.has_internal else ""))
        _table(doc, ["Voce", "Descrizione", "Ricavo", "Costo fornitore", "Costo interno", "Margine €", "Margine %", "Stato"],
               [(ml.sale.id, ml.sale.desc[:36], fmt_num(ml.revenue), fmt_num(ml.supplier_cost), fmt_num(ml.internal_cost),
                 fmt_num(ml.margin_total() if m.has_internal else ml.margin_supplier()),
                 fmt_pct(quantize(ml.pct(ml.margin_total() if m.has_internal else ml.margin_supplier(), ml.revenue)) if ml.pct(ml.margin_total() if m.has_internal else ml.margin_supplier(), ml.revenue) is not None else None),
                 ml.status) for ml in m.lines])
        doc.add_paragraph()
        _table(doc, ["Costo", "Categoria", "Fonte", "Descrizione", "Costo contratto € (netto IVA)", "Abbinamento", "Criticità"],
               [(cc.cost.id, cc.cost.category, cc.cost.source, cc.cost.desc[:40], fmt_num(cc.contract),
                 (a.allocs[cc.cost.id].criterion + ": " + ", ".join(a.allocs[cc.cost.id].targets)) if cc.cost.id in a.allocs else "",
                 "; ".join(cc.issues + cc.warnings)) for cc in m.costs])
        for msg in m.messages:
            doc.add_paragraph(msg)
    doc.add_heading("8. Limiti noti", 1)
    for t in ["La pagina non è calcolabile in modo affidabile da un DOCX: le posizioni sono indicate come sezione/tabella/riga.",
              "Tabelle annidate, caselle di testo, controlli di contenuto e revisioni tracciate sono conservati ma non analizzati.",
              "I controlli ortografici usano un dizionario di refusi configurabile e regole di punteggiatura, non un correttore completo.",
              "OCR dei PDF scansionati disponibile solo se installati tesseract e pytesseract; altrimenti griglia manuale."]:
        doc.add_paragraph(t, style="List Bullet")
    for n in s.notices:
        doc.add_paragraph(n, style="List Bullet")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
