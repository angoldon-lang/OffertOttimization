"""Calcoli economici su casi più ricchi: una tantum, canoni mensili, opzioni, sconto in riga, subtotali, IVA."""
import io
from decimal import Decimal as D

from docx import Document

from adreview.docx_utils import DocView, open_docx
from adreview.econ import extract_and_calc


def build(rows, months_text="Durata del contratto: 24 mesi."):
    d = Document()
    d.add_paragraph(months_text)
    t = d.add_table(rows=len(rows), cols=len(rows[0]))
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            t.cell(i, j).text = v
    buf = io.BytesIO(); d.save(buf)
    return DocView(open_docx(buf.getvalue()))


H = ["Descrizione", "Q.tà", "Prezzo unitario", "Sconto", "Importo", "Periodicità"]


def test_mix_una_tantum_canoni_opzioni_iva():
    v = build([H,
               ["Installazione", "2", "450,00 €", "", "900,00 €", "Una tantum"],
               ["Canone assistenza", "1", "100,00 €", "", "100,00 €", "Mensile"],
               ["Modulo extra (opzionale)", "1", "1.000,00 €", "", "1.000,00 €", "Una tantum"],
               ["Totale contratto 24 mesi", "", "", "", "3.300,00 €", ""],
               ["IVA 22%", "", "", "", "726,00 €", ""],
               ["Totale IVA inclusa", "", "", "", "4.026,00 €", ""]])
    e = extract_and_calc(v, 24)
    assert e.once == D("900.00") and e.annual == D("1200.00") and e.contract == D("3300.00")
    assert e.options == D("1000.00")                       # opzione esclusa dai totali
    st = {t.ctype: (t.status, t.calc) for t in e.totals}
    assert st["contratto"] == ("Conforme", D("3300.00"))
    assert st["iva"] == ("Conforme", D("726.00")) and st["lordo"] == ("Conforme", D("4026.00"))
    assert [r["importo"] for r in e.per_year()] == [D("2100.00"), D("1200.00")]


def test_errori_riga_iva_e_sconto_non_applicato():
    v = build([H,
               ["Licenza", "10", "100,00 €", "10%", "1.000,00 €", "Annuale"],
               ["Totale annuo", "", "", "", "1.000,00 €", ""],
               ["IVA 22%", "", "", "", "230,00 €", ""]])
    e = extract_and_calc(v, 12)
    l = e.lines[0]
    assert l.status == "Errore" and l.calc == D("900.00") and l.diff == D("100.00") and "sconto non applicato" in l.note
    assert [t for t in e.totals if t.ctype == "anno"][0].status == "Errore"
    iva = [t for t in e.totals if t.ctype == "iva"][0]
    assert iva.status == "Errore" and iva.calc == D("198.00")          # 22% sul netto ricalcolato (900), non su quello riportato


def test_durata_assente_rende_non_verificabile_il_totale_contratto():
    v = build([H, ["Canone", "1", "100,00 €", "", "100,00 €", "Mensile"], ["Totale triennale", "", "", "", "3.600,00 €", ""]], months_text="Nessuna durata.")
    e = extract_and_calc(v, None)
    assert [t for t in e.totals if t.ctype == "contratto"][0].status == "Non verificabile"
    assert e.contract is None


def test_periodicita_ambigua_da_confermare_e_conferma():
    v = build([H, ["Servizio X", "1", "500,00 €", "", "500,00 €", ""], ["Totale annuo", "", "", "", "500,00 €", ""]])
    e = extract_and_calc(v, 12)
    assert e.lines[0].periodicity is None and e.ambiguous
    assert [t for t in e.totals if t.ctype == "anno"][0].status == "Non verificabile"
    e2 = extract_and_calc(v, 12, {"lines": {"T1R2": {"periodicita": "annuale"}}})
    assert e2.lines[0].periodicity == "annuale" and e2.lines[0].period_src == "confermata"
    assert [t for t in e2.totals if t.ctype == "anno"][0].status == "Conforme"


def test_sconto_in_riga_e_subtotali():
    v = build([H,
               ["A", "1", "1.000,00 €", "", "1.000,00 €", "Una tantum"],
               ["B", "1", "500,00 €", "", "500,00 €", "Una tantum"],
               ["Subtotale", "", "", "", "1.500,00 €", ""],
               ["Sconto 10%", "", "", "", "150,00 €", ""],
               ["Totale", "", "", "", "1.350,00 €", ""]])
    e = extract_and_calc(v, 12)
    assert all(t.status == "Conforme" for t in e.totals), [(t.label, t.status, t.calc) for t in e.totals]


def test_rounding_configurabile():
    v = build([H, ["X", "3", "0,335 €", "", "1,01 €", "Una tantum"]])
    assert extract_and_calc(v, 12, decimals=2, rounding="half_up").lines[0].calc == D("1.01")
    assert extract_and_calc(v, 12, decimals=2, rounding="down").lines[0].calc == D("1.00")
    assert extract_and_calc(v, 12, decimals=3, rounding="half_up").lines[0].calc == D("1.005")


def test_periodi_con_date_in_inglese_righe_unite_e_durata_dedotta():
    """Offerta tipo subscription: righe-intestazione «31 Dec 2026 - 30 Dec 2027» (cella unita) e totale offerta."""
    d = Document()
    t = d.add_table(rows=7, cols=5)
    hdr = ["Part number", "Descrizione", "Q.tà", "Prezzo unitario", "Importo"]
    for j, v in enumerate(hdr):
        t.cell(0, j).text = v
    def period(i, txt):
        c = t.cell(i, 0).merge(t.cell(i, 4)); c.text = txt
    period(1, "31 Dec 2026 - 30 Dec 2027")
    for j, v in enumerate(["2078005", "SolarWinds NPM SLX", "1", "28.692,00", "28.692,00"]):
        t.cell(2, j).text = v
    period(3, "31 Dec 2027 - 30 Dec 2028")
    for j, v in enumerate(["2078005", "SolarWinds NPM SLX", "1", "28.692,00", "28.692,00"]):
        t.cell(4, j).text = v
    for j, v in enumerate(["", "", "", "", ""]):
        pass
    t.cell(5, 1).text = ""
    for j, v in enumerate(["", "Totale Offerta", "", "", "57.384,00"]):
        t.cell(6, j).text = v
    buf = io.BytesIO(); d.save(buf)
    e = extract_and_calc(DocView(open_docx(buf.getvalue())), None)
    assert len(e.lines) == 2 and all(l.periodicity == "contratto" and l.qty == 1 for l in e.lines)
    assert e.span_months == 24
    assert [x[2] for x in e.period_totals] == [D("28692.00"), D("28692.00")]
    tot = [x for x in e.totals if x.label == "Totale Offerta"][0]
    assert tot.status == "Conforme" and tot.calc == D("57384.00"), (tot.status, tot.calc, tot.note)
