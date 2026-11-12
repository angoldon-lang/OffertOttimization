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
