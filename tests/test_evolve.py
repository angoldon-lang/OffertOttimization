"""«Evolvi offerta»: da offerta tradizionale a offerta estesa (Oggetto a elenco, Soluzione proposta, tabella con prezzi unitari, condizioni)."""
import io
import re
import zipfile
from decimal import Decimal as D

import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from lxml import etree

from adreview import evolve as ev
from adreview.docx_utils import DocView, open_docx, visible_text
from conftest import docx_text


def bullet(doc, text):
    p = doc.add_paragraph(text, style="Body Text")
    ppr = p._p.get_or_add_pPr()
    ppr.append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'))
    return p


def traditional() -> bytes:
    d = Document()
    d.add_paragraph("Offerta per Fornitura Palo Alto")
    d.add_heading("Oggetto", 1)
    d.add_paragraph("Facendo seguito ai colloqui avvenuti, abbiamo il piacere di sottoporre la nostra offerta relativa a:", style="Body Text")
    bullet(d, "Fornitura Palo Alto")
    bullet(d, "Installazione e configurazione")
    d.add_paragraph().add_run().add_break(__import__("docx").enum.text.WD_BREAK.PAGE)
    d.add_heading("Offerta Economica", 1)
    t = d.add_table(rows=7, cols=3)
    t.style = "Table Grid"
    data = [("Descrizione", "Quantità", "Importo"),
            ("PAN-PWR-CORD-EU\nPower cord for Continental Europe", "2", "200,00 €"),
            ("PAN-PA-560\nPalo Alto Networks PA-560", "2", "11.300,00 €"),
            ("PAN-PA-560-BND PRECISIONAI-3YR\nPA-560, Precision AI Bundle, 36 months", "", "23.400,00 €"),
            ("PAN-SVC-BKLN-560-3YR\nPA-560, Partner enabled premium support, 36 months", "", "4.600,00 €"),
            ("Installazione e configurazione", "", "4.000,00 €"),
            ("Netto a Voi Riservato", "", "")]
    for i, row in enumerate(data):
        for j, v in enumerate(row):
            t.cell(i, j).text = v
    d.add_paragraph("Prezzi Iva esclusa ", style="Body Text")
    d.add_heading("Condizioni di vendita", 1)
    c = d.add_table(rows=4, cols=2)
    for i, (a, b) in enumerate([("TIPOLOGIA E CONDIZIONI DI PAGAMENTO", ""), ("FATTURAZIONE", "All'ordine"), ("VALIDITÀ CONTRATTO", ""), ("RINNOVO", "Automatico")]):
        c.cell(i, 0).text, c.cell(i, 1).text = a, b
    buf = io.BytesIO(); d.save(buf)
    return buf.getvalue()


def accept_reject(data, mode):
    z = zipfile.ZipFile(io.BytesIO(data))
    root = etree.fromstring(z.read("word/document.xml"))
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    drop = W + ("del" if mode == "accept" else "ins")
    for el in list(root.iter(drop)):
        if el.getparent() is not None and el.getparent().tag != W + "rPr" and el.getparent().tag != W + "trPr":
            el.getparent().remove(el)
    # righe di tabella marcate come inserite (reject) o cancellate (accept): scartale
    for tr in list(root.iter(W + "tr")):
        trpr = tr.find(W + "trPr")
        if trpr is not None and trpr.find(W + ("del" if mode == "accept" else "ins")) is not None:
            tr.getparent().remove(tr)
    out = []
    for p in root.iter(W + "p"):
        ppr = p.find(W + "pPr")
        mark = ppr.find(W + "rPr") if ppr is not None else None
        if mark is not None and mark.find(W + ("del" if mode == "accept" else "ins")) is not None:
            continue            # il paragrafo stesso era inserito/cancellato
        out.append("".join(t.text or "" for t in p.iter(W + "t", W + "delText")))
    return [x for x in out if x.strip()]


@pytest.fixture()
def sess(service):
    return service.create_session("tradizionale.docx", traditional(), "ad_offerte")


def test_valori_proposti_dalla_tradizionale(service, sess):
    inp = service.evolve_form(sess)
    rows = {r.title: r for r in inp.econ_rows}
    assert rows["PAN-PA-560-BND PRECISIONAI-3YR"].qty == 2 and rows["PAN-PA-560-BND PRECISIONAI-3YR"].unit == D("11700.00")   # q.tà mancante dedotta dalle altre righe
    assert rows["Installazione e configurazione"].qty == 1 and rows["Installazione e configurazione"].unit == D("4000.00")
    assert inp.total == D("43500.00")
    assert inp.solution_title == "Soluzione proposta: Palo Alto Networks PA-560"
    assert "36 mesi" in inp.econ_note
    assert inp.conditions["VALIDITÀ CONTRATTO"] == "36 mesi dalla data di attivazione"
    assert "TIPOLOGIA E CONDIZIONI DI PAGAMENTO" in inp.conditions and "RINNOVO" not in inp.conditions
    assert any("Installazione" in l for l in inp.oggetto)
    assert "## Supporto e servizi professionali" in inp.solution_markup


def test_markup():
    b = ev.parse_markup("Intro riga 1\nriga 2\n\n## Titolo\n- Termine – descrizione\n- solo testo\n| A | B |\n|---|---|\n| x | y |\n> nota")
    assert b[0] == ("p", "Intro riga 1 riga 2") and b[1] == ("h", "Titolo")
    assert b[2] == ("b", "Termine", "descrizione") and b[3] == ("b", "", "solo testo")
    assert b[4] == ("t", [["A", "B"], ["x", "y"]]) and b[5] == ("n", "nota")


def test_normalizzazione_righe():
    r = ev.normalize_rows([ev.EconRow("a", "", D(2), D("5650"), None), ev.EconRow("b", "", D(2), None, D("4600")), ev.EconRow("c", "", None, None, D("100")), ev.EconRow("", "", None, None, None)])
    assert [(x.qty, x.unit, x.amount) for x in r] == [(2, D("5650"), D("11300")), (2, D("2300"), D("4600")), (1, D("100"), D("100"))]


def test_applica_evoluzione(service, sess):
    inp = service.evolve_form(sess)
    inp.oggetto = ["Fornitura di n. 2 firewall Palo Alto Networks PA-560", "Installazione e configurazione in alta affidabilità (HA)", "Supporto sistemistico annuale"]
    inp.solution_markup = ("La soluzione proposta prevede la fornitura di due firewall.\n\n## Caratteristiche principali\n- App-ID – identificazione applicazioni\n\n"
                           "## Specifiche tecniche\n| Parametro | Valore |\n| Firewall throughput | 8,5 Gbps |\n| Storage | SSD 240 GB |")
    inp.econ_rows.append(ev.EconRow("Supporto sistemistico annuale", "Servizio AD Consulting", D(1), D("2000"), None))
    inp.conditions["TIPOLOGIA E CONDIZIONI DI PAGAMENTO"] = "Bonifico Bancario 30gg Fine Mese"
    inp.econ_rows = ev.normalize_rows(inp.econ_rows)
    rep = service.evolve_apply(sess, inp)
    assert rep["oggetto"] and rep["soluzione"] and rep["tabella"] and rep["nota"] and len(rep["condizioni"]) == 2
    doc = Document(io.BytesIO(sess.offer_bytes))
    texts = [visible_text(p) for p in doc.paragraphs]
    i_ogg, i_sol, i_eco = texts.index("Oggetto"), texts.index("Soluzione proposta: Palo Alto Networks PA-560"), texts.index("Offerta Economica")
    assert i_ogg < i_sol < i_eco
    assert "Fornitura Palo Alto" not in texts and "Fornitura di n. 2 firewall Palo Alto Networks PA-560" in texts
    assert "Prezzi IVA esclusa. Le sottoscrizioni e il supporto hanno durata di 36 mesi dalla data di attivazione." in texts
    assert not any(t.startswith("Prezzi Iva esclusa") for t in texts)
    # tabelle: specifiche (2 col) + economica (4 col) con prezzi unitari e totale 45.500
    cols = [len(t.columns) for t in doc.tables]
    assert cols[:2] == [2, 4]
    econ = doc.tables[1]
    assert [c.text for c in econ.rows[0].cells] == ["Descrizione", "Q.tà", "Prezzo unitario", "Importo"]
    assert econ.rows[-1].cells[-1].text == "45.500,00 €"
    cond = doc.tables[2]
    assert cond.rows[0].cells[1].text.startswith("Bonifico") and cond.rows[2].cells[1].text == "36 mesi dalla data di attivazione"
    # i controlli sul documento evoluto: totali coerenti, nessun errore di calcolo
    a = sess.analysis
    assert sess.evolved and a.econ.months == 36
    assert [t.status for t in a.econ.totals] == ["Conforme"] and a.econ.totals[0].calc == D("45500.00")     # somma degli importi di riga
    assert a.econ.contract == D("49500.00")                                                                     # canone annuale ×3 anni nel valore triennale
    assert not [f for f in a.findings if f.check == "calcoli" and f.status == "Errore"]
    assert a.verify_problems == []
    # aggiornamento indice all'apertura
    assert b"updateFields" in zipfile.ZipFile(io.BytesIO(sess.offer_bytes)).read("word/settings.xml")


def test_modifiche_tracciate_accetta_e_rifiuta(service, sess):
    inp = service.evolve_form(sess)
    inp.econ_rows.append(ev.EconRow("Supporto sistemistico annuale", "", D(1), D("2000"), None))
    inp.econ_rows = ev.normalize_rows(inp.econ_rows)
    sess.evolve_inputs = inp
    data = service.export_evolution_tracked(sess)
    x = zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml").decode()
    assert "<w:ins " in x and "<w:del " in x and "w:trPr><w:cantSplit/><w:tblHeader/><w:ins" in x.replace("\n", "").replace("  ", "")
    original = [visible_text(p) for p in Document(io.BytesIO(traditional())).paragraphs if visible_text(p).strip()]
    rej = accept_reject(data, "reject")
    for t in original:
        assert t.strip() in [r.strip() for r in rej] or t.strip() == "", t       # rifiutando tutto torna il testo originale
    accl = accept_reject(data, "accept")
    acc = "\n".join(accl)
    assert "Soluzione proposta" in acc and "45.500,00 €" in acc and "Fornitura Palo Alto" not in accl
    assert "Supporto sistemistico annuale" in acc
    # il file si apre
    Document(io.BytesIO(data))


def test_flusso_web_evolvi(service):
    from fastapi.testclient import TestClient
    from adreview.app import create_app
    c = TestClient(create_app(service.settings, service))
    r = c.post("/review", files=[("offer", ("t.docx", io.BytesIO(traditional())))], data={"profile": "ad_offerte"}, follow_redirects=False)
    sid = r.headers["location"].split("/")[2]
    page = c.get(f"/review/{sid}/evolvi").text
    assert "Evolvi offerta" in page and "Soluzione proposta: Palo Alto Networks PA-560" in page and "43.500,00" in page and "nuova riga" in page
    form = {"oggetto": "Fornitura firewall\nInstallazione", "sol_title": "Soluzione proposta: PA-560", "sol_markup": "Testo della soluzione.",
            "econ_note": "Prezzi IVA esclusa.", "remove_refs": "on", "cond|TIPOLOGIA E CONDIZIONI DI PAGAMENTO": "Bonifico 30gg", "cond|VALIDITÀ CONTRATTO": "36 mesi"}
    rows = [("PAN-PA-560", "Firewall", "2", "5.650,00", ""), ("Supporto", "Annuale", "1", "", "2.000,00")]
    for i, (t, d, q, u, a) in enumerate(rows):
        form.update({f"r|{i}|title": t, f"r|{i}|desc": d, f"r|{i}|qty": q, f"r|{i}|unit": u, f"r|{i}|amount": a})
    assert c.post(f"/review/{sid}/evolvi", data={**form, "action": "preview"}, follow_redirects=False).status_code == 303
    assert "13.300,00" in c.get(f"/review/{sid}/evolvi").text                       # 11.300 + 2.000, ricalcolato in anteprima
    assert c.post(f"/review/{sid}/evolvi", data={**form, "action": "apply"}, follow_redirects=False).status_code == 303
    s = service.get(sid)
    assert s.evolved and "13.300,00 €" in docx_text(s.analysis.client_docx)
    d = c.get(f"/review/{sid}/download/evoluzione.docx")
    assert d.status_code == 200 and b"<w:ins " in zipfile.ZipFile(io.BytesIO(d.content)).read("word/document.xml")
    # torna all'originale
    c.post(f"/review/{sid}/evolvi", data={"action": "revert"})
    assert not service.get(sid).evolved and "Fornitura Palo Alto" in docx_text(service.get(sid).analysis.client_docx)


def test_bozza_ai_non_invia_prezzi(service, sess):
    service.settings.llm_provider, service.settings.llm_api_key = "anthropic", "k"
    sent = []
    service.llm_transport = lambda p: sent.append(p["text"]) or "Intro AI.\n\n## Caratteristiche principali\n- Voce – testo"
    sess.llm_consent = True
    inp = service.evolve_form(sess)
    out = service.evolve_draft_ai(sess, inp)
    assert out.startswith("Intro AI") and sent
    assert "5.650" not in sent[0] and "11300" not in sent[0] and "200" not in sent[0].replace("PAN-PWR-CORD-EU", "")
    sess.llm_consent = False
    assert service.evolve_draft_ai(sess, inp) is None
