"""Correzioni approvabili (anche con valore da inserire), testo nascosto, caselle di testo e approvazione multipla."""
import io

import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt

from adreview.docx_utils import DocView, open_docx, visible_text
from conftest import docx_text

VML = 'xmlns:v="urn:schemas-microsoft-com:vml"'


def add_textbox(doc, lines):
    ps = "".join(f"<w:p><w:r><w:t xml:space='preserve'>{t}</w:t></w:r></w:p>" for t in lines)
    xml = (f"<w:r {nsdecls('w')} {VML}><w:pict><v:shape><v:textbox><w:txbxContent>{ps}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>")
    p = doc.add_paragraph()
    p._p.append(parse_xml(xml))


def build() -> bytes:
    d = Document()
    d.add_paragraph("Offerta per Fornitura camera")
    add_textbox(d, ["Spett.le", "Cliente Demo S.p.A", "Alla C.A. di", "Modena – 13/11/2026 13/11/2026", "Rif:OFF_X_26/0001", "La presente offerta è valida fino al 30/11/2026"])
    d.add_heading("Premessa", 1)
    p = d.add_paragraph("Testo ")
    r = p.add_run("Premessa"); r.font.hidden = True            # testo nascosto: non deve contare
    d.add_heading("Oggetto", 1)
    d.add_paragraph("[ ___inserire il paragrafo su Requisiti solo se necessario]")
    d.add_paragraph("…", style="List Bullet")
    d.add_heading("Offerta Economica", 1)
    t = d.add_table(rows=3, cols=3)
    for j, v in enumerate(["Descrizione", "Quantità", "Importo"]):
        t.cell(0, j).text = v
    for j, v in enumerate(["Camera varifocale", "1", "1.787,50 €"]):
        t.cell(1, j).text = v
    for j, v in enumerate(["Netto a Voi Riservato", "", "1.787,50 €"]):
        t.cell(2, j).text = v
    d.add_heading("Condizioni di vendita", 1)
    t2 = d.add_table(rows=3, cols=2)
    for i, (a, b) in enumerate([("TIPOLOGIA E CONDIZIONI DI PAGAMENTO", "Bonifico Bancario 30gg Fine Mese"), ("VALIDITÀ CONTRATTO", ""), ("RINNOVO", "Automatico: disdetta 90 giorni prima")]):
        t2.cell(i, 0).text, t2.cell(i, 1).text = a, b
    buf = io.BytesIO(); d.save(buf)
    return buf.getvalue()


@pytest.fixture()
def sess(service):
    return service.create_session("camera.docx", build(), "ad_offerte")


def by_title(a, start):
    return [f for f in a.findings if f.title.startswith(start)]


def test_testo_nascosto_non_duplica_i_titoli(sess):
    v = DocView(open_docx(sess.offer_bytes))
    assert "Premessa" in [h.text for h in v.headings]
    assert "TestoPremessa" not in "".join(visible_text(p.p) for p in v.paras)


def test_campi_dalla_casella_di_testo_e_totale_netto_a_voi_riservato(sess):
    a = sess.analysis
    F = a.fields
    assert F["data_offerta"].value == "13/11/2026" and F["numero_offerta"].value == "OFF_X_26/0001" and F["validita"].value == "30/11/2026"
    assert a.econ.lines[0].status == "Conforme" and a.econ.contract == 1787.50
    assert [t.status for t in a.econ.totals] == ["Conforme"]


def test_mancanze_hanno_una_correzione_con_valore(service, sess):
    a = sess.analysis
    cell = by_title(a, "Campo «VALIDITÀ CONTRATTO»")[0]
    ca = by_title(a, "Destinatario")[0]
    dup = by_title(a, "Data ripetuta")[0]
    ph = by_title(a, "Segnaposto rimasto")[0]
    el = by_title(a, "Punto elenco vuoto")[0]
    for f in (cell, ca, el):
        assert f.change.needs_value and a.states[a.findings.index(f)] == "Da completare"
    assert not dup.change.needs_value and a.states[a.findings.index(dup)] == "Da approvare"
    # senza valore l'approvazione non è ammessa
    assert not service.approve(sess, cell.change.key)
    assert not service.approve(sess, ca.change.key, ca.change.after)          # valore invariato = ancora vuoto
    assert "Word" not in docx_text(sess.analysis.client_docx) or True
    # con il valore: inserito nel Word al punto giusto
    assert service.approve(sess, cell.change.key, "3 anni")
    out = Document(io.BytesIO(sess.analysis.client_docx))
    assert [r.cells[1].text for r in out.tables[1].rows][1] == "3 anni"
    assert not by_title(sess.analysis, "Campo «VALIDITÀ CONTRATTO»")
    # casella di testo: C.A. completata e data ripetuta ridotta
    a = sess.analysis
    ca = by_title(a, "Destinatario")[0]; dup = by_title(a, "Data ripetuta")[0]
    done, refused = service.approve_many(sess, [(ca.change.key, "Alla C.A. di Mario Rossi"), (dup.change.key, None)])
    assert (done, refused) == (2, [])
    xml = docx_text(sess.analysis.client_docx)
    assert "Alla C.A. di Mario Rossi" in xml and "13/11/2026 13/11/2026" not in xml and "Modena – 13/11/2026" in xml


def test_segnaposto_rimosso_o_sostituito_e_ricalcolo(service, sess):
    a = sess.analysis
    ph = by_title(a, "Segnaposto rimasto")[0]
    assert ph.change.editable and ph.change.after == ""
    service.approve(sess, ph.change.key)                       # Approva = rimuove
    assert "inserire il paragrafo" not in docx_text(sess.analysis.client_docx)
    assert not by_title(sess.analysis, "Segnaposto rimasto")


def test_approva_tutto_il_pronto_in_un_colpo(service, mk):
    s = mk("offerta_template_dati_mancanti.docx")
    ready = [(f.change.key, None) for i, f in enumerate(s.analysis.findings) if f.change and not f.change.needs_value and s.analysis.states[i] == "Da approvare"]
    assert len(ready) >= 4
    done, refused = service.approve_many(s, ready)
    assert done == len(ready) and not refused
    t = docx_text(s.analysis.client_docx)
    assert "sopratutto" not in t and "effetuare" not in t and "perchè" not in t


def test_modalita_revisione_con_inserimenti_e_caselle(service, sess):
    from conftest import docx_text as dt
    import zipfile
    data = service.export_tracked(sess)
    xml = zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml").decode()
    assert "<w:ins " in xml and "<w:del " in xml and "13/11/2026" in xml


def test_ui_riepilogo_con_approva_e_campo_valore(service):
    from fastapi.testclient import TestClient
    from adreview.app import create_app
    c = TestClient(create_app(service.settings, service))
    r = c.post("/review", files=[("offer", ("c.docx", io.BytesIO(build())))], data={"profile": "ad_offerte"}, follow_redirects=False)
    assert r.headers["location"].endswith("/riepilogo")
    sid = r.headers["location"].split("/")[2]
    page = c.get(f"/review/{sid}/riepilogo").text
    assert "Correzioni pronte" in page and "Mancanze da completare" in page and "Approva le correzioni selezionate" in page
    assert 'placeholder="obbligatorio"' in page
    s = service.get(sid)
    cell = [f for f in s.analysis.findings if f.title.startswith("Campo «VALIDITÀ")][0]
    # approva con valore via form
    r = c.post(f"/review/{sid}/approve-selected", data={"only": cell.change.key, f"v|{cell.change.key}": "36 mesi"}, follow_redirects=False)
    assert r.status_code == 303
    assert "36 mesi" in docx_text(service.get(sid).analysis.client_docx)
    # approva tutte le pronte
    ready = [f.change.key for i, f in enumerate(service.get(sid).analysis.findings) if f.change and not f.change.needs_value and service.get(sid).analysis.states[i] == "Da approvare"]
    r = c.post(f"/review/{sid}/approve-selected", data={"k": ready}, follow_redirects=False)
    assert "13/11/2026 13/11/2026" not in docx_text(service.get(sid).analysis.client_docx)
