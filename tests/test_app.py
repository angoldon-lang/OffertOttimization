import io
import re

import pytest
from fastapi.testclient import TestClient

from adreview.app import create_app
from adreview.config import Settings

from conftest import DEMO, docx_text


@pytest.fixture()
def client():
    st = Settings.from_env()
    st.llm_provider = "none"
    return TestClient(create_app(st))


def f(name):
    return (DEMO / name).open("rb")


def upload(client, offer, template=None, quotes=(), **data):
    files = [("offer", (offer, f(offer)))]
    if template:
        files.append(("template", (template, f(template))))
    files += [("quotes", (q, f(q))) for q in quotes]
    r = client.post("/review", files=files, data={"profile": "demo", **data}, follow_redirects=False)
    assert r.status_code == 303, r.text[:300]
    return r.headers["location"].split("/")[2]


def test_flusso_completo_revisione_approvazione_export(client):
    sid = upload(client, "offerta_totale_errato.docx", quotes=["quotazione_fornitore_triennale.xlsx"],
                 ic_desc=["Installazione"], ic_qty=["1"], ic_unit=["500"], ic_per=["una_tantum"])
    for tab in ("controlli", "calcoli", "template", "margine", "riepilogo", "export"):
        r = client.get(f"/review/{sid}/{tab}")
        assert r.status_code == 200
    page = client.get(f"/review/{sid}/controlli").text
    assert "Totale errato" in page and "2.800,00" in page and "Approva e applica al Word" in page
    key = re.search(r"/change/([0-9a-f]{12})/approve", page).group(1)
    # prima dell'approvazione il DOCX cliente ha ancora il vecchio importo
    d0 = client.get(f"/review/{sid}/download/cliente.docx")
    assert d0.status_code == 200 and "2.800,00" in docx_text(d0.content)
    assert client.post(f"/review/{sid}/change/{key}/approve?next=controlli", follow_redirects=False).status_code == 303
    d1 = client.get(f"/review/{sid}/download/cliente.docx")
    assert "2.700,00" in docx_text(d1.content)
    # report + xlsx
    rep = client.get(f"/review/{sid}/download/report.docx")
    assert rep.status_code == 200 and "RISERVATO" in docx_text(rep.content)
    x = client.get(f"/review/{sid}/download/margine.xlsx")
    assert x.status_code == 200 and x.content[:2] == b"PK"
    assert "Installazione" not in docx_text(d1.content)
    # cancellazione dati
    assert client.post(f"/review/{sid}/delete", follow_redirects=False).status_code == 303
    assert client.get(f"/review/{sid}/controlli").status_code == 404


def test_senza_preventivo_margine_non_disponibile_ma_resto_si(client):
    sid = upload(client, "offerta_triennale_conforme.docx")
    assert client.get(f"/review/{sid}/download/margine.xlsx").status_code == 409
    assert client.get(f"/review/{sid}/download/cliente.docx").status_code == 200
    assert client.get(f"/review/{sid}/download/report.docx").status_code == 200
    assert "Non calcolabile" in client.get(f"/review/{sid}/riepilogo").text


def test_conferma_periodicita_e_modifica_proposta_da_tabella(client):
    sid = upload(client, "offerta_triennale_conforme.docx")
    # l'utente propone di cambiare la quantità da 10 a 12: resta una proposta finché non approvata
    r = client.post(f"/review/{sid}/lines", data={"mesi": "", "e|T1R2|qty": "12", "c|T1R2|periodicita": "annuale"}, follow_redirects=False)
    assert r.status_code == 303
    page = client.get(f"/review/{sid}/calcoli").text
    assert "Modifiche in attesa" in page and "Quantità" in page
    assert "<ins>12</ins>" in page
    assert ">10 <" in docx_text(client.get(f"/review/{sid}/download/cliente.docx").content) or "10" in docx_text(client.get(f"/review/{sid}/download/cliente.docx").content)
    key = re.search(r"/change/([0-9a-f]{12})/approve\?next=calcoli", page).group(1)
    client.post(f"/review/{sid}/change/{key}/approve?next=calcoli")
    page = client.get(f"/review/{sid}/calcoli").text
    # ricalcolo dopo l'approvazione: 12 × 100 × 0,9 = 1.080 ≠ importo riportato 900 → nuovo errore proposto
    assert "1.080,00" in page


def test_upload_non_valido(client):
    r = client.post("/review", files=[("offer", ("x.docx", io.BytesIO(b"non un docx")))], data={"profile": "demo"})
    assert r.status_code == 200 and "DOCX valido" in r.text
    r = client.post("/review", data={"profile": "demo"})
    assert r.status_code == 200 and "Seleziona" in r.text


def test_margine_ui_abbinamento_e_ripartizione(client):
    sid = upload(client, "offerta_due_voci.docx", quotes=["quotazione_costo_condiviso.xlsx"])
    page = client.get(f"/review/{sid}/margine").text
    assert "PARZIALE" in page
    cid = "Q1-2"
    client.post(f"/review/{sid}/margin/alloc/{cid}", data={"mode": "proportional", "targets": ["T1R2", "T1R3"]})
    page = client.get(f"/review/{sid}/margine").text
    assert "COMPLETO" in page or "completo" in page
    assert "2.700,00" in page
