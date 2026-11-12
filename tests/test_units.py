from datetime import date
from decimal import Decimal as D

from adreview.config import Settings
from adreview.extract import add_months, find_dates, parse_date
from adreview.llm import suggest_corrections
from adreview.money import fmt_eur, fmt_pct, parse_number, parse_percent, quantize


def test_formati_italiani():
    assert parse_number("1.234,56") == D("1234.56")
    assert parse_number("€ 2.700,00") == D("2700.00")
    assert parse_number("1.500") == D("1500")
    assert parse_number("12,5 %") == D("12.5")
    assert parse_percent("sconto 10%") == D(10) and parse_percent("33,33 %") == D("33.33")
    assert parse_number("abc") is None
    assert fmt_eur(D("2700")) == "2.700,00 €" and fmt_pct(D("33.333")) == "33,33%"


def test_arrotondamento_configurabile():
    assert quantize(D("2.345"), 2, "half_up") == D("2.35")
    assert quantize(D("2.345"), 2, "half_even") == D("2.34")
    assert quantize(D("2.349"), 2, "down") == D("2.34")
    assert quantize(D("2.341"), 2, "up") == D("2.35")


def test_date_italiane():
    assert parse_date("12/11/2026") == date(2026, 11, 12)
    assert parse_date("1° dicembre 2026") == date(2026, 12, 1)
    assert parse_date("31/02/2026") is None
    assert add_months(date(2026, 1, 31), 1) == date(2026, 2, 28)
    assert len(find_dates("dal 01/01/2027 al 31/12/2029")) == 2


def test_llm_valida_i_suggerimenti_e_funziona_con_trasporto_finto():
    st = Settings()
    paras = [(0, "Il cliente potrà effetuare l'attivazione."), (1, "Testo corretto.")]
    fake = lambda payload: '[{"idx":0,"before":"effetuare","after":"effettuare","reason":"refuso"},{"idx":1,"before":"inesistente","after":"x"}]'
    out = suggest_corrections(st, paras, fake)
    assert out == [{"idx": 0, "before": "effetuare", "after": "effettuare", "reason": "refuso"}]


def test_ai_non_configurata_non_invia_nulla(service, mk):
    called = []
    service.llm_transport = lambda p: called.append(p) or "[]"
    s = mk("offerta_template_dati_mancanti.docx", llm_consent=True)
    assert called == [] and not s.llm_consent        # consenso ignorato: servizio non configurato


def test_ai_configurata_con_consenso_propone_correzioni_da_approvare(service, mk):
    service.settings.llm_provider, service.settings.llm_api_key = "anthropic", "test"
    sent = []
    def fake(payload):
        sent.append(payload["text"])
        return '[{"idx": 0, "before": "OFFERTA COMMERCIALE", "after": "Offerta commerciale", "reason": "maiuscole"}]'
    service.llm_transport = fake
    s = mk("offerta_template_dati_mancanti.docx", llm_consent=True)
    assert sent and "2.700,00" not in sent[0] and "SW-XYZ" not in sent[0]          # nessun dato di tabella inviato
    f = [f for f in s.analysis.findings if f.title == "Suggerimento del modello linguistico"][0]
    assert s.analysis.states[s.analysis.findings.index(f)] == "Da approvare"


def test_documenti_atipici_non_bloccano_la_revisione(service):
    from docx import Document
    import io
    d = Document()
    buf = io.BytesIO(); d.save(buf)
    s = service.create_session("vuoto.docx", buf.getvalue(), "demo")
    a = s.analysis
    assert a.client_docx and a.fields["cliente"].value == "Non indicato"
    assert [c for c in a.controls if c["id"] == "calcoli"][0]["status"] in ("Da verificare", "Non verificabile")
    d = Document()
    d.add_paragraph("Offerta n. 1 – Totale € 100,00")
    t = d.add_table(rows=1, cols=1)
    inner = t.cell(0, 0).add_table(rows=1, cols=2)
    inner.cell(0, 0).text = "x"
    buf = io.BytesIO(); d.save(buf)
    s = service.create_session("annidata.docx", buf.getvalue(), "demo")
    assert any("annidate" in f.title for f in s.analysis.findings)
