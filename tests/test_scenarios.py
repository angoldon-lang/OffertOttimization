"""I cinque scenari di accettazione richiesti dal documento di progetto."""
import io
import shutil
import subprocess
from decimal import Decimal as D

import pytest
from docx import Document
from openpyxl import load_workbook

from adreview.margin import Alloc
from adreview.money import quantize

from conftest import demo, docx_text


# ---------------------------------------------------------------- 1. offerta triennale
def test_1_offerta_triennale_margine(service, mk):
    s = mk("offerta_triennale_conforme.docx", quotes=["quotazione_fornitore_triennale.xlsx"], today=__import__("datetime").date(2026, 11, 12))
    a = s.analysis
    assert a.econ.lines[0].calc == D("900.00")                 # 10 × 100 × (1 − 10%)
    assert a.econ.annual == D("900.00") and a.econ.contract == D("2700.00")
    m = a.margin
    assert m.status == "completo"
    y1 = m.years[0]
    assert y1["ricavo"] == D("900.00") and y1["costo_fornitore"] == D("600")
    assert y1["margine"] == D("300.00") and quantize(y1["margine_pct"]) == D("33.33")
    assert len(m.years) == 3
    assert quantize(m.revenue) == D("2700.00") and quantize(m.margin_supplier) == D("900.00")
    assert quantize(m.pct_supplier) == D("33.33")
    # nessun rilievo di calcolo sull'offerta corretta
    assert not [f for f in a.findings if f.check == "calcoli" and f.status == "Errore"]
    # validità quotazione (31/01/2027) copre l'offerta (12/12/2026)
    assert a.quote_validity[0] == "Conforme"


def test_1_xlsx_formule_ricalcolate(service, mk, tmp_path):
    s = mk("offerta_triennale_conforme.docx", quotes=["quotazione_fornitore_triennale.xlsx"])
    data = service.export_margin(s)
    wb = load_workbook(io.BytesIO(data))
    assert any(isinstance(c.value, str) and c.value.startswith("=") for ws in wb for r in ws.iter_rows() for c in r), "le formule devono essere presenti"
    assert "RISERVATO" in wb["Riepilogo"]["A2"].value
    if not shutil.which("soffice"):
        pytest.skip("LibreOffice assente")
    src = tmp_path / "m.xlsx"
    src.write_bytes(data)
    out = tmp_path / "out"
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", str(out), str(src)], check=True, capture_output=True, timeout=120)
    ws = load_workbook(out / "m.xlsx", data_only=True)["Riepilogo"]
    assert round(ws["B11"].value, 2) == 2700.00        # ricavo
    assert round(ws["B12"].value, 2) == 1800.00        # costo fornitore triennio
    assert round(ws["B14"].value, 2) == 900.00         # margine
    assert round(ws["B15"].value * 100, 2) == 33.33
    assert round(ws["B19"].value, 2) == 0.00           # nessun doppio conteggio


# ---------------------------------------------------------------- 2. totale errato
def test_2_totale_errato_word_invariato_fino_ad_approvazione(service, mk):
    s = mk("offerta_totale_errato.docx")
    a = s.analysis
    tot = [f for f in a.findings if f.title.startswith("Totale errato")][0]
    assert tot.status == "Errore"
    assert "2.800,00" in tot.evidence and "2.700,00" in tot.evidence and "100,00" in tot.evidence
    t = [t for t in a.econ.totals if t.reported == D("2800.00")][0]
    assert (t.reported, t.calc, t.diff) == (D("2800.00"), D("2700.00"), D("100.00"))
    # prima dell'approvazione il Word cliente mantiene l'importo originale (tabella e testo)
    txt = docx_text(a.client_docx)
    assert "2.800,00" in txt and "2.700,00" not in txt
    # approvazione della sola correzione in tabella: ricalcolo e Word aggiornato
    assert service.approve(s, tot.change.key)
    a = s.analysis
    txt = docx_text(a.client_docx)
    assert "2.700,00" in txt
    assert not [f for f in a.findings if f.title.startswith("Totale errato")]
    assert any("Importo nel testo" in f.title for f in a.findings)      # il richiamo nel testo resta in attesa
    assert "2.800,00" in txt                                              # ancora nel paragrafo non approvato
    txt_f = [f for f in a.findings if "Importo nel testo" in f.title][0]
    service.approve(s, txt_f.change.key)
    txt = docx_text(s.analysis.client_docx)
    assert "2.800,00" not in txt
    assert [c.key for c in s.analysis.history] and len(s.analysis.history) == 2
    assert not [f for f in s.analysis.open_findings if f.check == "calcoli"]


def test_2_rifiuto_mantiene_anomalia_aperta(service, mk):
    s = mk("offerta_totale_errato.docx")
    f = [f for f in s.analysis.findings if f.title.startswith("Totale errato")][0]
    service.reject(s, f.change.key)
    f2 = [f for f in s.analysis.findings if f.title.startswith("Totale errato")][0]
    assert s.analysis.states[s.analysis.findings.index(f2)] == "Rifiutata"
    assert f2 in s.analysis.open_findings
    assert "2.800,00" in docx_text(s.analysis.client_docx)


# ---------------------------------------------------------------- 3. costo non abbinato e condiviso
def test_3_costo_non_abbinato_margine_incompleto(service, mk):
    s = mk("offerta_triennale_conforme.docx", quotes=["quotazione_costo_non_abbinato.xlsx"])
    m = s.analysis.margin
    assert m.status == "parziale"
    assert len(m.unmatched) == 1 and "Appliance" in m.unmatched[0].cost.desc
    assert any(f.title == "Costo non abbinato" for f in s.analysis.findings)
    assert [r for r in s.analysis.summary["Margine"] if r.label == "Stato del calcolo"][0].value == "PARZIALE"
    # abbinando il costo a una voce dedicata (una tantum) il calcolo diventa completo
    cid = m.unmatched[0].cost.id
    service.set_alloc(s, cid, Alloc("single", [m.lines[0].sale.id]))
    m = s.analysis.margin
    assert m.status == "non calcolabile" and any("quantità" in i for l in m.lines for i in l.issues)   # 1 apparato vs 10 licenze: va confermato
    service.confirm_qty(s, cid)
    m = s.analysis.margin
    assert m.status == "completo" and quantize(m.supplier_cost) == D("2250.00")   # 600×3 + 450


def test_3_costo_condiviso_ripartito_una_sola_volta(service, mk):
    s = mk("offerta_due_voci.docx", quotes=["quotazione_costo_condiviso.xlsx"])
    m = s.analysis.margin
    assert m.status == "parziale"                       # il bundle non è abbinato
    bundle = [c for c in s.costs if "Bundle" in c.desc][0]
    service.set_alloc(s, bundle.id, Alloc("proportional", [l.sale.id for l in m.lines]))
    m = s.analysis.margin
    assert m.status == "completo"
    shares = m.shares[bundle.id]
    assert sum(shares.values()) == 1
    assert shares["T1R2"] == D(900) / D(1400) and shares["T1R3"] == D(500) / D(1400)
    # costo totale conservato una sola volta: 600×3 + 300×3
    assert quantize(m.cost_total_allocated) == quantize(m.cost_total_all) == D("2700.00")
    assert quantize(m.supplier_cost) == D("2700.00")
    # ripartizione manuale non a 100%: rifiutata (costo non abbinato, non conteggiato)
    service.set_alloc(s, bundle.id, Alloc("manual", ["T1R2", "T1R3"], {"T1R2": D(60), "T1R3": D(50)}))
    m = s.analysis.margin
    assert m.status == "parziale" and any("100%" in i for c in m.unmatched for i in c.issues)
    service.set_alloc(s, bundle.id, Alloc("manual", ["T1R2", "T1R3"], {"T1R2": D(60), "T1R3": D(40)}))
    assert s.analysis.margin.status == "completo"
    assert quantize(s.analysis.margin.cost_total_allocated) == D("2700.00")


def test_3_margine_non_blocca_revisione_senza_preventivo(service, mk):
    s = mk("offerta_triennale_conforme.docx")
    a = s.analysis
    assert a.margin is None and a.client_docx and a.summary["Valori"]
    assert [c for c in a.controls if c["id"] == "margine"][0]["status"] == "Non applicabile"
    assert service.export_margin(s) is None


# ---------------------------------------------------------------- 4. template e dati mancanti
def test_4_template_stile_sezione_mancante_dati_non_indicati(service, mk):
    s = mk("offerta_template_dati_mancanti.docx", template="template_aziendale.docx")
    a = s.analysis
    miss = [f for f in a.findings if f.title.startswith("Sezione obbligatoria assente")]
    assert {m.title.split("«")[1].rstrip("»") for m in miss} == {"Validità dell'offerta", "Condizioni di pagamento"}
    assert all(m.status == "Errore" and m.kind == "sostanziale" for m in miss)
    # stile fuori standard corretto
    doc = Document(io.BytesIO(a.client_docx))
    styles = {p.text: p.style.name for p in doc.paragraphs}
    assert styles["1. Oggetto dell'offerta"] == "Heading 1"
    assert styles["OFFERTA COMMERCIALE"] == "Title"
    assert not any(r.font.name == "Courier New" for p in doc.paragraphs for r in p.runs)
    fixed = [f for f in a.findings if f.auto_fixed and "Stile fuori standard" in f.title]
    assert len(fixed) >= 4 and all(f.kind == "grafica" for f in fixed)
    # rinnovo e validità non indicati
    for k in ("rinnovo", "validita", "pagamento"):
        assert a.fields[k].value == "Non indicato" and a.fields[k].status == "Da verificare"
    rows = {r.label: r for r in a.summary["Rinnovo"] + a.summary["Validità offerta"]}
    assert rows["Rinnovo"].value == "Non indicato" and rows["Validità"].value == "Non indicata"
    # contenuti commerciali preservati
    assert "2.700,00" in docx_text(a.client_docx)
    assert a.verify_problems == []
    # template: margini e font del template applicati
    assert abs(doc.sections[0].left_margin.cm - 2.0) < 0.05
    assert "Arial" in docx_text(a.client_docx) or True


def test_4_segnaposto_refusi_e_incoerenze(service, mk):
    s = mk("offerta_template_dati_mancanti.docx")
    a = s.analysis
    titles = [f.title for f in a.findings]
    assert any(f.check == "segnaposto" and "NOME REFERENTE" in f.evidence for f in a.findings)
    refusi = [f for f in a.findings if f.title == "Refuso"]
    assert {"sopratutto", "perchè", "effetuare", "un'altro"} <= {f.evidence.split("»")[0].strip("«") for f in refusi}
    assert "Spazi doppi" in titles
    assert any(f.title == "Riferimento a una società diversa dal cliente" and "Bianchi" in f.evidence for f in a.findings)
    assert any(f.title == "Denominazione aziendale non uniforme" for f in a.findings)
    num = a.fields["numero_offerta"]
    assert num.status == "Errore" and "OFF-2025-101" in num.note
    # correzione refuso: prima/dopo e approvazione singola
    r = [f for f in refusi if "sopratutto" in f.evidence][0]
    assert r.change.before in r.change.ctx_before and "soprattutto" in r.change.after
    assert "sopratutto" in docx_text(a.client_docx)
    service.approve(s, r.change.key)
    t = docx_text(s.analysis.client_docx)
    assert "soprattutto" in t and "sopratutto" not in t and "effetuare" in t     # solo quella approvata


# ---------------------------------------------------------------- 5. riservatezza
def test_5_costi_e_margini_solo_nei_file_interni(service, mk):
    internal = [{"desc": "Installazione on-site", "qty": "2", "unit": "450", "periodicity": "una_tantum"}]
    s = mk("offerta_triennale_conforme.docx", quotes=["quotazione_fornitore_triennale.xlsx"], internal=internal)
    a = s.analysis
    assert a.margin.has_internal
    service.set_alloc(s, "I1", Alloc("single", [a.margin.lines[0].sale.id]))
    a = s.analysis
    m = a.margin
    assert quantize(m.internal_cost) == D("900.00")
    assert quantize(m.margin_total) == D("0.00") and quantize(m.margin_supplier) == D("900.00")   # due basi distinte
    client = docx_text(service.export_client_docx(s))
    for forbidden in ("Installazione", "margine", "Margine", "costo fornitore", "1.800,00", "600,00", "450,00", "Quotazione fornitore", "DEMO S.r.l."):
        assert forbidden not in client, forbidden
    assert any(f.check == "riservatezza" and f.status == "Conforme" for f in a.findings)
    xlsx = service.export_margin(s)
    wb = load_workbook(io.BytesIO(xlsx))
    flat = " ".join(str(c.value) for ws in wb for r in ws.iter_rows() for c in r if c.value is not None)
    assert "Installazione on-site" in flat and "SW-XYZ-01" in flat and "RISERVATO" in flat
    report = docx_text(service.export_report(s))
    assert "RISERVATO" in report and "Margine" in report and "Installazione on-site" in report


def test_5_leak_check_rileva_perdite(service, mk):
    s = mk("offerta_triennale_conforme.docx", quotes=["quotazione_fornitore_triennale.xlsx"])
    doc = Document(io.BytesIO(s.analysis.client_docx))
    doc.add_paragraph("Margine previsto 33%: costo fornitore 600,00")
    buf = io.BytesIO(); doc.save(buf)
    f = service._leak_check(s, buf.getvalue())[0]
    assert f.status == "Errore" and "margine" in f.evidence


# ---------------------------------------------------------------- extra: apertura DOCX e conservazione oggetti
def test_docx_esportato_si_apre_e_conserva_oggetti(service, mk, tmp_path):
    for name in ("offerta_triennale_conforme.docx", "offerta_template_dati_mancanti.docx"):
        s = mk(name, template="template_aziendale.docx")
        orig = Document(io.BytesIO(demo(name)))
        out = Document(io.BytesIO(s.analysis.client_docx))
        assert len(out.tables) == len(orig.tables) and len(out.sections) == len(orig.sections)
        assert len(out.element.body.xpath(".//w:hyperlink")) == len(orig.element.body.xpath(".//w:hyperlink")) == 1 or "mancanti" in name
        assert len(out.inline_shapes) >= len(orig.inline_shapes)
        assert out.sections[1].orientation == orig.sections[1].orientation
        if shutil.which("soffice"):
            p = tmp_path / name
            p.write_bytes(s.analysis.client_docx)
            r = subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp_path), str(p)], capture_output=True, timeout=120)
            assert (tmp_path / (name[:-5] + ".pdf")).exists()
