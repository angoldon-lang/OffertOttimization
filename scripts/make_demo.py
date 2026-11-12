"""Genera i documenti demo FITTIZI in demo/ (offerte, template, quotazioni fornitore)."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "demo"
LOGO = ROOT / "profiles" / "assets" / "logo_demo.png"


def add_hyperlink(par, url: str, text: str):
    r_id = par.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), r_id)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    h.append(r)
    par._p.append(h)


def pseudo_heading(doc, text, font="Courier New", size=13):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.name = font
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor(0x99, 0x00, 0x00)
    return p


def body(doc, text, font=None, size=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    if font:
        r.font.name = font
    if size:
        r.font.size = Pt(size)
    return p


def table(doc, rows, widths=None, style="Table Grid"):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = style
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            t.cell(i, j).text = v
    return t


HEAD = ["Descrizione", "SKU", "Q.tà", "Prezzo unitario", "Sconto", "Importo", "Periodicità"]


def build_offer(path: Path, *, wrong_total=False, messy=False, missing=False, two_lines=False, with_extras=True):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    h = (lambda t, lvl=1: doc.add_heading(t, lvl)) if not messy else (lambda t, lvl=1: pseudo_heading(doc, t))
    if messy:
        pseudo_heading(doc, "OFFERTA COMMERCIALE", size=18)
    else:
        doc.add_heading("Offerta commerciale", 0)
    body(doc, "AD Consulting S.p.A." if not messy else "AD Consulting SpA", font="Courier New" if messy else None)
    body(doc, "Milano, 12/11/2026")
    body(doc, "Spett.le Rossi Impianti S.r.l.")
    body(doc, "Oggetto: Fornitura licenze software Gestionale XYZ – triennio 2027-2029")
    body(doc, "Offerta n. OFF-2026-045 del 12/11/2026" + ("" if not messy else ""))
    if with_extras:
        p = body(doc, "Per approfondimenti: ")
        add_hyperlink(p, "https://www.example.com/gestionale-xyz", "scheda prodotto")
        doc.add_picture(str(LOGO), width=Cm(3))

    h("1. Oggetto dell'offerta")
    body(doc, "Con la presente offriamo la fornitura delle licenze del software Gestionale XYZ "
              + ("sopratutto per le sedi operative, perchè richiesto dal cliente.  " if messy else "per le sedi operative del cliente.")
              + (" Il cliente potrà effetuare l'attivazione in autonomia." if messy else ""), font="Courier New" if messy else None, size=9 if messy else None)
    if messy:
        body(doc, "Come concordato con Bianchi S.p.A. il servizio è un'altro rispetto a quello descritto in precedenza. Offerta n. OFF-2025-101 collegata.")
        body(doc, "Referente: [NOME REFERENTE] – email: <indirizzo email>")

    h("2. Perimetro della fornitura")
    body(doc, "La fornitura comprende 10 licenze utente del modulo base, aggiornamenti e supporto standard.")

    h("3. Condizioni economiche")
    if two_lines:
        rows = [HEAD,
                ["Licenza Gestionale XYZ", "SW-XYZ-01", "10", "100,00 €", "10%", "900,00 €", "Annuale"],
                ["Supporto avanzato 24x7", "SUP-24X7", "1", "500,00 €", "", "500,00 €", "Annuale"],
                ["Totale annuo", "", "", "", "", "1.400,00 €", ""],
                ["Totale triennale (36 mesi)", "", "", "", "", "4.200,00 €", ""]]
    else:
        tot = "2.800,00 €" if wrong_total else "2.700,00 €"
        rows = [HEAD,
                ["Licenza Gestionale XYZ", "SW-XYZ-01", "10", "100,00 €", "10%", "900,00 €", "Annuale"],
                ["Totale annuo", "", "", "", "", "900,00 €", ""],
                ["Totale triennale (36 mesi)", "", "", "", "", tot, ""]]
    t = table(doc, rows)
    if messy:
        for row in t.rows:
            for c in row.cells:
                for r in c.paragraphs[0].runs:
                    r.font.name = "Courier New"
    body(doc, "Importi in euro, IVA esclusa.")
    if wrong_total:
        body(doc, "Il totale complessivo dell'offerta per il triennio è pari a € 2.800,00 IVA esclusa.")

    if not missing:
        h("4. Durata e rinnovo")
        body(doc, "Durata del contratto: 36 mesi dalla data di sottoscrizione. Il contratto si rinnova tacitamente per ulteriori 12 mesi, "
                  "salvo disdetta con preavviso di 60 giorni.")
        h("5. Validità dell'offerta")
        body(doc, "La presente offerta è valida 30 giorni dalla data dell'offerta.")
        h("6. Condizioni di pagamento")
        body(doc, "Pagamento: 30 gg d.f.f.m. con bonifico bancario.")
    else:
        h("4. Durata")
        body(doc, "Durata del contratto: 36 mesi dalla data di sottoscrizione.")
        # niente rinnovo, validità e condizioni di pagamento
    if with_extras:
        new = doc.add_section(WD_SECTION.NEW_PAGE)
        new.orientation = WD_ORIENT.LANDSCAPE
        new.page_width, new.page_height = Cm(29.7), Cm(21)
        h("Allegato A – Specifiche tecniche")
        body(doc, "Elenco dei moduli inclusi nella fornitura (allegato in orizzontale).")
    doc.save(path)


def build_template(path: Path):
    doc = Document()
    s = doc.sections[0]
    s.left_margin = s.right_margin = Cm(2.0)
    s.top_margin, s.bottom_margin = Cm(2.2), Cm(2.0)
    for name, size, color in [("Normal", 10.5, None), ("Heading 1", 15, "0B5394"), ("Heading 2", 12, "0B5394"), ("Title", 24, "0B5394")]:
        st = doc.styles[name]
        st.font.name = "Arial"
        st.font.size = Pt(size)
        rpr = st.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rf.set(qn(a), "Arial")
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
            if rf.get(qn(a)) is not None:
                del rf.attrib[qn(a)]
        if color:
            st.font.color.rgb = RGBColor.from_string(color)
    s.header.paragraphs[0].text = "AD Consulting S.p.A. – Offerta commerciale (template)"
    s.footer.paragraphs[0].text = "Documento riservato – AD Consulting S.p.A."
    doc.add_heading("Offerta commerciale", 0)
    doc.add_paragraph("Spett.le [CLIENTE]")
    doc.add_paragraph("Offerta n. [NUMERO OFFERTA] del [DATA]")
    doc.add_paragraph("Oggetto: [OGGETTO]")
    for t in ["Oggetto dell'offerta", "Perimetro della fornitura", "Condizioni economiche", "Durata e rinnovo",
              "Validità dell'offerta", "Condizioni di pagamento"]:
        doc.add_heading(t, 1)
        doc.add_paragraph("[Testo della sezione]")
        if t == "Condizioni economiche":
            tb = doc.add_table(rows=2, cols=4)
            for j, v in enumerate(["Descrizione", "Quantità", "Prezzo unitario", "Importo"]):
                tb.cell(0, j).text = v
    doc.save(path)


def quote_xlsx(path: Path, rows, *, valid="31/01/2027", currency="EUR", extra_meta=None, title="Quotazione fornitore DEMO S.r.l."):
    wb = Workbook()
    ws = wb.active
    ws.title = "Quotazione"
    ws.append([title])
    ws.append(["Validità quotazione:", valid])
    ws.append(["Valuta:", currency])
    for k, v in (extra_meta or []):
        ws.append([k, v])
    ws.append([])
    ws.append(["SKU", "Descrizione", "Quantità", "Costo unitario", "Totale", "Periodicità"])
    for r in rows:
        ws.append(r)
    wb.save(path)


def main():
    OUT.mkdir(exist_ok=True)
    build_offer(OUT / "offerta_triennale_conforme.docx")
    build_offer(OUT / "offerta_totale_errato.docx", wrong_total=True)
    build_offer(OUT / "offerta_template_dati_mancanti.docx", messy=True, missing=True)
    build_offer(OUT / "offerta_due_voci.docx", two_lines=True)
    build_template(OUT / "template_aziendale.docx")
    # quotazioni
    quote_xlsx(OUT / "quotazione_fornitore_triennale.xlsx",
               [["SW-XYZ-01", "Licenza Gestionale XYZ", 10, 60, 600, "Annuale"]])
    quote_xlsx(OUT / "quotazione_costo_non_abbinato.xlsx",
               [["SW-XYZ-01", "Licenza Gestionale XYZ", 10, 60, 600, "Annuale"],
                ["HW-APP-99", "Appliance di rete dedicata", 1, 450, 450, "Una tantum"]])
    quote_xlsx(OUT / "quotazione_costo_condiviso.xlsx",
               [["SW-XYZ-01", "Licenza Gestionale XYZ", 10, 60, 600, "Annuale"],
                ["BND-SUP-01", "Bundle supporto e manutenzione (comune)", 1, 300, 300, "Annuale"]])
    # PDF quotazione (testuale) via LibreOffice
    d = Document()
    d.add_heading("Quotazione fornitore DEMO S.r.l.", 1)
    d.add_paragraph("Validità quotazione: 31/01/2027 – Valuta: EUR – importi IVA esclusa")
    table(d, [["SKU", "Descrizione", "Quantità", "Costo unitario", "Totale", "Periodicità"],
              ["SW-XYZ-01", "Licenza Gestionale XYZ", "10", "60,00", "600,00", "Annuale"]])
    d.save(OUT / "_quotazione_pdf_src.docx")
    try:
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(OUT),
                        str(OUT / "_quotazione_pdf_src.docx")], check=True, capture_output=True, timeout=120)
        (OUT / "_quotazione_pdf_src.pdf").replace(OUT / "quotazione_fornitore.pdf")
    except Exception as exc:
        print("PDF demo non generato (LibreOffice assente?):", exc, file=sys.stderr)
    (OUT / "_quotazione_pdf_src.docx").unlink(missing_ok=True)
    print("Demo generati in", OUT)


if __name__ == "__main__":
    main()
