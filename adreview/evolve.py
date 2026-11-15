"""«Evolvi offerta»: trasforma un'offerta tradizionale (tabella grezza, Oggetto minimale, condizioni vuote) nel formato esteso:
Oggetto a elenco, sezione «Soluzione proposta» (testo, punti, tabella specifiche), tabella economica con prezzi unitari e totale, nota IVA,
condizioni compilate. Formattazione coerente con le offerte AD (Arial, blu 0067B4/3983C5). Può produrre anche una versione con
modifiche tracciate (inserimenti/cancellazioni) da accettare in Word."""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from decimal import Decimal
from xml.sax.saxutils import escape

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.text.paragraph import Paragraph

from .docx_utils import DocView, append_text, norm, visible_text
from .money import D, find_numbers, fmt_num, parse_number, quantize

EV = dict(font="Arial", blue="0067B4", blue_light="3983C5", line="D9D9D9", grey="595959", text="000000", row_text="1F1F1F", spec_label="404040")
WIDTHS_ECON = (5137, 900, 1750, 1850)
WIDTHS_SPEC = (4200, 5437)
SERVICE_WORDS = r"support|installazion|assistenza|servizi|manutenzion|configurazion|formazione|consulenz"
PRODUCT_STOP = r"bundle|support|subscription|installazion|cord|cavo|servizi|licen|assistenza|manutenzion"
AUTHOR = "AD Review"


@dataclass
class EconRow:
    title: str
    desc: str = ""
    qty: Decimal | None = None
    unit: Decimal | None = None
    amount: Decimal | None = None


@dataclass
class EvolveInputs:
    oggetto: list[str] = field(default_factory=list)
    solution_title: str = ""
    solution_markup: str = ""
    econ_rows: list[EconRow] = field(default_factory=list)
    econ_note: str = ""
    conditions: dict[str, str] = field(default_factory=dict)
    remove_dangling_refs: bool = True

    @property
    def total(self) -> Decimal:
        return sum((r.amount for r in self.econ_rows if r.amount is not None), Decimal(0))


# ------------------------------------------------------------------ valori proposti dalla tradizionale
def _split_title_desc(text: str) -> tuple[str, str]:
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    if not lines:
        return "", ""
    return lines[0], " ".join(lines[1:])


def normalize_rows(rows: list[EconRow]) -> list[EconRow]:
    """Regola: se prezzo unitario e quantità sono compilati, l'importo = q.tà × prezzo; altrimenti prezzo = importo ÷ q.tà."""
    out = []
    for r in rows:
        if not (r.title or r.desc or r.amount or r.unit):
            continue
        qty, unit, amount = r.qty, r.unit, r.amount
        if qty is not None and unit is not None:
            amount = quantize(qty * unit)
        elif qty and amount is not None:
            unit = quantize(amount / qty)
        elif amount is not None and qty is None:
            qty = Decimal(1)
            unit = amount
        out.append(EconRow(r.title.strip(), r.desc.strip(), qty, unit, amount))
    return out


def defaults_from_doc(view: DocView, econ, fields: dict) -> EvolveInputs:
    inp = EvolveInputs()
    # righe economiche dalla tabella esistente
    rows: list[EconRow] = []
    for l in econ.lines:
        if l.kind == "alternativa":
            continue
        cell = view.tables[l.table].rows[l.row][0].text if view.tables[l.table].rows[l.row] else l.desc
        title, desc = _split_title_desc(l.desc or cell)
        rows.append(EconRow(title, desc, l.qty, l.unit, l.reported if l.reported is not None else l.amount))
    known = [r.qty for r in rows if r.qty]
    mode = max(set(known), key=known.count) if known else Decimal(1)
    for r in rows:
        if r.qty is None:
            code_like = bool(re.match(r"^[A-Z0-9][A-Z0-9\-_/. ]{3,}$", r.title)) and any(ch.isdigit() for ch in r.title)
            r.qty = mode if code_like else Decimal(1)
    inp.econ_rows = normalize_rows(rows)
    # Oggetto: voci esistenti + una per riga
    inp.oggetto = [f"{_short(r)}{' (n. ' + fmt_num(r.qty, 0) + ')' if r.qty and r.qty != 1 else ''}" for r in inp.econ_rows]
    # prodotto principale
    main = [r for r in inp.econ_rows if r.amount and not re.search(PRODUCT_STOP, (r.title + " " + r.desc), re.I)]
    prod = max(main, key=lambda r: r.amount).desc if main else ""
    prod = prod or (max(main, key=lambda r: r.amount).title if main else "")
    inp.solution_title = f"Soluzione proposta: {prod}" if prod else "Soluzione proposta"
    # note e condizioni
    months = None
    for r in inp.econ_rows:
        m = re.search(r"(\d{1,3})\s*(?:mesi|months|month)\b", r.title + " " + r.desc, re.I)
        if m:
            months = int(m.group(1))
    inp.econ_note = "Prezzi IVA esclusa." + (f" Le sottoscrizioni e il supporto hanno durata di {months} mesi dalla data di attivazione." if months else "")
    for t in view.tables:
        for cells in t.rows:
            if (len(cells) == 2 and cells[0].text.strip() and not cells[1].text.strip() and not re.search(r"\d", cells[0].text)
                    and not re.search(r"(?i)netto|totale|importo", cells[0].text)):
                lab = cells[0].text.strip()
                if re.search(r"(?i)pagament", lab):
                    inp.conditions[lab] = "Bonifico Bancario 30gg Fine Mese"
                elif re.search(r"(?i)validit[àa]\s+contratto|durata", lab):
                    inp.conditions[lab] = f"{months} mesi dalla data di attivazione" if months else ""
                else:
                    inp.conditions[lab] = ""
    inp.solution_markup = skeleton_markup(inp.econ_rows)
    return inp


def _short(r: EconRow) -> str:
    base = r.desc or r.title
    base = re.split(r"[,;(]", base)[0].strip()
    return (f"{r.title} – {base}" if r.desc and re.match(r"^[A-Z0-9][A-Z0-9\-_/. ]{3,}$", r.title) else (r.title if not r.desc else f"{r.title} – {base}"))[:110]


def skeleton_markup(rows: list[EconRow]) -> str:
    out = ["[Descrivere la soluzione proposta: modello, finalità e benefici per il cliente.]", "",
           "## Caratteristiche principali", "- [Funzione] – [descrizione]", "",
           "## Specifiche tecniche", "| Parametro | Valore |", "| [parametro] | [valore] |", ""]
    serv = [r for r in rows if re.search(SERVICE_WORDS, r.title + " " + r.desc, re.I)]
    if serv:
        out.append("## Supporto e servizi professionali")
        for r in serv:
            out.append(f"- {r.title} – {r.desc or '[descrizione del servizio]'}")
    return "\n".join(out).strip()


# ------------------------------------------------------------------ mini-markup
def parse_markup(text: str) -> list[tuple]:
    blocks: list[tuple] = []
    para: list[str] = []
    table: list[list[str]] = []

    def flush():
        nonlocal para, table
        if para:
            blocks.append(("p", " ".join(para)))
            para = []
        if table:
            blocks.append(("t", table))
            table = []

    for raw in (text or "").replace("\r", "").split("\n"):
        line = raw.strip()
        if not line:
            flush()
            continue
        if line.startswith("##"):
            flush(); blocks.append(("h", line.lstrip("#").strip()))
        elif line.startswith(("- ", "• ", "* ")):
            flush()
            body = line[2:].strip()
            m = re.match(r"^(.{2,80}?)\s+[–—-]\s+(.+)$", body) or re.match(r"^(.{2,60}?):\s+(.+)$", body)
            blocks.append(("b", m.group(1), m.group(2)) if m else ("b", "", body))
        elif line.startswith("|"):
            if para:
                flush()
            cells = [c.strip() for c in line.strip("|").split("|")]
            if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
                table.append(cells)
        elif line.startswith(">"):
            flush(); blocks.append(("n", line.lstrip("> ").strip()))
        else:
            if table:
                flush()
            para.append(line)
    flush()
    return blocks


# ------------------------------------------------------------------ costruttori XML (stile offerte AD)
def _run(text: str, size: int = 20, bold=False, italic=False, color: str | None = None) -> str:
    rpr = (f'<w:rFonts w:ascii="{EV["font"]}" w:hAnsi="{EV["font"]}" w:cs="{EV["font"]}"/>' + ("<w:b/>" if bold else "") + ("<w:i/>" if italic else "") +
           f'<w:color w:val="{color or EV["text"]}"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/><w:lang w:val="it-IT"/>')
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


class Builder:
    def __init__(self, doc: Document):
        self.doc = doc
        self.body_style = self._style_id("Body Text") or self._style_id("Normal")
        self.normal_style = self._style_id("Normal")
        self.bullet_num = self._find_bullet_num()
        self.next_id = 9500

    def _style_id(self, name: str) -> str | None:
        try:
            return self.doc.styles[name].style_id
        except KeyError:
            return None

    def _find_bullet_num(self) -> str | None:
        for p in self.doc.paragraphs:
            ppr = p._p.pPr
            if ppr is not None and ppr.numPr is not None and ppr.numPr.numId is not None and p.style is not None and p.style.name == "Body Text":
                return str(ppr.numPr.numId.val)
        return None

    def uid(self) -> int:
        self.next_id += 1
        return self.next_id

    def _p(self, ppr: str, runs: str):
        return parse_xml(f"<w:p {nsdecls('w')}><w:pPr>{ppr}</w:pPr>{runs}</w:p>")

    def body(self, text: str):
        return self._p(f'<w:pStyle w:val="{self.body_style}"/><w:spacing w:before="60" w:after="100" w:line="252" w:lineRule="auto"/><w:ind w:right="300"/><w:jc w:val="both"/>', _run(text))

    def subheading(self, text: str):
        return self._p(f'<w:pStyle w:val="{self.body_style}"/><w:keepNext/><w:spacing w:before="240" w:after="100"/><w:jc w:val="left"/>',
                       _run(text, 24, bold=True, color=EV["blue"]))

    def bullet(self, lead: str, rest: str):
        num = f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{self.bullet_num}"/></w:numPr>' if self.bullet_num else ""
        runs = (_run(lead, bold=True) if lead else "") + _run((" – " if lead else "") + rest if lead else rest)
        if not self.bullet_num:
            runs = _run("• ") + runs
        return self._p(f'<w:pStyle w:val="{self.body_style}"/>{num}<w:spacing w:before="60" w:after="40" w:line="252" w:lineRule="auto"/><w:ind w:right="300"/><w:jc w:val="both"/>', runs)

    def note(self, text: str, size: int = 18):
        return self._p(f'<w:pStyle w:val="{self.body_style}"/><w:spacing w:before="120" w:after="0"/><w:jc w:val="left"/>', _run(text, size, italic=True, color=EV["grey"]))

    def page_break(self):
        return parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:pStyle w:val="{self.normal_style}"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>')

    def heading_like(self, template_heading, text: str):
        h = copy.deepcopy(template_heading)
        for el in list(h):
            if el.tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd"), qn("w:hyperlink")) or el.tag == qn("w:r"):
                h.remove(el)
        # run di testo: copia il primo run del titolo originale (formattazione)
        src = next((r for r in template_heading.iterchildren(qn("w:r"))), None)
        run = copy.deepcopy(src) if src is not None else OxmlElement("w:r")
        for k in list(run):
            if k.tag != qn("w:rPr"):
                run.remove(k)
        rpr = run.find(qn("w:rPr"))
        if rpr is not None:
            for t in rpr.findall(qn("w:rStyle")) + rpr.findall(qn("w:vanish")):
                rpr.remove(t)
        t = OxmlElement("w:t"); t.text = text; run.append(t)
        n = self.uid()
        bs = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{n}" w:name="_Toc{90000 + n}"/>')
        be = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{n}"/>')
        h.append(bs); h.append(be); h.append(run)
        return h

    # tabelle
    def _cell(self, w: int, paras: list[tuple], fill: str | None, border_color: str, v="top", span: int = 1) -> str:
        shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
        gs = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
        ps = ""
        for text, size, bold, color, align, before, italic in paras:
            ps += (f'<w:p><w:pPr><w:pStyle w:val="{self.normal_style}"/><w:spacing w:before="{before}" w:after="0"/><w:jc w:val="{align}"/></w:pPr>'
                   + _run(text, size, bold, italic, color) + "</w:p>")
        return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{gs}<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/>'
                f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/><w:right w:val="nil"/></w:tcBorders>{shd}<w:vAlign w:val="{v}"/></w:tcPr>{ps}</w:tc>')

    def _tbl(self, widths, rows_xml: str):
        grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
        return parse_xml(
            f'<w:tbl {nsdecls("w")}><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:top w:w="90" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar>'
            f'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows_xml}</w:tbl>')

    def spec_table(self, rows: list[list[str]]):
        w0, w1 = WIDTHS_SPEC
        head = rows[0] if rows else ["Parametro", "Valore"]
        hdr = (f'<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>'
               + self._cell(w0, [(head[0], 19, True, "FFFFFF", "left", 0, False)], EV["blue_light"], EV["blue_light"], "center")
               + self._cell(w1, [((head[1] if len(head) > 1 else ""), 19, True, "FFFFFF", "left", 0, False)], EV["blue_light"], EV["blue_light"], "center") + "</w:tr>")
        body = ""
        for r in rows[1:]:
            a = r[0] if r else ""
            b = " ".join(r[1:]) if len(r) > 1 else ""
            body += ("<w:tr><w:trPr><w:cantSplit/></w:trPr>" + self._cell(w0, [(a, 19, False, EV["spec_label"], "left", 0, False)], None, EV["line"], "center")
                     + self._cell(w1, [(b, 19, True, EV["text"], "left", 0, False)], None, EV["line"], "center") + "</w:tr>")
        return self._tbl(WIDTHS_SPEC, hdr + body)

    def econ_table(self, rows: list[EconRow], total: Decimal):
        w = WIDTHS_ECON
        heads = [("Descrizione", "left"), ("Q.tà", "center"), ("Prezzo unitario", "right"), ("Importo", "right")]
        hdr = ('<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>'
               + "".join(self._cell(w[i], [(h, 19, True, "FFFFFF", al, 0, False)], EV["blue_light"], EV["blue_light"], "center") for i, (h, al) in enumerate(heads)) + "</w:tr>")
        body = ""
        eur = lambda v: "" if v is None else f"{fmt_num(v)} €"
        for r in rows:
            paras = [(r.title, 19, True, EV["row_text"], "left", 0, False)]
            if r.desc:
                paras.append((r.desc, 16, False, EV["grey"], "left", 30, False))
            body += ("<w:tr><w:trPr><w:cantSplit/></w:trPr>" + self._cell(w[0], paras, None, EV["line"])
                     + self._cell(w[1], [(fmt_num(r.qty, 0 if r.qty is not None and r.qty == r.qty.to_integral_value() else 2) if r.qty is not None else "", 19, False, EV["text"], "center", 0, False)], None, EV["line"])
                     + self._cell(w[2], [(eur(r.unit), 19, False, EV["text"], "right", 0, False)], None, EV["line"])
                     + self._cell(w[3], [(eur(r.amount), 19, True, EV["text"], "right", 0, False)], None, EV["line"]) + "</w:tr>")
        body += ("<w:tr><w:trPr><w:cantSplit/></w:trPr>" + self._cell(sum(w[:3]), [("Netto a Voi riservato (IVA esclusa)", 21, True, "FFFFFF", "left", 0, False)], EV["blue"], EV["blue"], "center", span=3)
                 + self._cell(w[3], [(eur(total), 21, True, "FFFFFF", "right", 0, False)], EV["blue"], EV["blue"], "center") + "</w:tr>")
        return self._tbl(w, hdr + body)


# ------------------------------------------------------------------ revisioni tracciate sugli elementi inseriti/cancellati
class Tracker:
    def __init__(self, enabled: bool):
        self.enabled = enabled
        self.n = 7000
        self.date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    def _attrs(self, el):
        self.n += 1
        el.set(qn("w:id"), str(self.n)); el.set(qn("w:author"), AUTHOR); el.set(qn("w:date"), self.date)
        return el

    def _mark_p(self, p, kind: str):
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr"); p.insert(0, ppr)
        rpr = ppr.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr"); ppr.append(rpr)
        rpr.insert(0, self._attrs(OxmlElement(f"w:{kind}")))

    def _wrap_runs(self, p, kind: str):
        for r in [c for c in p.iterchildren(qn("w:r"))]:
            if kind == "del":
                for t in r.findall(qn("w:t")):
                    t.tag = qn("w:delText")
            w = self._attrs(OxmlElement(f"w:{kind}"))
            r.addprevious(w)
            w.append(r)

    def inserted(self, el):
        if not self.enabled:
            return el
        paras = [el] if el.tag == qn("w:p") else list(el.iter(qn("w:p")))
        if el.tag == qn("w:tbl"):
            for tr in el.iter(qn("w:tr")):
                trpr = tr.find(qn("w:trPr"))
                if trpr is None:
                    trpr = OxmlElement("w:trPr"); tr.insert(0, trpr)
                trpr.append(self._attrs(OxmlElement("w:ins")))
        for p in paras:
            self._mark_p(p, "ins")
            self._wrap_runs(p, "ins")
        return el

    def deleted_paragraph(self, p):
        if self.enabled:
            self._mark_p(p, "del")
            self._wrap_runs(p, "del")

    def deleted_table(self, tbl):
        if not self.enabled:
            return
        for tr in tbl.iter(qn("w:tr")):
            trpr = tr.find(qn("w:trPr"))
            if trpr is None:
                trpr = OxmlElement("w:trPr"); tr.insert(0, trpr)
            trpr.append(self._attrs(OxmlElement("w:del")))
        for p in tbl.iter(qn("w:p")):
            self._mark_p(p, "del")
            self._wrap_runs(p, "del")


# ------------------------------------------------------------------ applicazione
def _heading_idx(doc: Document, pattern: str):
    body = doc.element.body
    for el in body.iterchildren(qn("w:p")):
        p = Paragraph(el, doc._body)
        if p.style is not None and re.match(r"(?i)heading|titolo", p.style.name) and re.match(pattern, norm(visible_text(p))):
            return el
    return None


def apply_evolution(doc: Document, inp: EvolveInputs, tracked: bool = False) -> dict:
    """Modifica `doc` in place. Ritorna un riepilogo di ciò che è stato fatto/non trovato."""
    b, tr = Builder(doc), Tracker(tracked)
    rep = {"oggetto": False, "soluzione": False, "tabella": False, "nota": False, "condizioni": [], "avvisi": []}
    body = doc.element.body
    rows = normalize_rows(inp.econ_rows)
    # --- Oggetto: sostituisce le voci esistenti con l'elenco
    h_ogg = _heading_idx(doc, r"oggetto")
    h_eco = _heading_idx(doc, r"offerta economica|condizioni economiche")
    if h_ogg is not None and inp.oggetto:
        kids = list(body.iterchildren())
        start = kids.index(h_ogg) + 1
        end = kids.index(h_eco) if h_eco is not None else len(kids)
        item_tpl = None
        intro = None
        last = None
        for el in kids[start:end]:
            if el.tag != qn("w:p"):
                continue
            p = Paragraph(el, doc._body)
            t = visible_text(p).strip()
            if not t:
                continue
            if intro is None:
                intro = el
                last = el
                continue
            if el.xpath("./w:pPr/w:numPr"):
                item_tpl = item_tpl if item_tpl is not None else copy.deepcopy(el)
            tr.deleted_paragraph(el)
            if not tracked:
                body.remove(el)
            else:
                last = el
        anchor = last if last is not None else h_ogg
        for line in [l.strip() for l in inp.oggetto if l.strip()]:
            if item_tpl is not None:
                el = copy.deepcopy(item_tpl)
                runs = list(el.iterchildren(qn("w:r")))
                for r in runs[1:]:
                    el.remove(r)
                if runs:
                    for t in runs[0].findall(qn("w:t")):
                        runs[0].remove(t)
                    tt = OxmlElement("w:t"); tt.text = line; tt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); runs[0].append(tt)
                for c in list(el.iterchildren()):
                    if c.tag in (qn("w:ins"), qn("w:del")):
                        el.remove(c)
                    ppr = el.find(qn("w:pPr"))
                    if ppr is not None and ppr.find(qn("w:rPr")) is not None:
                        for x in ppr.find(qn("w:rPr")).findall(qn("w:ins")) + ppr.find(qn("w:rPr")).findall(qn("w:del")):
                            ppr.find(qn("w:rPr")).remove(x)
            else:
                el = b.bullet("", line)
            anchor.addnext(tr.inserted(el))
            anchor = el
        rep["oggetto"] = True
    # --- Soluzione proposta: prima dell'interruzione di pagina che precede «Offerta Economica»
    blocks = parse_markup(inp.solution_markup)
    if h_eco is not None and (blocks or inp.solution_title.strip()):
        prev = h_eco.getprevious()
        tpl_h = h_eco
        ref = prev if (prev is not None and prev.tag == qn("w:p") and prev.xpath('.//w:br[@w:type="page"]')) else h_eco
        elems = [b.page_break()]                      # la sezione parte su una nuova pagina
        elems.append(b.heading_like(tpl_h, inp.solution_title.strip() or "Soluzione proposta"))
        for blk in blocks:
            if blk[0] == "p":
                elems.append(b.body(blk[1]))
            elif blk[0] == "h":
                elems.append(b.subheading(blk[1]))
            elif blk[0] == "b":
                elems.append(b.bullet(blk[1], blk[2]))
            elif blk[0] == "n":
                elems.append(b.note(blk[1], 16))
            elif blk[0] == "t":
                elems.append(b.spec_table(blk[1]))
                elems.append(b.body(""))
        for el in elems:
            ref.addprevious(tr.inserted(el))
        rep["soluzione"] = True
    # --- tabella economica
    if h_eco is not None and rows:
        kids = list(body.iterchildren())
        i0 = kids.index(h_eco)
        old = next((k for k in kids[i0 + 1:] if k.tag == qn("w:tbl")), None)
        new = tr.inserted(b.econ_table(rows, sum((r.amount or Decimal(0) for r in rows), Decimal(0))))
        if old is not None:
            old.addprevious(new)
            tr.deleted_table(old)
            if not tracked:
                body.remove(old)
            anchor_tbl = new
            # nota: sostituisce il paragrafo «Prezzi IVA esclusa» o ne aggiunge uno dopo la tabella
            nxt = None
            for k in list(new.itersiblings())[:12]:
                if k.tag == qn("w:p") and re.search(r"(?i)prezzi\s+iva", "".join(k.xpath(".//w:t/text()")) + "".join(k.xpath(".//w:delText/text()"))):
                    nxt = k
                    break
            if inp.econ_note.strip():
                note = tr.inserted(b.note(inp.econ_note.strip()))
                if nxt is not None:
                    nxt.addprevious(note)
                    tr.deleted_paragraph(nxt)
                    if not tracked:
                        body.remove(nxt)
                else:
                    new.addnext(note)
                rep["nota"] = True
        else:
            h_eco.addnext(new)
            rep["avvisi"].append("Nessuna tabella economica esistente: ne ho creata una nuova sotto il titolo.")
        rep["tabella"] = True
    # --- condizioni: compila le celle vuote
    from .docx_utils import _Cell
    from .tracked import _Ids, tracked_append
    ids = _Ids(8000)
    for tbl in doc.tables:
        for tr_el in tbl._tbl.tr_lst:
            tcs = tr_el.tc_lst
            if len(tcs) != 2:
                continue
            lab = "".join(tcs[0].xpath(".//w:t/text()")).strip()
            val = inp.conditions.get(lab)
            if val and not "".join(tcs[1].xpath(".//w:t/text()")).strip():
                par = _Cell(tcs[1], tbl).paragraphs[0]
                if tracked:
                    tracked_append(par, val, ids, tr.date)
                else:
                    append_text(par, val)
                rep["condizioni"].append(lab)
    # --- riferimenti a paragrafi inesistenti (es. «Perimetro: si veda il paragrafo Perimetro del servizio.»)
    if inp.remove_dangling_refs:
        heads = [norm(visible_text(Paragraph(e, doc._body))) for e in body.iterchildren(qn("w:p"))
                 if (Paragraph(e, doc._body).style is not None and re.match(r"(?i)heading|titolo", Paragraph(e, doc._body).style.name))]
        for el in list(body.iterchildren(qn("w:p"))):
            t = visible_text(Paragraph(el, doc._body)).strip()
            m = re.match(r"(?i)^perimetro\s*:\s*si veda il paragrafo\s+(.+?)\.?$", t)
            if m and not any(norm(m.group(1)) in h for h in heads):
                tr.deleted_paragraph(el)
                if not tracked:
                    el.getparent().remove(el)
                rep["avvisi"].append(f"Rimosso il rimando a un paragrafo inesistente: «{t}».")
    # i sommari (indice) si aggiornano all'apertura in Word
    st = doc.settings.element
    if st.find(qn("w:updateFields")) is None:
        uf = OxmlElement("w:updateFields"); uf.set(qn("w:val"), "true"); st.append(uf)
    return rep
