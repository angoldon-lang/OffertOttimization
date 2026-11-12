"""Impaginazione del DOCX: stili nativi, tabelle, margini, intestazione/piè di pagina, numerazione, copertina."""
from __future__ import annotations

import copy
import re
from collections import Counter
from dataclasses import dataclass

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

from .docx_utils import DocView, Location, norm
from .econ import header_roles
from .findings import Finding
from .template import Standard

DIRECT_FMT = ("w:rFonts", "w:sz", "w:szCs", "w:color", "w:highlight", "w:shd")


@dataclass
class Action:
    check: str
    title: str
    location: Location
    evidence: str
    proposal: str = "Sistemato automaticamente"
    status: str = "Conforme"
    kind: str = "grafica"

    def to_finding(self) -> Finding:
        return Finding(self.check, self.title, self.status, self.location, self.evidence, self.proposal,
                       auto_fixed=True, kind=self.kind)


def _set_font(style, name: str | None, size=None, color: str | None = None, bold: bool | None = None):
    if name:
        style.font.name = name
        rpr = style.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rf.set(qn(a), name)
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
            if rf.get(qn(a)) is not None:
                del rf.attrib[qn(a)]
    if size:
        style.font.size = Pt(size)
    if color:
        style.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        style.font.bold = bold


def _get_style(doc: Document, name: str, builtin_base: str | None = "Normal"):
    try:
        return doc.styles[name]
    except KeyError:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        if builtin_base:
            try:
                st.base_style = doc.styles[builtin_base]
            except KeyError:
                pass
        st.quick_style = True
        return st


def merge_template_styles(doc: Document, tdoc: Document) -> int:
    dst = doc.styles.element
    n = 0
    for s in tdoc.styles.element.findall(qn("w:style")):
        sid = s.get(qn("w:styleId"))
        if s.xpath("./w:pPr/w:numPr"):
            continue      # evita riferimenti a numerazioni del template assenti nell'offerta
        old = dst.xpath(f'./w:style[@w:styleId="{sid}"]')
        new = copy.deepcopy(s)
        if old:
            dst.replace(old[0], new)
        else:
            dst.append(new)
        n += 1
    td = tdoc.styles.element.find(qn("w:docDefaults"))
    dd = dst.find(qn("w:docDefaults"))
    if td is not None:
        if dd is not None:
            dst.replace(dd, copy.deepcopy(td))
        else:
            dst.insert(0, copy.deepcopy(td))
    return n


def ensure_table_style(doc: Document, std: Standard) -> str:
    sid = "ADTabella"
    if doc.styles.element.xpath(f'./w:style[@w:styleId="{sid}"]'):
        return sid
    pal = std.palette
    has_normal_table = bool(doc.styles.element.xpath('./w:style[@w:styleId="TableNormal"]'))
    xml = f'''<w:style {nsdecls("w")} w:type="table" w:customStyle="1" w:styleId="{sid}">
  <w:name w:val="AD Tabella"/>{'<w:basedOn w:val="TableNormal"/>' if has_normal_table else ''}<w:uiPriority w:val="59"/><w:qFormat/>
  <w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders>
    <w:top w:val="single" w:sz="6" w:space="0" w:color="{pal['primary']}"/>
    <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{pal['primary']}"/>
    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="C9D3E0"/></w:tblBorders>
    <w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>
  <w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="{pal['table_header_text']}"/></w:rPr>
    <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{pal['primary']}"/></w:tcPr></w:tblStylePr>
  <w:tblStylePr w:type="band2Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{pal['table_band']}"/></w:tcPr></w:tblStylePr>
</w:style>'''
    doc.styles.element.append(parse_xml(xml))
    return sid


def _strip_direct(p, counters: Counter, fonts: Counter, sizes: Counter, keep_bold=False) -> int:
    n = 0
    for r in p.runs:
        rpr = r._r.rPr
        if rpr is None:
            continue
        if rpr.find(qn("w:rStyle")) is not None:
            continue
        for tag in DIRECT_FMT:
            for el in rpr.findall(qn(tag)):
                if tag == "w:rFonts":
                    nm = el.get(qn("w:ascii")) or el.get(qn("w:hAnsi"))
                    if nm:
                        fonts[nm] += 1
                elif tag == "w:sz":
                    try:
                        sizes[int(el.get(qn("w:val"))) / 2] += 1
                    except (TypeError, ValueError):
                        pass
                rpr.remove(el)
                n += 1
        if not keep_bold:
            for tag in ("w:b", "w:bCs"):
                for el in rpr.findall(qn(tag)):
                    rpr.remove(el)
    counters["runs"] += 1 if n else 0
    return n


def _add_field(par, instr: str, placeholder: str = "1"):
    def r(child):
        run = OxmlElement("w:r")
        run.append(child)
        par._p.append(run)
        return run
    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin"); r(b)
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {instr} "; r(it)
    s = OxmlElement("w:fldChar"); s.set(qn("w:fldCharType"), "separate"); r(s)
    t = OxmlElement("w:t"); t.text = placeholder; r(t)
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end"); r(e)


def _bottom_border(par, color: str, top: bool = False):
    ppr = par._p.get_or_add_pPr()
    bd = OxmlElement("w:pBdr")
    el = OxmlElement("w:top" if top else "w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "4"), ("w:color", color)):
        el.set(qn(k), v)
    bd.append(el)
    ppr.append(bd)


def _hf_text(part) -> str:
    return re.sub(r"\s+", " ", " ".join(p.text for p in part.paragraphs)).strip()


def _clear(part):
    for p in list(part.paragraphs):
        p._p.getparent().remove(p._p)
    for t in list(getattr(part, "tables", [])):
        t._tbl.getparent().remove(t._tbl)


def apply_format(doc: Document, std: Standard, view: DocView, fields: dict, logo_in_header: bool = True) -> list[Action]:
    acts: list[Action] = []
    pal, fonts, sizes = std.palette, std.fonts, std.sizes

    # ---- 1. stili
    if std.template_doc is not None:
        n = merge_template_styles(doc, std.template_doc)
        acts.append(Action("template", "Stili del template applicati", Location(kind="doc"),
                           f"{n} stili nativi (Normale, Titoli, ecc.) allineati al template «{std.template_name}»."))
    normal = doc.styles["Normal"]
    if std.template_doc is None:
        _set_font(normal, fonts["body"], sizes["body"], pal["text"])
    for lvl, key in ((1, "h1"), (2, "h2"), (3, "h3")):
        st = _get_style(doc, f"Heading {lvl}")
        if std.template_doc is None or not any(s.get(qn("w:styleId")) == f"Heading{lvl}" for s in std.template_doc.styles.element.findall(qn("w:style"))):
            _set_font(st, fonts["heading"], sizes.get(key, 12), pal["primary"] if lvl < 3 else pal["secondary"], True)
            st.paragraph_format.space_before = Pt(14 if lvl == 1 else 10)
            st.paragraph_format.space_after = Pt(6)
            st.paragraph_format.keep_with_next = True
    title = _get_style(doc, "Title")
    if std.template_doc is None:
        _set_font(title, fonts["heading"], sizes["title"], pal["primary"], True)
    tabpar = _get_style(doc, "AD Testo Tabella")
    _set_font(tabpar, None, sizes.get("table", 9.5))
    tabpar.paragraph_format.space_before = Pt(2)
    tabpar.paragraph_format.space_after = Pt(2)
    tbl_style_id = ensure_table_style(doc, std)
    # la definizione di stile `ADTabella` per Normal: se Normal definisce già dimensioni, il paragrafo AD Tabella le sovrascrive

    # ---- 2. paragrafi
    fonts_removed, sizes_removed = Counter(), Counter()
    counters = Counter()
    first_nonempty = next((p for p in view.paras if p.text.strip()), None)
    paras_by_idx = doc.paragraphs
    heading_fixes = 0
    for pi in view.paras:
        p = paras_by_idx[pi.idx]
        if not p.text.strip() and not p.runs:
            continue
        before_style = pi.style
        n_before = counters["runs"]
        stripped = _strip_direct(p, counters, fonts_removed, sizes_removed, keep_bold=not pi.heading_level and pi.style.lower() not in ("title",))
        if pi.heading_level and not pi.in_heading_style:
            is_title = pi is first_nonempty and re.search(r"offerta|preventivo|proposta", norm(pi.text)) is not None
            if is_title:
                p.style = doc.styles["Title"]
                new_style, what = "Title", "titolo documento"
            else:
                lvl = min(pi.heading_level, 3)
                p.style = doc.styles[f"Heading {lvl}"]
                new_style, what = f"Heading {lvl}", "titolo di sezione"
            # il grassetto manuale è assorbito dallo stile
            for r in p.runs:
                if r._r.rPr is not None:
                    for tag in ("w:b", "w:bCs", "w:i"):
                        for el in r._r.rPr.findall(qn(tag)):
                            r._r.rPr.remove(el)
            heading_fixes += 1
            acts.append(Action("impaginazione", f"Stile fuori standard: {what} senza stile Word", Location(pi.section, paragraph=pi.idx, kind="par"),
                               f"«{pi.text[:70]}» era «{before_style}» con formattazione manuale", f"Applicato lo stile nativo «{new_style}»"))
        elif pi.style.lower() in ("normal", "normale") and pi.heading_level is None:
            pass
    # in tabella: stile paragrafo dedicato
    n_tables = 0
    tables_total = len(view.tables)
    direct_tbl = 0
    for ti in view.tables:
        tbl = doc.tables[ti.idx]
        tblPr = tbl._tbl.tblPr
        # stile tabella
        ts = tblPr.find(qn("w:tblStyle"))
        if ts is None:
            ts = OxmlElement("w:tblStyle")
            tblPr.insert(0, ts)
        old_style = ts.get(qn("w:val"))
        ts.set(qn("w:val"), tbl_style_id)
        for tag in ("w:tblBorders",):
            for el in tblPr.findall(qn(tag)):
                tblPr.remove(el)
                direct_tbl += 1
        header_row = None
        for r in range(min(3, len(ti.rows))):
            rl = header_roles([c.text for c in ti.rows[r]])
            if len(rl) >= 2 and ({"amount", "unit"} & set(rl.values())):
                header_row = r
                break
        look = tblPr.find(qn("w:tblLook"))
        if look is None:
            look = OxmlElement("w:tblLook")
            tblPr.append(look)
        has_header = header_row == 0 or (len(ti.rows) > 2 and len(ti.rows[0]) >= 3 and not any(re.search(r"\d", c.text) for c in ti.rows[0]))
        look.set(qn("w:val"), "0420" if has_header else "0400")
        look.set(qn("w:firstRow"), "1" if has_header else "0")
        look.set(qn("w:lastRow"), "0")
        look.set(qn("w:firstColumn"), "0")
        look.set(qn("w:lastColumn"), "0")
        look.set(qn("w:noHBand"), "0")
        look.set(qn("w:noVBand"), "1")
        if has_header:
            trPr = tbl._tbl.tr_lst[0].get_or_add_trPr()
            if trPr.find(qn("w:tblHeader")) is None:
                trPr.append(OxmlElement("w:tblHeader"))
        for tr in tbl._tbl.tr_lst:
            for tc in tr.tc_lst:
                tcPr = tc.tcPr
                if tcPr is not None:
                    for tag in ("w:shd", "w:tcBorders"):
                        for el in tcPr.findall(qn(tag)):
                            tcPr.remove(el)
                            direct_tbl += 1
        for row in tbl.rows:
            for cell in {id(c._tc): c for c in row.cells}.values():
                for cp in cell.paragraphs:
                    _strip_direct(cp, counters, fonts_removed, sizes_removed, keep_bold=True)
                    if cp.style is None or cp.style.name.lower() in ("normal", "normale"):
                        cp.style = doc.styles["AD Testo Tabella"]
        if has_header:   # colore/grassetto espliciti: lo stile Normale definisce un colore che prevarrebbe su quello della tabella
            for cell in {id(c._tc): c for c in tbl.rows[0].cells}.values():
                for cp in cell.paragraphs:
                    for r in cp.runs:
                        r.bold = True
                        r.font.color.rgb = RGBColor.from_string(pal["table_header_text"])
        n_tables += 1
        acts.append(Action("impaginazione", "Stile tabella uniformato", Location(ti.section, table=ti.idx, kind="cell"),
                           f"Tabella {ti.idx + 1}: stile «{old_style or 'nessuno'}» → «AD Tabella»"
                           + (", riga di intestazione ripetuta" if has_header else "")))
    if tables_total:
        # un'unica azione aggregata per le tabelle
        tbl_acts = [a for a in acts if a.title == "Stile tabella uniformato"]
        for a in tbl_acts:
            acts.remove(a)
        first = tbl_acts[0]
        acts.append(Action("impaginazione", f"Stile tabella uniformato ({tables_total} tabelle)", first.location,
                           "; ".join(a.evidence for a in tbl_acts)
                           + (f". Rimossa formattazione diretta di bordi/sfondi ({direct_tbl} elementi)." if direct_tbl else ""),
                           "Applicato lo stile tabella «AD Tabella» (intestazione colorata, righe alternate); larghezze di colonna invariate"))
    if fonts_removed or sizes_removed:
        bf = fonts["body"]
        diff_fonts = {f: c for f, c in fonts_removed.items() if f != bf}
        parts = []
        if diff_fonts:
            parts.append("font " + ", ".join(f"«{f}» ({c})" for f, c in diff_fonts.items()))
        if sizes_removed:
            parts.append("dimensioni " + ", ".join(f"{s:g} pt ({c})" for s, c in sizes_removed.items()))
        if parts:
            acts.append(Action("impaginazione", "Formattazione manuale fuori standard ricondotta agli stili", Location(kind="doc"),
                               "Rimossa formattazione diretta: " + "; ".join(parts) + ". Grassetto/corsivo/link conservati.",
                               f"Font {fonts['body']} {sizes['body']:g} pt dallo stile Normale"))

    # ---- 3. sezioni, margini, dimensione pagina
    m = std.margins_cm
    changed = []
    for i, sec in enumerate(doc.sections):
        old = (sec.top_margin.cm if sec.top_margin is not None else None, sec.bottom_margin.cm if sec.bottom_margin is not None else None,
               sec.left_margin.cm if sec.left_margin is not None else None, sec.right_margin.cm if sec.right_margin is not None else None)
        sec.top_margin, sec.bottom_margin = Cm(m["top"]), Cm(m["bottom"])
        sec.left_margin, sec.right_margin = Cm(m["left"]), Cm(m["right"])
        sec.header_distance, sec.footer_distance = Cm(1.2), Cm(1.0)
        if sec.orientation == WD_ORIENT.LANDSCAPE:
            sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
        else:
            sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        new = (m["top"], m["bottom"], m["left"], m["right"])
        if any(o is None or abs(o - n) > 0.05 for o, n in zip(old, new)):
            changed.append(f"sezione {i + 1}: " + "/".join("-" if o is None else f"{o:.1f}" for o in old) + " → " + "/".join(f"{n:.1f}" for n in new))
    if changed:
        acts.append(Action("impaginazione", "Margini di pagina allineati allo standard", Location(kind="doc"),
                           "Margini sup/inf/sx/dx (cm): " + "; ".join(changed),
                           f"Margini {m['top']}/{m['bottom']}/{m['left']}/{m['right']} cm; formato A4; sezioni e orientamenti preservati"))

    # ---- 4. copertina
    first_sec = doc.sections[0]
    has_title = any(pi.style.lower() == "title" or (pi is first_nonempty and pi.heading_level and re.search(r"offerta|preventivo|proposta", norm(pi.text))) for pi in view.paras[:4])
    cover_added = False
    if std.cover and not has_title:
        _add_cover(doc, std, fields)
        first_sec.different_first_page_header_footer = True
        cover_added = True
        acts.append(Action("impaginazione", "Copertina aggiunta", Location(kind="doc"),
                           "Il documento non aveva un frontespizio.", "Copertina con logo, titolo, cliente, numero e data offerta"))

    # ---- 5. intestazione, piè di pagina, numerazione
    old_h, old_f = _hf_text(first_sec.header), _hf_text(first_sec.footer)
    first_sec.header.is_linked_to_previous = False
    first_sec.footer.is_linked_to_previous = False
    hdr, ftr = first_sec.header, first_sec.footer
    _clear(hdr)
    _clear(ftr)
    width = first_sec.page_width - first_sec.left_margin - first_sec.right_margin
    hp = hdr.add_paragraph()
    hp.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)
    logo_done = False
    if logo_in_header and std.logo_path:
        try:
            hp.add_run().add_picture(str(std.logo_path), height=Cm(0.9))
            logo_done = True
        except Exception:
            pass
    run = hp.add_run("\t" + std.header_text if std.header_text else "")
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor.from_string(pal["secondary"])
    _bottom_border(hp, pal["primary"])
    fp = ftr.add_paragraph()
    fp.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)
    _bottom_border(fp, pal["primary"], top=True)
    r1 = fp.add_run(std.footer_text)
    r1.font.size = Pt(8.5)
    r2 = fp.add_run("\tPagina ")
    r2.font.size = Pt(8.5)
    _add_field(fp, "PAGE")
    fp.add_run(" di ").font.size = Pt(8.5)
    _add_field(fp, "NUMPAGES")
    for r in fp.runs:
        r.font.size = Pt(8.5)
    for sec in doc.sections[1:]:
        sec.header.is_linked_to_previous = True
        sec.footer.is_linked_to_previous = True
    src = std.sources.get("intestazione", "profilo")
    ev = f"Intestazione: «{std.header_text}»" + (" con logo" if logo_done else " (logo non disponibile)") + f"; piè di pagina: «{std.footer_text}» + «Pagina X di Y» (rif. {src})."
    if (old_h and old_h != std.header_text) or (old_f and old_f != std.footer_text):
        acts.append(Action("template", "Intestazione/piè di pagina sostituiti (verificare il contenuto precedente)", Location(kind="header"),
                           f"Prima — intestazione: «{old_h or '—'}»; piè: «{old_f or '—'}». Dopo — {ev}",
                           "Contenuti precedenti sostituiti dallo standard: verificare che non contenessero informazioni da conservare",
                           status="Da verificare"))
    else:
        acts.append(Action("template", "Intestazione, piè di pagina e numerazione applicati", Location(kind="header"), ev))
    return acts


def _add_cover(doc: Document, std: Standard, fields: dict):
    body = doc.element.body
    pal = std.palette
    new = []
    def para(text="", style=None, size=None, color=None, bold=None, align=None, after=None):
        p = doc.add_paragraph()
        if style:
            p.style = doc.styles[style]
        if text:
            r = p.add_run(text)
            if size:
                r.font.size = Pt(size)
            if color:
                r.font.color.rgb = RGBColor.from_string(color)
            if bold is not None:
                r.bold = bold
        if after is not None:
            p.paragraph_format.space_after = Pt(after)
        new.append(p)
        return p
    if std.logo_path:
        p = para()
        p.add_run().add_picture(str(std.logo_path), width=Cm(6))
    for _ in range(5):
        para()
    para("Offerta commerciale", "Title")
    def val(k):
        f = fields.get(k)
        return f.value if f and f.present else None
    if val("oggetto"):
        para(val("oggetto"), None, 15, pal["secondary"], False, after=18)
    for lab, k in (("Cliente", "cliente"), ("Offerta n.", "numero_offerta"), ("Data", "data_offerta")):
        if val(k):
            para(f"{lab} {val(k)}" if lab == "Offerta n." else f"{lab}: {val(k)}", None, 12)
    para()
    para(std.profile.company_name, None, 12, pal["primary"], True)
    br = para()
    br.add_run().add_break(__import__("docx").enum.text.WD_BREAK.PAGE)
    for i, p in enumerate(new):
        body.remove(p._p)
        body.insert(i, p._p)


def verify_docx(expected_bytes: bytes, new_bytes: bytes) -> list[str]:
    """Controlla che il DOCX esportato si apra e conservi tabelle, immagini, link, sezioni e testo."""
    from .docx_utils import open_docx
    problems = []
    try:
        new = open_docx(new_bytes)
    except Exception as exc:
        return [f"Il file esportato non si apre: {exc}"]
    exp = open_docx(expected_bytes)
    nv, ev = DocView(new), DocView(exp)
    if len(nv.tables) != len(ev.tables):
        problems.append(f"Tabelle: {len(ev.tables)} → {len(nv.tables)}")
    for name, xp in (("immagini", ".//a:blip"), ("link", ".//w:hyperlink"), ("interruzioni di sezione", ".//w:sectPr")):
        a = len(exp.element.body.xpath(xp))
        b = len(new.element.body.xpath(xp))
        if name == "immagini":
            b -= 0
        if a != b and not (name == "immagini" and b > a):
            problems.append(f"{name.capitalize()}: {a} → {b}")
    new_texts = Counter(p.text.strip() for p in nv.paras if p.text.strip())
    for p in ev.paras:
        t = p.text.strip()
        if t and new_texts[t] <= 0:
            problems.append(f"Paragrafo non più presente: «{t[:60]}»")
        elif t:
            new_texts[t] -= 1
    for t_old, t_new in zip(ev.tables, nv.tables):
        a = [c.text for r in t_old.rows for c in r]
        b = [c.text for r in t_new.rows for c in r]
        if a != b:
            problems.append(f"Contenuto della tabella {t_old.idx + 1} alterato")
    return problems
