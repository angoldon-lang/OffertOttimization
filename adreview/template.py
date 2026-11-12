"""Standard di impaginazione (template + profilo) e confronto struttura offerta/template."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher

from docx import Document
from docx.oxml.ns import qn

from .config import Profile
from .docx_utils import DocView, Location, norm, open_docx
from .econ import header_roles
from .findings import Finding


@dataclass
class Standard:
    profile: Profile
    fonts: dict
    sizes: dict
    palette: dict
    margins_cm: dict
    header_text: str
    footer_text: str
    cover: bool
    logo_path: object
    required_sections: list
    template_doc: Document | None = None
    template_name: str = ""
    template_placeholders: list = field(default_factory=list)
    template_table_roles: list = field(default_factory=list)   # [set(ruoli)]
    template_table_headers: list = field(default_factory=list)
    sources: dict = field(default_factory=dict)

    @property
    def has_template(self) -> bool:
        return self.template_doc is not None


def _style_font(doc: Document, name: str):
    try:
        st = doc.styles[name]
    except KeyError:
        return None, None
    fname = st.font.name
    if not fname:
        rf = st.element.xpath("./w:rPr/w:rFonts")
        if rf:
            fname = rf[0].get(qn("w:ascii"))
    size = st.font.size.pt if st.font.size else None
    return fname, size


def _clean_hf(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def build_standard(profile: Profile, template_bytes: bytes | None, template_name: str = "") -> Standard:
    std = Standard(profile, dict(profile.fonts), dict(profile.sizes), dict(profile.palette), dict(profile.margins_cm),
                   profile.header_text, profile.footer_text, profile.cover, profile.logo_path,
                   [dict(s) for s in profile.required_sections])
    std.sources = {k: "profilo" for k in ("font", "dimensioni", "margini", "intestazione", "piè di pagina", "sezioni", "colori")}
    if not template_bytes:
        return std
    tdoc = open_docx(template_bytes)
    std.template_doc, std.template_name = tdoc, template_name
    tview = DocView(tdoc)
    bf, bs = _style_font(tdoc, "Normal")
    hf, hs = _style_font(tdoc, "Heading 1")
    if bf:
        std.fonts["body"] = bf
        std.sources["font"] = "template"
    if hf:
        std.fonts["heading"] = hf
    if bs:
        std.sizes["body"] = bs
        std.sources["dimensioni"] = "template"
    if hs:
        std.sizes["h1"] = hs
    try:
        c = tdoc.styles["Heading 1"].font.color.rgb
        if c is not None:
            std.palette["primary"] = str(c)
            std.sources["colori"] = "template"
    except Exception:
        pass
    s0 = tdoc.sections[0]
    if s0.left_margin is not None:
        std.margins_cm = {"top": s0.top_margin.cm, "bottom": s0.bottom_margin.cm, "left": s0.left_margin.cm, "right": s0.right_margin.cm}
        std.sources["margini"] = "template"
    ht = _clean_hf(" ".join(p.text for p in s0.header.paragraphs)) if not s0.header.is_linked_to_previous else ""
    ft = _clean_hf(" ".join(p.text for p in s0.footer.paragraphs)) if not s0.footer.is_linked_to_previous else ""
    if ht:
        std.header_text, std.sources["intestazione"] = ht, "template"
    if ft:
        std.footer_text, std.sources["piè di pagina"] = ft, "template"
    heads = [h for h in tview.headings if h.heading_level == 1] or tview.headings
    if heads:
        prof_secs = profile.required_sections
        secs = []
        for h in heads:
            title = re.sub(r"^\d+(?:\.\d+)*[.)]?\s*", "", h.text.strip())
            al = []
            for ps in prof_secs:
                if norm(ps["title"]) == norm(title) or norm(title) in [norm(a) for a in ps.get("aliases", [])]:
                    al = ps.get("aliases", [])
            secs.append({"title": title, "aliases": al})
        std.required_sections = secs
        std.sources["sezioni"] = "template"
    for loc, text in tview.scan():
        for m in re.finditer(r"\[([^\]\n]{2,40})\]", text):
            std.template_placeholders.append(m.group(1))
    for t in tview.tables:
        if t.rows:
            roles = header_roles([c.text for c in t.rows[0]])
            if len(roles) >= 2:
                std.template_table_roles.append(set(roles.values()))
                std.template_table_headers.append([c.text for c in t.rows[0]])
    return std


def _strip_num(t: str) -> str:
    return re.sub(r"^\d+(?:\.\d+)*[.)]?\s*", "", t.strip())


def match_section(sec: dict, headings) -> int | None:
    titles = [norm(sec["title"])] + [norm(a) for a in sec.get("aliases", [])]
    best, bi = 0.0, None
    for i, h in enumerate(headings):
        nh = norm(_strip_num(h.text))
        for t in titles:
            if not t:
                continue
            if nh == t:
                return i
            score = SequenceMatcher(None, nh, t).ratio()
            if len(t) >= 6 and (t in nh or nh in t) and len(nh) >= 5:
                score = max(score, 0.85)
            if score > best:
                best, bi = score, i
    return bi if best >= 0.82 else None


def compare_template(view: DocView, std: Standard, fields: dict) -> tuple[list[Finding], dict]:
    out: list[Finding] = []
    info: dict = {"sections": [], "missing": [], "order_ok": True, "source": std.sources.get("sezioni")}
    heads = [h for h in view.headings if h.heading_level]
    origin = f"template «{std.template_name}»" if std.sources.get("sezioni") == "template" else "profilo aziendale"
    positions = []
    for sec in std.required_sections:
        i = match_section(sec, heads)
        info["sections"].append({"title": sec["title"], "found": heads[i].text if i is not None else None})
        if i is None:
            info["missing"].append(sec["title"])
            out.append(Finding("template", f"Sezione obbligatoria assente: «{sec['title']}»", "Errore", Location(kind="doc"),
                               f"Nessun titolo corrispondente a «{sec['title']}» (rif. {origin}).",
                               "Aggiungere la sezione: il contenuto commerciale va redatto dall'utente (non inserito automaticamente).",
                               kind="sostanziale", needs_input=True))
        else:
            positions.append((i, sec["title"], heads[i]))
    idxs = [p[0] for p in positions]
    if idxs != sorted(idxs):
        info["order_ok"] = False
        exp = [p[1] for p in positions]
        got = [p[1] for p in sorted(positions)]
        out.append(Finding("template", "Ordine delle sezioni diverso dal riferimento", "Da verificare",
                           Location(positions[0][2].section, paragraph=positions[0][2].idx, kind="par"),
                           f"Atteso: {' → '.join(exp)}. Nel documento: {' → '.join(got)}.",
                           "Riordinare le sezioni (spostamento non eseguito automaticamente per non alterare il testo)", kind="sostanziale"))
    elif std.required_sections and not info["missing"]:
        out.append(Finding("template", "Sezioni obbligatorie presenti e nell'ordine atteso", "Conforme", Location(kind="doc"),
                           f"{len(std.required_sections)} sezioni verificate (rif. {origin})."))
    # colonne tabelle prezzi
    expected_roles = std.template_table_roles[0] if std.template_table_roles else None
    exp_headers = std.template_table_headers[0] if std.template_table_headers else (std.profile.table_headers_expected or [])
    if expected_roles is None and exp_headers:
        expected_roles = set(header_roles(list(exp_headers)).values())
    if expected_roles:
        for t in view.tables:
            for r in range(min(3, len(t.rows))):
                roles = set(header_roles([c.text for c in t.rows[r]]).values())
                if len(roles) >= 2 and ({"amount", "unit"} & roles):
                    miss = expected_roles - roles - {"sku", "period"}
                    if miss:
                        names = {"desc": "Descrizione", "qty": "Quantità", "unit": "Prezzo unitario", "amount": "Importo", "disc": "Sconto"}
                        out.append(Finding("template", "Colonne attese assenti nella tabella economica", "Da verificare",
                                           Location(t.section, table=t.idx, row=r, kind="cell"),
                                           "Mancano: " + ", ".join(names.get(m, m) for m in sorted(miss)),
                                           "Aggiungere le colonne previste dal riferimento", kind="sostanziale"))
                    break
    # campi attesi dal template non coperti
    text_all = norm(view.full_text())
    keymap = {"cliente": "cliente", "numero offerta": "numero_offerta", "data": "data_offerta", "oggetto": "oggetto"}
    for ph in dict.fromkeys(std.template_placeholders):
        n = norm(ph)
        if n in keymap:
            fr = fields.get(keymap[n])
            if fr is not None and not fr.present and not fr.required:
                out.append(Finding("template", f"Campo atteso dal template non trovato: «{ph}»", "Da verificare", Location(kind="doc"),
                                   "Il template prevede il campo, l'offerta non lo contiene.", "Inserire il dato", kind="sostanziale"))
        elif n not in ("testo della sezione",) and len(n) > 2 and n not in text_all:
            out.append(Finding("template", f"Campo atteso dal template non trovato: «{ph}»", "Da verificare", Location(kind="doc"),
                               "Il template prevede il campo, non individuato nel testo dell'offerta.", "Inserire il dato o confermare che non serve",
                               kind="sostanziale"))
    return out, info
