"""Lettura strutturata del DOCX (paragrafi, tabelle, sezioni), posizioni e applicazione di modifiche approvate."""
from __future__ import annotations

import hashlib
import io
import re
from dataclasses import dataclass, field
from typing import Iterator

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class DocxError(Exception):
    pass


def open_docx(data: bytes) -> Document:
    try:
        return Document(io.BytesIO(data))
    except Exception as exc:
        raise DocxError(f"Il file non è un DOCX valido ({exc.__class__.__name__}).") from exc


def save_docx(doc: Document) -> bytes:
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


@dataclass(frozen=True)
class Location:
    section: str = ""
    paragraph: int | None = None     # indice 0-based tra i paragrafi di primo livello
    table: int | None = None         # 0-based
    row: int | None = None           # 0-based
    col: int | None = None
    kind: str = "doc"                # doc | par | cell | header | footer | file

    @property
    def label(self) -> str:
        if self.kind == "box":
            return "Casella di testo / controllo di contenuto"
        if self.kind == "header":
            return "Intestazione / piè di pagina"
        parts = []
        if self.section:
            parts.append(f"Sezione «{self.section}»")
        if self.table is not None:
            t = f"Tabella {self.table + 1}"
            if self.row is not None:
                t += f", riga {self.row + 1}"
            if self.col is not None:
                t += f", colonna {self.col + 1}"
            parts.append(t)
        elif self.paragraph is not None:
            parts.append(f"paragrafo {self.paragraph + 1}")
        return " · ".join(parts) if parts else "Documento"


@dataclass
class Change:
    kind: str                 # cell | para
    before: str
    after: str
    label: str
    category: str             # prezzo | quantita | sconto | totale | data | clausola | testo
    table: int | None = None
    row: int | None = None
    col: int | None = None
    para: int | None = None
    origin: str = "auto"      # auto | utente | ai
    ctx_before: str = ""
    reason: str = ""

    @property
    def key(self) -> str:
        raw = f"{self.kind}|{self.table}|{self.row}|{self.col}|{self.para}|{self.before}|{self.after}"
        return hashlib.sha1(raw.encode()).hexdigest()[:12]

    @property
    def ctx_after(self) -> str:
        return self.ctx_before.replace(self.before, self.after, 1) if self.ctx_before else self.after


@dataclass
class CInfo:
    tc: object
    text: str
    row: int
    col: int          # indice della cella nella riga (w:tc)
    grid: int         # colonna di griglia di partenza
    span: int = 1


@dataclass
class PInfo:
    idx: int
    p: Paragraph
    text: str
    style: str
    section: str
    heading_level: int | None
    is_list: bool
    in_heading_style: bool


@dataclass
class TInfo:
    idx: int
    tbl: Table
    section: str
    rows: list[list[CInfo]]
    nested: int = 0


def style_name(p: Paragraph) -> str:
    try:
        return p.style.name if p.style is not None else "Normal"
    except Exception:
        return "Normal"


def _all_bold(p: Paragraph) -> bool:
    runs = [r for r in p.runs if r.text.strip()]
    return bool(runs) and all(bool(r.bold) for r in runs)


def heading_level(p: Paragraph, text: str, sname: str) -> int | None:
    m = re.match(r"(?:heading|titolo)\s*(\d)", sname, re.I)
    if m and text.strip():
        return int(m.group(1))
    if sname.lower() in ("title", "titolo", "subtitle", "sottotitolo"):
        return None
    ppr = p._p.pPr
    if ppr is not None and ppr.numPr is not None:
        return None
    t = text.strip()
    if not t or len(t) > 90 or t.endswith((".", ";", ",")):
        return None
    if re.search(r":\s*\S", t):
        return None
    if sname.lower() not in ("normal", "body text", "default", "corpo del testo", "normale", "no spacing", "heading_pseudo"):
        return None
    numbered = re.match(r"^(\d+(?:\.\d+)*)[.)]?\s+(\S.*)$", t)
    if _all_bold(p) and len(t) <= 80:
        if numbered:
            return numbered.group(1).count(".") + 1
        return 1
    if numbered and len(t) <= 70 and numbered.group(2)[0].isupper() and not re.search(r"\d[.,]\d", t):
        return numbered.group(1).count(".") + 1
    letters = re.sub(r"[^A-Za-zÀ-ÿ]", "", t)
    if (len(letters) >= 4 and t.isupper() and len(t) <= 60 and not re.search(r"\d|\bS\.?P\.?A\.?\b|\bS\.?R\.?L\.?\b|\bP\.?\s?IVA\b", t)
            and getattr(p, "_after_heading", True)):
        return 1
    return None


class DocView:
    """Vista in sola lettura ordinata del corpo del documento."""

    def __init__(self, doc: Document):
        self.doc = doc
        self.paras: list[PInfo] = []
        self.tables: list[TInfo] = []
        self.items: list[tuple[str, object]] = []
        self.headings: list[PInfo] = []
        self.has_tracked_changes = bool(doc.element.body.xpath(".//w:ins | .//w:del"))
        section = ""
        pi = ti = 0
        body = doc.element.body
        for child in body.iterchildren():
            if child.tag == qn("w:p"):
                p = Paragraph(child, doc._body)
                text = p.text
                sname = style_name(p)
                p._after_heading = bool(self.headings)      # maiuscolo come titolo solo dopo un primo titolo vero (evita copertine)
                lvl = heading_level(p, text, sname)
                is_list = child.pPr is not None and child.pPr.numPr is not None
                if lvl:
                    section = text.strip()
                info = PInfo(pi, p, text, sname, section, lvl, is_list, bool(re.match(r"(heading|titolo)\s*\d", sname, re.I)))
                self.paras.append(info)
                self.items.append(("p", info))
                if lvl:
                    self.headings.append(info)
                pi += 1
            elif child.tag == qn("w:tbl"):
                tbl = Table(child, doc._body)
                rows = []
                for r, tr in enumerate(child.tr_lst):
                    cells = []
                    grid = 0
                    for c, tc in enumerate(tr.tc_lst):
                        cell = _Cell(tc, tbl)
                        txt = "\n".join(pp.text for pp in cell.paragraphs).strip()
                        span = tc.grid_span if hasattr(tc, "grid_span") else 1
                        cells.append(CInfo(tc, txt, r, c, grid, span or 1))
                        grid += span or 1
                    rows.append(cells)
                nested = len(child.xpath("./w:tr/w:tc/w:tbl"))
                info = TInfo(ti, tbl, section, rows, nested)
                self.tables.append(info)
                self.items.append(("t", info))
                ti += 1

    # --- scansione testuale
    def scan(self) -> Iterator[tuple[Location, str]]:
        for kind, it in self.items:
            if kind == "p":
                for line in it.text.split("\n"):
                    if line.strip():
                        yield Location(it.section, paragraph=it.idx, kind="par"), line.strip()
            else:
                for r, cells in enumerate(it.rows):
                    texts = [c.text for c in cells]
                    nonempty = [t for t in texts if t]
                    if len(cells) == 2 and len(nonempty) == 2 and ":" not in texts[0]:
                        yield Location(it.section, table=it.idx, row=r, kind="cell"), f"{texts[0]}: {texts[1]}"
                    else:
                        for c in cells:
                            if c.text:
                                yield Location(it.section, table=it.idx, row=r, col=c.col, kind="cell"), c.text.replace("\n", " ")

    def scan_extra(self) -> Iterator[tuple[Location, str]]:
        """Testo in punti che il resto dell'analisi non modifica: caselle di testo, controlli di contenuto (SDT),
        intestazioni e piè di pagina. Solo lettura (usato per riconoscere campi come cliente/numero/data)."""
        seen = set()
        xp = ('.//w:txbxContent//w:p[not(ancestor::*[local-name()="Fallback"])] | ./w:sdt//w:p')
        for pel in self.doc.element.body.xpath(xp):
            if id(pel) in seen:
                continue
            seen.add(id(pel))
            t = "".join(pel.xpath(".//w:t/text()")).strip()
            if t:
                yield Location(kind="box"), t
        done = set()
        for sec in self.doc.sections:
            for part in (sec.header, sec.footer):
                if part.is_linked_to_previous or id(part._element) in done:
                    continue
                done.add(id(part._element))
                for pp in part.paragraphs:
                    if pp.text.strip():
                        yield Location(kind="header"), pp.text.strip()

    def full_text(self) -> str:
        return "\n".join(t for _, t in self.scan())


def replace_in_paragraph(p: Paragraph, old: str, new: str) -> bool:
    """Sostituisce `old` con `new` preservando la formattazione del primo run coinvolto."""
    runs = p.runs
    texts = [r.text for r in runs]
    full = "".join(texts)
    i = full.find(old)
    if i < 0 or not old:
        return False
    end = i + len(old)
    pos = 0
    a = b = None
    off_a = off_b = 0
    for k, t in enumerate(texts):
        s, e = pos, pos + len(t)
        if a is None and i < e:
            a, off_a = k, i - s
        if end <= e:
            b, off_b = k, end - s
            break
        pos = e
    if a is None or b is None:
        return False
    if a == b:
        runs[a].text = texts[a][:off_a] + new + texts[a][off_b:]
    else:
        runs[a].text = texts[a][:off_a] + new
        for k in range(a + 1, b):
            runs[k].text = ""
        runs[b].text = texts[b][off_b:]
    return True


def apply_changes(doc: Document, changes: list[Change]) -> list[tuple[Change, bool]]:
    results = []
    paras = doc.paragraphs
    tables = doc.tables
    for ch in changes:
        ok = False
        try:
            if ch.kind == "para" and ch.para is not None and ch.para < len(paras):
                ok = replace_in_paragraph(paras[ch.para], ch.before, ch.after)
            elif ch.kind == "cell":
                tbl = tables[ch.table]
                tc = tbl._tbl.tr_lst[ch.row].tc_lst[ch.col]
                for pp in _Cell(tc, tbl).paragraphs:
                    if replace_in_paragraph(pp, ch.before, ch.after):
                        ok = True
                        break
        except (IndexError, KeyError):
            ok = False
        results.append((ch, ok))
    return results


def norm(s: str) -> str:
    import unicodedata
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()
