"""Modalità revisione: le modifiche proposte diventano revisioni tracciate di Word (w:ins / w:del) e i rilievi diventano commenti,
così l'utente accetta o rifiuta direttamente in Word."""
from __future__ import annotations

import copy
from datetime import datetime, timezone

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell
from docx.text.run import Run

from .docx_utils import Change, DocView, Location, apply_changes, box_paragraphs, vis, visible_text
from .findings import Finding

AUTHOR = "AD Review"
INITIALS = "AD"


class _Ids:
    def __init__(self, start: int = 9000):
        self.n = start

    def next(self) -> str:
        self.n += 1
        return str(self.n)


def _simple(r_el) -> bool:
    kids = [c.tag for c in r_el if c.tag != qn("w:rPr")]
    return kids == [qn("w:t")]


def _clone(r_el, text: str, tag: str = "w:t"):
    c = copy.deepcopy(r_el)
    for k in list(c):
        if k.tag != qn("w:rPr"):
            c.remove(k)
    t = OxmlElement(tag)
    t.set(qn("xml:space"), "preserve")
    t.text = text
    c.append(t)
    return c


def _strip_font(r_el) -> None:
    rpr = r_el.find(qn("w:rPr"))
    if rpr is not None:
        for tag in ("w:rFonts", "w:sz", "w:szCs", "w:color", "w:highlight", "w:shd"):
            for el in rpr.findall(qn(tag)):
                rpr.remove(el)


def _minimal(full: str, i: int, old: str, new: str) -> tuple[int, str, str]:
    """Riduce la sostituzione alla sola parte che cambia (estesa alla parola intera, salvo cambi di soli spazi)."""
    k = 0
    while k < min(len(old), len(new)) and old[k] == new[k]:
        k += 1
    m = 0
    while m < min(len(old), len(new)) - k and old[len(old) - 1 - m] == new[len(new) - 1 - m]:
        m += 1
    s0, e0 = i + k, i + len(old) - m
    o, n = old[k:len(old) - m], new[k:len(new) - m]
    if (o + n).strip() and o.strip() == o and n.strip() == n:
        ls, re_ = s0, e0
        while ls > 0 and (full[ls - 1].isalnum() or full[ls - 1] in "'’"):
            ls -= 1
        while re_ < len(full) and (full[re_].isalnum() or full[re_] in "'’"):
            re_ += 1
        o, n = full[ls:s0] + o + full[e0:re_], full[ls:s0] + n + full[e0:re_]
        s0 = ls
    return s0, o, n


def tracked_append(p, new: str, ids: _Ids, date: str) -> Run | None:
    if not new:
        return None
    parent = p._p
    vis_runs = [r for r in p.runs if vis(r)]
    src = vis_runs[-1]._r if vis_runs else None
    el = _clone(src, new) if src is not None else _clone(OxmlElement("w:r"), new)
    _strip_font(el)
    w_ins = OxmlElement("w:ins")
    for kk, v in (("w:id", ids.next()), ("w:author", AUTHOR), ("w:date", date)):
        w_ins.set(qn(kk), v)
    w_ins.append(el)
    if src is not None and src.getparent() is parent:
        src.addnext(w_ins)
    else:
        parent.append(w_ins)
    return Run(el, p)


def tracked_replace(p, old: str, new: str, ids: _Ids, date: str) -> Run | None:
    runs = p.runs
    texts = [vis(r) for r in runs]
    full = "".join(texts)
    if not old:
        return tracked_append(p, new, ids, date)
    i = full.find(old)
    if i < 0:
        return None
    i, old, new = _minimal(full, i, old, new)
    end, pos = i + len(old), 0
    a = b = None
    off_a = off_b = 0
    for k, t in enumerate(texts):
        s, e = pos, pos + len(t)
        if a is None and (i < e or (not old and i <= e)):
            a, off_a = k, i - s
        if a is not None and end <= e:
            b, off_b = k, end - s
            break
        pos = e
    if a is None or b is None or not all(_simple(runs[k]._r) for k in range(a, b + 1)):
        return None
    parent = p._p
    dels = []
    anchor = None
    for k in range(a, b + 1):
        r_el, T = runs[k]._r, texts[k]
        if not T:
            continue
        lo = off_a if k == a else 0
        hi = off_b if k == b else len(T)
        idx = parent.index(r_el)
        parent.remove(r_el)
        seq = []
        if lo > 0:
            seq.append(_clone(r_el, T[:lo]))
        if hi > lo:
            d = _clone(r_el, T[lo:hi], "w:delText")
            dels.append(d)
            seq.append(d)
        elif anchor is None:
            anchor = len(seq)          # punto di inserimento puro
        if hi < len(T):
            seq.append(_clone(r_el, T[hi:]))
        for j, el in enumerate(seq):
            parent.insert(idx + j, el)
        if anchor is not None and not dels:
            anchor = idx + anchor
    if dels:
        w_del = OxmlElement("w:del")
        for kk, v in (("w:id", ids.next()), ("w:author", AUTHOR), ("w:date", date)):
            w_del.set(qn(kk), v)
        first_idx = parent.index(dels[0])
        for d in dels:
            parent.remove(d)
            w_del.append(d)
        parent.insert(first_idx, w_del)
        ins_at = first_idx + 1
        tmpl = dels[0]
    else:
        ins_at = anchor
        tmpl = parent[max(ins_at - 1, 0)] if len(parent) else None
    w_ins = OxmlElement("w:ins")
    for kk, v in (("w:id", ids.next()), ("w:author", AUTHOR), ("w:date", date)):
        w_ins.set(qn(kk), v)
    new_run = _clone(tmpl if tmpl is not None and tmpl.tag == qn("w:r") or (tmpl is not None and tmpl.tag == qn("w:r")) else runs[a]._r, new) if new else None
    if new_run is None:
        return Run(dels[0], p) if dels else None
    _strip_font(new_run)
    w_ins.append(new_run)
    parent.insert(ins_at, w_ins)
    return Run(new_run, p)


def apply_changes_tracked(doc: Document, changes: list[Change]) -> list[tuple[Change, Run | None]]:
    ids, date = _Ids(), datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    out = []
    paras, tables = doc.paragraphs, doc.tables
    for ch in changes:
        run = None
        try:
            if ch.kind == "para" and ch.para is not None and ch.para < len(paras):
                run = tracked_replace(paras[ch.para], ch.before, ch.after, ids, date)
            elif ch.kind == "cell":
                tbl = tables[ch.table]
                tcs = tbl._tbl.tr_lst[ch.row].tc_lst
                for tc in ([tcs[ch.col]] if ch.col is not None else list(tcs)):
                    pars = _Cell(tc, tbl).paragraphs
                    for pp in (pars[:1] if not ch.before else pars):
                        run = tracked_replace(pp, ch.before, ch.after, ids, date)
                        if run:
                            break
                    if run:
                        break
            elif ch.kind == "box":
                for pp in box_paragraphs(doc):
                    t = visible_text(pp)
                    if (not ch.ctx_before or t.strip() == ch.ctx_before.strip()) and (ch.before in t if ch.before else bool(t)):
                        r_ = tracked_replace(pp, ch.before, ch.after, ids, date)
                        run = run or r_
        except (IndexError, KeyError):
            run = None
        out.append((ch, run))
    return out


NOISY = ("Periodicità/unità da confermare", "Riga non ricalcolabile")


def _anchor_runs(doc: Document, loc: Location):
    try:
        if loc.paragraph is not None and loc.table is None:
            p = doc.paragraphs[loc.paragraph]
            return [r for r in p.runs if r.text.strip()][:1] and [r for r in p.runs if r.text.strip()]
        if loc.table is not None and loc.row is not None:
            tbl = doc.tables[loc.table]
            tcs = tbl._tbl.tr_lst[loc.row].tc_lst
            tc = tcs[loc.col] if loc.col is not None and loc.col < len(tcs) else tcs[0]
            for pp in _Cell(tc, tbl).paragraphs:
                rs = [r for r in pp.runs if r.text.strip()]
                if rs:
                    return rs
    except (IndexError, KeyError):
        return None
    return None


def add_review_comments(doc: Document, findings: list[Finding], tracked: list[tuple[Change, Run | None]], max_comments: int = 120) -> dict:
    n = 0
    stats = {"comments": 0, "general": 0, "skipped_noise": 0}
    for ch, run in tracked:
        if run is not None and ch.reason:
            doc.add_comment(run, text=f"{ch.label}: {ch.reason}", author=AUTHOR, initials=INITIALS)
            stats["comments"] += 1
    general = []
    for f in findings:
        if f.status not in ("Errore", "Da verificare", "Non verificabile") or f.auto_fixed or f.check in ("margine", "riservatezza") or f.change is not None:
            continue
        if f.title in NOISY:
            stats["skipped_noise"] += 1
            continue
        text = f"[{f.status}] {f.title}: {f.evidence}" + (f" → {f.proposal}" if f.proposal else "")
        runs = _anchor_runs(doc, f.location) if (f.location.paragraph is not None or f.location.table is not None) else None
        if runs and stats["comments"] < max_comments:
            doc.add_comment(runs, text=text, author=AUTHOR, initials=INITIALS)
            stats["comments"] += 1
        else:
            general.append(text)
    if stats["skipped_noise"]:
        general.append(f"{stats['skipped_noise']} righe con periodicità/unità da confermare o non ricalcolabili: vedi la schermata Calcoli dell'applicazione.")
    first = next((r for p in doc.paragraphs for r in p.runs if r.text.strip()), None)
    if general and first is not None:
        doc.add_comment(first, text="Rilievi generali (non legati a un punto preciso):\n- " + "\n- ".join(general[:25]) +
                        (f"\n… e altri {len(general) - 25}." if len(general) > 25 else ""), author=AUTHOR, initials=INITIALS)
        stats["comments"] += 1
        stats["general"] = len(general)
    return stats


def enable_track_revisions(doc: Document) -> None:
    st = doc.settings.element
    if st.find(qn("w:trackRevisions")) is None:
        el = OxmlElement("w:trackRevisions")
        st.insert(0, el)
