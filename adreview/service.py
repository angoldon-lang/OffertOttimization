"""Orchestrazione: sessioni di revisione, analisi deterministica ricalcolata dopo ogni modifica, approvazioni ed esportazioni."""
from __future__ import annotations

import re
import secrets
import time
import zipfile
import io
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation

from . import llm
from .checks import check_econ, check_fields, offer_expiry
from .config import Profile, Settings
from .docx_utils import Change, DocView, Location, apply_changes, norm, open_docx, save_docx
from .econ import Econ, change_for_cell, extract_and_calc
from .extract import FieldResult, Occ, extract_fields, fmt_date
from .findings import CONTROLS, SEVERITY, Finding, worst
from .formatter import apply_format, verify_docx
from .margin import Alloc, MarginResult, SaleRow, check_quote_validity, compute_margin, sales_from_econ, suggest_allocs
from .margin_xlsx import export_margin_xlsx
from .money import D, fmt_eur, parse_number, quantize
from .report import build_report_docx
from .suppliers import CostLine, QuoteResult, parse_quotation
from .summary import build_summary
from .tracked import add_review_comments, apply_changes_tracked, enable_track_revisions
from .template import Standard, build_standard, compare_template
from .textcheck import check_consistency, check_placeholders, check_typos

OPEN_STATUSES = ("Errore", "Da verificare", "Non verificabile")


@dataclass
class Session:
    id: str
    offer_name: str
    offer_bytes: bytes
    profile_id: str
    template_name: str = ""
    template_bytes: bytes | None = None
    created: float = field(default_factory=time.time)
    quotes: list[QuoteResult] = field(default_factory=list)
    costs: list[CostLine] = field(default_factory=list)
    approved: list[Change] = field(default_factory=list)
    change_state: dict[str, str] = field(default_factory=dict)       # key -> approved|rejected|failed
    user_changes: dict[str, Change] = field(default_factory=dict)    # modifiche proposte dall'utente (in attesa)
    acks: set[str] = field(default_factory=set)
    confirm: dict = field(default_factory=lambda: {"mesi": None, "lines": {}})
    allocs: dict[str, Alloc] = field(default_factory=dict)
    fx: dict[str, Decimal] = field(default_factory=dict)
    no_cost_ok: set[str] = field(default_factory=set)
    qty_confirmed: set[str] = field(default_factory=set)
    margin_base: str = "solo_fornitore"
    rounding: str | None = None
    decimals: int | None = None
    threshold: Decimal | None = None
    llm_consent: bool = False
    llm_changes: list[Change] = field(default_factory=list)
    llm_done: bool = False
    notices: list[str] = field(default_factory=list)
    today: date | None = None
    analysis: "Analysis | None" = None


@dataclass
class Analysis:
    fields: dict
    econ: Econ
    findings: list[Finding]
    controls: list[dict]
    summary: dict
    margin: MarginResult | None
    sales: list[SaleRow]
    client_docx: bytes
    verify_problems: list[str]
    std: Standard
    template_info: dict
    changes: dict[str, Change]
    states: dict[str, str]
    open_findings: list[Finding]
    expiry: date | None
    quote_validity: tuple[str, str] | None
    allocs: dict[str, Alloc]
    decimals: int
    rounding: str
    threshold: Decimal
    history: list[Change]


def review_state(f: Finding, s: Session) -> str:
    if f.change is not None:
        st = s.change_state.get(f.change.key)
        if st == "rejected":
            return "Rifiutata"
        if st == "failed":
            return "Non applicabile al Word"
        return "Da approvare"
    if f.auto_fixed:
        return "Sistemato automaticamente"
    if f.status == "Conforme":
        return "Chiuso"
    if f.status == "Non applicabile":
        return "—"
    if f.id in s.acks:
        return "Preso in carico"
    return "Da trattare" if f.status != "Non verificabile" else "Da verificare a mano"


class Service:
    def __init__(self, settings: Settings, profiles: dict[str, Profile]):
        self.settings = settings
        self.profiles = profiles
        self.sessions: dict[str, Session] = {}
        self.llm_transport = None      # iniettabile nei test

    # ------------------------------------------------------------ sessioni
    def create_session(self, offer_name: str, offer: bytes, profile_id: str, template: tuple[str, bytes] | None = None,
                       quotes: list[tuple[str, bytes]] | None = None, internal_costs: list[dict] | None = None,
                       llm_consent: bool = False, today: date | None = None) -> Session:
        if profile_id not in self.profiles:
            raise ValueError("Profilo aziendale sconosciuto")
        open_docx(offer)    # valida
        if template:
            open_docx(template[1])
        s = Session(secrets.token_urlsafe(12), offer_name, offer, profile_id, template[0] if template else "", template[1] if template else None,
                    llm_consent=llm_consent and self.settings.llm_enabled, today=today)
        for qi, (name, data) in enumerate(quotes or [], 1):
            q = parse_quotation(name, data)
            for n, c in enumerate(q.lines, 1):
                c.id = f"Q{qi}-{n}"
            s.quotes.append(q)
            s.costs += q.lines
        for n, ic in enumerate(internal_costs or [], 1):
            qty, unit = parse_number(str(ic.get("qty", ""))), parse_number(str(ic.get("unit", "")))
            if qty is None and unit is None:
                continue
            s.costs.append(CostLine(f"I{n}", "inserimento manuale (costo diretto interno)", "", category="interno", desc=ic.get("desc", "") or "Costo interno",
                                    qty=qty, unit_cost=unit, periodicity=ic.get("periodicity") or "una_tantum", manual=True))
        self.sessions[s.id] = s
        self.analyze(s)
        return s

    def get(self, sid: str) -> Session | None:
        return self.sessions.get(sid)

    def delete(self, sid: str) -> None:
        self.sessions.pop(sid, None)

    # ------------------------------------------------------------ analisi
    def _dec(self, s: Session) -> tuple[int, str, Decimal]:
        return (s.decimals if s.decimals is not None else self.settings.decimals,
                s.rounding or self.settings.rounding,
                s.threshold if s.threshold is not None else (self.profiles[s.profile_id].margin_threshold_pct or self.settings.margin_threshold))

    def analyze(self, s: Session) -> Analysis:
        profile = self.profiles[s.profile_id]
        decimals, rounding, threshold = self._dec(s)
        today = s.today or date.today()
        std = build_standard(profile, s.template_bytes, s.template_name)

        # 1) documento = originale + modifiche approvate
        doc_a = open_docx(s.offer_bytes)
        applied = apply_changes(doc_a, s.approved)
        keep = []
        for ch, ok in applied:
            if ok:
                keep.append(ch)
            else:
                s.change_state[ch.key] = "failed"
                s.notices.append(f"La modifica «{ch.before} → {ch.after}» non è applicabile automaticamente al Word (testo non trovato/modificato): intervenire a mano.")
        s.approved = keep
        expected_bytes = save_docx(doc_a)
        view = DocView(doc_a)

        # 2) campi, calcoli
        f0 = extract_fields(view, profile)
        months = s.confirm.get("mesi") or f0["durata"].data.get("months")
        econ = extract_and_calc(view, months, s.confirm, decimals, rounding)
        derived = False
        if months is None and econ.span_months:
            months, derived = econ.span_months, True
            econ = extract_and_calc(view, months, s.confirm, decimals, rounding)
        fields = extract_fields(view, profile, has_prices=bool(econ.lines), months_override=s.confirm.get("mesi"))
        if derived and not fields["durata"].present:
            d = fields["durata"]
            d.data["months"], d.data["derived"] = months, True
            d.value = f"{months} mesi (ricavata dai periodi in tabella)"
            d.status = "Da verificare"
            d.note = "Durata non dichiarata nel testo: dedotta dai periodi indicati nella tabella economica; confermarla."
            d.occurrences = [Occ(d.value, econ.periods[0][2], "periodi in tabella")]
        findings: list[Finding] = []
        if view.has_tracked_changes:
            findings.append(Finding("impaginazione", "Revisioni tracciate presenti nel documento", "Non verificabile", Location(kind="doc"),
                                    "Il testo inserito/eliminato come revisione non viene analizzato né modificato.", "Accettare o rifiutare le revisioni in Word e ricaricare il file"))
        if any(t.nested for t in view.tables):
            findings.append(Finding("impaginazione", "Tabelle annidate non analizzate", "Non verificabile", Location(kind="doc"),
                                    "Le tabelle dentro altre tabelle sono conservate ma non analizzate.", "Verificare manualmente"))
        findings += check_fields(fields, today)
        findings += check_econ(econ, months is not None, bool(view.tables))

        # 3) testo
        findings += check_typos(view, profile)
        extra = [("intestazione/piè di pagina", " ".join(p.text for sec in doc_a.sections for part in (sec.header, sec.footer) for p in part.paragraphs))]
        findings += check_placeholders(view, extra)
        corr = {}
        for c in s.approved:
            if c.category in ("totale", "prezzo"):
                b, n = parse_number(c.before), parse_number(c.after)
                if b is not None and n is not None:
                    corr[b] = n
        findings += check_consistency(view, profile, fields, econ, today, corr)
        n_text = len([f for f in findings if f.check == "testo"])
        if n_text == 0:
            findings.append(Finding("testo", "Nessun refuso o errore grammaticale rilevato dalle regole locali", "Conforme", Location(kind="doc"),
                                    "Controllo deterministico (dizionario di refusi + regole di punteggiatura)."))
        if not any(f.check == "segnaposto" and f.status == "Errore" for f in findings):
            findings.append(Finding("segnaposto", "Nessun segnaposto rimasto", "Conforme", Location(kind="doc"), "Cercati [..], {{..}}, XXX, «da compilare»."))
        if not any(f.check == "coerenza" for f in findings):
            findings.append(Finding("coerenza", "Denominazioni, aliquote e riferimenti coerenti", "Conforme", Location(kind="doc"), "Nessuna incoerenza rilevata dalle regole locali."))
        # AI opzionale
        if s.llm_consent and self.settings.llm_enabled and not s.llm_done:
            self._run_llm(s, view)
        for ch in s.llm_changes:
            if s.change_state.get(ch.key) == "approved":
                continue
            findings.append(Finding("testo", "Suggerimento del modello linguistico", "Da verificare", Location(view.paras[ch.para].section if ch.para < len(view.paras) else "", paragraph=ch.para, kind="par"),
                                    f"«{ch.before}» → «{ch.after}» ({ch.reason})", "Approvare o rifiutare la correzione proposta", ch))
        # modifiche proposte dall'utente dalla tabella di revisione
        for ch in s.user_changes.values():
            if s.change_state.get(ch.key) in ("approved",):
                continue
            loc = Location(kind="cell", table=ch.table, row=ch.row, col=ch.col)
            findings.append(Finding("calcoli", f"Modifica proposta dall'utente: {ch.label}", "Da verificare", loc,
                                    f"{ch.ctx_before} → {ch.ctx_after}", "Approvare per riportarla nel Word; il ricalcolo avviene dopo l'approvazione", ch))

        # 4) template + impaginazione (eseguita su una copia)
        doc_f = open_docx(expected_bytes)
        view_f = DocView(doc_f)
        tfind, tinfo = compare_template(view, std, fields)
        findings += tfind
        acts = apply_format(doc_f, std, view_f, fields)
        findings += [a.to_finding() for a in acts]
        client = save_docx(doc_f)
        problems = verify_docx(expected_bytes, client)
        findings.append(Finding("impaginazione", "Verifica del DOCX esportato", "Conforme" if not problems else "Errore", Location(kind="doc"),
                                "Il file si riapre; tabelle, immagini, link, interruzioni di sezione e testo conservati." if not problems else "; ".join(problems[:5]),
                                "" if not problems else "Non usare il file esportato: verificare l'originale"))

        # 5) margine
        sales = sales_from_econ(econ)
        expiry, _, _ = offer_expiry(fields)
        margin, allocs, qv = None, {}, None
        if s.costs:
            sug = suggest_allocs(s.costs, sales)
            allocs = {c.id: (s.allocs[c.id] if c.id in s.allocs else sug.get(c.id, Alloc())) for c in s.costs}
            margin = compute_margin(sales, s.costs, allocs, months, s.fx, threshold, s.no_cost_ok, s.qty_confirmed)
            qv = check_quote_validity(expiry, [(q.filename, q.valid_until) for q in s.quotes])
            findings += self._margin_findings(s, margin, qv, threshold)
        else:
            findings.append(Finding("margine", "Nessuna quotazione fornitore o costo caricato", "Non applicabile", Location(kind="doc"),
                                    "La revisione Word, la verifica dei calcoli e il riepilogo restano disponibili.", ""))
        # 6) riservatezza della copia cliente
        findings += self._leak_check(s, client)

        # 7) finalizza
        uniq: dict[str, Change] = {}
        for f in findings:
            if f.change is not None:
                uniq.setdefault(f.change.key, f.change)
        controls = []
        for cid, name in CONTROLS.items():
            fl = [f for f in findings if f.check == cid]
            st = worst([f.status for f in fl]) if fl else "Non applicabile"
            if cid == "margine" and margin is None:
                st = "Non applicabile"
            controls.append({"id": cid, "name": name, "status": st, "n": len([f for f in fl if f.status in OPEN_STATUSES]), "total": len(fl)})
        open_f = [f for f in findings if f.status in OPEN_STATUSES and not f.auto_fixed and not (f.id in s.acks and f.change is None)]
        summary = build_summary(fields, econ, margin, s.margin_base, qv)
        states = {i: review_state(f, s) for i, f in enumerate(findings)}
        a = Analysis(fields, econ, findings, controls, summary, margin, sales, client, problems, std, tinfo, uniq,
                     states, open_f, expiry, qv, allocs, decimals, rounding, threshold, list(s.approved))
        s.analysis = a
        return a

    # ------------------------------------------------------------ margine
    def _margin_findings(self, s: Session, m: MarginResult, qv, threshold: Decimal) -> list[Finding]:
        out = []
        loc = Location(kind="file")
        st = {"completo": "Conforme", "parziale": "Da verificare", "non calcolabile": "Non verificabile"}[m.status]
        out.append(Finding("margine", f"Margine complessivo: {m.status}", st, loc,
                           ("Voci incluse: " + (", ".join(m.included) or "nessuna")
                            + ("; escluse: " + "; ".join(f"{i} ({r})" for i, r in m.excluded) if m.excluded else "")),
                           "" if m.status == "completo" else "Completare abbinamenti/dati mancanti nella schermata Margine", needs_input=m.status != "completo"))
        for u in m.unmatched:
            out.append(Finding("margine", "Costo non abbinato", "Da verificare", loc, f"«{u.cost.desc or u.cost.sku}» – {u.cost.source}: {'; '.join(u.issues)}",
                               "Abbinare il costo a una voce o ripartirlo con un criterio esplicito", needs_input=True))
        for ml in m.lines + m.options:
            if ml.status != "completa":
                out.append(Finding("margine", f"Voce {ml.sale.id}: margine {ml.status}", "Da verificare", Location(kind="cell", table=None), "; ".join(ml.issues), "Completare i dati nella schermata Margine", needs_input=True))
            if ml.revenue is not None and ml.status == "completa":
                mt = ml.margin_total() if m.has_internal and s.margin_base == "con_interni" else ml.margin_supplier()
                pct = ml.pct(mt, ml.revenue)
                if mt is not None and mt < 0:
                    out.append(Finding("margine", f"Margine negativo sulla voce {ml.sale.id}", "Errore", loc, f"{ml.sale.desc}: margine {fmt_eur(quantize(mt))}", "Rivedere prezzo o costo"))
                elif pct is not None and pct < threshold:
                    out.append(Finding("margine", f"Margine sotto soglia sulla voce {ml.sale.id}", "Da verificare", loc, f"{ml.sale.desc}: {quantize(pct)}% < soglia {threshold}%", "Valutare il prezzo"))
        for cc in m.costs:
            for w in cc.warnings:
                out.append(Finding("margine", "Costo riproporzionato", "Da verificare", loc, f"{cc.cost.desc}: {w}", "Verificare la base di confronto"))
        for q in s.quotes:
            for w in q.warnings:
                out.append(Finding("margine", "Estrazione quotazione incerta o incompleta", "Non verificabile", loc, f"{q.filename}: {w}", "Usare la griglia manuale per correggere/inserire i dati", needs_input=True))
            for c in q.lines:
                if c.uncertain:
                    out.append(Finding("margine", "Riga di costo estratta con incertezza", "Da verificare", loc, f"{c.source}: {c.note}", "Verificare nella griglia dei costi"))
        out.append(Finding("margine", "Validità della quotazione fornitore vs offerta", qv[0], loc, qv[1], ""))
        return out

    def _leak_check(self, s: Session, client: bytes) -> list[Finding]:
        """Nessun dato di costo/margine/fornitore deve comparire nel DOCX cliente (se non già presente nell'offerta originale)."""
        def alltext(b: bytes) -> str:
            with zipfile.ZipFile(io.BytesIO(b)) as z:
                xml = " ".join(z.read(n).decode("utf8", "ignore") for n in z.namelist() if n.startswith("word/") and n.endswith(".xml"))
            return norm(re.sub(r"<[^>]+>", " ", xml))
        if not s.costs:
            return [Finding("riservatezza", "Copia cliente priva di dati interni", "Conforme", Location(kind="doc"),
                            "Nessun costo, margine o quotazione fornitore è mai inserito nel DOCX cliente.")]
        orig, new = alltext(s.offer_bytes), alltext(client)
        tokens = set()
        for c in s.costs:
            for t in (c.desc, c.sku):
                if t and len(t) > 3:
                    tokens.add(norm(t))
            tot = c.total_effective
            for v in (tot, c.unit_cost):
                if v is not None and v > 0:
                    tokens.add(norm(f"{quantize(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")))
        tokens |= {"margine", "marginalita", "costo fornitore", "costi interni"}
        leaks = sorted(t for t in tokens if t and t in new and t not in orig)
        if leaks:
            return [Finding("riservatezza", "Possibile dato interno nella copia cliente", "Errore", Location(kind="doc"), "Trovati: " + ", ".join(leaks[:8]), "Non inviare il file al cliente")]
        return [Finding("riservatezza", "Copia cliente priva di dati interni", "Conforme", Location(kind="doc"),
                        "Verificati testo del corpo, intestazioni e piè di pagina: nessun costo, margine o dato di quotazione fornitore.")]

    def _run_llm(self, s: Session, view: DocView) -> None:
        paras = [(p.idx, p.text) for p in view.paras if p.text.strip()]
        try:
            sug = llm.suggest_corrections(self.settings, paras, self.llm_transport)
        except Exception as exc:
            s.notices.append(f"Servizio AI non raggiungibile ({exc.__class__.__name__}): i controlli deterministici restano validi.")
            s.llm_done = True
            return
        by = {p.idx: p for p in view.paras}
        for it in sug:
            p = by[it["idx"]]
            s.llm_changes.append(Change("para", it["before"], it["after"], "Suggerimento AI", "testo", para=p.idx, origin="ai", ctx_before=p.text, reason=it["reason"]))
        s.llm_done = True

    # ------------------------------------------------------------ azioni utente
    def approve(self, s: Session, key: str) -> bool:
        a = s.analysis
        ch = (a.changes.get(key) if a else None) or s.user_changes.get(key)
        if ch is None or s.change_state.get(key) == "approved":
            return False
        s.change_state[key] = "approved"
        s.approved.append(ch)
        s.user_changes.pop(key, None)
        self.analyze(s)
        return True

    def reject(self, s: Session, key: str) -> None:
        s.change_state[key] = "rejected"
        s.user_changes.pop(key, None)
        self.analyze(s)

    def reopen(self, s: Session, key: str) -> None:
        if s.change_state.get(key) == "rejected":
            s.change_state.pop(key)
        self.analyze(s)

    def ack(self, s: Session, fid: str, on: bool = True) -> None:
        (s.acks.add if on else s.acks.discard)(fid)
        self.analyze(s)

    def set_confirmations(self, s: Session, months: str | None, lines: dict[str, dict]) -> None:
        if months is not None and months.strip() != "":
            n = parse_number(months)
            s.confirm["mesi"] = int(n) if n and n > 0 else None
        elif months is not None:
            s.confirm["mesi"] = None
        for lid, vals in lines.items():
            cur = s.confirm["lines"].setdefault(lid, {})
            for k in ("periodicita", "kind"):
                v = vals.get(k)
                if v:
                    cur[k] = v
                elif k in vals:
                    cur.pop(k, None)
        self.analyze(s)

    def propose_line_edits(self, s: Session, edits: dict[str, dict[str, str]]) -> int:
        """edits: line_id -> {qty, unit, disc, amount} (testo utente). Crea modifiche in attesa di approvazione."""
        a = s.analysis
        n = 0
        by = {l.id: l for l in a.econ.lines}
        mapping = {"qty": ("quantita", "col_qty", "qty", None), "unit": ("prezzo", "col_unit", "unit", 2),
                   "disc": ("sconto", "col_disc", "disc", None), "amount": ("prezzo", "col_amount", "reported", 2)}
        for lid, vals in edits.items():
            l = by.get(lid)
            if not l:
                continue
            for fld, (cat, colattr, cur_attr, dec) in mapping.items():
                raw = (vals.get(fld) or "").strip()
                if not raw:
                    continue
                new = parse_number(raw)
                cur = getattr(l, cur_attr)
                col = getattr(l, colattr)
                if new is None or col is None or (cur is not None and new == cur):
                    continue
                key = {"qty": "qty", "unit": "unit", "disc": "disc", "amount": "amount"}[fld]
                ctx = l.cell_text[key]
                ch = change_for_cell(l.table, l.row, col, ctx, new, f"{ {'qty': 'Quantità', 'unit': 'Prezzo unitario', 'disc': 'Sconto', 'amount': 'Importo'}[fld]} {l.source}",
                                     cat, decimals=dec, origin="utente", reason="modifica manuale in tabella di revisione")
                if ch and ch.key not in s.change_state:
                    s.user_changes[ch.key] = ch
                    n += 1
        self.analyze(s)
        return n

    # --- margine
    def set_margin_settings(self, s: Session, **kw) -> None:
        if "fx" in kw:
            s.fx = {k.upper(): v for k, v in kw["fx"].items() if v and v > 0}
        if kw.get("margin_base") in ("solo_fornitore", "con_interni"):
            s.margin_base = kw["margin_base"]
        if kw.get("threshold") is not None:
            s.threshold = kw["threshold"]
        self.analyze(s)

    def update_cost(self, s: Session, cid: str, **vals) -> None:
        c = next((c for c in s.costs if c.id == cid), None)
        if not c:
            return
        for k, v in vals.items():
            setattr(c, k, v)
        c.uncertain = False if vals else c.uncertain
        self.analyze(s)

    def add_cost(self, s: Session, category: str, **vals) -> CostLine:
        n = len([c for c in s.costs if c.manual]) + 1
        c = CostLine(f"{'I' if category == 'interno' else 'M'}{n}x{secrets.token_hex(2)}", "inserimento manuale", "", category=category, manual=True, **vals)
        s.costs.append(c)
        self.analyze(s)
        return c

    def remove_cost(self, s: Session, cid: str) -> None:
        s.costs = [c for c in s.costs if c.id != cid]
        s.allocs.pop(cid, None)
        self.analyze(s)

    def set_alloc(self, s: Session, cid: str, a: Alloc) -> None:
        a.suggested = False
        s.allocs[cid] = a
        self.analyze(s)

    def confirm_qty(self, s: Session, cid: str, on: bool = True) -> None:
        (s.qty_confirmed.add if on else s.qty_confirmed.discard)(cid)
        self.analyze(s)

    def confirm_no_cost(self, s: Session, sid: str, on: bool = True) -> None:
        (s.no_cost_ok.add if on else s.no_cost_ok.discard)(sid)
        self.analyze(s)

    def confirm_suggestions(self, s: Session) -> None:
        for cid, a in (s.analysis.allocs if s.analysis else {}).items():
            if a.suggested:
                a.suggested = False
                s.allocs[cid] = a
        self.analyze(s)

    # ------------------------------------------------------------ export
    def export_client_docx(self, s: Session) -> bytes:
        return s.analysis.client_docx

    def export_tracked(self, s: Session) -> bytes:
        """DOCX in modalità revisione: modifiche proposte come revisioni tracciate + rilievi come commenti.
        Le modifiche già approvate sono applicate come testo normale; costi e margini non sono mai inclusi."""
        a = s.analysis
        doc = open_docx(s.offer_bytes)
        apply_changes(doc, s.approved)
        pending = [ch for key, ch in a.changes.items() if s.change_state.get(key) not in ("approved", "rejected", "failed")]
        tracked = apply_changes_tracked(doc, pending)
        stats = add_review_comments(doc, a.findings, tracked)
        apply_format(doc, a.std, DocView(doc), a.fields)
        enable_track_revisions(doc)
        data = save_docx(doc)
        leak = self._leak_check(s, data)[0]
        if leak.status == "Errore":
            raise RuntimeError("Il DOCX in revisione conterrebbe dati interni: " + leak.evidence)
        return data

    def export_report(self, s: Session) -> bytes:
        return build_report_docx(s, s.analysis, self.profiles[s.profile_id])

    def export_margin(self, s: Session) -> bytes | None:
        a = s.analysis
        if a.margin is None:
            return None
        econ_lines = {l.id: l for l in a.econ.lines}
        qwarn = [w for q in s.quotes for w in q.warnings]
        meta = {"numero": a.fields["numero_offerta"].value, "cliente": a.fields["cliente"].value, "quote_warnings": qwarn,
                "validity": a.quote_validity[1] if a.quote_validity else "", "fx": s.fx, "rounding": f"{a.rounding}, {a.decimals} decimali"}
        return export_margin_xlsx(a.margin, a.sales, s.costs, a.allocs, econ_lines, meta)
