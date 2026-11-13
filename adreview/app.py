"""Applicazione web FastAPI (UI server-rendered, nessun dato inviato all'esterno salvo AI esplicitamente configurata)."""
from __future__ import annotations

import io
import re
from datetime import date
from decimal import Decimal
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import llm
from .config import Settings, load_profiles
from .econ import KINDS, PERIODS
from .extract import fmt_date
from .findings import CONTROLS
from .margin import Alloc
from .money import fmt_eur, fmt_num, fmt_pct, parse_number, quantize
from .service import Service

BASE = Path(__file__).resolve().parent
TABS = [("riepilogo", "Riepilogo e correzioni"), ("controlli", "Tutti i controlli"), ("calcoli", "Calcoli"), ("template", "Template e impaginazione"),
        ("margine", "Margine (riservato)"), ("export", "Esporta")]
COST_PERIODS = {"una_tantum": "Una tantum", "annuale": "Annuale", "mensile": "Mensile", "trimestrale": "Trimestrale",
                "semestrale": "Semestrale", "periodo_quotazione": "Totale per il periodo della quotazione"}
STATUS_CLASS = {"Conforme": "ok", "Da verificare": "warn", "Errore": "err", "Non verificabile": "unk", "Non applicabile": "na"}


def create_app(settings: Settings | None = None, service: Service | None = None) -> FastAPI:
    settings = settings or Settings.from_env()
    svc = service or Service(settings, load_profiles(settings))
    app = FastAPI(title="Revisione offerte commerciali – AD Consulting", docs_url=None, redoc_url=None)
    app.state.service = svc
    app.mount("/static", StaticFiles(directory=BASE / "static"), name="static")
    tpl = Jinja2Templates(directory=str(BASE / "templates"))

    def dec(s):  # decimali configurati per la sessione
        return s.analysis.decimals if s and s.analysis else settings.decimals

    env = tpl.env
    env.filters["eur"] = lambda v, d=2: fmt_eur(v, d)
    env.filters["num"] = lambda v, d=2: fmt_num(v, d)
    env.filters["pct"] = lambda v: fmt_pct(None if v is None else quantize(v))
    env.filters["qnum"] = lambda v: "—" if v is None else fmt_num(quantize(v))
    env.filters["qpct"] = lambda v: "n.a." if v is None else fmt_pct(quantize(v))
    env.filters["dt"] = fmt_date
    env.filters["cls"] = lambda st: STATUS_CLASS.get(st, "na")
    from markupsafe import Markup, escape

    def _hl(ctx: str, frag: str, tag: str):
        ctx, frag = ctx or frag, frag
        i = ctx.find(frag)
        if i >= 0 and len(ctx) > len(frag) + 140:          # mostra solo il contesto vicino alla modifica
            a, b = max(0, i - 60), min(len(ctx), i + len(frag) + 60)
            ctx = ("…" if a else "") + ctx[a:b] + ("…" if b < len(ctx) else "")
            i = ctx.find(frag)
        if not frag:
            return Markup("<span class='muted'>(vuoto)</span>")
        if i < 0:
            return Markup(f"<{tag}>{{}}</{tag}>").format(ctx)
        return Markup(f"{{}}<{tag}>{{}}</{tag}>{{}}").format(ctx[:i], frag, ctx[i + len(frag):])

    env.filters["diff_before"] = lambda ch: _hl(ch.ctx_before, ch.before, "del")
    env.filters["diff_after"] = lambda ch: _hl(ch.ctx_after, ch.after, "ins")
    env.globals.update(TABS=TABS, PERIODS=PERIODS, KINDS=KINDS, COST_PERIODS=COST_PERIODS, CONTROLS=CONTROLS)

    def get_session(sid: str):
        s = svc.get(sid)
        if not s:
            raise HTTPException(404, "Sessione non trovata o scaduta")
        return s

    def back(request: Request, sid: str, default: str = "riepilogo"):
        nxt = (request.query_params.get("next") or default)
        if not re.fullmatch(r"[a-z]+", nxt):
            nxt = default
        return RedirectResponse(f"/review/{sid}/{nxt}", status_code=303)

    def render(request: Request, name: str, **ctx):
        resp = tpl.TemplateResponse(request, name, ctx)
        resp.headers["Cache-Control"] = "no-store"
        return resp

    @app.get("/healthz")
    def health():
        return {"ok": True}

    @app.get("/", response_class=HTMLResponse)
    def home(request: Request):
        return render(request, "home.html", profiles=svc.profiles, settings=settings, ai_text=llm.disclosure(settings),
                      max_mb=settings.max_upload_mb)

    @app.post("/review")
    async def upload(request: Request):
        form = await request.form()
        offer = form.get("offer")
        if offer is None or not getattr(offer, "filename", ""):
            return render(request, "home.html", profiles=svc.profiles, settings=settings, ai_text=llm.disclosure(settings),
                          max_mb=settings.max_upload_mb, error="Seleziona l'offerta cliente in formato DOCX.")
        limit = settings.max_upload_mb * 1024 * 1024
        async def read(f):
            data = await f.read()
            if len(data) > limit:
                raise ValueError(f"{f.filename}: file troppo grande (max {settings.max_upload_mb} MB)")
            return data
        try:
            offer_b = await read(offer)
            tmpl = form.get("template")
            tmpl_t = (tmpl.filename, await read(tmpl)) if tmpl is not None and getattr(tmpl, "filename", "") else None
            quotes = []
            for q in form.getlist("quotes"):
                if getattr(q, "filename", ""):
                    quotes.append((q.filename, await read(q)))
            ics = []
            for d_, q_, u_, p_ in zip(form.getlist("ic_desc"), form.getlist("ic_qty"), form.getlist("ic_unit"), form.getlist("ic_per")):
                if q_.strip() or u_.strip():
                    ics.append({"desc": d_, "qty": q_, "unit": u_, "periodicity": p_})
            s = svc.create_session(offer.filename, offer_b, form.get("profile", ""), tmpl_t, quotes, ics, llm_consent=form.get("llm_consent") == "on")
        except Exception as exc:
            return render(request, "home.html", profiles=svc.profiles, settings=settings, ai_text=llm.disclosure(settings),
                          max_mb=settings.max_upload_mb, error=str(exc))
        return RedirectResponse(f"/review/{s.id}/riepilogo", status_code=303)

    @app.get("/review/{sid}")
    def review_root(sid: str):
        return RedirectResponse(f"/review/{sid}/riepilogo", status_code=303)

    @app.get("/review/{sid}/{tab}", response_class=HTMLResponse)
    def review(request: Request, sid: str, tab: str):
        s = get_session(sid)
        if tab not in dict(TABS):
            raise HTTPException(404)
        a = s.analysis
        q = request.query_params
        flt_status, flt_check = q.get("esito", ""), q.get("check", "")
        rows = [(i, f) for i, f in enumerate(a.findings)
                if (not flt_status or f.status == flt_status) and (not flt_check or f.check == flt_check)
                and not (tab == "controlli" and not flt_status and f.status == "Conforme" and q.get("tutti") != "1")]
        fixes = [(i, f) for i, f in enumerate(a.findings) if f.change and a.states[i] in ("Da approvare", "Da completare")]
        rejected = [(i, f) for i, f in enumerate(a.findings) if f.change and a.states[i] == "Rifiutata"]
        manual = [(i, f) for i, f in enumerate(a.findings) if not f.change and f.status in ("Errore", "Da verificare", "Non verificabile")
                  and not f.auto_fixed and f.id not in s.acks and f.title not in ("Riga non ricalcolabile", "Periodicità/unità da confermare")
                  and f.check not in ("margine",)]
        return render(request, f"{tab}.html", fixes=fixes, ready=sorted([x for x in fixes if not x[1].change.needs_value], key=lambda x: x[1].change.category == "testo"),
                      needs=[x for x in fixes if x[1].change.needs_value], manual=manual, rejected=rejected, s=s, a=a, tab=tab, rows=rows, flt_status=flt_status, flt_check=flt_check,
                      show_all=q.get("tutti") == "1", profile=svc.profiles[s.profile_id], settings=settings, ai_text=llm.disclosure(settings),
                      n_conformi=len([f for f in a.findings if f.status == "Conforme"]))

    # --- approvazioni
    @app.post("/review/{sid}/change/{key}/{action}")
    def change_action(request: Request, sid: str, key: str, action: str):
        s = get_session(sid)
        if action == "approve":
            svc.approve(s, key)
        elif action == "reject":
            svc.reject(s, key)
        elif action == "reopen":
            svc.reopen(s, key)
        else:
            raise HTTPException(400)
        return back(request, sid)

    @app.post("/review/{sid}/approve-selected")
    async def approve_selected(request: Request, sid: str):
        s = get_session(sid)
        form = await request.form()
        only, rej = form.get("only"), form.get("reject")
        if rej:
            svc.reject(s, rej)
            return back(request, sid, "riepilogo")
        keys = [only] if only else form.getlist("k")
        items = [(k, form.get(f"v|{k}")) for k in keys]
        done, refused = svc.approve_many(s, items)
        if refused:
            s.notices.append(f"{len(refused)} correzione/i non applicate: serve scrivere il valore da inserire nel campo.")
        elif done:
            s.notices.append(f"{done} correzione/i applicate al Word.")
        return back(request, sid, "riepilogo")

    @app.post("/review/{sid}/ack/{fid}")
    async def ack(request: Request, sid: str, fid: str):
        s = get_session(sid)
        form = await request.form()
        svc.ack(s, fid, form.get("on", "1") == "1")
        return back(request, sid)

    # --- conferme e modifiche nella tabella di revisione
    @app.post("/review/{sid}/lines")
    async def lines(request: Request, sid: str):
        s = get_session(sid)
        form = await request.form()
        conf, edits = {}, {}
        for k, v in form.items():
            m = re.fullmatch(r"c\|(.+)\|(periodicita|kind)", k)
            if m:
                conf.setdefault(m.group(1), {})[m.group(2)] = v
            m = re.fullmatch(r"e\|(.+)\|(qty|unit|disc|amount)", k)
            if m:
                edits.setdefault(m.group(1), {})[m.group(2)] = v
        svc.set_confirmations(s, form.get("mesi"), conf)
        n = svc.propose_line_edits(s, edits)
        if n:
            s.notices.append(f"{n} modifica/e proposta/e: approvarle nell'elenco «Modifiche in attesa» per riportarle nel Word.")
        return back(request, sid, "calcoli")

    # --- margine
    def dparse(v: str | None):
        return parse_number(v) if v and v.strip() else None

    @app.post("/review/{sid}/margin/quotes")
    async def add_quotes(request: Request, sid: str):
        from .suppliers import parse_quotation
        s = get_session(sid)
        form = await request.form()
        for q in form.getlist("quotes"):
            if getattr(q, "filename", ""):
                qr = parse_quotation(q.filename, await q.read())
                qi = len(s.quotes) + 1
                for n, c in enumerate(qr.lines, 1):
                    c.id = f"Q{qi}-{n}"
                s.quotes.append(qr)
                s.costs += qr.lines
        svc.analyze(s)
        return back(request, sid, "margine")

    @app.post("/review/{sid}/margin/settings")
    async def margin_settings(request: Request, sid: str):
        s = get_session(sid)
        form = await request.form()
        fx = {k[3:]: dparse(v) for k, v in form.items() if k.startswith("fx_") and dparse(v)}
        svc.set_margin_settings(s, fx=fx, margin_base=form.get("margin_base"), threshold=dparse(form.get("threshold")))
        return back(request, sid, "margine")

    @app.post("/review/{sid}/margin/cost/{cid}")
    async def margin_cost(request: Request, sid: str, cid: str):
        s = get_session(sid)
        form = await request.form()
        g = form.get
        pm = g("period_months", "").strip()
        vals = dict(sku=g("sku", ""), desc=g("desc", ""), qty=dparse(g("qty")), unit_cost=dparse(g("unit_cost")), total=dparse(g("total")),
                    currency=(g("currency", "EUR").strip().upper() or "EUR"), periodicity=g("periodicity") or None,
                    period_months=int(pm) if pm.isdigit() else None, vat_included=g("vat_included") == "on", vat_pct=dparse(g("vat_pct")),
                    base_confirmed=g("base_confirmed") == "on")
        svc.update_cost(s, cid, **vals)
        return back(request, sid, "margine")

    @app.post("/review/{sid}/margin/add")
    async def margin_add(request: Request, sid: str):
        s = get_session(sid)
        form = await request.form()
        g = form.get
        cat = "interno" if g("category") == "interno" else "fornitore"
        svc.add_cost(s, cat, sku=g("sku", ""), desc=g("desc", "") or ("Costo interno" if cat == "interno" else "Costo fornitore"),
                     qty=dparse(g("qty")), unit_cost=dparse(g("unit_cost")), currency=(g("currency", "EUR").upper() or "EUR"),
                     periodicity=g("periodicity") or "una_tantum")
        return back(request, sid, "margine")

    @app.post("/review/{sid}/margin/delete/{cid}")
    def margin_delete(request: Request, sid: str, cid: str):
        svc.remove_cost(get_session(sid), cid)
        return back(request, sid, "margine")

    @app.post("/review/{sid}/margin/alloc/{cid}")
    async def margin_alloc(request: Request, sid: str, cid: str):
        s = get_session(sid)
        form = await request.form()
        mode = form.get("mode", "none")
        targets = form.getlist("targets")
        pcts = {}
        for t in targets:
            v = dparse(form.get(f"pct|{t}"))
            if v is not None:
                pcts[t] = v
        if mode == "single":
            targets = targets[:1]
        svc.set_alloc(s, cid, Alloc(mode if targets or mode == "none" else "none", targets, pcts))
        return back(request, sid, "margine")

    @app.post("/review/{sid}/margin/confirm-suggestions")
    def margin_confirm(request: Request, sid: str):
        svc.confirm_suggestions(get_session(sid))
        return back(request, sid, "margine")

    @app.post("/review/{sid}/margin/confirm-qty/{cid}")
    def margin_qty(request: Request, sid: str, cid: str):
        svc.confirm_qty(get_session(sid), cid)
        return back(request, sid, "margine")

    @app.post("/review/{sid}/margin/no-cost/{lid}")
    def margin_nocost(request: Request, sid: str, lid: str):
        svc.confirm_no_cost(get_session(sid), lid)
        return back(request, sid, "margine")

    # --- download
    def attach(data: bytes, name: str, mime: str):
        return Response(data, media_type=mime, headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(name)}", "Cache-Control": "no-store"})

    DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    @app.get("/review/{sid}/download/{what}")
    def download(sid: str, what: str):
        s = get_session(sid)
        stem = re.sub(r"\.docx$", "", s.offer_name, flags=re.I)
        if what == "cliente.docx":
            return attach(svc.export_client_docx(s), f"{stem} - revisionata.docx", DOCX)
        if what == "revisione.docx":
            return attach(svc.export_tracked(s), f"{stem} - in revisione.docx", DOCX)
        if what == "report.docx":
            return attach(svc.export_report(s), f"{stem} - report interno RISERVATO.docx", DOCX)
        if what == "margine.xlsx":
            data = svc.export_margin(s)
            if data is None:
                raise HTTPException(409, "Analisi di marginalità non disponibile: carica una quotazione fornitore o inserisci i costi.")
            return attach(data, f"{stem} - marginalita RISERVATO.xlsx", XLSX)
        raise HTTPException(404)

    @app.post("/review/{sid}/delete")
    def delete(sid: str):
        svc.delete(sid)
        return RedirectResponse("/", status_code=303)

    return app


def main():
    import uvicorn
    settings = Settings.from_env()
    uvicorn.run(create_app(settings), host=settings.host, port=settings.port)
