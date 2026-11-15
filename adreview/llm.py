"""Integrazione OPZIONALE con un modello linguistico esterno. Disattivata di default.

Dati inviati (solo se configurata E se l'utente acconsente per quella revisione): il testo dei paragrafi del corpo del documento
(esclusi tabelle, importi delle tabelle, quotazioni fornitore, costi interni e margini). Tutti i controlli deterministici
funzionano anche senza questo servizio."""
from __future__ import annotations

import json
import re
import urllib.request

from .config import Settings

SYSTEM = ("Sei un revisore di offerte commerciali in italiano. Individua SOLO refusi, errori grammaticali e incoerenze testuali evidenti. "
          "Non modificare prezzi, quantità, sconti, date o clausole. Rispondi esclusivamente con un array JSON di oggetti "
          '{"idx": <numero paragrafo>, "before": "<frammento esatto da sostituire>", "after": "<frammento corretto>", "reason": "<motivo breve>"}. '
          "Se non ci sono errori rispondi []. Il frammento 'before' deve comparire identico nel paragrafo.")


def disclosure(settings: Settings) -> str:
    if not settings.llm_enabled:
        return "Servizio AI non configurato: vengono eseguiti solo i controlli deterministici locali. Nessun dato lascia questo computer."
    return (f"Servizio AI configurato ({settings.llm_provider}, modello {settings.llm_model}, endpoint {settings.llm_base_url}). "
            "Se acconsenti, verrà inviato SOLO il testo dei paragrafi del corpo dell'offerta (non le tabelle, non gli importi delle tabelle, "
            "non le quotazioni fornitore, non i costi interni, non i margini). Le correzioni proposte restano soggette ad approvazione.")


def _http_transport(settings: Settings):
    def call(payload: dict) -> str:
        body = json.dumps({"model": settings.llm_model, "max_tokens": 2000, "system": SYSTEM,
                           "messages": [{"role": "user", "content": payload["text"]}]}).encode()
        req = urllib.request.Request(settings.llm_base_url.rstrip("/") + "/v1/messages", data=body, method="POST",
                                     headers={"content-type": "application/json", "x-api-key": settings.llm_api_key,
                                              "anthropic-version": "2023-06-01"})
        with urllib.request.urlopen(req, timeout=60) as r:
            data = json.loads(r.read())
        return "".join(b.get("text", "") for b in data.get("content", []))
    return call


def paragraphs_payload(paras: list[tuple[int, str]], limit: int = 12000) -> tuple[str, list[int]]:
    out, ids, n = [], [], 0
    for idx, t in paras:
        line = f"[{idx}] {t}"
        if n + len(line) > limit:
            break
        out.append(line)
        ids.append(idx)
        n += len(line)
    return "\n".join(out), ids


def suggest_corrections(settings: Settings, paras: list[tuple[int, str]], transport=None) -> list[dict]:
    """Ritorna suggerimenti validati {idx, before, after, reason}. Solleva eccezione su errori di rete (gestita dal chiamante)."""
    text, ids = paragraphs_payload(paras)
    if not text:
        return []
    call = transport or _http_transport(settings)
    raw = call({"text": text})
    m = re.search(r"\[.*\]", raw, re.S)
    if not m:
        return []
    items = json.loads(m.group(0))
    by_idx = dict(paras)
    out = []
    for it in items:
        try:
            idx, b, a = int(it["idx"]), str(it["before"]), str(it["after"])
        except (KeyError, ValueError, TypeError):
            continue
        if idx in by_idx and b and b != a and b in by_idx[idx]:
            out.append({"idx": idx, "before": b, "after": a, "reason": str(it.get("reason", ""))[:200]})
    return out


DRAFT_SYSTEM = ("Sei un redattore di offerte commerciali IT in italiano per AD Consulting. Scrivi il testo della sezione «Soluzione proposta» "
                "usando ESCLUSIVAMENTE questo mini-formato: paragrafi separati da riga vuota; sottotitoli con «## Titolo»; punti elenco «- Termine – descrizione»; "
                "tabella specifiche con righe «| Parametro | Valore |» (prima riga = intestazione); note con «> testo». "
                "Non inventare prestazioni o prezzi: nelle specifiche inserisci solo valori ufficiali del produttore di cui sei certo, altrimenti scrivi [da verificare]. "
                "Includi: introduzione, «## Caratteristiche principali», «## Specifiche tecniche», eventuale «## Servizi di sicurezza inclusi» e «## Supporto e servizi professionali».")


def draft_disclosure(settings: Settings) -> str:
    if not settings.llm_enabled:
        return "Servizio AI non configurato: la bozza della soluzione va scritta a mano (resta disponibile lo scheletro precompilato)."
    return ("Se usi «Genera bozza con AI» vengono inviati SOLO codici, descrizioni e quantità delle righe della tabella economica (mai prezzi, cliente, "
            f"costi o margini) a {settings.llm_base_url} (modello {settings.llm_model}). La bozza va sempre riletta prima di applicarla.")


def draft_solution(settings: Settings, rows: list[tuple[str, str, str]], transport=None) -> str:
    """rows: (titolo/codice, descrizione, quantità). Ritorna il testo nel mini-formato."""
    text = "Righe dell'offerta (codice – descrizione – quantità):\n" + "\n".join(f"- {t} – {d} – q.tà {q}" for t, d, q in rows)
    if transport is not None:
        return transport({"text": text, "system": DRAFT_SYSTEM}).strip()
    body = json.dumps({"model": settings.llm_model, "max_tokens": 3500, "system": DRAFT_SYSTEM, "messages": [{"role": "user", "content": text}]}).encode()
    req = urllib.request.Request(settings.llm_base_url.rstrip("/") + "/v1/messages", data=body, method="POST",
                                 headers={"content-type": "application/json", "x-api-key": settings.llm_api_key, "anthropic-version": "2023-06-01"})
    with urllib.request.urlopen(req, timeout=90) as r:
        data = json.loads(r.read())
    return "".join(b.get("text", "") for b in data.get("content", [])).strip()
