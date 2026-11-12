# Revisione offerte commerciali – AD Consulting (MVP interno)

Applicazione web per il team commerciale. Si carica un'offerta Word e si ottiene:

1. una **copia DOCX impaginata** secondo template/profilo aziendale (stili nativi Word, tabelle, margini, intestazione/piè di pagina, numerazione, copertina opzionale);
2. **verifiche** su testo (refusi, grammatica, segnaposto, denominazioni), campi richiesti, **calcoli economici** (quantità × prezzo, sconti, subtotali, totali per anno/contratto, IVA) e confronto con il template;
3. un **riepilogo commerciale** interno con il riferimento al punto dell'offerta da cui proviene ogni dato;
4. se si allega una quotazione fornitore (XLSX/PDF) o si inseriscono costi interni, un **report riservato di marginalità** (XLSX con formule).

Tre download: **DOCX cliente**, **report interno dei controlli** (DOCX, riservato), **XLSX di marginalità** (riservato, se disponibile).
Costi, margini e quotazioni non vengono mai scritti nel DOCX cliente (c'è anche un controllo automatico anti-perdite).

> I documenti e i valori in `demo/` e `profiles/ad_consulting_demo.json` (palette, font, logo, margini, soglie) sono **fittizi**: sostituire con i parametri ufficiali.

## Avvio rapido

Richiede Python ≥ 3.10 (sviluppato con 3.13). LibreOffice è facoltativo (serve solo ai test di apertura/ricalcolo e a rigenerare il PDF demo).

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp config.example.env .env          # opzionale
python -m adreview                  # http://127.0.0.1:8000
```

In alternativa: `uvicorn adreview.app:create_app --factory --host 127.0.0.1 --port 8000`.

Prova guidata con i demo (`python scripts/make_demo.py` li rigenera):

| Cosa provare | File da caricare |
|---|---|
| Offerta corretta + margine 33,33% | `offerta_triennale_conforme.docx` + `quotazione_fornitore_triennale.xlsx` (anche `quotazione_fornitore.pdf`) |
| Totale Word errato (2.800 invece di 2.700) | `offerta_totale_errato.docx` |
| Costo non abbinato → margine parziale | `offerta_triennale_conforme.docx` + `quotazione_costo_non_abbinato.xlsx` |
| Costo condiviso da ripartire | `offerta_due_voci.docx` + `quotazione_costo_condiviso.xlsx` (abbinare il «Bundle» con ripartizione proporzionale) |
| Template, stile fuori standard, sezioni e dati mancanti | `offerta_template_dati_mancanti.docx` + template `template_aziendale.docx` |

## Flusso utente

1. **Carica** offerta (obbligatoria), template, quotazioni, costi diretti interni (es. giornate × costo giornaliero), scegli il profilo. Il template è il riferimento per stili/struttura; il profilo fornisce le regole mancanti (logo, palette, checklist campi, sezioni obbligatorie).
2. **Controlli**: esito per ciascun controllo (*Conforme / Da verificare / Errore / Non verificabile / Non applicabile*) e, per ogni rilievo, posizione (sezione · tabella · riga: la pagina non è determinabile in modo affidabile da un DOCX), evidenza, correzione proposta e stato della revisione.
3. **Approvazioni**: ogni correzione che tocca prezzi, quantità, sconti, date, clausole o testo mostra *prima/dopo* e va approvata **singolarmente**. Il DOCX cliente conserva i valori originali finché non approvi; dopo ogni approvazione tutto viene ricalcolato.
4. **Calcoli**: tabella editabile delle righe economiche con fonte, formula, arrotondamento, valore riportato/calcolato/differenza. Periodicità, tipo (base/opzione/alternativa) e durata ambigui vanno confermati qui; le modifiche numeriche diventano proposte da approvare.
5. **Template e impaginazione**: sezioni obbligatorie (presenza e ordine), colonne attese, segnaposto rimasti; elenco di ciò che è stato sistemato automaticamente (differenze grafiche) e di ciò che richiede intervento (mancanze sostanziali).
6. **Margine** (riservato): griglia costi (modificabile, con fonte per riga), abbinamento suggerito/confermabile (uno-a-molti, molti-a-uno, ripartizione uguale/proporzionale/percentuali manuali, il costo è contato una sola volta), cambio manuale per valute ≠ EUR, base di confronto per durate diverse, validità della quotazione vs offerta.
7. **Riepilogo** in un'unica schermata (valori, durata, rinnovo, validità, margine) e **Esporta**.

I controlli sono indipendenti: senza quotazione fornitore restano disponibili revisione Word, calcoli e riepilogo.

## Regole di calcolo

- Aritmetica `Decimal`; arrotondamento configurabile (`AD_ROUNDING`, `AD_DECIMALS`); formati italiani (`1.234,56`, `10%`).
- Riga = quantità × prezzo unitario × (1 − sconto%). Totali: *generico* (somma righe/subtotali − sconti), *annuo* (canoni annualizzati), *contratto* (una tantum + canoni × durata), IVA (aliquota × netto ricalcolato), lordo (netto + IVA, solo se la riga IVA è presente). Le opzioni sono escluse dai totali.
- Margine € = ricavo netto − costo diretto; margine % = margine € ÷ ricavo netto × 100; per i totali margine complessivo ÷ ricavo complessivo; ricavo 0 → «n.a.». Due basi distinte: *solo costi fornitore* e *con costi interni* (se valorizzati).
- Un costo mancante, non abbinato, con cambio/IVA/periodicità/durata/quantità da confermare lascia il margine **parziale**; il report dice quali voci comprende.
- Scadenza dell'offerta calcolata solo se data di partenza e criterio sono chiari («30 giorni dalla data dell'offerta»).

## Configurazione e AI

`config.example.env` elenca le variabili. I profili sono JSON in `profiles/` (font, dimensioni, palette, margini, logo, copertina, header/footer, sezioni obbligatorie con alias, campi richiesti, dizionario refusi, soglia margine).

Il servizio AI è **disattivato di default**: tutti i controlli deterministici funzionano offline. Se configurato (`AD_LLM_PROVIDER=anthropic` + chiave) e se l'utente acconsente per quella revisione, viene inviato **solo il testo dei paragrafi del corpo** (mai tabelle, importi, quotazioni, costi interni, margini); i suggerimenti sono validati (il frammento deve esistere nel paragrafo) e restano proposte da approvare. La UI mostra sempre cosa verrebbe inviato.

## Test

```bash
python -m pytest -q
```

`tests/test_scenarios.py` copre i cinque scenari richiesti (triennale 900/300/33,33% e 2.700/900; totale errato con Word invariato fino ad approvazione; costo non abbinato e costo condiviso senza doppio conteggio; template/stile/sezione assente/dati «Non indicato»; riservatezza dell'export), inclusi apertura del DOCX in LibreOffice e ricalcolo delle formule dell'XLSX.

## Struttura

```
adreview/  app.py (FastAPI + UI) · service.py (sessioni, analisi, approvazioni, export) · docx_utils.py · extract.py (campi)
           econ.py (righe/totali) · textcheck.py · template.py · formatter.py (impaginazione) · suppliers.py (XLSX/PDF/OCR)
           margin.py · margin_xlsx.py · summary.py · report.py · llm.py · config.py · money.py
profiles/  profilo demo + logo fittizio     demo/  documenti fittizi     scripts/make_demo.py     tests/
```

## Limiti noti

- **Posizione**: nessun numero di pagina (non affidabile da DOCX senza motore di impaginazione): si indicano sezione, tabella e riga.
- **DOCX complessi**: tabelle annidate, caselle di testo, controlli di contenuto (SDT), revisioni tracciate, note a piè di pagina e campi sono *conservati ma non analizzati* né modificati (segnalati quando rilevati). I titoli senza stile Word sono riconosciuti con euristiche (grassetto breve, numerazione, maiuscolo): possibili falsi positivi/negativi. Le larghezze di colonna non vengono toccate; le intestazioni/piè di pagina esistenti sono sostituite dallo standard (il contenuto precedente è riportato nel rilievo); logo da profilo (le immagini nell'intestazione del template non sono copiate); le sezioni successive alla prima ereditano intestazione/piè di pagina. Dal template si importano gli stili (esclusi quelli con numerazione), margini, font e testi di intestazione/piè, non temi o numerazioni.
- **Refusi/grammatica**: dizionario configurabile + regole di punteggiatura, non un correttore completo; per una copertura più ampia serve il servizio AI.
- **Quotazioni**: XLSX con intestazione riconoscibile (SKU/descrizione/quantità/costo/totale/periodicità); PDF con tabelle o righe testuali regolari. Layout diversi richiedono la griglia manuale. L'**OCR** dei PDF scansionati funziona solo con `tesseract` + `pytesseract` installati (non provato in questo ambiente: senza OCR l'app lo segnala e propone l'inserimento manuale). Formule XLSX senza valore memorizzato sono segnalate.
- **Margine**: riproporzione lineare per durate diverse (solo se confermata), cambi fissi inseriti dall'utente, ripartizione per anno solo con durata intera e periodicità note; non gestisce prezzi/costi variabili per anno.
- **Stato**: sessioni solo in memoria (si perdono al riavvio). **Nessuna autenticazione**: l'app ascolta su localhost; per uso condiviso metterla dietro SSO/reverse proxy.
- Verificato con python-docx e LibreOffice; non provato con Microsoft Word. L'integrazione AI è testata solo con trasporto simulato (nessuna chiamata reale).
