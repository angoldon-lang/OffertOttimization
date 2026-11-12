"""Export XLSX del report di marginalità (formule vive, dati e riferimenti alle fonti). File riservato: mai incluso nel DOCX cliente."""
from __future__ import annotations

import io
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .econ import PERIODS
from .margin import Alloc, MarginLine, MarginResult, SaleRow
from .suppliers import CostLine

HEAD_FILL = PatternFill("solid", fgColor="1F3A5F")
HEAD_FONT = Font(bold=True, color="FFFFFF")
RISERVATO = "RISERVATO – USO INTERNO AD Consulting – non inviare al cliente"


def _f(d):
    return None if d is None else float(d)


def _head(ws, row, headers):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")


def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def export_margin_xlsx(res: MarginResult, sales: list[SaleRow], costs: list[CostLine], allocs: dict[str, Alloc],
                       econ_lines: dict, meta: dict) -> bytes:
    wb = Workbook()
    wsR = wb.active
    wsR.title = "Riepilogo"
    wsV = wb.create_sheet("Voci")
    wsA = wb.create_sheet("Ricavi")
    wsC = wb.create_sheet("Costi")
    wsP = wb.create_sheet("Ripartizione")
    wsF = wb.create_sheet("Fonti e avvisi")
    years = Decimal(res.months) / 12 if res.months else None

    # ---------- Ricavi
    _head(wsA, 1, ["ID", "Descrizione", "SKU", "Tipo", "Periodicità", "Quantità", "Prezzo unitario €", "Sconto %",
                   "Importo riga € (formula)", "Moltiplicatore contratto", "Ricavo contratto € netto IVA", "Ricavo annuo ricorrente €",
                   "Una tantum €", "Fonte (offerta)"])
    for i, s in enumerate(sales, 2):
        el = econ_lines.get(s.id)
        wsA.cell(i, 1, s.id); wsA.cell(i, 2, s.desc); wsA.cell(i, 3, s.sku); wsA.cell(i, 4, s.kind)
        wsA.cell(i, 5, PERIODS.get(s.periodicity or "", "da confermare"))
        if el is not None and el.qty is not None and el.unit is not None:
            wsA.cell(i, 6, _f(el.qty)); wsA.cell(i, 7, _f(el.unit)); wsA.cell(i, 8, _f(el.disc or 0))
            wsA.cell(i, 9, f"=ROUND(F{i}*G{i}*(1-H{i}/100),2)")
        else:
            wsA.cell(i, 9, _f(el.amount) if el is not None else None)
        mult = None
        if s.revenue is not None and el is not None and el.amount:
            mult = s.revenue / el.amount
        wsA.cell(i, 10, _f(mult))
        wsA.cell(i, 11, f"=IF(ISNUMBER(J{i}),I{i}*J{i},\"n.d.\")")
        wsA.cell(i, 12, _f(s.annual) if s.annual is not None else "n.d.")
        wsA.cell(i, 13, _f(s.once) if s.once is not None else 0)
        wsA.cell(i, 14, s.source)
    _widths(wsA, [10, 38, 14, 11, 16, 10, 14, 9, 16, 14, 18, 18, 14, 22])

    # ---------- Costi
    _head(wsC, 1, ["ID", "Fonte", "Categoria", "SKU", "Descrizione", "Quantità", "Costo unitario", "Totale riportato", "Valuta",
                   "Cambio verso EUR", "IVA inclusa", "IVA %", "Periodicità", "Mesi coperti dalla quotazione", "Moltiplicatore contratto",
                   "Costo contratto € netto IVA (formula)", "Note / criticità"])
    ccalc = {cc.cost.id: cc for cc in res.costs}
    for i, c in enumerate(costs, 2):
        cc = ccalc.get(c.id)
        tot = c.total_effective
        vals = [c.id, c.source, c.category, c.sku, c.desc, _f(c.qty), _f(c.unit_cost), _f(tot), c.currency, _f(cc.fx if cc and cc.fx else 1),
                "Sì" if c.vat_included else "No", _f(c.vat_pct) if c.vat_pct is not None else 0,
                PERIODS.get(c.periodicity or "", "da confermare") if c.periodicity != "periodo_quotazione" else "Periodo quotazione",
                c.period_months, _f(cc.mult) if cc and cc.mult is not None else None]
        for j, v in enumerate(vals, 1):
            wsC.cell(i, j, v)
        wsC.cell(i, 16, f"=IF(ISNUMBER(O{i}),H{i}*J{i}/(1+IF(K{i}=\"Sì\",L{i}/100,0))*O{i},0)")
        wsC.cell(i, 17, "; ".join((cc.issues if cc else []) + ([c.note] if c.note else [])))
    _widths(wsC, [10, 34, 11, 13, 34, 9, 12, 14, 8, 10, 9, 7, 16, 14, 13, 18, 50])

    # ---------- Ripartizione
    _head(wsP, 1, ["ID costo", "ID voce di vendita", "Quota %", "Criterio di ripartizione", "Importo contratto € (formula)", "Categoria costo"])
    r = 2
    for cid, shares in res.shares.items():
        a = allocs.get(cid, Alloc())
        for sid, sh in shares.items():
            wsP.cell(r, 1, cid); wsP.cell(r, 2, sid); wsP.cell(r, 3, _f(sh * 100)); wsP.cell(r, 4, a.criterion)
            wsP.cell(r, 5, f"=VLOOKUP(A{r},Costi!$A:$P,16,FALSE)*C{r}/100")
            wsP.cell(r, 6, f"=VLOOKUP(A{r},Costi!$A:$C,3,FALSE)")
            r += 1
    _widths(wsP, [16, 16, 10, 44, 20, 14])
    last_p = max(r - 1, 2)

    # ---------- Voci
    _head(wsV, 1, ["ID", "Descrizione", "Tipo", "Ricavo contratto € netto IVA", "Costo fornitore €", "Costo interno €",
                   "Margine solo fornitore €", "Margine solo fornitore %", "Margine con costi interni €", "Margine con costi interni %",
                   "Stato", "Sotto soglia / negativo", "Avvisi", "Fonte (offerta)"])
    allm = res.lines + res.options
    n_ricavi = len(sales) + 1
    for i, ml in enumerate(allm, 2):
        wsV.cell(i, 1, ml.sale.id); wsV.cell(i, 2, ml.sale.desc); wsV.cell(i, 3, ml.sale.kind)
        wsV.cell(i, 4, f"=VLOOKUP(A{i},Ricavi!$A$2:$K${n_ricavi},11,FALSE)")
        wsV.cell(i, 5, f"=SUMIFS(Ripartizione!$E$2:$E${last_p},Ripartizione!$B$2:$B${last_p},A{i},Ripartizione!$F$2:$F${last_p},\"fornitore\")")
        wsV.cell(i, 6, f"=SUMIFS(Ripartizione!$E$2:$E${last_p},Ripartizione!$B$2:$B${last_p},A{i},Ripartizione!$F$2:$F${last_p},\"interno\")")
        wsV.cell(i, 7, f"=IF(ISNUMBER(D{i}),D{i}-E{i},\"n.d.\")")
        wsV.cell(i, 8, f"=IF(ISNUMBER(G{i}),IF(D{i}=0,\"n.a.\",G{i}/D{i}),\"n.d.\")")
        wsV.cell(i, 9, f"=IF(ISNUMBER(D{i}),D{i}-E{i}-F{i},\"n.d.\")")
        wsV.cell(i, 10, f"=IF(ISNUMBER(I{i}),IF(D{i}=0,\"n.a.\",I{i}/D{i}),\"n.d.\")")
        wsV.cell(i, 11, {"completa": "Completa", "parziale": "Parziale", "mancante": "Costo mancante"}[ml.status])
        wsV.cell(i, 12, f"=IF(ISNUMBER(J{i}),IF(I{i}<0,\"NEGATIVO\",IF(J{i}*100<Riepilogo!$B$5,\"SOTTO SOGLIA\",\"\")),\"\")")
        wsV.cell(i, 13, "; ".join(ml.issues + ml.warnings)); wsV.cell(i, 14, ml.sale.source)
        for col in (8, 10):
            wsV.cell(i, col).number_format = "0.00%"
        for col in (4, 5, 6, 7, 9):
            wsV.cell(i, col).number_format = "#,##0.00"
    _widths(wsV, [10, 36, 10, 16, 14, 14, 16, 14, 16, 14, 14, 16, 60, 24])
    last_v = len(allm) + 1

    # ---------- Riepilogo
    wsR["A1"] = "Report di marginalità – offerta " + str(meta.get("numero", ""))
    wsR["A1"].font = Font(bold=True, size=14)
    wsR["A2"] = RISERVATO
    wsR["A2"].font = Font(bold=True, color="C8102E")
    wsR["A3"] = "Cliente"; wsR["B3"] = meta.get("cliente", "")
    wsR["A4"] = "Durata contratto (mesi)"; wsR["B4"] = res.months if res.months else "n.d."
    wsR["A5"] = "Soglia margine % (configurabile)"; wsR["B5"] = _f(res.threshold)
    wsR["A6"] = "Stato del calcolo"; wsR["B6"] = res.status.upper()
    wsR["A7"] = "Voci incluse nei totali"; wsR["B7"] = ", ".join(res.included) or "—"
    wsR["A8"] = "Voci escluse"; wsR["B8"] = "; ".join(f"{i}: {m}" for i, m in res.excluded) or "—"
    _head(wsR, 10, ["Totale (solo voci incluse, tipo base)", "Valore"])
    inc = f'Voci!$C$2:$C${last_v},"base",Voci!$K$2:$K${last_v},"Completa"'
    wsR["A11"] = "Ricavo contratto € netto IVA"; wsR["B11"] = f"=SUMIFS(Voci!$D$2:$D${last_v},{inc})"
    wsR["A12"] = "Costo fornitore €"; wsR["B12"] = f"=SUMIFS(Voci!$E$2:$E${last_v},{inc})"
    wsR["A13"] = "Costo interno €"; wsR["B13"] = f"=SUMIFS(Voci!$F$2:$F${last_v},{inc})"
    wsR["A14"] = "Margine solo fornitore €"; wsR["B14"] = "=B11-B12"
    wsR["A15"] = "Margine solo fornitore % (margine ÷ ricavo)"; wsR["B15"] = '=IF(B11=0,"n.a.",B14/B11)'
    wsR["A16"] = "Margine con costi interni €"; wsR["B16"] = "=B11-B12-B13"
    wsR["A17"] = "Margine con costi interni %"; wsR["B17"] = '=IF(B11=0,"n.a.",B16/B11)'
    for c in ("B15", "B17"):
        wsR[c].number_format = "0.00%"
    for c in ("B11", "B12", "B13", "B14", "B16"):
        wsR[c].number_format = "#,##0.00"
    wsR["A18"] = "Costi interni valorizzati"; wsR["B18"] = "Sì" if res.has_internal else "No"
    wsR["A19"] = "Controllo anti doppio conteggio: costo totale caricato − costo ripartito (deve essere 0)"
    wsR["B19"] = f"=ROUND(SUM(Costi!$P$2:$P${len(costs) + 1})-SUMIFS(Ripartizione!$E$2:$E${last_p},Ripartizione!$A$2:$A${last_p},\"<>\"),2)" if costs else 0
    wsR["B19"].number_format = "#,##0.00"
    if res.status != "completo":
        wsR["A21"] = "ATTENZIONE: margine PARZIALE/INCOMPLETO – include solo le voci elencate in B7."
        wsR["A21"].font = Font(bold=True, color="C8102E")
    if res.years:
        _head(wsR, 23, ["Anno", "Ricavo €", "Costo fornitore €", "Costo interno €", "Margine €", "Margine %"])
        for k, y in enumerate(res.years):
            rr = 24 + k
            wsR.cell(rr, 1, y["anno"]); wsR.cell(rr, 2, _f(y["ricavo"])); wsR.cell(rr, 3, _f(y["costo_fornitore"]))
            wsR.cell(rr, 4, _f(y["costo_interno"])); wsR.cell(rr, 5, f"=B{rr}-C{rr}-D{rr}")
            wsR.cell(rr, 6, f'=IF(B{rr}=0,"n.a.",E{rr}/B{rr})'); wsR.cell(rr, 6).number_format = "0.00%"
    _widths(wsR, [62, 40, 18, 16, 14, 12])

    # ---------- Fonti e avvisi
    _head(wsF, 1, ["Tipo", "Dettaglio"])
    rows = [("Riservatezza", RISERVATO)]
    rows += [("Avviso", m) for m in res.messages]
    rows += [("Voce esclusa", f"{i}: {m}") for i, m in res.excluded]
    rows += [("Quotazione", w) for w in meta.get("quote_warnings", [])]
    rows += [("Validità quotazione", meta.get("validity", ""))]
    rows += [("Cambi usati", ", ".join(f"{k}={v}" for k, v in meta.get("fx", {}).items()) or "nessuno (solo EUR)")]
    rows += [("Arrotondamento", meta.get("rounding", ""))]
    if res.years is None:
        rows += [("Anni", res.years_note)]
    for i, (a, b) in enumerate(rows, 2):
        wsF.cell(i, 1, a); wsF.cell(i, 2, b)
    _widths(wsF, [24, 120])
    for ws in wb.worksheets:
        ws.freeze_panes = "A2" if ws.title != "Riepilogo" else None
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
