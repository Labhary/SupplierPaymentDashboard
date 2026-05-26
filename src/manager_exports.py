from __future__ import annotations

import json
from datetime import datetime
from io import BytesIO

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from src.metrics import currency
from src.ui import (
    _dataframe_to_excel_bytes,
    format_money_columns,
    prepare_forecast_export,
    prepare_invoice_display,
    prepare_invoice_export,
)


def _date_text(value: object) -> str:
    if pd.isna(value):
        return ""
    timestamp = pd.to_datetime(value, errors="coerce")
    if pd.isna(timestamp):
        return ""
    return timestamp.strftime("%Y-%m-%d")


def _manager_records(invoices: pd.DataFrame) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    for _, row in invoices.iterrows():
        records.append(
            {
                "supplier": str(row.get("Nom du fournisseur", "")),
                "reference": str(row.get("Numéro réf. fournisseur", "")),
                "due_date": _date_text(row.get("Date d'échéance")),
                "amount": float(row.get("Montant", 0) or 0),
                "accounting_date": _date_text(row.get("Date comptable")),
                "code": str(row.get("Code Affaire", "")),
                "project": str(row.get("Nom Projet", "")),
                "source_month": str(row.get("Mois source", "")),
                "month": str(row.get("Mois", "")),
                "year": str(row.get("Année", "")),
                "project_code": str(row.get("Projet / Code Affaire", "")),
            }
        )
    return records


def manager_chart_datasets(records: list[dict[str, object]]) -> dict[str, list[dict[str, object]]]:
    def aggregate(key: str) -> list[dict[str, object]]:
        totals: dict[str, float] = {}
        for record in records:
            label = str(record.get(key) or "")
            if label:
                totals[label] = totals.get(label, 0.0) + float(record.get("amount") or 0)
        return [{"label": label, "amount": amount} for label, amount in totals.items()]

    monthly = sorted(aggregate("month"), key=lambda item: item["label"])
    suppliers = sorted(aggregate("supplier"), key=lambda item: item["amount"], reverse=True)[:10]
    projects = sorted(aggregate("project_code"), key=lambda item: item["amount"], reverse=True)[:10]
    return {"monthly": monthly, "suppliers": suppliers, "projects": projects}


def manager_supplier_detail(records: list[dict[str, object]], supplier: str | None = None) -> dict[str, object]:
    rows = [row for row in records if not supplier or row.get("supplier") == supplier]
    total = sum(float(row.get("amount") or 0) for row in rows)
    suppliers = {str(row.get("supplier") or "") for row in rows if row.get("supplier")}
    projects = {str(row.get("project_code") or "") for row in rows if row.get("project_code")}
    amounts = [float(row.get("amount") or 0) for row in rows]

    if supplier:
        project_totals: dict[str, float] = {}
        for row in rows:
            project = str(row.get("project_code") or "")
            project_totals[project] = project_totals.get(project, 0.0) + float(row.get("amount") or 0)
        latest_due = max([str(row.get("due_date") or "") for row in rows], default="")
        latest_accounting = max([str(row.get("accounting_date") or "") for row in rows], default="")
        top_project = max(project_totals.items(), key=lambda item: item[1])[0] if project_totals else "N/A"
        return {
            "supplier": supplier,
            "invoice_count": len(rows),
            "total_amount": total,
            "average_invoice": total / len(rows) if rows else 0.0,
            "top_project": top_project,
            "latest_due_date": latest_due,
            "latest_accounting_date": latest_accounting,
        }

    supplier_totals: dict[str, float] = {}
    for row in rows:
        supplier_name = str(row.get("supplier") or "")
        supplier_totals[supplier_name] = supplier_totals.get(supplier_name, 0.0) + float(row.get("amount") or 0)
    top5 = sum(sorted(supplier_totals.values(), reverse=True)[:5])
    return {
        "supplier_count": len(suppliers),
        "average_invoice": total / len(rows) if rows else 0.0,
        "highest_invoice": max(amounts) if amounts else 0.0,
        "project_count": len(projects),
        "supplier_concentration": top5 / total if total else 0.0,
    }


def manager_dynamic_insights(records: list[dict[str, object]]) -> list[dict[str, str]]:
    if not records:
        return [{"level": "INFO", "text": "Aucune donnée disponible dans le périmètre exporté."}]

    datasets = manager_chart_datasets(records)
    total = sum(float(row.get("amount") or 0) for row in records)
    avg = total / len(records)
    insights: list[dict[str, str]] = []
    if datasets["monthly"]:
        month = max(datasets["monthly"], key=lambda item: item["amount"])
        share = month["amount"] / total if total else 0
        level = "CRITICAL" if share >= 0.45 else "WARNING" if share >= 0.30 else "INFO"
        insights.append({"level": level, "text": f"{month['label']} concentre {share:.0%} du besoin de décaissement."})
    if datasets["suppliers"]:
        top5 = sum(item["amount"] for item in datasets["suppliers"][:5])
        share = top5 / total if total else 0
        level = "CRITICAL" if share >= 0.70 else "WARNING" if share >= 0.50 else "INFO"
        insights.append({"level": level, "text": f"Les 5 premiers fournisseurs représentent {share:.0%} des engagements."})
    if datasets["projects"]:
        project = max(datasets["projects"], key=lambda item: item["amount"])
        share = project["amount"] / total if total else 0
        level = "WARNING" if share >= 0.35 else "INFO"
        insights.append({"level": level, "text": f"{project['label']} est le projet / code affaire dominant ({share:.0%})."})

    large = [row for row in records if float(row.get("amount") or 0) >= avg * 2 and avg > 0]
    if large:
        insights.append({"level": "WARNING", "text": f"{len(large)} facture(s) dépassent deux fois le montant moyen."})
    else:
        insights.append({"level": "INFO", "text": "Aucune facture atypiquement élevée détectée dans la sélection."})
    return insights


def manager_records_to_csv(records: list[dict[str, object]]) -> str:
    columns = ["supplier", "reference", "due_date", "amount", "accounting_date", "code", "project", "month"]
    lines = [";".join(columns)]
    for record in records:
        values = [str(record.get(column, "")).replace(";", ",") for column in columns]
        lines.append(";".join(values))
    return "\n".join(lines)


def _json_for_html_script(payload: dict[str, object]) -> str:
    text = json.dumps(payload, ensure_ascii=False)
    return (
        text.replace("&", "\\u0026")
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("\u2028", "\\u2028")
        .replace("\u2029", "\\u2029")
    )


def build_interactive_html_report(
    filters: dict[str, object],
    invoices: pd.DataFrame,
    kpis: dict[str, str],
    executive_metrics: dict[str, object],
    insights: list[str],
) -> bytes:
    records = _manager_records(invoices)
    top_project = executive_metrics["most_expensive_project"]
    payload = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "filters": format_filters_summary(filters),
        "records": records,
        "datasets": manager_chart_datasets(records),
        "initial_kpis": {
            "total": kpis.get("Total à payer", "0.00 MAD"),
            "invoice_count": kpis.get("Nombre de factures", "0"),
            "supplier_count": kpis.get("Nombre de fournisseurs", "0"),
            "highest_month": kpis.get("Mois le plus élevé", "N/A").replace("\n", " - "),
            "top_project": f"{top_project['label']} - {currency(top_project['amount'])}",
        },
        "insights": insights,
        "dynamic_insights": manager_dynamic_insights(records),
    }
    payload_json = _json_for_html_script(payload)
    html = """<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Synthèse des décaissements fournisseurs</title>
  <style>
    :root{color-scheme:dark;--bg:#101522;--panel:#182131;--panel2:#111a29;--line:#2a384b;--text:#f8fafc;--muted:#b8c3d4;--blue:#2f6fed;--teal:#20a39e;--amber:#f59e0b;--red:#ef4444}
    *{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font-family:Segoe UI,Arial,sans-serif}
    main{max-width:1340px;margin:auto;padding:26px}header{display:flex;justify-content:space-between;gap:20px;margin-bottom:20px}
    h1{margin:0 0 6px;font-size:28px}h2{margin:0 0 14px;font-size:18px}.subtitle,.muted{color:var(--muted);font-size:13px}
    .brand{padding:18px 20px;border:1px solid var(--line);border-radius:8px;background:#172033;box-shadow:0 10px 30px rgba(0,0,0,.22)}
    .panel,.card{background:var(--panel);border:1px solid var(--line);border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.18)}
    .panel{padding:16px;margin-bottom:16px}.filters{display:grid;grid-template-columns:repeat(5,minmax(140px,1fr));gap:12px;align-items:end}
    label{display:block;color:var(--muted);font-size:11px;font-weight:800;text-transform:uppercase;margin-bottom:6px}
    select,input{width:100%;background:#0b1220;color:var(--text);border:1px solid var(--line);border-radius:7px;padding:10px;outline:none}
    select option{white-space:normal}select:focus,input:focus{border-color:var(--blue);box-shadow:0 0 0 2px rgba(47,111,237,.18)}
    button{background:var(--blue);color:white;border:0;border-radius:7px;padding:10px 14px;font-weight:800;cursor:pointer}button.secondary{background:#24324b}
    .chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px}.chip{background:#0b1220;border:1px solid var(--line);border-radius:999px;padding:6px 10px;color:var(--muted);font-size:12px}
    .kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:12px}.card{padding:15px;min-height:104px}.label{color:var(--muted);font-size:11px;font-weight:800;text-transform:uppercase}.value{font-size:21px;font-weight:850;margin-top:8px;line-height:1.18}
    .grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}.chart-box{height:360px}.mini-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}.mini{background:#0b1220;border:1px solid var(--line);border-radius:8px;padding:11px}
    .insights{display:grid;grid-template-columns:repeat(2,1fr);gap:10px}.insight{background:var(--panel2);border:1px solid var(--line);border-left:5px solid var(--blue);border-radius:8px;padding:12px}.insight.WARNING{border-left-color:var(--amber)}.insight.CRITICAL{border-left-color:var(--red)}
    .badge{display:inline-block;font-size:10px;font-weight:900;border-radius:999px;padding:4px 8px;margin-bottom:6px;background:rgba(47,111,237,.18);color:#bfdbfe}.WARNING .badge{background:rgba(245,158,11,.18);color:#fde68a}.CRITICAL .badge{background:rgba(239,68,68,.18);color:#fecaca}
    .table-tools{display:grid;grid-template-columns:1fr 130px 150px;gap:10px;margin-bottom:10px}.table-wrap{max-height:460px;overflow:auto;border:1px solid var(--line);border-radius:8px}
    .filter-note{display:none;color:#fde68a;margin-top:10px}.no-results{display:none;background:#0b1220;border:1px solid var(--line);border-radius:8px;padding:16px;margin-bottom:10px;color:var(--muted)}
    table{width:100%;border-collapse:collapse;font-size:12px}th,td{padding:9px;border-bottom:1px solid var(--line);text-align:left}th{position:sticky;top:0;background:#111827;color:var(--muted);cursor:pointer;z-index:1}tbody tr:nth-child(even){background:rgba(255,255,255,.025)}tbody tr:hover{background:rgba(47,111,237,.12)}
    .pager{display:flex;justify-content:space-between;align-items:center;margin-top:10px}.runtime-error{display:none;background:rgba(239,68,68,.14);border:1px solid rgba(239,68,68,.55);border-left:5px solid var(--red);border-radius:8px;padding:14px;margin-bottom:16px;color:#fecaca;white-space:pre-wrap}
    @media(max-width:1050px){.kpis,.filters,.grid2,.insights{grid-template-columns:1fr}.table-tools{grid-template-columns:1fr}header{flex-direction:column}}
  </style>
</head>
<body>
<main>
  <section id="runtimeError" class="runtime-error"></section>
  <header><div class="brand"><h1>Synthèse des décaissements fournisseurs</h1><div class="subtitle">Synthèse interactive locale, en lecture seule, basée sur le périmètre exporté.</div></div><div class="panel"><div class="label">Généré le</div><div id="generated" class="value"></div></div></header>
  <section class="panel"><h2>Filtres actifs à l'export</h2><div id="sourceFilters" class="muted"></div><div id="activeChips" class="chips"></div></section>
  <section class="panel"><div class="filters"><div><label>Mois</label><select id="monthFilter"></select></div><div><label>Fournisseur</label><select id="supplierFilter"></select></div><div><label>Projet</label><select id="projectFilter"></select></div><div><label>Recherche</label><input id="searchBox" placeholder="Référence, fournisseur, projet..."></div><button id="resetBtn">Réinitialiser</button></div><p class="muted" id="resultCount"></p><p class="filter-note" id="filterNotice"></p></section>
  <section class="kpis"><div class="card"><div class="label">Total à payer</div><div class="value" id="kpiTotal"></div></div><div class="card"><div class="label">Nombre de factures</div><div class="value" id="kpiInvoices"></div></div><div class="card"><div class="label">Nombre de fournisseurs</div><div class="value" id="kpiSuppliers"></div></div><div class="card"><div class="label">Mois le plus élevé</div><div class="value" id="kpiMonth"></div></div><div class="card"><div class="label">Projet / code affaire principal</div><div class="value" id="kpiProject"></div></div></section>
  <section class="grid2"><div class="panel"><h2>Prévision mensuelle</h2><div class="chart-box"><canvas id="monthlyChart"></canvas></div></div><div class="panel"><h2>Top fournisseurs</h2><div class="chart-box"><canvas id="supplierChart"></canvas></div></div></section>
  <section class="grid2"><div class="panel"><h2>Projet / code affaire</h2><div class="chart-box"><canvas id="projectChart"></canvas></div></div><div class="panel"><h2>Détail fournisseur</h2><div id="supplierDetail" class="mini-grid"></div></div></section>
  <section class="panel"><h2>Insights exécutifs</h2><div id="insights" class="insights"></div></section>
  <section class="panel"><h2>Factures filtrées</h2><div id="tableChips" class="chips"></div><div class="table-tools"><input id="tableSearch" placeholder="Rechercher dans la table filtrée"><select id="pageSize"><option>25</option><option>50</option><option>100</option></select><button class="secondary" id="csvBtn">Exporter CSV</button></div><div id="noResults" class="no-results"><p>Aucune facture ne correspond aux filtres sélectionnés.</p><button class="secondary" id="noResultsReset">Réinitialiser les filtres</button></div><div class="table-wrap"><table id="detailTable"></table></div><div class="pager"><button class="secondary" id="prevPage">Précédent</button><span class="muted" id="pageInfo"></span><button class="secondary" id="nextPage">Suivant</button></div></section>
</main>
<script id="report-data" type="application/json">__PAYLOAD__</script>
<script>
const el = {runtimeError:document.getElementById('runtimeError'),generated:document.getElementById('generated'),sourceFilters:document.getElementById('sourceFilters'),activeChips:document.getElementById('activeChips'),monthFilter:document.getElementById('monthFilter'),supplierFilter:document.getElementById('supplierFilter'),projectFilter:document.getElementById('projectFilter'),searchBox:document.getElementById('searchBox'),tableSearch:document.getElementById('tableSearch'),pageSize:document.getElementById('pageSize'),resultCount:document.getElementById('resultCount'),filterNotice:document.getElementById('filterNotice'),tableChips:document.getElementById('tableChips'),noResults:document.getElementById('noResults'),noResultsReset:document.getElementById('noResultsReset'),kpiTotal:document.getElementById('kpiTotal'),kpiInvoices:document.getElementById('kpiInvoices'),kpiSuppliers:document.getElementById('kpiSuppliers'),kpiMonth:document.getElementById('kpiMonth'),kpiProject:document.getElementById('kpiProject'),supplierDetail:document.getElementById('supplierDetail'),insights:document.getElementById('insights'),detailTable:document.getElementById('detailTable'),pageInfo:document.getElementById('pageInfo'),resetBtn:document.getElementById('resetBtn'),prevPage:document.getElementById('prevPage'),nextPage:document.getElementById('nextPage'),csvBtn:document.getElementById('csvBtn')};
let data = null
let records = []
let sortKey = 'due_date'
let sortDir = 1
let page = 1
function showRuntimeError(error){const message=error&&error.message?error.message:String(error);if(el.runtimeError){el.runtimeError.style.display='block';el.runtimeError.textContent='Erreur lors du chargement du rapport interactif.\\n'+message}console.error(error)}
window.onerror=function(message,source,line,column,error){showRuntimeError(error||message);return false};
window.onunhandledrejection=function(event){showRuntimeError(event.reason||event)};
function parsePayload(){const payload=document.getElementById('report-data');data=JSON.parse(payload.textContent);records=Array.isArray(data.records)?data.records:[]}
function money(v){const amount=Number(v)||0;return new Intl.NumberFormat('fr-FR',{minimumFractionDigits:2,maximumFractionDigits:2}).format(amount)+' MAD'}
function unique(values){const seen={};const out=[];values.forEach(function(value){const label=String(value||'');if(label&&!seen[label]){seen[label]=true;out.push(label)}});return out.sort()}
function sumBy(rows,key){const totals={};rows.forEach(function(row){const label=String(row[key]||'');if(label){totals[label]=(totals[label]||0)+Number(row.amount||0)}});return totals}
function topEntries(totals, limit){const entries=[];Object.keys(totals).forEach(function(key){entries.push([key,totals[key]])});entries.sort(function(a,b){return b[1]-a[1]});return entries.slice(0,limit)}
function fillSelect(select, values){const selected=select['value'];select['innerHTML']='';const all=document.createElement('option');all.value='';all.textContent='Tous';select.appendChild(all);values.forEach(function(value){const option=document.createElement('option');option.value=value;option.textContent=value;option.title=value;select.appendChild(option)});if(selected&&values.indexOf(selected)!==-1){select['value']=selected}}
function recordText(record){const values=[];Object.keys(record).forEach(function(key){values.push(record[key])});return values.join(' ')}
function searchMatches(record){const text=recordText(record).toLowerCase();const mainSearch=el.searchBox['value'].toLowerCase();const tableSearch=el.tableSearch['value'].toLowerCase();return(!mainSearch||text.indexOf(mainSearch)!==-1)&&(!tableSearch||text.indexOf(tableSearch)!==-1)}
function compatibleRows(skipKey){const month=skipKey==='month'?'':el.monthFilter['value'];const supplier=skipKey==='supplier'?'':el.supplierFilter['value'];const project=skipKey==='project'?'':el.projectFilter['value'];return records.filter(function(record){return(!month||record.month===month)&&(!supplier||record.supplier===supplier)&&(!project||record.project_code===project)&&searchMatches(record)})}
function compatibleValues(key,skipKey){return unique(compatibleRows(skipKey).map(function(record){return record[key]}))}
function clearProjectIfIncompatible(){const value=el.projectFilter['value'];const projects=compatibleValues('project_code','project');if(value&&projects.indexOf(value)===-1){el.projectFilter['value']='';return true}return false}
function clearSupplierIfIncompatible(){const value=el.supplierFilter['value'];const suppliers=compatibleValues('supplier','supplier');if(value&&suppliers.indexOf(value)===-1){el.supplierFilter['value']='';return true}return false}
function clearMonthDependentFilters(){const supplierChanged=clearSupplierIfIncompatible();const projectChanged=clearProjectIfIncompatible();return supplierChanged||projectChanged}
function clearIncompatibleFilters(){const supplierChanged=clearSupplierIfIncompatible();const projectChanged=clearProjectIfIncompatible();return supplierChanged||projectChanged}
function updateCascadingFilters(showNotice){const adjusted=clearIncompatibleFilters();fillSelect(el.monthFilter,compatibleValues('month','month'));fillSelect(el.supplierFilter,compatibleValues('supplier','supplier'));fillSelect(el.projectFilter,compatibleValues('project_code','project'));if(el.filterNotice){el.filterNotice.style.display=adjusted&&showNotice?'block':'none';el.filterNotice.textContent=adjusted?'Certains filtres incompatibles ont été réinitialisés.':''}}
function resetFilters(){['monthFilter','supplierFilter','projectFilter','searchBox','tableSearch'].forEach(function(id){el[id]['value']=''});page=1;updateCascadingFilters(false);render()}
function activeRows(){const month=el.monthFilter['value'];const supplier=el.supplierFilter['value'];const project=el.projectFilter['value'];return records.filter(function(record){return(!month||record.month===month)&&(!supplier||record.supplier===supplier)&&(!project||record.project_code===project)&&searchMatches(record)})}
function dataset(rows, key, limit, asc){const items=topEntries(sumBy(rows,key),limit).map(function(entry){return{label:entry[0],amount:entry[1]}});if(asc){items.sort(function(a,b){return String(a.label).localeCompare(String(b.label))})}return items}
function roundRect(ctx,x,y,w,h,r){const rr=Math.min(r,Math.abs(w)/2,Math.abs(h)/2);ctx.beginPath();ctx.moveTo(x+rr,y);ctx.lineTo(x+w-rr,y);ctx.quadraticCurveTo(x+w,y,x+w,y+rr);ctx.lineTo(x+w,y+h-rr);ctx.quadraticCurveTo(x+w,y+h,x+w-rr,y+h);ctx.lineTo(x+rr,y+h);ctx.quadraticCurveTo(x,y+h,x,y+h-rr);ctx.lineTo(x,y+rr);ctx.quadraticCurveTo(x,y,x+rr,y);ctx.closePath()}
function drawChart(canvasId, items, horizontal){const canvas=document.getElementById(canvasId);const box=canvas.parentElement;const ctx=canvas.getContext('2d');const dpr=window.devicePixelRatio||1;const w=box.clientWidth||620;const h=box.clientHeight||340;canvas.width=w*dpr;canvas.height=h*dpr;canvas.style.width=w+'px';canvas.style.height=h+'px';ctx.setTransform(dpr,0,0,dpr,0,0);ctx.clearRect(0,0,w,h);ctx.font='12px Segoe UI, Arial';ctx.fillStyle='#aebbd0';ctx.textBaseline='middle';if(!items.length){ctx.fillText('Aucune donnée à afficher.',18,30);return}let max=1;items.forEach(function(item){max=Math.max(max,Number(item.amount)||0)});const blue='#2f6fed';const teal='#20a39e';const grid='rgba(255,255,255,.08)';const line='#273449';if(horizontal){const left=150;const right=118;const topPad=18;const rowH=Math.max(27,Math.min(42,(h-topPad-18)/items.length));const barMax=Math.max(20,w-left-right);ctx.strokeStyle=grid;ctx.lineWidth=1;for(let i=0;i<=4;i++){const gx=left+barMax*i/4;ctx.beginPath();ctx.moveTo(gx,topPad-4);ctx.lineTo(gx,h-14);ctx.stroke()}items.forEach(function(item,index){const y=topPad+index*rowH;const barW=barMax*((Number(item.amount)||0)/max);const grad=ctx.createLinearGradient(left,0,left+barW,0);grad.addColorStop(0,blue);grad.addColorStop(1,teal);ctx.fillStyle='#aebbd0';ctx.textAlign='right';ctx.fillText(String(item.label).slice(0,22),left-10,y+rowH/2);ctx.fillStyle=grad;roundRect(ctx,left,y+7,barW,Math.max(10,rowH-15),6);ctx.fill();ctx.fillStyle='#f8fafc';ctx.textAlign='left';ctx.fillText(money(item.amount),left+barW+8,y+rowH/2)})}else{const left=54;const right=18;const topPad=20;const bottom=68;const plotW=w-left-right;const plotH=h-topPad-bottom;const barGap=10;const barW=Math.max(12,(plotW/items.length)-barGap);ctx.strokeStyle=grid;ctx.lineWidth=1;ctx.textAlign='right';ctx.fillStyle='#aebbd0';for(let step=0;step<=4;step++){const gy=topPad+plotH-plotH*step/4;ctx.beginPath();ctx.moveTo(left,gy);ctx.lineTo(w-right,gy);ctx.stroke();ctx.fillText(step===0?'':money(max*step/4).replace(' MAD',''),left-6,gy+4)}ctx.strokeStyle=line;ctx.beginPath();ctx.moveTo(left,topPad);ctx.lineTo(left,topPad+plotH);ctx.lineTo(w-right,topPad+plotH);ctx.stroke();items.forEach(function(item,index){const x=left+index*(barW+barGap)+barGap/2;const barH=plotH*((Number(item.amount)||0)/max);const y=topPad+plotH-barH;const grad=ctx.createLinearGradient(0,y,0,topPad+plotH);grad.addColorStop(0,teal);grad.addColorStop(1,blue);ctx.fillStyle=grad;roundRect(ctx,x,y,barW,barH,6);ctx.fill();ctx.save();ctx.translate(x+barW/2,topPad+plotH+12);ctx.rotate(-Math.PI/5);ctx.fillStyle='#aebbd0';ctx.textAlign='right';ctx.fillText(String(item.label).slice(0,16),0,0);ctx.restore()})}}
function computeInsights(rows){const total=rows.reduce(function(acc,row){return acc+Number(row.amount||0)},0);const avg=total/(rows.length||1);const out=[];const month=topEntries(sumBy(rows,'month'),1)[0];if(month){const monthShare=month[1]/(total||1);out.push({level:monthShare>=0.45?'CRITICAL':monthShare>=0.30?'WARNING':'INFO',text:month[0]+' concentre '+Math.round(monthShare*100)+'% du besoin de décaissement.'})}const supplierTop=topEntries(sumBy(rows,'supplier'),5).reduce(function(acc,item){return acc+item[1]},0);const supplierShare=supplierTop/(total||1);out.push({level:supplierShare>=0.70?'CRITICAL':supplierShare>=0.50?'WARNING':'INFO',text:'Les 5 premiers fournisseurs représentent '+Math.round(supplierShare*100)+'% des engagements.'});const project=topEntries(sumBy(rows,'project_code'),1)[0];if(project){out.push({level:project[1]/(total||1)>=0.35?'WARNING':'INFO',text:project[0]+' est le projet / code affaire dominant.'})}const large=rows.filter(function(row){return Number(row.amount||0)>avg*2});out.push({level:large.length?'WARNING':'INFO',text:large.length?large.length+' facture(s) dépassent deux fois le montant moyen.':'Aucune facture atypiquement élevée détectée.'});return out}
function renderInsights(rows){el.insights['innerHTML']='';computeInsights(rows).forEach(function(item){const card=document.createElement('div');card.className='insight '+item.level;const badge=document.createElement('div');badge.className='badge';badge.textContent=item.level;const text=document.createElement('div');text.textContent=item.text;card.appendChild(badge);card.appendChild(text);el.insights.appendChild(card)})}
function renderKpis(rows){const total=rows.reduce(function(acc,row){return acc+Number(row.amount||0)},0);const month=topEntries(sumBy(rows,'month'),1)[0]||['N/A',0];const project=topEntries(sumBy(rows,'project_code'),1)[0]||['N/A',0];el.kpiTotal.textContent=money(total);el.kpiInvoices.textContent=String(rows.length);el.kpiSuppliers.textContent=String(unique(rows.map(function(row){return row.supplier})).length);el.kpiMonth['innerHTML']='';el.kpiMonth.appendChild(document.createTextNode(month[0]));el.kpiMonth.appendChild(document.createElement('br'));el.kpiMonth.appendChild(document.createTextNode(money(month[1])));el.kpiProject['innerHTML']='';el.kpiProject.appendChild(document.createTextNode(project[0]));el.kpiProject.appendChild(document.createElement('br'));el.kpiProject.appendChild(document.createTextNode(money(project[1])))}
function renderSupplier(rows){const selected=el.supplierFilter['value'];const base=selected?rows.filter(function(row){return row.supplier===selected}):rows;const total=base.reduce(function(acc,row){return acc+Number(row.amount||0)},0);const avg=total/(base.length||1);let cards=[];if(selected){const firstProject=topEntries(sumBy(base,'project_code'),1)[0];const dates=base.map(function(row){return row.due_date||''}).sort();cards=[['Fournisseur',selected],['Factures',base.length],['Montant total',money(total)],['Facture moyenne',money(avg)],['Top projet',firstProject?firstProject[0]:'N/A'],['Dernière échéance',dates.length?dates[dates.length-1]:'']]}else{const amounts=base.map(function(row){return Number(row.amount||0)});let high=0;amounts.forEach(function(amount){high=Math.max(high,amount)});const top5=topEntries(sumBy(base,'supplier'),5).reduce(function(acc,item){return acc+item[1]},0);cards=[['Fournisseurs',unique(base.map(function(row){return row.supplier})).length],['Facture moyenne',money(avg)],['Plus forte facture',money(high)],['Projets',unique(base.map(function(row){return row.project_code})).length],['Concentration top 5',Math.round(top5/(total||1)*100)+'%']]}el.supplierDetail['innerHTML']='';cards.forEach(function(card){const box=document.createElement('div');box.className='mini';const label=document.createElement('div');label.className='label';label.textContent=card[0];const value=document.createElement('div');value.className='value';value.textContent=String(card[1]);box.appendChild(label);box.appendChild(value);el.supplierDetail.appendChild(box)})}
function renderFilterChips(){const chips=[];if(el.monthFilter['value']){chips.push('Mois: '+el.monthFilter['value'])}if(el.supplierFilter['value']){chips.push('Fournisseur: '+el.supplierFilter['value'])}if(el.projectFilter['value']){chips.push('Projet: '+el.projectFilter['value'])}if(el.searchBox['value']){chips.push('Recherche: '+el.searchBox['value'])}if(el.tableSearch['value']){chips.push('Table: '+el.tableSearch['value'])}if(!chips.length){chips.push('Aucun filtre interne')}[el.activeChips,el.tableChips].forEach(function(container){container['innerHTML']='';chips.forEach(function(value){const chip=document.createElement('span');chip.className='chip';chip.textContent=value;container.appendChild(chip)})})}
function renderTable(rows){const list=rows.slice();list.sort(function(a,b){let av=a[sortKey];let bv=b[sortKey];if(sortKey==='amount'){av=Number(av)||0;bv=Number(bv)||0}else{av=String(av||'');bv=String(bv||'')}return av>bv?sortDir:av<bv?-sortDir:0});const size=Number(el.pageSize['value'])||25;const pages=Math.max(1,Math.ceil(list.length/size));page=Math.min(page,pages);const pageRows=list.slice((page-1)*size,page*size);const columns=[['supplier','Fournisseur'],['reference','Référence'],['due_date','Échéance'],['amount','Montant'],['accounting_date','Date comptable'],['code','Code'],['project','Projet'],['month','Mois']];el.detailTable['innerHTML']='';const thead=document.createElement('thead');const headRow=document.createElement('tr');columns.forEach(function(column){const th=document.createElement('th');th.setAttribute('data-key',column[0]);th.textContent=column[1];th.onclick=function(){const key=th.getAttribute('data-key');sortDir=sortKey===key?-sortDir:1;sortKey=key;render()};headRow.appendChild(th)});thead.appendChild(headRow);const tbody=document.createElement('tbody');pageRows.forEach(function(row){const tr=document.createElement('tr');columns.forEach(function(column){const td=document.createElement('td');td.textContent=column[0]==='amount'?money(row.amount):String(row[column[0]]||'');tr.appendChild(td)});tbody.appendChild(tr)});el.detailTable.appendChild(thead);el.detailTable.appendChild(tbody);el.pageInfo.textContent='Page '+page+' / '+pages+' - '+list.length+' ligne(s)'}
function exportCsv(){const rows=activeRows();const cols=['supplier','reference','due_date','amount','accounting_date','code','project','month'];const lines=[cols.join(';')];rows.forEach(function(row){const values=cols.map(function(col){return String(row[col]===undefined||row[col]===null?'':row[col]).replace(/;/g,',')});lines.push(values.join(';'))});const link=document.createElement('a');link.href=URL.createObjectURL(new Blob([lines.join('\\n')],{type:'text/csv;charset=utf-8'}));link.download='factures_filtrees_manager.csv';link.click()}
function render(){console.log('render start');document.body.classList.add('updating');const rows=activeRows();el.resultCount['textContent']=rows.length+' facture(s) dans la vue active.';renderFilterChips();if(el.noResults){el.noResults.style.display=rows.length?'none':'block'}renderKpis(rows);drawChart('monthlyChart',dataset(rows,'month',24,true),false);drawChart('supplierChart',dataset(rows,'supplier',10,false),true);drawChart('projectChart',dataset(rows,'project_code',10,false),true);renderSupplier(rows);renderInsights(rows);renderTable(rows);setTimeout(function(){document.body.classList.remove('updating')},120);console.log('render success')}
function init(){parsePayload();console.log('data loaded');console.log('row count',records.length);el.generated['textContent']=data.generated_at||'';el.sourceFilters['innerHTML']='';(data.filters||[]).forEach(function(filter){const item=document.createElement('div');item.textContent=filter;el.sourceFilters.appendChild(item)});fillSelect(el.monthFilter,unique(records.map(function(record){return record.month})));fillSelect(el.supplierFilter,unique(records.map(function(record){return record.supplier})));fillSelect(el.projectFilter,unique(records.map(function(record){return record.project_code})));el.monthFilter.addEventListener('input',function(){page=1;clearMonthDependentFilters();updateCascadingFilters(true);render()});el.supplierFilter.addEventListener('input',function(){page=1;clearProjectIfIncompatible();updateCascadingFilters(true);render()});el.projectFilter.addEventListener('input',function(){page=1;clearSupplierIfIncompatible();updateCascadingFilters(true);render()});['searchBox','tableSearch','pageSize'].forEach(function(id){el[id].addEventListener('input',function(){page=1;updateCascadingFilters(true);render()})});el.resetBtn['onclick']=resetFilters;el.noResultsReset['onclick']=resetFilters;el.prevPage.onclick=function(){page=Math.max(1,page-1);render()};el.nextPage.onclick=function(){page=page+1;render()};el.csvBtn['onclick']=exportCsv;window.addEventListener('resize',render);updateCascadingFilters(false);render()}
function start(){try{init()}catch(error){showRuntimeError(error)}}
if(document.readyState==='loading'){document.addEventListener('DOMContentLoaded',start,{once:true})}else{start()}
</script>
</body></html>"""
    return html.replace("__PAYLOAD__", payload_json).encode("utf-8")


def format_filters_summary(filters: dict[str, object]) -> list[str]:
    labels = [
        ("Fournisseur", filters.get("suppliers") or []),
        ("Code Affaire", filters.get("codes") or []),
        ("Nom Projet", filters.get("project_names") or []),
        ("Mois calculé", filters.get("months") or []),
    ]
    lines = [f"{label}: {', '.join(values)}" for label, values in labels if values]
    min_amount = float(filters.get("min_amount") or 0)
    if min_amount > 0:
        lines.append(f"Montant minimum: {currency(min_amount)}")
    return lines or ["Aucun filtre actif"]


def build_manager_excel(
    invoices: pd.DataFrame,
    forecast: pd.DataFrame,
    top_suppliers: pd.DataFrame,
) -> bytes:
    return _dataframe_to_excel_bytes(
        {
            "Factures filtrées": prepare_invoice_export(invoices),
            "Prévision mensuelle": prepare_forecast_export(forecast),
            "Top fournisseurs": top_suppliers.copy(),
        }
    )


def _styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "ManagerTitle",
            parent=base["Title"],
            fontName="Helvetica-Bold",
            fontSize=18,
            alignment=TA_CENTER,
            spaceAfter=12,
        ),
        "heading": ParagraphStyle(
            "ManagerHeading",
            parent=base["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=12,
            textColor=colors.HexColor("#172033"),
            spaceBefore=8,
            spaceAfter=6,
        ),
        "body": ParagraphStyle("ManagerBody", parent=base["BodyText"], fontSize=9, leading=12),
    }


def _chart_image(kind: str, data: pd.DataFrame) -> BytesIO | None:
    if data.empty:
        return None

    fig, ax = plt.subplots(figsize=(8.2, 3.8), dpi=150)
    color = "#2f6fed"
    accent = "#20a39e"

    if kind == "monthly":
        ax.bar(data["Mois"], data["Montant"], color=color)
        ax.set_title("Décaissements prévisionnels par mois")
        ax.set_ylabel("MAD")
        ax.tick_params(axis="x", rotation=35)
    elif kind == "suppliers":
        plot_df = data.sort_values("Montant").tail(10)
        ax.barh(plot_df["Nom du fournisseur"], plot_df["Montant"], color=color)
        ax.set_title("Top fournisseurs")
        ax.set_xlabel("MAD")
    elif kind == "projects":
        plot_df = data.sort_values("Montant").tail(10)
        ax.barh(plot_df["Projet / Code Affaire"], plot_df["Montant"], color=accent)
        ax.set_title("Montant par projet / code affaire")
        ax.set_xlabel("MAD")
    elif kind == "pareto":
        plot_df = data.sort_values("Montant", ascending=False).head(10)
        ax.bar(plot_df["Nom du fournisseur"], plot_df["Contribution"] * 100, color=color)
        ax.set_title("Contribution fournisseurs")
        ax.set_ylabel("Contribution (%)")
        ax.tick_params(axis="x", rotation=40)
    else:
        plt.close(fig)
        return None

    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    output = BytesIO()
    fig.savefig(output, format="png", bbox_inches="tight")
    plt.close(fig)
    output.seek(0)
    return output


def _add_chart(story: list, kind: str, data: pd.DataFrame, width: float = 17 * cm) -> None:
    image = _chart_image(kind, data)
    if image is not None:
        story.append(Image(image, width=width, height=7.5 * cm))
        story.append(Spacer(1, 0.25 * cm))


def _table_from_dataframe(df: pd.DataFrame, max_rows: int = 10) -> Table:
    display = df.head(max_rows).copy()
    values = [display.columns.tolist()] + display.astype(str).values.tolist()
    table = Table(values, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#172033")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
            ]
        )
    )
    return table


def _build_pdf(story: list, page_size=A4) -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=page_size,
        rightMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm,
    )
    doc.build(story)
    return buffer.getvalue()


def build_executive_pdf(
    filters: dict[str, object],
    kpis: dict[str, str],
    executive_metrics: dict[str, object],
    insights: list[str],
    forecast: pd.DataFrame,
    top_suppliers: pd.DataFrame,
) -> bytes:
    styles = _styles()
    story = [Paragraph("Synthèse exécutive - Décaissements fournisseurs", styles["title"])]

    story.append(Paragraph("Filtres sélectionnés", styles["heading"]))
    for line in format_filters_summary(filters):
        story.append(Paragraph(line, styles["body"]))

    story.append(Paragraph("Indicateurs clés", styles["heading"]))
    top_project = executive_metrics["most_expensive_project"]
    rows = [
        ["Total à payer", kpis.get("Total à payer", "N/A")],
        ["Nombre de factures", kpis.get("Nombre de factures", "0")],
        ["Nombre de fournisseurs", kpis.get("Nombre de fournisseurs", "0")],
        ["Mois le plus élevé", kpis.get("Mois le plus élevé", "N/A").replace("\n", " - ")],
        ["Projet / code affaire principal", f"{top_project['label']} - {currency(top_project['amount'])}"],
    ]
    story.append(_table_from_dataframe(pd.DataFrame(rows, columns=["Indicateur", "Valeur"]), max_rows=20))

    story.append(Paragraph("Top 5 fournisseurs", styles["heading"]))
    story.append(_table_from_dataframe(format_money_columns(top_suppliers.head(5), ["Montant"]), max_rows=5))

    story.append(Paragraph("Prévision mensuelle", styles["heading"]))
    _add_chart(story, "monthly", forecast)

    story.append(Paragraph("Insights", styles["heading"]))
    for insight in insights:
        story.append(Paragraph(f"- {insight}", styles["body"]))

    return _build_pdf(story)


def build_visuals_pdf(
    forecast: pd.DataFrame,
    top_suppliers: pd.DataFrame,
    project_amounts: pd.DataFrame,
    pareto: pd.DataFrame,
) -> bytes:
    styles = _styles()
    story = [Paragraph("Visualisations - Décaissements fournisseurs", styles["title"])]
    _add_chart(story, "monthly", forecast)
    _add_chart(story, "suppliers", top_suppliers)
    story.append(PageBreak())
    _add_chart(story, "projects", project_amounts)
    _add_chart(story, "pareto", pareto)
    return _build_pdf(story, page_size=landscape(A4))


def build_full_pdf(
    filters: dict[str, object],
    kpis: dict[str, str],
    executive_metrics: dict[str, object],
    insights: list[str],
    forecast: pd.DataFrame,
    top_suppliers: pd.DataFrame,
    project_amounts: pd.DataFrame,
    pareto: pd.DataFrame,
    invoices: pd.DataFrame,
    max_rows: int = 30,
) -> bytes:
    styles = _styles()
    story = [Paragraph("Rapport complet - Décaissements fournisseurs", styles["title"])]
    story.append(Paragraph("Synthèse exécutive", styles["heading"]))
    top_project = executive_metrics["most_expensive_project"]
    rows = [
        ["Total à payer", kpis.get("Total à payer", "N/A")],
        ["Nombre de factures", kpis.get("Nombre de factures", "0")],
        ["Nombre de fournisseurs", kpis.get("Nombre de fournisseurs", "0")],
        ["Mois le plus élevé", kpis.get("Mois le plus élevé", "N/A").replace("\n", " - ")],
        ["Projet / code affaire principal", f"{top_project['label']} - {currency(top_project['amount'])}"],
    ]
    story.append(_table_from_dataframe(pd.DataFrame(rows, columns=["Indicateur", "Valeur"]), max_rows=20))
    story.append(Paragraph("Filtres", styles["heading"]))
    for line in format_filters_summary(filters):
        story.append(Paragraph(line, styles["body"]))
    story.append(Paragraph("Insights", styles["heading"]))
    for insight in insights:
        story.append(Paragraph(f"- {insight}", styles["body"]))

    story.append(PageBreak())
    story.append(Paragraph("Visualisations", styles["heading"]))
    _add_chart(story, "monthly", forecast)
    _add_chart(story, "suppliers", top_suppliers)
    _add_chart(story, "projects", project_amounts)
    _add_chart(story, "pareto", pareto)

    story.append(PageBreak())
    story.append(Paragraph("Factures filtrées", styles["heading"]))
    display = prepare_invoice_display(invoices)
    if len(display) > max_rows:
        story.append(Paragraph(f"Table limitée aux {max_rows} premières lignes sur {len(display)}.", styles["body"]))
    story.append(_table_from_dataframe(display, max_rows=max_rows))
    return _build_pdf(story, page_size=landscape(A4))
