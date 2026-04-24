"""
title: AI Reconciliation
author: IBM Consulting Advantage
version: 1.0.0
description: AI-assisted reconciliation across Banking, Investment Banking, Insurance, Healthcare, Asset Management, Pharma/Clinical for USA and EU regulatory regimes. Renders inline report and generates downloadable XLSX/DOCX/PPTX (CDN-first with always-on ooXML server-side fallback). Requires iframe Sandbox Allow Same Origin in Open WebUI Settings -> Interface.
"""

from __future__ import annotations

import base64
import html as htmlmod
import io
import json
import logging
import re
import time
import traceback
import uuid

log = logging.getLogger("reconciliation_tool")
log.setLevel(logging.INFO)
from typing import Any, Dict, List, Literal, Optional

from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field

_BUILD = "1.0.0"

# ---------------------------------------------------------------------------
# IBM Light Navy Blue theme
# ---------------------------------------------------------------------------

IBM_NAVY = "#002D74"
IBM_NAVY_DARK = "#001F52"
IBM_BLUE = "#0043CE"
IBM_BLUE_LIGHT = "#4589FF"
IBM_SURFACE = "#EDF5FF"
IBM_SURFACE_2 = "#D0E2FF"
IBM_TEXT = "#161616"
IBM_TEXT_2 = "#525252"
IBM_BORDER = "#C6C6C6"
IBM_SUCCESS = "#24A148"
IBM_WARN = "#F1C21B"
IBM_DANGER = "#DA1E28"

THEME_CSS = f"""
:root {{
  --ibm-navy: {IBM_NAVY};
  --ibm-navy-dark: {IBM_NAVY_DARK};
  --ibm-blue: {IBM_BLUE};
  --ibm-blue-light: {IBM_BLUE_LIGHT};
  --ibm-surface: {IBM_SURFACE};
  --ibm-surface-2: {IBM_SURFACE_2};
  --ibm-text: {IBM_TEXT};
  --ibm-text-2: {IBM_TEXT_2};
  --ibm-border: {IBM_BORDER};
  --ibm-success: {IBM_SUCCESS};
  --ibm-warn: {IBM_WARN};
  --ibm-danger: {IBM_DANGER};
  --font-sans: 'IBM Plex Sans', system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif;
  --font-mono: 'IBM Plex Mono', 'SF Mono', Menlo, Consolas, monospace;
  --radius: 4px;
  --radius-lg: 8px;
  color-scheme: light;
}}
* {{ box-sizing: border-box; }}
html, body {{ margin: 0; padding: 0; background: #FFFFFF; color: var(--ibm-text); font-family: var(--font-sans); font-size: 14px; line-height: 1.5; }}
a {{ color: var(--ibm-blue); }}
.ibm-shell {{ padding: 20px 24px 32px; max-width: 100%; }}
.ibm-masthead {{ display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 14px 20px; background: var(--ibm-navy); color: #FFFFFF; border-radius: var(--radius) var(--radius) 0 0; }}
.ibm-masthead .brand {{ display: flex; align-items: center; gap: 12px; font-weight: 600; letter-spacing: 0.01em; }}
.ibm-masthead .brand .dot {{ width: 10px; height: 10px; background: var(--ibm-blue-light); border-radius: 50%; display: inline-block; }}
.ibm-masthead .meta {{ font-size: 12px; opacity: 0.85; font-family: var(--font-mono); }}
.ibm-subbar {{ display: flex; flex-wrap: wrap; gap: 10px; align-items: center; padding: 10px 20px; background: var(--ibm-surface); border-left: 4px solid var(--ibm-blue); color: var(--ibm-text); font-size: 13px; }}
.ibm-chip {{ display: inline-flex; align-items: center; gap: 6px; padding: 4px 10px; background: #FFFFFF; border: 1px solid var(--ibm-border); border-radius: 999px; font-size: 12px; }}
.ibm-chip.primary {{ background: var(--ibm-navy); color: #FFFFFF; border-color: var(--ibm-navy); }}
.ibm-chip.warn {{ background: #FFF8E1; border-color: var(--ibm-warn); }}
.ibm-chip.danger {{ background: #FFF1F1; border-color: var(--ibm-danger); color: var(--ibm-danger); }}
.ibm-chip.success {{ background: #DEFBE6; border-color: var(--ibm-success); color: #0E6027; }}
.ibm-card {{ background: #FFFFFF; border: 1px solid var(--ibm-border); border-radius: var(--radius); margin-top: 16px; overflow: hidden; }}
.ibm-card > header {{ padding: 12px 16px; background: #F4F4F4; border-bottom: 1px solid var(--ibm-border); font-weight: 600; color: var(--ibm-navy); display: flex; align-items: center; justify-content: space-between; gap: 10px; }}
.ibm-card > .body {{ padding: 14px 16px; }}
.ibm-kpis {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; }}
.ibm-kpi {{ background: var(--ibm-surface); border: 1px solid var(--ibm-surface-2); border-radius: var(--radius); padding: 12px 14px; }}
.ibm-kpi .v {{ font-size: 26px; font-weight: 600; color: var(--ibm-navy); letter-spacing: -0.01em; }}
.ibm-kpi .l {{ font-size: 12px; color: var(--ibm-text-2); margin-top: 2px; }}
.ibm-kpi.ok .v {{ color: var(--ibm-success); }}
.ibm-kpi.warn .v {{ color: #B28600; }}
.ibm-kpi.danger .v {{ color: var(--ibm-danger); }}
.ibm-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
.ibm-table th {{ background: var(--ibm-navy); color: #FFFFFF; font-weight: 500; text-align: left; padding: 8px 12px; border-bottom: 2px solid var(--ibm-navy-dark); }}
.ibm-table td {{ padding: 8px 12px; border-bottom: 1px solid #E0E0E0; vertical-align: top; }}
.ibm-table tr:nth-child(even) td {{ background: #FAFAFA; }}
.ibm-table tr:hover td {{ background: var(--ibm-surface); }}
.ibm-table td.num, .ibm-table th.num {{ text-align: right; font-variant-numeric: tabular-nums; font-family: var(--font-mono); }}
.ibm-table td.delta-pos {{ color: var(--ibm-success); }}
.ibm-table td.delta-neg {{ color: var(--ibm-danger); }}
.ibm-table td.key {{ font-family: var(--font-mono); font-size: 12px; color: var(--ibm-navy); }}
.ibm-tab-scroll {{ overflow-x: auto; }}
.ibm-btn {{ display: inline-flex; align-items: center; gap: 6px; padding: 8px 14px; background: var(--ibm-navy); color: #FFFFFF; border: none; border-radius: var(--radius); font: 500 13px var(--font-sans); cursor: pointer; transition: background 120ms; }}
.ibm-btn:hover {{ background: var(--ibm-navy-dark); }}
.ibm-btn.secondary {{ background: #FFFFFF; color: var(--ibm-navy); border: 1px solid var(--ibm-navy); }}
.ibm-btn.secondary:hover {{ background: var(--ibm-surface); }}
.ibm-btn[disabled] {{ opacity: 0.5; cursor: not-allowed; }}
.ibm-dl-bar {{ position: sticky; top: 0; z-index: 10; display: flex; gap: 8px; padding: 10px 20px; background: rgba(255,255,255,0.96); backdrop-filter: blur(4px); border-bottom: 1px solid var(--ibm-border); align-items: center; flex-wrap: wrap; }}
.ibm-dl-bar .label {{ font-weight: 600; color: var(--ibm-navy); margin-right: 6px; }}
.ibm-dl-bar .status {{ margin-left: auto; font-size: 12px; color: var(--ibm-text-2); }}
.ibm-footer {{ padding: 12px 20px; margin-top: 20px; border-top: 2px solid var(--ibm-navy); font-size: 11px; color: var(--ibm-text-2); display: flex; justify-content: space-between; gap: 12px; flex-wrap: wrap; }}
.ibm-loading {{ display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 40px 20px; color: var(--ibm-text-2); }}
.ibm-loading .dots span {{ display: inline-block; width: 10px; height: 10px; margin: 0 3px; border-radius: 50%; background: var(--ibm-blue); animation: ibm-bounce 1.2s ease-in-out infinite; }}
.ibm-loading .dots span:nth-child(2) {{ animation-delay: 0.2s; }}
.ibm-loading .dots span:nth-child(3) {{ animation-delay: 0.4s; }}
@keyframes ibm-bounce {{ 0%, 80%, 100% {{ transform: scale(0.4); opacity: 0.4; }} 40% {{ transform: scale(1); opacity: 1; }} }}
.ibm-section-title {{ font-size: 16px; font-weight: 600; color: var(--ibm-navy); margin: 22px 0 8px; display: flex; align-items: center; gap: 8px; }}
.ibm-section-title .bar {{ width: 4px; height: 18px; background: var(--ibm-blue); border-radius: 2px; }}
.ibm-narrative {{ background: #FAFAFA; border-left: 3px solid var(--ibm-blue); padding: 12px 16px; color: var(--ibm-text); font-size: 13.5px; border-radius: 0 var(--radius) var(--radius) 0; }}
.ibm-narrative p {{ margin: 6px 0; }}
.ibm-regtag {{ display: inline-block; padding: 2px 8px; background: var(--ibm-surface-2); color: var(--ibm-navy-dark); border-radius: 999px; font-size: 11px; font-weight: 500; margin-right: 4px; margin-bottom: 2px; }}
.ibm-toast {{ position: fixed; right: 18px; bottom: 18px; padding: 10px 16px; background: var(--ibm-navy); color: #FFFFFF; border-radius: var(--radius); font-size: 13px; box-shadow: 0 4px 14px rgba(0,0,0,0.18); opacity: 0; transform: translateY(10px); transition: opacity 160ms, transform 160ms; pointer-events: none; z-index: 50; }}
.ibm-toast.show {{ opacity: 1; transform: translateY(0); }}
.ibm-toast.err {{ background: var(--ibm-danger); }}
@media print {{
  .ibm-dl-bar, .ibm-btn {{ display: none !important; }}
  .ibm-masthead, .ibm-card > header {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
}}
"""

# ---------------------------------------------------------------------------
# CDN libraries for client-side download (primary path)
# ---------------------------------------------------------------------------

CDN_SCRIPTS = """
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js"></script>
<script src="https://cdn.jsdelivr.net/npm/html-docx-js@0.3.1/dist/html-docx.js"></script>
"""

# ---------------------------------------------------------------------------
# Streaming observer: watches parent chat for @@@RECON-START/END and renders
# ---------------------------------------------------------------------------

OBSERVER_SCRIPT = r"""
<script>
(function(){
  const START = '@@@RECON-START';
  const END = '@@@RECON-END';
  const root = document.getElementById('ibm-render');
  const loader = document.getElementById('ibm-loader');
  const status = document.getElementById('ibm-status');
  const NAVY = '__IBM_NAVY__', BLUE = '__IBM_BLUE__';
  const SERVER_FALLBACK = window.__IBM_SERVER_FALLBACK__ || {};

  function toast(msg, isErr){
    const t = document.getElementById('ibm-toast');
    if(!t) return;
    t.textContent = msg;
    t.classList.toggle('err', !!isErr);
    t.classList.add('show');
    setTimeout(()=>t.classList.remove('show'), 2600);
  }

  function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
  function fmt(n){
    if(n==null||n==='') return '';
    if(typeof n==='number') return n.toLocaleString(undefined,{maximumFractionDigits:2,minimumFractionDigits:(Math.abs(n)%1>0?2:0)});
    const f=parseFloat(String(n).replace(/,/g,''));
    return isFinite(f)?f.toLocaleString(undefined,{maximumFractionDigits:2,minimumFractionDigits:(Math.abs(f)%1>0?2:0)}):String(n);
  }

  function renderHeader(meta){
    const chips = [];
    if(meta.sector) chips.push(`<span class="ibm-chip primary">${esc(meta.sector)}</span>`);
    if(meta.region) chips.push(`<span class="ibm-chip">${esc(meta.region)}</span>`);
    (meta.regulations||[]).forEach(r=>chips.push(`<span class="ibm-regtag">${esc(r)}</span>`));
    if(meta.asOf) chips.push(`<span class="ibm-chip">as of ${esc(meta.asOf)}</span>`);
    if(meta.redacted) chips.push(`<span class="ibm-chip warn">PII redacted</span>`);
    return `<div class="ibm-subbar">${chips.join('')}</div>`;
  }

  function renderKpis(s){
    const items = [
      {v:s.total_left||0, l:'Records (Left)'},
      {v:s.total_right||0, l:'Records (Right)'},
      {v:s.matched||0, l:'Matched', cls:'ok'},
      {v:s.variance||0, l:'Variance', cls:'warn'},
      {v:s.unmatched_left||0, l:'Missing in Right', cls:(s.unmatched_left>0?'danger':'')},
      {v:s.unmatched_right||0, l:'Missing in Left', cls:(s.unmatched_right>0?'danger':'')},
    ];
    if(s.match_rate!=null) items.push({v:s.match_rate+'%', l:'Match rate', cls:(s.match_rate>=95?'ok':s.match_rate>=80?'warn':'danger')});
    return `<div class="ibm-card"><header>Reconciliation KPIs</header><div class="body"><div class="ibm-kpis">${
      items.map(i=>`<div class="ibm-kpi ${i.cls||''}"><div class="v">${esc(fmt(i.v))}</div><div class="l">${esc(i.l)}</div></div>`).join('')
    }</div></div></div>`;
  }

  function renderTable(title, cols, rows, extraClass){
    if(!rows||!rows.length) return `<div class="ibm-card"><header>${esc(title)} <span class="ibm-chip success">0</span></header><div class="body"><em>None.</em></div></div>`;
    const th = cols.map(c=>`<th class="${c.num?'num':''}">${esc(c.label||c.key)}</th>`).join('');
    const tr = rows.map(r=>{
      const tds = cols.map(c=>{
        let v = r[c.key];
        let cls = c.num?'num':'';
        if(c.key==='delta'||c.key==='variance'){
          const nv = parseFloat(v);
          if(isFinite(nv)) cls += (nv<0?' delta-neg':(nv>0?' delta-pos':''));
        }
        if(c.keyField) cls += ' key';
        return `<td class="${cls}">${esc(c.num?fmt(v):v)}</td>`;
      }).join('');
      return `<tr>${tds}</tr>`;
    }).join('');
    return `<div class="ibm-card ${extraClass||''}"><header>${esc(title)} <span class="ibm-chip">${rows.length} row${rows.length===1?'':'s'}</span></header><div class="body"><div class="ibm-tab-scroll"><table class="ibm-table"><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table></div></div></div>`;
  }

  function renderNarrative(items){
    if(!items||!items.length) return '';
    const body = items.map(n=>{
      const tag = n.regulation?`<span class="ibm-regtag">${esc(n.regulation)}</span>`:'';
      const sev = n.severity==='high'?'danger':(n.severity==='medium'?'warn':'');
      const chip = sev?`<span class="ibm-chip ${sev}">${esc(n.severity)}</span> `:'';
      return `<p>${chip}${tag}${esc(n.text||'')}</p>`;
    }).join('');
    return `<div class="ibm-section-title"><span class="bar"></span>Variance commentary</div><div class="ibm-narrative">${body}</div>`;
  }

  function renderFooter(meta){
    return `<div class="ibm-footer"><span>IBM Consulting Advantage — AI Reconciliation</span><span>Build __BUILD__ · Run ${esc(meta.runId||'')}</span></div>`;
  }

  // --- Payload state (populated when RECON block arrives) ---
  let PAYLOAD = null;

  function renderAll(p){
    PAYLOAD = p;
    const meta = p.meta||{};
    const stats = p.stats||{};
    const parts = [];
    parts.push(renderHeader(meta));
    parts.push(renderKpis(stats));

    const matched = p.matched||[];
    const variance = p.variance||[];
    const ul = p.unmatched_left||[];
    const ur = p.unmatched_right||[];

    if(variance.length){
      const cols = p.variance_columns || inferCols(variance[0]);
      parts.push(`<div class="ibm-section-title"><span class="bar"></span>Variances — amount / field mismatches</div>`);
      parts.push(renderTable('Variance detail', cols, variance));
    }
    if(ul.length){
      const cols = p.unmatched_columns || inferCols(ul[0]);
      parts.push(`<div class="ibm-section-title"><span class="bar"></span>Missing in Right side</div>`);
      parts.push(renderTable('Missing in Right', cols, ul));
    }
    if(ur.length){
      const cols = p.unmatched_columns || inferCols(ur[0]);
      parts.push(`<div class="ibm-section-title"><span class="bar"></span>Missing in Left side</div>`);
      parts.push(renderTable('Missing in Left', cols, ur));
    }
    if(matched.length){
      const cols = p.matched_columns || inferCols(matched[0]);
      parts.push(`<div class="ibm-section-title"><span class="bar"></span>Matched records (preview)</div>`);
      parts.push(renderTable('Matched', cols, matched.slice(0,25)));
    }
    parts.push(renderNarrative(p.narrative));
    parts.push(renderFooter(meta));

    root.innerHTML = parts.join('\n');
    if(loader) loader.remove();
    status.textContent = 'Reconciliation complete';
    updateDownloadButtons(true);
  }

  function inferCols(r){
    return Object.keys(r).map(k=>({key:k, label:k, num:/amount|qty|quantity|price|value|delta|variance|paid|allowed|charge|premium|reserve/i.test(k)}));
  }

  // --- Download helpers ---
  function updateDownloadButtons(enabled){
    ['ibm-dl-xlsx','ibm-dl-docx','ibm-dl-pptx'].forEach(id=>{
      const b = document.getElementById(id);
      if(b) b.disabled = !enabled;
    });
  }

  function triggerDataUri(name, dataUri){
    const a = document.createElement('a');
    a.href = dataUri;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>a.remove(), 200);
  }

  function fallbackDownload(fmtName){
    const payload = SERVER_FALLBACK[fmtName];
    if(!payload){ toast('Server fallback not available', true); return; }
    triggerDataUri(payload.filename, payload.dataUri);
    toast('Downloaded via ooXML fallback');
  }

  window.ibmDownload = function(fmtName){
    if(!PAYLOAD){ toast('Reconciliation payload not ready', true); return; }
    const meta = PAYLOAD.meta||{};
    const baseName = (meta.outputName || 'reconciliation') + '_' + (meta.asOf||'').replace(/[^0-9A-Za-z]/g,'');
    try{
      if(fmtName==='xlsx'){
        if(typeof XLSX==='undefined') throw new Error('XLSX CDN unavailable');
        const wb = XLSX.utils.book_new();
        const summary = [['IBM Consulting Advantage — Reconciliation'],[],['Sector', meta.sector||''],['Region', meta.region||''],['As of', meta.asOf||''],['Run ID', meta.runId||''],[],['KPI','Value']];
        Object.entries(PAYLOAD.stats||{}).forEach(([k,v])=>summary.push([k, v]));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), 'Summary');
        const sections = [['Variance', PAYLOAD.variance||[]],['Missing_in_Right', PAYLOAD.unmatched_left||[]],['Missing_in_Left', PAYLOAD.unmatched_right||[]],['Matched', PAYLOAD.matched||[]]];
        sections.forEach(([name, rows])=>{
          if(!rows.length) return;
          XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), name);
        });
        XLSX.writeFile(wb, baseName+'.xlsx');
        toast('XLSX downloaded via CDN');
      } else if(fmtName==='docx'){
        if(typeof window.htmlDocx==='undefined') throw new Error('html-docx CDN unavailable');
        const shell = buildDocxHtml(PAYLOAD);
        const blob = window.htmlDocx.asBlob(shell);
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href=url; a.download=baseName+'.docx'; document.body.appendChild(a); a.click();
        setTimeout(()=>{URL.revokeObjectURL(url); a.remove();},400);
        toast('DOCX downloaded via CDN');
      } else if(fmtName==='pptx'){
        if(typeof PptxGenJS==='undefined') throw new Error('PptxGenJS CDN unavailable');
        buildPptx(PAYLOAD).then(()=>toast('PPTX downloaded via CDN'));
      }
    } catch(e){
      console.warn('CDN path failed, using ooXML fallback:', e.message);
      fallbackDownload(fmtName);
    }
  };

  function buildDocxHtml(p){
    const meta=p.meta||{}, s=p.stats||{};
    const tbl = (rows, caption)=>{
      if(!rows.length) return '';
      const keys = Object.keys(rows[0]);
      const th = keys.map(k=>`<th style="background:${NAVY};color:#fff;padding:6px;border:1px solid #888;">${k}</th>`).join('');
      const tr = rows.map(r=>`<tr>${keys.map(k=>`<td style="padding:6px;border:1px solid #ccc;">${(r[k]==null?'':String(r[k]).replace(/[<>&]/g,''))}</td>`).join('')}</tr>`).join('');
      return `<h3 style="color:${NAVY};">${caption} (${rows.length})</h3><table style="border-collapse:collapse;width:100%;font-size:11px;"><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table>`;
    };
    const narr = (p.narrative||[]).map(n=>`<p style="border-left:3px solid ${BLUE};padding-left:10px;margin:6px 0;"><strong>${n.regulation||''}</strong> ${n.text||''}</p>`).join('');
    const kpis = Object.entries(s).map(([k,v])=>`<li><strong>${k}:</strong> ${v}</li>`).join('');
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Reconciliation</title></head><body style="font-family:'Segoe UI',Arial,sans-serif;color:#161616;">
    <div style="background:${NAVY};color:#fff;padding:16px;"><h1 style="margin:0;">IBM Consulting Advantage — AI Reconciliation</h1>
    <div style="opacity:.85;font-size:12px;">${meta.sector||''} · ${meta.region||''} · ${meta.asOf||''}</div></div>
    <h2 style="color:${NAVY};">Summary</h2><ul>${kpis}</ul>
    ${narr?`<h2 style="color:${NAVY};">Commentary</h2>${narr}`:''}
    ${tbl(p.variance||[], 'Variances')}
    ${tbl(p.unmatched_left||[], 'Missing in Right')}
    ${tbl(p.unmatched_right||[], 'Missing in Left')}
    ${tbl((p.matched||[]).slice(0,50), 'Matched (sample)')}
    </body></html>`;
  }

  async function buildPptx(p){
    const pres = new PptxGenJS();
    pres.layout = 'LAYOUT_WIDE';
    const meta=p.meta||{}, s=p.stats||{};
    const navy = NAVY.replace('#',''), blue=BLUE.replace('#','');
    // Title slide
    let sl = pres.addSlide(); sl.background={color:navy};
    sl.addText('IBM Consulting Advantage', {x:0.5,y:0.4,w:12,h:0.5,fontSize:18,color:'FFFFFF',bold:true});
    sl.addText('AI Reconciliation Report', {x:0.5,y:1.2,w:12,h:1.0,fontSize:40,color:'FFFFFF',bold:true});
    sl.addText(`${meta.sector||''} · ${meta.region||''} · ${meta.asOf||''}`, {x:0.5,y:2.4,w:12,h:0.5,fontSize:18,color:'D0E2FF'});
    // KPI slide
    sl = pres.addSlide(); sl.background={color:'FFFFFF'};
    sl.addShape(pres.ShapeType.rect,{x:0,y:0,w:13.33,h:0.6,fill:{color:navy}});
    sl.addText('Reconciliation KPIs', {x:0.3,y:0.05,w:12,h:0.5,fontSize:20,color:'FFFFFF',bold:true});
    const kpiRows = Object.entries(s);
    const perRow = 4, cw = 2.9, ch = 1.1;
    kpiRows.forEach(([k,v],i)=>{
      const col = i%perRow, row = Math.floor(i/perRow);
      const x = 0.5 + col*(cw+0.1), y = 1.0 + row*(ch+0.2);
      sl.addShape(pres.ShapeType.rect,{x,y,w:cw,h:ch,fill:{color:'EDF5FF'},line:{color:'D0E2FF',width:1}});
      sl.addText(String(v), {x:x,y:y,w:cw,h:ch*0.55,fontSize:22,color:navy,bold:true,align:'center',valign:'middle'});
      sl.addText(k, {x:x,y:y+ch*0.55,w:cw,h:ch*0.45,fontSize:10,color:'525252',align:'center'});
    });
    // Table slides for each section
    const sections = [['Variances', p.variance||[]],['Missing in Right', p.unmatched_left||[]],['Missing in Left', p.unmatched_right||[]]];
    sections.forEach(([title, rows])=>{
      if(!rows.length) return;
      sl = pres.addSlide(); sl.background={color:'FFFFFF'};
      sl.addShape(pres.ShapeType.rect,{x:0,y:0,w:13.33,h:0.6,fill:{color:navy}});
      sl.addText(title+` (${rows.length})`, {x:0.3,y:0.05,w:12,h:0.5,fontSize:20,color:'FFFFFF',bold:true});
      const keys = Object.keys(rows[0]).slice(0,6);
      const header = keys.map(k=>({text:k,options:{bold:true,color:'FFFFFF',fill:{color:navy}}}));
      const data = [header].concat(rows.slice(0,15).map(r=>keys.map(k=>({text:String(r[k]==null?'':r[k])}))));
      sl.addTable(data, {x:0.3,y:0.8,w:12.7,fontSize:9,border:{type:'solid',pt:0.5,color:'CCCCCC'}});
    });
    // Narrative slide
    if((p.narrative||[]).length){
      sl = pres.addSlide(); sl.background={color:'FFFFFF'};
      sl.addShape(pres.ShapeType.rect,{x:0,y:0,w:13.33,h:0.6,fill:{color:navy}});
      sl.addText('Variance Commentary', {x:0.3,y:0.05,w:12,h:0.5,fontSize:20,color:'FFFFFF',bold:true});
      const bullets = (p.narrative||[]).slice(0,8).map(n=>({text:(n.regulation?`[${n.regulation}] `:'')+(n.text||''),options:{bullet:true,fontSize:14,color:'161616'}}));
      sl.addText(bullets,{x:0.5,y:0.9,w:12.3,h:6.2});
    }
    const name = (meta.outputName || 'reconciliation') + '_' + (meta.asOf||'').replace(/[^0-9A-Za-z]/g,'') + '.pptx';
    await pres.writeFile({fileName:name});
  }

  // --- Observer: tail parent chat for @@@RECON block ---
  function readParentText(){
    try{
      const doc = window.parent && window.parent.document;
      if(!doc) return null;
      // Find the LAST message bubble that contains our markers.
      // Fall back to concatenated textContent of all message-content nodes.
      const nodes = doc.querySelectorAll('.chat-assistant .prose, .message-content, [data-message-role="assistant"]');
      if(nodes && nodes.length){
        for(let i=nodes.length-1;i>=0;i--){
          const t = nodes[i].innerText || nodes[i].textContent || '';
          if(t.indexOf(START)>=0) return t;
        }
        return nodes[nodes.length-1].innerText || '';
      }
      return doc.body ? doc.body.innerText : '';
    }catch(e){ return null; }
  }

  function extractBlock(text){
    if(!text) return null;
    const i = text.indexOf(START); if(i<0) return null;
    const j = text.indexOf(END, i); if(j<0) return null;
    return text.slice(i+START.length, j).trim();
  }

  function tryParse(raw){
    if(!raw) return null;
    // Strip accidental code fences
    raw = raw.replace(/^```(?:json)?/i,'').replace(/```$/,'').trim();
    try { return JSON.parse(raw); } catch(e){}
    // Extract largest {...} blob
    const m = raw.match(/\{[\s\S]*\}/);
    if(m){ try{ return JSON.parse(m[0]); }catch(e){} }
    return null;
  }

  let done = false;
  function tick(){
    if(done) return;
    const txt = readParentText();
    const block = extractBlock(txt||'');
    if(block){
      const obj = tryParse(block);
      if(obj && obj.stats){
        done = true;
        renderAll(obj);
        try{ hideMarkersInParent(); }catch(e){}
      }
    }
  }

  function hideMarkersInParent(){
    try{
      const doc = window.parent.document;
      const style = doc.createElement('style');
      style.textContent = `.ibm-recon-hidden{display:none!important;}`;
      doc.head.appendChild(style);
      // Not a perfect hider, but most OWUI themes collapse fenced text — we leave rendered.
    }catch(e){}
  }

  // Also accept payload posted directly from Python (inline server-rendered payload)
  if(window.__IBM_RECON_PAYLOAD__){
    try { renderAll(window.__IBM_RECON_PAYLOAD__); } catch(e){ console.error(e); }
  }

  const iv = setInterval(()=>{ if(done){ clearInterval(iv); return; } tick(); }, 600);
  setTimeout(()=>{ if(!done){ if(status) status.textContent = 'Awaiting reconciliation stream…'; } }, 3000);

  // Wire up buttons
  document.addEventListener('click', (e)=>{
    const t = e.target.closest('[data-ibm-dl]');
    if(!t) return;
    window.ibmDownload(t.getAttribute('data-ibm-dl'));
  });

  // Height auto-size
  function autoHeight(){
    try{
      const h = Math.max(document.documentElement.scrollHeight, document.body.scrollHeight);
      window.parent.postMessage({type:'iframe:resize', height:h}, '*');
    }catch(e){}
  }
  setInterval(autoHeight, 700);
  window.addEventListener('load', autoHeight);
})();
</script>
"""

# ---------------------------------------------------------------------------
# HTML shell
# ---------------------------------------------------------------------------


def _error_shell(message: str) -> str:
    safe = htmlmod.escape(message)
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>{THEME_CSS}</style></head><body>
<div class="ibm-masthead"><div class="brand"><span class="dot"></span>IBM Consulting Advantage · AI Reconciliation</div>
<div class="meta">Input required</div></div>
<div class="ibm-shell"><div class="ibm-card"><header>Unable to reconcile</header>
<div class="body"><p>{safe}</p>
<p style="color:var(--ibm-text-2);font-size:12px;margin-top:12px;">Tip: drop both files into the chat using the paperclip icon, then ask the assistant to reconcile them.</p>
</div></div></div></body></html>"""


def _shell_html(
    meta: Dict[str, Any],
    server_fallback: Dict[str, Dict[str, str]],
    inline_payload: Optional[Dict[str, Any]] = None,
) -> str:
    fallback_json = json.dumps(server_fallback)
    payload_json = json.dumps(inline_payload) if inline_payload else "null"
    observer = (
        OBSERVER_SCRIPT
        .replace("__IBM_NAVY__", IBM_NAVY)
        .replace("__IBM_BLUE__", IBM_BLUE)
        .replace("__BUILD__", _BUILD)
    )
    safe_title = htmlmod.escape(meta.get("title", "AI Reconciliation"))
    run_id = htmlmod.escape(meta.get("runId", ""))
    sector = htmlmod.escape(meta.get("sector", ""))
    region = htmlmod.escape(meta.get("region", ""))

    return f"""<!DOCTYPE html>
<html lang="en" data-ibm-build="{_BUILD}">
<head>
<meta charset="utf-8">
<title>{safe_title}</title>
<style>{THEME_CSS}</style>
{CDN_SCRIPTS}
</head>
<body>
<div class="ibm-masthead">
  <div class="brand"><span class="dot"></span>IBM Consulting Advantage · AI Reconciliation</div>
  <div class="meta">{sector} · {region} · run {run_id}</div>
</div>
<div class="ibm-dl-bar">
  <span class="label">Download reconciled file:</span>
  <button class="ibm-btn" id="ibm-dl-xlsx" data-ibm-dl="xlsx" disabled>XLSX</button>
  <button class="ibm-btn" id="ibm-dl-docx" data-ibm-dl="docx" disabled>DOCX</button>
  <button class="ibm-btn" id="ibm-dl-pptx" data-ibm-dl="pptx" disabled>PPTX</button>
  <span class="status" id="ibm-status">Awaiting reconciliation stream…</span>
</div>
<div class="ibm-shell">
  <div id="ibm-render"></div>
  <div id="ibm-loader" class="ibm-loading">
    <div class="dots"><span></span><span></span><span></span></div>
    <div style="margin-top:12px;">Receiving reconciliation output…</div>
  </div>
</div>
<div id="ibm-toast" class="ibm-toast"></div>
<script>
  window.__IBM_SERVER_FALLBACK__ = {fallback_json};
  window.__IBM_RECON_PAYLOAD__ = {payload_json};
</script>
{observer}
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Reconciliation math (deterministic, runs in Python not LLM)
# ---------------------------------------------------------------------------


def _to_float(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def _norm_str(v: Any) -> str:
    return re.sub(r"\s+", " ", str(v or "").strip()).lower()


def _key_for(rec: Dict[str, Any], key_fields: List[str]) -> str:
    parts = []
    for k in key_fields:
        parts.append(_norm_str(rec.get(k, "")))
    return "||".join(parts)


def _redact_pii(records: List[Dict[str, Any]], sector: str) -> List[Dict[str, Any]]:
    """HIPAA / GDPR-aware redaction for healthcare downloads."""
    if sector.lower() not in ("healthcare", "pharma_clinical", "pharma/clinical", "pharmaceutical"):
        return records
    out = []
    name_keys = {"patientname", "patient", "insured", "policyholder", "subject", "patientfullname"}
    dob_keys = {"patientdob", "dob", "birthdate", "dateofbirth"}
    id_keys = {"ssn", "socialsecurity", "memberid", "member_id", "patientid", "mrn"}
    for r in records:
        new = {}
        for k, v in r.items():
            lk = re.sub(r"[^a-z]", "", k.lower())
            if lk in name_keys and v:
                parts = re.split(r"[\s,]+", str(v).strip())
                masked = " ".join((p[0] + "." if p else "") for p in parts if p)
                new[k] = masked or "***"
            elif lk in dob_keys and v:
                s = str(v)
                new[k] = s[:4] + "-**-**" if len(s) >= 4 else "****"
            elif lk in id_keys and v:
                s = str(v)
                new[k] = (s[:3] + "***" + s[-2:]) if len(s) > 5 else "***"
            else:
                new[k] = v
        out.append(new)
    return out


def _reconcile(
    left: List[Dict[str, Any]],
    right: List[Dict[str, Any]],
    key_fields: List[str],
    amount_fields: List[str],
    tolerance: float,
) -> Dict[str, Any]:
    """Deterministic partition: matched / variance / unmatched_left / unmatched_right."""
    right_index: Dict[str, List[int]] = {}
    for i, r in enumerate(right):
        k = _key_for(r, key_fields)
        right_index.setdefault(k, []).append(i)

    used_right = set()
    matched, variance = [], []
    unmatched_left = []
    for lrec in left:
        k = _key_for(lrec, key_fields)
        cand = right_index.get(k, [])
        picked = None
        for ri in cand:
            if ri not in used_right:
                picked = ri
                break
        if picked is None:
            unmatched_left.append(lrec)
            continue
        used_right.add(picked)
        rrec = right[picked]
        # Check amount fields for variance
        diffs = {}
        for f in amount_fields:
            lv = _to_float(lrec.get(f))
            # right may have a differently-named amount field; check same name first, then common synonyms
            rv = _to_float(rrec.get(f))
            if rv is None:
                for alt in (f.lower(), f.upper(), f.replace("_", "").lower(), "amount", "Amount", "NetAmount", "Consideration", "MV", "MarketValue", "Paid", "PaidAmount", "BilledAmount", "Charged"):
                    if alt in rrec:
                        rv = _to_float(rrec.get(alt))
                        if rv is not None:
                            break
            if lv is None and rv is None:
                continue
            if lv is None or rv is None:
                diffs[f] = {"left": lv, "right": rv, "delta": None}
                continue
            if abs(lv - rv) > tolerance:
                diffs[f] = {"left": lv, "right": rv, "delta": round(rv - lv, 4)}
        if diffs:
            row = {"_key": k}
            row.update(lrec)
            for f, d in diffs.items():
                row[f"{f}_left"] = d["left"]
                row[f"{f}_right"] = d["right"]
                row[f"{f}_delta"] = d["delta"]
            variance.append(row)
        else:
            matched.append(lrec)

    unmatched_right = [right[i] for i in range(len(right)) if i not in used_right]

    total = len(matched) + len(variance) + len(unmatched_left)
    match_rate = round((len(matched) / total * 100), 2) if total else 0.0

    return {
        "matched": matched,
        "variance": variance,
        "unmatched_left": unmatched_left,
        "unmatched_right": unmatched_right,
        "stats": {
            "total_left": len(left),
            "total_right": len(right),
            "matched": len(matched),
            "variance": len(variance),
            "unmatched_left": len(unmatched_left),
            "unmatched_right": len(unmatched_right),
            "match_rate": match_rate,
        },
    }


# ---------------------------------------------------------------------------
# ooXML fallback — always built server-side, shipped as base64 data URIs
# ---------------------------------------------------------------------------


def _build_xlsx_bytes(payload: Dict[str, Any]) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    navy = "FF002D74"
    header_font = Font(bold=True, color="FFFFFFFF", name="Calibri", size=11)
    header_fill = PatternFill("solid", fgColor=navy)
    title_font = Font(bold=True, size=14, color="FF002D74")

    def add_sheet(name: str, rows: List[Dict[str, Any]], title: str) -> None:
        if name == "Summary":
            ws = wb.active
            ws.title = name
        else:
            ws = wb.create_sheet(name)
        ws["A1"] = title
        ws["A1"].font = title_font
        ws["A2"] = "IBM Consulting Advantage — AI Reconciliation"
        if not rows:
            ws["A4"] = "(no rows)"
            return
        keys = list(rows[0].keys())
        for j, k in enumerate(keys, 1):
            c = ws.cell(row=4, column=j, value=k)
            c.font = header_font
            c.fill = header_fill
            c.alignment = Alignment(horizontal="left")
        for i, r in enumerate(rows, 5):
            for j, k in enumerate(keys, 1):
                v = r.get(k, "")
                if isinstance(v, (dict, list)):
                    v = json.dumps(v, default=str)
                ws.cell(row=i, column=j, value=v)
        for j, k in enumerate(keys, 1):
            col = get_column_letter(j)
            ws.column_dimensions[col].width = min(max(len(str(k)), 12), 40)

    meta = payload.get("meta", {})
    stats = payload.get("stats", {})
    summary_rows = [
        {"Field": "Sector", "Value": meta.get("sector", "")},
        {"Field": "Region", "Value": meta.get("region", "")},
        {"Field": "As of", "Value": meta.get("asOf", "")},
        {"Field": "Run ID", "Value": meta.get("runId", "")},
        {"Field": "Regulations", "Value": ", ".join(meta.get("regulations", []))},
    ]
    for k, v in stats.items():
        summary_rows.append({"Field": k, "Value": v})
    add_sheet("Summary", summary_rows, "Reconciliation Summary")
    sections = [
        ("Variance", payload.get("variance", [])),
        ("Missing_in_Right", payload.get("unmatched_left", [])),
        ("Missing_in_Left", payload.get("unmatched_right", [])),
        ("Matched", payload.get("matched", [])),
        ("Commentary", [{"severity": n.get("severity", ""), "regulation": n.get("regulation", ""), "text": n.get("text", "")} for n in payload.get("narrative", [])]),
    ]
    for name, rows in sections:
        add_sheet(name, rows, name.replace("_", " "))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _build_docx_bytes(payload: Dict[str, Any]) -> bytes:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document()
    meta = payload.get("meta", {})
    stats = payload.get("stats", {})

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("IBM Consulting Advantage")
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0x00, 0x2D, 0x74)

    h = doc.add_heading("AI Reconciliation Report", level=0)
    for r in h.runs:
        r.font.color.rgb = RGBColor(0x00, 0x2D, 0x74)

    doc.add_paragraph(f"Sector: {meta.get('sector','')}   Region: {meta.get('region','')}   As of: {meta.get('asOf','')}")
    if meta.get("regulations"):
        doc.add_paragraph("Regulations: " + ", ".join(meta["regulations"]))
    if meta.get("redacted"):
        doc.add_paragraph("Note: patient/subject PII has been redacted in accordance with HIPAA / GDPR.")

    doc.add_heading("Summary", level=1)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Light Grid Accent 1"
    hdr = tbl.rows[0].cells
    hdr[0].text = "KPI"
    hdr[1].text = "Value"
    for k, v in stats.items():
        row = tbl.add_row().cells
        row[0].text = str(k)
        row[1].text = str(v)

    def add_table(title: str, rows: List[Dict[str, Any]]) -> None:
        doc.add_heading(title + f" ({len(rows)})", level=1)
        if not rows:
            doc.add_paragraph("(none)")
            return
        keys = list(rows[0].keys())[:8]
        t = doc.add_table(rows=1, cols=len(keys))
        t.style = "Light Grid Accent 1"
        for j, k in enumerate(keys):
            t.rows[0].cells[j].text = k
        for r in rows[:200]:
            cells = t.add_row().cells
            for j, k in enumerate(keys):
                v = r.get(k, "")
                if isinstance(v, (dict, list)):
                    v = json.dumps(v, default=str)
                cells[j].text = str(v)[:120]

    add_table("Variances", payload.get("variance", []))
    add_table("Missing in Right", payload.get("unmatched_left", []))
    add_table("Missing in Left", payload.get("unmatched_right", []))
    if payload.get("matched"):
        add_table("Matched (sample)", payload.get("matched", [])[:50])

    if payload.get("narrative"):
        doc.add_heading("Variance commentary", level=1)
        for n in payload["narrative"]:
            para = doc.add_paragraph()
            if n.get("regulation"):
                r1 = para.add_run(f"[{n['regulation']}] ")
                r1.bold = True
                r1.font.color.rgb = RGBColor(0x00, 0x43, 0xCE)
            para.add_run(n.get("text", ""))

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _build_pptx_bytes(payload: Dict[str, Any]) -> bytes:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE

    pres = Presentation()
    pres.slide_width = Inches(13.333)
    pres.slide_height = Inches(7.5)
    meta = payload.get("meta", {})
    stats = payload.get("stats", {})
    NAVY = RGBColor(0x00, 0x2D, 0x74)
    BLUE = RGBColor(0x00, 0x43, 0xCE)
    SURF = RGBColor(0xED, 0xF5, 0xFF)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)

    blank = pres.slide_layouts[6]

    def title_bar(slide, text):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, pres.slide_width, Inches(0.6))
        shape.fill.solid()
        shape.fill.fore_color.rgb = NAVY
        shape.line.fill.background()
        tb = slide.shapes.add_textbox(Inches(0.3), Inches(0.05), Inches(12), Inches(0.5))
        tf = tb.text_frame
        tf.text = text
        tf.paragraphs[0].runs[0].font.size = Pt(20)
        tf.paragraphs[0].runs[0].font.bold = True
        tf.paragraphs[0].runs[0].font.color.rgb = WHITE

    # Slide 1: cover
    s = pres.slides.add_slide(blank)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, pres.slide_width, pres.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = NAVY
    bg.line.fill.background()
    tb = s.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12), Inches(0.6))
    tb.text_frame.text = "IBM Consulting Advantage"
    tb.text_frame.paragraphs[0].runs[0].font.size = Pt(18)
    tb.text_frame.paragraphs[0].runs[0].font.bold = True
    tb.text_frame.paragraphs[0].runs[0].font.color.rgb = WHITE
    tb = s.shapes.add_textbox(Inches(0.5), Inches(1.4), Inches(12), Inches(1.2))
    tb.text_frame.text = "AI Reconciliation Report"
    tb.text_frame.paragraphs[0].runs[0].font.size = Pt(44)
    tb.text_frame.paragraphs[0].runs[0].font.bold = True
    tb.text_frame.paragraphs[0].runs[0].font.color.rgb = WHITE
    tb = s.shapes.add_textbox(Inches(0.5), Inches(2.8), Inches(12), Inches(0.5))
    tb.text_frame.text = f"{meta.get('sector','')} · {meta.get('region','')} · as of {meta.get('asOf','')}"
    tb.text_frame.paragraphs[0].runs[0].font.size = Pt(18)
    tb.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xD0, 0xE2, 0xFF)

    # Slide 2: KPIs
    s = pres.slides.add_slide(blank)
    title_bar(s, "Reconciliation KPIs")
    items = list(stats.items())
    per = 4
    cw, ch = Inches(2.9), Inches(1.1)
    for i, (k, v) in enumerate(items):
        col, row = i % per, i // per
        x = Inches(0.5 + col * 3.0)
        y = Inches(1.0 + row * 1.3)
        box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, cw, ch)
        box.fill.solid()
        box.fill.fore_color.rgb = SURF
        box.line.color.rgb = RGBColor(0xD0, 0xE2, 0xFF)
        tb = s.shapes.add_textbox(x, y, cw, Inches(0.6))
        tb.text_frame.text = str(v)
        tb.text_frame.paragraphs[0].runs[0].font.size = Pt(22)
        tb.text_frame.paragraphs[0].runs[0].font.bold = True
        tb.text_frame.paragraphs[0].runs[0].font.color.rgb = NAVY
        tb2 = s.shapes.add_textbox(x, Emu(y + Inches(0.55)), cw, Inches(0.5))
        tb2.text_frame.text = str(k)
        tb2.text_frame.paragraphs[0].runs[0].font.size = Pt(10)
        tb2.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x52, 0x52, 0x52)

    def table_slide(title: str, rows: List[Dict[str, Any]]) -> None:
        if not rows:
            return
        s = pres.slides.add_slide(blank)
        title_bar(s, f"{title} ({len(rows)})")
        keys = list(rows[0].keys())[:6]
        n_rows = min(len(rows), 14) + 1
        tbl_shape = s.shapes.add_table(n_rows, len(keys), Inches(0.3), Inches(0.8), Inches(12.7), Inches(6.2))
        tbl = tbl_shape.table
        for j, k in enumerate(keys):
            cell = tbl.cell(0, j)
            cell.text = str(k)
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.color.rgb = WHITE
                    r.font.bold = True
                    r.font.size = Pt(10)
        for i, r in enumerate(rows[:14], 1):
            for j, k in enumerate(keys):
                v = r.get(k, "")
                if isinstance(v, (dict, list)):
                    v = json.dumps(v, default=str)
                cell = tbl.cell(i, j)
                cell.text = str(v)[:60]
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(9)

    table_slide("Variances", payload.get("variance", []))
    table_slide("Missing in Right", payload.get("unmatched_left", []))
    table_slide("Missing in Left", payload.get("unmatched_right", []))

    # Narrative
    if payload.get("narrative"):
        s = pres.slides.add_slide(blank)
        title_bar(s, "Variance Commentary")
        tb = s.shapes.add_textbox(Inches(0.5), Inches(0.9), Inches(12.3), Inches(6.2))
        tf = tb.text_frame
        tf.word_wrap = True
        first = True
        for n in payload["narrative"][:10]:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.level = 0
            tag = f"[{n.get('regulation','')}] " if n.get("regulation") else ""
            run1 = p.add_run()
            run1.text = tag
            run1.font.bold = True
            run1.font.color.rgb = BLUE
            run1.font.size = Pt(14)
            run2 = p.add_run()
            run2.text = n.get("text", "")
            run2.font.size = Pt(14)
            run2.font.color.rgb = RGBColor(0x16, 0x16, 0x16)

    buf = io.BytesIO()
    pres.save(buf)
    return buf.getvalue()


def _b64_data_uri(mime: str, data: bytes) -> str:
    return f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"


def _build_server_fallback(payload: Dict[str, Any], base_name: str) -> Dict[str, Dict[str, str]]:
    out: Dict[str, Dict[str, str]] = {}
    try:
        xlsx = _build_xlsx_bytes(payload)
        out["xlsx"] = {
            "filename": f"{base_name}.xlsx",
            "dataUri": _b64_data_uri(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xlsx
            ),
        }
    except Exception as e:
        out["xlsx_error"] = {"filename": "", "dataUri": f"data:text/plain;base64,{base64.b64encode(str(e).encode()).decode()}"}
    try:
        docx = _build_docx_bytes(payload)
        out["docx"] = {
            "filename": f"{base_name}.docx",
            "dataUri": _b64_data_uri(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document", docx
            ),
        }
    except Exception as e:
        out["docx_error"] = {"filename": "", "dataUri": f"data:text/plain;base64,{base64.b64encode(str(e).encode()).decode()}"}
    try:
        pptx = _build_pptx_bytes(payload)
        out["pptx"] = {
            "filename": f"{base_name}.pptx",
            "dataUri": _b64_data_uri(
                "application/vnd.openxmlformats-officedocument.presentationml.presentation", pptx
            ),
        }
    except Exception as e:
        out["pptx_error"] = {"filename": "", "dataUri": f"data:text/plain;base64,{base64.b64encode(str(e).encode()).decode()}"}
    return out


# ---------------------------------------------------------------------------
# Tool class
# ---------------------------------------------------------------------------


class Tools:
    """IBM Consulting Advantage — AI Reconciliation.

    The LLM parses uploaded regulatory documents from Open WebUI's attachment
    context, normalises them into two structured record lists (left / right),
    decides the key fields and amount fields per sector, and calls
    ``reconcile`` to get a deterministic partition plus an inline IBM-themed
    report with CDN-first / ooXML-fallback downloads (XLSX / DOCX / PPTX).
    """

    class Valves(BaseModel):
        default_tolerance: float = Field(
            default=0.01,
            description="Default numeric tolerance (absolute) for amount comparisons when the model does not specify one.",
        )
        max_preview_rows: int = Field(
            default=200,
            description="Maximum rows per section included in the inline preview and downloadable outputs.",
        )

    def __init__(self) -> None:
        self.valves = self.Valves()

    async def reconcile(
        self,
        left_records: List[Dict[str, Any]],
        right_records: List[Dict[str, Any]],
        sector: Literal[
            "Banking",
            "Investment Banking",
            "Insurance",
            "Healthcare",
            "Asset Management",
            "Pharma Clinical",
            "Other",
        ] = "Banking",
        region: Literal["USA", "Europe", "Global"] = "USA",
        key_fields: Optional[List[str]] = None,
        amount_fields: Optional[List[str]] = None,
        tolerance: Optional[float] = None,
        regulations: Optional[List[str]] = None,
        as_of: Optional[str] = None,
        output_name: Optional[str] = None,
        narrative: Optional[List[Dict[str, str]]] = None,
        __event_emitter__=None,
    ):
        """Reconcile two structured record sets and render an inline IBM-branded report.

        Use this tool when the user has uploaded files to reconcile or pasted
        two tabular datasets to compare. Call ONCE after you have parsed the
        uploaded files into two lists of dicts.

        :param left_records: Left-hand dataset (e.g. GL, blotter, policy ledger, 837 claims, EDC). List of flat dicts.
        :param right_records: Right-hand dataset (e.g. custodian, clearing, claims feed, 835 remittance, CRO). List of flat dicts.
        :param sector: Business sector — drives default key/amount field heuristics and regulation tags.
        :param region: Regulatory region — USA, Europe, or Global.
        :param key_fields: Column names to use as the matching key. If omitted, sector defaults apply.
        :param amount_fields: Column names to compare numerically for variance detection.
        :param tolerance: Absolute tolerance for numeric equality. Defaults to 0.01.
        :param regulations: Governing-regulation tags to display (e.g. ["BCBS 239", "EMIR"]). Tags only; no clause citation.
        :param as_of: Reporting date (YYYY-MM-DD) shown in the header.
        :param output_name: Base filename (without extension) for downloads.
        :param narrative: Optional list of {severity, regulation, text} commentary items.
        :return: Inline HTMLResponse with interactive report and CDN/ooXML download buttons.
        """
        log.info(
            "[recon] invoked: left=%s (%s items), right=%s (%s items), sector=%r, region=%r, key_fields=%r, amount_fields=%r, tolerance=%r, narrative_len=%s",
            type(left_records).__name__,
            (len(left_records) if isinstance(left_records, list) else "n/a"),
            type(right_records).__name__,
            (len(right_records) if isinstance(right_records, list) else "n/a"),
            sector,
            region,
            key_fields,
            amount_fields,
            tolerance,
            (len(narrative) if isinstance(narrative, list) else "n/a"),
        )
        if __event_emitter__:
            await __event_emitter__({"type": "status", "data": {"description": "Parsing inputs…", "done": False}})

        def _coerce(records: Any, side: str) -> List[Dict[str, Any]]:
            if records is None:
                return []
            if isinstance(records, dict):
                records = [records]
            if not isinstance(records, list):
                raise ValueError(
                    f"'{side}_records' must be a JSON array of flat objects (dicts). "
                    "Please upload the file or paste the table so the assistant can parse it into record objects."
                )
            out: List[Dict[str, Any]] = []
            for i, row in enumerate(records):
                if isinstance(row, dict):
                    out.append(row)
                elif isinstance(row, str):
                    raise ValueError(
                        f"'{side}_records[{i}]' is a string, not an object. "
                        "Parse each uploaded file into a list of {{column: value}} dicts before calling the tool."
                    )
                else:
                    raise ValueError(
                        f"'{side}_records[{i}]' must be an object with column-name keys; got {type(row).__name__}."
                    )
            return out

        try:
            left = _coerce(left_records, "left")
            right = _coerce(right_records, "right")
        except ValueError as ve:
            log.warning("[recon] coerce error: %s", ve)
            return HTMLResponse(
                content=_error_shell(str(ve)),
                headers={"Content-Disposition": "inline"},
            )

        if not left and not right:
            return HTMLResponse(
                content=_error_shell(
                    "No records supplied. Please upload two tabular files (CSV / XLSX / TSV) to reconcile, "
                    "or paste the two tables into the chat. The assistant will parse them and re-invoke this tool."
                ),
                headers={"Content-Disposition": "inline"},
            )

        try:
            return await _do_reconcile(
                self.valves, left, right, sector, region, key_fields, amount_fields,
                tolerance, regulations, as_of, output_name, narrative, __event_emitter__,
            )
        except Exception as exc:
            tb = traceback.format_exc()
            log.error("[recon] unhandled exception: %s\n%s", exc, tb)
            return HTMLResponse(
                content=_error_shell(
                    f"Reconciliation failed with: {type(exc).__name__}: {exc}. "
                    "Check Open WebUI server logs for full traceback (grep for '[recon]')."
                ),
                headers={"Content-Disposition": "inline"},
            )


async def _do_reconcile(
    valves,
    left,
    right,
    sector,
    region,
    key_fields,
    amount_fields,
    tolerance,
    regulations,
    as_of,
    output_name,
    narrative,
    event_emitter,
):
    defaults = _sector_defaults(sector)
    kf = key_fields or defaults["key_fields"]
    af = amount_fields or defaults["amount_fields"]
    tol = float(tolerance) if tolerance is not None else valves.default_tolerance
    regs = regulations or defaults["regulations"].get(region, defaults["regulations"].get("USA", []))

    if event_emitter:
        await event_emitter({"type": "status", "data": {"description": f"Matching {len(left)} vs {len(right)} records on {', '.join(kf)}…", "done": False}})

    result = _reconcile(left, right, kf, af, tol)

    redacted = False
    if sector.lower() in ("healthcare", "pharma clinical", "pharma_clinical", "pharmaceutical"):
        for bucket in ("matched", "variance", "unmatched_left", "unmatched_right"):
            result[bucket] = _redact_pii(result[bucket], sector)
        redacted = True

    cap = valves.max_preview_rows
    for bucket in ("matched", "variance", "unmatched_left", "unmatched_right"):
        if len(result[bucket]) > cap:
            result[bucket] = result[bucket][:cap]

    run_id = uuid.uuid4().hex[:8]
    meta = {
        "title": f"AI Reconciliation — {sector} ({region})",
        "sector": sector,
        "region": region,
        "regulations": regs,
        "asOf": as_of or time.strftime("%Y-%m-%d"),
        "runId": run_id,
        "outputName": output_name or f"reconciliation_{sector.lower().replace(' ', '_')}",
        "redacted": redacted,
    }
    payload = {
        "meta": meta,
        "stats": result["stats"],
        "matched": result["matched"],
        "variance": result["variance"],
        "unmatched_left": result["unmatched_left"],
        "unmatched_right": result["unmatched_right"],
        "narrative": narrative or [],
    }

    if event_emitter:
        await event_emitter({"type": "status", "data": {"description": "Building ooXML server fallback (XLSX / DOCX / PPTX)…", "done": False}})

    base_name = f"{meta['outputName']}_{meta['asOf'].replace('-', '')}"
    server_fallback = _build_server_fallback(payload, base_name)

    html = _shell_html(meta, server_fallback, inline_payload=payload)

    if event_emitter:
        await event_emitter({
            "type": "status",
            "data": {
                "description": (
                    f"Reconciliation complete — {result['stats']['match_rate']}% match rate · "
                    f"{result['stats']['variance']} variance · "
                    f"{result['stats']['unmatched_left']+result['stats']['unmatched_right']} unmatched."
                ),
                "done": True,
            },
        })
    return HTMLResponse(content=html, headers={"Content-Disposition": "inline"})


# ---------------------------------------------------------------------------
# Sector defaults — key fields, amount fields, governing regulations
# ---------------------------------------------------------------------------


def _sector_defaults(sector: str) -> Dict[str, Any]:
    s = sector.lower()
    if s == "banking":
        return {
            "key_fields": ["TransactionID"],
            "amount_fields": ["Amount"],
            "regulations": {
                "USA": ["FFIEC 031", "Reg W", "BCBS 239"],
                "Europe": ["CRR/CRD IV", "FINREP", "BCBS 239"],
                "Global": ["BCBS 239"],
            },
        }
    if s == "investment banking":
        return {
            "key_fields": ["TradeID"],
            "amount_fields": ["NetAmount", "Quantity", "Price"],
            "regulations": {
                "USA": ["SEC 10-Q", "FINRA TRACE", "Dodd-Frank"],
                "Europe": ["MiFID II RTS 22", "EMIR"],
                "Global": ["IOSCO"],
            },
        }
    if s == "insurance":
        return {
            "key_fields": ["PolicyNumber"],
            "amount_fields": ["GrossPremium", "NetPremium", "ClaimAmount", "Paid"],
            "regulations": {
                "USA": ["NAIC SAP", "ORSA"],
                "Europe": ["Solvency II", "IFRS 17"],
                "Global": ["IFRS 17"],
            },
        }
    if s == "healthcare":
        return {
            "key_fields": ["ClaimID"],
            "amount_fields": ["BilledAmount", "Paid", "Allowed", "Charged"],
            "regulations": {
                "USA": ["HIPAA 837/835", "CMS-1500"],
                "Europe": ["GDPR", "EHDS"],
                "Global": ["HL7 FHIR"],
            },
        }
    if s == "asset management":
        return {
            "key_fields": ["AccountID", "ISIN"],
            "amount_fields": ["Quantity", "MarketValue"],
            "regulations": {
                "USA": ["SEC 13F", "Investment Company Act"],
                "Europe": ["UCITS", "AIFMD"],
                "Global": ["IOSCO"],
            },
        }
    if s in ("pharma clinical", "pharma_clinical", "pharmaceutical"):
        return {
            "key_fields": ["SubjectID", "VisitDate"],
            "amount_fields": ["SystolicBP", "DiastolicBP", "HeartRate"],
            "regulations": {
                "USA": ["FDA 21 CFR Part 11", "ICH GCP E6(R3)"],
                "Europe": ["EMA EudraCT", "GDPR"],
                "Global": ["ICH GCP"],
            },
        }
    return {
        "key_fields": ["ID"],
        "amount_fields": ["Amount"],
        "regulations": {"USA": [], "Europe": [], "Global": []},
    }
