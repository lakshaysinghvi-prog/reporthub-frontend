import React, { useState, useMemo, useRef, useEffect, useCallback } from "react";
import _ from "lodash";
// API module — calls Railway backend. Stripped in standalone HTML (functions are global there).
import { login as apiLogin, logout as apiLogout, getUsers, createUser, updatePassword, updateRole,
         deleteUser, approveUser, getReports, createReport, updateReport, updateReportConfig, deleteReport as apiDeleteReport,
         publishReport as apiPublishReport, unpublishReport as apiUnpublishReport,
         getReportData, fetchUrlViaProxy,
         getOAuthStatus, startMicrosoftAuth, startGoogleAuth, disconnectOAuth,
         getCustomCredentials, saveCustomCredentials, testCustomCredentials, deleteCustomCredentials,
         getReportAccess, setReportAccess,
         getPublishedReports, getPublishedReportData,
         getRefreshSchedule, setRefreshSchedule,
         getReportFields,
         toggleCollab, getCollabColumns, createCollabColumn, updateCollabColumn, deleteCollabColumn,
         getCollabCycles, openCollabCycle, closeCollabCycle, renameCollabCycle, deleteCollabCycle, reopenCollabCycle,
         exportCollabCycle,
         getCollabValues, upsertCollabValue, submitCollabValue, reviewCollabValue,
         getCollabAudit, getCollabHistory } from "./api.js";

// ── Palette (warm maroon / cream - matches vendor dashboard reference) ─────────
const T = {
  bgPage:   "#F0E8DC", bgCard:   "#FFFFFF", bgHeader: "#5C2D1A",
  bgStat:   "#FBF5EE", bgAlt:    "#F5EEE4", bgTableH: "#EDE0CF",
  border:   "#D4BEA0", borderDk: "#A07850", borderHd: "#7A4520",
  primary:  "#5C2D1A", secondary:"#8B5E3C", accent:   "#C8922A",
  active:   "#4A1F10", text:     "#2C1810", textMd:   "#7A5C4A",
  textLt:   "#F5EFE6", numColor: "#4A2010", success:  "#2D6A4F",
  danger:   "#A32D2D", warning:  "#BA7517",
  tagR:"#534AB7", tagC:"#0F6E56", tagV:"#8B5A2B", tagF:"#185FA5", tagK:"#4A3060",
};

// ── Parse JWT payload to extract user id ──────────────────────────────────
function getMyUserId() {
  try {
    const t = localStorage.getItem('rh_token');
    if (!t) return null;
    return JSON.parse(atob(t.split('.')[1])).id || null;
  } catch(e) { return null; }
}

// ── useViewport hook — true if mobile (<700px) ────────────────────────────────
function useViewport() {
  const [isMobile, setIsMobile] = useState(typeof window !== "undefined" ? window.innerWidth < 700 : false);
  useEffect(() => {
    const onResize = () => setIsMobile(window.innerWidth < 700);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);
  return isMobile;
}


// ── Number formats ─────────────────────────────────────────────────────────────
// Extract a human-friendly source label from a cloud storage URL
function getSourceLabel(url) {
  if (!url) return '☁ Cloud';
  if (url.includes('drive.google.com') || url.includes('docs.google.com')) return '☁ Google Drive';
  if (url.includes('sharepoint.com') || url.includes('my.sharepoint.com')) return '☁ SharePoint';
  if (url.includes('onedrive.live.com') || url.includes('1drv.ms')) return '☁ OneDrive';
  if (url.includes('dropbox.com')) return '☁ Dropbox';
  try { return '☁ ' + new URL(url).hostname.replace('www.',''); } catch(e) { return '☁ Cloud'; }
}

const NUM_FORMATS = [
  { key:"Cr",    label:"Crores",    div:1e7, suffix:" Cr", dec:2 },
  { key:"L",     label:"Lakhs",     div:1e5, suffix:" L",  dec:2 },
  { key:"M",     label:"Millions",  div:1e6, suffix:" M",  dec:2 },
  { key:"K",     label:"Thousands", div:1e3, suffix:" K",  dec:1 },
  { key:"units", label:"Units",     div:1,   suffix:"",    dec:0 },
];

const AGGS=["sum","avg","count","min","max","general"];
const MAX_ROWS=100000, DRILL_PAGE=25, SLICER_SEARCH=30, SLICER_MAX=500, BLANK_THRESH=0.95;
const isMoneyField=f=>/sale|revenue|profit|price|amount|cost|income|spend|budget|fee|net|gross|pay|earn|cash|value|due|paid|deduct|bill/i.test(f);

function fmtNum(n, agg, field, fmtKey, isCurrency) {
  // "general" = raw number exactly as-is, no currency, no scaling (like Excel General)
  if (agg === "general") {
    if (Number.isInteger(n)) return n.toLocaleString();
    return parseFloat(n.toFixed(4)).toLocaleString(undefined,{maximumFractionDigits:4});
  }
  if (agg === "count") return Math.round(n).toLocaleString();
  const fmt = NUM_FORMATS.find(f => f.key === fmtKey) || NUM_FORMATS[4];
  const pfx = isMoneyField(field) ? "\u20B9" : "";
  if (fmt.key === "units") return pfx + Math.round(n).toLocaleString();
  const v = n / fmt.div;
  return pfx + v.toFixed(fmt.dec) + fmt.suffix;
}

// ── CDN loader ─────────────────────────────────────────────────────────────────
function useLibs() {
  const [libs, setLibs] = useState({ XLSX:null, Papa:null });
  useEffect(() => {
    const st = { XLSX:window.XLSX||null, Papa:window.Papa||null };
    const tick = () => { if (st.XLSX && st.Papa) setLibs({XLSX:st.XLSX, Papa:st.Papa}); };
    tick();
    if (!st.XLSX) { const s=document.createElement("script"); s.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"; s.onload=()=>{st.XLSX=window.XLSX;tick();}; document.head.appendChild(s); }
    if (!st.Papa) { const s=document.createElement("script"); s.src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"; s.onload=()=>{st.Papa=window.Papa;tick();}; document.head.appendChild(s); }
  }, []);
  return libs;
}

// ── Sanitization ───────────────────────────────────────────────────────────────
const junkRe=/^(__EMPTY|Column\d+|Unnamed:\s*\d+|undefined)(\s*_\d+)?$/i;
const fmtDate=d=>{try{const y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,"0"),dd=String(d.getDate()).padStart(2,"0");return y+"-"+m+"-"+dd;}catch(e){return "";}};

function sanitizeRows(rawRows) {
  if (!rawRows.length) return {rows:[],fields:[]};
  const rawFields=Object.keys(rawRows[0]), colMap={}, seen={};
  rawFields.forEach(k=>{
    let c=String(k).trim().replace(/\s+/g," ");
    if (!c||junkRe.test(c)){colMap[k]=null;return;}
    if (seen[c]){seen[c]++;colMap[k]=c+" ("+(seen[c])+")";}
    else{seen[c]=1;colMap[k]=c;}
  });
  const good=rawFields.filter(k=>colMap[k]);
  const mapped=rawRows.map(row=>{
    const out={};
    good.forEach(k=>{
      const v=row[k],key=colMap[k];
      if (v instanceof Date) out[key]=fmtDate(v);
      else if (v===null||v===undefined) out[key]="";
      else if (typeof v==="number") out[key]=isFinite(v)?v:"";
      else { const s=String(v).trim(); out[key]=/^\s*-\s*$/.test(s)?"":s; }
    });
    return out;
  });
  const cleanFields=good.map(k=>colMap[k]), nCols=cleanFields.length;
  const rows=mapped.filter(row=>{
    const empty=cleanFields.filter(f=>row[f]===""||row[f]===null||row[f]===undefined).length;
    return empty/nCols<BLANK_THRESH;
  });
  const usedFields=cleanFields.filter(f=>{
    const empty=rows.filter(r=>r[f]===""||r[f]===null||r[f]===undefined).length;
    return (empty/Math.max(rows.length,1))<0.98;
  });
  const finalRows=rows.map(row=>{const o={};usedFields.forEach(f=>{o[f]=row[f];});return o;});
  return {rows:finalRows, fields:usedFields};
}

function detectNumFields(rows,fields) {
  const nums=new Set();
  const sample=rows.slice(0,300);
  fields.forEach(f=>{
    const vals=sample.map(r=>r[f]).filter(v=>v!==""&&v!==null&&v!==undefined);
    if (!vals.length) return;
    const nc=vals.filter(v=>{
      if (typeof v==="number") return true;
      const s=String(v).trim().replace(/[$,\u20B9]/g,"");
      return !isNaN(parseFloat(s))&&isFinite(s)&&!/^0\d{3,}/.test(s);
    }).length;
    if (nc/vals.length>=0.75) nums.add(f);
  });
  return nums;
}

function autoConfig(fields, numFields, name) {
  const dims=fields.filter(f=>!numFields.has(f));
  const nums=fields.filter(f=>numFields.has(f));
  return {
    name:name||"New Report",
    rows:dims.slice(0,1), columns:dims.length>1?dims.slice(1,2):[],
    values:nums.slice(0,3).map(f=>({field:f,agg:"sum"})),
    filters:dims.slice(dims.length>1?2:1,5)
  };
}

// ── Sample data ────────────────────────────────────────────────────────────────
let _s=271828;
const rng=()=>{_s=(_s*214013+2531011)&0x7fffffff;return _s/0x7fffffff;};
const SR=["North","South","East","West"],SC=["Electronics","Clothing","Food","Home"];
const SP={Electronics:["Laptop","Phone","Tablet","Earbuds"],Clothing:["Jacket","Shoes","T-Shirt","Jeans"],Food:["Coffee","Snacks","Juice","Tea"],Home:["Lamp","Chair","Cushion","Planter"]};
const SM=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const SQ={Jan:"Q1",Feb:"Q1",Mar:"Q1",Apr:"Q2",May:"Q2",Jun:"Q2",Jul:"Q3",Aug:"Q3",Sep:"Q3",Oct:"Q4",Nov:"Q4",Dec:"Q4"};
function makeSample(){
  const rows=[];
  SR.forEach(r=>SC.forEach(c=>SP[c].forEach(p=>SM.forEach(m=>{
    const s=2000+Math.round(rng()*8000),u=20+Math.round(rng()*180);
    rows.push({Region:r,Category:c,Product:p,Month:m,Quarter:SQ[m],Sales:s,Units:u,Profit:Math.round(s*(0.15+rng()*0.3))});
  }))));
  const fields=Object.keys(rows[0]),numFields=new Set(["Sales","Units","Profit"]);
  return{rows,fields,numFields,config:{name:"Sales Performance Report",rows:["Region"],columns:["Quarter"],values:[{field:"Sales",agg:"sum"},{field:"Units",agg:"sum"}],filters:["Category","Product"]}};
}

// ── Pivot engine ──────────────────────────────────────────────────────────────
function doAgg(rows,field,type){
  if (!rows.length) return 0;
  const v=rows.map(r=>{
    const x=r[field];
    if (typeof x==="number") return x;
    const n=parseFloat(String(x||"").replace(/[$,\u20B9]/g,""));
    return isNaN(n)?0:n;
  });
  if (type==="sum") return _.sum(v);
  if (type==="avg") return _.mean(v);
  if (type==="count") return rows.length;
  if (type==="min") return Math.min(...v);
  if (type==="max") return Math.max(...v);
  return _.sum(v);
}

function runPivot(data,config,filters) {
  try {
    // Filter by ALL active filters — configured slicers AND card filter clicks
    const allFilterKeys=[...new Set([...config.filters,...Object.keys(filters).filter(k=>filters[k]&&filters[k].length)])];
    const filtered=data.filter(row=>allFilterKeys.every(f=>{
      const s=filters[f];
      if(s==null||!Array.isArray(s)||s.length===0) return true;
      // Special numeric filter values
      if(Array.isArray(s)&&(s.includes("__has__")||s.includes("__zero__"))){
        const v=row[f];
        const isZero = v===null||v===undefined||v===""||Number(v)===0;
        if(s.includes("__has__")&&!isZero) return true;
        if(s.includes("__zero__")&&isZero) return true;
        return false;
      }
      return Array.isArray(s)&&s.includes(String(row[f]||""));
    }));
    const rFs=config.rows, cF=config.columns[0], vals=config.values;
    if (!rFs.length||!vals.length) return null;
    const compute=sub=>vals.map(v=>doAgg(sub,v.field,v.agg));
    // Normalise: group case-insensitively + trim, keep first-seen original case for display
    const seenRk=new Map();
    filtered.forEach(r=>{
      const k=rFs.map(f=>String(r[f]||"").trim().toLowerCase()).join("\0");
      if (!seenRk.has(k)) seenRk.set(k,rFs.map(f=>String(r[f]||"").trim()));
    });
    const rowKeys=[...seenRk.values()].sort((a,b)=>a.join("\0").localeCompare(b.join("\0")));
    const colValSeen=new Map();
    if(cF) filtered.forEach(r=>{const raw=String(r[cF]||"").trim();const k=raw.toLowerCase();if(!colValSeen.has(k))colValSeen.set(k,raw);});
    const colVals=cF?[...colValSeen.values()].sort():[];
    const norm=s=>String(s||"").trim().toLowerCase();
    const cells={};
    rowKeys.forEach(rk=>{
      const rkStr=rk.join("\0");
      const rd=filtered.filter(r=>rFs.every((f,i)=>norm(r[f])===norm(rk[i])));
      cells[rkStr]={};
      colVals.forEach(cv=>{cells[rkStr][cv]=compute(rd.filter(r=>norm(r[cF])===norm(cv)));});
      cells[rkStr]["__total__"]=compute(rd);
    });
    const colTotals={};
    colVals.forEach(cv=>{colTotals[cv]=compute(filtered.filter(r=>norm(r[cF])===norm(cv)));});
    return{rowKeys,colVals,cells,colTotals,grandTotals:compute(filtered),rFs,cF,vals,count:filtered.length};
  } catch(e){return{error:e.message};}
}

// ── Export helpers ─────────────────────────────────────────────────────────────
function exportExcel(result, config, numFmt) {
  if (!window.XLSX) { alert("XLSX library not loaded yet. Please wait a moment."); return; }
  const XLSX = window.XLSX;
  const {rowKeys, colVals, cells, grandTotals, colTotals, rFs, cF, vals} = result;
  const hasGroups = colVals.length > 0;
  // Build header rows
  const hdr1 = rFs.join(" / ") + (cF ? " by " + cF : "");
  const rows = [];
  // Column header row
  const colHdr = [...rFs.map(()=>"")];
  if (hasGroups) {
    colVals.forEach(cv => vals.forEach(v => colHdr.push(cv + " - " + v.field)));
    vals.forEach(v => colHdr.push("Total - " + v.field));
  } else {
    vals.forEach(v => colHdr.push(v.field + " (" + v.agg + ")"));
  }
  rows.push(colHdr);
  // Data rows
  rowKeys.forEach(rk => {
    const rkStr = rk.join("\0");
    const row = [...rk];
    if (hasGroups) {
      colVals.forEach(cv => vals.forEach((_,vi) => row.push(((cells[rkStr]||{})[cv]||[])[vi]||0)));
      vals.forEach((_,vi) => row.push(((cells[rkStr]||{})["__total__"]||[])[vi]||0));
    } else {
      vals.forEach((_,vi) => row.push(((cells[rkStr]||{})["__total__"]||[])[vi]||0));
    }
    rows.push(row);
  });
  // Grand total row
  const gtRow = [...rFs.map((f,i)=>i===0?"Grand Total":"")];
  if (hasGroups) {
    colVals.forEach(cv => vals.forEach((_,vi) => gtRow.push((colTotals[cv]||[])[vi]||0)));
    vals.forEach((_,vi) => gtRow.push(grandTotals[vi]||0));
  } else {
    vals.forEach((_,vi) => gtRow.push(grandTotals[vi]||0));
  }
  rows.push(gtRow);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, config.name.slice(0,31));
  XLSX.writeFile(wb, config.name.replace(/[\/:*?"<>|]/g,"-") + ".xlsx");
}

function exportPDF(config) {
  const style = `<style>
    body{font-family:Arial,sans-serif;font-size:11px;color:#2C1810;background:#fff}
    h2{color:#5C2D1A;margin-bottom:4px;font-size:16px}
    p{color:#7A5C4A;font-size:10px;margin-bottom:12px}
    table{border-collapse:collapse;width:100%}
    th{background:#5C2D1A;color:#F5EFE6;padding:7px 10px;text-align:right;font-size:10px}
    th:first-child{text-align:left}
    td{padding:6px 10px;border-bottom:1px solid #D4BEA0;text-align:right;font-size:10px}
    td:first-child{text-align:left;font-weight:600}
    tr:nth-child(even) td{background:#F5EEE4}
    tfoot td{font-weight:700;background:#EDE0CF;border-top:2px solid #A07850}
    @media print{body{margin:0}}
  </style>`;
  // Find the pivot table — look for the main report area table, not drill-down
  // The pivot table is inside a div with overflowX:auto that is NOT inside a fixed modal
  let tableEl = null;
  const allTables = document.querySelectorAll("table");
  for (const t of allTables) {
    // Skip tables inside fixed-position modals (drill-down, settings)
    let el = t.parentElement;
    let inModal = false;
    while (el) {
      const pos = getComputedStyle(el).position;
      if (pos === "fixed") { inModal = true; break; }
      el = el.parentElement;
    }
    if (!inModal && t.querySelector("thead th")) { tableEl = t; break; }
  }
  if (!tableEl) { alert("No pivot table found. Make sure a report is loaded."); return; }
  const win = window.open("","_blank","width=900,height=700");
  if (!win) { alert("Pop-up blocked. Please allow pop-ups for this site and try again."); return; }
  win.document.write("<html><head><title>"+config.name+"</title>"+style+"</head><body>");
  win.document.write("<h2>"+config.name+"</h2>");
  win.document.write("<p>Exported "+new Date().toLocaleString()+"</p>");
  win.document.write(tableEl.outerHTML);
  win.document.write("</body></html>");
  win.document.close();
  setTimeout(()=>win.print(), 600);
}

// ── Drill-down column filter (Excel-style per-column filter in drill-down) ────
function DrillColFilter({field, data, active, onChange, numFields, activeSort, onSort}) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const sortDir = activeSort||"az";
  const setSortDir = dir => { onSort&&onSort(field,dir); };
  const ref = useRef(null);
  const panelRef = useRef(null);
  const searchRef = useRef(null);
  const looksNum = numFields.has(field);
  const rawOpts = useMemo(()=>data&&data.length?_.uniq(data.map(r=>String(r[field]??""))):[]  ,[data,field]);
  const [noneMode, setNoneMode] = useState(false);

  // ── Draggable panel position ─────────────────────────────────────────────────
  const [pos, setPos] = useState(null); // {x, y} — null means use default anchor
  const dragState = useRef(null); // {startX, startY, origX, origY}

  const startDrag = (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    const panel = panelRef.current;
    if (!panel) return;
    const rect = panel.getBoundingClientRect();
    dragState.current = { startX: e.clientX, startY: e.clientY, origX: rect.left, origY: rect.top };
    const onMove = (ev) => {
      if (!dragState.current) return;
      const dx = ev.clientX - dragState.current.startX;
      const dy = ev.clientY - dragState.current.startY;
      const newX = Math.max(0, Math.min(window.innerWidth - rect.width, dragState.current.origX + dx));
      const newY = Math.max(0, Math.min(window.innerHeight - 80, dragState.current.origY + dy));
      setPos({x: newX, y: newY});
    };
    const onUp = () => {
      dragState.current = null;
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  };

  // Reset position when closed
  useEffect(() => { if (!open) { setPos(null); setSearch(""); } }, [open]);

  const sorted = useMemo(()=>{
    const o=[...rawOpts];
    if (sortDir==="az") o.sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
    else if (sortDir==="za") o.sort((a,b)=>b.localeCompare(a,undefined,{numeric:true}));
    else if (sortDir==="09") o.sort((a,b)=>parseFloat(a||0)-parseFloat(b||0));
    else o.sort((a,b)=>parseFloat(b||0)-parseFloat(a||0));
    return o;
  },[rawOpts,sortDir]);
  const vis = search ? sorted.filter(o=>o.toLowerCase().includes(search.toLowerCase())).slice(0,200) : sorted.slice(0,200);

  const noFilter = active == null || (Array.isArray(active) && active.length === 0);
  const isAllSelected = noFilter;
  // Always ensure effectiveActive is an array — guard against malformed data
  const effectiveActive = Array.isArray(noneMode ? [] : (noFilter ? rawOpts : (Array.isArray(active) ? active : [])))
    ? (noneMode ? [] : (noFilter ? rawOpts : (Array.isArray(active) ? active : [])))
    : [];
  const toggle = v => {
    if (noneMode) { setNoneMode(false); onChange([v]); return; }
    const cur = noFilter ? rawOpts : (Array.isArray(active) ? active : rawOpts);
    if (!Array.isArray(cur)) { onChange([String(v)]); return; }  // safety guard
    const sv = String(v);
    const next = cur.includes(sv) ? cur.filter(x=>x!==sv) : [...cur,sv];
    onChange(!rawOpts.length||next.length >= rawOpts.length ? undefined : next);
  };
  const applySearch = () => {
    if (!search.trim()) { setSearch(""); return; }
    if (vis.length > 0 && vis.length < rawOpts.length) onChange(vis);
    setSearch(""); setOpen(false);
  };
  const partial = !noFilter && !noneMode && Array.isArray(active) && active.length > 0 && active.length < rawOpts.length;

  useEffect(()=>{
    if (open && searchRef.current) setTimeout(()=>searchRef.current&&searchRef.current.focus(),30);
  },[open]);

  useEffect(()=>{
    if (!open) return;
    const h=e=>{
      if (panelRef.current&&!panelRef.current.contains(e.target)&&
          ref.current&&!ref.current.contains(e.target)) setOpen(false);
    };
    const t=setTimeout(()=>document.addEventListener("click",h),10);
    return()=>{clearTimeout(t);document.removeEventListener("click",h);};
  },[open]);

  // Compute default anchor position from the trigger button
  const getDefaultPos = () => {
    if (!ref.current) return {x:0, y:0};
    const r = ref.current.getBoundingClientRect();
    return {
      x: Math.min(r.left, window.innerWidth - 280),
      y: Math.min(r.bottom + 4, window.innerHeight - 400),
    };
  };
  const panelPos = pos || (open ? getDefaultPos() : {x:0,y:0});

  const isSorted=activeSort&&activeSort!=="az";
  const SortBtn=({dir,label})=>(
    <button onClick={()=>{setSortDir(dir);}} style={{padding:"2px 6px",border:"1px solid "+(sortDir===dir&&isSorted?T.primary:T.border),borderRadius:3,fontSize:10,cursor:"pointer",
      background:sortDir===dir&&isSorted?T.primary:"none",color:sortDir===dir&&isSorted?T.textLt:T.textMd,fontWeight:sortDir===dir&&isSorted?700:400}}>
      {label}
    </button>
  );

  return (
    <div ref={ref} style={{position:"relative",display:"inline-block"}}>
      <button onClick={()=>setOpen(o=>!o)} title={"Filter/sort "+field} style={{
        width:18,height:16,padding:0,border:"none",background:"none",cursor:"pointer",
        color:partial||isSorted?T.accent:T.textMd,fontSize:11,display:"flex",alignItems:"center",justifyContent:"center",
        fontWeight:partial||isSorted?700:400}}>
        {partial?"▼":isSorted?(activeSort==="za"||activeSort==="90"?"↓":"↑"):"⊟"}
      </button>
      {open&&(
        <div ref={panelRef} style={{
          position:"fixed", left:panelPos.x+"px", top:panelPos.y+"px",
          zIndex:9999,background:T.bgCard,border:"1px solid "+T.border,
          borderRadius:8,width:280,
          boxShadow:"0 8px 32px rgba(92,45,26,0.25)",
          overflow:"hidden",userSelect:"none"}}>
          {/* ── Drag handle ── */}
          <div onMouseDown={startDrag}
            style={{display:"flex",alignItems:"center",justifyContent:"space-between",
              padding:"6px 10px",background:T.bgHeader,cursor:"grab",
              borderBottom:"0.5px solid "+T.borderHd}}>
            <span style={{fontSize:11,fontWeight:700,color:T.textLt,
              overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:200}}>
              ⠿ {field}
            </span>
            <button onClick={()=>setOpen(false)}
              style={{background:"rgba(255,255,255,0.15)",border:"none",borderRadius:4,
                color:T.textLt,cursor:"pointer",fontSize:13,width:20,height:20,
                display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
              ×
            </button>
          </div>
          {/* ── Sort ── */}
          <div style={{padding:"7px 10px",background:T.bgStat,borderBottom:"0.5px solid "+T.border}}>
            <div style={{fontSize:10,color:T.textMd,fontWeight:600,marginBottom:4}}>Sort</div>
            <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
              <SortBtn dir="az" label="A→Z"/>
              <SortBtn dir="za" label="Z→A"/>
              {looksNum&&<SortBtn dir="09" label="0→9"/>}
              {looksNum&&<SortBtn dir="90" label="9→0"/>}
            </div>
          </div>
          {/* ── Search ── */}
          <div style={{padding:"6px 10px",borderBottom:"0.5px solid "+T.border}}>
            <input ref={searchRef} value={search} onChange={e=>setSearch(e.target.value)}
              placeholder={"Search "+rawOpts.length+" values..."}
              onKeyDown={e=>{
                if (e.key==="Enter") { e.preventDefault(); applySearch(); }
                if (e.key==="Escape") { setSearch(""); setOpen(false); }
              }}
              style={{width:"100%",padding:"5px 8px",border:"0.5px solid "+T.border,borderRadius:4,
                fontSize:12,background:T.bgCard,color:T.text,boxSizing:"border-box",outline:"none"}}/>
            {search&&(
              <div style={{fontSize:10,color:T.textMd,marginTop:3,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span>{vis.length} match{vis.length!==1?"es":""} · Enter to filter</span>
                <button onClick={applySearch}
                  style={{fontSize:10,fontWeight:700,color:T.primary,background:"none",border:"none",cursor:"pointer"}}>
                  Apply ↵
                </button>
              </div>
            )}
          </div>
          {/* ── Clear / All / None ── */}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",
            padding:"5px 10px",borderBottom:"0.5px solid "+T.border,gap:4}}>
            <button onClick={()=>{setNoneMode(false);setSearch("");onChange(undefined);}}
              style={{fontSize:10,background:"none",border:"none",cursor:"pointer",color:T.textMd}}>Clear</button>
            <span style={{fontSize:10,color:T.textMd}}>
              {search?vis.length+" of "+rawOpts.length:rawOpts.length+" values"}
            </span>
            <button onClick={()=>{setNoneMode(false);setSearch("");onChange(undefined);}}
              style={{fontSize:10,background:"none",border:"none",cursor:"pointer",color:T.primary,fontWeight:600}}>All</button>
            <button onClick={()=>{setNoneMode(true);setSearch("");}}
              style={{fontSize:10,background:"none",border:"none",cursor:"pointer",color:T.textMd}}>None</button>
          </div>
          {/* ── Options list ── */}
          <div style={{maxHeight:220,overflowY:"auto"}}>
            {vis.map(o=>(
              <label key={o} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 10px",
                cursor:"pointer",fontSize:11,
                background:Array.isArray(effectiveActive)&&effectiveActive.includes(o)?"rgba(92,45,26,0.05)":undefined,color:T.text}}>
                <input type="checkbox" checked={Array.isArray(effectiveActive)&&effectiveActive.includes(o)} onChange={()=>toggle(o)}
                  style={{width:12,height:12,accentColor:T.primary}}/>
                <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{o||"(blank)"}</span>
              </label>
            ))}
            {vis.length===0&&<div style={{padding:"10px",fontSize:11,color:T.textMd,textAlign:"center"}}>No matches</div>}
          </div>
        </div>
      )}
    </div>
  );
}


function DrillDown({data,target,fields,numFields,onClose,numFmt,savedHiddenCols,savedColFmts,onSaveHiddenCols,onSaveColFilters,configValues,activeFilters}) {
  const [page,setPage]=useState(0);
  const [pageSize,setPageSize]=useState(25); // 25|50|100|"all"
  const [colFiltersSaved,setColFiltersSaved]=useState(false); // feedback after saving col layout
  const [hiddenCols,setHiddenCols]=useState(()=>new Set(savedHiddenCols||[]));
  const [showColPicker,setShowColPicker]=useState(false);
  const [colFilters,setColFilters]=useState(()=>{
    try{const s=localStorage.getItem("rh_drill_colf_"+(target&&target.rowKey||""));return s?JSON.parse(s):{};}catch(e){return{};}
  }); // {field: [selectedValues]}
  const [rowSort,setRowSort]=useState({}); // {field: dir} — only one active at a time
  const [colWidths,startColResize]=useColResize(130);
  const [layoutSaved,setLayoutSaved]=useState(false); // flash confirmation after save
  const {rowKey,colVal,rFs,cF,metricLabel}=target;
  // Per-column format: restore saved fmts, then default metric fields to builder numFmt
  const [colFmts,setColFmts]=useState(()=>{
    const defaults={};
    (configValues||[]).forEach(v=>{ defaults[v.field]=numFmt; });
    return {...defaults,...(savedColFmts||{})};
  });
  const getColFmt=(f)=>colFmts[f]||"general";
  const setColFmt=(f,fmt)=>setColFmts(p=>({...p,[f]:fmt}));
  // Pre-filter by active slicer filters from the parent pivot so drill-down
  // shows the same subset of rows the pivot table is based on
  const filteredBySlicers=useMemo(()=>{
    if (!activeFilters||Object.keys(activeFilters).length===0) return data;
    return data.filter(row=>
      Object.entries(activeFilters).every(([f,sel])=>{
        if (sel==null||!Array.isArray(sel)||sel.length===0) return true;
        if (sel.includes("__has__")){const v=row[f];return v!==null&&v!==undefined&&v!==""&&Number(v)!==0;}
        return sel.includes(String(row[f]??""));
      })
    );
  },[data,activeFilters]);

  const normD=s=>String(s||"").trim().toLowerCase();
  const baseRows=useMemo(()=>filteredBySlicers.filter(row=>
    rFs.every((f,i)=>normD(row[f])===normD(rowKey[i]))&&
    (!cF||!colVal||colVal==="__total__"||normD(row[cF])===normD(colVal))
  ),[filteredBySlicers,target]);
  // Apply per-column filters
  const rows=useMemo(()=>baseRows.filter(row=>
    Object.entries(colFilters).every(([f,sel])=>!sel.length||sel.includes(String(row[f]||"")))
  ),[baseRows,colFilters]);
  const showAll=pageSize==="all";
  const effectivePageSize=showAll?rows.length:pageSize;
  const totalPages=showAll?1:Math.ceil(rows.length/effectivePageSize);
  // Apply active row sort
  const sortedRows=useMemo(()=>{
    const [sf,sd]=Object.entries(rowSort)[0]||[];
    if (!sf) return rows;
    const isNum=numFields.has(sf);
    return [...rows].sort((a,b)=>{
      const av=isNum?+a[sf]||0:String(a[sf]||"");
      const bv=isNum?+b[sf]||0:String(b[sf]||"");
      if (sd==="az"||sd==="09") return isNum?av-bv:String(av).localeCompare(String(bv),undefined,{numeric:true});
      return isNum?bv-av:String(bv).localeCompare(String(av),undefined,{numeric:true});
    });
  },[rows,rowSort,numFields]);
  const visible=showAll?sortedRows:sortedRows.slice(page*effectivePageSize,(page+1)*effectivePageSize);
  // Maintain original field order from source Excel; no cap
  const visibleCols=fields.filter(f=>!hiddenCols.has(f));
  const title=[...rFs.map((f,i)=>f+": "+rowKey[i]),cF&&colVal&&colVal!=="__total__"?cF+": "+colVal:null].filter(Boolean).join(" / ");
  const toggleCol=f=>setHiddenCols(s=>{const n=new Set(s);n.has(f)?n.delete(f):n.add(f);return n;});
  const setColFilter=(f,sel)=>{
    setColFilters(p=>{const n={...p};if(sel==null||!Array.isArray(sel)||sel.length===0)delete n[f];else n[f]=sel;return n;});
    setPage(0);
  };
  const hasColFilters=Object.values(colFilters).some(v=>Array.isArray(v)&&v.length>0);
  // Column totals for visible (filtered) rows
  const colSums=useMemo(()=>{
    const s={};
    visibleCols.forEach(f=>{
      if (numFields.has(f)) s[f]=_.sum(rows.map(r=>+r[f]||0));
    });
    return s;
  },[rows,visibleCols,numFields]);
  return(
    <div style={{position:"fixed",inset:0,zIndex:500,display:"flex",alignItems:"flex-end",background:"rgba(44,24,16,0.5)"}}>
      <div style={{width:"100%",background:T.bgCard,borderRadius:"14px 14px 0 0",boxShadow:"0 -8px 40px rgba(92,45,26,0.25)",maxHeight:"80vh",display:"flex",flexDirection:"column"}}>
        {/* Header */}
        <div style={{padding:"12px 20px",background:T.bgHeader,borderRadius:"14px 14px 0 0",display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
          <div style={{flex:1}}>
            <div style={{fontWeight:700,fontSize:15,color:T.textLt}}>Drill-down: {metricLabel}</div>
            <div style={{fontSize:11,color:"rgba(245,239,230,0.65)",marginTop:2}}>{title}</div>
          </div>
          <span style={{fontSize:12,color:"rgba(245,239,230,0.6)"}}>
            {rows.length.toLocaleString()} of {baseRows.length.toLocaleString()} rows
            {hasColFilters&&<span style={{marginLeft:6,background:"rgba(200,146,42,0.4)",padding:"1px 6px",borderRadius:8,fontSize:10}}>filtered</span>}
            {" · "}{visibleCols.length}/{fields.length} cols
          </span>
          {/* Column visibility picker */}
          <div style={{position:"relative"}}>
            <button onClick={()=>setShowColPicker(p=>!p)}
              style={{padding:"4px 10px",border:"1px solid rgba(255,255,255,0.25)",borderRadius:6,background:"rgba(255,255,255,0.12)",cursor:"pointer",fontSize:11,color:T.textLt,fontWeight:600}}>
              Columns {showColPicker?"v":"v"}
            </button>
            {showColPicker&&(
              <div style={{position:"absolute",right:0,top:"calc(100% + 6px)",background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,width:240,maxHeight:320,overflowY:"auto",boxShadow:"0 8px 24px rgba(92,45,26,0.2)",zIndex:600}}>
                <div style={{padding:"8px 12px",borderBottom:"0.5px solid "+T.border,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <span style={{fontSize:11,fontWeight:700,color:T.primary}}>Show / hide columns</span>
                  <div style={{display:"flex",gap:8}}>
                    {onSaveHiddenCols&&(
                      layoutSaved
                        ? <span style={{fontSize:10,color:T.success,fontWeight:700,display:"flex",alignItems:"center",gap:3}}>
                            ✓ Layout saved
                          </span>
                        : <button onClick={()=>{
                            onSaveHiddenCols([...hiddenCols],colFmts);
                            setLayoutSaved(true);
                            setTimeout(()=>setLayoutSaved(false),2500);
                          }} style={{fontSize:10,color:T.primary,background:"none",border:"1px solid "+T.primary,borderRadius:4,padding:"2px 8px",cursor:"pointer",fontWeight:700}}>
                            Save layout
                          </button>
                    )}
                    <button onClick={()=>setHiddenCols(new Set())} style={{fontSize:10,color:T.textMd,background:"none",border:"none",cursor:"pointer"}}>Show all</button>
                  </div>
                </div>
                {fields.map(f=>(
                  <label key={f} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 12px",cursor:"pointer",fontSize:12,color:T.text,background:hiddenCols.has(f)?"rgba(92,45,26,0.04)":undefined}}>
                    <input type="checkbox" checked={!hiddenCols.has(f)} onChange={()=>toggleCol(f)} style={{accentColor:T.primary,width:13,height:13}}/>
                    <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{f}</span>
                    <span style={{fontSize:10,color:numFields.has(f)?T.tagV:T.textMd,fontWeight:600}}>{numFields.has(f)?"#":"Aa"}</span>
                  </label>
                ))}
              </div>
            )}
          </div>
          <button onClick={onClose} style={{width:28,height:28,borderRadius:6,border:"none",background:"rgba(255,255,255,0.15)",cursor:"pointer",fontSize:16,color:T.textLt,display:"flex",alignItems:"center",justifyContent:"center"}}>x</button>
        </div>
        {/* Hint + status bar */}
        <div style={{fontSize:11,color:T.textMd,padding:"5px 14px",background:layoutSaved?"rgba(45,106,79,0.1)":T.bgStat,borderBottom:"0.5px solid "+T.border,flexShrink:0,display:"flex",alignItems:"center",gap:10,transition:"background 0.3s"}}>
          {layoutSaved
            ? <span style={{color:T.success,fontWeight:600,display:"flex",alignItems:"center",gap:5}}>
                ✓ Column layout saved — this will be remembered when the report is saved
              </span>
            : <span>Columns in original Excel order · Click ⋏ on headers to filter/sort · Scroll right to see all</span>
          }
          {!layoutSaved&&hasColFilters&&<button onClick={()=>setColFilters({})} style={{fontSize:10,color:T.danger,background:"none",border:"none",cursor:"pointer",textDecoration:"underline",flexShrink:0}}>Clear column filters</button>}
          {onSaveColFilters&&hasColFilters&&(
            colFiltersSaved
              ? <span style={{fontSize:10,color:T.success,fontWeight:700,flexShrink:0}}>✓ Saved</span>
              : <button onClick={()=>{onSaveColFilters(colFilters);setColFiltersSaved(true);setTimeout(()=>setColFiltersSaved(false),3000);}}
                  style={{fontSize:10,color:T.primary,background:"none",border:"1px solid "+T.primary,
                    borderRadius:4,padding:"1px 7px",cursor:"pointer",fontWeight:700,flexShrink:0}}>
                  💾 Save col filters
                </button>
          )}
        </div>
        {/* Table — full horizontal scroll, all columns, original order */}
        <div style={{overflowX:"auto",flex:1,overflowY:"auto"}}>
          <table style={{borderCollapse:"collapse",fontSize:12,tableLayout:"fixed",minWidth:"100%"}}>
            <thead style={{position:"sticky",top:0,zIndex:5}}><tr style={{background:T.bgTableH}}>
              {visibleCols.map(f=>{
                const fActive=Array.isArray(colFilters[f])&&colFilters[f].length>0;
                return(
                  <th key={f} style={{padding:"8px 12px",textAlign:numFields.has(f)?"right":"left",fontWeight:700,fontSize:11,
                    color:fActive?T.accent:numFields.has(f)?T.tagV:T.primary,
                    borderBottom:"1px solid "+T.border,whiteSpace:"nowrap",
                    position:"sticky",top:0,background:fActive?"rgba(200,146,42,0.12)":T.bgTableH,zIndex:2,position:"relative",width:colWidths[f]||undefined,minWidth:60}}>
                    <div style={{display:"flex",alignItems:"center",gap:4,justifyContent:numFields.has(f)?"flex-end":"flex-start"}}>
                      <span>{f}</span>
                      <DrillColFilter field={f} data={baseRows} active={colFilters[f]||[]} onChange={sel=>setColFilter(f,sel)} numFields={numFields}
                        activeSort={rowSort[f]} onSort={(fld,dir)=>setRowSort({[fld]:dir})}/>
                    </div>
                    <ResizeHandle onMouseDown={e=>startColResize(f,e)}/>
                    {numFields.has(f)&&(
                      <select value={getColFmt(f)} onChange={e=>setColFmt(f,e.target.value)}
                        onClick={e=>e.stopPropagation()}
                        style={{fontSize:10,border:"1px solid rgba(245,239,230,0.5)",background:"rgba(0,0,0,0.25)",
                          color:"#fff",cursor:"pointer",padding:"1px 3px",marginTop:3,display:"block",
                          width:"100%",borderRadius:3,fontWeight:600}}>
                        <option value="general">General</option>
                        <option value="Cr">Crores</option>
                        <option value="L">Lakhs</option>
                        <option value="M">Millions</option>
                        <option value="K">Thousands</option>
                        <option value="units">Units</option>
                      </select>
                    )}
                  </th>
                );
              })}
            </tr></thead>
            <tbody>
              {visible.map((row,i)=>(
                <tr key={i} style={{background:i%2===0?T.bgCard:T.bgAlt}}>
                  {visibleCols.map(f=>(
                    <td key={f} style={{padding:"7px 13px",borderBottom:"0.5px solid "+T.border,
                      textAlign:numFields.has(f)?"right":"left",color:T.text,
                      overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                      {row[f]===""||row[f]===null||row[f]===undefined
                        ?<span style={{color:T.textMd}}>-</span>
                        :numFields.has(f)?fmtNum(+row[f],getColFmt(f)==="general"?"general":"sum",f,getColFmt(f)):String(row[f])}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
            {/* Totals row — updates with column filters */}
            <tfoot>
              <tr style={{background:T.bgTableH}}>
                {visibleCols.map((f,i)=>{
                  const isNum=numFields.has(f);
                  return(
                    <td key={f} style={{padding:"7px 12px",fontWeight:700,fontSize:11,
                      textAlign:isNum?"right":"left",borderTop:"1px solid "+T.borderDk,
                      color:isNum?T.primary:T.textMd,background:T.bgTableH,whiteSpace:"nowrap"}}>
                      {i===0?"Total ("+rows.length.toLocaleString()+" rows)":
                        isNum?fmtNum(colSums[f]||0,getColFmt(f)==="general"?"general":"sum",f,getColFmt(f)):""}
                    </td>
                  );
                })}
              </tr>
            </tfoot>
          </table>
        </div>
        {/* Pagination footer */}
        <div style={{padding:"8px 20px",borderTop:"0.5px solid "+T.border,display:"flex",alignItems:"center",gap:10,flexShrink:0,flexWrap:"wrap"}}>
          {/* Page size selector */}
          <div style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:T.textMd}}>
            <span>Show:</span>
            {[25,50,100,"all"].map(sz=>(
              <button key={sz} onClick={()=>{setPageSize(sz);setPage(0);}}
                style={{padding:"3px 9px",border:"1px solid "+(pageSize===sz?T.primary:T.border),borderRadius:5,
                  background:pageSize===sz?T.primary:"none",color:pageSize===sz?T.textLt:T.text,
                  fontSize:11,fontWeight:pageSize===sz?700:400,cursor:"pointer"}}>
                {sz==="all"?"All":sz}
              </button>
            ))}
            <span style={{color:T.textMd}}>rows</span>
          </div>
          <span style={{fontSize:12,color:T.textMd,flex:1,textAlign:"center"}}>
            {showAll
              ? "Showing all "+rows.length.toLocaleString()+" rows"
              : "Page "+(page+1)+" of "+totalPages+" · rows "+(page*effectivePageSize+1)+"–"+Math.min((page+1)*effectivePageSize,rows.length)+" of "+rows.length.toLocaleString()}
          </span>
          {!showAll&&(<>
            <button onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0}
              style={{padding:"4px 12px",border:"0.5px solid "+T.border,borderRadius:5,background:"none",cursor:page===0?"not-allowed":"pointer",opacity:page===0?0.4:1,fontSize:12,color:T.text}}>Prev</button>
            <button onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={page===totalPages-1}
              style={{padding:"4px 12px",border:"0.5px solid "+T.border,borderRadius:5,background:"none",cursor:page===totalPages-1?"not-allowed":"pointer",opacity:page===totalPages-1?0.4:1,fontSize:12,color:T.text}}>Next</button>
          </>)}
        </div>
      </div>
    </div>
  );
}

// ── Quick filter cards ─────────────────────────────────────────────────────────
function QuickFilterCards({field,data,activeFilters,onFilter,primaryVal,numFmt,numFields,cardAgg}) {
  // Determine if this card field is numeric (KPI mode) or dimension (filter mode)
  const isNumericField = numFields && numFields.has(field);
  // For numeric fields, use the card's own configured agg (sum/count/avg/min/max)
  // For dimension fields, use primaryVal metric broken down by dimension values
  const displayVal = isNumericField
    ? {field, agg: cardAgg||"sum"}  // use card-specific agg
    : primaryVal;                    // use primary metric per dimension value

  const opts = useMemo(()=>_.uniq(data.map(r=>String(r[field]||""))).sort(),[data,field]);
  const tooManyOpts = opts.length > 20;
  const defaultMode = (isNumericField || tooManyOpts) ? "summary" : "breakdown";
  const [mode, setMode] = useState(defaultMode);
  const active = activeFilters || [];
  const allActive = active.length === 0;

  const cardBase = {
    flexShrink:0, padding:"10px 14px", borderRadius:8, textAlign:"left",
    cursor:"pointer", border:"1px solid "+T.border, transition:"all 0.15s",
  };
  const cardOn  = {...cardBase, background:T.primary, border:"2px solid "+T.primary,
    boxShadow:"0 2px 8px rgba(92,45,26,0.25)", transform:"translateY(-1px)"};
  const cardOff = {...cardBase, background:T.bgCard};
  const cardKpi = {...cardBase, background:T.bgStat, cursor:"default"};

  // ── NUMERIC FIELD: single clickable tile — toggle has-value filter ────────────
  if (isNumericField) {
    const total = fmtNum(doAgg(data, displayVal.field, displayVal.agg), displayVal.agg, displayVal.field, numFmt);
    // "Has value" rows = field is non-null AND non-zero (positive OR negative)
    const withVal = data.filter(r => {
      const v = r[field];
      return v !== null && v !== undefined && v !== "" && Number(v) !== 0;
    });
    const isOn = active.includes("__has__");
    const displayData = isOn ? withVal : data;
    const displayTotal = fmtNum(doAgg(displayData,displayVal.field,displayVal.agg),displayVal.agg,displayVal.field,numFmt);
    return (
      <div>
        <div style={{fontSize:10,fontWeight:700,color:T.textMd,textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:6}}>
          {field}
        </div>
        <button onClick={()=>onFilter(isOn?[]:["__has__"])}
          title={isOn?"Click to clear — show all rows":"Click to filter — only rows with a value"}
          style={{...(isOn?cardOn:cardOff), minWidth:140, width:"100%", textAlign:"left"}}>
          {isOn && (
            <div style={{fontSize:9,color:"rgba(245,239,230,0.75)",marginBottom:3,fontStyle:"italic"}}>
              Filtered · click to clear
            </div>
          )}
          <div style={{fontSize:9,color:isOn?"rgba(245,239,230,0.7)":T.textMd,marginBottom:2}}>
            {displayVal.agg} of {displayVal.field}
          </div>
          <div style={{fontSize:17,fontWeight:700,color:isOn?T.textLt:T.numColor}}>{displayTotal}</div>
          <div style={{fontSize:9,color:isOn?"rgba(245,239,230,0.6)":T.textMd,marginTop:2}}>
            {displayData.length.toLocaleString()} rows{isOn?" (with value)":""}
          </div>
        </button>
      </div>
    );
  }

  // ── DIMENSION FIELD: summary mode (single card) ───────────────────────────────
  if (mode === "summary") {
    const total = fmtNum(doAgg(data, displayVal.field, displayVal.agg), displayVal.agg, displayVal.field, numFmt);
    return (
      <div>
        <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}>
          <span style={{fontSize:10,fontWeight:700,color:T.textMd,textTransform:"uppercase",letterSpacing:"0.8px"}}>{field}</span>
          {!tooManyOpts && (
            <button onClick={()=>setMode("breakdown")}
              style={{fontSize:9,color:T.primary,background:"none",border:"1px solid "+T.border,borderRadius:3,padding:"1px 6px",cursor:"pointer"}}>
              Expand ▸
            </button>
          )}
          {tooManyOpts && <span style={{fontSize:9,color:T.textMd,fontStyle:"italic"}}>{opts.length} values</span>}
        </div>
        <button onClick={allActive ? undefined : ()=>onFilter([])}
          style={{...(allActive?cardKpi:cardOn), cursor:allActive?"default":"pointer", width:"100%", textAlign:"left"}}>
          {!allActive && (
            <div style={{fontSize:9,color:"rgba(245,239,230,0.75)",marginBottom:3,fontStyle:"italic"}}>
              {active.join(", ")} · click to clear
            </div>
          )}
          <div style={{fontSize:9,color:allActive?T.textMd:"rgba(245,239,230,0.7)",marginBottom:2}}>
            {displayVal.agg} of {displayVal.field}
          </div>
          <div style={{fontSize:17,fontWeight:700,color:allActive?T.numColor:T.textLt}}>{total}</div>
          <div style={{fontSize:9,color:allActive?T.textMd:"rgba(245,239,230,0.6)",marginTop:2}}>
            {data.length.toLocaleString()} rows
          </div>
        </button>
      </div>
    );
  }

  // ── DIMENSION FIELD: breakdown mode (one card per unique value) ───────────────
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}>
        <span style={{fontSize:10,fontWeight:700,color:T.textMd,textTransform:"uppercase",letterSpacing:"0.8px"}}>{field}</span>
        <button onClick={()=>setMode("summary")}
          style={{fontSize:9,color:T.textMd,background:"none",border:"1px solid "+T.border,borderRadius:3,padding:"1px 6px",cursor:"pointer"}}>
          ◂ Collapse
        </button>
        {!allActive && (
          <button onClick={()=>onFilter([])}
            style={{fontSize:9,color:T.textMd,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>
            Clear
          </button>
        )}
      </div>
      <div style={{display:"flex",gap:6,overflowX:"auto",paddingBottom:4}}>
        {/* All button */}
        <button onClick={()=>onFilter([])} style={allActive?{...cardOn,minWidth:80}:{...cardOff,minWidth:80}}>
          <div style={{fontSize:9,color:allActive?"rgba(245,239,230,0.7)":T.textMd,marginBottom:2}}>All</div>
          <div style={{fontSize:14,fontWeight:700,color:allActive?T.textLt:T.numColor}}>
            {fmtNum(doAgg(data,displayVal.field,displayVal.agg),displayVal.agg,displayVal.field,numFmt)}
          </div>
          <div style={{fontSize:9,color:allActive?"rgba(245,239,230,0.6)":T.textMd,marginTop:2}}>{data.length.toLocaleString()} rows</div>
        </button>
        {/* One card per unique value */}
        {opts.map(val=>{
          const on = active.includes(val);
          const subset = data.filter(r=>String(r[field]||"")===val);
          return (
            <button key={val} onClick={()=>on?onFilter([]):onFilter([val])}
              style={on?{...cardOn,minWidth:80}:{...cardOff,minWidth:80}}>
              <div style={{fontSize:9,color:on?"rgba(245,239,230,0.7)":T.textMd,marginBottom:2,
                maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                {val||"(blank)"}
              </div>
              <div style={{fontSize:14,fontWeight:700,color:on?T.textLt:T.numColor}}>
                {fmtNum(doAgg(subset,displayVal.field,displayVal.agg),displayVal.agg,displayVal.field,numFmt)}
              </div>
              <div style={{fontSize:9,color:on?"rgba(245,239,230,0.6)":T.textMd,marginTop:2}}>
                {subset.length.toLocaleString()} rows
              </div>
            </button>
          );
        })}
      </div>
    </div>
  );
}
function Slicer({field,active,onChange,data}) {
  const [open,setOpen]=useState(false);
  const [search,setSearch]=useState("");
  const [sortDir,setSortDir]=useState("az"); // "az"|"za"|"09"|"90"
  const ref=useRef(null);
  const rawOpts=useMemo(()=>_.uniq(data.map(r=>String(r[field]||""))),[field,data]);
  const [noneMode,setNoneMode]=useState(false);
  // Detect if field looks numeric for numeric sort option
  const looksNumeric=useMemo(()=>{
    const sample=rawOpts.slice(0,20).filter(o=>o!=="");
    return sample.length>0&&sample.filter(o=>!isNaN(parseFloat(o))&&isFinite(o)).length/sample.length>0.7;
  },[rawOpts]);
  const sortedOpts=useMemo(()=>{
    const opts=[...rawOpts];
    if (sortDir==="az") opts.sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
    else if (sortDir==="za") opts.sort((a,b)=>b.localeCompare(a,undefined,{numeric:true}));
    else if (sortDir==="09") opts.sort((a,b)=>parseFloat(a||0)-parseFloat(b||0));
    else opts.sort((a,b)=>parseFloat(b||0)-parseFloat(a||0));
    return opts;
  },[rawOpts,sortDir]);
  const tooMany=sortedOpts.length>SLICER_MAX;
  const needsSearch=sortedOpts.length>SLICER_SEARCH;
  const visOpts=search?sortedOpts.filter(o=>o.toLowerCase().includes(search.toLowerCase())).slice(0,300):sortedOpts.slice(0,300);
  // active=undefined → no filter (all rows pass, treated as all-selected)
  // active=[] → explicitly none (nothing passes, all unchecked)
  // active=[x,y] → filter to x and y
  const noFilter = active == null || active === undefined || (Array.isArray(active) && active.length === 0);
  const isAllSelected = noFilter;
  const effectiveActive = noneMode ? [] : (noFilter ? rawOpts : active);
  const toggle=v=>{
    if(noneMode){setNoneMode(false);onChange([v]);return;}
    const cur = noFilter ? rawOpts : active;
    const next=cur.includes(v)?cur.filter(x=>x!==v):[...cur,v];
    // If all explicitly selected → revert to "no filter" (undefined)
    onChange(next.length>=rawOpts.length ? undefined : next);
  };
  const partial=!isAllSelected&&!noneMode&&active&&active.length>0&&active.length<rawOpts.length;
  useEffect(()=>{
    if (!open) return;
    const h=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};
    const t=setTimeout(()=>document.addEventListener("click",h),10);
    return()=>{clearTimeout(t);document.removeEventListener("click",h);};
  },[open]);
  if (tooMany) return(
    <span style={{display:"inline-flex",alignItems:"center",gap:6,padding:"6px 12px",background:T.bgStat,border:"0.5px solid "+T.border,borderRadius:6,fontSize:12,color:T.textMd}}>
      {field} <span style={{fontSize:10}}>({sortedOpts.length.toLocaleString()} - too many)</span>
    </span>
  );
  const SortBtn=({dir,label})=>(
    <button onClick={()=>setSortDir(dir)} style={{padding:"3px 8px",border:"1px solid "+(sortDir===dir?T.primary:T.border),borderRadius:4,fontSize:11,cursor:"pointer",
      background:sortDir===dir?T.primary:"none",color:sortDir===dir?T.textLt:T.textMd,fontWeight:sortDir===dir?700:400}}>
      {label}
    </button>
  );
  return(
    <div ref={ref} style={{position:"relative"}}>
      <button onClick={()=>setOpen(o=>!o)} style={{display:"flex",alignItems:"center",gap:6,
        background:partial?T.primary:T.bgCard,border:"1px solid "+(partial?T.primary:T.border),
        borderRadius:6,padding:"6px 12px",cursor:"pointer",fontSize:13,color:partial?T.textLt:T.text,fontWeight:partial?600:400}}>
        {field}
        {partial&&<span style={{background:"rgba(255,255,255,0.25)",color:T.textLt,borderRadius:10,padding:"1px 7px",fontSize:11,fontWeight:600}}>{active&&active.length}</span>}
        <span style={{fontSize:9,opacity:0.5}}>{open?"▲":"▼"}</span>
      </button>
      {open&&(
        <div style={{position:"absolute",top:"calc(100% + 5px)",left:0,zIndex:9999,background:T.bgCard,border:"1px solid "+T.border,borderRadius:8,minWidth:260,maxWidth:340,boxShadow:"0 8px 28px rgba(92,45,26,0.2)",overflow:"hidden"}}>
          {/* Sort row — like Excel filter */}
          <div style={{padding:"8px 12px",borderBottom:"0.5px solid "+T.border,background:T.bgStat}}>
            <div style={{fontSize:10,color:T.textMd,fontWeight:600,marginBottom:5}}>Sort</div>
            <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
              <SortBtn dir="az" label="A → Z"/>
              <SortBtn dir="za" label="Z → A"/>
              {looksNumeric&&<SortBtn dir="09" label="0 → 9"/>}
              {looksNumeric&&<SortBtn dir="90" label="9 → 0"/>}
            </div>
          </div>
          {/* Search */}
          <div style={{padding:"7px 10px",borderBottom:"0.5px solid "+T.border}}>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder={"Search "+sortedOpts.length+" values..."}
              style={{width:"100%",padding:"5px 9px",border:"0.5px solid "+T.border,borderRadius:5,fontSize:12,background:T.bgCard,color:T.text,boxSizing:"border-box",outline:"none"}}/>
          </div>
          {/* Select all / clear */}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"5px 12px",borderBottom:"0.5px solid "+T.border}}>
            <button onClick={()=>{setNoneMode(false);onChange(undefined);}} style={{fontSize:11,background:"none",border:"none",cursor:"pointer",color:T.textMd}}>Clear all</button>
            <span style={{fontSize:10,color:T.textMd}}>{sortedOpts.length} values</span>
            <button onClick={()=>setNoneMode(true)} style={{fontSize:11,background:"none",border:"none",cursor:"pointer",color:T.textMd}}>None</button>
            <button onClick={()=>{setNoneMode(false);onChange(undefined);}} style={{fontSize:11,background:"none",border:"none",cursor:"pointer",color:T.primary,fontWeight:600}}>All</button>
          </div>
          {/* Checkbox list */}
          <div style={{maxHeight:250,overflowY:"auto"}}>
            {visOpts.map(o=>(
              <label key={o} style={{display:"flex",alignItems:"center",gap:9,padding:"6px 12px",cursor:"pointer",fontSize:12,background:effectiveActive.includes(o)?"rgba(92,45,26,0.05)":undefined,color:T.text}}>
                <input type="checkbox" checked={effectiveActive.includes(o)} onChange={()=>toggle(o)} style={{width:13,height:13,accentColor:T.primary,flexShrink:0}}/>
                <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1}}>{o||"(blank)"}</span>
              </label>
            ))}
            {!search&&sortedOpts.length>300&&<div style={{padding:"6px 12px",fontSize:10,color:T.textMd,borderTop:"0.5px solid "+T.border}}>Showing 300 of {sortedOpts.length} - type to search all</div>}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Column resize hook ────────────────────────────────────────────────────────
// Returns [widths, startResize] — widths is {colKey: px}, startResize(key, e)
function useColResize(defaultWidth=120) {
  const [widths,setWidths]=useState({});
  const startResize=useCallback((key,e)=>{
    e.preventDefault();
    e.stopPropagation();
    const startX=e.clientX;
    const startW=widths[key]||defaultWidth;
    const onMove=me=>{
      const newW=Math.max(50,startW+(me.clientX-startX));
      setWidths(w=>({...w,[key]:newW}));
    };
    const onUp=()=>{
      document.removeEventListener("mousemove",onMove);
      document.removeEventListener("mouseup",onUp);
      document.body.style.cursor="";
      document.body.style.userSelect="";
    };
    document.body.style.cursor="col-resize";
    document.body.style.userSelect="none";
    document.addEventListener("mousemove",onMove);
    document.addEventListener("mouseup",onUp);
  },[widths]);
  return [widths,startResize];
}

// Resize handle element — attach onMouseDown={e=>startResize(key,e)}
const ResizeHandle=({onMouseDown})=>(
  <div onMouseDown={onMouseDown}
    style={{position:"absolute",right:0,top:0,bottom:0,width:6,cursor:"col-resize",
      zIndex:10,display:"flex",alignItems:"center",justifyContent:"center"}}
    title="Drag to resize column">
    <div style={{width:2,height:"60%",background:"rgba(255,255,255,0.3)",borderRadius:1}}/>
  </div>
);

// ── Pivot table ────────────────────────────────────────────────────────────────
// ── Chart View — visualisations of pivot result ─────────────────────────────────
function ChartView({result, numFmt, chartType, onChartTypeChange}) {
  const [topN, setTopN] = useState(20); // how many rows to show before collapsing to "Others"

  if (!result || result.error) return null;
  const { cells, vals, rowKeys, colVals } = result;
  const hasGroups = colVals && colVals.length > 0;
  const nV = vals.length;

  if (!rowKeys.length || !nV) {
    return (
      <div style={{padding:"40px 20px",textAlign:"center",color:T.textMd,background:T.bgCard,
        borderRadius:10,border:"1px solid "+T.border}}>
        Configure rows and values in the builder to see charts.
      </div>
    );
  }

  // Build chart-friendly data: one entry per row
  const allChartData = rowKeys.map(rk => {
    const rkStr = rk.join("\0");
    const cellRow = cells[rkStr] || {};
    const out = { name: rk.join(" · "), values: [] };
    if (hasGroups) {
      colVals.forEach(cv => {
        vals.forEach((v, vi) => {
          out.values.push({
            label: (nV > 1) ? `${cv} — ${v.field}` : String(cv),
            value: ((cellRow[cv] || [])[vi]) || 0,
          });
        });
      });
    } else {
      vals.forEach((v, vi) => {
        out.values.push({ label: v.field, value: ((cellRow["__total__"] || [])[vi]) || 0 });
      });
    }
    return out;
  });

  // Sort by sum of all series values descending — biggest bars always in Top N
  const sorted = [...allChartData].sort((a,b) => {
    const sumA = a.values.reduce((s,v)=>s+(v.value||0),0);
    const sumB = b.values.reduce((s,v)=>s+(v.value||0),0);
    return sumB - sumA;
  });
  const needsGrouping = sorted.length > topN;
  let chartData;
  if (needsGrouping) {
    const top = sorted.slice(0, topN);
    const rest = sorted.slice(topN);
    // Aggregate "Others" by summing each series
    const othersValues = (top[0]?.values||[]).map((_,si) => ({
      label: top[0]?.values[si]?.label || "",
      value: rest.reduce((s,d) => s + (d.values[si]?.value||0), 0),
    }));
    chartData = [...top, { name: `Others (${rest.length})`, values: othersValues, isOthers: true }];
  } else {
    chartData = allChartData; // already reasonable number
  }

  const series = chartData[0] ? chartData[0].values.map(v => v.label) : [];
  const palette = [T.primary, T.accent, "#7B5C3E", "#A07850", "#5C2D1A", "#C8922A", "#8B6B4A", "#3D1F11"];

  const allValues = chartData.flatMap(d => d.values.map(v => v.value));
  const maxV = Math.max(...allValues, 0);
  const minV = Math.min(...allValues, 0);
  const range = maxV - minV || 1;
  const fmtShort = (n) => fmtNum(n, "sum", "", numFmt).replace("₹","").trim();

  const W = 900, H = 400;
  const padL = 75, padR = 20, padT = 30, padB = 90;  // padT = room for value labels
  const innerW = W - padL - padR;
  const innerH = H - padT - padB;
  const groupW = innerW / Math.max(chartData.length, 1);
  const yTicks = 5;
  const yTickValues = Array.from({length:yTicks+1},(_,i)=>minV+(range*i/yTicks));
  const yPos = (v) => padT + innerH - ((v - minV) / range) * innerH;

  let chartEl = null;

  if (chartType === "bar") {
    const barW = Math.max(3, (groupW * 0.8) / Math.max(series.length, 1));
    chartEl = (
      <svg viewBox={`0 0 ${W} ${H}`} style={{width:"100%",height:"auto",maxHeight:400,display:"block"}}>
        {yTickValues.map((tv,i)=>(
          <g key={i}>
            <line x1={padL} x2={W-padR} y1={yPos(tv)} y2={yPos(tv)} stroke={T.border} strokeDasharray="3 3"/>
            <text x={padL-8} y={yPos(tv)+4} fontSize="10" fill={T.textMd} textAnchor="end">{fmtShort(tv)}</text>
          </g>
        ))}
        <line x1={padL} x2={W-padR} y1={yPos(0)} y2={yPos(0)} stroke={T.borderDk}/>
        {chartData.map((d,gi)=>{
          const groupX = padL + gi*groupW + (groupW - barW*series.length)/2;
          const isOthers = d.isOthers;
          return(
            <g key={gi}>
              {d.values.map((v,si)=>{
                const x = groupX + si*barW;
                const y = v.value >= 0 ? yPos(v.value) : yPos(0);
                const h = Math.max(1, Math.abs(yPos(v.value) - yPos(0)));
                return(
                  <g key={si}>
                    <rect x={x} y={y} width={Math.max(1,barW-1)} height={h}
                      fill={isOthers?"#AAAAAA":palette[si%palette.length]} rx={1} opacity={isOthers?0.6:1}>
                      <title>{d.name} · {v.label}: {fmtNum(v.value,"sum","",numFmt)}</title>
                    </rect>
                    {/* Value label above bar — only show if bar is wide enough */}
                    {barW > 18 && v.value !== 0 && (
                      <text x={x + (barW-1)/2} y={y - 3}
                        fontSize="9" fill={T.textMd} textAnchor="middle"
                        style={{pointerEvents:"none"}}>
                        {fmtShort(v.value)}
                      </text>
                    )}
                  </g>
                );
              })}
              <text x={padL+gi*groupW+groupW/2} y={H-padB+14} fontSize={chartData.length>30?8:10} fill={d.isOthers?"#888":T.text}
                textAnchor="middle"
                transform={`rotate(-35 ${padL+gi*groupW+groupW/2} ${H-padB+14})`}>
                {d.name.length>18?d.name.slice(0,18)+"…":d.name}
              </text>
            </g>
          );
        })}
      </svg>
    );
  } else if (chartType === "line" || chartType === "area") {
    chartEl = (
      <svg viewBox={`0 0 ${W} ${H}`} style={{width:"100%",height:"auto",maxHeight:400,display:"block"}}>
        {yTickValues.map((tv,i)=>(
          <g key={i}>
            <line x1={padL} x2={W-padR} y1={yPos(tv)} y2={yPos(tv)} stroke={T.border} strokeDasharray="3 3"/>
            <text x={padL-8} y={yPos(tv)+4} fontSize="10" fill={T.textMd} textAnchor="end">{fmtShort(tv)}</text>
          </g>
        ))}
        <line x1={padL} x2={W-padR} y1={yPos(0)} y2={yPos(0)} stroke={T.borderDk}/>
        {series.map((sLabel,si)=>{
          const points = chartData.map((d,gi)=>{
            const x = padL + gi*groupW + groupW/2;
            const v = d.values[si] ? d.values[si].value : 0;
            return [x, yPos(v)];
          });
          const pathD = points.map((p,i)=>(i===0?"M":"L")+p[0].toFixed(1)+","+p[1].toFixed(1)).join(" ");
          if (chartType === "area") {
            const areaD = pathD + ` L${points[points.length-1][0]},${yPos(0)} L${points[0][0]},${yPos(0)} Z`;
            return(
              <g key={si}>
                <path d={areaD} fill={palette[si%palette.length]} fillOpacity="0.25"/>
                <path d={pathD} stroke={palette[si%palette.length]} strokeWidth="2" fill="none"/>
              </g>
            );
          }
          return(
            <path key={si} d={pathD} stroke={palette[si%palette.length]} strokeWidth="2.5" fill="none"/>
          );
        })}
        {chartData.map((d,gi)=>(
          <text key={gi} x={padL+gi*groupW+groupW/2} y={H-padB+14} fontSize={chartData.length>30?8:10} fill={T.text}
            textAnchor="middle" transform={`rotate(-35 ${padL+gi*groupW+groupW/2} ${H-padB+14})`}>
            {d.name.length>18?d.name.slice(0,18)+"…":d.name}
          </text>
        ))}
      </svg>
    );
  } else if (chartType === "pie") {
    const pieData = chartData.map(d=>({name:d.name,value:d.values[0]?d.values[0].value:0,isOthers:d.isOthers})).filter(d=>d.value>0);
    const total = pieData.reduce((s,d)=>s+d.value,0);
    if (total === 0) {
      chartEl = <div style={{textAlign:"center",padding:40,color:T.textMd}}>No positive values to chart</div>;
    } else {
      const cx=W/2, cy=H/2-10, r=Math.min(innerW,innerH)/2.5;
      let cumAngle=-Math.PI/2;
      const slices=pieData.map((d,i)=>{
        const angle=(d.value/total)*Math.PI*2;
        const x1=cx+r*Math.cos(cumAngle), y1=cy+r*Math.sin(cumAngle);
        const x2=cx+r*Math.cos(cumAngle+angle), y2=cy+r*Math.sin(cumAngle+angle);
        const large=angle>Math.PI?1:0;
        const path=`M${cx},${cy} L${x1.toFixed(1)},${y1.toFixed(1)} A${r},${r} 0 ${large},1 ${x2.toFixed(1)},${y2.toFixed(1)} Z`;
        const labelAngle=cumAngle+angle/2;
        const lx=cx+(r+20)*Math.cos(labelAngle), ly=cy+(r+20)*Math.sin(labelAngle);
        cumAngle+=angle;
        return {path,color:d.isOthers?"#AAAAAA":palette[i%palette.length],lx,ly,label:d.name,value:d.value,pct:(d.value/total*100)};
      });
      chartEl=(
        <svg viewBox={`0 0 ${W} ${H}`} style={{width:"100%",height:"auto",maxHeight:400,display:"block"}}>
          {slices.map((s,i)=>(
            <g key={i}>
              <path d={s.path} fill={s.color} stroke={T.bgCard} strokeWidth="1.5">
                <title>{s.label}: {fmtNum(s.value,"sum","",numFmt)} ({s.pct.toFixed(1)}%)</title>
              </path>
              {s.pct>=3&&(
                <text x={s.lx} y={s.ly} fontSize="10" fill={T.text}
                  textAnchor={s.lx>cx?"start":"end"}>
                  {s.label.length>14?s.label.slice(0,14)+"…":s.label} ({s.pct.toFixed(0)}%)
                </text>
              )}
            </g>
          ))}
        </svg>
      );
    }
  }

  return (
    <div style={{background:T.bgCard,borderRadius:10,border:"1px solid "+T.border,padding:"16px",
      boxShadow:"0 2px 8px rgba(92,45,26,0.08)"}}>
      {/* Header: title + chart type toggle */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
          <span style={{fontWeight:700,fontSize:13,color:T.primary}}>Chart visualisation</span>
          {needsGrouping&&(
            <span style={{fontSize:11,color:T.textMd}}>
              Top {topN} of {allChartData.length} rows · others grouped
            </span>
          )}
        </div>
        <div style={{display:"flex",gap:4,background:T.bgStat,borderRadius:7,padding:3,border:"0.5px solid "+T.border}}>
          {[{k:"bar",l:"📊 Bar"},{k:"line",l:"📈 Line"},{k:"area",l:"📉 Area"},{k:"pie",l:"🥧 Pie"}].map(b=>(
            <button key={b.k} onClick={()=>onChartTypeChange(b.k)}
              style={{padding:"4px 12px",border:"none",borderRadius:5,
                background:chartType===b.k?T.primary:"none",color:chartType===b.k?T.textLt:T.textMd,
                cursor:"pointer",fontSize:11,fontWeight:600}}>
              {b.l}
            </button>
          ))}
        </div>
      </div>

      {/* Top N slider — only shown when there are many rows */}
      {allChartData.length > 10&&(
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,
          padding:"7px 12px",background:T.bgStat,borderRadius:7,border:"0.5px solid "+T.border}}>
          <span style={{fontSize:11,color:T.textMd,whiteSpace:"nowrap"}}>Show top</span>
          <input type="range" min={5} max={Math.min(100,allChartData.length)} step={5} value={topN}
            onChange={e=>setTopN(Number(e.target.value))}
            style={{flex:1,accentColor:T.primary,cursor:"pointer"}}/>
          <span style={{fontSize:12,fontWeight:700,color:T.primary,minWidth:28,textAlign:"right"}}>{topN}</span>
          <span style={{fontSize:11,color:T.textMd,whiteSpace:"nowrap"}}>
            of {allChartData.length} rows
          </span>
        </div>
      )}

      {chartEl}

      {/* Legend — show even for single series so field name is clear */}
      {chartType!=="pie"&&series.length>0&&(
        <div style={{display:"flex",flexWrap:"wrap",gap:14,marginTop:10,justifyContent:"center",
          padding:"6px 0",borderTop:"0.5px solid "+T.border}}>
          {series.map((s,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",gap:6,fontSize:11,color:T.text}}>
              <span style={{width:14,height:10,background:palette[i%palette.length],
                borderRadius:2,display:"inline-block",flexShrink:0}}/>
              <span style={{fontWeight:600}}>{s}</span>
            </div>
          ))}
          {needsGrouping&&(
            <div style={{display:"flex",alignItems:"center",gap:6,fontSize:11,color:T.textMd}}>
              <span style={{width:14,height:10,background:"#AAAAAA",borderRadius:2,display:"inline-block",flexShrink:0}}/>
              <span>Others ({allChartData.length - topN} rows grouped)</span>
            </div>
          )}
        </div>
      )}
    </div>
  );
}


function PivotTable({result,onDrillDown,numFmt,colOrder,onColReorder,colFilter,colExcluded,onColFilter,pivotFilters,onPivotFilter,pivotSort,onPivotSort}) {
  // ── ALL hooks must come before any conditional return (Rules of Hooks) ──────
  const [dragOverCol,setDragOverCol]=useState(null);
  const [colWidths,startColResize]=useColResize(120);
  const [valSort,setValSort]=useState(null); // {field,dir:"asc"|"desc"} — sort rows by metric

  // Derive vals safely — null when result not ready
  const vals = result&&!result.error ? result.vals : [];
  const colVals = result&&!result.error ? result.colVals : [];
  const hasGroups = colVals.length>0;

  // Reorder value metrics when no column field (drag-to-reorder value columns)
  // Must run unconditionally — guard internally with safe fallbacks
  const orderedVals=useMemo(()=>{
    if (!vals.length||hasGroups||!colOrder) return vals;
    const reordered=colOrder.map(n=>vals.find(v=>v.field===n)).filter(Boolean);
    return reordered.length===vals.length?reordered:vals;
  },[vals,colOrder,hasGroups]);

  // Derive rawRowKeys safely (null guard) — must be BEFORE early returns
  const rawRowKeys = result&&!result.error ? result.rowKeys : [];
  const rFsSafe    = result&&!result.error ? result.rFs    : [];

  // Apply pivot row field filters + sort — MUST be before early returns (hook rule)
  const rowKeys=useMemo(()=>{
    let rk=rawRowKeys;
    if (pivotFilters&&rFsSafe.length) {
      rk=rk.filter(rk=>rFsSafe.every((_f,i)=>{
        const sel=Array.isArray(pivotFilters[i])&&pivotFilters[i].length>0?pivotFilters[i]:null;
        return sel===null||sel.includes(rk[i]);
      }));
    }
    if (pivotSort&&pivotSort.fieldIdx!=null) {
      const {fieldIdx,dir}=pivotSort;
      rk=[...rk].sort((a,b)=>{
        const av=a[fieldIdx]||"",bv=b[fieldIdx]||"";
        const r=av.localeCompare(bv,undefined,{numeric:true});
        return dir==="za"?-r:r;
      });
    }
    return rk;
  },[rawRowKeys,pivotFilters,pivotSort,rFsSafe]);

  // ── Early returns AFTER all hooks ──────────────────────────────────────────
  if (!result) return(
    <div style={{textAlign:"center",padding:"48px 24px",fontSize:13,color:T.textMd,background:T.bgStat,borderRadius:10,border:"1px dashed "+T.border}}>
      Assign at least one Row field (R) and one Value field (V) to generate the pivot.
    </div>
  );
  if (result.error) return(
    <div style={{padding:"14px",background:"rgba(163,45,45,0.08)",border:"1px solid rgba(163,45,45,0.3)",borderRadius:8,fontSize:13,color:T.danger}}>Pivot error: {result.error}</div>
  );

  const {cells,colTotals,grandTotals,rFs,cF}=result;
  const nV=vals.length;
  // Apply value-column sort: reorder rows by a metric descending/ascending
  const sortedRowKeys=valSort
    ? (()=>{
        const vi=vals.findIndex(v=>v.field===valSort.field);
        if (vi===-1) return rowKeys;
        return [...rowKeys].sort((a,b)=>{
          const av=((cells[a.join("\0")]||{})["__total__"]||[])[vi]||0;
          const bv=((cells[b.join("\0")]||{})["__total__"]||[])[vi]||0;
          return valSort.dir==="asc"?av-bv:bv-av;
        });
      })()
    : rowKeys;

  // Use colFilter (filtered col values from Report), respecting drag reorder
  const orderedColVals=(()=>{
    const base = colFilter&&colFilter.length>=0 ? colFilter : colVals;
    if (!colOrder||!colOrder.length) return base;
    const ordered = colOrder.filter(v=>base.includes(v));
    return ordered.length===base.length ? ordered : base;
  })();
  const totalCells=rowKeys.length*Math.max(orderedColVals.length,1)*nV;
  if (totalCells>50000) return(
    <div style={{padding:"14px",background:"rgba(200,146,42,0.08)",border:"1px solid rgba(200,146,42,0.35)",borderRadius:8,fontSize:13,color:T.warning}}>
      Too many combinations ({rowKeys.length.toLocaleString()} rows x {Math.max(orderedColVals.length,1)} cols). Add filters or choose fields with fewer unique values.
    </div>
  );
  // In no-group mode, orderedVals may differ from vals (user reordered).
  // vi must reference the ORIGINAL vals index (cells are stored in original order).
  const origIdx=v=>vals.findIndex(ov=>ov.field===v.field);
  const flatCols=hasGroups
    ?[...orderedColVals.flatMap(cv=>orderedVals.map((v,_i)=>({key:cv,   vi:origIdx(v),isTotal:false}))),
      ...orderedVals.map((v,_i)=>                          ({key:"__total__",vi:origIdx(v),isTotal:true }))]
    :orderedVals.map((v,_i)=>                              ({key:"__total__",vi:origIdx(v),isTotal:false}));
  const effectiveVals=hasGroups?vals:orderedVals;
  const getCell=(s,col)=>((cells[s]||{})[col.key]||effectiveVals.map(()=>0))[col.vi]||0;
  // Grand totals from VISIBLE rows only (plain computation — no hook needed here)
  const visibleGrandTotals=sortedRowKeys.map
    ? effectiveVals.map((_,vi)=>sortedRowKeys.reduce((sum,rk)=>{
        const rkStr=rk.join("\0"); // must match the \0 separator used in cells keys
        return sum+(((cells[rkStr]||{})["__total__"]||[])[vi]||0);
      },0))
    : grandTotals;
  const visibleColTotals=(()=>{
    const out={};
    colVals.forEach(cv=>{
      out[cv]=effectiveVals.map((_,vi)=>sortedRowKeys.reduce((sum,rk)=>{
        const rkStr=rk.join("\0"); // must match \0 separator
        return sum+(((cells[rkStr]||{})[cv]||[])[vi]||0);
      },0));
    });
    return out;
  })();
  const getGrand=col=>(col.key==="__total__"?visibleGrandTotals:(visibleColTotals[col.key]||effectiveVals.map(()=>0)))[col.vi]||0;
  const lBorder=i=>i===0||flatCols[i-1].key!==flatCols[i].key?"1px solid "+T.borderDk:"none";
  const thStyle={padding:"10px 14px",fontWeight:700,fontSize:12,color:T.textLt,whiteSpace:"nowrap",background:T.bgHeader,borderBottom:"1px solid "+T.borderHd};
  // Column group drag handlers (only active when onColReorder is provided)
  const colDragStart=(e,cv)=>{if(onColReorder)e.dataTransfer.setData("pivotCol",cv);};
  const colDragOver=(e,cv)=>{if(onColReorder){e.preventDefault();setDragOverCol(cv);}};
  const colDrop=(e,cv)=>{
    if(!onColReorder)return;
    const from=e.dataTransfer.getData("pivotCol");
    setDragOverCol(null);
    if(from&&from!==cv)onColReorder(from,cv);
  };
  return(
    <div style={{overflowX:"auto",overflowY:"auto",maxHeight:"70vh",borderRadius:10,border:"1px solid "+T.border,boxShadow:"0 2px 8px rgba(92,45,26,0.08)"}}>
      <div style={{fontSize:11,color:T.textMd,padding:"5px 14px",background:T.bgStat,borderBottom:"0.5px solid "+T.border}}>
        {onDrillDown?"Click any cell to drill down  ·  ":""}{onColReorder?"Drag column headers to reorder":""}
      </div>
      <table style={{borderCollapse:"collapse",minWidth:"100%"}}>
        <thead style={{position:"sticky",top:0,zIndex:5}}>
          {hasGroups&&(
            <tr>
              {rFs.map((rf,ri)=>(
                <th key={ri} style={{...thStyle,textAlign:"left",borderBottom:nV>1?"0.5px solid "+T.borderHd:"1px solid "+T.borderHd,
                  position:"relative",background:(pivotSort&&pivotSort.fieldIdx===ri)||((pivotFilters&&pivotFilters[ri]||[]).length>0)?"rgba(200,146,42,0.2)":T.bgHeader}}>
                  <div style={{display:"flex",alignItems:"center",gap:4}}>
                    <span>{rf}{ri===0&&cF?<span style={{opacity:0.6,fontWeight:400}}> / {cF}</span>:null}</span>
                    {onPivotFilter&&<DrillColFilter
                      field={rf}
                      data={result.rowKeys.map(rk=>({[rf]:rk[ri]}))}
                      active={Array.isArray(pivotFilters&&pivotFilters[ri])?pivotFilters[ri]:[]}
                      onChange={sel=>onPivotFilter(ri,sel)}
                      numFields={new Set()}
                      activeSort={pivotSort&&pivotSort.fieldIdx===ri?pivotSort.dir:undefined}
                      onSort={(_,dir)=>onPivotSort&&onPivotSort({fieldIdx:ri,dir})}/>}
                  </div>
                  <ResizeHandle onMouseDown={e=>startColResize("row_"+ri,e)}/>
                </th>
              ))}
              {[...orderedColVals.map(cv=>({cv,isT:false})),{cv:"Total",isT:true}].map((g,i)=>(
                <th key={i} colSpan={nV}
                  draggable={!!onColReorder&&!g.isT}
                  onDragStart={e=>colDragStart(e,g.cv)}
                  onDragOver={e=>colDragOver(e,g.cv)}
                  onDragLeave={()=>setDragOverCol(null)}
                  onDrop={e=>colDrop(e,g.cv)}
                  style={{...thStyle,textAlign:"center",borderLeft:"1px solid "+T.borderHd,
                    borderBottom:nV>1?"0.5px solid "+T.borderHd:"1px solid "+T.borderHd,
                    background:g.isT?"#3D1A0E":dragOverCol===g.cv?"rgba(200,146,42,0.3)":T.bgHeader,
                    cursor:onColReorder&&!g.isT?"grab":"default",
                    outline:dragOverCol===g.cv?"2px dashed "+T.accent:"none",
                    transition:"background 0.1s",
                    position:"relative",width:colWidths["grp_"+g.cv]||undefined,minWidth:60}}>
                  {!g.isT&&onColReorder&&<span style={{opacity:0.4,fontSize:9,marginRight:4}}>⋮</span>}
                  {g.cv}
                  {!g.isT&&<ResizeHandle onMouseDown={e=>startColResize("grp_"+g.cv,e)}/>}
                </th>
              ))}
            </tr>
          )}
          <tr>
            {!hasGroups?rFs.map((rf,ri)=>(
              <th key={ri} style={{...thStyle,textAlign:"left",position:"relative",
                background:(pivotSort&&pivotSort.fieldIdx===ri)||((pivotFilters&&pivotFilters[ri]||[]).length>0)?"rgba(200,146,42,0.2)":T.bgHeader}}>
                <div style={{display:"flex",alignItems:"center",gap:4}}>
                  <span>{rf}</span>
                  {onPivotFilter&&<DrillColFilter
                    field={rf}
                    data={result.rowKeys.map(rk=>({[rf]:rk[ri]}))}
                    active={pivotFilters&&pivotFilters[ri]||[]}
                    onChange={sel=>onPivotFilter(ri,sel)}
                    numFields={new Set()}
                    activeSort={pivotSort&&pivotSort.fieldIdx===ri?pivotSort.dir:undefined}
                    onSort={(_,dir)=>onPivotSort&&onPivotSort({fieldIdx:ri,dir})}/>}
                </div>
                <ResizeHandle onMouseDown={e=>startColResize("row_"+ri,e)}/>
              </th>
            )):<th colSpan={rFs.length} style={{...thStyle}}></th>}
            {flatCols.map((col,i)=>{
              const v=effectiveVals[col.vi];
              const isDraggable=!!onColReorder&&!hasGroups&&effectiveVals.length>1;
              return(
                <th key={i}
                  draggable={isDraggable}
                  onDragStart={e=>{if(isDraggable)e.dataTransfer.setData("pivotCol",v.field);}}
                  onDragOver={e=>{if(isDraggable){e.preventDefault();setDragOverCol(v.field);}}}
                  onDragLeave={()=>setDragOverCol(null)}
                  onDrop={e=>{
                    if(!isDraggable)return;
                    const from=e.dataTransfer.getData("pivotCol");
                    setDragOverCol(null);
                    if(from&&from!==v.field)onColReorder(from,v.field);
                  }}
                  style={{...thStyle,textAlign:"right",borderLeft:lBorder(i),
                    background:col.isTotal&&hasGroups?"#3D1A0E":dragOverCol===v.field?"rgba(200,146,42,0.3)":T.bgHeader,
                    cursor:isDraggable?"grab":"default",
                    outline:dragOverCol===v.field?"2px dashed "+T.accent:"none",
                    position:"relative",width:colWidths["val_"+v.field]||undefined,minWidth:70}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"flex-end",gap:4}}>
                    {isDraggable&&<span style={{opacity:0.4,fontSize:9}}>{"⋮"}</span>}
                    <div style={{textAlign:"right"}}>
                      <div style={{display:"flex",alignItems:"center",gap:4,justifyContent:"flex-end"}}>
                        {v.field}
                        <button onClick={e=>{e.stopPropagation();setValSort(vs=>vs&&vs.field===v.field?(vs.dir==="asc"?{field:v.field,dir:"desc"}:null):{field:v.field,dir:"desc"});}}
                          title={"Sort by "+v.field}
                          style={{background:"none",border:"none",cursor:"pointer",color:"rgba(245,239,230,0.7)",fontSize:11,padding:"0 2px",lineHeight:1,flexShrink:0}}>
                          {valSort&&valSort.field===v.field?(valSort.dir==="desc"?"↓":"↑"):"⇅"}
                        </button>
                      </div>
                      <div style={{fontSize:10,fontWeight:400,opacity:0.65,marginTop:2}}>{v.agg}</div>
                    </div>
                  </div>
                  <ResizeHandle onMouseDown={e=>startColResize("val_"+v.field,e)}/>
                </th>
              );
            })}
          </tr>
        </thead>
        <tbody>
          {sortedRowKeys.map((rk,ri)=>{
            const rkStr=rk.join("\0");
            return(
              <tr key={rkStr} style={{background:ri%2===0?T.bgCard:T.bgAlt}}>
                {rk.map((v,i)=>(
                  <td key={i} style={{padding:"9px 14px",fontSize:13,fontWeight:600,borderBottom:"0.5px solid "+T.border,paddingLeft:i>0?28:14,color:T.text,
                    width:colWidths["row_"+i]||undefined,minWidth:80}}>
                    {i>0&&<span style={{opacity:0.3,marginRight:6,fontWeight:400}}>L</span>}
                    {v||<span style={{color:T.textMd}}>(blank)</span>}
                  </td>
                ))}
                {flatCols.map((col,i)=>{
                  const v=getCell(rkStr,col);
                  return(
                    <td key={i}
                      onClick={()=>onDrillDown&&onDrillDown(rk,col.key,vals[col.vi].agg+" of "+vals[col.vi].field)}
                      onMouseEnter={e=>{if(onDrillDown)e.currentTarget.style.background="rgba(92,45,26,0.08)";}}
                      onMouseLeave={e=>{if(onDrillDown)e.currentTarget.style.background=col.isTotal&&hasGroups?T.bgAlt:"";}}
                      style={{padding:"9px 14px",textAlign:"right",fontSize:13,borderBottom:"0.5px solid "+T.border,borderLeft:lBorder(i),
                        fontWeight:col.isTotal&&hasGroups?700:400,color:col.isTotal&&hasGroups?T.primary:T.text,
                        background:col.isTotal&&hasGroups?T.bgAlt:undefined,cursor:onDrillDown?"pointer":undefined}}>
                      {fmtNum(v,effectiveVals[col.vi].agg,effectiveVals[col.vi].field,numFmt)}
                    </td>
                  );
                })}
              </tr>
            );
          })}
        </tbody>
        <tfoot>
          <tr style={{background:T.bgTableH}}>
            <td colSpan={rFs.length} style={{padding:"11px 14px",fontWeight:700,fontSize:13,color:T.primary,borderTop:"1px solid "+T.border}}>Grand Total</td>
            {flatCols.map((col,i)=>(
              <td key={i} style={{padding:"11px 14px",textAlign:"right",fontWeight:700,fontSize:13,borderLeft:lBorder(i),color:col.isTotal?T.primary:T.secondary,borderTop:"1px solid "+T.border}}>
                {fmtNum(getGrand(col),effectiveVals[col.vi].agg,effectiveVals[col.vi].field,numFmt)}
              </td>
            ))}
          </tr>
        </tfoot>
      </table>
    </div>
  );
}

// ── Format selector ────────────────────────────────────────────────────────────
function FormatSelector({value,onChange,allowedFmts}) {
  // If allowedFmts is set and non-empty, restrict to those keys; first = default
  const allowed = allowedFmts&&allowedFmts.length>0
    ? NUM_FORMATS.filter(f=>allowedFmts.includes(f.key))
    : NUM_FORMATS;
  // If only 1 format allowed, hide the selector entirely
  if (allowed.length<=1) return null;
  return(
    <div style={{display:"flex",alignItems:"center",gap:6,padding:"4px",background:T.bgStat,borderRadius:8,border:"1px solid "+T.border}}>
      <span style={{fontSize:11,color:T.textMd,paddingLeft:6,fontWeight:500,whiteSpace:"nowrap"}}>Show in:</span>
      {allowed.map(f=>(
        <button key={f.key} onClick={()=>onChange(f.key)} style={{
          padding:"4px 10px",borderRadius:6,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,
          background:value===f.key?T.primary:"transparent",
          color:value===f.key?T.textLt:T.textMd}}>
          {f.label}
        </button>
      ))}
    </div>
  );
}

// ── Report ─────────────────────────────────────────────────────────────────────
function Report({config,data,fields,numFields,showExport,cardFields,onDrillHiddenColsChange,onColExcludedChange,tabs,activeTabIdx,onTabChange,onTabsChange,onTabDelete,onFiltersChange,onSaveFilters,onSaveColFilters,externalFilters,externalPivotFilters,onExternalFiltersChange,onExternalPivotFiltersChange}) {
  // Defensive: ensure numFields is always a Set
  // Priority: 1) DB numFields  2) config.values (admin-declared)  3) auto-detect
  numFields = useMemo(()=>{
    // Always start from DB numFields as base
    const base = numFields instanceof Set ? new Set(numFields)
      : new Set(Array.isArray(numFields)?numFields:Object.values(numFields||{}));

    // ALWAYS merge config.values fields — these are admin-declared numeric fields.
    // Must NOT short-circuit before this: DB may be missing fields (e.g. Adv Paid)
    // even when other numeric fields ARE in DB, so we always union both sources.
    if(config&&config.tabs&&Array.isArray(config.tabs)){
      config.tabs.forEach(t=>(t.config?.values||[]).forEach(v=>v.field&&base.add(v.field)));
    }
    (config?.values||[]).forEach(v=>v.field&&base.add(v.field));
    (cardFields||[]).forEach(cf=>cf.field&&base.add(cf.field));

    // Auto-detect supplement — catches numeric cols not in V zone / cardFields
    if(data&&data.length>0){
      const sample=data.slice(0,50);
      Object.keys(sample[0]||{}).forEach(f=>{
        if(base.has(f)) return;
        const vals=sample.map(r=>r[f]).filter(v=>v!=null&&v!=="");
        if(!vals.length) return;
        const numCount=vals.filter(v=>
          typeof v==="number"||(typeof v==="string"&&!isNaN(parseFloat(v))&&isFinite(v)&&v.trim()!=="")
        ).length;
        if(numCount/vals.length>0.8) base.add(f);
      });
    }
    return base;
  },[numFields,data,config,cardFields]);
  // ── Per-tab filter helpers ──────────────────────────────────────────────────
  // Get saved slicer filters for a tab (from config.tabs[i] or top-level)
  // Strip null/empty-array values from a filters object (they mean "no filter")
  const sanitiseFilters=(f)=>{
    if(!f||typeof f!=="object") return {};
    const out={};
    Object.entries(f).forEach(([k,v])=>{
      if(v==null) return;
      if(Array.isArray(v)&&v.length===0) return;
      out[k]=v;
    });
    return out;
  };
  const getSavedFilters=React.useCallback((idx)=>{
    if (tabs&&tabs[idx]&&tabs[idx].config&&tabs[idx].config.defaultFilters)
      return sanitiseFilters(tabs[idx].config.defaultFilters);
    return idx===0?sanitiseFilters(config.defaultFilters||{}):{};
  },[tabs,config.defaultFilters]);

  // Get saved pivot filters for a tab — always returns {rowIdx:[array]}
  const getSavedPivotFilters=React.useCallback((idx)=>{
    const validate=(pf)=>{
      if (!pf||typeof pf!=="object"||Array.isArray(pf)) return {};
      // Each value must be an array — filter out bad entries
      const out={};
      Object.entries(pf).forEach(([k,v])=>{if(Array.isArray(v)&&v.length>0)out[k]=v;});
      return out;
    };
    if (tabs&&tabs[idx]&&tabs[idx].config&&tabs[idx].config.defaultPivotFilters)
      return validate(tabs[idx].config.defaultPivotFilters);
    if (idx===0) return validate(config.defaultPivotFilters);
    return {};
  },[tabs,config.defaultPivotFilters]);

  // Per-tab slicer filter state — keyed by tab index
  const [tabFilters,setTabFilters]=useState(()=>{
    const init={};
    const count=(tabs&&tabs.length>0)?tabs.length:1;
    for(let i=0;i<count;i++) init[i]=getSavedFilters(i);
    return init;
  });
  // If external filter control provided (UserView), use it — else use internal tabFilters
  const filters=externalFilters!==undefined?externalFilters:(tabFilters[activeTabIdx]||{});
  const setFilters=(updater)=>{
    const cur=filters;
    const next=typeof updater==="function"?updater(cur):updater;
    if (onExternalFiltersChange) {onExternalFiltersChange(next);return;}
    setTabFilters(prev=>({...prev,[activeTabIdx]:next}));
  };
  // Init newly visited tab from its saved defaults (don't overwrite existing session state)
  useEffect(()=>{
    setTabFilters(prev=>{
      if (prev[activeTabIdx]!==undefined) return prev;
      return {...prev,[activeTabIdx]:getSavedFilters(activeTabIdx)};
    });
  },[activeTabIdx]);

  // When a completely new report is loaded (config._reportId changes), reinit all tab filters
  const prevReportId=useRef(config._reportId);
  useEffect(()=>{
    if (config._reportId!==prevReportId.current) {
      prevReportId.current=config._reportId;
      const count=(tabs&&tabs.length>0)?tabs.length:1;
      const init={};
      for(let i=0;i<count;i++) init[i]=getSavedFilters(i);
      setTabFilters(init);
      const pinit={};
      for(let i=0;i<count;i++) pinit[i]=getSavedPivotFilters(i);
      setTabPivotFilters(pinit);
    }
  },[config._reportId]);
  const [drill,setDrill]=useState(null);
  const [numFmt,setNumFmt]=useState(()=>(config.allowedFmts&&config.allowedFmts.length>0)?config.allowedFmts[0]:"Cr");
  const [colOrder,setColOrder]=useState(null);
  const [excludedColVals,setExcludedColVals]=useState(()=>new Set(config.colExcluded||[])); // persist to config
  const [showColFilter,setShowColFilter]=useState(false);
  const colFilterRef=useRef(null);
  const [adHocFields,setAdHocFields]=useState([]); // extra filters user adds in view mode
  const [drillHiddenCols,setDrillHiddenCols]=useState(()=>{
    if(config._reportId){
      try{
        const stored=localStorage.getItem("rh_drill_cols_"+config._reportId);
        if(stored)return JSON.parse(stored);
      }catch(e){}
    }
    return config.drillHiddenCols||[];
  });
  // drillColFmts: from localStorage (user-specific) or config (admin-saved)
  const drillColFmts=useMemo(()=>{
    if(config._reportId){
      try{
        const stored=localStorage.getItem("rh_drill_fmts_"+config._reportId);
        if(stored)return JSON.parse(stored);
      }catch(e){}
    }
    return config.drillColFmts||{};
  },[config._reportId,config.drillColFmts]);
  // Per-tab pivot filter state — keyed by tab index, values are {rowIdx:[vals]}
  const [tabPivotFilters,setTabPivotFilters]=useState(()=>{
    const init={};
    const count=(tabs&&tabs.length>0)?tabs.length:1;
    for(let i=0;i<count;i++) init[i]=getSavedPivotFilters(i);
    return init;
  });
  // pivotFilters is always a plain object {rowIdx: [arrayOfValues]}
  const pivotFilters=externalPivotFilters!==undefined?externalPivotFilters:(tabPivotFilters[activeTabIdx]||{});
  const setPivotFilters=(updater)=>{
    const cur=pivotFilters;
    const next=typeof updater==="function"?updater(cur):updater;
    const safe=(next&&typeof next==="object"&&!Array.isArray(next))?next:{};
    if (onExternalPivotFiltersChange){onExternalPivotFiltersChange(safe);return;}
    setTabPivotFilters(prev=>Object.assign({},prev,{[activeTabIdx]:safe}));
  };
  const [pivotSort,setPivotSort]=useState(null); // {fieldIdx, dir}
  const [viewMode,setViewMode]=useState("table"); // "table" | "chart"
  const [chartType,setChartType]=useState("bar"); // bar | line | area | pie
  const [tabDragIdx,setTabDragIdx]=useState(null);
  const [showAdHocPicker,setShowAdHocPicker]=useState(false);
  const adHocRef=useRef(null);
  const [filtersSaved,setFiltersSaved]=useState(false);
  const [colFiltersSaved,setColFiltersSaved]=useState(false);
  const result=useMemo(()=>runPivot(data,config,filters),[config,data,filters]);
  // For chart: apply pivotFilters on top of base result so chart matches table view
  const chartResult=useMemo(()=>{
    if (!result||result.error||!Object.keys(pivotFilters).length) return result;
    const rFs=config.rows||[];
    const activeFilters={...filters};
    Object.entries(pivotFilters).forEach(([idx,sel])=>{
      if (sel&&sel.length&&rFs[parseInt(idx)]) activeFilters[rFs[parseInt(idx)]]=sel;
    });
    return runPivot(data,config,activeFilters);
  },[result,pivotFilters,data,config,filters]);
  useEffect(()=>{if(result&&!result.error&&result.colVals){setColOrder(null);setExcludedColVals(new Set(config.colExcluded||[]));}},[config.columns,config.rows]);
  const setF=(f,v)=>setFilters(p=>{
    const next={...p,[f]:v};
    onFiltersChange&&onFiltersChange(next,externalPivotFilters!==undefined?externalPivotFilters:tabPivotFilters);
    return next;
  });
  // Also fire when filters are cleared
  const clearFilters=()=>{
    // ONLY clears current tab — other tabs completely untouched
    const idx=activeTabIdx;
    setTabFilters(prev=>Object.assign({},prev,{[idx]:{}}));
    setTabPivotFilters(prev=>Object.assign({},prev,{[idx]:{}}));
    setAdHocFields([]);
  };
  const hasActive=Object.values(filters).some(v=>Array.isArray(v)&&v.length>0);
  const cardFieldNames=useMemo(()=>(cardFields||[]).map(x=>typeof x==="string"?x:x.field),[cardFields]);
  const slicerFields=(config.filters||[]).filter(f=>!cardFieldNames.includes(f));
  const primaryVal=(config.values||[])[0]||{field:"",agg:"sum"};
  // All dimension fields available for ad-hoc filtering
  const dimFields=useMemo(()=>fields.filter(f=>!numFields.has(f)),[fields,numFields]);
  // Ad-hoc fields not already in configured slicers or card fields
  const addableFields=dimFields.filter(f=>!slicerFields.includes(f)&&!cardFieldNames.includes(f)&&!adHocFields.includes(f));
  useEffect(()=>{
    if (!showAdHocPicker) return;
    const h=e=>{if(adHocRef.current&&!adHocRef.current.contains(e.target))setShowAdHocPicker(false);};
    const t=setTimeout(()=>document.addEventListener("click",h),10);
    return()=>{clearTimeout(t);document.removeEventListener("click",h);};
  },[showAdHocPicker]);
  // Filtered colVals for PivotTable (excludes hidden column values)
  // When excludedColVals changes, lift to parent so it gets saved with report
  const updateExcluded=(fn)=>{
    setExcludedColVals(prev=>{
      const next=fn instanceof Set?fn:fn(prev);
      onColExcludedChange&&onColExcludedChange([...next]);
      return next;
    });
  };

  const filteredColVals=useMemo(()=>{
    if (!result||result.error||!result.colVals) return [];
    return result.colVals.filter(cv=>!excludedColVals.has(cv));
  },[result,excludedColVals]);

  // Close col-filter dropdown when clicking outside
  useEffect(()=>{
    if (!showColFilter) return;
    const h=e=>{if(colFilterRef.current&&!colFilterRef.current.contains(e.target))setShowColFilter(false);};
    const t=setTimeout(()=>document.addEventListener('click',h),10);
    return()=>{clearTimeout(t);document.removeEventListener('click',h);};
  },[showColFilter]);

  function handleColReorder(from,to) {
    const hasColField=result&&result.colVals&&result.colVals.length>0;
    const base=hasColField
      ?(colOrder||[...filteredColVals])
      :(colOrder||(config.values||[]).map(v=>v.field));
    const fi=base.indexOf(from),ti=base.indexOf(to);
    if(fi===-1||ti===-1)return;
    const arr=[...base];arr.splice(fi,1);arr.splice(ti,0,from);
    setColOrder(arr);
  }

  return(
    <div>
      {/* Tab strip — multi-section reports. Shown for admin (onTabsChange present) or when tabs exist */}
      {(onTabsChange||(tabs&&tabs.length>0))&&(
        <div style={{display:"flex",gap:0,marginBottom:14,borderBottom:"2px solid "+T.border,overflowX:"auto",alignItems:"flex-end"}}>
          {(tabs||[]).map((t,i)=>(
            <div key={t.id||i}
              draggable={!!onTabsChange}
              onDragStart={()=>setTabDragIdx(i)}
              onDragOver={e=>e.preventDefault()}
              onDrop={()=>{
                if(tabDragIdx===null||tabDragIdx===i||!onTabsChange)return;
                const arr=[...(tabs||[])];
                const [moved]=arr.splice(tabDragIdx,1);
                arr.splice(i,0,moved);
                // Adjust active index to follow the moved tab
                let newActive=activeTabIdx;
                if(activeTabIdx===tabDragIdx) newActive=i;
                else if(tabDragIdx<activeTabIdx&&i>=activeTabIdx) newActive=activeTabIdx-1;
                else if(tabDragIdx>activeTabIdx&&i<=activeTabIdx) newActive=activeTabIdx+1;
                onTabsChange(arr);
                if(newActive!==activeTabIdx) onTabChange&&onTabChange(newActive);
                setTabDragIdx(null);
              }}
              onDragEnd={()=>setTabDragIdx(null)}
              style={{position:"relative",display:"flex",alignItems:"center",
                opacity:tabDragIdx===i?0.5:1,
                cursor:onTabsChange?"grab":"default"}}>
              <button onClick={()=>onTabChange&&onTabChange(i)}
                style={{padding:"8px 16px",border:"none",
                  background:i===activeTabIdx?T.bgCard:"none",
                  color:i===activeTabIdx?T.primary:T.textMd,
                  fontWeight:i===activeTabIdx?700:500,
                  fontSize:13,cursor:onTabsChange?"grab":"pointer",
                  borderTopLeftRadius:8,borderTopRightRadius:8,
                  borderBottom:i===activeTabIdx?"2px solid "+T.primary:"none",
                  marginBottom:i===activeTabIdx?-2:0,whiteSpace:"nowrap",
                  display:"flex",alignItems:"center",gap:6}}>
                {onTabsChange&&<span style={{opacity:0.4,fontSize:10}}>⋮⋮</span>}
                {t.name||"Untitled"}
              </button>
              {onTabsChange&&(tabs||[]).length>1&&i===activeTabIdx&&(
                <button onClick={(e)=>{
                    e.stopPropagation();
                    if(confirm("Delete tab "+(t.name||"Untitled")+"?")){
                      if(onTabDelete){
                        onTabDelete(i); // single atomic delete+switch — avoids setConfig race
                      } else {
                        const nt=tabs.filter((_,idx)=>idx!==i);
                        onTabsChange(nt);
                        if(activeTabIdx>=nt.length)onTabChange(Math.max(0,nt.length-1));
                      }
                    }
                  }}
                  title="Delete tab"
                  style={{position:"absolute",top:6,right:4,background:"none",border:"none",
                    cursor:"pointer",fontSize:11,color:T.textMd,padding:"0 4px",lineHeight:1}}>
                  ×
                </button>
              )}
            </div>
          ))}
          {onTabsChange&&(
            <button onClick={()=>{
                const existing=tabs||[];
                const name=prompt("Name for new tab?","Tab "+(existing.length+1));
                if(!name)return;
                // First time: snapshot current config as "Tab 1" before adding new
                let newTabs;
                if (existing.length===0) {
                  const tab1={id:"t"+Date.now()+"a",name:"Tab 1",config:{...config},cardFields:[...(cardFields||[])]};
                  const tab2={id:"t"+Date.now()+"b",name,config:{...config},cardFields:[...(cardFields||[])]};
                  newTabs=[tab1,tab2];
                  onTabsChange(newTabs);
                  onTabChange&&onTabChange(1); // switch to new tab
                } else {
                  const newTab={id:"t"+Date.now(),name,config:{...config},cardFields:[...(cardFields||[])]};
                  newTabs=[...existing,newTab];
                  onTabsChange(newTabs);
                  onTabChange&&onTabChange(existing.length);
                }
              }}
              title="Add new tab"
              style={{padding:"8px 12px",border:"1px dashed "+T.border,background:"none",
                color:T.textMd,fontSize:12,cursor:"pointer",borderRadius:6,marginLeft:6,marginBottom:2}}>
              + Tab
            </button>
          )}
          {onTabsChange&&(tabs||[]).length>0&&(
            <button onClick={()=>{
                const cur=tabs[activeTabIdx];
                const newName=prompt("Rename tab:",cur.name||"");
                if(newName===null||!newName.trim())return;
                const nt=tabs.map((t,i)=>i===activeTabIdx?{...t,name:newName.trim()}:t);
                onTabsChange(nt);
              }}
              title="Rename current tab"
              style={{padding:"5px 10px",border:"none",background:"none",color:T.textMd,fontSize:11,cursor:"pointer",marginLeft:4,marginBottom:6}}>
              ✏ Rename
            </button>
          )}
        </div>
      )}

      {/* Format selector + view toggle + export row */}
      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:16,flexWrap:"wrap"}}>
        {/* Table / Chart toggle */}
        <div style={{display:"flex",gap:3,background:T.bgStat,borderRadius:7,padding:3,border:"0.5px solid "+T.border}}>
          <button onClick={()=>setViewMode("table")}
            style={{padding:"5px 14px",border:"none",borderRadius:5,
              background:viewMode==="table"?T.primary:"none",color:viewMode==="table"?T.textLt:T.textMd,
              cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:4}}>
            <span>▦</span> Table
          </button>
          <button onClick={()=>setViewMode("chart")}
            style={{padding:"5px 14px",border:"none",borderRadius:5,
              background:viewMode==="chart"?T.primary:"none",color:viewMode==="chart"?T.textLt:T.textMd,
              cursor:"pointer",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:4}}>
            <span>📊</span> Chart
          </button>
        </div>
        <FormatSelector value={numFmt} onChange={setNumFmt} allowedFmts={config.allowedFmts}/>
        {hasActive&&<button onClick={clearFilters} style={{fontSize:12,color:T.textMd,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>Clear all filters</button>}
        {Object.keys(pivotFilters).length>0&&<button onClick={()=>setPivotFilters({})} style={{fontSize:12,color:T.danger,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>Clear col filters</button>}
        {showExport&&result&&!result.error&&(
          <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center"}}>
            <button onClick={()=>exportExcel(result,config,numFmt)}
              style={{padding:"6px 14px",background:T.bgHeader,color:T.textLt,border:"none",borderRadius:6,cursor:"pointer",fontSize:12,fontWeight:600}}>
              ↓ Export Excel
            </button>
            <button onClick={()=>exportPDF(config)}
              style={{padding:"6px 14px",background:"none",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:12,color:T.text}}>
              ↓ Export PDF
            </button>
          </div>
        )}
      </div>

      {/* KPI stat cards — clickable to filter table to rows where that field has a value */}
      {result&&!result.error&&(
        <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
          {result.vals.map((v,i)=>{
            const isOn=filters[v.field]&&filters[v.field].includes("__has__");
            return(
              <button key={i} onClick={()=>setF(v.field, isOn ? undefined : ["__has__"])}
                title={isOn?"Click to clear filter":"Click to show only rows where "+v.field+" has a value"}
                style={{background:isOn?T.primary:(i===0?T.bgHeader:T.bgCard),
                  borderRadius:8,padding:"12px 16px",flex:1,minWidth:120,
                  border:"2px solid "+(isOn?T.accent:i===0?T.primary:T.border),
                  boxShadow:isOn?"0 2px 8px rgba(92,45,26,0.25)":"0 1px 4px rgba(92,45,26,0.1)",
                  cursor:"pointer",textAlign:"left",transition:"all 0.15s",
                  transform:isOn?"translateY(-1px)":"none"}}>
                <div style={{fontSize:10,color:isOn?"rgba(245,239,230,0.6)":i===0?"rgba(245,239,230,0.7)":T.textMd,
                  marginBottom:4,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.5px"}}>
                  {v.agg} of {v.field}
                  {isOn&&<span style={{marginLeft:6,fontStyle:"italic",fontSize:9}}>· filtered</span>}
                </div>
                <div style={{fontSize:20,fontWeight:700,color:isOn?T.textLt:i===0?T.textLt:T.numColor}}>
                  {fmtNum(result.grandTotals[i],v.agg,v.field,numFmt)}
                </div>
                {isOn&&<div style={{fontSize:9,color:"rgba(245,239,230,0.6)",marginTop:3}}>click to clear</div>}
              </button>
            );
          })}
          <div style={{background:T.bgCard,borderRadius:8,padding:"12px 16px",flex:1,minWidth:120,border:"1px solid "+T.border}}>
            <div style={{fontSize:10,color:T.textMd,marginBottom:4,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.5px"}}>Records</div>
            <div style={{fontSize:20,fontWeight:700,color:T.numColor}}>{result.count.toLocaleString()}</div>
          </div>
        </div>
      )}

      {/* Card filter container — all card groups in one horizontal panel */}
      {(cardFields||[]).length>0&&(
        <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,padding:"12px 16px",marginBottom:14,
          overflowX:"auto"}}>
          <div style={{display:"flex",gap:24,minWidth:0,alignItems:"flex-start"}}>
            {(cardFields||[]).map(cf=>{
              const f=typeof cf==="string"?cf:cf.field;
              const cardAgg=typeof cf==="string"?"sum":cf.agg;
              // Cross-filter: show data filtered by all OTHER active card/slicer filters
              const otherFilters=Object.fromEntries(Object.entries(filters).filter(([k])=>k!==f));
              const otherKeys=[...new Set([...config.filters,...Object.keys(otherFilters).filter(k=>otherFilters[k]&&otherFilters[k].length)])];
              const cardData=otherKeys.length?data.filter(row=>otherKeys.every(ff=>{
                const s=otherFilters[ff];
                if(s==null||!Array.isArray(s)||s.length===0) return true;
                if(s.includes("__has__")||s.includes("__zero__")){
                  const v=row[ff];
                  const isZero=v===null||v===undefined||v===""||Number(v)===0;
                  if(s.includes("__has__")&&!isZero) return true;
                  if(s.includes("__zero__")&&isZero) return true;
                  return false;
                }
                return Array.isArray(s)&&s.includes(String(row[ff]||""));
              })):data;
              // Override primaryVal with card's own agg for numeric fields
              const cardPrimary=numFields&&numFields.has(f)?{field:f,agg:cardAgg}:primaryVal;
              return(
                <div key={f} style={{flexShrink:0,minWidth:140}}>
                  <QuickFilterCards field={f} data={cardData} activeFilters={filters[f]||[]}
                    onFilter={v=>setF(f,v)} numFmt={numFmt} numFields={numFields}
                    primaryVal={cardPrimary} cardAgg={cardAgg}/>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Slicers + column value filter — in one unified filter bar */}
      {(slicerFields.length>0||adHocFields.length>0||addableFields.length>0||(result&&result.cF))&&(
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14,flexWrap:"wrap",position:"relative",zIndex:200}}>
          <span style={{fontSize:12,color:T.textMd,fontWeight:600}}>Filters:</span>
          {slicerFields.map(f=><Slicer key={f} field={f} active={filters[f]} onChange={v=>setF(f,v)} data={data}/>)}
          {adHocFields.map(f=>(
            <div key={f} style={{position:"relative",display:"inline-flex",alignItems:"center",gap:2}}>
              <Slicer field={f} active={filters[f]} onChange={v=>setF(f,v)} data={data}/>
              <button onClick={()=>{setAdHocFields(af=>af.filter(x=>x!==f));setF(f,[]);}}
                title="Remove this filter"
                style={{width:16,height:16,borderRadius:"50%",border:"0.5px solid "+T.border,background:T.bgStat,cursor:"pointer",fontSize:10,color:T.textMd,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                ×
              </button>
            </div>
          ))}
          {addableFields.length>0&&(
            <div ref={adHocRef} style={{position:"relative"}}>
              <button onClick={()=>setShowAdHocPicker(p=>!p)}
                style={{display:"flex",alignItems:"center",gap:4,padding:"5px 10px",border:"1px dashed "+T.borderDk,borderRadius:6,background:"none",cursor:"pointer",fontSize:12,color:T.textMd}}>
                + Add filter
              </button>
              {showAdHocPicker&&(
                <div style={{position:"absolute",top:"calc(100% + 4px)",left:0,zIndex:300,background:T.bgCard,border:"1px solid "+T.border,borderRadius:8,minWidth:200,maxHeight:260,overflowY:"auto",boxShadow:"0 6px 20px rgba(92,45,26,0.18)"}}>
                  <div style={{padding:"7px 12px",borderBottom:"0.5px solid "+T.border,fontSize:11,fontWeight:700,color:T.textMd}}>Add a filter field</div>
                  {addableFields.map(f=>(
                    <button key={f} onClick={()=>{setAdHocFields(af=>[...af,f]);setShowAdHocPicker(false);}}
                      style={{display:"block",width:"100%",textAlign:"left",padding:"7px 12px",border:"none",background:"none",cursor:"pointer",fontSize:12,color:T.text}}>
                      {f} <span style={{fontSize:10,color:numFields.has(f)?T.tagV:T.textMd}}>{numFields.has(f)?"#":"Aa"}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          )}
          {/* Column value filter chip — looks like a Slicer but controls which col groups show */}
          {result&&result.cF&&result.colVals&&result.colVals.length>0&&(
            <div ref={colFilterRef} style={{position:"relative"}}>
              <button onClick={()=>setShowColFilter(v=>!v)}
                style={{display:"flex",alignItems:"center",gap:4,padding:"5px 10px",
                  border:"1px solid "+(excludedColVals.size>0?T.primary:T.borderDk),
                  borderRadius:6,background:excludedColVals.size>0?"rgba(92,45,26,0.08)":"none",
                  cursor:"pointer",fontSize:12,
                  color:excludedColVals.size>0?T.primary:T.text,fontWeight:excludedColVals.size>0?600:400}}>
                {result.cF} {excludedColVals.size>0?"("+excludedColVals.size+" hidden)":""}
                <span style={{fontSize:9,color:T.textMd,marginLeft:2}}>▾</span>
              </button>
              {showColFilter&&(
                <div style={{position:"absolute",top:"calc(100% + 5px)",left:0,zIndex:9999,
                  background:T.bgCard,border:"1px solid "+T.border,borderRadius:8,
                  minWidth:240,boxShadow:"0 6px 24px rgba(92,45,26,0.2)",overflow:"hidden"}}>
                  <div style={{padding:"8px 12px",borderBottom:"0.5px solid "+T.border,
                    display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                    <span style={{fontSize:11,fontWeight:700,color:T.primary}}>{result.cF} columns</span>
                    <div style={{display:"flex",gap:10}}>
                      <button onClick={()=>updateExcluded(new Set())}
                        style={{fontSize:10,color:T.textMd,background:"none",border:"none",cursor:"pointer"}}>All</button>
                      <button onClick={()=>updateExcluded(new Set(result.colVals))}
                        style={{fontSize:10,color:T.textMd,background:"none",border:"none",cursor:"pointer"}}>None</button>
                    </div>
                  </div>
                  <div style={{maxHeight:250,overflowY:"auto"}}>
                    {result.colVals.map(cv=>{
                      const hidden=excludedColVals.has(cv);
                      const label=(cv===""||cv===null||cv===undefined)?"(blank)":String(cv);
                      return(
                        <label key={String(cv)} style={{display:"flex",alignItems:"center",gap:8,
                          padding:"6px 12px",cursor:"pointer",
                          background:!hidden?"none":"rgba(92,45,26,0.04)"}}>
                          <input type="checkbox" checked={!hidden}
                            onChange={()=>updateExcluded(prev=>{
                              const n=new Set(prev);
                              n.has(cv)?n.delete(cv):n.add(cv);
                              return n;
                            })}
                            style={{accentColor:T.primary,width:13,height:13,cursor:"pointer"}}/>
                          <span style={{fontSize:12,flex:1,color:hidden?T.textMd:T.text}}>{label}</span>
                        </label>
                      );
                    })}
                  </div>
                  {excludedColVals.size>0&&(
                    <div style={{padding:"6px 12px",borderTop:"0.5px solid "+T.border,fontSize:10,color:T.textMd}}>
                      {excludedColVals.size} hidden · Total column shows all data
                    </div>
                  )}
                </div>
              )}
            </div>
          )}
          {hasActive&&<button onClick={clearFilters} style={{fontSize:11,color:T.textMd,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>Clear all</button>}
          {onSaveFilters&&(
            filtersSaved
              ? <span style={{fontSize:11,color:T.success,fontWeight:700,marginLeft:4}}>✓ Saved</span>
              : <button onClick={()=>{onSaveFilters(filters,pivotFilters);setFiltersSaved(true);setTimeout(()=>setFiltersSaved(false),3000);}}
                  style={{fontSize:11,color:T.primary,background:"none",border:"1px solid "+T.primary,
                    borderRadius:4,padding:"2px 8px",cursor:"pointer",fontWeight:700,marginLeft:4}}>
                  💾 Save filters
                </button>
          )}
        </div>
      )}


      {viewMode==="table"
        ? <PivotTable result={result} numFmt={numFmt}
            colOrder={colOrder&&result&&result.colVals?colOrder:undefined}
            onColReorder={result&&!result.error&&
              ((result.colVals&&result.colVals.length>1)||((!result.cF)&&result.vals&&result.vals.length>1))
              ?handleColReorder:undefined}
            colExcluded={excludedColVals}
            colFilter={result&&result.colVals?filteredColVals:undefined}
            onColFilter={result&&result.cF&&result.colVals&&result.colVals.length>0?()=>setShowColFilter(v=>!v):undefined}
            pivotFilters={Object.keys(pivotFilters).length?pivotFilters:null}
            onPivotFilter={(idx,sel)=>setPivotFilters(p=>({...p,[idx]:sel}))}
            pivotSort={pivotSort}
            onPivotSort={setPivotSort}
            onDrillDown={(rowKey,colVal,label)=>setDrill({rowKey,colVal,rFs:result.rFs,cF:result.cF,metricLabel:label})}/>
        : <ChartView result={chartResult} numFmt={numFmt} chartType={chartType} onChartTypeChange={setChartType}/>
      }

      {drill&&<DrillDown data={data} target={drill} fields={fields} numFields={numFields} numFmt={numFmt}
        activeFilters={filters}
        savedHiddenCols={drillHiddenCols}
        savedColFmts={drillColFmts}
        configValues={config.values||[]}
        onSaveHiddenCols={(cols,fmts)=>{setDrillHiddenCols(cols);onDrillHiddenColsChange&&onDrillHiddenColsChange(cols,fmts);}}
        onSaveColFilters={cf=>{
          try{localStorage.setItem("rh_drill_colf_"+(drill.rowKey||""),JSON.stringify(cf));}catch(e){}
        }}
        onClose={()=>setDrill(null)}/>}
    </div>
  );
}

// ── Draggable field tag ────────────────────────────────────────────────────────
function DragTag({fieldName, color, onRemove, extra, onReorder, zone}) {
  const [over, setOver]=useState(false);
  return(
    <span
      draggable
      onDragStart={e=>{e.dataTransfer.setData("text/plain",zone+":"+fieldName);e.dataTransfer.effectAllowed="move";}}
      onDragOver={e=>{e.preventDefault();setOver(true);}}
      onDragLeave={()=>setOver(false)}
      onDrop={e=>{e.preventDefault();setOver(false);const raw=e.dataTransfer.getData("text/plain");const parts=raw.split(":");if(parts[0]===zone&&parts[1]!==fieldName)onReorder(parts[1],fieldName);}}
      style={{display:"inline-flex",alignItems:"center",gap:4,borderRadius:20,padding:"4px 8px 4px 10px",fontSize:12,fontWeight:600,maxWidth:180,cursor:"grab",
        background:over?"rgba(0,0,0,0.08)":"rgba(0,0,0,0.06)",color,
        outline:over?"2px dashed "+color:"none",transition:"outline 0.1s"}}>
      <span style={{opacity:0.5,fontSize:10,marginRight:2}}>:</span>
      <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color}}>{fieldName}</span>
      {extra}
      <button onClick={e=>{e.stopPropagation();onRemove();}} style={{background:"none",border:"none",cursor:"pointer",color,fontSize:14,lineHeight:1,padding:"0 2px",marginLeft:1,flexShrink:0}}>x</button>
    </span>
  );
}

// ── Zone box with drag-and-drop reorder ────────────────────────────────────────
function ZoneBox({label, color, fields, onRemove, isValues, onAggChange, onReorder, zone, emptyMsg}) {
  return(
    <div style={{background:T.bgCard,border:"1px solid "+color+"50",borderRadius:10,padding:12}}>
      <div style={{fontSize:10,fontWeight:700,color,marginBottom:8,textTransform:"uppercase",letterSpacing:"1px",display:"flex",alignItems:"center",gap:6}}>
        {label}
        <span style={{fontSize:9,opacity:0.6,fontWeight:400}}>drag to reorder</span>
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:6,minHeight:30}}>
        {isValues ? fields.map(v=>(
          <DragTag key={v.field} fieldName={v.field} color={color} zone={zone}
            onRemove={()=>onRemove(v.field)} onReorder={onReorder}
            extra={<>
              <select value={v.agg} onChange={e=>onAggChange&&onAggChange(v.field,e.target.value)}
                style={{fontSize:10,border:"none",background:"transparent",color,cursor:"pointer",padding:"0 2px",marginLeft:3}}>
                {AGGS.map(a=><option key={a} value={a}>{a}</option>)}
              </select>

            </>}/>
        )) : fields.map(f=>(
          <DragTag key={f} fieldName={f} color={color} zone={zone}
            onRemove={()=>onRemove(f)} onReorder={onReorder}/>
        ))}
        {!fields.length&&<span style={{fontSize:12,color:T.textMd,fontStyle:"italic"}}>{emptyMsg}</span>}
      </div>
    </div>
  );
}

// ── Field row (with type toggle + R/C/V/F/K buttons) ──────────────────────────
function FieldRow({field, isNum, status, onToggle, onToggleType, onToggleCard}) {
  const btns=[
    {zone:"rows",   L:"R", color:T.tagR, on:status.rows},
    {zone:"columns",L:"C", color:T.tagC, on:status.cols},
    ...(isNum?[{zone:"values",L:"V",color:T.tagV,on:status.vals}]:[]),
    {zone:"filters",L:"F", color:T.tagF, on:status.filters},
    {zone:"cards",  L:"K", color:T.tagK, on:status.card},
  ];
  const anyOn=status.rows||status.cols||status.vals||status.filters||status.card;
  return(
    <div style={{display:"flex",alignItems:"center",gap:5,padding:"6px 8px",borderRadius:6,background:anyOn?T.bgAlt:"transparent",marginBottom:1}}>
      <button onClick={onToggleType} title="Toggle numeric / dimension" style={{
        width:28,padding:"2px 3px",borderRadius:4,fontSize:10,fontWeight:700,cursor:"pointer",border:"none",flexShrink:0,
        background:isNum?"rgba(139,90,43,0.15)":"rgba(83,74,183,0.12)",color:isNum?T.tagV:T.tagR}}>
        {isNum?"#":"Aa"}
      </button>
      <span style={{fontSize:12,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:anyOn?T.secondary:T.text}} title={field}>{field}</span>
      <div style={{display:"flex",gap:3,flexShrink:0}}>
        {btns.map(b=>(
          <button key={b.zone} onClick={()=>b.zone==="cards"?onToggleCard&&onToggleCard(field):onToggle(b.zone,field)}
            title={(b.on?"Remove from ":"Add to ")+b.zone}
            style={{width:22,height:22,borderRadius:4,fontSize:10,fontWeight:700,cursor:"pointer",border:"none",
              background:b.on?b.color:T.bgTableH, color:b.on?"white":T.textMd}}>
            {b.L}
          </button>
        ))}
      </div>
    </div>
  );
}

// ── App header ─────────────────────────────────────────────────────────────────
function AppHeader({role, onLogout, children}) {
  const isMobile = useViewport();
  return(
    <div style={{position:"sticky",top:0,zIndex:50,background:T.bgHeader,borderBottom:"2px solid "+T.borderHd,
      padding:isMobile?"0 12px":"0 20px",display:"flex",alignItems:"center",gap:isMobile?6:12,height:52,
      boxShadow:"0 2px 12px rgba(44,24,16,0.3)"}}>
      <span style={{fontWeight:700,fontSize:isMobile?14:15,color:T.textLt,letterSpacing:"-0.3px"}}>
        <span style={{color:T.accent}}>Report</span>Hub
      </span>
      {!isMobile&&<span style={{color:"rgba(245,239,230,0.3)"}}>|</span>}
      {!isMobile&&<span style={{fontSize:11,color:T.textLt,background:"rgba(255,255,255,0.12)",padding:"2px 10px",borderRadius:4,fontWeight:500}}>{role}</span>}
      <div style={{flex:1}}/>{children}
      <button onClick={onLogout} style={{padding:isMobile?"5px 10px":"5px 14px",background:"rgba(255,255,255,0.12)",border:"1px solid rgba(255,255,255,0.2)",borderRadius:6,cursor:"pointer",fontSize:12,color:T.textLt}}>{isMobile?"Out":"Logout"}</button>
    </div>
  );
}

// ── Upload Tab ─────────────────────────────────────────────────────────────────
// ── Custom App Credentials Panel ────────────────────────────────────────────────
// Lets admin store their own Azure / Google Cloud app credentials so ReportHub
// can access their SharePoint / Google Drive without blanket IT admin approval.
function CustomCredentialsPanel() {
  const [saved, setSaved] = useState(null); // {microsoft:{...}, google:{...}}
  const [editing, setEditing] = useState(null); // "microsoft" | "google" | null
  const [form, setForm] = useState({clientId:"", clientSecret:"", tenantId:""});
  const [showSecret, setShowSecret] = useState(false);
  const [testing, setTesting] = useState(false);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState(""); // {text, ok}

  useEffect(()=>{
    getCustomCredentials().then(setSaved).catch(()=>setSaved({}));
  },[]);

  const inp = {width:"100%",padding:"8px 10px",border:"1px solid "+T.border,borderRadius:6,
    fontSize:12,background:T.bgCard,color:T.text,boxSizing:"border-box",outline:"none"};

  const startEdit = (provider) => {
    setEditing(provider); setShowSecret(false); setMsg("");
    setForm({clientId:"", clientSecret:"", tenantId: saved?.[provider]?.tenantId||""});
  };

  const handleTest = async () => {
    if (!form.clientId || !form.clientSecret) { setMsg({text:"Enter Client ID and Secret first", ok:false}); return; }
    setTesting(true); setMsg("");
    try {
      const r = await testCustomCredentials(editing, form.clientId, form.clientSecret, form.tenantId);
      setMsg({text: r.message || (r.ok?"Verified":"Failed"), ok: r.ok, error: r.error});
    } catch(e) { setMsg({text: e.message, ok: false}); }
    finally { setTesting(false); }
  };

  const handleSave = async () => {
    if (!form.clientId || !form.clientSecret) { setMsg({text:"Client ID and Secret are required", ok:false}); return; }
    setSaving(true); setMsg("");
    try {
      await saveCustomCredentials(editing, form.clientId, form.clientSecret, form.tenantId);
      const updated = await getCustomCredentials();
      setSaved(updated);
      setEditing(null);
      setMsg({text:"Saved successfully", ok:true});
    } catch(e) { setMsg({text: e.message, ok:false}); }
    finally { setSaving(false); }
  };

  const handleDelete = async (provider) => {
    if (!confirm("Remove custom "+provider+" credentials? ReportHub will revert to shared app credentials.")) return;
    await deleteCustomCredentials(provider);
    const updated = await getCustomCredentials();
    setSaved(updated);
    setMsg({text:"Credentials removed", ok:true});
  };

  const providers = [
    {
      key: "microsoft",
      label: "Microsoft (SharePoint / OneDrive)",
      icon: "🪟",
      fields: [
        {key:"tenantId", label:"Tenant ID", placeholder:"yourcompany.onmicrosoft.com or Azure Tenant ID",
          hint:"Found in Azure Portal → Azure Active Directory → Overview → Tenant ID"},
        {key:"clientId", label:"Client ID (Application ID)", placeholder:"xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
          hint:"Azure Portal → App registrations → your app → Application (client) ID"},
        {key:"clientSecret", label:"Client Secret", placeholder:"Paste your client secret value", isSecret:true,
          hint:"Azure Portal → your app → Certificates & secrets → New client secret"},
      ],
      setupSteps: [
        "Azure Portal → App registrations → New registration",
        "Name: 'ReportHub', Accounts: Single tenant (your org only)",
        "API permissions → Microsoft Graph → Application → Sites.Read.All → Grant consent",
        "Certificates & secrets → New client secret → copy Value",
      ],
    },
    {
      key: "google",
      label: "Google (Drive / Sheets)",
      icon: "🔵",
      fields: [
        {key:"clientId", label:"Client ID", placeholder:"xxxxxxxxxxxx-xxx.apps.googleusercontent.com",
          hint:"Google Cloud Console → APIs & Services → Credentials → OAuth 2.0 client"},
        {key:"clientSecret", label:"Client Secret", placeholder:"GOCSPX-...", isSecret:true,
          hint:"Shown alongside the Client ID in Google Cloud Console"},
      ],
      setupSteps: [
        "Google Cloud Console → New project → Enable Google Drive API",
        "OAuth consent screen → External → Add your email as test user (or Publish)",
        "Credentials → Create OAuth client → Web → Add redirect URI from Settings",
        "Copy Client ID and Client Secret to this panel",
      ],
    },
  ];

  return (
    <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,overflow:"hidden",marginBottom:14}}>
      <div style={{padding:"10px 16px",background:T.bgTableH,borderBottom:"0.5px solid "+T.border,display:"flex",alignItems:"center",gap:8}}>
        <span style={{fontSize:14}}>🔧</span>
        <span style={{fontWeight:700,fontSize:13,color:T.primary}}>Custom app credentials</span>
        <span style={{fontSize:11,color:T.textMd}}>For orgs where blanket IT approval isn't available</span>
      </div>
      {msg&&<div style={{padding:"8px 16px",fontSize:12,
        color:msg.ok?T.success:"#A32D2D",
        background:msg.ok?"rgba(45,106,79,0.07)":"rgba(163,45,45,0.07)",
        borderBottom:"0.5px solid "+T.border}}>
        {msg.ok?"✅ ":"❌ "}{msg.text}{msg.error&&" — "+msg.error}
      </div>}

      {editing ? (
        /* ── Edit form ── */
        <div style={{padding:16}}>
          <div style={{fontWeight:700,fontSize:13,color:T.primary,marginBottom:12}}>
            {providers.find(p=>p.key===editing)?.icon} Configure {providers.find(p=>p.key===editing)?.label}
          </div>
          {/* Setup guide */}
          <div style={{background:T.bgStat,borderRadius:8,padding:"10px 14px",marginBottom:14,
            border:"0.5px solid "+T.border}}>
            <div style={{fontSize:11,fontWeight:700,color:T.textMd,marginBottom:6}}>Setup steps in {editing==="microsoft"?"Azure Portal":"Google Cloud Console"}:</div>
            {providers.find(p=>p.key===editing)?.setupSteps.map((s,i)=>(
              <div key={i} style={{fontSize:11,color:T.textMd,marginBottom:3,display:"flex",gap:6}}>
                <span style={{color:T.accent,fontWeight:700,flexShrink:0}}>{i+1}.</span>{s}
              </div>
            ))}
          </div>
          {/* Form fields */}
          <div style={{display:"flex",flexDirection:"column",gap:10,marginBottom:14}}>
            {providers.find(p=>p.key===editing)?.fields.map(f=>(
              <div key={f.key}>
                <div style={{fontSize:11,fontWeight:600,color:T.textMd,marginBottom:4}}>{f.label}</div>
                <div style={{position:"relative"}}>
                  <input type={f.isSecret&&!showSecret?"password":"text"}
                    value={form[f.key]} onChange={e=>setForm(prev=>({...prev,[f.key]:e.target.value}))}
                    placeholder={f.placeholder} style={inp}/>
                  {f.isSecret&&(
                    <button type="button" onClick={()=>setShowSecret(v=>!v)}
                      style={{position:"absolute",right:8,top:"50%",transform:"translateY(-50%)",
                        background:"none",border:"none",cursor:"pointer",fontSize:13,color:T.textMd}}>
                      {showSecret?"🙈":"👁"}
                    </button>
                  )}
                </div>
                <div style={{fontSize:10,color:T.textMd,marginTop:3,lineHeight:1.4}}>{f.hint}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            <button onClick={handleTest} disabled={testing||saving}
              style={{padding:"6px 16px",background:"none",border:"1px solid "+T.border,
                borderRadius:6,cursor:testing||saving?"not-allowed":"pointer",fontSize:12,color:T.text,
                opacity:testing||saving?0.6:1}}>
              {testing?"Testing…":"Test connection"}
            </button>
            <button onClick={handleSave} disabled={saving||testing}
              style={{padding:"6px 18px",background:T.primary,color:T.textLt,border:"none",
                borderRadius:6,cursor:saving||testing?"not-allowed":"pointer",fontSize:12,fontWeight:700,
                opacity:saving||testing?0.6:1}}>
              {saving?"Saving…":"Save credentials"}
            </button>
            <button onClick={()=>{setEditing(null);setMsg("");}}
              style={{padding:"6px 14px",background:"none",border:"none",cursor:"pointer",
                fontSize:12,color:T.textMd,textDecoration:"underline"}}>
              Cancel
            </button>
          </div>
        </div>
      ) : (
        /* ── Summary view ── */
        <div style={{display:"flex",flexWrap:"wrap"}}>
          {providers.map((p,i)=>{
            const s = saved?.[p.key];
            return(
              <div key={p.key} style={{flex:"1 1 260px",padding:"12px 16px",
                borderRight:i===0?"0.5px solid "+T.border:"none"}}>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:18}}>{p.icon}</span>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:600,fontSize:12,color:T.text}}>{p.label}</div>
                    {s?.clientIdMasked
                      ? <div style={{fontSize:10,color:T.success,marginTop:2}}>
                          ✓ Custom credentials saved · {s.clientIdMasked}
                          {s.tenantId&&<span style={{marginLeft:6,color:T.textMd}}>{s.tenantId}</span>}
                        </div>
                      : <div style={{fontSize:10,color:T.textMd,marginTop:2}}>Using shared ReportHub app</div>
                    }
                  </div>
                  <div style={{display:"flex",gap:6,flexShrink:0}}>
                    <button onClick={()=>startEdit(p.key)}
                      style={{padding:"4px 10px",background:T.primary,color:T.textLt,border:"none",
                        borderRadius:5,cursor:"pointer",fontSize:11,fontWeight:600}}>
                      {s?.clientIdMasked?"Update":"Configure"}
                    </button>
                    {s?.clientIdMasked&&(
                      <button onClick={()=>handleDelete(p.key)}
                        style={{padding:"4px 8px",background:"none",border:"1px solid rgba(163,45,45,0.4)",
                          borderRadius:5,cursor:"pointer",fontSize:11,color:"#A32D2D"}}>
                        Remove
                      </button>
                    )}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ── OAuth Connection Panel ──────────────────────────────────────────────────────
function OAuthPanel() {
  const [status,setStatus]=useState(null);
  const [loading,setLoading]=useState(false);
  const [msg,setMsg]=useState("");

  useEffect(()=>{
    getOAuthStatus().then(setStatus).catch(()=>{});
  },[]);

  async function connect(provider) {
    setLoading(true); setMsg("");
    try {
      const {url}=await(provider==="microsoft"?startMicrosoftAuth():startGoogleAuth());
      const popup=window.open(url,"oauth_"+provider,"width=600,height=700,left=200,top=80");
      if (!popup){setMsg("Pop-up blocked — allow pop-ups for this site, then try again.");setLoading(false);return;}
      const done=await new Promise(resolve=>{
        const h=e=>{if(e.data&&(e.data.type==="oauth-success"||e.data.type==="oauth-error")){window.removeEventListener("message",h);resolve(e.data);}};
        window.addEventListener("message",h);
        const t=setInterval(()=>{if(popup.closed){clearInterval(t);window.removeEventListener("message",h);resolve({type:"closed"});}},500);
      });
      const s=await getOAuthStatus();
      setStatus(s);
      if(done.type==="oauth-success"||s[provider]?.connected)setMsg("✅ "+(provider==="microsoft"?"Microsoft":"Google")+" account connected!");
      else if(done.type==="oauth-error")setMsg("❌ Failed: "+done.error);
      else setMsg("Window closed — try again if connection didn't complete.");
    }catch(e){setMsg("Error: "+e.message);}
    finally{setLoading(false);}
  }

  async function disconnect(provider){
    if(!confirm("Disconnect "+provider+" account?"))return;
    await disconnectOAuth(provider);
    const s=await getOAuthStatus();setStatus(s);setMsg(provider+" disconnected.");
  }

  if (!status) return null;

  const providers=[
    {key:"microsoft",label:"Microsoft OneDrive / SharePoint",icon:"🪟"},
    {key:"google",label:"Google Drive / Sheets",icon:"🔵"},
  ];

  return(
    <div style={{background:T.bgCard,borderRadius:10,border:"1px solid "+T.border,overflow:"hidden",marginBottom:14}}>
      <div style={{padding:"10px 16px",background:T.bgTableH,borderBottom:"0.5px solid "+T.border,display:"flex",alignItems:"center",gap:8}}>
        <span style={{fontSize:14}}>🔐</span>
        <span style={{fontWeight:700,fontSize:13,color:T.primary}}>Cloud storage accounts</span>
        <span style={{fontSize:11,color:T.textMd}}>Connect once · access files without sharing</span>
      </div>
      {msg&&<div style={{padding:"8px 16px",fontSize:12,
        color:msg.startsWith("✅")?T.success:msg.startsWith("❌")?"#A32D2D":T.textMd,
        background:msg.startsWith("✅")?"rgba(45,106,79,0.08)":msg.startsWith("❌")?"rgba(163,45,45,0.07)":T.bgStat,
        borderBottom:"0.5px solid "+T.border}}>{msg}</div>}
      <div style={{display:"flex",flexWrap:"wrap"}}>
        {providers.map((p,i)=>{
          const info=status[p.key]||{};
          return(
            <div key={p.key} style={{flex:"1 1 200px",padding:"12px 16px",borderRight:i===0?"0.5px solid "+T.border:"none"}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <span style={{fontSize:18}}>{p.icon}</span>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:600,fontSize:12,color:T.text}}>{p.label}</div>
                  <div style={{fontSize:10,marginTop:2,color:info.connected?T.success:T.textMd}}>
                    {info.connected?("✓ Connected"+(info.connectedAt?" · "+new Date(info.connectedAt).toLocaleDateString():""))
                      :info.configured?"Not connected":"⚠ Not configured in Railway"}
                  </div>
                </div>
                {info.connected
                  ?<button onClick={()=>disconnect(p.key)} disabled={loading}
                    style={{padding:"4px 10px",background:"none",border:"1px solid rgba(163,45,45,0.4)",borderRadius:5,cursor:"pointer",fontSize:11,color:"#A32D2D",flexShrink:0}}>
                    Disconnect
                  </button>
                  :info.configured
                    ?<button onClick={()=>connect(p.key)} disabled={loading}
                      style={{padding:"4px 12px",background:T.primary,color:T.textLt,border:"none",borderRadius:5,cursor:loading?"not-allowed":"pointer",fontSize:11,fontWeight:600,flexShrink:0,opacity:loading?0.6:1}}>
                      Connect
                    </button>
                    :<span style={{fontSize:10,color:"#A32D2D",flexShrink:0}}>Setup needed</span>
                }
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}


function UploadTab({libs, onDataLoaded, onDataRefresh, existingConfig, savedReports, savedLinks, onQuickRefresh, onDeleteLink, onUpdateLink}) {
  const [phase,setPhase]=useState("drop");
  const [dragOver,setDragOver]=useState(false);
  const [fileInfo,setFileInfo]=useState(null);
  const [sheetNames,setSheetNames]=useState([]);
  const [workbook,setWorkbook]=useState(null);
  const [rangeOverride,setRangeOverride]=useState(""); // manual cell range e.g. "A1:AM5000"
  const [refreshingLinkUrl,setRefreshingLinkUrl]=useState(null); // URL currently being refreshed
  const [editingLink,setEditingLink]=useState(null);   // {origUrl, url, sheet} — inline edit state
  const [confirmDeleteUrl,setConfirmDeleteUrl]=useState(null); // URL pending delete confirmation
  const [schema,setSchema]=useState([]);
  const [previewRows,setPreviewRows]=useState([]);
  const [allRows,setAllRows]=useState([]);
  const [allFields,setAllFields]=useState([]);
  const [parseError,setParseError]=useState("");
  const [parseStats,setParseStats]=useState(null);
  const [refreshUrl,setRefreshUrl]=useState("");
  const [refreshSheet,setRefreshSheet]=useState("");
  const [lastRefresh,setLastRefresh]=useState(null);
  const fileRef=useRef(null);
  const libsReady=!!(libs.XLSX&&libs.Papa);
  const [showRefreshPicker,setShowRefreshPicker]=useState(false);
  const [pendingRefreshData,setPendingRefreshData]=useState(null);
  const [selectedRefreshIds,setSelectedRefreshIds]=useState(new Set());
  const [pendingLinkSave,setPendingLinkSave]=useState(null); // {url,sheet} to save after Load

  function applySchema(rows,fields,name,blankRowsRemoved=0) {
    if (!rows.length){setParseError("No data rows found after cleaning.");setPhase("error");return;}
    const numFields=detectNumFields(rows,fields);
    const scm=fields.map(f=>({
      field:f,type:numFields.has(f)?"num":"dim",
      sample:_.uniq(rows.slice(0,5).map(r=>String(r[f]||"")).filter(Boolean)).slice(0,3),
      nullPct:Math.round(rows.filter(r=>r[f]===""||r[f]===null||r[f]===undefined).length/rows.length*100),
      uniqueCount:_.uniq(rows.map(r=>String(r[f]||""))).length,
    }));
    setAllRows(rows);setAllFields(fields);setPreviewRows(rows.slice(0,8));setSchema(scm);
    setParseStats({rows:rows.length,fields:fields.length,name,blankRowsRemoved});setPhase("preview");
  }

  function processRaw(rawRows,name) {
    try{
      const{rows,fields}=sanitizeRows(rawRows);
      const blankRowsRemoved = rawRows.length - rows.length;
      applySchema(rows,fields,name,blankRowsRemoved);
    }
    catch(e){setParseError("Cleaning error: "+e.message);setPhase("error");}
  }

  function loadSheet(wb,sheetName) {
    setPhase("parsing");
    setTimeout(()=>{
      try{
        const ws=wb.Sheets[sheetName];
        if (!ws){setParseError("Sheet not found: "+sheetName);setPhase("error");return;}
        if (rangeOverride.trim()) {
          try { libs.XLSX.utils.decode_range(rangeOverride.trim()); ws["!ref"]=rangeOverride.trim().toUpperCase(); }
          catch(e) { /* invalid range — ignore */ }
        } else if (ws["!ref"]){
          const r=libs.XLSX.utils.decode_range(ws["!ref"]);
          if (r.e.r>MAX_ROWS){r.e.r=MAX_ROWS;ws["!ref"]=libs.XLSX.utils.encode_range(r);}
        }
        const raw=libs.XLSX.utils.sheet_to_json(ws,{defval:null,raw:true,cellDates:true});
        processRaw(raw,sheetName);
      }catch(e){setParseError("Sheet error: "+e.message);setPhase("error");}
    },60);
  }

  async function handleFile(file) {
    if (!libsReady){setParseError("Libraries loading, please wait.");return;}
    const ext=file.name.split(".").pop().toLowerCase();
    if (!["csv","txt","xlsx","xls","xlsm","ods"].includes(ext)){setParseError("Unsupported file type: ."+ext);setPhase("error");return;}
    setParseError("");setFileInfo({name:file.name,size:file.size});setPhase("parsing");
    try{
      if (ext==="csv"||ext==="txt"){
        libs.Papa.parse(file,{header:true,skipEmptyLines:true,dynamicTyping:true,
          complete:res=>processRaw(res.data,file.name.replace(/\.[^.]+$/,"")),
          error:err=>{setParseError(err.message);setPhase("error");}});
      }else{
        const buf=await file.arrayBuffer();
        const wb=libs.XLSX.read(buf,{type:"array",cellDates:true});
        setWorkbook(wb);
        setWorkbook(wb);
        if (wb.SheetNames.length===1){setSheetNames(wb.SheetNames);setPhase("sheet");}
        else{setSheetNames(wb.SheetNames);setPhase("sheet");}
      }
    }catch(e){setParseError("Read error: "+e.message);setPhase("error");}
  }

  // Detect Microsoft / SharePoint URLs
  const isMsUrl = url => url.includes("sharepoint.com") || url.includes("onedrive.live.com") ||
    url.includes("1drv.ms") || url.includes("office.com") || url.includes("microsoftonline.com");

  // Try to parse an ArrayBuffer as an Excel file, return rows + sheetNames or null
  function tryParseXlsx(buf, sheet) {
    try {
      const wb = libs.XLSX.read(buf, { type: "array", cellDates: true });
      const wsName = sheet && wb.SheetNames.includes(sheet) ? sheet : wb.SheetNames[0];
      const ws = wb.Sheets[wsName];
      if (!ws) return null;
      if (ws["!ref"]) {
        const r = libs.XLSX.utils.decode_range(ws["!ref"]);
        if (r.e.r > 100000) { r.e.r = 100000; ws["!ref"] = libs.XLSX.utils.encode_range(r); }
      }
      return { rows: libs.XLSX.utils.sheet_to_json(ws, { defval: null, cellDates: true, raw: true }), sheetNames: wb.SheetNames };
    } catch(e) { return null; }
  }

  // Strategy A: Browser fetch with credentials (session cookies)
  async function fetchBrowser(url, sheet) {
    // cache:"no-store" forces the browser to always hit the network, never serve a
    // stale cached XLSX — critical for SharePoint where the file gets updated in-place.
    const resp = await fetch(url, { credentials: "include", redirect: "follow", cache: "no-store" });
    if (!resp.ok) throw new Error("HTTP " + resp.status);
    const ct = resp.headers.get("content-type") || "";
    if (ct.includes("text/html")) throw new Error("got-html");
    const buf = await resp.arrayBuffer();
    const result = tryParseXlsx(buf, sheet);
    if (!result) throw new Error("parse-failed");
    return result;
  }

  async function fetchFromUrl(url, sheet) {
    // Strategy 1: Browser fetch with session cookies — works for org accounts already signed in
    try {
      const result = await fetchBrowser(url, sheet);
      console.log("Browser fetch succeeded:", result.rows.length, "rows");
      return result;
    } catch(e) {
      console.log("Browser fetch:", e.message);
    }
    // Strategy 2: Backend proxy (OneDrive Sharing API + download=1) — for public links
    return await fetchUrlViaProxy(url, sheet||undefined);
  }

  async function handleUrl(urlOverride, sheetOverride) {
    const url = (urlOverride||refreshUrl).trim();
    const sheet = sheetOverride||refreshSheet;
    if (!url){setParseError("Enter a URL first.");setPhase("error");return;}
    setPhase("parsing");setParseError("");
    try{
      const result = await fetchFromUrl(url, sheet);
      setLastRefresh(new Date());
      const urlName = url.split("/").pop().split("?")[0]||"Imported";
      // If multiple sheets and no sheet was specified, show picker
      if (!sheet && result.sheetNames && result.sheetNames.length > 1) {
        setSheetNames(result.sheetNames);
        // Store the fetched rows keyed by sheet name for immediate use after pick
        // We re-fetch with the chosen sheet name via handleUrl(url, chosenSheet)
        setPhase("url-sheet"); // new phase: sheet picker for URL-loaded files
        setParseError(url);    // reuse parseError to store the URL
        return;
      }
      processRaw(result.rows, urlName);
    }catch(e){
      const msg = e.message||"Unknown error";
      // If it looks like an auth/login problem AND it is a Microsoft URL → offer popup login
      const isAuthErr = msg.includes("401")||msg.includes("403")||
        msg.includes("sign-in")||msg.includes("preview page")||msg.includes("login")||
        msg.includes("got-html")||msg.includes("Got a login")||msg.includes("needs_auth");
      if (isAuthErr) {
        setParseError("needs_auth_refresh");
        setPhase("error");
        return;
      }
      setParseError(msg + (msg.includes("404") ? " — check the link is correct." : ""));
      setPhase("error");
    }
  }

  // Popup login: open the OneDrive/SharePoint URL in a popup so the user can sign in,
  // then retry the browser fetch after the popup closes
  async function handlePopupLogin() {
    const url = parseError; // URL stored here when phase === "login-required"
    const sheet = refreshSheet;
    const popup = window.open(url, "ms_login", "width=900,height=650,left=200,top=100");
    if (!popup) { setParseError("Pop-up was blocked. Please allow pop-ups for this site and try again."); setPhase("error"); return; }
    setPhase("popup-waiting");
    // Poll until popup closes
    await new Promise(resolve => {
      const t = setInterval(() => { if (popup.closed) { clearInterval(t); resolve(); } }, 500);
    });
    // Popup closed — retry browser fetch (user should now have session cookies)
    setPhase("parsing");
    setParseError("");
    try {
      const result = await fetchBrowser(url, sheet);
      setLastRefresh(new Date());
      processRaw(result.rows, url.split("/").pop().split("?")[0]||"Imported");
    } catch(e) {
      setParseError("Still could not access the file after sign-in. " +
        "If the file requires organisational permissions that ReportHub does not have, " +
        "please download it and use the file upload button instead.");
      setPhase("error");
    }
  }

  const onDrop=useCallback(e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)handleFile(f);},[libs]);

  function toggleType(field){setSchema(s=>s.map(item=>item.field===field?{...item,type:item.type==="num"?"dim":"num"}:item));}

  function confirmLoad() {
    const numFields=new Set(schema.filter(s=>s.type==="num").map(s=>s.field));
    const fields=schema.map(s=>s.field); // preserve original chronological order
    const name=parseStats&&parseStats.name?parseStats.name:"Report";
    // Always auto-generate a fresh pivot config from the new data's fields.
    // We deliberately do NOT inherit the full existingConfig (tabs, rows, columns,
    // values, filters) because "Load fresh / reset builder" means exactly that —
    // a clean slate. Inheriting a stale config from a deleted or unrelated report
    // caused old tabs to appear in new reports.
    // Only carry over sourceLinks so the saved URL is remembered.
    let baseConfig=autoConfig(fields,numFields,name);
    if (existingConfig?.sourceLinks?.length) {
      baseConfig.sourceLinks=existingConfig.sourceLinks;
    }
    if (refreshUrl.trim()) {
      const newLink={url:refreshUrl.trim(),sheet:refreshSheet||"",label:name,lastRefreshed:Date.now()};
      const existing=baseConfig.sourceLinks||[];
      baseConfig={...baseConfig,sourceLinks:[...existing.filter(x=>x.url!==newLink.url),newLink]};
    }
    const rows=allRows.map(r=>{
      const out={...r};
      fields.forEach(f=>{if(numFields.has(f)){const v=r[f];if(typeof v!=="number"){const n=parseFloat(String(v||"").replace(/[$,₹]/g,""));out[f]=isNaN(n)?0:n;}}});
      return out;
    });
    onDataLoaded({rows,fields,numFields,config:baseConfig});
  }

  const fmtSize=b=>b>1048576?(b/1048576).toFixed(1)+" MB":(b/1024).toFixed(1)+" KB";
  const inp={width:"100%",padding:"8px 11px",border:"1px solid "+T.border,borderRadius:7,fontSize:13,background:T.bgCard,color:T.text,boxSizing:"border-box",outline:"none"};

  return(
    <div style={{padding:20,maxWidth:960,margin:"0 auto"}}>

      {(phase==="drop"||phase==="error")&&(<>
          <CustomCredentialsPanel/>
          <OAuthPanel/>
        <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)} onDrop={onDrop}
          onClick={()=>libsReady&&fileRef.current.click()}
          style={{border:"2px dashed "+(dragOver?T.primary:T.border),borderRadius:14,padding:"52px 24px",textAlign:"center",
            cursor:libsReady?"pointer":"not-allowed",background:dragOver?"rgba(92,45,26,0.04)":T.bgCard,transition:"border-color 0.15s"}}>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.xlsm,.csv,.txt,.ods" style={{display:"none"}}
            onChange={e=>{const f=e.target.files[0];if(f)handleFile(f);e.target.value="";}}/>
          <div style={{fontSize:36,marginBottom:12}}>📂</div>
          {libsReady?(<>
            <div style={{fontWeight:700,fontSize:16,marginBottom:6,color:T.text}}>Drop your file here, or click to browse</div>
            <div style={{fontSize:13,color:T.textMd}}>Supports .xlsx .xls .xlsm .csv .ods</div>
            <div style={{fontSize:12,color:T.textMd,opacity:0.7,marginTop:4}}>Blank rows removed · Dates converted · Range-inflated files handled (capped at 100k rows)</div>
          </>):<div style={{fontSize:13,color:T.textMd}}>Loading parsers...</div>}
        </div>

        {phase==="error"&&parseError&&(
          <div style={{marginTop:14,padding:"12px 16px",background:"rgba(163,45,45,0.07)",border:"1px solid rgba(163,45,45,0.25)",borderRadius:8,fontSize:13,color:T.danger,display:"flex",alignItems:"center",gap:10}}>
            {parseError==="needs_auth_refresh"
              ? <><span style={{flex:1}}>⚠️ Your Google/Microsoft session expired. Reconnect your account in the Connect panel above, then retry.</span>
                  <button onClick={()=>{setPhase("drop");setParseError("");}}
                    style={{fontSize:12,color:T.primary,background:T.primary+"22",border:"1px solid "+T.primary,borderRadius:5,padding:"3px 10px",cursor:"pointer",fontWeight:700,flexShrink:0}}>
                    Reconnect &amp; retry
                  </button></>
              : <><span style={{flex:1}}>{parseError}</span>
                  <button onClick={()=>{setPhase("drop");setParseError("");}} style={{fontSize:12,color:T.danger,background:"none",border:"none",cursor:"pointer",textDecoration:"underline",flexShrink:0}}>Try again</button></>}
          </div>
        )}

        {/* ── Saved links — one-click refresh ──────────────────────────── */}
        {savedLinks&&savedLinks.length>0&&(
          <div style={{marginTop:18,background:T.bgCard,borderRadius:10,border:"1px solid "+T.border,overflow:"hidden"}}>
            <div style={{padding:"10px 16px",background:T.bgTableH,borderBottom:"0.5px solid "+T.border,display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontSize:14}}>⚡</span>
              <span style={{fontWeight:700,fontSize:13,color:T.primary}}>Saved links — quick refresh</span>
              <span style={{fontSize:11,color:T.textMd}}>No need to open the report first · refresh runs in the background</span>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:0}}>
              {savedLinks.map((lk,idx)=>{
                const linkedReport = savedReports&&savedReports.find(r=>r.id===lk.reportId);
                const isEditing = editingLink&&editingLink.origUrl===lk.url&&editingLink.reportId===lk.reportId;
                const isPendingDelete = confirmDeleteUrl===lk.url+"|"+lk.reportId;
                return(
                <div key={idx} style={{borderBottom:idx<savedLinks.length-1?"0.5px solid "+T.border:"none",background:idx%2===0?T.bgCard:T.bgStat}}>
                  {/* ── Normal row ── */}
                  {!isEditing&&!isPendingDelete&&(
                  <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 16px"}}>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:2,flexWrap:"wrap"}}>
                        <span style={{fontWeight:600,fontSize:12,color:T.text}}>{lk.label}</span>
                        {linkedReport&&<span style={{fontSize:10,color:T.textMd,background:T.bgStat,
                          border:"1px solid "+T.border,borderRadius:4,padding:"1px 6px"}}>
                          {linkedReport.name}
                        </span>}
                      </div>
                      <div style={{fontSize:10,color:T.textMd,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:340}}>{lk.url}</div>
                      {lk.sheet&&<div style={{fontSize:10,color:T.textMd}}>Sheet: {lk.sheet}</div>}
                      <div style={{fontSize:10,color:lk.lastRefreshed?T.success:T.textMd,marginTop:1}}>
                        {lk.lastRefreshed
                          ? "✓ Last refreshed: "+new Date(lk.lastRefreshed).toLocaleString()
                          : "Not yet refreshed — click ↻ Refresh to pull latest data"}
                      </div>
                    </div>
                    <div style={{display:"flex",flexDirection:"column",gap:5,flexShrink:0}}>
                      <button onClick={()=>{
                          if(refreshingLinkUrl) return;
                          setRefreshingLinkUrl(lk.url);
                          Promise.resolve(onQuickRefresh&&onQuickRefresh(lk))
                            .finally(()=>setRefreshingLinkUrl(null));
                        }}
                        disabled={!!refreshingLinkUrl}
                        style={{padding:"5px 14px",minWidth:100,
                          background:refreshingLinkUrl===lk.url?"#7a4a3a":T.primary,
                          color:T.textLt,border:"none",borderRadius:6,
                          cursor:refreshingLinkUrl?"not-allowed":"pointer",
                          fontSize:12,fontWeight:600,whiteSpace:"nowrap",
                          opacity:refreshingLinkUrl&&refreshingLinkUrl!==lk.url?0.4:1,
                          display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>
                        {refreshingLinkUrl===lk.url
                          ? <><span style={{display:"inline-block",animation:"spin 0.8s linear infinite",fontSize:14}}>⟳</span> Saving…</>
                          : "↻ Refresh"}
                      </button>
                      <button onClick={()=>setEditingLink({origUrl:lk.url,reportId:lk.reportId,url:lk.url,sheet:lk.sheet||""})}
                        style={{padding:"4px 10px",background:"none",border:"1px solid "+T.borderDk,
                          borderRadius:6,cursor:"pointer",fontSize:11,color:T.primary,textAlign:"center"}}>
                        ✏ Edit URL
                      </button>
                      <button onClick={()=>setConfirmDeleteUrl(lk.url+"|"+lk.reportId)}
                        style={{padding:"4px 10px",background:"none",border:"1px solid rgba(163,45,45,0.4)",
                          borderRadius:6,cursor:"pointer",fontSize:11,color:"#A32D2D",textAlign:"center"}}>
                        🗑 Remove
                      </button>
                    </div>
                  </div>
                  )}
                  {/* ── Inline edit row ── */}
                  {isEditing&&(
                  <div style={{padding:"12px 16px",background:"rgba(92,45,26,0.04)",display:"flex",flexDirection:"column",gap:8}}>
                    <div style={{fontWeight:600,fontSize:12,color:T.primary}}>Edit link for {linkedReport?.name||lk.label}</div>
                    <input value={editingLink.url} onChange={e=>setEditingLink(el=>({...el,url:e.target.value}))}
                      placeholder="SharePoint / OneDrive URL"
                      style={{padding:"7px 10px",border:"1px solid "+T.border,borderRadius:6,fontSize:12,
                        background:T.bgCard,color:T.text,outline:"none",width:"100%",boxSizing:"border-box"}}/>
                    <input value={editingLink.sheet} onChange={e=>setEditingLink(el=>({...el,sheet:e.target.value}))}
                      placeholder="Sheet name (optional)"
                      style={{padding:"7px 10px",border:"1px solid "+T.border,borderRadius:6,fontSize:12,
                        background:T.bgCard,color:T.text,outline:"none",width:"100%",boxSizing:"border-box"}}/>
                    <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                      <button onClick={async()=>{
                          const newUrl=editingLink.url.trim();
                          const newSheet=editingLink.sheet.trim();
                          if(!newUrl) return;
                          const r2=savedReports.find(x=>x.id===editingLink.reportId);
                          if(!r2){setEditingLink(null);return;}
                          const cfg2=r2.config||{};
                          const existing=getSourceLinks(cfg2);
                          const newLabel=newUrl.split("/").filter(Boolean).pop().split("?")[0]||r2.name;
                          const updLinks=existing.some(x=>x.url===editingLink.origUrl)
                            ? existing.map(x=>x.url===editingLink.origUrl
                                ?{...x,url:newUrl,sheet:newSheet,label:newLabel,lastRefreshed:null}
                                :x)
                            : [...existing,{url:newUrl,sheet:newSheet,label:newLabel,lastRefreshed:null}];
                          await (onUpdateLink&&onUpdateLink({reportId:editingLink.reportId,updLinks,cfg:cfg2,newUrl,newSheet}));
                          setEditingLink(null);
                          setPhase("drop");
                        }}
                        style={{padding:"6px 18px",background:T.primary,color:T.textLt,border:"none",borderRadius:6,cursor:"pointer",fontSize:12,fontWeight:600}}>
                        Save
                      </button>
                      <button onClick={()=>setEditingLink(null)}
                        style={{padding:"6px 14px",background:"none",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:12,color:T.text}}>
                        Cancel
                      </button>
                    </div>
                    <div style={{fontSize:11,color:T.textMd,marginTop:6,lineHeight:1.5}}>
                      Saving will update the URL and automatically fetch fresh data from the new source.
                    </div>
                  </div>
                  )}
                  {/* ── Inline delete confirm ── */}
                  {isPendingDelete&&(
                  <div style={{padding:"12px 16px",background:"rgba(163,45,45,0.05)",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                    <span style={{fontSize:12,color:T.danger,flex:1}}>Remove this URL link from <strong>{linkedReport?.name||lk.label}</strong>? The report keeps its current data but won't auto-refresh.</span>
                    <div style={{display:"flex",gap:8,flexShrink:0}}>
                      <button onClick={()=>{onDeleteLink&&onDeleteLink(lk);setConfirmDeleteUrl(null);}}
                        style={{padding:"5px 16px",background:T.danger,color:"#fff",border:"none",borderRadius:6,cursor:"pointer",fontSize:12,fontWeight:700}}>
                        Yes, remove
                      </button>
                      <button onClick={()=>setConfirmDeleteUrl(null)}
                        style={{padding:"5px 12px",background:"none",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:12,color:T.text}}>
                        Cancel
                      </button>
                    </div>
                  </div>
                  )}
                </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ── New URL / add link ─────────────────────────────────────────── */}
        <div style={{marginTop:14,padding:"16px 18px",background:T.bgCard,borderRadius:10,border:"1px solid "+T.border}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
            <span style={{fontSize:14}}>🔗</span>
            <span style={{fontWeight:700,fontSize:13,color:T.text}}>{savedLinks&&savedLinks.length>0?"Add another URL":"Load from URL"}</span>
            {lastRefresh&&<span style={{fontSize:11,color:T.textMd,marginLeft:"auto"}}>Last: {lastRefresh.toLocaleTimeString()}</span>}
          </div>
          <div style={{fontSize:12,color:T.textMd,marginBottom:10,lineHeight:1.55}}>
            Paste a link from <strong>OneDrive</strong>, <strong>SharePoint</strong>, <strong>Google Drive</strong>, or <strong>Dropbox</strong>.
            The app first tries your browser session (works if you are signed into OneDrive in this browser),
            then falls back to a server-side download for publicly shared files.
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            <input value={refreshUrl} onChange={e=>setRefreshUrl(e.target.value)}
              placeholder="Paste share link here..."
              style={{...inp,flex:"2 1 300px"}}/>
            <input value={refreshSheet} onChange={e=>setRefreshSheet(e.target.value)}
              placeholder="Sheet name (optional)"
              style={{...inp,flex:"1 1 140px"}}/>
            <button onClick={()=>handleUrl()} disabled={!refreshUrl.trim()||!libsReady}
              style={{padding:"8px 16px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,
                cursor:refreshUrl.trim()&&libsReady?"pointer":"not-allowed",fontSize:13,fontWeight:600,
                opacity:refreshUrl.trim()&&libsReady?1:0.5,whiteSpace:"nowrap"}}>
              Load
            </button>
          </div>
          <div style={{fontSize:11,color:T.textMd,marginTop:8,lineHeight:1.5}}>
            <strong>Internal OneDrive/SharePoint (org account):</strong> Open the file in a browser tab while signed in, then use the file upload button above instead —
            the app will pick it up locally without sign-in issues.
          </div>
        </div>

        <div style={{marginTop:12,padding:"14px 18px",background:T.bgCard,borderRadius:10,border:"1px solid "+T.border,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
          <div>
            <div style={{fontWeight:700,fontSize:13,color:T.text,marginBottom:2}}>No file? Try built-in sample data</div>
            <div style={{fontSize:12,color:T.textMd}}>768 rows · Region x Category x Product x Month · Sales, Units, Profit</div>
          </div>
          <button onClick={()=>onDataLoaded(makeSample())}
            style={{padding:"8px 18px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,cursor:"pointer",fontSize:13,fontWeight:600,whiteSpace:"nowrap"}}>
            Load sample data
          </button>
        </div>
      </>)}

      {phase==="parsing"&&(
        <div style={{textAlign:"center",padding:"80px 24px"}}>
          <div style={{fontSize:36,marginBottom:16,animation:"spin 1s linear infinite",display:"inline-block"}}>⚙️</div>
          <div style={{fontWeight:700,fontSize:15,marginBottom:6,color:T.text}}>Parsing and cleaning file...</div>
          {fileInfo&&<div style={{fontSize:13,color:T.textMd}}>{fileInfo.name} · {fmtSize(fileInfo.size)}</div>}
          <div style={{fontSize:12,color:T.textMd,marginTop:8,lineHeight:1.7,opacity:0.8}}>
            Capping range at 100k rows · Removing blank rows · Converting dates<br/>
            Large files may take 5-15 seconds
          </div>
          <style>{"@keyframes spin{to{transform:rotate(360deg)}}"}</style>
        </div>
      )}

      {/* ── Sign-in required — popup login flow ──────────────────────────── */}
      {phase==="login-required"&&(
        <div style={{marginTop:14,background:T.bgCard,borderRadius:10,border:"1px solid "+T.accent,overflow:"hidden"}}>
          <div style={{padding:"14px 18px",background:"rgba(200,146,42,0.1)",borderBottom:"0.5px solid "+T.accent,display:"flex",alignItems:"center",gap:10}}>
            <span style={{fontSize:20}}>🔐</span>
            <div>
              <div style={{fontWeight:700,fontSize:13,color:T.text}}>Sign-in required to access this file</div>
              <div style={{fontSize:11,color:T.textMd,marginTop:2}}>The file is on OneDrive or SharePoint and requires your organisational account.</div>
            </div>
          </div>
          <div style={{padding:"16px 18px"}}>
            <div style={{fontSize:12,color:T.textMd,marginBottom:14,lineHeight:1.65}}>
              Click the button below. A <strong>Microsoft sign-in window</strong> will open in a popup.
              Sign into your account there, then close that window — ReportHub will automatically
              retry downloading the file using your session.
            </div>
            <div style={{background:T.bgStat,borderRadius:8,padding:"10px 14px",fontSize:11,color:T.textMd,marginBottom:14,border:"0.5px solid "+T.border}}>
              <strong>URL:</strong> <span style={{wordBreak:"break-all",fontSize:10}}>{parseError}</span>
            </div>
            <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
              <button onClick={handlePopupLogin}
                style={{padding:"9px 22px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,
                  cursor:"pointer",fontSize:13,fontWeight:700,display:"flex",alignItems:"center",gap:8}}>
                <span>🪟</span> Open sign-in window
              </button>
              <button onClick={()=>{setPhase("drop");setParseError("");}}
                style={{padding:"9px 16px",background:"none",border:"1px solid "+T.border,borderRadius:7,cursor:"pointer",fontSize:13,color:T.text}}>
                Cancel
              </button>
            </div>
            <div style={{fontSize:11,color:T.textMd,marginTop:12,lineHeight:1.5}}>
              <strong>Tip:</strong> After signing in once, future refreshes will work automatically
              as long as you remain signed in.
              If the popup is blocked by your browser, click the address bar icon to allow pop-ups for this site.
            </div>
          </div>
        </div>
      )}

      {phase==="popup-waiting"&&(
        <div style={{textAlign:"center",padding:"60px 24px"}}>
          <div style={{fontSize:40,marginBottom:14}}>🪟</div>
          <div style={{fontWeight:700,fontSize:15,marginBottom:8,color:T.text}}>Waiting for sign-in...</div>
          <div style={{fontSize:13,color:T.textMd,lineHeight:1.6}}>
            Please complete sign-in in the popup window.<br/>
            Once you close it, the file will be downloaded automatically.
          </div>
        </div>
      )}

      {phase==="sheet"&&(
        <div>
          <div style={{fontWeight:700,fontSize:16,color:T.text,marginBottom:4}}>Select a sheet</div>
          <div style={{fontSize:13,color:T.textMd,marginBottom:16}}>{fileInfo&&fileInfo.name} has {sheetNames.length} sheets.</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:10}}>
            {sheetNames.map((name,i)=>(
              <button key={name} onClick={()=>loadSheet(workbook,name)}
                style={{padding:"16px 18px",textAlign:"left",background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,cursor:"pointer",display:"flex",alignItems:"center",gap:12,color:T.text}}
                onMouseEnter={e=>e.currentTarget.style.borderColor=T.primary}
                onMouseLeave={e=>e.currentTarget.style.borderColor=T.border}>
                <span style={{width:36,height:36,background:T.bgStat,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>📄</span>
                <div><div style={{fontWeight:600,fontSize:13}}>{name}</div><div style={{fontSize:11,color:T.textMd}}>Sheet {i+1}</div></div>
              </button>
            ))}
          </div>
          {/* Manual range override */}
          <div style={{marginTop:14,padding:"12px 16px",background:T.bgStat,borderRadius:8,border:"1px solid "+T.border}}>
            <div style={{fontSize:12,fontWeight:700,color:T.textMd,marginBottom:6}}>
              📐 Manual range override <span style={{fontWeight:400}}>— if not all rows are loading</span>
            </div>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <input value={rangeOverride} onChange={e=>setRangeOverride(e.target.value.toUpperCase())}
                placeholder="e.g. A1:AM10000"
                style={{flex:1,padding:"6px 10px",border:"1px solid "+T.border,borderRadius:6,
                  fontSize:13,background:T.bgCard,color:T.text,outline:"none",fontFamily:"monospace"}}/>
              {rangeOverride&&<button onClick={()=>setRangeOverride("")}
                style={{padding:"6px 10px",background:"none",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:12,color:T.textMd}}>Clear</button>}
            </div>
            <div style={{fontSize:11,color:T.textMd,marginTop:4}}>
              In Excel: press Ctrl+End to jump to last used cell. Enter A1:[that cell] here, then click a sheet.
            </div>
          </div>
          <button onClick={()=>{setPhase("drop");setWorkbook(null);setRangeOverride("");}} style={{marginTop:14,fontSize:13,color:T.textMd,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>Different file</button>
        </div>
      )}

      {/* Sheet picker for URL-loaded files (no local workbook available) */}
      {phase==="url-sheet"&&(
        <div style={{background:T.bgCard,borderRadius:10,border:"1px solid "+T.border,overflow:"hidden"}}>
          <div style={{padding:"12px 16px",background:T.bgTableH,borderBottom:"0.5px solid "+T.border}}>
            <div style={{fontWeight:700,fontSize:15,color:T.primary,marginBottom:2}}>Select a sheet</div>
            <div style={{fontSize:12,color:T.textMd}}>
              This workbook has {sheetNames.length} sheets. Which one should be loaded?
            </div>
          </div>
          <div style={{padding:"12px 16px",display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:10}}>
            {sheetNames.map((name,i)=>(
              <button key={name}
                onClick={async()=>{
                  const url=parseError;
                  setRefreshSheet(name);
                  setPhase("parsing"); setParseError("");
                  try{
                    // Pass rangeOverride to backend so it reads correct row range
                    const result=await fetchUrlViaProxy(url, name, rangeOverride.trim()||undefined);
                    setLastRefresh(new Date());
                    processRaw(result.rows, url.split("/").pop().split("?")[0]||"Imported");
                  }catch(e){setParseError(e.message);setPhase("error");}
                }}
                style={{padding:"14px 16px",textAlign:"left",background:T.bgCard,border:"1px solid "+T.border,
                  borderRadius:10,cursor:"pointer",display:"flex",alignItems:"center",gap:12,color:T.text}}
                onMouseEnter={e=>e.currentTarget.style.borderColor=T.primary}
                onMouseLeave={e=>e.currentTarget.style.borderColor=T.border}>
                <span style={{width:34,height:34,background:T.bgStat,borderRadius:8,display:"flex",
                  alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>📄</span>
                <div>
                  <div style={{fontWeight:600,fontSize:13}}>{name}</div>
                  <div style={{fontSize:11,color:T.textMd}}>Sheet {i+1}</div>
                </div>
              </button>
            ))}
          </div>
          {/* Range override for URL-loaded files */}
          <div style={{padding:"12px 16px",borderTop:"0.5px solid "+T.border,background:T.bgStat}}>
            <div style={{fontSize:12,fontWeight:700,color:T.textMd,marginBottom:6}}>
              📐 Manual range override <span style={{fontWeight:400}}>— if not all rows load</span>
            </div>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <input value={rangeOverride} onChange={e=>setRangeOverride(e.target.value.toUpperCase())}
                placeholder="e.g. A1:AM10000  (leave blank for auto)"
                style={{flex:1,padding:"6px 10px",border:"1px solid "+T.border,borderRadius:6,
                  fontSize:13,background:T.bgCard,color:T.text,outline:"none",fontFamily:"monospace"}}/>
              {rangeOverride&&<button onClick={()=>setRangeOverride("")}
                style={{padding:"5px 10px",background:"none",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:12,color:T.textMd}}>Clear</button>}
            </div>
            <div style={{fontSize:11,color:T.textMd,marginTop:4}}>
              Open file in Excel → Ctrl+End → note the last cell (e.g. AM6321) → enter A1:AM6321 here, then click a sheet above.
            </div>
          </div>
          <div style={{padding:"10px 16px",borderTop:"0.5px solid "+T.border}}>
            <button onClick={()=>{setPhase("drop");setParseError("");setSheetNames([]);setRangeOverride("");}}
              style={{fontSize:13,color:T.textMd,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>
              Different URL
            </button>
          </div>
        </div>
      )}

      {phase==="preview"&&(
        <div>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12,marginBottom:20}}>
            <div>
              <div style={{fontWeight:700,fontSize:16,color:T.text,marginBottom:4}}>Parsed successfully</div>
              <div style={{fontSize:13,color:T.textMd}}>
                <strong>{parseStats&&parseStats.rows&&parseStats.rows.toLocaleString()}</strong> rows · <strong>{parseStats&&parseStats.fields}</strong> columns
                · Column order preserved from source file
              </div>
              {parseStats&&parseStats.blankRowsRemoved>0&&(
                <div style={{fontSize:11,color:T.textMd,marginTop:3,
                  background:"rgba(200,146,42,0.1)",borderRadius:4,padding:"3px 8px",display:"inline-block"}}>
                  ⚠ {parseStats.blankRowsRemoved} blank section rows removed (totals/spacers) — data rows are correct
                </div>
              )}
              {/* Range override — re-parse with custom range if row count looks wrong */}
              <div style={{display:"flex",alignItems:"center",gap:8,marginTop:8,flexWrap:"wrap"}}>
                <span style={{fontSize:11,color:T.textMd,fontWeight:600}}>Row count wrong?</span>
                <input value={rangeOverride} onChange={e=>setRangeOverride(e.target.value.toUpperCase())}
                  placeholder="Range override e.g. A1:AM6321"
                  style={{padding:"5px 9px",border:"1px solid "+T.border,borderRadius:5,fontSize:12,
                    background:T.bgCard,color:T.text,outline:"none",fontFamily:"monospace",width:200}}/>
                <button onClick={()=>{
                    if(!workbook||!refreshSheet) return;
                    loadSheet(workbook, refreshSheet||sheetNames[0]||"");
                  }}
                  disabled={!rangeOverride.trim()||!workbook}
                  style={{padding:"5px 12px",background:rangeOverride.trim()&&workbook?T.primary:"none",
                    color:rangeOverride.trim()&&workbook?T.textLt:T.textMd,
                    border:"1px solid "+(rangeOverride.trim()&&workbook?T.primary:T.border),
                    borderRadius:5,cursor:rangeOverride.trim()&&workbook?"pointer":"not-allowed",
                    fontSize:12,fontWeight:600,opacity:rangeOverride.trim()&&workbook?1:0.5}}>
                  Re-parse
                </button>
                {rangeOverride&&<button onClick={()=>setRangeOverride("")}
                  style={{padding:"5px 10px",background:"none",border:"none",cursor:"pointer",fontSize:11,color:T.textMd,textDecoration:"underline"}}>Clear</button>}
              </div>
            </div>
            <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
              <button onClick={()=>{setPhase("drop");setAllRows([]);setSchema([]);}}
                style={{padding:"8px 16px",background:"none",border:"1px solid "+T.border,borderRadius:7,cursor:"pointer",fontSize:13,color:T.text}}>Different file</button>
              {onDataRefresh&&savedReports&&savedReports.length>0&&(
                <button onClick={()=>{
                  // Build the data payload, then show report picker
                  const numFields=new Set(schema.filter(s=>s.type==="num").map(s=>s.field));
                  const fields=schema.map(s=>s.field);
                  const rows=allRows.map(r=>{
                    const out={...r};
                    fields.forEach(f=>{if(numFields.has(f)){const v=r[f];if(typeof v!=="number"){const n=parseFloat(String(v||"").replace(/[$,₹]/g,""));out[f]=isNaN(n)?0:n;}}});
                    return out;
                  });
                  setPendingRefreshData({rows,fields,numFields});
                  setShowRefreshPicker(true);
                }}
                style={{padding:"8px 20px",background:"none",border:"2px solid "+T.primary,borderRadius:7,cursor:"pointer",fontSize:13,fontWeight:600,color:T.primary}}>
                  ↻ Update existing report
                </button>
              )}
              {/* Report picker modal — multi-select checkboxes */}
              {showRefreshPicker&&pendingRefreshData&&(
                <div style={{position:"fixed",inset:0,zIndex:700,background:"rgba(44,24,16,0.55)",display:"flex",alignItems:"center",justifyContent:"center"}}>
                  <div style={{background:T.bgCard,borderRadius:12,width:"min(560px,94vw)",maxHeight:"82vh",display:"flex",flexDirection:"column",boxShadow:"0 12px 40px rgba(44,24,16,0.3)"}}>
                    {/* Header */}
                    <div style={{padding:"16px 20px",background:T.bgHeader,borderRadius:"12px 12px 0 0",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                      <div>
                        <div style={{fontWeight:700,fontSize:15,color:T.textLt}}>Select reports to update</div>
                        <div style={{fontSize:11,color:"rgba(245,239,230,0.65)",marginTop:2}}>
                          New data: {pendingRefreshData.rows.length.toLocaleString()} rows · {pendingRefreshData.fields.length} columns
                        </div>
                      </div>
                      <button onClick={()=>{setShowRefreshPicker(false);setPendingRefreshData(null);setSelectedRefreshIds(new Set());}}
                        style={{border:"none",background:"rgba(255,255,255,0.15)",color:T.textLt,borderRadius:6,width:28,height:28,cursor:"pointer",fontSize:16}}>×</button>
                    </div>
                    {/* Info bar + select all */}
                    <div style={{padding:"8px 16px",borderBottom:"0.5px solid "+T.border,fontSize:12,color:T.textMd,background:T.bgStat,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                      <span>Tick the reports whose data rows should be replaced. Builder layout stays unchanged.</span>
                      <button onClick={()=>setSelectedRefreshIds(prev=>prev.size===savedReports.length?new Set():new Set(savedReports.map(r=>r.id)))}
                        style={{fontSize:11,color:T.primary,background:"none",border:"none",cursor:"pointer",fontWeight:600,flexShrink:0,marginLeft:12}}>
                        {selectedRefreshIds.size===savedReports.length?"Deselect all":"Select all"}
                      </button>
                    </div>
                    {/* Report list with checkboxes */}
                    <div style={{overflowY:"auto",padding:"10px 14px",display:"flex",flexDirection:"column",gap:6}}>
                      {savedReports.map(r=>{
                        const checked=selectedRefreshIds.has(r.id);
                        return(
                          <label key={r.id} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 14px",
                            background:checked?"rgba(92,45,26,0.06)":T.bgCard,
                            border:"1px solid "+(checked?T.primary:T.border),borderRadius:8,cursor:"pointer"}}>
                            <input type="checkbox" checked={checked}
                              onChange={()=>setSelectedRefreshIds(prev=>{
                                const n=new Set(prev);
                                n.has(r.id)?n.delete(r.id):n.add(r.id);
                                return n;
                              })}
                              style={{width:16,height:16,accentColor:T.primary,flexShrink:0,cursor:"pointer"}}/>
                            <div style={{width:34,height:34,background:r.isPublished?T.primary:T.bgStat,borderRadius:8,
                              display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,flexShrink:0}}>
                              {r.isPublished?"📤":"📊"}
                            </div>
                            <div style={{flex:1,minWidth:0}}>
                              <div style={{fontWeight:600,fontSize:13,color:T.text,display:"flex",alignItems:"center",gap:8}}>
                                {r.name}
                                {r.isPublished&&<span style={{background:T.primary,color:T.textLt,borderRadius:8,padding:"1px 7px",fontSize:10,fontWeight:600}}>Published</span>}
                              </div>
                              <div style={{fontSize:11,color:T.textMd,marginTop:2}}>
                                {r.rows.toLocaleString()} rows · Rows: {(r.config&&r.config.rows||[]).join(", ")||"—"} · Values: {(r.config&&r.config?.values||[]).map(v=>v.field).join(", ")||"—"}
                              </div>
                            </div>
                          </label>
                        );
                      })}
                    </div>
                    {/* Footer with action button */}
                    <div style={{padding:"12px 16px",borderTop:"0.5px solid "+T.border,display:"flex",alignItems:"center",justifyContent:"space-between",gap:10}}>
                      <span style={{fontSize:12,color:T.textMd}}>
                        {selectedRefreshIds.size===0?"No reports selected":selectedRefreshIds.size+" report"+(selectedRefreshIds.size>1?"s":"")+" selected"}
                      </span>
                      <div style={{display:"flex",gap:8}}>
                        <button onClick={()=>{setShowRefreshPicker(false);setPendingRefreshData(null);setSelectedRefreshIds(new Set());}}
                          style={{padding:"7px 16px",background:"none",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:13,color:T.text}}>
                          Cancel
                        </button>
                        <button disabled={selectedRefreshIds.size===0}
                          onClick={async()=>{
                            const ids=[...selectedRefreshIds];
                            const data=pendingRefreshData; // capture before clearing
                            setShowRefreshPicker(false);
                            setSelectedRefreshIds(new Set());
                            setPendingRefreshData(null);
                            // onDataRefresh (AdminView scope) handles loading/toast/tab switch
                            for(const id of ids){
                              await onDataRefresh(data,id);
                            }
                          }}
                          style={{padding:"7px 18px",background:selectedRefreshIds.size>0?T.primary:"rgba(92,45,26,0.3)",
                            color:T.textLt,border:"none",borderRadius:6,cursor:selectedRefreshIds.size>0?"pointer":"not-allowed",
                            fontSize:13,fontWeight:700,opacity:selectedRefreshIds.size>0?1:0.6}}>
                          ↻ Update {selectedRefreshIds.size>1?selectedRefreshIds.size+" reports":"report"}
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              )}
              <button onClick={confirmLoad}
                style={{padding:"8px 20px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,cursor:"pointer",fontSize:13,fontWeight:600}}>
                {existingConfig?"Load fresh (reset builder)":"Load into builder"}
              </button>
            </div>
          </div>

          <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,overflow:"hidden",marginBottom:16}}>
            <div style={{padding:"10px 16px",background:T.bgTableH,borderBottom:"1px solid "+T.border,display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontWeight:700,fontSize:13,color:T.primary}}>Column schema</span>
              <span style={{fontSize:12,color:T.textMd}}>click type badge to toggle · fields appear in Excel column order</span>
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr style={{background:T.bgTableH}}>
                  {["#","Column","Type","Null %","Unique values","Slicer OK?","Sample values"].map(h=>(
                    <th key={h} style={{padding:"8px 13px",textAlign:"left",fontSize:11,fontWeight:700,color:T.textMd,borderBottom:"0.5px solid "+T.border,whiteSpace:"nowrap"}}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>
                  {schema.map((item,i)=>{
                    const slicerOk=item.uniqueCount<=SLICER_MAX;
                    return(
                      <tr key={item.field} style={{background:i%2===0?T.bgCard:T.bgAlt}}>
                        <td style={{padding:"9px 13px",fontSize:11,color:T.textMd,borderBottom:"0.5px solid "+T.border,fontWeight:600}}>{i+1}</td>
                        <td style={{padding:"9px 13px",fontWeight:700,fontSize:13,borderBottom:"0.5px solid "+T.border,color:T.text}}>{item.field}</td>
                        <td style={{padding:"9px 13px",borderBottom:"0.5px solid "+T.border}}>
                          <button onClick={()=>toggleType(item.field)}
                            style={{padding:"2px 9px",borderRadius:4,fontSize:11,fontWeight:700,cursor:"pointer",border:"none",
                              background:item.type==="num"?"rgba(139,90,43,0.14)":"rgba(83,74,183,0.10)",
                              color:item.type==="num"?T.tagV:T.tagR}}>
                            {item.type==="num"?"numeric":"dimension"}
                          </button>
                        </td>
                        <td style={{padding:"9px 13px",fontSize:13,borderBottom:"0.5px solid "+T.border,color:item.nullPct>20?T.danger:T.textMd}}>{item.nullPct}%</td>
                        <td style={{padding:"9px 13px",fontSize:13,borderBottom:"0.5px solid "+T.border,color:T.text}}>{item.uniqueCount.toLocaleString()}</td>
                        <td style={{padding:"9px 13px",fontSize:12,borderBottom:"0.5px solid "+T.border,color:slicerOk?T.success:T.warning,fontWeight:600}}>{slicerOk?"Yes":"Too many"}</td>
                        <td style={{padding:"9px 13px",fontSize:12,color:T.textMd,borderBottom:"0.5px solid "+T.border}}>
                          {item.sample.map((v,j)=><span key={j} style={{display:"inline-block",background:T.bgStat,borderRadius:4,padding:"1px 6px",marginRight:4,fontSize:11,border:"0.5px solid "+T.border,maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",verticalAlign:"middle"}}>{v}</span>)}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,overflow:"hidden"}}>
            <div style={{padding:"10px 16px",background:T.bgTableH,borderBottom:"1px solid "+T.border}}>
              <span style={{fontWeight:700,fontSize:13,color:T.primary}}>Data preview</span>
              <span style={{fontSize:12,color:T.textMd,marginLeft:8}}>first 8 rows</span>
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{borderCollapse:"collapse",minWidth:"100%"}}>
                <thead><tr style={{background:T.bgTableH}}>
                  {schema.map(item=><th key={item.field} style={{padding:"8px 13px",textAlign:item.type==="num"?"right":"left",fontSize:11,fontWeight:700,color:item.type==="num"?T.tagV:T.primary,borderBottom:"0.5px solid "+T.border,whiteSpace:"nowrap"}}>{item.field}</th>)}
                </tr></thead>
                <tbody>
                  {previewRows.map((row,i)=>(
                    <tr key={i} style={{background:i%2===0?T.bgCard:T.bgAlt}}>
                      {schema.map(item=><td key={item.field} style={{padding:"7px 13px",fontSize:12,textAlign:item.type==="num"?"right":"left",borderBottom:"0.5px solid "+T.border,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:T.text}}>
                        {row[item.field]===""||row[item.field]===null||row[item.field]===undefined?<span style={{color:T.textMd}}>-</span>:String(row[item.field])}
                      </td>)}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Admin View ─────────────────────────────────────────────────────────────────
function AdminView({onLogout,savedReports,publishedId,onSaveReport,onPublishReport,onUnpublishReport,onDeleteReport,onLoadReportData,onReloadReports,currentUser,currentRole}) {
  const libs=useLibs();
  const isMobile=useViewport();
  const [dataset,setDataset]=useState(null);
  const [config,setConfig]=useState(null);
  const [typeOverrides,setTypeOverrides]=useState({});
  const [cardFields,setCardFields]=useState([]);
  const [tab,setTab]=useState("upload");
  const [toast,setToast]=useState("");
  const [showSettings,setShowSettings]=useState(false);
  const [apiLoading,setApiLoading]=useState(false);
  const [activeReportId,setActiveReportId]=useState(null); // id of report currently open in builder
  const [saveDialog,setSaveDialog]=useState(false); // show overwrite/new dialog
  const [activeTabIdx,setActiveTabIdx]=useState(0); // active tab in multi-tab report
  const [accessPanel,setAccessPanel]=useState(null); // {id, name} of report whose access is being managed
  const [collabPanel,setCollabPanel]=useState(null); // {id,name,config} for collab setup
  const [collabViewPanel,setCollabViewPanel]=useState(null); // {id,name,report} for collab data view
  // Global filter store — survives Settings navigation
  const [adminGlobalFilters,setAdminGlobalFilters]=useState({});
  const [adminGlobalPivotFilters,setAdminGlobalPivotFilters]=useState({});

  const adminFilterKey=(tabIdx)=>(activeReportId||"new")+":"+tabIdx;

  const effectiveNumFields=useMemo(()=>{
    if (!dataset) return new Set();
    const s=new Set(dataset.numFields);
    Object.entries(typeOverrides).forEach(([f,t])=>{if(t==="num")s.add(f);else s.delete(f);});
    return s;
  },[dataset,typeOverrides]);

  function onDataLoaded(ds){setDataset(ds);setConfig(ds.config);setTypeOverrides({});setCardFields([]);setActiveReportId(null);setActiveTabIdx(0);setTab("builder");}
  async function onDataRefresh(ds, targetId, {skipFeedback=false}={}) {
    // targetId = which saved report to update data for
    // skipFeedback=true when called from onQuickRefresh (which owns its own loading/toast)
    const r = savedReports.find(x=>x.id===targetId);
    if (!r) { showToast("Report not found."); return; }
    if (!skipFeedback) setApiLoading(true);
    try {
      // Use ds.config if provided (e.g. onQuickRefresh bakes in fresh lastRefreshed timestamp)
      // so handleSaveReport writes it to DB in a single PUT — no separate updateReportConfig needed.
      const saveConfig = ds.config || r.config || {};
      const id = await onSaveReport({
        name: r.name,
        dataset: {...ds, numFields: ds.numFields},
        config: saveConfig,
        cardFields: r.cardFields||[],
        updateId: targetId,
      });
      // Update builder dataset + config so the builder has something to render.
      // handleSaveReport already cleared dataCache + localStorage cache.
      if (activeReportId === targetId || !activeReportId) {
        // Rows: merge fresh data over whatever was there (or start fresh if dataset was null)
        setDataset(prev => prev ? {...prev,...ds} : ds);
        setActiveReportId(id);
        // Config: always restore from the report so the builder is never blank.
        // This is the same initialisation logic as openSavedReport.
        const rtabs = r.config?.tabs;
        if (rtabs && rtabs.length > 0 && rtabs[0]?.config) {
          setConfig({...r.config, ...rtabs[0].config, name: r.config.name, tabs: rtabs});
          setCardFields(rtabs[0].cardFields || r.cardFields || []);
        } else {
          setConfig(r.config || {});
          setCardFields(r.cardFields || []);
        }
        setActiveTabIdx(0);
      }
      if (!skipFeedback) {
        showToast("✓ Data updated: "+r.name);
        setTab("builder"); // bring user to builder so they can see the refreshed pivot
      }
      return id;
    } catch(e) {
      if (!skipFeedback) showToast("Update failed: "+e.message);
      throw e;
    } finally {
      if (!skipFeedback) setApiLoading(false);
    }
  }
  async function openSavedReport(id) {
    const r=savedReports.find(x=>x.id===id);
    if (!r) return;
    setApiLoading(true);
    try {
      const data=await onLoadReportData(id);
      const ds={rows:data.rows,fields:data.fields,numFields:data.numFields};
      setDataset(ds);
      setTypeOverrides({});
      setActiveReportId(id);
      setActiveTabIdx(0); // always start on Tab 0 when opening a report

      // When the report has tabs, initialize the builder with Tab 0's structural
      // config (rows/columns/values etc.) merged in — never rely on top-level
      // structural fields which may be stripped or belong to a different tab.
      const tabs = r.config && r.config?.tabs;
      // Seed adminGlobalFilters from each tab's saved defaultFilters
      // so filters are restored after page refresh or report switch
      const newAdminFilters = {};
      const newAdminPivotFilters = {};
      if (tabs && tabs.length > 0) {
        tabs.forEach((t, i) => {
          const key = r.id + ":" + i;
          if (t.config && t.config.defaultFilters)
            newAdminFilters[key] = t.config.defaultFilters;
          if (t.config && t.config.defaultPivotFilters)
            newAdminPivotFilters[key] = t.config.defaultPivotFilters;
        });
      } else {
        const key = r.id + ":0";
        if (r.config.defaultFilters) newAdminFilters[key] = r.config.defaultFilters;
        if (r.config.defaultPivotFilters) newAdminPivotFilters[key] = r.config.defaultPivotFilters;
      }
      setAdminGlobalFilters(prev=>({...prev,...newAdminFilters}));
      setAdminGlobalPivotFilters(prev=>({...prev,...newAdminPivotFilters}));
      if (tabs && tabs.length > 0 && tabs[0] && tabs[0].config) {
        const tab0 = tabs[0];
        setConfig({
          ...r.config,
          ...tab0.config,
          name: r.config?.name,
          tabs: tabs,
        });
        setCardFields(tab0.cardFields || r.cardFields || []);
      } else {
        setConfig(r.config);
        setCardFields(r.cardFields || []);
      }

      setTab("builder");
    } catch(e){showToast("Load error: "+e.message);}
    finally{setApiLoading(false);}
  }
  const showToast=msg=>{setToast(msg);setTimeout(()=>setToast(""),3000);};
  async function doSave() {
    if (!dataset||!config){showToast("Nothing to save yet.");return;}
    // If this dataset came from an existing saved report, offer overwrite or new
    if (activeReportId) {
      setSaveDialog(true);
    } else {
      await commitSave(false);
    }
  }
  async function commitSave(overwrite) {
    setSaveDialog(false);
    setApiLoading(true);
    try{
      // Sync active tab's live builder state into its tabs slot
      let configToSave = {...config};
      if (config.tabs && config.tabs.length > 0) {
        const updatedTabs = config.tabs.map((t,i)=>{
          if (i!==activeTabIdx) return t;
          const existingTabCfg = t.config || {};
          return {
            ...t,
            config:{
              ...config,
              // Preserve defaultFilters from tab config — commitSave must not wipe filters
              // set by "Save filters". config.defaultFilters is at top level and may be
              // undefined even after "Save filters" (which writes to tabs[i].config).
              defaultFilters: existingTabCfg.defaultFilters || config.defaultFilters,
              defaultPivotFilters: existingTabCfg.defaultPivotFilters || config.defaultPivotFilters,
              tabs:undefined,
              name:undefined,
            },
            cardFields:[...cardFields],
          };
        });
        // CRITICAL: strip all tab-structural fields from top-level config.
        // They must live ONLY inside tabs[i].config — never at top level.
        // This prevents any tab's rows/columns/values bleeding into other tabs.
        const {
          rows, columns, values, filters,
          defaultFilters, defaultPivotFilters,
          colExcluded,
          ...topLevelOnly
        } = config;
        configToSave = {...topLevelOnly, tabs: updatedTabs};
      }
      // If overwrite: update in-place (keeps publish status + access assignments)
      // If new: create fresh
      const updateId = overwrite && activeReportId ? activeReportId : null;
      const id=await onSaveReport({name:config.name,dataset:{...dataset,numFields:effectiveNumFields},config:configToSave,cardFields,updateId});
      setActiveReportId(id);
      showToast(overwrite?"Report updated! Publish status and access preserved.":"Report saved as new!");
    }catch(e){showToast("Save failed: "+e.message);}
    finally{setApiLoading(false);}
  }
  async function doPublish(id) {
    setApiLoading(true);
    try{
      await onPublishReport(id);
      showToast("Report published!");
    }catch(e){showToast("Publish failed: "+e.message);}
    finally{setApiLoading(false);}
  }
  async function doUnpublish(id) {
    setApiLoading(true);
    try{
      await onUnpublishReport(id);
      showToast("Report unpublished.");
    }catch(e){showToast("Unpublish failed: "+e.message);}
    finally{setApiLoading(false);}
  }

  function toggleFieldType(field) {
    const curNum=effectiveNumFields.has(field);
    setTypeOverrides(p=>({...p,[field]:curNum?"dim":"num"}));
    if (curNum) setConfig(c=>({...c,values:c.values.filter(v=>v.field!==field)}));
  }

  function toggleCard(field){setCardFields(cf=>cf.some(x=>x.field===field)?cf.filter(x=>x.field!==field):[...cf,{field,agg:"sum"}]);}
  function setCardAgg(field,agg){setCardFields(cf=>cf.map(x=>x.field===field?{...x,agg}:x));}

  function toggleField(zone,field) {
    setConfig(c=>{
      let rows=[...c.rows],cols=[...c.columns],vals=[...c.values],filters=[...c.filters];
      if (zone==="rows"){if(rows.includes(field))rows=rows.filter(f=>f!==field);else{cols=cols.filter(f=>f!==field);rows=[...rows,field];}}
      else if(zone==="columns"){if(cols.includes(field))cols=cols.filter(f=>f!==field);else{rows=rows.filter(f=>f!==field);cols=[...cols,field];}}
      else if(zone==="values"){if(vals.some(v=>v.field===field))vals=vals.filter(v=>v.field!==field);else vals=[...vals,{field,agg:"sum"}];}
      else if(zone==="filters"){if(filters.includes(field))filters=filters.filter(f=>f!==field);else filters=[...filters,field];}
      return{...c,rows,columns:cols,values:vals,filters};
    });
  }

  function removeFrom(zone,field){
    setConfig(c=>({...c,[zone]:zone==="values"?c.values.filter(v=>v.field!==field):c[zone].filter(f=>f!==field)}));
  }

  function setAgg(field,agg,isCurrency){setConfig(c=>({...c,values:c.values.map(v=>v.field===field?{...v,agg,...(isCurrency!==undefined?{isCurrency}:{})}:v)}));}

  function reorderInZone(zone,fromField,toField) {
    setConfig(c=>{
      if (zone==="values"){
        const arr=[...c.values];
        const fi=arr.findIndex(v=>v.field===fromField), ti=arr.findIndex(v=>v.field===toField);
        if (fi===-1||ti===-1) return c;
        const [mv]=arr.splice(fi,1); arr.splice(ti,0,mv);
        return{...c,values:arr};
      }
      const arr=[...c[zone]];
      const fi=arr.indexOf(fromField), ti=arr.indexOf(toField);
      if (fi===-1||ti===-1) return c;
      arr.splice(fi,1); arr.splice(ti,0,fromField);
      return{...c,[zone]:arr};
    });
  }

  // Apply saved tab filters to the live preview so it matches what users see
  const previewFilters=adminGlobalFilters[adminFilterKey(activeTabIdx)]||{};
  const preview=useMemo(()=>dataset&&config?runPivot(dataset.rows,config,previewFilters):[],[dataset,config,previewFilters]);
  const fieldStatus=useMemo(()=>{
    if (!dataset||!config) return {};
    const z={};
    dataset.fields.forEach(f=>{z[f]={rows:config.rows.includes(f),cols:config.columns.includes(f),vals:config.values.some(v=>v.field===f),filters:config.filters.includes(f),card:cardFields.some(x=>x.field===f)};});
    return z;
  },[dataset,config,cardFields]);

  const isSubAdminUser = currentRole === 'subadmin_user';
  const isSuperAdmin   = currentRole === 'admin';

  const TABS=[
    ["upload","Upload"],
    ["builder","Report Builder",!dataset],
    ["preview","User Preview",!dataset],
    ["data","Raw Data",!dataset],
    ["reports","Reports ("+savedReports.length+")"],
    ...(isSubAdminUser ? [["myreports","📋 My Reports"]] : []),
    ...(!isSubAdminUser ? [["workflow","🤝 Workflow"]] : []),
  ];
  const tabBtn=(t,l,disabled)=>(
    <button key={t} onClick={()=>!disabled&&setTab(t)} style={{padding:"11px 16px",background:"none",border:"none",cursor:disabled?"not-allowed":"pointer",fontSize:13,
      borderBottom:tab===t?"2px solid "+T.accent:"2px solid transparent",
      fontWeight:tab===t?700:400,color:disabled?T.textMd:tab===t?T.textLt:"rgba(245,239,230,0.6)",opacity:disabled?0.4:1}}>
      {l}
    </button>
  );

  const roleLabel = currentRole === 'admin' ? 'Super Admin' : currentRole === 'subadmin' ? 'Sub-Admin' : currentRole === 'subadmin_user' ? 'Admin+User' : 'Admin';

  return(
    <div style={{minHeight:"100vh",background:T.bgPage,fontFamily:"system-ui,sans-serif"}}>
      <AppHeader role={roleLabel} onLogout={onLogout}>
        {toast&&<span style={{fontSize:12,color:T.textLt,background:"rgba(45,106,79,0.5)",padding:"4px 12px",borderRadius:6,fontWeight:500,border:"1px solid rgba(45,106,79,0.6)"}}>{toast}</span>}
        {dataset&&config&&<button onClick={doSave} disabled={apiLoading} style={{padding:"6px 14px",background:"rgba(255,255,255,0.15)",color:T.textLt,border:"1px solid rgba(255,255,255,0.25)",borderRadius:6,cursor:apiLoading?"wait":"pointer",fontSize:12,fontWeight:600,opacity:apiLoading?0.6:1}}>
          {apiLoading?"Saving…":"Save Report"}
        </button>}
        <button onClick={()=>setShowSettings(true)} title="User management & settings"
          style={{padding:"6px 12px",background:"rgba(255,255,255,0.12)",color:T.textLt,border:"1px solid rgba(255,255,255,0.2)",borderRadius:6,cursor:"pointer",fontSize:12}}>
          ⚙ Settings
        </button>
      </AppHeader>

      <div style={{position:"sticky",top:52,zIndex:40,background:T.bgHeader,borderBottom:"1px solid "+T.borderHd,
        padding:"0 8px",display:"flex",overflowX:"auto",WebkitOverflowScrolling:"touch",
        scrollbarWidth:"none",msOverflowStyle:"none"}}>
        {TABS.map(([t,l,d])=>tabBtn(t,l,d))}
      </div>

      {tab==="upload"&&<UploadTab libs={libs} onDataLoaded={onDataLoaded} onDataRefresh={savedReports.length?onDataRefresh:null}
        existingConfig={config} savedReports={savedReports}
        savedLinks={savedReports.flatMap(r=>getSourceLinks(r.config).map(lk=>({...lk,reportId:r.id,label:lk.label||r.name})))}
        onDeleteLink={async(lk)=>{
          const r=savedReports.find(x=>x.id===lk.reportId);
          if(!r){showToast("Report not found");return;}
          try{
            const cfg=r.config||{};
            const newLinks=getSourceLinks(cfg).filter(x=>x.url!==lk.url);
            await updateReportConfig(lk.reportId,{...cfg,sourceLinks:newLinks});
            await onReloadReports();
            showToast("✓ URL link removed from "+r.name);
          }catch(e){showToast("Failed to remove link: "+e.message);}
        }}
        onUpdateLink={async({reportId,updLinks,cfg,newUrl,newSheet})=>{
          const r=savedReports.find(x=>x.id===reportId);
          if(!r){showToast("Report not found");return;}
          try{
            await updateReportConfig(reportId,{...cfg,sourceLinks:updLinks});
            if(newUrl){
              // Auto-fetch from the new URL so the user doesn't have to click ↻ manually
              setApiLoading(true);
              try{
                let result;
                try{
                  const resp=await fetch(newUrl,{credentials:"include",redirect:"follow",cache:"no-store"});
                  if(resp.ok){const ct=resp.headers.get("content-type")||"";if(!ct.includes("text/html")){const buf=await resp.arrayBuffer();const wb=window.XLSX.read(buf,{type:"array",cellDates:true});const wsName=newSheet&&wb.SheetNames.includes(newSheet)?newSheet:wb.SheetNames[0];const ws=wb.Sheets[wsName];if(ws){const rows=window.XLSX.utils.sheet_to_json(ws,{defval:null,cellDates:true,raw:true});result={rows,sheetNames:wb.SheetNames};}}}
                }catch(e){console.log("browser fetch failed:",e.message);}
                if(!result){result=await fetchUrlViaProxy(newUrl,newSheet||undefined);}
                const {rows:cleanRows,fields:cleanFields}=sanitizeRows(result.rows);
                const rc=r.config||{};
                const allValFields=[...(rc.values||[]).map(v=>v.field),...((rc.tabs||[]).flatMap(t=>(t.config?.values||[]).map(v=>v.field)))];
                const nfArr=[...new Set(allValFields.length?allValFields:cleanFields.filter(k=>typeof cleanRows[0]?.[k]==="number"))];
                const ts=Date.now();
                const tsLinks=updLinks.map(x=>x.url===newUrl?{...x,lastRefreshed:ts}:x);
                const freshConfig={...cfg,sourceLinks:tsLinks};
                await onDataRefresh({rows:cleanRows,fields:cleanFields,numFields:new Set(nfArr),config:freshConfig},reportId,{skipFeedback:true});
                showToast("✓ Link updated + "+cleanRows.length.toLocaleString()+" rows loaded: "+r.name);
              }catch(e){
                const msg=e.message||"";
                if(msg.includes("Connect your Microsoft")||msg.includes("needs_auth")||msg.includes("connect your Microsoft"))
                  showToast("✓ URL saved — reconnect Microsoft to refresh data");
                else
                  showToast("✓ URL saved — auto-refresh failed ("+msg+"). Use ↻ to retry.");
              }finally{setApiLoading(false);}
            } else {
              showToast("✓ Link updated for "+r.name);
            }
            await onReloadReports();
          }catch(e){showToast("Failed to update link: "+e.message);}
        }}
        onQuickRefresh={async(lk)=>{
          // Quick refresh: fetch data + update the linked report directly
          setApiLoading(true);
          try{
            if(!lk||!lk.url){showToast("No URL to refresh");return;}
            // Try browser first then backend proxy.
            // cache:"no-store" forces the browser to always hit the network so we
            // never serve a stale cached XLSX when the SharePoint file was updated.
            let result;
            try{
              const resp=await fetch(lk.url,{credentials:"include",redirect:"follow",cache:"no-store"});
              if(resp.ok){const ct=resp.headers.get("content-type")||"";if(!ct.includes("text/html")){const buf=await resp.arrayBuffer();const wb=window.XLSX.read(buf,{type:"array",cellDates:true});const wsName=lk.sheet&&wb.SheetNames.includes(lk.sheet)?lk.sheet:wb.SheetNames[0];const ws=wb.Sheets[wsName];if(ws){
                      if(lk.rangeOverride&&lk.rangeOverride.trim()){
                        try{window.XLSX.utils.decode_range(lk.rangeOverride);ws["!ref"]=lk.rangeOverride.trim().toUpperCase();}catch(e){}
                      }
                      const rows=window.XLSX.utils.sheet_to_json(ws,{defval:null,cellDates:true,raw:true});result={rows,sheetNames:wb.SheetNames};}}}
            }catch(e){console.log("browser fetch failed:",e.message);}
            if(!result){result=await fetchUrlViaProxy(lk.url,lk.sheet||undefined,lk.rangeOverride||undefined);}
            // Sanitize rows (trim column names, drop blank rows) so field names
            // always match what the config stored (same path as the initial upload).
            const {rows:cleanRows,fields:cleanFields}=sanitizeRows(result.rows);
            // Build numFields from the target report's config
            const r=savedReports.find(x=>x.id===lk.reportId);
            if(!r){showToast("Report not found for this link");setApiLoading(false);return;}
            const rc=r.config||{};
            const allValFields=[
              ...(rc.values||[]).map(v=>v.field),
              ...((rc.tabs||[]).flatMap(t=>(t.config?.values||[]).map(v=>v.field))),
            ];
            const nfArr=[...new Set(allValFields.length
              ? allValFields
              : cleanFields.filter(k=>typeof cleanRows[0]?.[k]==="number")
            )];
            // Bake the new timestamp into the config BEFORE saving so
            // handleSaveReport writes the fresh timestamp to the DB in one shot
            // (avoids the race where a separate updateReportConfig call loses to onReloadReports).
            const ts = Date.now();
            const rCfg = r?.config||{};
            const tsLinks = getSourceLinks(rCfg).map(x=>x.url===lk.url?{...x,lastRefreshed:ts}:x);
            if (!tsLinks.find(x=>x.url===lk.url)) tsLinks.push({...lk,lastRefreshed:ts});
            const freshConfig = {...rCfg, sourceLinks:tsLinks};
            await onDataRefresh({rows:cleanRows,fields:cleanFields,numFields:new Set(nfArr),config:freshConfig},lk.reportId,{skipFeedback:true});
            // Mirror the new timestamp into local config state so the UI updates instantly
            setConfig(cfg=>{
              if (!cfg) return cfg;
              const existing = getSourceLinks(cfg);
              const sl = existing.map(x=>x.url===lk.url?{...x,lastRefreshed:ts}:x);
              if (!existing.find(x=>x.url===lk.url)) sl.push({...lk,lastRefreshed:ts});
              return {...cfg, sourceLinks:sl};
            });
            await onReloadReports();
            showToast("✓ "+result.rows.length.toLocaleString()+" rows refreshed: "+lk.label);
          }catch(e){
            // Provide actionable message for Microsoft/Google auth failures
            const msg=e.message||"";
            if(msg.includes("Connect your Microsoft")||msg.includes("needs_auth")||msg.includes("connect your Microsoft"))
              showToast("⚠ SharePoint connection expired — go to Upload → Connect Microsoft Account to reconnect");
            else if(msg.includes("Connect your Google")||msg.includes("google"))
              showToast("⚠ Google Drive connection expired — go to Upload → Connect Google Account to reconnect");
            else
              showToast("Refresh failed: "+msg);
          }
          finally{setApiLoading(false);}
        }}/>}

      {tab==="builder"&&dataset&&config&&(
        <div style={{padding:isMobile?12:20,display:"grid",
          gridTemplateColumns:isMobile?"1fr":"290px 1fr",
          gap:isMobile?12:20,alignItems:"start"}}>

          {/* Left panel */}
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,padding:14}}>
              <div style={{fontWeight:700,fontSize:13,color:T.primary,marginBottom:2}}>{dataset.fields.length} fields · {dataset.rows.length.toLocaleString()} rows</div>
              <div style={{fontSize:11,color:T.textMd,marginBottom:10}}>{config.name}</div>

              {/* Legend */}
              <div style={{display:"flex",flexWrap:"wrap",gap:7,marginBottom:12,padding:"9px 10px",background:T.bgStat,borderRadius:8,border:"0.5px solid "+T.border}}>
                {[{L:"#/Aa",c:T.tagV,t:"Type toggle"},{L:"R",c:T.tagR,t:"Rows"},{L:"C",c:T.tagC,t:"Cols"},{L:"V",c:T.tagV,t:"Values"},{L:"F",c:T.tagF,t:"Filters"},{L:"K",c:T.tagK,t:"Card filter"}].map(b=>(
                  <div key={b.L} style={{display:"flex",alignItems:"center",gap:4,fontSize:10,color:T.textMd}}>
                    <span style={{padding:"1px 5px",borderRadius:3,background:b.c,color:"white",fontSize:9,fontWeight:700}}>{b.L}</span>{b.t}
                  </div>
                ))}
              </div>

              <div style={{borderTop:"0.5px solid "+T.border,paddingTop:10,display:"flex",flexDirection:"column",maxHeight:520,overflowY:"auto"}}>
                <FieldSearch fields={dataset.fields} numFields={effectiveNumFields}
                  fieldStatus={fieldStatus} onToggle={toggleField}
                  onToggleType={f=>toggleFieldType(f)} onToggleCard={f=>toggleCard(f)}/>
              </div>
            </div>

            <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,padding:14}}>
              <div style={{fontWeight:700,fontSize:13,color:T.primary,marginBottom:8}}>Report Name</div>
              <input value={config.name} onChange={e=>setConfig(c=>({...c,name:e.target.value}))}
                style={{width:"100%",padding:"7px 10px",border:"1px solid "+T.border,borderRadius:6,fontSize:13,background:T.bgStat,color:T.text,boxSizing:"border-box",outline:"none"}}/>
            </div>

            {/* Number format options */}
            <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,padding:14}}>
              <div style={{fontWeight:700,fontSize:13,color:T.primary,marginBottom:4}}>Number Formats</div>
              <div style={{fontSize:11,color:T.textMd,marginBottom:8}}>Which units users can switch between (first = default)</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                {NUM_FORMATS.map(f=>{
                  const cur=config.allowedFmts||[];
                  const on=cur.length===0||cur.includes(f.key);
                  return(
                    <label key={f.key} style={{display:"flex",alignItems:"center",gap:5,
                      padding:"5px 12px",border:"1px solid "+(on?T.primary:T.border),
                      borderRadius:6,cursor:"pointer",background:on?"rgba(92,45,26,0.07)":"none",
                      fontSize:12,color:on?T.primary:T.textMd,fontWeight:on?600:400,userSelect:"none"}}>
                      <input type="checkbox" checked={on}
                        onChange={()=>{
                          let next;
                          if(cur.length===0){
                            // Was "all" — uncheck this one
                            next=NUM_FORMATS.map(x=>x.key).filter(k=>k!==f.key);
                          } else if(on){
                            next=cur.filter(k=>k!==f.key);
                            if(next.length===0) return; // keep at least one
                          } else {
                            next=[...cur,f.key];
                            if(next.length===NUM_FORMATS.length) next=[];
                          }
                          setConfig(c=>({...c,allowedFmts:next}));
                        }}
                        style={{accentColor:T.primary,cursor:"pointer"}}/>
                      {f.label}
                    </label>
                  );
                })}
              </div>
              <div style={{fontSize:10,color:T.textMd,marginTop:6}}>
                {(!config.allowedFmts||config.allowedFmts.length===0)
                  ? "All 5 options shown · Recommended: check only Crores + Units"
                  : `${config.allowedFmts.length} option${config.allowedFmts.length>1?"s":""} shown · Default: ${NUM_FORMATS.find(f=>f.key===(config.allowedFmts||[])[0])?.label||""}`}
              </div>
            </div>
          </div>

          {/* Right panel */}
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
              <ZoneBox label="Row Labels (R)" color={T.tagR} zone="rows" fields={config.rows}
                onRemove={f=>removeFrom("rows",f)} onReorder={(a,b)=>reorderInZone("rows",a,b)}
                emptyMsg="Press R on any field"/>
              <ZoneBox label="Column Labels (C)" color={T.tagC} zone="columns" fields={config.columns}
                onRemove={f=>removeFrom("columns",f)} onReorder={(a,b)=>reorderInZone("columns",a,b)}
                emptyMsg="Press C on any field"/>
            </div>
            <ZoneBox label="Values (V) — multiple metrics, drag to reorder" color={T.tagV} zone="values"
              fields={config.values} isValues onAggChange={setAgg}
              onRemove={f=>removeFrom("values",f)} onReorder={(a,b)=>reorderInZone("values",a,b)}
              emptyMsg="Press V on a numeric field"/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
              <ZoneBox label="Filters / Slicers (F)" color={T.tagF} zone="filters" fields={config.filters}
                onRemove={f=>removeFrom("filters",f)} onReorder={(a,b)=>reorderInZone("filters",a,b)}
                emptyMsg="Press F on any field"/>
              <ZoneBox label="Card Filters (K) — Power BI style" color={T.tagK} zone="cards" fields={cardFields}
                isValues onAggChange={setCardAgg}
                onRemove={f=>setCardFields(cf=>cf.filter(x=>x.field!==f))}
                onReorder={(a,b)=>setCardFields(cf=>{
                  const fi=cf.findIndex(x=>x.field===a),ti=cf.findIndex(x=>x.field===b);
                  if(fi===-1||ti===-1)return cf;const arr=[...cf];arr.splice(fi,1);arr.splice(ti,0,cf[fi]);return arr;
                })}
                emptyMsg="Press K on any field"/>
            </div>
            <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,padding:14}}>
              <div style={{fontWeight:700,fontSize:13,color:T.primary,marginBottom:12}}>Live Preview</div>
              <PivotTable result={preview} numFmt="Cr"/>
            </div>
          </div>
        </div>
      )}

      {tab==="preview"&&dataset&&config&&(
        <div style={{padding:20}}>
          <div style={{fontWeight:700,fontSize:18,color:T.primary,marginBottom:3}}>{config.name}</div>
          <div style={{fontSize:12,color:T.textMd,marginBottom:18}}>Preview — what users see · click cells to drill down</div>
          <Report key={"preview_"+(activeReportId||"new")} config={config} data={dataset.rows} fields={dataset.fields} numFields={effectiveNumFields} showExport cardFields={cardFields}
            externalFilters={adminGlobalFilters[adminFilterKey(activeTabIdx)]}
            externalPivotFilters={adminGlobalPivotFilters[adminFilterKey(activeTabIdx)]}
            onExternalFiltersChange={(f)=>setAdminGlobalFilters(prev=>({...prev,[adminFilterKey(activeTabIdx)]:f}))}
            onExternalPivotFiltersChange={(pf)=>setAdminGlobalPivotFilters(prev=>({...prev,[adminFilterKey(activeTabIdx)]:pf}))}
            tabs={config.tabs||null}
            activeTabIdx={activeTabIdx}
            onTabChange={(idx)=>{
              if(idx===activeTabIdx)return;
              // Use functional setConfig so we always read latest state —
              // prevents stale-closure overwrites when called alongside onTabsChange
              setConfig(prevCfg=>{
                const prevTabs=prevCfg.tabs||[];
                // Save current active tab edits
                const savedTabs=prevTabs.map((t,i)=>
                  i===activeTabIdx
                    ? {...t,config:{...prevCfg,tabs:undefined,name:undefined},cardFields:[...cardFields]}
                    : t
                );
                const target=savedTabs[idx];
                if(target){
                  if(target.cardFields)setCardFields([...target.cardFields]);
                  return {...prevCfg,...target.config,name:prevCfg.name,tabs:savedTabs};
                }
                // Target doesn't exist yet (being created in same batch) — just return as-is
                return prevCfg;
              });
              setActiveTabIdx(idx);
            }}
            onTabDelete={(delIdx)=>{
              // Atomic: remove tab + load new active tab in one setConfig call
              // avoids the setConfig race that made tabs reappear after deletion
              setConfig(prevCfg=>{
                const nt=(prevCfg.tabs||[]).filter((_,i)=>i!==delIdx);
                if(!nt.length){if(setCardFields)setCardFields([]);return {...prevCfg,tabs:undefined};}
                const newIdx=delIdx>=nt.length?nt.length-1:delIdx;
                const target=nt[newIdx];
                if(target){
                  if(target.cardFields)setCardFields([...target.cardFields]);
                  return {...prevCfg,...target.config,name:prevCfg.name,tabs:nt};
                }
                return {...prevCfg,tabs:nt};
              });
              setActiveTabIdx(prev=>prev>=((config.tabs||[]).length-1)?Math.max(0,(config.tabs||[]).length-2):prev);
            }}
            onTabsChange={(newTabs)=>{
              // Called for: add/delete/rename/reorder. Always preserve active-tab edits.
              const synced=newTabs.map((t,i)=>{
                // If this tab IS the active one AND has not been replaced wholesale, sync from current config
                if (i===activeTabIdx && t && t.id===(config.tabs?.[activeTabIdx]?.id))
                  return {...t,config:{...config,tabs:undefined,name:undefined},cardFields:[...cardFields]};
                return t;
              });
              setConfig(c=>({...c,tabs:synced}));
            }}
            onDrillHiddenColsChange={(cols,fmts)=>{
              setConfig(prev=>{
                const newCfg={...prev,drillHiddenCols:cols,...(fmts?{drillColFmts:fmts}:{})};
                if(activeReportId){
                  (async()=>{
                    try{
                      await updateReportConfig(activeReportId,newCfg);
                      await onReloadReports();
                      showToast("✓ Drill-down layout saved!");
                    }catch(e){showToast("Layout save failed: "+e.message);}
                  })();
                } else {
                  showToast("Drill-down layout saved — click Save Report to persist.");
                }
                return newCfg;
              });
            }}
            onColExcludedChange={async(cols)=>{
              setConfig(prev=>{
                const newCfg={...prev,colExcluded:cols};
                if(activeReportId){
                  // Build save config the same way as commitSave — strip structural if tabbed
                  let cfgToSave=newCfg;
                  if(newCfg.tabs&&newCfg.tabs.length>0){
                    const{rows,columns,values,filters,defaultFilters,defaultPivotFilters,colExcluded:_ce,...topOnly}=newCfg;
                    const updatedTabs=newCfg.tabs.map((t,i)=>
                      i===activeTabIdx?{...t,config:{...config,tabs:undefined,name:undefined,colExcluded:cols},cardFields:[...cardFields]}:t
                    );
                    cfgToSave={...topOnly,tabs:updatedTabs};
                  }
                  (async()=>{
                    try{await updateReportConfig(activeReportId,cfgToSave);}
                    catch(e){/* silent */}
                  })();
                }
                return newCfg;
              });
            }}
            onFiltersChange={(f,pf)=>setConfig(c=>({...c,defaultFilters:f,...(pf!==undefined?{defaultPivotFilters:pf}:{})}))}
            onSaveFilters={async(f,pf)=>{
              if (!activeReportId){showToast("Save the report first before locking filters.");return;}
              setApiLoading(true);
              setConfig(prev=>{
                let newCfg;
                if (prev.tabs&&prev.tabs.length>0) {
                  // Use prev (guaranteed fresh) instead of stale closure `config`
                  const activeTabCfg = prev.tabs[activeTabIdx]?.config || {};
                  const updatedTabs=prev.tabs.map((t,i)=>
                    i===activeTabIdx
                      ? {
                          ...t,
                          config:{
                            ...activeTabCfg,   // tab's own saved structural state
                            defaultFilters:f,  // new filter — always wins
                            defaultPivotFilters:pf||{},
                          },
                          cardFields: t.cardFields||[],
                        }
                      : t
                  );
                  // Local state: keep all fields intact so builder doesn't break
                  newCfg={...prev,tabs:updatedTabs};
                  // DB payload: strip structural from top level (tabs[i].config has them)
                  const {rows:_r,columns:_c,values:_v,filters:_f2,defaultFilters:_df,defaultPivotFilters:_dpf,colExcluded:_ce,...topOnly}=prev;
                  const cfgForDB={...topOnly,tabs:updatedTabs};
                  (async()=>{
                    try{
                      await updateReportConfig(activeReportId,cfgForDB);
                      await onReloadReports();
                      showToast("✓ Filters saved!");
                    }catch(e){showToast("Save failed: "+e.message);}
                    finally{setApiLoading(false);}
                  })();
                } else {
                  newCfg={...prev,defaultFilters:f,defaultPivotFilters:{0:pf||{}}};
                  (async()=>{
                    try{
                      await updateReportConfig(activeReportId,newCfg);
                      await onReloadReports();
                      showToast("✓ Filters saved!");
                    }catch(e){showToast("Save failed: "+e.message);}
                    finally{setApiLoading(false);}
                  })();
                }
                return newCfg;
              });
            }}/>
        </div>
      )}

      {tab==="data"&&dataset&&(
        <div style={{padding:20}}>
          <div style={{fontSize:13,color:T.textMd,marginBottom:12}}>First 100 of {dataset.rows.length.toLocaleString()} rows · {dataset.fields.length} columns (in original order)</div>
          <div style={{overflowX:"auto",borderRadius:10,border:"1px solid "+T.border}}>
            <table style={{borderCollapse:"collapse",minWidth:"100%",fontSize:12}}>
              <thead><tr style={{background:T.bgHeader}}>
                {dataset.fields.map(f=><th key={f} style={{padding:"9px 13px",textAlign:effectiveNumFields.has(f)?"right":"left",fontWeight:700,fontSize:11,color:effectiveNumFields.has(f)?T.accent:T.textLt,borderBottom:"1px solid "+T.borderHd,whiteSpace:"nowrap"}}>{f}</th>)}
              </tr></thead>
              <tbody>{dataset.rows.slice(0,100).map((row,i)=>(
                <tr key={i} style={{background:i%2===0?T.bgCard:T.bgAlt}}>
                  {dataset.fields.map(f=><td key={f} style={{padding:"7px 13px",borderBottom:"0.5px solid "+T.border,textAlign:effectiveNumFields.has(f)?"right":"left",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:T.text}}>
                    {effectiveNumFields.has(f)?(+row[f]).toLocaleString():String(row[f]||"")}
                  </td>)}
                </tr>
              ))}</tbody>
            </table>
          </div>
        </div>
      )}
      {tab==="reports"&&(
        <ReportsTab
          savedReports={savedReports}
          publishedId={publishedId}
          onOpen={openSavedReport}
          onDelete={onDeleteReport}
          onPublish={doPublish}
          onUnpublish={doUnpublish}
          onReload={onReloadReports}
          onAccessPanel={(id,name)=>setAccessPanel({id,name})}/>
      )}
      {tab==="myreports"&&isSubAdminUser&&(
        <MyReportsViewer
          savedReports={savedReports}
          onLoadReportData={onLoadReportData}
          onLogout={onLogout}/>
      )}
      {tab==="workflow"&&!isSubAdminUser&&(
        <WorkflowListTab
          savedReports={savedReports}
          currentUser={currentUser}
          currentRole={currentRole}
          onSetup={(r)=>setCollabPanel({id:r.id,name:r.name,report:r})}
          onOpenView={(r)=>setCollabViewPanel({id:r.id,name:r.name,report:r})}
          onReload={onReloadReports}/>
      )}
      {showSettings&&<SettingsPanel currentUser={currentUser} currentRole={currentRole} onClose={()=>setShowSettings(false)}/>}
      {accessPanel&&<ReportAccessPanel reportId={accessPanel.id} reportName={accessPanel.name} onClose={()=>setAccessPanel(null)}/>}
      {collabPanel&&<CollabSetupPanel report={collabPanel.report} currentUser={currentUser} currentRole={currentRole} onClose={()=>{setCollabPanel(null);onReloadReports();}}/>}
      {collabViewPanel&&<CollabDataView report={collabViewPanel.report} currentUser={currentUser} currentRole={currentRole} onClose={()=>setCollabViewPanel(null)}/>}
      {saveDialog&&(
        <div style={{position:"fixed",inset:0,zIndex:600,background:"rgba(44,24,16,0.5)",display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:T.bgCard,borderRadius:12,padding:28,width:"min(420px,90vw)",boxShadow:"0 12px 40px rgba(44,24,16,0.3)"}}>
            <div style={{fontWeight:700,fontSize:16,color:T.primary,marginBottom:8}}>Save Report</div>
            <div style={{fontSize:13,color:T.textMd,marginBottom:20,lineHeight:1.6}}>
              This report was previously saved. Do you want to <strong>overwrite</strong> the existing version, or <strong>save as a new</strong> report?
            </div>
            <div style={{display:"flex",gap:10,justifyContent:"flex-end",flexWrap:"wrap"}}>
              <button onClick={()=>setSaveDialog(false)} style={{padding:"8px 16px",background:"none",border:"1px solid "+T.border,borderRadius:7,cursor:"pointer",fontSize:13,color:T.text}}>
                Cancel
              </button>
              <button onClick={()=>commitSave(false)} style={{padding:"8px 16px",background:"none",border:"1px solid "+T.primary,borderRadius:7,cursor:"pointer",fontSize:13,color:T.primary,fontWeight:600}}>
                Save as New
              </button>
              <button onClick={()=>commitSave(true)} style={{padding:"8px 18px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,cursor:"pointer",fontSize:13,fontWeight:700}}>
                Overwrite
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Phase 3: Workflow List Tab ────────────────────────────────────────────────
// Shows all reports; collab-enabled ones get an "Open Workflow" button
function WorkflowListTab({savedReports,currentUser,currentRole,onSetup,onOpenView,onReload}) {
  const [toggling,setToggling]=useState({});
  const isBuilder=['admin','subadmin'].includes(currentRole);

  const handleToggle=async(r)=>{
    setToggling(t=>({...t,[r.id]:true}));
    try{
      await toggleCollab(r.id,!r.config?.collab_enabled);
      onReload();
    }catch(e){alert("Error: "+e.message);}
    finally{setToggling(t=>({...t,[r.id]:false}));}
  };

  const collabReports=savedReports.filter(r=>r.config?.collab_enabled);
  const otherReports=savedReports.filter(r=>!r.config?.collab_enabled);

  const ReportCard=({r})=>(
    <div key={r.id} style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,padding:"14px 16px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
      <div style={{flex:1,minWidth:140}}>
        <div style={{fontWeight:600,fontSize:14,color:T.text}}>{r.name}</div>
        <div style={{fontSize:11,color:T.textMd,marginTop:2}}>{r.row_count||0} rows · {r.config?.collab_enabled?"🟢 Collab ON":"⚪ Standard"}</div>
      </div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        {r.config?.collab_enabled&&(
          <>
            {isBuilder&&<button onClick={()=>onSetup(r)}
              style={{padding:"6px 14px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600}}>
              ⚙ Setup Columns & Cycle
            </button>}
            <button onClick={()=>onOpenView(r)}
              style={{padding:"6px 14px",background:"#3B5998",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600}}>
              📋 Open Workflow View
            </button>
          </>
        )}
        {isBuilder&&<button onClick={()=>handleToggle(r)} disabled={!!toggling[r.id]}
          style={{padding:"6px 14px",background:r.config?.collab_enabled?"#A32D2D":"#2D6A4F",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600,opacity:toggling[r.id]?0.5:1}}>
          {toggling[r.id]?"...":(r.config?.collab_enabled?"Disable Collab":"Enable Collab")}
        </button>}
      </div>
    </div>
  );

  return(
    <div style={{padding:"20px 16px",maxWidth:900,margin:"0 auto"}}>
      <div style={{fontSize:18,fontWeight:700,color:T.primary,marginBottom:6}}>🤝 Collaborative Workflow</div>
      <div style={{fontSize:13,color:T.textMd,marginBottom:20,lineHeight:1.6}}>
        Enable collab mode on any report to add input columns, open monthly cycles, and track approvals.
      </div>
      {collabReports.length>0&&(
        <>
          <div style={{fontSize:12,fontWeight:700,color:T.secondary,textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>Collab-Enabled Reports</div>
          <div style={{display:"flex",flexDirection:"column",gap:10,marginBottom:20}}>
            {collabReports.map(r=><ReportCard key={r.id} r={r}/>)}
          </div>
        </>
      )}
      {isBuilder&&otherReports.length>0&&(
        <>
          <div style={{fontSize:12,fontWeight:700,color:T.textMd,textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>Standard Reports (collab disabled)</div>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {otherReports.map(r=><ReportCard key={r.id} r={r}/>)}
          </div>
        </>
      )}
      {savedReports.length===0&&<div style={{color:T.textMd,fontSize:14,padding:"20px 0"}}>No reports yet. Create one in the Report Builder first.</div>}
    </div>
  );
}

// ── Field search wrapper for the Report Builder left-panel field list ─────────
function FieldSearch({fields,numFields,fieldStatus,onToggle,onToggleType,onToggleCard}){
  const [q,setQ]=useState("");
  const vis=q.trim()?fields.filter(f=>f.toLowerCase().includes(q.toLowerCase())):fields;
  return(
    <>
      {fields.length>8&&(
        <input value={q} onChange={e=>setQ(e.target.value)}
          placeholder={`Search ${fields.length} fields…`}
          style={{padding:"5px 9px",border:"0.5px solid "+T.border,borderRadius:5,fontSize:12,
            background:T.bgStat,color:T.text,outline:"none",marginBottom:6,width:"100%",boxSizing:"border-box"}}/>
      )}
      {vis.map(f=>(
        <FieldRow key={f} field={f} isNum={numFields.has(f)}
          status={fieldStatus[f]||{}} onToggle={onToggle}
          onToggleType={()=>onToggleType(f)} onToggleCard={()=>onToggleCard(f)}/>
      ))}
      {vis.length===0&&q&&<div style={{padding:"8px 4px",fontSize:11,color:T.textMd}}>No fields match "{q}"</div>}
    </>
  );
}

// ── Mini builder zone helpers (used by CollabSetupPanel) ─────────────────────
function ZoneEditor({label,color,hint,fields,allFields,onAdd,onRemove,onReorder}){
  const [open,setOpen]=useState(false);
  const [q,setQ]=useState("");
  const [dragIdx,setDragIdx]=useState(null);
  const available=allFields.filter(f=>!fields.includes(f));
  const visAvail=q.trim()?available.filter(f=>f.toLowerCase().includes(q.toLowerCase())):available;
  return(
    <div style={{background:T.bgAlt,border:"1px solid "+T.border,borderRadius:8,padding:10,minHeight:80}}>
      <div style={{fontSize:11,fontWeight:700,color,marginBottom:3,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span>{label}</span>
        <button onClick={()=>{setOpen(o=>!o);setQ("");}}
          style={{fontSize:10,background:color,color:"#fff",border:"none",borderRadius:4,padding:"1px 8px",cursor:"pointer"}}>+ Add</button>
      </div>
      <div style={{fontSize:10,color:T.textMd,marginBottom:6}}>{hint}</div>
      <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
        {fields.map((f,i)=>(
          <span key={f} draggable
            onDragStart={()=>setDragIdx(i)}
            onDragOver={e=>e.preventDefault()}
            onDrop={()=>{
              if(dragIdx===null||dragIdx===i)return;
              const n=[...fields];n.splice(dragIdx,1);n.splice(i,0,fields[dragIdx]);
              onReorder&&onReorder(n);setDragIdx(null);
            }}
            onDragEnd={()=>setDragIdx(null)}
            style={{display:"inline-flex",alignItems:"center",gap:3,padding:"2px 8px",
              background:dragIdx===i?"rgba(0,0,0,0.3)":color,color:"#fff",borderRadius:10,fontSize:11,
              cursor:"grab",opacity:dragIdx===i?0.5:1,userSelect:"none"}}>
            <span style={{fontSize:9,opacity:0.6,marginRight:1}}>⠿</span>
            {f}
            <button onClick={()=>onRemove(f)} style={{background:"none",border:"none",color:"rgba(255,255,255,0.8)",cursor:"pointer",fontSize:13,padding:0,lineHeight:1}}>×</button>
          </span>
        ))}
        {fields.length===0&&<span style={{fontSize:11,color:T.textMd,fontStyle:"italic"}}>Empty — click + Add</span>}
      </div>
      {open&&(
        <div style={{marginTop:6,borderTop:"1px solid "+T.border,paddingTop:6}}>
          <input value={q} onChange={e=>setQ(e.target.value)} placeholder={`Search ${available.length} fields…`}
            style={{width:"100%",padding:"4px 8px",border:"0.5px solid "+T.border,borderRadius:5,fontSize:11,
              background:T.bgCard,color:T.text,boxSizing:"border-box",outline:"none",marginBottom:4}}/>
          <div style={{maxHeight:130,overflowY:"auto"}}>
            {visAvail.map(f=>(
              <div key={f} onClick={()=>{onAdd(f);setOpen(false);setQ("");}}
                style={{padding:"4px 6px",fontSize:11,cursor:"pointer",borderRadius:4,color:T.text,userSelect:"none"}}
                onMouseEnter={e=>e.currentTarget.style.background=T.bgCard}
                onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                {f}
              </div>
            ))}
            {visAvail.length===0&&<div style={{fontSize:11,color:T.textMd,padding:"4px 6px"}}>{q?"No matches":"All fields assigned"}</div>}
          </div>
          <button onClick={()=>{setOpen(false);setQ("");}}
            style={{width:"100%",marginTop:4,padding:"3px 0",fontSize:10,background:"none",border:"none",cursor:"pointer",color:T.textMd,borderTop:"1px solid "+T.border}}>
            Close ✕
          </button>
        </div>
      )}
    </div>
  );
}

function ValuesZoneEditor({label,color,hint,values,allFields,onAdd,onRemove,onAggChange,onReorder}){
  const [open,setOpen]=useState(false);
  const [q,setQ]=useState("");
  const [dragIdx,setDragIdx]=useState(null);
  const assigned=values.map(v=>v.field);
  const available=allFields.filter(f=>!assigned.includes(f));
  const visAvail=q.trim()?available.filter(f=>f.toLowerCase().includes(q.toLowerCase())):available;
  const AGGS=["sum","count","avg","min","max"];
  return(
    <div style={{background:T.bgAlt,border:"1px solid "+T.border,borderRadius:8,padding:10,minHeight:80}}>
      <div style={{fontSize:11,fontWeight:700,color,marginBottom:3,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span>{label}</span>
        <button onClick={()=>setOpen(o=>!o)}
          style={{fontSize:10,background:color,color:"#fff",border:"none",borderRadius:4,padding:"1px 8px",cursor:"pointer"}}>+ Add</button>
      </div>
      <div style={{fontSize:10,color:T.textMd,marginBottom:6}}>{hint}</div>
      <div style={{display:"flex",flexDirection:"column",gap:5}}>
        {values.map((v,i)=>(
          <div key={v.field} draggable
            onDragStart={()=>setDragIdx(i)}
            onDragOver={e=>e.preventDefault()}
            onDrop={()=>{
              if(dragIdx===null||dragIdx===i)return;
              const n=[...values];n.splice(dragIdx,1);n.splice(i,0,values[dragIdx]);
              onReorder&&onReorder(n);setDragIdx(null);
            }}
            onDragEnd={()=>setDragIdx(null)}
            style={{display:"flex",alignItems:"center",gap:5,cursor:"grab",opacity:dragIdx===i?0.4:1}}>
            <span style={{fontSize:9,color:T.textMd,flexShrink:0}}>⠿</span>
            <span style={{flex:1,padding:"2px 8px",background:color,color:"#fff",borderRadius:10,fontSize:11,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{v.field}</span>
            <select value={v.agg} onChange={e=>onAggChange(v.field,e.target.value)}
              style={{fontSize:10,padding:"2px 4px",border:"1px solid "+T.border,borderRadius:4,background:T.bgCard,color:T.text}}>
              {AGGS.map(a=><option key={a} value={a}>{a}</option>)}
            </select>
            <button onClick={()=>onRemove(v.field)}
              style={{background:"none",border:"none",color:T.textMd,cursor:"pointer",fontSize:15,padding:0,lineHeight:1}}>×</button>
          </div>
        ))}
        {values.length===0&&<span style={{fontSize:11,color:T.textMd,fontStyle:"italic"}}>Empty — click + Add</span>}
      </div>
      {open&&(
        <div style={{marginTop:6,borderTop:"1px solid "+T.border,paddingTop:6}}>
          <input value={q} onChange={e=>setQ(e.target.value)} placeholder={`Search ${available.length} fields…`}
            style={{width:"100%",padding:"4px 8px",border:"0.5px solid "+T.border,borderRadius:5,fontSize:11,
              background:T.bgCard,color:T.text,boxSizing:"border-box",outline:"none",marginBottom:4}}/>
          <div style={{maxHeight:130,overflowY:"auto"}}>
            {visAvail.map(f=>(
              <div key={f} onClick={()=>{onAdd(f);setOpen(false);setQ("");}}
                style={{padding:"4px 6px",fontSize:11,cursor:"pointer",borderRadius:4,color:T.text,userSelect:"none"}}
                onMouseEnter={e=>e.currentTarget.style.background=T.bgCard}
                onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                {f}
              </div>
            ))}
            {visAvail.length===0&&<div style={{fontSize:11,color:T.textMd,padding:"4px 6px"}}>{q?"No matches":"All fields assigned"}</div>}
          </div>
          <button onClick={()=>{setOpen(false);setQ("");}}
            style={{width:"100%",marginTop:4,padding:"3px 0",fontSize:10,background:"none",border:"none",cursor:"pointer",color:T.textMd,borderTop:"1px solid "+T.border}}>
            Close ✕
          </button>
        </div>
      )}
    </div>
  );
}

// ── Phase 3: Collab Setup Panel (modal) ───────────────────────────────────────
// Builder configures collab columns, row key, and manages cycles
function CollabSetupPanel({report,currentUser,currentRole,onClose}) {
  const [columns,setColumns]=useState([]);
  const [cycles,setCycles]=useState([]);
  const [allUsers,setAllUsers]=useState([]);
  const [loading,setLoading]=useState(true);
  const [saving,setSaving]=useState(false);
  const [msg,setMsg]=useState("");
  const [editingCol,setEditingCol]=useState(null); // null | 'new' | {existing col}
  const [newCycleLabel,setNewCycleLabel]=useState("");
  const [cycleHistViewers,setCycleHistViewers]=useState([]);
  const [openingCycle,setOpeningCycle]=useState(false);
  const [renamingCycle,setRenamingCycle]=useState(null); // {id, label} being renamed
  const [deletingCycle,setDeletingCycle]=useState(null); // cycleId being deleted
  const [colDragIdx,setColDragIdx]=useState(null);       // drag-and-drop index for collab columns
  const [dataFields,setDataFields]=useState([]);
  const _setupCfg=typeof report?.config==="string"?JSON.parse(report.config||"{}"):(report?.config||{});
  const [viewRows,setViewRows]=useState(_setupCfg.collab_rows||((_setupCfg.collab_row_key)?[_setupCfg.collab_row_key]:[]));
  const [viewCols,setViewCols]=useState(_setupCfg.collab_cols||[]);
  const [viewValues,setViewValues]=useState(_setupCfg.collab_values||(Array.isArray(_setupCfg.collab_display_fields)?_setupCfg.collab_display_fields:[]).map(f=>({field:f,agg:"sum"})));
  const [viewFilters,setViewFilters]=useState(_setupCfg.collab_filters||[]);
  const [savingView,setSavingView]=useState(false);

  // ColForm state
  const blankForm={label:"",col_type:"input",inputter_ids:[],reviewer_ids:[],validation_config:null};
  const [colForm,setColForm]=useState(blankForm);

  useEffect(()=>{
    (async()=>{
      // Each call is independent — one failure never blocks others

      // 1. Collab columns
      try{
        const cols=await getCollabColumns(report.id);
        setColumns(cols);
      }catch(e){console.warn("collab columns load error:",e.message);}

      // 2. Cycles
      try{
        const cycs=await getCollabCycles(report.id);
        setCycles(cycs);
      }catch(e){}

      // 3. Users for inputter/reviewer assignment (silent fail — form still usable)
      try{
        const users=await getUsers();
        if(Array.isArray(users))setAllUsers(users);
      }catch(e){}

      // 4. Field names — lightweight endpoint first, then full data fallback
      try{
        const fields=await getReportFields(report.id);
        if(Array.isArray(fields)&&fields.length>0){setDataFields(fields);}
        else throw new Error("empty");
      }catch(e){
        try{
          const rd=await getReportData(report.id);
          const fields=rd.fields?.length?rd.fields
                       :rd.rows?.length?Object.keys(rd.rows[0])
                       :[];
          if(fields.length>0)setDataFields(fields);
        }catch(e2){}
      }

      setLoading(false);
    })();
  },[report.id]);

  const saveViewConfig=async()=>{
    setSavingView(true);
    try{
      const newCfg={...(report.config||{}),
        collab_rows:viewRows, collab_cols:viewCols,
        collab_values:viewValues, collab_filters:viewFilters,
        collab_row_key:viewRows[0]||null,             // backward compat
        collab_display_fields:viewValues.map(v=>v.field), // backward compat
      };
      await updateReportConfig(report.id,newCfg,report.name);
      setMsg("✓ View config saved");setTimeout(()=>setMsg(""),2000);
    }catch(e){setMsg("Error: "+e.message);}
    setSavingView(false);
  };

  const openCycle=async()=>{
    if(!newCycleLabel.trim())return setMsg("Period label required (e.g. May 2026)");
    setOpeningCycle(true);
    try{
      const c=await openCollabCycle(report.id,newCycleLabel.trim(),cycleHistViewers);
      setCycles(prev=>[c,...prev]);setNewCycleLabel("");setCycleHistViewers([]);
      setMsg("✓ Cycle opened: "+c.period_label);
    }catch(e){setMsg("Error: "+e.message);}
    setOpeningCycle(false);
  };

  const closeCycle=async(cycleId)=>{
    if(!window.confirm("Close this cycle? All values will be frozen and it will move to History."))return;
    setSaving(true);
    try{
      const c=await closeCollabCycle(report.id,cycleId);
      setCycles(prev=>prev.map(x=>x.id===c.id?c:x));
      setMsg("✓ Cycle closed");
    }catch(e){setMsg("Error: "+e.message);}
    setSaving(false);
  };

  const reopenCycle=async(cycleId)=>{
    if(!window.confirm("Reopen this cycle? It will become active again and accept new values."))return;
    setSaving(true);
    try{
      const c=await reopenCollabCycle(report.id,cycleId);
      setCycles(prev=>prev.map(x=>x.id===c.id?c:x));
      setMsg("✓ Cycle reopened");
    }catch(e){setMsg("Error: "+e.message);}
    setSaving(false);
  };

  const saveCol=async()=>{
    if(!colForm.label.trim())return setMsg("Column label required");
    setSaving(true);
    try{
      const vc=colForm.validation_config;
      const normVc=vc?.field?{field:vc.field,rule:vc.rule||"lte",...(vc.rule==="pct"?{pct_max:vc.pct_max||100}:{})}:null;
      const payload={...colForm,validation_config:normVc,ref_column:null,col_order:columns.length};
      let saved;
      if(editingCol&&editingCol.id){
        saved=await updateCollabColumn(report.id,editingCol.id,payload);
        setColumns(prev=>prev.map(c=>c.id===saved.id?saved:c));
      }else{
        saved=await createCollabColumn(report.id,payload);
        setColumns(prev=>[...prev,saved]);
      }
      setEditingCol(null);setColForm(blankForm);setMsg("✓ Column saved");
    }catch(e){setMsg("Error: "+e.message);}
    setSaving(false);
  };

  const deleteCol=async(colId)=>{
    if(!window.confirm("Delete this column? All values for this column will be lost."))return;
    try{
      await deleteCollabColumn(report.id,colId);
      setColumns(prev=>prev.filter(c=>c.id!==colId));
      setMsg("✓ Column deleted");
    }catch(e){setMsg("Error: "+e.message);}
  };

  const openEdit=(col)=>{
    setEditingCol(col);
    setColForm({
      label:col.label,col_type:col.col_type,
      inputter_ids:Array.isArray(col.inputter_ids)?col.inputter_ids:JSON.parse(col.inputter_ids||'[]'),
      reviewer_ids:Array.isArray(col.reviewer_ids)?col.reviewer_ids:JSON.parse(col.reviewer_ids||'[]'),
      validation_config:col.validation_config||(col.validation_config===null?null:null),
    });
  };

  // dataFields loaded via getReportFields in useEffect above
  const statusColor={open:"#2D6A4F",closed:"#A32D2D"};

  return(
    <div style={{position:"fixed",inset:0,zIndex:700,background:"rgba(44,24,16,0.55)",display:"flex",alignItems:"flex-start",justifyContent:"center",overflowY:"auto",padding:"20px 8px"}}>
      <div style={{background:T.bgPage,borderRadius:14,width:"min(860px,98vw)",boxShadow:"0 16px 60px rgba(44,24,16,0.4)",marginBottom:20}}>
        {/* Header */}
        <div style={{background:T.bgHeader,borderRadius:"14px 14px 0 0",padding:"16px 20px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{color:T.textLt,fontWeight:700,fontSize:15}}>⚙ Workflow Setup</div>
            <div style={{color:"rgba(245,239,230,0.7)",fontSize:12,marginTop:2}}>{report.name}</div>
          </div>
          <button onClick={onClose} style={{background:"none",border:"none",color:T.textLt,fontSize:20,cursor:"pointer",lineHeight:1}}>✕</button>
        </div>

        <div style={{padding:"20px"}}>
          {msg&&<div style={{background:"#E8F5E9",color:"#2D6A4F",border:"1px solid #A5D6A7",borderRadius:7,padding:"8px 14px",fontSize:13,marginBottom:16}}>{msg}</div>}
          {loading&&<div style={{color:T.textMd,padding:20}}>Loading…</div>}

          {/* ── View Configuration (Mini Report Builder) ── */}
          {!loading&&(
            <div style={{marginBottom:20,background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,padding:16}}>
              <div style={{fontWeight:700,color:T.primary,fontSize:14,marginBottom:4}}>📐 View Configuration</div>
              <div style={{fontSize:12,color:T.textMd,marginBottom:14,lineHeight:1.5}}>
                Configure how data appears in the Workflow View — same field zones as the Report Builder.
                The view always summarizes by <strong>Rows (R)</strong> with drill-down, shows <strong>Values (V)</strong> as reference columns,
                and exposes <strong>Filters (F)</strong> as filter pills.
              </div>

              {/* All fields overview */}
              <div style={{marginBottom:12}}>
                <div style={{fontSize:11,fontWeight:600,color:T.textMd,textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:6}}>
                  All Fields ({dataFields.length})
                </div>
                <div style={{display:"flex",flexWrap:"wrap",gap:5,padding:"8px 10px",background:T.bgAlt,borderRadius:7,border:"1px solid "+T.border,minHeight:36}}>
                  {dataFields.map(f=>{
                    const inR=viewRows.includes(f);
                    const inC=viewCols.includes(f);
                    const inV=viewValues.some(v=>v.field===f);
                    const inF=viewFilters.includes(f);
                    const zone=inR?"R":inC?"C":inV?"V":inF?"F":null;
                    const zc={R:T.tagR,C:T.tagC,V:T.tagV,F:T.tagF};
                    return(
                      <span key={f} style={{padding:"2px 9px",borderRadius:11,fontSize:11,
                        background:zone?zc[zone]:T.bgCard,color:zone?"#fff":T.text,
                        border:"1px solid "+(zone?zc[zone]:T.border)}}>
                        {zone&&<span style={{fontSize:9,fontWeight:700,marginRight:3,opacity:0.85}}>[{zone}]</span>}
                        {f}
                      </span>
                    );
                  })}
                  {dataFields.length===0&&<span style={{fontSize:12,color:T.textMd}}>⚠ No fields — upload data first</span>}
                </div>
              </div>

              {/* Zone editors */}
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
                <ZoneEditor label="🏷 Rows (R)" color={T.tagR}
                  hint="Field(s) that identify each row — e.g. Vendor Name. Drag pills to reorder."
                  fields={viewRows} allFields={dataFields}
                  onAdd={f=>{if(!viewRows.includes(f))setViewRows(r=>[...r,f]);}}
                  onRemove={f=>setViewRows(r=>r.filter(x=>x!==f))}
                  onReorder={setViewRows}/>

                <ValuesZoneEditor label="📊 Values (V)" color={T.tagV}
                  hint="Numeric reference columns — drag to reorder."
                  values={viewValues} allFields={dataFields}
                  onAdd={f=>{if(!viewValues.some(v=>v.field===f))setViewValues(v=>[...v,{field:f,agg:"sum"}]);}}
                  onRemove={f=>setViewValues(v=>v.filter(x=>x.field!==f))}
                  onAggChange={(f,agg)=>setViewValues(v=>v.map(x=>x.field===f?{...x,agg}:x))}
                  onReorder={setViewValues}/>

                <ZoneEditor label="📋 Columns (C)" color={T.tagC}
                  hint="Optional: column grouping field (cross-tab like builder)."
                  fields={viewCols} allFields={dataFields}
                  onAdd={f=>{if(!viewCols.includes(f))setViewCols(c=>[...c,f]);}}
                  onRemove={f=>setViewCols(c=>c.filter(x=>x!==f))}
                  onReorder={setViewCols}/>

                <ZoneEditor label="🔍 Filters (F)" color={T.tagF}
                  hint="Fields exposed as filter pills in the Workflow View toolbar."
                  fields={viewFilters} allFields={dataFields}
                  onAdd={f=>{if(!viewFilters.includes(f))setViewFilters(v=>[...v,f]);}}
                  onRemove={f=>setViewFilters(v=>v.filter(x=>x!==f))}
                  onReorder={setViewFilters}/>
              </div>

              <button onClick={saveViewConfig} disabled={savingView}
                style={{padding:"8px 18px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,cursor:"pointer",fontSize:13,fontWeight:700,opacity:savingView?0.6:1}}>
                {savingView?"Saving…":"Save View Config"}
              </button>
            </div>
          )}

          {/* ── Collab Columns ── */}
          {!loading&&(
            <div style={{marginBottom:24}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                <div style={{fontWeight:700,color:T.primary,fontSize:14}}>Collaboration Columns</div>
                {!editingCol&&<button onClick={()=>{setEditingCol('new');setColForm(blankForm);}}
                  style={{padding:"6px 14px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600}}>
                  + Add Column
                </button>}
              </div>
              {/* Column form */}
              {editingCol&&(
                <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,padding:16,marginBottom:14}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:12}}>
                    <div>
                      <label style={{fontSize:12,fontWeight:600,color:T.textMd}}>Column Label *</label>
                      <input value={colForm.label} onChange={e=>setColForm(f=>({...f,label:e.target.value}))}
                        placeholder="e.g. Recommended Amount"
                        style={{width:"100%",padding:"7px 10px",border:"1px solid "+T.border,borderRadius:6,fontSize:13,boxSizing:"border-box",marginTop:4}}/>
                    </div>
                    <div>
                      <label style={{fontSize:12,fontWeight:600,color:T.textMd}}>Type</label>
                      <select value={colForm.col_type} onChange={e=>setColForm(f=>({...f,col_type:e.target.value}))}
                        style={{width:"100%",padding:"7px 10px",border:"1px solid "+T.border,borderRadius:6,fontSize:13,marginTop:4}}>
                        <option value="input">Input Only (no approval)</option>
                        <option value="workflow">Workflow (submit → approve/reject/hold)</option>
                      </select>
                    </div>
                  </div>
                  <UserMultiSelect label="Inputters (who can enter values)" value={colForm.inputter_ids}
                    allUsers={allUsers} onChange={v=>setColForm(f=>({...f,inputter_ids:v}))}/>
                  {colForm.col_type==="workflow"&&(
                    <div style={{marginTop:10}}>
                      <UserMultiSelect label="Reviewers (who can approve/reject/hold)" value={colForm.reviewer_ids}
                        allUsers={allUsers} onChange={v=>setColForm(f=>({...f,reviewer_ids:v}))}/>
                    </div>
                  )}
                  {/* ── Validation Rule ── */}
                  <div style={{marginTop:14,padding:"12px 14px",background:T.bgAlt,borderRadius:8,border:"1px solid "+T.border}}>
                    <div style={{fontSize:12,fontWeight:700,color:T.textMd,marginBottom:8}}>🔒 Input Validation (optional)</div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,alignItems:"end"}}>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:T.textMd}}>Reference Field (from report data)</label>
                        <select value={colForm.validation_config?.field||""}
                          onChange={e=>setColForm(f=>({...f,validation_config:e.target.value?{field:e.target.value,rule:f.validation_config?.rule||"lte",...(f.validation_config?.pct_max?{pct_max:f.validation_config.pct_max}:{})}:null}))}
                          style={{width:"100%",padding:"6px 9px",border:"1px solid "+T.border,borderRadius:6,fontSize:12,marginTop:3}}>
                          <option value="">— No validation —</option>
                          {dataFields.map(f=><option key={f} value={f}>{f}</option>)}
                        </select>
                      </div>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:T.textMd}}>Rule</label>
                        <select value={colForm.validation_config?.rule||"lte"}
                          onChange={e=>setColForm(f=>({...f,validation_config:f.validation_config?{...f.validation_config,rule:e.target.value}:null}))}
                          disabled={!colForm.validation_config?.field}
                          style={{width:"100%",padding:"6px 9px",border:"1px solid "+T.border,borderRadius:6,fontSize:12,marginTop:3}}>
                          <option value="lte">≤ Cannot exceed (max)</option>
                          <option value="gte">≥ Cannot be less than (min)</option>
                          <option value="pct">% Percentage of field</option>
                        </select>
                      </div>
                    </div>
                    {colForm.validation_config?.field&&colForm.validation_config?.rule==="pct"&&(
                      <div style={{marginTop:8}}>
                        <label style={{fontSize:11,fontWeight:600,color:T.textMd}}>Max % allowed</label>
                        <input type="number" value={colForm.validation_config?.pct_max||100}
                          onChange={e=>setColForm(f=>({...f,validation_config:{...f.validation_config,pct_max:parseFloat(e.target.value)||100}}))}
                          style={{padding:"5px 8px",border:"1px solid "+T.border,borderRadius:5,fontSize:12,width:80,marginLeft:8}}/>
                        <span style={{fontSize:11,color:T.textMd,marginLeft:4}}>% of the reference field value</span>
                      </div>
                    )}
                    {colForm.validation_config?.field&&(
                      <div style={{fontSize:11,color:T.primary,marginTop:6}}>
                        Preview: Input value must be
                        {colForm.validation_config.rule==="lte"&&<strong> ≤ {colForm.validation_config.field}</strong>}
                        {colForm.validation_config.rule==="gte"&&<strong> ≥ {colForm.validation_config.field}</strong>}
                        {colForm.validation_config.rule==="pct"&&<strong> ≤ {colForm.validation_config.pct_max||100}% of {colForm.validation_config.field}</strong>}
                      </div>
                    )}
                  </div>
                  <div style={{display:"flex",gap:8,marginTop:14,justifyContent:"flex-end"}}>
                    <button onClick={()=>{setEditingCol(null);setColForm(blankForm);}}
                      style={{padding:"7px 14px",background:"none",border:"1px solid "+T.border,borderRadius:7,cursor:"pointer",fontSize:12}}>Cancel</button>
                    <button onClick={saveCol} disabled={saving}
                      style={{padding:"7px 16px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,cursor:saving?"wait":"pointer",fontSize:12,fontWeight:700,opacity:saving?0.6:1}}>
                      {saving?"Saving…":"Save Column"}
                    </button>
                  </div>
                </div>
              )}
              {/* Column list — draggable to reorder */}
              {columns.length===0&&!editingCol&&<div style={{color:T.textMd,fontSize:13,padding:"10px 0"}}>No collab columns defined yet.</div>}
              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                {columns.map((col,ci)=>(
                  <div key={col.id} draggable
                    onDragStart={()=>setColDragIdx(ci)}
                    onDragOver={e=>e.preventDefault()}
                    onDrop={async()=>{
                      if(colDragIdx===null||colDragIdx===ci)return;
                      const reordered=[...columns];
                      reordered.splice(colDragIdx,1);reordered.splice(ci,0,columns[colDragIdx]);
                      setColumns(reordered);setColDragIdx(null);
                      // Persist new col_order
                      try{await Promise.all(reordered.map((c,idx)=>updateCollabColumn(report.id,c.id,{...c,col_order:idx})));}
                      catch(e){setMsg("Error saving order: "+e.message);}
                    }}
                    onDragEnd={()=>setColDragIdx(null)}
                    style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:8,padding:"10px 14px",
                      display:"flex",alignItems:"center",gap:10,cursor:"grab",
                      opacity:colDragIdx===ci?0.4:1,
                      boxShadow:colDragIdx===ci?"inset 0 0 0 2px "+T.accent:"none"}}>
                    <span style={{fontSize:14,color:T.textMd,flexShrink:0,cursor:"grab"}}>⠿</span>
                    <div style={{flex:1}}>
                      <span style={{fontWeight:600,fontSize:13,color:T.text}}>{col.label}</span>
                      <span style={{marginLeft:8,fontSize:11,background:col.col_type==="workflow"?"#534AB7":"#185FA5",color:"#fff",padding:"2px 7px",borderRadius:10}}>
                        {col.col_type==="workflow"?"Workflow":"Input Only"}
                      </span>
                    </div>
                    <button onClick={()=>openEdit(col)} style={{padding:"4px 10px",background:"none",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:11}}>Edit</button>
                    <button onClick={()=>deleteCol(col.id)} style={{padding:"4px 10px",background:"none",border:"1px solid #A32D2D",color:"#A32D2D",borderRadius:6,cursor:"pointer",fontSize:11}}>Delete</button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ── Cycle Management ── */}
          {!loading&&(
            <div>
              <div style={{fontWeight:700,color:T.primary,fontSize:14,marginBottom:12}}>Monthly Cycles</div>
              {/* Open new cycle */}
              <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,padding:14,marginBottom:14}}>
                  <div style={{fontSize:13,fontWeight:600,color:T.text,marginBottom:10}}>Open New Cycle</div>
                  <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end"}}>
                    <div style={{flex:1,minWidth:160}}>
                      <label style={{fontSize:12,fontWeight:600,color:T.textMd}}>Period Label *</label>
                      <input value={newCycleLabel} onChange={e=>setNewCycleLabel(e.target.value)}
                        placeholder="e.g. May 2026"
                        style={{width:"100%",padding:"7px 10px",border:"1px solid "+T.border,borderRadius:6,fontSize:13,boxSizing:"border-box",marginTop:4}}/>
                    </div>
                    <div style={{flex:2,minWidth:200}}>
                      <UserMultiSelect label="History viewers (can see closed cycle)" value={cycleHistViewers}
                        allUsers={allUsers} onChange={setCycleHistViewers}/>
                    </div>
                    <button onClick={openCycle} disabled={openingCycle}
                      style={{padding:"8px 18px",background:T.accent,color:T.textLt,border:"none",borderRadius:7,cursor:openingCycle?"wait":"pointer",fontSize:13,fontWeight:700,height:36,marginBottom:0,opacity:openingCycle?0.6:1}}>
                      {openingCycle?"Opening…":"Open Cycle"}
                    </button>
                  </div>
                </div>
              {/* Cycle list */}
              {cycles.length===0&&<div style={{color:T.textMd,fontSize:13,padding:"8px 0"}}>No cycles yet.</div>}
              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                {cycles.map(c=>{
                  const isRenaming=renamingCycle?.id===c.id;
                  return(
                  <div key={c.id} style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:8,padding:"10px 16px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                    <div style={{width:10,height:10,borderRadius:"50%",background:statusColor[c.status]||T.textMd,flexShrink:0}}/>
                    <div style={{flex:1,minWidth:160}}>
                      {isRenaming?(
                        <div style={{display:"flex",gap:6,alignItems:"center"}}>
                          <input value={renamingCycle.label} autoFocus
                            onChange={e=>setRenamingCycle(r=>({...r,label:e.target.value}))}
                            onKeyDown={async e=>{
                              if(e.key==="Enter"){
                                try{const updated=await renameCollabCycle(report.id,c.id,renamingCycle.label);setCycles(prev=>prev.map(x=>x.id===updated.id?updated:x));setRenamingCycle(null);}catch(err){setMsg("Error: "+err.message);}
                              } else if(e.key==="Escape") setRenamingCycle(null);
                            }}
                            style={{padding:"4px 8px",border:"1px solid "+T.accent,borderRadius:5,fontSize:13,fontWeight:600,color:T.text,background:T.bgCard,outline:"none",flex:1}}/>
                          <button onClick={async()=>{
                            try{const updated=await renameCollabCycle(report.id,c.id,renamingCycle.label);setCycles(prev=>prev.map(x=>x.id===updated.id?updated:x));setRenamingCycle(null);}catch(err){setMsg("Error: "+err.message);}
                          }} style={{padding:"4px 10px",background:T.accent,color:T.textLt,border:"none",borderRadius:5,cursor:"pointer",fontSize:12,fontWeight:600}}>Save</button>
                          <button onClick={()=>setRenamingCycle(null)} style={{padding:"4px 10px",background:"none",border:"1px solid "+T.border,borderRadius:5,cursor:"pointer",fontSize:12}}>Cancel</button>
                        </div>
                      ):(
                        <>
                          <span style={{fontWeight:600,fontSize:13,color:T.text}}>{c.period_label}</span>
                          <span style={{marginLeft:8,fontSize:11,color:statusColor[c.status],fontWeight:600}}>{c.status.toUpperCase()}</span>
                          {c.closed_at&&<span style={{marginLeft:8,fontSize:11,color:T.textMd}}>Closed {new Date(c.closed_at).toLocaleDateString()}</span>}
                        </>
                      )}
                    </div>
                    {!isRenaming&&(
                      <button onClick={()=>setRenamingCycle({id:c.id,label:c.period_label})}
                        style={{padding:"4px 10px",background:"none",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:11,color:T.textMd}}>
                        ✏ Rename
                      </button>
                    )}
                    {c.status==="open"&&!isRenaming&&(
                      <button onClick={()=>closeCycle(c.id)} disabled={saving}
                        style={{padding:"5px 12px",background:"#A32D2D",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600}}>
                        Close Cycle
                      </button>
                    )}
                    {c.status==="closed"&&!isRenaming&&(
                      <button onClick={()=>reopenCycle(c.id)} disabled={saving}
                        style={{padding:"5px 12px",background:T.success,color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600}}>
                        ↩ Reopen
                      </button>
                    )}
                    {/* Super Admin only: delete a closed cycle */}
                    {currentRole==="admin"&&c.status==="closed"&&!isRenaming&&(
                      <button onClick={async()=>{
                        if(!window.confirm(`Delete cycle "${c.period_label}"? All entered values will be permanently lost.`))return;
                        setDeletingCycle(c.id);
                        try{await deleteCollabCycle(report.id,c.id);setCycles(prev=>prev.filter(x=>x.id!==c.id));setMsg("✓ Cycle deleted");}
                        catch(e){setMsg("Error: "+e.message);}
                        setDeletingCycle(null);
                      }} disabled={deletingCycle===c.id}
                        style={{padding:"5px 12px",background:"none",border:"1px solid #A32D2D",color:"#A32D2D",borderRadius:7,cursor:"pointer",fontSize:12,opacity:deletingCycle===c.id?0.5:1}}>
                        🗑 Delete
                      </button>
                    )}
                  </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ── Helper: multi-select user picker ──────────────────────────────────────────
function UserMultiSelect({label,value=[],allUsers=[],onChange}) {
  const safeValue=Array.isArray(value)?value:[];
  const toggleUser=(id)=>{
    if(safeValue.includes(id)) onChange(safeValue.filter(x=>x!==id));
    else onChange([...safeValue,id]);
  };
  const eligible=Array.isArray(allUsers)?allUsers.filter(u=>u.role!=='admin'):[];
  return(
    <div>
      <label style={{fontSize:12,fontWeight:600,color:T.textMd}}>{label}</label>
      <div style={{display:"flex",flexWrap:"wrap",gap:6,marginTop:4}}>
        {eligible.map(u=>(
          <button key={u.id} onClick={()=>toggleUser(u.id)}
            style={{padding:"4px 10px",borderRadius:14,fontSize:11,cursor:"pointer",fontWeight:safeValue.includes(u.id)?700:400,
              background:safeValue.includes(u.id)?T.primary:"none",
              color:safeValue.includes(u.id)?T.textLt:T.text,
              border:"1px solid "+(safeValue.includes(u.id)?T.primary:T.border)}}>
            {u.username} ({u.role})
          </button>
        ))}
        {eligible.length===0&&<span style={{color:T.textMd,fontSize:12}}>No users available</span>}
      </div>
    </div>
  );
}

// ── Phase 3: Collab Data View (modal) ────────────────────────────────────────
// Shows flat table of data rows + collab columns; inputters can enter values, reviewers can approve
function CollabDataView({report,currentUser,currentRole,onClose}) {
  const [columns,setColumns]=useState([]);
  const [cycles,setCycles]=useState([]);
  const [activeCycle,setActiveCycle]=useState(null);
  const [values,setValues]=useState({});
  const [dataRows,setDataRows]=useState([]);
  const [loading,setLoading]=useState(true);
  const [saving,setSaving]=useState({});
  const [draftMap,setDraftMap]=useState({});
  const [auditRow,setAuditRow]=useState(null);
  const [auditData,setAuditData]=useState([]);
  const [reviewModal,setReviewModal]=useState(null);
  const [reviewRemarks,setReviewRemarks]=useState("");
  const [reviewValue,setReviewValue]=useState(""); // approver's editable value
  const [msg,setMsg]=useState("");
  const [histMode,setHistMode]=useState(false);
  // Filter/sort/display state
  const [search,setSearch]=useState("");
  const [sortCol,setSortCol]=useState(null);
  const [showNonZeroOnly,setShowNonZeroOnly]=useState(false);
  const [numFmt,setNumFmt]=useState("units");
  const [page,setPage]=useState(0);
  const PAGE_SIZE=100;
  const [colFilters,setColFilters]=useState({});      // {field: string[]|undefined} — pre-grouping
  const [collabFilters,setCollabFilters]=useState({}); // {colId: string[]|undefined} — post-grouping on collab values
  const [colSorts,setColSorts]=useState({});           // {field: 'az'|'za'|'09'|'90'}
  const [expanded,setExpanded]=useState(new Set());   // level-1 group keys that are expanded (hierarchical mode)
  const [drillDown,setDrillDown]=useState(null);      // {rowKey,colVal,rFs,cF,metricLabel} — raw record drill-down panel
  const [cellErrors,setCellErrors]=useState({});      // {dk: errorMessage} — per-cell validation popup
  // Fresh config from API (overrides stale prop after Save View Config)
  const [freshConfig,setFreshConfig]=useState(null);

  // Determine user's roles in each column
  const myId=currentUser?.id||getMyUserId(); // fallback: decode from JWT
  const isBuilder=['admin','subadmin','subadmin_user'].includes(currentRole);

  // Derive view config — prefer freshly-loaded config to avoid stale props
  const _ac=freshConfig||(typeof report?.config==="string"?JSON.parse(report.config||"{}"):(report?.config||{}));
  const viewRows=_ac.collab_rows||((_ac.collab_row_key)?[_ac.collab_row_key]:[]);
  const viewValues=_ac.collab_values||(Array.isArray(_ac.collab_display_fields)?_ac.collab_display_fields:[]).map(f=>({field:f,agg:"sum"}));
  const viewCols=_ac.collab_cols||[]; // C zone — cross-tab column field (first one used)
  const cField=viewCols[0]||null;    // primary C field for cross-tab
  const viewFilters=_ac.collab_filters||[];
  const rowKey=viewRows[0]||null; // primary display field
  const computeRK=(row,ri)=>viewRows.length===0?String(ri):viewRows.map(f=>String(row[f]||"")).join("||");
  const numFieldsForFilter=useMemo(()=>new Set(viewValues.map(v=>v.field)),[viewValues]);
  // All raw field names + numeric field detection for the DrillDown component
  const allDataFields=useMemo(()=>dataRows.length?Object.keys(dataRows[0]).filter(k=>!k.startsWith("__")):[],[dataRows]);
  const allNumFields=useMemo(()=>{
    if(!dataRows.length)return new Set();
    const first=dataRows[0];
    return new Set(Object.keys(first).filter(k=>{const v=first[k];return typeof v==="number"||(v!==null&&v!==undefined&&v!==""&&!isNaN(Number(v)));}));
  },[dataRows]);
  // Unique C field values for cross-tab column headers
  const cVals=useMemo(()=>{
    if(!cField)return[];
    const seen=new Map();
    dataRows.forEach(r=>{const raw=String(r[cField]||"").trim();const k=raw.toLowerCase();if(!seen.has(k))seen.set(k,raw);});
    return [...seen.values()].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
  },[cField,dataRows]);

  useEffect(()=>{loadAll();},[report.id]);

  const loadAll=async()=>{
    setLoading(true);
    try{
      const [cols,cycs,rd]=await Promise.all([
        getCollabColumns(report.id),
        getCollabCycles(report.id),
        getReportData(report.id)
      ]);
      setColumns(cols);
      setCycles(cycs);
      setDataRows(rd.rows||[]);
      // Fetch fresh report config to get latest row key / display fields
      try{
        const fields=await getReportFields(report.id); // re-use existing endpoint
        // Also get fresh config from reports list
        const reports=await getReports();
        const fresh=reports.find(r=>r.id===report.id);
        if(fresh){
          const cfg=typeof fresh.config==="string"?JSON.parse(fresh.config||"{}"):(fresh.config||{});
          setFreshConfig(cfg);
        }
      }catch(e2){/* silent — fall back to prop */}
      const open=cycs.find(c=>c.status==="open");
      if(open){ setActiveCycle(open); await loadValues(open.id); }
    }catch(e){setMsg("Error: "+e.message);}
    setLoading(false);
  };

  const loadValues=async(cycleId)=>{
    const vals=await getCollabValues(report.id,cycleId);
    const map={};
    vals.forEach(v=>{ map[v.row_key+"__"+v.col_id]=v; });
    setValues(map);
  };

  const switchCycle=async(c)=>{
    setActiveCycle(c);
    setValues({});
    setDraftMap({});
    if(c){await loadValues(c.id);}
  };

  const draftKey=(rk,cid)=>rk+"__"+cid;
  const autoSaveTimers=useRef({}); // debounce handles per cell

  // Submit all pending workflow values across all rows (not just current page)
  const submitAll=async()=>{
    if(!activeCycle||activeCycle.status==="closed")return;
    const pending=Object.values(values).filter(v=>
      v.status==="pending"&&v.value!==null&&v.value!==undefined&&v.value!==""
    ).map(v=>{
      const col=columns.find(c=>c.id===v.col_id||c.id===parseInt(v.col_id));
      return col&&col.col_type==="workflow"&&canInput(col)?{rk:v.row_key,col}:null;
    }).filter(Boolean);
    if(!pending.length){setMsg("No pending values to submit");setTimeout(()=>setMsg(""),2500);return;}
    try{
      await Promise.all(pending.map(({rk,col})=>submitCollabValue(report.id,activeCycle.id,rk,col.id)));
      await loadValues(activeCycle.id);
      setMsg(`✓ Submitted ${pending.length} value${pending.length!==1?"s":""} for review`);
      setTimeout(()=>setMsg(""),3000);
    }catch(e){setMsg("Error: "+e.message);}
  };

  const canInput=(col)=>{
    if(isBuilder)return true;
    const ids=Array.isArray(col.inputter_ids)?col.inputter_ids:JSON.parse(col.inputter_ids||'[]');
    if(ids.length===0)return true; // no restriction — all authenticated users can input
    return ids.includes(myId);
  };
  const canReview=(col)=>{
    if(['admin','subadmin'].includes(currentRole))return true;
    const ids=Array.isArray(col.reviewer_ids)?col.reviewer_ids:JSON.parse(col.reviewer_ids||'[]');
    return ids.includes(myId);
  };

  // Validate input value against the column's validation_config
  const validateInput=(rk,col,val)=>{
    const vc=col.validation_config;
    if(!vc||!vc.field||!vc.rule)return null;
    // Sum the reference field across ALL rows matching this row key (vendor summary level, not individual bill)
    const matchingRows=dataRows.filter(r=>viewRows.map(f=>String(r[f]||"")).join("||")===rk||String(r[viewRows[0]]||"")===rk);
    if(!matchingRows.length)return null;
    const refVal=matchingRows.reduce((sum,r)=>sum+(parseFloat(r[vc.field])||0),0);
    const inputVal=parseFloat(val)||0;
    if(vc.rule==="lte"&&inputVal>refVal)
      return `${col.label}: value (${inputVal.toLocaleString()}) cannot exceed ${vc.field} (${refVal.toLocaleString()})`;
    if(vc.rule==="gte"&&inputVal<refVal)
      return `${col.label}: value (${inputVal.toLocaleString()}) cannot be less than ${vc.field} (${refVal.toLocaleString()})`;
    if(vc.rule==="pct"){const max=(vc.pct_max||100)/100*refVal;if(inputVal>max)return `${col.label}: value exceeds ${vc.pct_max||100}% of ${vc.field} (max: ${max.toLocaleString()})`;}
    return null;
  };

  const saveDraft=async(rk,col)=>{
    const dk=draftKey(rk,col.id);
    const draft=draftMap[dk]||{};
    const val=draft.value!==undefined?draft.value:null;
    if(val===null||val==="")return;
    // If this draft was set by a validation revert (not user input), skip saving it
    if(draft.isReverted)return;
    const validationErr=validateInput(rk,col,val);
    if(validationErr){
      if(autoSaveTimers.current[dk])clearTimeout(autoSaveTimers.current[dk]);
      setCellErrors(e=>({...e,[dk]:validationErr}));
      // Mark draft as reverted-to-zero so blur/auto-save won't save it
      setDraftMap(d=>({...d,[dk]:{value:"0",isReverted:true}}));
      return;
    }
    setCellErrors(e=>{const n={...e};delete n[dk];return n;});
    setSaving(s=>({...s,[dk]:true}));
    try{
      const saved=await upsertCollabValue(report.id,activeCycle.id,{row_key:rk,col_id:col.id,value:parseFloat(val)||0,remarks:null});
      setValues(v=>({...v,[dk]:saved}));
      // Only clear the draft if the user has NOT typed more digits since this save started.
      // If they have (race: typed while API was in-flight), keep the newer draft so it gets saved by its own timer.
      setDraftMap(d=>{
        const cur=d[dk];
        if(!cur||String(cur.value)===String(val)){const nd={...d};delete nd[dk];return nd;}
        return d;
      });
    }catch(e){setMsg("Error: "+e.message);}
    setSaving(s=>({...s,[dk]:false}));
  };

  const submitValue=async(rk,col)=>{
    const dk=draftKey(rk,col.id);
    setSaving(s=>({...s,[dk]:true}));
    try{
      const saved=await submitCollabValue(report.id,activeCycle.id,rk,col.id);
      setValues(v=>({...v,[dk]:saved}));
      setMsg("✓ Submitted for review");setTimeout(()=>setMsg(""),2000);
    }catch(e){setMsg("Error: "+e.message);}
    setSaving(s=>({...s,[dk]:false}));
  };

  const doReview=async(action)=>{
    if(!reviewModal)return;
    const {rowKey:rk,col_id}=reviewModal;
    const dk=draftKey(rk,col_id);
    setSaving(s=>({...s,[dk]:true}));
    try{
      // Pass modified_value only when action is 'modified'
      const modVal=action==="modified"?parseFloat(reviewValue)||0:undefined;
      const saved=await reviewCollabValue(report.id,activeCycle.id,rk,col_id,action,reviewRemarks||null,modVal);
      setValues(v=>({...v,[dk]:saved}));
      setReviewModal(null);setReviewRemarks("");setReviewValue("");
      setMsg("✓ "+action.charAt(0).toUpperCase()+action.slice(1));setTimeout(()=>setMsg(""),2000);
    }catch(e){setMsg("Error: "+e.message);}
    setSaving(s=>({...s,[dk]:false}));
  };

  const openAudit=async(rk)=>{
    setAuditRow(rk);
    if(activeCycle){
      const a=await getCollabAudit(report.id,activeCycle.id,rk);
      setAuditData(a);
    }
  };

  const statusBadge=(status)=>{
    const map={
      pending:  {bg:"#FFF3CD",color:"#856404"},
      submitted:{bg:"#CCE5FF",color:"#004085"},
      approved: {bg:"#D4EDDA",color:"#155724"},
      rejected: {bg:"#F8D7DA",color:"#721C24"},
      hold:     {bg:"#E2D9F3",color:"#4A2080"},
      modified: {bg:"#FFE5B4",color:"#7A3E00"}, // amber — reviewer changed the value
    };
    const s=map[status]||{bg:T.bgAlt,color:T.textMd};
    return<span style={{fontSize:10,fontWeight:700,background:s.bg,color:s.color,padding:"2px 6px",borderRadius:8,textTransform:"uppercase"}}>{status||"—"}</span>;
  };

  // Number formatter for display columns
  const fmtNum=(v)=>{
    if(v===null||v===undefined||v==="")return "—";
    const n=Number(v);
    if(isNaN(n))return String(v);
    if(numFmt==="Cr") return (n/1e7).toLocaleString('en-IN',{maximumFractionDigits:2})+" Cr";
    if(numFmt==="L")  return (n/1e5).toLocaleString('en-IN',{maximumFractionDigits:2})+" L";
    return n.toLocaleString('en-IN',{maximumFractionDigits:2});
  };

  // Sort toggle
  const toggleSort=(field)=>{
    setSortCol(s=>{
      const next=s&&s.field===field?{field,dir:s.dir==="asc"?"desc":"asc"}:{field,dir:"asc"};
      return next;
    });
    setPage(0);
  };

  // Total Approval = sum of approved + modified amounts across all workflow cols for this row
  const totalApprovalBadge=(wfCols,rk)=>{
    if(!wfCols.length)return null;
    const entries=wfCols.map(col=>values[rk+"__"+col.id]);
    const approvedSum=entries.reduce((sum,v)=>{
      if(v&&['approved','modified'].includes(v.status)){
        const amt=v.reviewer_value!=null?parseFloat(v.reviewer_value)||0:parseFloat(v.value)||0;
        return sum+amt;
      }
      return sum;
    },0);
    const allDone=entries.every(v=>v&&['approved','modified','rejected'].includes(v.status));
    const hasRejected=entries.some(v=>v?.status==='rejected');
    const hasPending=entries.some(v=>!v||v.status==='pending');
    const hasSubmitted=entries.some(v=>v?.status==='submitted');
    const bg=allDone?(hasRejected?"#F8D7DA":"#D4EDDA"):hasSubmitted?"#CCE5FF":"#FFF3CD";
    const clr=allDone?(hasRejected?"#721C24":"#155724"):hasSubmitted?"#004085":"#856404";
    const label=fmtNum(approvedSum);
    return<span style={{fontSize:12,fontWeight:700,background:bg,color:clr,padding:"3px 10px",borderRadius:8,whiteSpace:"nowrap"}}>{label}</span>;
  };

  // Filter chain
  let filteredRows=dataRows;
  if(search.trim()){
    const q=search.trim().toLowerCase();
    const flds=[...viewRows,...viewValues.map(v=>v.field)].filter(Boolean);
    filteredRows=filteredRows.filter(r=>flds.some(f=>String(r[f]||"").toLowerCase().includes(q)));
  }
  const activeColFilters=Object.entries(colFilters).filter(([,v])=>Array.isArray(v)&&v.length>0);
  if(activeColFilters.length){
    filteredRows=filteredRows.filter(row=>activeColFilters.every(([f,vals])=>vals.includes(String(row[f]??""))));
  }
  if(showNonZeroOnly){
    filteredRows=filteredRows.filter((row,ri)=>{
      const rk=computeRK(row,ri);
      return columns.some(col=>{const v=values[rk+"__"+col.id];return v&&Number(v.value)!==0;});
    });
  }
  // Sort
  if(sortCol){
    filteredRows=[...filteredRows].sort((a,b)=>{
      const av=a[sortCol.field],bv=b[sortCol.field];
      const an=Number(av),bn=Number(bv);
      const cmp=!isNaN(an)&&!isNaN(bn)?an-bn:String(av??"").localeCompare(String(bv??""),undefined,{numeric:true});
      return sortCol.dir==="asc"?cmp:-cmp;
    });
  }
  // Hierarchical pivot grouping — like Excel pivot table
  // viewRows[0] = outer level (e.g. Project), viewRows[1+] = inner level (e.g. Vendor Name)
  const aggGroup=(rows)=>{
    const g={__rows:rows,__count:rows.length,__cGroups:{}};
    viewValues.forEach(({field,agg})=>{ g[field]=doAgg(rows,field,agg); });
    if(cField){
      cVals.forEach(cv=>{
        const cr=rows.filter(r=>String(r[cField]||"").trim().toLowerCase()===cv.toLowerCase());
        g.__cGroups[cv]={};
        viewValues.forEach(({field,agg})=>{ g.__cGroups[cv][field]=cr.length>0?doAgg(cr,field,agg):null; });
      });
    }
    return g;
  };
  let displayRows=filteredRows;
  const isHierarchical=viewRows.length>=2;
  if(viewRows.length===1){
    // Single-level: each unique value of viewRows[0] is one row (case-insensitive grouping)
    const groups={};const order=[];
    filteredRows.forEach((row,ri)=>{
      const rawK=String(row[viewRows[0]]||"").trim();
      const k=rawK.toLowerCase();
      if(!groups[k]){groups[k]={...row,__rk:rawK,...aggGroup([row]),__level:1};order.push(k);}
      else{groups[k].__rows.push(row);groups[k].__count++;}
    });
    order.forEach(k=>{ Object.assign(groups[k],aggGroup(groups[k].__rows)); });
    displayRows=order.map(k=>groups[k]);
  } else if(viewRows.length>=2){
    // Two-level: viewRows[0] = outer, viewRows[1+] = inner (case-insensitive grouping)
    const L1={};const L1Order=[];
    filteredRows.forEach((row)=>{
      const rawK1=String(row[viewRows[0]]||"").trim();
      const k1=rawK1.toLowerCase();
      const rawInner=viewRows.slice(1).map(f=>String(row[f]||"").trim()).join("||");
      const innerKey=rawInner.toLowerCase();
      const rk2=rawK1+"||"+rawInner;
      if(!L1[k1]){
        L1[k1]={...row,__rk:rawK1,__rows:[],__count:0,__level:1,__cGroups:{},__inner:{},__innerOrder:[]};
        L1Order.push(k1);
      }
      L1[k1].__rows.push(row);L1[k1].__count++;
      if(!L1[k1].__inner[innerKey]){
        L1[k1].__inner[innerKey]={...row,__rk:rk2,__rows:[],__count:0,__level:2,__cGroups:{}};
        L1[k1].__innerOrder.push(innerKey);
      }
      L1[k1].__inner[innerKey].__rows.push(row);
      L1[k1].__inner[innerKey].__count++;
    });
    L1Order.forEach(k1=>{
      Object.assign(L1[k1],aggGroup(L1[k1].__rows));
      // Preserve __inner and __innerOrder which aggGroup doesn't touch
      L1[k1].__innerOrder.forEach(k2=>{
        Object.assign(L1[k1].__inner[k2],aggGroup(L1[k1].__inner[k2].__rows));
      });
    });
    displayRows=L1Order.map(k1=>L1[k1]);
  }
  // Post-grouping filter: collab column value filters
  const activeCollabFilters=Object.entries(collabFilters).filter(([,v])=>Array.isArray(v)&&v.length>0);
  if(activeCollabFilters.length){
    displayRows=displayRows.filter((row,ri)=>{
      const rk=row.__rk||computeRK(row,ri);
      return activeCollabFilters.every(([colId,vals])=>{
        const v=values[rk+"__"+colId];
        const cell=v?.value!==undefined&&v.value!==null?String(v.value):"";
        return vals.includes(cell);
      });
    });
  }
  const totalPages=Math.ceil(displayRows.length/PAGE_SIZE);
  const pagedRows=displayRows.slice(page*PAGE_SIZE,(page+1)*PAGE_SIZE);

  const cycleLabel=c=>c.period_label+(c.status==="closed"?" (Closed)":"");

  // ── Excel download for workflow view ──
  const downloadWorkflowExcel=async()=>{
    if(!activeCycle)return;
    try{
      setMsg("Preparing download…");
      const exp=await exportCollabCycle(report.id,activeCycle.id);
      const {columns:cols,values:vals,dataRows:dRows}=exp;
      // Build value map
      const valMap={};vals.forEach(v=>{valMap[v.row_key+"__"+v.col_id]=v;});
      // Build header row
      const wfColsExp=cols.filter(c=>c.col_type==="workflow");
      const header=[...viewRows,...viewValues.map(v=>v.field),...cols.map(c=>c.label),...wfColsExp.map(c=>"Approved — "+c.label),...wfColsExp.map(c=>"Status — "+c.label),"Total Approval"];
      // Build data rows — one per unique row key
      const rkMap={};
      dRows.forEach(row=>{
        const rk=viewRows.map(f=>String(row[f]||"")).join("||");
        if(!rkMap[rk])rkMap[rk]={row,rk};
      });
      const effectiveAmt=(v)=>{
        if(!v||!['approved','modified','rejected'].includes(v.status))return "";
        if(v.status==='rejected')return 0;
        return v.reviewer_value!=null?parseFloat(v.reviewer_value)||0:parseFloat(v.value)||0;
      };
      const sheetRows=[header,...Object.values(rkMap).map(({row,rk})=>{
        const wfCols=cols.filter(c=>c.col_type==="workflow");
        const approvedSum=wfCols.reduce((s,col)=>{
          const v=valMap[rk+"__"+col.id];
          if(!v||!['approved','modified'].includes(v.status))return s;
          return s+(v.reviewer_value!=null?parseFloat(v.reviewer_value)||0:parseFloat(v.value)||0);
        },0);
        return[
          ...viewRows.map(f=>row[f]||""),
          ...viewValues.map(({field})=>row[field]||0),
          ...cols.map(col=>{const v=valMap[rk+"__"+col.id];return v?.value??"";}),
          ...wfCols.map(col=>effectiveAmt(valMap[rk+"__"+col.id])),
          ...wfCols.map(col=>{const v=valMap[rk+"__"+col.id];return v?.status||"";}),
          approvedSum,
        ];
      })];
      // Use XLSX if available
      if(window.XLSX){
        const ws=window.XLSX.utils.aoa_to_sheet(sheetRows);
        const wb=window.XLSX.utils.book_new();
        window.XLSX.utils.book_append_sheet(wb,ws,"Workflow");
        window.XLSX.writeFile(wb,`${report.name}_${activeCycle.period_label}_workflow.xlsx`);
      } else {
        // Fallback CSV
        const csv=sheetRows.map(r=>r.map(c=>String(c).includes(",")?`"${c}"`:c).join(",")).join("\n");
        const a=document.createElement("a");a.href="data:text/csv;charset=utf-8,"+encodeURIComponent(csv);
        a.download=`${report.name}_${activeCycle.period_label}_workflow.csv`;a.click();
      }
      setMsg("✓ Downloaded");setTimeout(()=>setMsg(""),2000);
    }catch(e){setMsg("Error: "+e.message);}
  };

  // ── Review modal derived values — computed at component scope so they're always reactive ──
  const reviewOrigVal=reviewModal?parseFloat(String(reviewModal.currentVal||"0"))||0:0;
  const reviewCurVal=parseFloat(reviewValue||"0")||0;
  const reviewIsChanged=!!reviewModal&&(reviewCurVal!==reviewOrigVal);
  const reviewIsZero=reviewIsChanged&&reviewCurVal===0;
  const reviewCanModify=reviewIsChanged&&!reviewIsZero;
  const reviewCanApprove=!reviewIsChanged;

  // ── isOnlyReviewer: user is exclusively a reviewer (no inputter rights) ──
  const isOnlyReviewer=!isBuilder&&columns.length>0&&columns.filter(c=>c.col_type==="workflow").every(col=>{
    const iIds=Array.isArray(col.inputter_ids)?col.inputter_ids:JSON.parse(col.inputter_ids||'[]');
    const rIds=Array.isArray(col.reviewer_ids)?col.reviewer_ids:JSON.parse(col.reviewer_ids||'[]');
    return rIds.includes(myId)&&iIds.length>0&&!iIds.includes(myId);
  });
  const allCycles=[...cycles];

  return(
    <div style={{position:"fixed",inset:0,zIndex:700,background:"rgba(44,24,16,0.55)",display:"flex",flexDirection:"column",overflow:"hidden"}}>
      {/* Header */}
      <div style={{background:T.bgHeader,padding:"14px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
        <div>
          <div style={{color:T.textLt,fontWeight:700,fontSize:15}}>📋 Workflow View — {report.name}</div>
          {activeCycle&&<div style={{color:"rgba(245,239,230,0.7)",fontSize:12,marginTop:2}}>
            Cycle: <strong style={{color:T.textLt}}>{activeCycle.period_label}</strong>
            <span style={{marginLeft:8,fontSize:11,color:activeCycle.status==="open"?"#88FFAA":"#FF9999",fontWeight:700}}>{activeCycle.status.toUpperCase()}</span>
          </div>}
        </div>
        <div style={{display:"flex",gap:10,alignItems:"center"}}>
          {allCycles.length>1&&(
            <select value={activeCycle?.id||""} onChange={e=>{const c=allCycles.find(x=>String(x.id)===e.target.value);switchCycle(c||null);}}
              style={{padding:"5px 10px",borderRadius:6,border:"none",fontSize:12,background:"rgba(255,255,255,0.15)",color:T.textLt}}>
              {allCycles.map(c=><option key={c.id} value={c.id} style={{background:T.bgHeader}}>{cycleLabel(c)}</option>)}
            </select>
          )}
          {activeCycle&&(
            <button onClick={downloadWorkflowExcel}
              style={{padding:"5px 12px",background:"rgba(255,255,255,0.15)",color:T.textLt,border:"1px solid rgba(255,255,255,0.3)",borderRadius:6,cursor:"pointer",fontSize:12,fontWeight:600}}>
              ⬇ Excel
            </button>
          )}
          <button onClick={onClose} style={{background:"none",border:"none",color:T.textLt,fontSize:20,cursor:"pointer"}}>✕</button>
        </div>
      </div>

      {msg&&<div style={{background:"#E8F5E9",color:"#2D6A4F",padding:"6px 20px",fontSize:13,flexShrink:0}}>{msg}</div>}

      {/* Filter toolbar */}
      {!loading&&dataRows.length>0&&(
        <div style={{background:T.bgAlt,borderBottom:"1px solid "+T.border,padding:"8px 14px",flexShrink:0}}>
          {/* Row 1: Search + controls + pagination */}
          <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap",marginBottom:viewFilters.length>0?8:0}}>
            <input value={search} onChange={e=>{setSearch(e.target.value);setPage(0);}} placeholder="🔍 Search…"
              style={{padding:"5px 10px",border:"1px solid "+T.border,borderRadius:7,fontSize:12,minWidth:180,flex:1,maxWidth:280,outline:"none"}}/>
            <label style={{display:"flex",alignItems:"center",gap:5,fontSize:12,cursor:"pointer",color:T.textMd,whiteSpace:"nowrap"}}>
              <input type="checkbox" checked={showNonZeroOnly} onChange={e=>{setShowNonZeroOnly(e.target.checked);setPage(0);}}/>
              Has values only
            </label>
            <div style={{display:"flex",gap:0,border:"1px solid "+T.border,borderRadius:7,overflow:"hidden"}}>
              {[{k:"units",l:"Units"},{k:"L",l:"Lakhs"},{k:"Cr",l:"Crores"}].map(f=>(
                <button key={f.k} onClick={()=>setNumFmt(f.k)}
                  style={{padding:"4px 10px",border:"none",background:numFmt===f.k?T.primary:T.bgCard,
                    color:numFmt===f.k?T.textLt:T.textMd,cursor:"pointer",fontSize:11,fontWeight:numFmt===f.k?700:400}}>
                  {f.l}
                </button>
              ))}
            </div>
            {/* Submit All button — only for inputters, hidden for pure reviewers */}
            {activeCycle?.status==="open"&&columns.some(c=>c.col_type==="workflow")&&!isOnlyReviewer&&(
              <button onClick={submitAll}
                style={{padding:"5px 14px",background:"#185FA5",color:"#fff",border:"none",borderRadius:7,
                  cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap",flexShrink:0}}>
                ✓ Submit All Pending
              </button>
            )}
            <span style={{fontSize:11,color:T.textMd,marginLeft:"auto",whiteSpace:"nowrap"}}>
              {displayRows.length.toLocaleString()} rows{totalPages>1&&` · Page ${page+1}/${totalPages}`}
            </span>
            {/* Smart pagination */}
            {totalPages>1&&(
              <div style={{display:"flex",gap:3,alignItems:"center",flexShrink:0}}>
                <button disabled={page===0} onClick={()=>setPage(0)}
                  style={{padding:"3px 7px",border:"1px solid "+T.border,borderRadius:5,cursor:page===0?"not-allowed":"pointer",fontSize:11,background:T.bgCard,opacity:page===0?0.4:1}}>«</button>
                <button disabled={page===0} onClick={()=>setPage(p=>p-1)}
                  style={{padding:"3px 7px",border:"1px solid "+T.border,borderRadius:5,cursor:page===0?"not-allowed":"pointer",fontSize:11,background:T.bgCard,opacity:page===0?0.4:1}}>‹</button>
                {/* Page number pills — up to 5 visible */}
                {Array.from({length:totalPages},(_,i)=>i)
                  .filter(i=>Math.abs(i-page)<=2||i===0||i===totalPages-1)
                  .reduce((acc,i,idx,arr)=>{
                    if(idx>0&&i-arr[idx-1]>1)acc.push("…");
                    acc.push(i);return acc;
                  },[])
                  .map((item,idx)=>item==="…"
                    ?<span key={"e"+idx} style={{fontSize:11,color:T.textMd,padding:"0 2px"}}>…</span>
                    :<button key={item} onClick={()=>setPage(item)}
                      style={{padding:"3px 8px",border:"1px solid "+(item===page?T.primary:T.border),borderRadius:5,
                        cursor:"pointer",fontSize:11,fontWeight:item===page?700:400,
                        background:item===page?T.primary:T.bgCard,
                        color:item===page?T.textLt:T.text,minWidth:28}}>
                      {item+1}
                    </button>
                  )}
                <button disabled={page===totalPages-1} onClick={()=>setPage(p=>p+1)}
                  style={{padding:"3px 7px",border:"1px solid "+T.border,borderRadius:5,cursor:page===totalPages-1?"not-allowed":"pointer",fontSize:11,background:T.bgCard,opacity:page===totalPages-1?0.4:1}}>›</button>
                <button disabled={page===totalPages-1} onClick={()=>setPage(totalPages-1)}
                  style={{padding:"3px 7px",border:"1px solid "+T.border,borderRadius:5,cursor:page===totalPages-1?"not-allowed":"pointer",fontSize:11,background:T.bgCard,opacity:page===totalPages-1?0.4:1}}>»</button>
                {/* Direct page jump */}
                <span style={{fontSize:11,color:T.textMd}}>Go:</span>
                <input type="number" min={1} max={totalPages} defaultValue={page+1}
                  key={page}
                  onKeyDown={e=>{if(e.key==="Enter"){const v=parseInt(e.target.value)-1;if(v>=0&&v<totalPages)setPage(v);}}}
                  onBlur={e=>{const v=parseInt(e.target.value)-1;if(v>=0&&v<totalPages)setPage(v);}}
                  style={{width:40,padding:"3px 5px",border:"1px solid "+T.border,borderRadius:5,fontSize:11,textAlign:"center"}}/>
              </div>
            )}
          </div>
          {/* Row 2: Slicer filter pills (same as builder report) */}
          {viewFilters.length>0&&(
            <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:11,fontWeight:600,color:T.textMd}}>Filters:</span>
              {viewFilters.map(f=>(
                <Slicer key={f} field={f} data={dataRows}
                  active={colFilters[f]}
                  onChange={v=>{ setColFilters(p=>({...p,[f]:v})); setPage(0); }}/>
              ))}
              {Object.values(colFilters).some(v=>v!==undefined&&v!==null)&&(
                <button onClick={()=>{setColFilters({});setPage(0);}}
                  style={{fontSize:11,padding:"5px 10px",background:"none",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",color:T.textMd}}>
                  Clear all
                </button>
              )}
            </div>
          )}
        </div>
      )}

      {loading&&<div style={{padding:30,color:T.textMd,fontSize:14,textAlign:"center"}}>⏳ Loading workflow data…</div>}

      {!loading&&!activeCycle&&(
        <div style={{padding:30,color:T.textMd,fontSize:14}}>
          No open cycle found. Ask the report builder to open a cycle for this period.
          {cycles.length>0&&<div style={{marginTop:8}}>You can still view past cycles using the dropdown above.</div>}
        </div>
      )}

      {/* Main table */}
      {!loading&&(activeCycle||cycles.length>0)&&columns.length>0&&(
        <div style={{flex:1,overflow:"auto",padding:"0 0 20px 0"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:T.bgTableH,position:"sticky",top:0,zIndex:2}}>
                {/* Row label columns — with DrillColFilter (search, select-all, sort) */}
                {viewRows.map((rf,i)=>(
                  <th key={rf}
                    style={{padding:"9px 14px",textAlign:"left",fontWeight:600,color:T.text,
                      borderBottom:"2px solid "+T.border,minWidth:180,userSelect:"none",position:"relative",
                      borderLeft:i>0?"1px solid "+T.border:"none"}}>
                    <div style={{display:"flex",alignItems:"center",gap:4}}>
                      <span onClick={()=>toggleSort(rf)} style={{cursor:"pointer",flex:1,display:"flex",alignItems:"center",gap:3}}>
                        {rf}
                        <span style={{fontSize:10,opacity:0.6}}>{sortCol?.field===rf?(sortCol.dir==="asc"?"↑":"↓"):"⇅"}</span>
                      </span>
                      <DrillColFilter field={rf} data={dataRows}
                        active={colFilters[rf]}
                        onChange={v=>{ setColFilters(p=>({...p,[rf]:v})); setPage(0); }}
                        numFields={numFieldsForFilter}
                        activeSort={colSorts[rf]}
                        onSort={(f,d)=>{ setColSorts(s=>({...s,[f]:d})); setSortCol({field:f,dir:d==="za"||d==="90"?"desc":"asc"}); }}/>
                    </div>
                  </th>
                ))}
                {viewRows.length===0&&(
                  <th style={{padding:"9px 14px",fontWeight:400,color:T.textMd,fontStyle:"italic",fontSize:11,borderBottom:"2px solid "+T.border}}>
                    Row — set Rows (R) in ⚙ Setup
                  </th>
                )}
                {/* C zone — text reference columns */}
                {viewCols.map((cf,i)=>(
                  <th key={cf}
                    style={{padding:"9px 14px",textAlign:"left",fontWeight:600,color:T.tagC,
                      borderBottom:"2px solid "+T.border,minWidth:120,userSelect:"none",
                      borderLeft:"1px solid "+T.border}}>
                    <div style={{display:"flex",alignItems:"center",gap:4}}>
                      <span style={{flex:1}}>{cf}</span>
                      <DrillColFilter field={cf} data={dataRows}
                        active={colFilters[cf]}
                        onChange={v=>{ setColFilters(p=>({...p,[cf]:v})); setPage(0); }}
                        numFields={numFieldsForFilter}
                        activeSort={colSorts[cf]}
                        onSort={(f,d)=>{ setColSorts(s=>({...s,[f]:d})); setSortCol({field:f,dir:d==="za"||d==="90"?"desc":"asc"}); }}/>
                    </div>
                    <div style={{fontWeight:400,fontSize:10,color:T.textMd}}>ref</div>
                  </th>
                ))}
                {/* Separator before values */}
                {viewValues.length>0&&<th style={{width:0,padding:0,borderBottom:"2px solid "+T.border,borderRight:"2px solid "+T.borderDk}}/>}
                {cField?(
                  /* ── C zone: cross-tab column headers ── */
                  <>
                    {cVals.map(cv=>(
                      viewValues.map(({field,agg},vi)=>(
                        <th key={cv+"_"+field}
                          style={{padding:"8px 10px",textAlign:"right",fontWeight:600,color:T.text,
                            borderBottom:"2px solid "+T.border,minWidth:100,userSelect:"none",
                            borderLeft:vi===0?"2px solid "+T.border:"1px solid rgba(0,0,0,0.06)"}}>
                          <div style={{fontSize:10,color:T.tagC,fontWeight:700,marginBottom:2,textAlign:"center"}}>{cv}</div>
                          <div style={{display:"flex",justifyContent:"flex-end",alignItems:"center",gap:2}}>
                            <span onClick={()=>toggleSort(field)} style={{cursor:"pointer",fontSize:12}}>{field}</span>
                          </div>
                          <div style={{fontWeight:400,fontSize:10,color:T.textMd,textAlign:"right"}}>{agg}</div>
                        </th>
                      ))
                    ))}
                    {/* Total columns */}
                    {viewValues.map(({field,agg},vi)=>(
                      <th key={"tot_"+field}
                        style={{padding:"8px 10px",textAlign:"right",fontWeight:700,color:T.text,
                          borderBottom:"2px solid "+T.border,minWidth:100,
                          borderLeft:vi===0?"2px solid "+T.borderDk:"1px solid rgba(0,0,0,0.08)",
                          background:"rgba(92,45,26,0.04)"}}>
                        <div style={{fontSize:10,color:T.textMd,fontWeight:700,marginBottom:2,textAlign:"center"}}>Total</div>
                        <div style={{textAlign:"right"}}>{field}</div>
                        <div style={{fontWeight:400,fontSize:10,color:T.textMd,textAlign:"right"}}>{agg}</div>
                      </th>
                    ))}
                  </>
                ):(
                  /* ── No C zone: plain value columns ── */
                  viewValues.map(({field,agg})=>(
                    <th key={field}
                      style={{padding:"9px 14px",textAlign:"right",fontWeight:600,color:T.text,
                        borderBottom:"2px solid "+T.border,minWidth:120,userSelect:"none",position:"relative"}}>
                      <div style={{display:"flex",justifyContent:"flex-end",alignItems:"center",gap:3}}>
                        <DrillColFilter field={field} data={dataRows}
                          active={colFilters[field]}
                          onChange={v=>{ setColFilters(p=>({...p,[field]:v})); setPage(0); }}
                          numFields={numFieldsForFilter}
                          activeSort={colSorts[field]}
                          onSort={(f,d)=>{ setColSorts(s=>({...s,[f]:d})); setSortCol({field:f,dir:d==="za"||d==="90"?"desc":"asc"}); }}/>
                        <span onClick={()=>toggleSort(field)} style={{cursor:"pointer",display:"flex",alignItems:"center",gap:3}}>
                          {field}
                          <span style={{fontSize:10,opacity:0.6}}>{sortCol?.field===field?(sortCol.dir==="asc"?"↑":"↓"):"⇅"}</span>
                        </span>
                      </div>
                      <div style={{fontWeight:400,fontSize:10,color:T.textMd,textAlign:"right"}}>{agg}</div>
                    </th>
                  ))
                )}
                {/* Separator before collab */}
                <th style={{width:0,padding:0,borderBottom:"2px solid "+T.border,borderRight:"2px solid "+T.borderDk}}/>
                {/* Collab input columns — with DrillColFilter on collab values */}
                {columns.map(col=>{
                  const colData=displayRows.map((row,ri)=>{
                    const rk=row.__rk||computeRK(row,page*PAGE_SIZE+ri);
                    const v=values[rk+"__"+col.id];
                    return{[col.id]:v?.value!==undefined&&v.value!==null?String(v.value):""};
                  });
                  return(
                    <th key={col.id} style={{padding:"9px 14px",textAlign:"left",fontWeight:600,color:T.text,borderBottom:"2px solid "+T.border,minWidth:150}}>
                      <div style={{display:"flex",alignItems:"center",gap:4}}>
                        <span style={{flex:1}}>{col.label}</span>
                        <DrillColFilter field={col.id} data={colData}
                          active={collabFilters[col.id]}
                          onChange={v=>{ setCollabFilters(p=>({...p,[col.id]:v})); setPage(0); }}
                          numFields={new Set()}
                          activeSort={null} onSort={()=>{}}/>
                      </div>
                      <div style={{fontWeight:400,fontSize:10,color:T.textMd}}>{col.col_type==="workflow"?"Workflow":"Input Only"}</div>
                    </th>
                  );
                })}
                {/* Auto-paired approval columns (one per workflow col) */}
                {columns.filter(c=>c.col_type==="workflow").map(col=>{
                  const statusData=displayRows.map((row,ri)=>{
                    const rk=row.__rk||computeRK(row,page*PAGE_SIZE+ri);
                    const v=values[rk+"__"+col.id];
                    return{["appr_"+col.id]:v?.status||"pending"};
                  });
                  return(
                    <th key={"appr_"+col.id} style={{padding:"9px 14px",textAlign:"center",fontWeight:600,color:"#534AB7",borderBottom:"2px solid "+T.border,minWidth:140,borderLeft:"1px solid "+T.border}}>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4}}>
                        <span>Approval — {col.label}</span>
                        <DrillColFilter field={"appr_"+col.id} data={statusData}
                          active={collabFilters["appr_"+col.id]}
                          onChange={v=>{ setCollabFilters(p=>({...p,["appr_"+col.id]:v})); setPage(0); }}
                          numFields={new Set()} activeSort={null} onSort={()=>{}}/>
                      </div>
                      <div style={{fontWeight:400,fontSize:10,color:T.textMd}}>per column</div>
                    </th>
                  );
                })}
                {/* Total Approval */}
                {columns.some(c=>c.col_type==="workflow")&&(
                  <th style={{padding:"9px 14px",textAlign:"center",fontWeight:600,color:T.text,borderBottom:"2px solid "+T.border,minWidth:120,borderLeft:"2px solid "+T.borderDk}}>
                    Total Approval
                  </th>
                )}
                <th style={{padding:"9px 14px",textAlign:"center",fontWeight:600,color:T.text,borderBottom:"2px solid "+T.border,width:60}}>Trail</th>
              </tr>
            </thead>
            <tbody>
              {pagedRows.map((row,ri)=>{
                const rowBg=ri%2===0?T.bgCard:T.bgAlt;
                const rk=row.__rk||computeRK(row,page*PAGE_SIZE+ri);
                const wfCols=columns.filter(c=>c.col_type==="workflow");
                const isL1=isHierarchical&&row.__level===1;
                const isL1Expanded=isL1&&expanded.has(rk);
                const innerRows=isL1?(row.__innerOrder||[]).map(k2=>row.__inner[k2]):[];

                // Reusable V cell renderer (flat or cross-tab)
                const renderVCells=(r,bg)=>cField?(
                  <>
                    {cVals.map(cv=>viewValues.map(({field},vi)=>{
                      const cv_val=r.__cGroups?.[cv]?.[field];
                      return<td key={cv+"_"+field} style={{padding:"7px 10px",textAlign:"right",color:cv_val!=null?T.numColor:T.textMd,fontWeight:500,whiteSpace:"nowrap",background:bg,borderLeft:vi===0?"2px solid "+T.border:"1px solid rgba(0,0,0,0.06)"}}>
                        {cv_val!=null?fmtNum(cv_val):"—"}</td>;
                    }))}
                    {viewValues.map(({field},vi)=><td key={"tot_"+field} style={{padding:"7px 10px",textAlign:"right",color:T.numColor,fontWeight:700,whiteSpace:"nowrap",borderLeft:vi===0?"2px solid "+T.borderDk:"1px solid rgba(0,0,0,0.08)",background:"rgba(92,45,26,0.04)"}}>{fmtNum(r[field])}</td>)}
                  </>
                ):viewValues.map(({field})=><td key={field} style={{padding:"8px 14px",textAlign:"right",color:T.numColor,fontWeight:500,background:bg,whiteSpace:"nowrap",opacity:0.9}}>{fmtNum(r[field])}</td>);

                // Reusable collab input cells — auto-save on type (debounced), no remarks, no manual save button
                const renderCollabCells=(cellRk)=>columns.map(col=>{
                  const dk=draftKey(cellRk,col.id);
                  const existing=values[dk];const draft=draftMap[dk];
                  // Always show raw number in the input — no formatting (avoids focus/cursor instability)
                  const displayVal=draft?.value!==undefined?String(draft.value):(existing?.value!==undefined?String(existing.value):"");
                  const isSav=saving[dk];const cycleClosed=activeCycle?.status==="closed";const canI=!cycleClosed&&canInput(col);
                  const isDirty=draft?.value!==undefined;
                  const cellErr=cellErrors[dk];
                  return<td key={col.id} style={{padding:"6px 14px",verticalAlign:"middle"}}>
                    {canI?(
                      <div style={{position:"relative",display:"inline-block"}}>
                        <input type="text" inputMode="numeric" value={displayVal} placeholder="0"
                          onFocus={e=>e.target.select()}
                          onChange={e=>{
                            const val=e.target.value.replace(/[^0-9.\-]/g,"");
                            // Clear error and isReverted flag as soon as user types something new
                            if(cellErr)setCellErrors(er=>({...er,[dk]:null}));
                            setDraftMap(d=>({...d,[dk]:{value:val}})); // fresh object clears isReverted
                            if(autoSaveTimers.current[dk])clearTimeout(autoSaveTimers.current[dk]);
                            autoSaveTimers.current[dk]=setTimeout(()=>saveDraft(cellRk,col),800);
                          }}
                          onBlur={()=>{
                            if(autoSaveTimers.current[dk])clearTimeout(autoSaveTimers.current[dk]);
                            saveDraft(cellRk,col);
                          }}
                          disabled={isSav}
                          style={{width:110,padding:"5px 9px",borderRadius:6,fontSize:13,outline:"none",
                            border:"1px solid "+(cellErr?"#A32D2D":isDirty?"#C8922A":T.border),
                            background:cellErr?"#FFF0F0":isDirty?"#FFFDE7":T.bgCard}}/>
                        {isSav&&<span style={{position:"absolute",right:-18,top:7,fontSize:10,color:T.textMd}}>…</span>}
                        {cellErr&&(
                          <div style={{position:"absolute",top:"calc(100% + 4px)",left:0,zIndex:200,
                            background:"#5C2D1A",color:"#FFF5EE",borderRadius:7,
                            padding:"7px 28px 7px 10px",
                            fontSize:11,whiteSpace:"normal",wordBreak:"break-word",
                            boxShadow:"0 4px 14px rgba(44,24,16,0.4)",
                            maxWidth:260,lineHeight:1.5,minWidth:160}}>
                            ⚠ {cellErr}
                            <button onClick={()=>setCellErrors(er=>({...er,[dk]:null}))}
                              style={{position:"absolute",top:5,right:5,
                                background:"rgba(255,255,255,0.25)",border:"none",
                                color:"#fff",cursor:"pointer",fontSize:12,fontWeight:900,
                                lineHeight:1,borderRadius:3,padding:"2px 5px",display:"flex",
                                alignItems:"center",justifyContent:"center"}}>✕</button>
                          </div>
                        )}
                      </div>
                    ):(
                      <span style={{fontSize:13,color:T.text,minWidth:60,display:"block",textAlign:"right",paddingRight:4}}>
                        {existing?.value!==undefined&&existing.value!==null?fmtNum(existing.value):"—"}
                      </span>
                    )}
                  </td>;
                });

                // Reusable approval cells — status badge + amount after action + Review button
                const renderApprovalCells=(cellRk)=>wfCols.map(col=>{
                  const dk=draftKey(cellRk,col.id);const existing=values[dk];
                  const cycleClosed=activeCycle?.status==="closed";const canR=!cycleClosed&&canReview(col);
                  const effectiveVal=existing?.reviewer_value!=null?existing.reviewer_value:existing?.value;
                  const showAmt=existing&&['approved','modified','hold'].includes(existing.status)&&effectiveVal!=null;
                  const rejAmt=existing?.status==='rejected';
                  return<td key={"appr_"+col.id} style={{padding:"6px 14px",verticalAlign:"middle",textAlign:"center",borderLeft:"1px solid "+T.border}}>
                    {existing?statusBadge(existing.status):<span style={{fontSize:11,color:T.textMd}}>—</span>}
                    {/* Show the effective amount after action */}
                    {showAmt&&<div style={{fontSize:12,fontWeight:600,color:T.numColor,marginTop:3}}>{fmtNum(effectiveVal)}</div>}
                    {rejAmt&&<div style={{fontSize:12,color:T.textMd,marginTop:3}}>0</div>}
                    {canR&&existing?.status==="submitted"&&(
                      <button onClick={()=>{setReviewModal({rowKey:cellRk,col_id:col.id,colLabel:col.label,currentVal:existing.value});setReviewRemarks("");setReviewValue(String(existing.value??""));}}
                        style={{marginTop:4,padding:"3px 10px",background:"#2D6A4F",color:"#fff",border:"none",borderRadius:5,cursor:"pointer",fontSize:11,fontWeight:600,display:"block",margin:"4px auto 0"}}>
                        Review
                      </button>
                    )}
                  </td>;
                });

                // Drill-down helper for a specific row grouping
                const openDrill=(drillRfs,drillRowKey,label)=>setDrillDown({rowKey:drillRowKey,colVal:null,rFs:drillRfs,cF:null,metricLabel:label});
                const vCellStyle=(bg,vi,extra)=>({padding:"8px 12px",textAlign:"right",color:T.numColor,fontWeight:500,
                  background:bg,whiteSpace:"nowrap",cursor:"pointer",userSelect:"none",
                  borderLeft:vi===0?"none":"none",...extra});
                return(
                  <React.Fragment key={rk}>
                  {/* ── Main summary row (level-1 or single-level) ── */}
                  <tr style={{background:isL1?T.bgStat:rowBg,borderBottom:"1px solid "+T.border}}>
                    {/* Row label cells — NOT clickable */}
                    {viewRows.map((rf,i)=>(
                      <td key={rf} style={{padding:"8px 14px",fontWeight:600,color:T.text,
                        maxWidth:220,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                        borderLeft:i>0?"1px solid "+T.border:"none"}}
                        title={String(row[rf]||"")}>
                        {i===0&&isL1?(
                          <span style={{display:"flex",alignItems:"center",gap:6}}>
                            <button onClick={()=>setExpanded(s=>{const n=new Set(s);n.has(rk)?n.delete(rk):n.add(rk);return n;})}
                              title={isL1Expanded?"Collapse":"Expand"}
                              style={{background:isL1Expanded?T.primary:"none",border:"1px solid "+T.border,borderRadius:4,width:18,height:18,cursor:"pointer",
                                fontSize:11,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,
                                color:isL1Expanded?T.textLt:T.primary,fontWeight:700}}>
                              {isL1Expanded?"−":"+"}
                            </button>
                            <span>{row[rf]||rk}</span>
                            <span style={{fontSize:10,color:T.textMd,fontWeight:400}}>({row.__count})</span>
                          </span>
                        ):(String(row[rf]||""))}
                      </td>
                    ))}
                    {viewRows.length===0&&<td style={{padding:"8px 14px",color:T.textMd}}>{page*PAGE_SIZE+ri+1}</td>}
                    {/* C zone cells */}
                    {viewCols.map(cf=>(
                      <td key={cf} style={{padding:"8px 14px",color:T.text,borderLeft:"1px solid "+T.border,
                        maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}
                        title={String(row[cf]||"")}>{String(row[cf]||"—")}</td>
                    ))}
                    {/* Separator */}
                    {viewValues.length>0&&<td style={{width:0,padding:0,borderRight:"2px solid "+T.borderDk,background:isL1?T.bgStat:rowBg}}/>}
                    {/* Value cells — CLICKABLE → drill-down into raw records */}
                    {cField?(
                      <>
                        {cVals.map(cv=>viewValues.map(({field},vi)=>{
                          const cv_val=row.__cGroups?.[cv]?.[field];
                          return<td key={cv+"_"+field}
                            onClick={()=>openDrill(isL1?[viewRows[0]]:viewRows,isL1?[String(row[viewRows[0]]||"")]:viewRows.map(f=>String(row[f]||"")),field+" · "+cv)}
                            style={{...vCellStyle(isL1?T.bgStat:rowBg,vi,{borderLeft:vi===0?"2px solid "+T.border:"1px solid rgba(0,0,0,0.06)",color:cv_val!=null?T.numColor:T.textMd})}}
                            title="Click to see raw records">
                            {cv_val!=null?fmtNum(cv_val):"—"}</td>;
                        }))}
                        {viewValues.map(({field},vi)=><td key={"tot_"+field}
                          onClick={()=>openDrill(isL1?[viewRows[0]]:viewRows,isL1?[String(row[viewRows[0]]||"")]:viewRows.map(f=>String(row[f]||"")),field+" (Total)")}
                          style={{...vCellStyle("rgba(92,45,26,0.04)",vi,{borderLeft:vi===0?"2px solid "+T.borderDk:"1px solid rgba(0,0,0,0.08)",fontWeight:700})}}
                          title="Click to see raw records">{fmtNum(row[field])}</td>)}
                      </>
                    ):(
                      viewValues.map(({field})=><td key={field}
                        onClick={()=>openDrill(isL1?[viewRows[0]]:viewRows,isL1?[String(row[viewRows[0]]||"")]:viewRows.map(f=>String(row[f]||"")),field)}
                        style={{...vCellStyle(isL1?T.bgStat:rowBg,0,{})}}
                        onMouseEnter={e=>e.currentTarget.style.background="rgba(92,45,26,0.08)"}
                        onMouseLeave={e=>e.currentTarget.style.background=isL1?T.bgStat:rowBg}
                        title="Click to see raw records">{fmtNum(row[field])}</td>)
                    )}
                    {/* Separator before collab */}
                    <td style={{width:0,padding:0,borderRight:"2px solid "+T.borderDk,background:isL1?T.bgStat:rowBg}}/>
                    {/* Collab inputs — ALWAYS on main rows (primary level), never on sub-rows */}
                    {renderCollabCells(rk)}
                    {renderApprovalCells(rk)}
                    {columns.some(c=>c.col_type==="workflow")&&<td style={{padding:"8px 14px",textAlign:"center",borderLeft:"2px solid "+T.borderDk}}>{totalApprovalBadge(wfCols,rk)}</td>}
                    <td style={{padding:"8px 14px",textAlign:"center"}}>
                      <button onClick={()=>openAudit(rk)} title="View audit trail"
                        style={{background:"none",border:"1px solid "+T.border,borderRadius:5,cursor:"pointer",fontSize:11,padding:"3px 7px",color:T.textMd}}>
                        🕵
                      </button>
                    </td>
                  </tr>

                  {/* ── Level-2 sub-rows (expand from level-1) — values only, NO collab inputs ── */}
                  {isL1Expanded&&innerRows.map((inner,ii)=>{
                    const innerRk=inner.__rk;
                    const innerBg=ii%2===0?T.bgCard:T.bgAlt;
                    return(
                      <tr key={innerRk} style={{background:innerBg,borderBottom:"0.5px solid "+T.border}}>
                        {/* Row label cells — indented, NOT clickable */}
                        {viewRows.map((rf,i)=>(
                          <td key={rf} style={{padding:"6px 14px 6px "+(i===0?32:14)+"px",
                            color:i===0?T.textMd:T.text,fontWeight:i===0?400:500,fontSize:12,
                            overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                            borderLeft:i>0?"1px solid "+T.border:"none"}}>
                            {i===0&&<span style={{color:T.textMd,marginRight:3,fontSize:10}}>⤷</span>}
                            {String(inner[rf]||"")}
                          </td>
                        ))}
                        {viewCols.map(cf=><td key={cf} style={{padding:"6px 14px",fontSize:12,color:T.text,borderLeft:"1px solid "+T.border}}>{String(inner[cf]||"—")}</td>)}
                        {/* Separator */}
                        {viewValues.length>0&&<td style={{width:0,padding:0,borderRight:"2px solid "+T.borderDk,background:innerBg}}/>}
                        {/* Value cells — CLICKABLE → drill-down for this specific sub-group */}
                        {cField?(
                          <>
                            {cVals.map(cv=>viewValues.map(({field},vi)=>{
                              const cv_val=inner.__cGroups?.[cv]?.[field];
                              return<td key={cv+"_"+field}
                                onClick={()=>openDrill(viewRows,viewRows.map(f=>String(inner.__rows[0]?.[f]||"")),field+" · "+cv)}
                                style={{...vCellStyle(innerBg,vi,{borderLeft:vi===0?"2px solid "+T.border:"1px solid rgba(0,0,0,0.06)",color:cv_val!=null?T.numColor:T.textMd,fontSize:12})}}
                                title="Click to see raw records">{cv_val!=null?fmtNum(cv_val):"—"}</td>;
                            }))}
                            {viewValues.map(({field},vi)=><td key={"tot_"+field}
                              onClick={()=>openDrill(viewRows,viewRows.map(f=>String(inner.__rows[0]?.[f]||"")),field)}
                              style={{...vCellStyle("rgba(92,45,26,0.04)",vi,{borderLeft:vi===0?"2px solid "+T.borderDk:"1px solid rgba(0,0,0,0.08)",fontWeight:600,fontSize:12})}}
                              title="Click to see raw records">{fmtNum(inner[field])}</td>)}
                          </>
                        ):(
                          viewValues.map(({field})=><td key={field}
                            onClick={()=>openDrill(viewRows,viewRows.map(f=>String(inner.__rows[0]?.[f]||"")),field)}
                            style={{...vCellStyle(innerBg,0,{fontSize:12})}}
                            onMouseEnter={e=>e.currentTarget.style.background="rgba(92,45,26,0.08)"}
                            onMouseLeave={e=>e.currentTarget.style.background=innerBg}
                            title="Click to see raw records">{fmtNum(inner[field])}</td>)
                        )}
                        {/* Collab inputs at the granular (L2) level — full composite row key */}
                        <td style={{width:0,padding:0,borderRight:"2px solid "+T.borderDk,background:innerBg}}/>
                        {renderCollabCells(innerRk)}
                        {renderApprovalCells(innerRk)}
                        {columns.some(c=>c.col_type==="workflow")&&<td style={{padding:"8px 14px",textAlign:"center",borderLeft:"2px solid "+T.borderDk}}>{totalApprovalBadge(wfCols,innerRk)}</td>}
                        <td style={{padding:"8px 14px",textAlign:"center"}}>
                          <button onClick={()=>openAudit(innerRk)} title="View audit trail"
                            style={{background:"none",border:"1px solid "+T.border,borderRadius:5,cursor:"pointer",fontSize:11,padding:"3px 7px",color:T.textMd}}>
                            🕵
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                  </React.Fragment>
                );
              })}
              {pagedRows.length===0&&(
                <tr><td colSpan={viewRows.length+viewCols.length+viewValues.length+columns.length+columns.filter(c=>c.col_type==="workflow").length+3} style={{padding:20,color:T.textMd,textAlign:"center"}}>No data rows.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {!loading&&columns.length===0&&(
        <div style={{padding:30,color:T.textMd,fontSize:14}}>No collab columns defined yet. Use ⚙ Setup Columns & Cycle in the Workflow tab.</div>
      )}

      {/* ── Drill-Down — reuses the exact same DrillDown panel as the Report Builder ── */}
      {drillDown&&(
        <DrillDown
          data={dataRows}
          target={drillDown}
          fields={allDataFields}
          numFields={allNumFields}
          onClose={()=>setDrillDown(null)}
          numFmt={numFmt}
          savedHiddenCols={[]}
          savedColFmts={{}}
          onSaveHiddenCols={null}
          onSaveColFilters={null}
          configValues={viewValues}
          activeFilters={colFilters}
        />
      )}

      {/* Review Modal — uses component-scope reviewIsChanged/reviewCanModify for reliable reactivity */}
      {reviewModal&&(
        <div style={{position:"absolute",inset:0,zIndex:800,background:"rgba(44,24,16,0.5)",display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:T.bgCard,borderRadius:12,padding:24,width:"min(440px,92vw)",boxShadow:"0 12px 40px rgba(0,0,0,0.3)"}}>
            <div style={{fontWeight:700,fontSize:15,color:T.primary,marginBottom:14}}>Review: {reviewModal.colLabel}</div>

            {/* Editable value — approver can modify */}
            <div style={{marginBottom:14}}>
              <label style={{fontSize:12,fontWeight:600,color:T.textMd,display:"block",marginBottom:4}}>
                Submitted Value — modify if needed:
              </label>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <input type="text" inputMode="numeric" value={reviewValue}
                  onChange={e=>setReviewValue(e.target.value.replace(/[^0-9.\-]/g,""))}
                  style={{flex:1,padding:"8px 12px",border:"2px solid "+(reviewIsChanged?"#C8922A":T.border),
                    borderRadius:7,fontSize:15,fontWeight:600,color:T.text,
                    background:reviewIsChanged?"#FFFDE7":T.bgCard,outline:"none"}}/>
                {reviewIsChanged&&(
                  <span style={{fontSize:11,color:"#7A3E00",fontWeight:600,whiteSpace:"nowrap"}}>
                    {reviewIsZero?"→ Reject":"Was: "+reviewModal.currentVal}
                  </span>
                )}
              </div>
              {reviewCanModify&&<div style={{fontSize:11,color:"#7A3E00",marginTop:4}}>✏ Changed — will be saved as <strong>MODIFIED</strong></div>}
              {reviewIsZero&&<div style={{fontSize:11,color:"#721C24",marginTop:4}}>⚠ Value is 0 — will be saved as <strong>REJECTED</strong></div>}
            </div>

            {/* Reviewer remarks */}
            <div style={{marginBottom:16}}>
              <label style={{fontSize:12,fontWeight:600,color:T.textMd,display:"block",marginBottom:4}}>Remarks (optional)</label>
              <textarea value={reviewRemarks} onChange={e=>setReviewRemarks(e.target.value)}
                rows={2} placeholder="Add remarks for inputter..."
                style={{width:"100%",padding:"7px 10px",border:"1px solid "+T.border,borderRadius:7,fontSize:12,boxSizing:"border-box",resize:"vertical"}}/>
            </div>

            {/* Action buttons */}
            <div style={{display:"flex",gap:8,flexWrap:"wrap",justifyContent:"flex-end",alignItems:"center"}}>
              <button onClick={()=>{setReviewModal(null);setReviewRemarks("");setReviewValue("");}}
                style={{padding:"7px 14px",background:"none",border:"1px solid "+T.border,borderRadius:7,cursor:"pointer",fontSize:12}}>
                Cancel
              </button>
              <button onMouseDown={e=>e.preventDefault()} onClick={()=>doReview("hold")}
                style={{padding:"7px 14px",background:"#6B5B95",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600}}>
                ⏸ Hold
              </button>
              <button onMouseDown={e=>e.preventDefault()} onClick={()=>doReview("rejected")}
                style={{padding:"7px 14px",background:reviewIsZero?"#721C24":"#A32D2D",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600,
                  boxShadow:reviewIsZero?"0 0 0 3px rgba(163,45,45,0.4)":"none"}}>
                ❌ Reject
              </button>
              {reviewCanModify&&(
                <button onMouseDown={e=>e.preventDefault()} onClick={()=>doReview("modified")}
                  style={{padding:"7px 14px",background:"#8B4400",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600,
                    boxShadow:"0 0 0 3px rgba(139,68,0,0.35)"}}>
                  ✏ Modified
                </button>
              )}
              <button onMouseDown={e=>e.preventDefault()} onClick={()=>doReview("approved")} disabled={!reviewCanApprove}
                style={{padding:"7px 14px",background:reviewCanApprove?"#2D6A4F":"#ccc",color:"#fff",border:"none",borderRadius:7,
                  cursor:reviewCanApprove?"pointer":"not-allowed",fontSize:12,fontWeight:700,opacity:reviewCanApprove?1:0.5}}>
                ✅ Approve
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Audit Trail Panel */}
      {auditRow&&(
        <div style={{position:"absolute",inset:0,zIndex:800,background:"rgba(44,24,16,0.5)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:T.bgCard,borderRadius:12,padding:0,width:"min(600px,95vw)",maxHeight:"80vh",display:"flex",flexDirection:"column",boxShadow:"0 12px 40px rgba(0,0,0,0.3)"}}>
            <div style={{background:T.bgHeader,borderRadius:"12px 12px 0 0",padding:"14px 18px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{color:T.textLt,fontWeight:700,fontSize:14}}>🕵 Audit Trail — {auditRow}</span>
              <button onClick={()=>setAuditRow(null)} style={{background:"none",border:"none",color:T.textLt,fontSize:18,cursor:"pointer"}}>✕</button>
            </div>
            <div style={{overflowY:"auto",flex:1,padding:14}}>
              {auditData.length===0&&<div style={{color:T.textMd,fontSize:13}}>No audit entries for this row.</div>}
              {auditData.map(a=>(
                <div key={a.id} style={{borderBottom:"1px solid "+T.border,padding:"8px 0",fontSize:12}}>
                  <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
                    <span style={{fontWeight:700,color:T.text}}>{a.username||"(deleted user)"}</span>
                    <span style={{background:T.bgAlt,color:T.textMd,padding:"2px 7px",borderRadius:8,fontSize:10,textTransform:"uppercase"}}>{a.action}</span>
                    {a.value!==null&&a.value!==undefined&&<span style={{color:T.numColor,fontWeight:600}}>{a.value}</span>}
                    <span style={{color:T.textMd,marginLeft:"auto"}}>{new Date(a.created_at).toLocaleString()}</span>
                  </div>
                  {a.remarks&&<div style={{color:T.textMd,marginTop:3,fontStyle:"italic"}}>{a.remarks}</div>}
                  {a.col_id&&<div style={{color:T.textMd,fontSize:11}}>Column ID: {a.col_id}</div>}
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Auto-Refresh Schedule Panel ───────────────────────────────────────────────
function AutoRefreshPanel({reportId, reportName, onClose}) {
  const [schedule,setSchedule]=useState(null);
  const [saving,setSaving]=useState(false);
  const [msg,setMsg]=useState("");
  const intervals=[
    {val:0,label:"Off"},
    {val:5,label:"5 minutes"},
    {val:15,label:"15 minutes"},
    {val:30,label:"30 minutes"},
    {val:60,label:"1 hour"},
    {val:240,label:"4 hours"},
    {val:1440,label:"Daily (24 hrs)"},
  ];
  useEffect(()=>{
    getRefreshSchedule(reportId)
      .then(s=>setSchedule(s))
      .catch(()=>setSchedule({interval_minutes:0,enabled:false,last_run:null}));
  },[reportId]);

  async function save(){
    setSaving(true);setMsg("");
    try{
      await setRefreshSchedule(reportId, schedule.interval_minutes, schedule.interval_minutes>0);
      setMsg("✓ Schedule saved!");
    }catch(e){setMsg("Error: "+e.message);}
    finally{setSaving(false);}
  }

  const inp={padding:"8px 12px",border:"1px solid "+T.border,borderRadius:6,fontSize:13,
    background:T.bgCard,color:T.text,outline:"none",cursor:"pointer",width:"100%",boxSizing:"border-box"};
  return(
    <div style={{position:"fixed",inset:0,zIndex:650,background:"rgba(44,24,16,0.55)",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:T.bgCard,borderRadius:14,width:"min(420px,92vw)",boxShadow:"0 12px 48px rgba(44,24,16,0.35)"}}>
        <div style={{padding:"16px 20px",background:T.bgHeader,borderRadius:"14px 14px 0 0",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <span style={{fontWeight:700,fontSize:15,color:T.textLt}}>⏱ Auto-Refresh Schedule</span>
          <button onClick={onClose} style={{border:"none",background:"rgba(255,255,255,0.15)",color:T.textLt,borderRadius:6,width:28,height:28,cursor:"pointer",fontSize:16}}>×</button>
        </div>
        <div style={{padding:20}}>
          <div style={{fontSize:13,color:T.textMd,marginBottom:16}}>
            Report: <strong style={{color:T.text}}>{reportName}</strong>
          </div>
          {!schedule
            ?<div style={{textAlign:"center",padding:20,color:T.textMd}}>Loading…</div>
            :<>
              <div style={{marginBottom:16}}>
                <div style={{fontSize:11,fontWeight:600,color:T.textMd,marginBottom:6}}>Refresh interval (server-side)</div>
                <select value={schedule.interval_minutes} onChange={e=>setSchedule(s=>({...s,interval_minutes:+e.target.value}))} style={inp}>
                  {intervals.map(i=><option key={i.val} value={i.val}>{i.label}</option>)}
                </select>
                <div style={{fontSize:11,color:T.textMd,marginTop:6}}>
                  {schedule.interval_minutes===0
                    ?"Auto-refresh is disabled. Data only updates on manual refresh."
                    :`Server will re-fetch the Excel source every ${intervals.find(i=>i.val===schedule.interval_minutes)?.label?.toLowerCase()}.`}
                </div>
              </div>
              {schedule.last_run&&(
                <div style={{fontSize:11,color:T.textMd,marginBottom:12,padding:"8px 12px",background:T.bgStat,borderRadius:7,border:"0.5px solid "+T.border}}>
                  Last run: {new Date(schedule.last_run).toLocaleString()}
                  {schedule.next_run&&<span style={{marginLeft:10}}>· Next: {new Date(schedule.next_run).toLocaleString()}</span>}
                </div>
              )}
              {msg&&<div style={{padding:"8px 12px",borderRadius:7,fontSize:12,marginBottom:12,
                background:msg.startsWith("✓")?"rgba(45,106,79,0.1)":"rgba(163,45,45,0.1)",
                color:msg.startsWith("✓")?T.success:T.danger}}>{msg}</div>}
              <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
                <button onClick={onClose} style={{padding:"8px 16px",background:"none",border:"1px solid "+T.border,borderRadius:7,cursor:"pointer",fontSize:13,color:T.text}}>Close</button>
                <button onClick={save} disabled={saving} style={{padding:"8px 18px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,cursor:saving?"wait":"pointer",fontSize:13,fontWeight:700,opacity:saving?0.6:1}}>
                  {saving?"Saving…":"Save Schedule"}
                </button>
              </div>
            </>}
        </div>
      </div>
    </div>
  );
}

// ── MyReportsViewer — embedded viewer for subadmin_user role ──────────────────
// Shows only published reports that are assigned to the current user.
function MyReportsViewer({savedReports,onLoadReportData}) {
  // Only show published reports that were NOT created by the current user (i.e., assigned)
  const publishedReports = savedReports.filter(r => r.isPublished);
  const [activeId,setActiveId]=useState(publishedReports[0]?.id||null);
  const [loadedData,setLoadedData]=useState({});
  const [dataLoading,setDataLoading]=useState(false);
  const [refreshing,setRefreshing]=useState(false);
  const [lastRefreshed,setLastRefreshed]=useState(null);
  const [refreshError,setRefreshError]=useState("");
  const [userActiveTabIdx,setUserActiveTabIdx]=useState(0);
  const autoRefreshRef=useRef(null);
  const numFieldsRef=useRef(new Set());
  const AUTO_REFRESH_MS=5*60*1000;
  const currentMeta=publishedReports.find(r=>r.id===activeId)||publishedReports[0]||null;

  useEffect(()=>{
    if(!activeId&&publishedReports.length)setActiveId(publishedReports[0].id);
  },[publishedReports.length]);

  useEffect(()=>{
    if(!currentMeta)return;
    setDataLoading(true);setRefreshError("");setUserActiveTabIdx(0);
    onLoadReportData(currentMeta.id).then(data=>{
      if(data.numFields instanceof Set&&data.numFields.size>0)numFieldsRef.current=data.numFields;
      setLoadedData(p=>({...p,[currentMeta.id]:data}));
      setDataLoading(false);
    }).catch(()=>setDataLoading(false));
    // row_count in deps: when admin saves new data the count updates → this
    // effect re-fires → preview re-fetches from the freshly-written cache/DB
  },[currentMeta?.id, currentMeta?.rows]);

  async function refreshFromSource(silent=false){
    if(!currentMeta)return;
    const links=currentMeta.config?.sourceLinks||[];
    if(!links.length)return;
    if(!silent)setRefreshing(true);setRefreshError("");
    try{
      const result=await fetchUrlViaProxy(links[0].url,links[0].sheet||undefined);
      const freshFields=result.fields||(result.rows.length?Object.keys(result.rows[0]):[]);
      // [] is truthy — must check .length>0 or backend's empty array would replace
      // numFieldsRef with an empty Set, turning all values to 0.
      const backendNF = result.numFields&&result.numFields.length>0 ? new Set(result.numFields) : null;
      const freshNumFields = backendNF && backendNF.size>0 ? backendNF : numFieldsRef.current;
      if(freshNumFields.size>0) numFieldsRef.current=freshNumFields;
      setLoadedData(p=>({...p,[currentMeta.id]:{rows:result.rows,fields:freshFields,numFields:freshNumFields}}));
      setLastRefreshed(new Date());
    }catch(e){if(!silent)setRefreshError(e.message);}
    finally{setRefreshing(false);}
  }

  useEffect(()=>{
    if(autoRefreshRef.current)clearInterval(autoRefreshRef.current);
    const links=currentMeta?.config?.sourceLinks||[];
    if(!links.length)return;
    const t=setTimeout(()=>refreshFromSource(true),1500);
    autoRefreshRef.current=setInterval(()=>refreshFromSource(true),AUTO_REFRESH_MS);
    return()=>{clearTimeout(t);if(autoRefreshRef.current)clearInterval(autoRefreshRef.current);};
  },[currentMeta?.id]);

  const currentData=currentMeta?loadedData[currentMeta.id]:null;

  // Filter/sort helpers
  const [globalFilters,setGlobalFilters]=useState({});
  const [globalPivotFilters,setGlobalPivotFilters]=useState({});
  const filterKey=(id,t)=>id+":"+t;
  const getStoredFilter=(id,t,meta)=>globalFilters[filterKey(id,t)]||(meta?.config?.defaultFilters||{});
  const getStoredPivotFilter=(id,t,meta)=>globalPivotFilters[filterKey(id,t)]||(meta?.config?.defaultPivotFilters||{});

  if(!publishedReports.length)return(
    <div style={{padding:40,textAlign:"center"}}>
      <div style={{fontSize:36,marginBottom:12}}>📋</div>
      <div style={{fontWeight:700,fontSize:16,color:T.primary,marginBottom:8}}>No reports assigned to you yet</div>
      <div style={{fontSize:13,color:T.textMd}}>Ask your Super Admin to publish and assign reports to your account.</div>
    </div>
  );

  return(
    <div style={{padding:20}}>
      {/* Report selector */}
      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:16,flexWrap:"wrap"}}>
        <span style={{fontWeight:700,fontSize:14,color:T.primary}}>View Report:</span>
        <select value={activeId||""} onChange={e=>{setActiveId(e.target.value);setUserActiveTabIdx(0);}}
          style={{padding:"6px 10px",border:"1px solid "+T.border,borderRadius:7,background:T.bgCard,color:T.text,fontSize:13,cursor:"pointer",outline:"none",maxWidth:280}}>
          {publishedReports.map(r=><option key={r.id} value={r.id}>{r.name}</option>)}
        </select>
        {currentMeta?.config?.sourceLinks?.length>0&&(
          <button onClick={()=>refreshFromSource(false)} disabled={refreshing}
            style={{padding:"6px 14px",background:refreshing?"rgba(92,45,26,0.08)":T.primary,color:refreshing?T.primary:T.textLt,
              border:"1px solid "+T.primary,borderRadius:7,cursor:refreshing?"not-allowed":"pointer",fontSize:12,fontWeight:600}}>
            <span style={{display:"inline-block",animation:refreshing?"spin 0.8s linear infinite":"none"}}>↻</span>
            {refreshing?" Updating...":" Refresh"}
          </button>
        )}
        {lastRefreshed&&<span style={{fontSize:11,color:T.success}}>✓ Refreshed {lastRefreshed.toLocaleTimeString()}</span>}
      </div>
      {refreshError&&<div style={{padding:"8px 12px",background:"rgba(163,45,45,0.07)",border:"1px solid rgba(163,45,45,0.25)",borderRadius:7,fontSize:12,color:T.danger,marginBottom:10}}>{refreshError}</div>}
      {dataLoading&&<div style={{padding:40,textAlign:"center",color:T.textMd,fontSize:13}}>⏳ Loading report…</div>}
      {!dataLoading&&currentMeta&&currentData&&(
        <Report
          key={currentMeta.id}
          externalFilters={getStoredFilter(currentMeta.id,userActiveTabIdx,currentMeta)}
          externalPivotFilters={getStoredPivotFilter(currentMeta.id,userActiveTabIdx,currentMeta)}
          onExternalFiltersChange={f=>setGlobalFilters(p=>({...p,[filterKey(currentMeta.id,userActiveTabIdx)]:f}))}
          onExternalPivotFiltersChange={pf=>setGlobalPivotFilters(p=>({...p,[filterKey(currentMeta.id,userActiveTabIdx)]:pf}))}
          config={(()=>{
            const topCfg=currentMeta.config||{};
            const tabCfg=topCfg.tabs&&topCfg.tabs[userActiveTabIdx];
            if(tabCfg){
              const rw={allowedFmts:topCfg.allowedFmts,name:topCfg.name};
              const ts=tabCfg.config||{};
              return{...rw,...ts,defaultFilters:ts.defaultFilters??topCfg.defaultFilters??{},
                colExcluded:ts.colExcluded||[],_reportId:currentMeta.id};
            }
            return{...topCfg,_reportId:currentMeta.id};
          })()}
          data={currentData.rows} fields={currentData.fields} numFields={currentData.numFields}
          showExport cardFields={currentMeta.cardFields||[]}
          tabs={currentMeta.config?.tabs||null}
          activeTabIdx={userActiveTabIdx}
          onTabChange={setUserActiveTabIdx}/>
      )}
    </div>
  );
}

// ── User view ──────────────────────────────────────────────────────────────────
function UserView({onLogout,savedReports,onLoadReportData,currentUser,currentRole}) {
  const isMobileView=useViewport();
  const [activeId,setActiveId]=useState(null);
  const [collabViewReport,setCollabViewReport]=useState(null);
  const [dataLoading,setDataLoading]=useState(false);
  const [refreshing,setRefreshing]=useState(false); // background refresh (doesn't blank report)
  const [lastRefreshed,setLastRefreshed]=useState(null);
  const [refreshError,setRefreshError]=useState("");
  const [userActiveTabIdx,setUserActiveTabIdx]=useState(0);
  const autoRefreshRef=useRef(null);
  const numFieldsRef=useRef(new Set()); // always current numFields — avoids stale closure in refreshFromSource

  // ── localStorage cache helpers ──────────────────────────────────────────────
  function cacheKey(id){return "rh_data_"+id;}
  function saveCache(id,data){
    try{localStorage.setItem(cacheKey(id),JSON.stringify({
      rows:data.rows,fields:data.fields,
      numFields:[...(data.numFields instanceof Set?data.numFields:new Set(data.numFields||[]))],
      ts:Date.now()
    }));}catch(e){/* quota exceeded — ignore */}
  }
  function loadCache(id){
    try{
      const raw=localStorage.getItem(cacheKey(id));
      if(!raw)return null;
      const d=JSON.parse(raw);
      const nf=new Set(d.numFields||[]);
      // Discard corrupted cache entries where numFields is empty —
      // forces a fresh DB fetch which auto-detects numeric fields.
      if(nf.size===0) { localStorage.removeItem(cacheKey(id)); return null; }
      return {...d,numFields:nf};
    }catch(e){return null;}
  }

  // Initialise loadedData from localStorage cache for instant display
  const [loadedData,setLoadedData]=useState(()=>{
    const init={};
    try{
      Object.keys(localStorage).filter(k=>k.startsWith("rh_data_")).forEach(k=>{
        const id=k.replace("rh_data_","");
        const d=loadCache(id);
        if(d)init[id]=d;
      });
    }catch(e){}
    return init;
  });

  // Global filter store keyed by "reportId:tabIdx" — survives navigation and tab switches
  const [globalFilters,setGlobalFilters]=useState({}); // {"id:tabIdx": {field:vals}}
  const [globalPivotFilters,setGlobalPivotFilters]=useState({}); // {"id:tabIdx": {rowIdx:[vals]}}
  const getFilterKey=(id,tabIdx)=>id+":"+tabIdx;
  const getStoredFilter=(id,tabIdx,meta)=>{
    const k=getFilterKey(id,tabIdx);
    if (globalFilters[k]!==undefined) return globalFilters[k];
    // Fall back to saved defaultFilters for this tab
    const cfg=meta&&meta.config;
    if (cfg&&cfg.tabs&&cfg.tabs[tabIdx]&&cfg.tabs[tabIdx].config&&cfg.tabs[tabIdx].config.defaultFilters)
      return cfg.tabs[tabIdx].config.defaultFilters;
    return (tabIdx===0&&cfg&&cfg.defaultFilters)||{};
  };
  const getStoredPivotFilter=(id,tabIdx,meta)=>{
    const k=getFilterKey(id,tabIdx);
    if (globalPivotFilters[k]!==undefined) return globalPivotFilters[k];
    const cfg=meta&&meta.config;
    if (cfg&&cfg.tabs&&cfg.tabs[tabIdx]&&cfg.tabs[tabIdx].config&&cfg.tabs[tabIdx].config.defaultPivotFilters)
      return cfg.tabs[tabIdx].config.defaultPivotFilters;
    return {};
  };
  const setStoredFilter=(id,tabIdx,val)=>setGlobalFilters(prev=>({...prev,[getFilterKey(id,tabIdx)]:val}));
  const setStoredPivotFilter=(id,tabIdx,val)=>setGlobalPivotFilters(prev=>({...prev,[getFilterKey(id,tabIdx)]:val}));

  const publishedReports=useMemo(()=>savedReports.filter(r=>r.isPublished),[savedReports]);
  const currentMeta=useMemo(()=>{
    if (activeId) return savedReports.find(r=>r.id===activeId)||publishedReports[0]||null;
    return publishedReports[0]||null;
  },[activeId,savedReports,publishedReports]);

  useEffect(()=>{ setUserActiveTabIdx(0); },[currentMeta?.id]);

  // Load data when report changes — show cache instantly, fetch DB in background
  // Also re-fires when row_count changes (admin saved new data) so preview updates.
  useEffect(()=>{
    if (!currentMeta) return;
    const id=currentMeta.id;
    const cached=loadedData[id];
    if (!cached){
      // No cache — show spinner, load from DB
      setDataLoading(true);
      onLoadReportData(id)
        .then(data=>{
          // Only update numFieldsRef if we got a non-empty set
          if(data.numFields instanceof Set && data.numFields.size>0) numFieldsRef.current=data.numFields;
          // Never overwrite a valid numFields with an empty one (guards against null num_fields in DB)
          const safeData = (data.numFields instanceof Set && data.numFields.size===0 && numFieldsRef.current.size>0)
            ? {...data, numFields: numFieldsRef.current}
            : data;
          setLoadedData(p=>({...p,[id]:safeData}));
          saveCache(id,safeData);
        })
        .catch(e=>console.error("Load error",e))
        .finally(()=>setDataLoading(false));
    } else {
      // Cache hit — display immediately, silently refresh from DB in background
      if(cached.numFields instanceof Set && cached.numFields.size>0) numFieldsRef.current=cached.numFields;
      setDataLoading(false);
      onLoadReportData(id)
        .then(data=>{
          if(data.numFields instanceof Set && data.numFields.size>0) numFieldsRef.current=data.numFields;
          const safeData = (data.numFields instanceof Set && data.numFields.size===0 && numFieldsRef.current.size>0)
            ? {...data, numFields: numFieldsRef.current}
            : data;
          setLoadedData(p=>({...p,[id]:safeData}));
          saveCache(id,safeData);
        })
        .catch(()=>{/* silent background refresh failed — keep showing cache */});
    }
  // row_count in deps: handleSaveReport writes fresh data to dataCache + localStorage
  // and loadAllReports updates row_count → this effect re-fires → preview gets fresh data
  },[currentMeta?.id, currentMeta?.rows]);

  // Refresh from source URL — old data stays visible, spinner overlay only
  async function refreshFromSource(silent=false) {
    if (!currentMeta) return;
    const links = getSourceLinks(currentMeta?.config);
    if (!links.length) return;
    if (!silent) setRefreshing(true);
    setRefreshError("");
    try {
      const lk = links[0];
      const result = await fetchUrlViaProxy(lk.url, lk.sheet||undefined);
      // Use fields/numFields returned by the backend (parsed from actual Excel headers)
      // Backend normalizes column headers (trims whitespace) so they match stored config.
      const freshFields = result.fields || (result.rows.length ? Object.keys(result.rows[0]) : []);
      // [] is truthy — must check .length>0 or an empty backend array would replace
      // numFieldsRef with an empty Set, causing all values to show as 0.
      const backendNF2 = result.numFields&&result.numFields.length>0 ? new Set(result.numFields) : null;
      const freshNumFields = backendNF2 && backendNF2.size>0 ? backendNF2 : numFieldsRef.current;

      // Merge with config.values to ensure value fields are always treated as numeric
      // (covers currency-formatted cells the backend detection missed)
      const configValueFields = [];
      const mc = currentMeta?.config;
      if (mc) {
        (mc.values||[]).forEach(v=>v.field&&configValueFields.push(v.field));
        (mc.tabs||[]).forEach(t=>(t.config?.values||[]).forEach(v=>v.field&&configValueFields.push(v.field)));
      }
      configValueFields.forEach(f=>freshNumFields.add(f));

      if (freshFields.length === 0) {
        if (!silent) setRefreshError("Could not read columns from file.");
        setRefreshing(false);
        return;
      }
      numFieldsRef.current = freshNumFields;
      const newData = {rows:result.rows, fields:freshFields, numFields:freshNumFields};
      setLoadedData(p=>({...p,[currentMeta.id]:newData}));
      saveCache(currentMeta.id, newData); // persist so next page load is instant
      setLastRefreshed(new Date());
    } catch(e) {
      if (!silent) setRefreshError(e.message);
    } finally {
      setRefreshing(false);
    }
  }

  // ── Auto-refresh every N minutes when report has a source link ──────────────
  const AUTO_REFRESH_MS = 5 * 60 * 1000; // 5 minutes
  useEffect(()=>{
    if (autoRefreshRef.current) clearInterval(autoRefreshRef.current);
    const links = currentMeta&&getSourceLinks(currentMeta?.config);
    if (links.length===0) return;
    // Immediately do a silent refresh when report is selected (picks up latest without user action)
    const t = setTimeout(()=>refreshFromSource(true), 1500);
    // Then auto-refresh every 5 minutes
    autoRefreshRef.current = setInterval(()=>refreshFromSource(true), AUTO_REFRESH_MS);
    return()=>{clearTimeout(t); if(autoRefreshRef.current)clearInterval(autoRefreshRef.current);};
  },[currentMeta?.id]);

  const currentData=currentMeta?loadedData[currentMeta.id]:null;

  if (!publishedReports.length) return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:T.bgPage}}>
      <div style={{textAlign:"center"}}>
        <div style={{fontSize:44,marginBottom:14}}>📋</div>
        <div style={{fontWeight:700,fontSize:16,color:T.text,marginBottom:8}}>No published reports yet</div>
        <div style={{fontSize:13,color:T.textMd}}>Ask your admin to publish a report from the Reports tab.</div>
      </div>
    </div>
  );

  return(<>
    <div style={{minHeight:"100vh",background:T.bgPage,fontFamily:"system-ui,sans-serif"}}>
      <AppHeader role="User" onLogout={onLogout}>
        {publishedReports.length>0&&(
          <div style={{display:"flex",alignItems:"center",gap:6,flex:1,minWidth:0}}>
            <span style={{fontSize:11,color:"rgba(245,239,230,0.6)",flexShrink:0}}>Report:</span>
            <select value={activeId||publishedReports[0]?.id||""}
              onChange={e=>setActiveId(e.target.value)}
              style={{padding:"4px 8px",border:"1px solid rgba(255,255,255,0.25)",borderRadius:6,background:"rgba(255,255,255,0.1)",
                color:T.textLt,fontSize:12,cursor:"pointer",outline:"none",minWidth:0,flex:1,maxWidth:260}}>
              {publishedReports.map(r=>(
                <option key={r.id} value={r.id}>{r.name}</option>
              ))}
            </select>
          </div>
        )}
      </AppHeader>
      {dataLoading&&(
        <div style={{padding:"40px",textAlign:"center"}}>
          <div style={{fontSize:30,animation:"spin 1s linear infinite",display:"inline-block"}}>⚙️</div>
          <div style={{color:T.textMd,marginTop:10,fontSize:13}}>Loading report data…</div>
          <style>{"@keyframes spin{to{transform:rotate(360deg)}}"}</style>
        </div>
      )}
      {!dataLoading&&currentMeta&&currentData?(
        <div style={{padding:isMobileView?10:20}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:8}}>
            <div style={{display:"flex",alignItems:"baseline",gap:8,flexWrap:"wrap"}}>
              <div style={{fontWeight:700,fontSize:isMobileView?15:18,color:T.primary}}>{currentMeta.config&&currentMeta.config.name||currentMeta.name}</div>
              <span style={{fontSize:11,background:T.primary,color:T.textLt,padding:"2px 8px",borderRadius:10,fontWeight:600}}>Published</span>
              {getSourceLinks(currentMeta?.config).length>0&&(
                <span style={{fontSize:10,background:"rgba(0,100,200,0.09)",color:"#0064C8",
                  border:"1px solid rgba(0,100,200,0.22)",borderRadius:4,padding:"2px 8px",fontWeight:600}}>
                  {getSourceLabel(getSourceLinks(currentMeta?.config)[0]?.url)}
                  {getSourceLinks(currentMeta?.config)[0]?.sheet&&" · "+getSourceLinks(currentMeta?.config)[0]?.sheet}
                </span>
              )}
            </div>
            <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              <div style={{fontSize:11,color:T.textMd}}>
                {currentData.rows.length.toLocaleString()} records · {currentData.fields.length} fields
                {lastRefreshed&&<span style={{marginLeft:8,color:T.success}}>· Refreshed {lastRefreshed.toLocaleTimeString()}</span>}
              </div>
              {currentMeta.config?.collab_enabled&&(
                <button onClick={()=>setCollabViewReport(currentMeta)}
                  style={{padding:"6px 14px",background:"#3B5998",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontSize:12,fontWeight:600}}>
                  🤝 Workflow
                </button>
              )}
              {(getSourceLinks(currentMeta?.config).length>0)&&(
                <button onClick={()=>refreshFromSource(false)} disabled={refreshing}
                  title="Pull latest data from Google Drive / OneDrive"
                  style={{display:"flex",alignItems:"center",gap:6,padding:"6px 14px",
                    background:refreshing?"rgba(92,45,26,0.08)":T.primary,
                    color:refreshing?T.primary:T.textLt,
                    border:"1px solid "+T.primary,borderRadius:7,
                    cursor:refreshing?"not-allowed":"pointer",
                    fontSize:12,fontWeight:600,transition:"all 0.2s"}}>
                  <span style={{display:"inline-block",animation:refreshing?"spin 0.8s linear infinite":"none",fontSize:14}}>↻</span>
                  {refreshing?"Updating...":"Refresh"}
                </button>
              )}
            </div>
          </div>
          {refreshError&&(
            <div style={{padding:"8px 12px",background:"rgba(163,45,45,0.07)",border:"1px solid rgba(163,45,45,0.25)",
              borderRadius:7,fontSize:12,color:"#A32D2D",marginBottom:10,display:"flex",alignItems:"center",gap:8}}>
              <span>⚠</span><span>{refreshError}</span>
              <button onClick={()=>setRefreshError("")} style={{marginLeft:"auto",background:"none",border:"none",cursor:"pointer",color:"#A32D2D",fontSize:14}}>×</button>
            </div>
          )}
          <div style={{fontSize:12,color:T.textMd,marginBottom:14}}>Click cells to drill down</div>
          <Report
            key={currentMeta.id}
            externalFilters={getStoredFilter(currentMeta.id,userActiveTabIdx,currentMeta)}
            externalPivotFilters={getStoredPivotFilter(currentMeta.id,userActiveTabIdx,currentMeta)}
            onExternalFiltersChange={(f)=>setStoredFilter(currentMeta.id,userActiveTabIdx,f)}
            onExternalPivotFiltersChange={(pf)=>setStoredPivotFilter(currentMeta.id,userActiveTabIdx,pf)}
            config={(()=>{
              const topCfg=currentMeta.config||{};
              const tabCfg=topCfg.tabs&&topCfg.tabs[userActiveTabIdx];
              if (tabCfg) {
                // EXPLICIT whitelist of report-wide (non-structural) fields from topCfg.
                // Structural fields (rows/columns/values/filters etc.) come ONLY from
                // tabCfg.config — never from the top level. This is the only safe approach.
                const reportWide = {
                  allowedFmts:    topCfg.allowedFmts,
                  drillHiddenCols:topCfg.drillHiddenCols,
                  drillColFmts:   topCfg.drillColFmts,
                  name:           topCfg.name,
                };
                const tabStructural = tabCfg.config || {};
                const merged = {
                  ...reportWide,
                  ...tabStructural,
                  // defaultFilters: tab-level wins; top-level fallback for old data
                  defaultFilters: tabStructural.defaultFilters != null
                    ? tabStructural.defaultFilters
                    : (topCfg.defaultFilters || {}),
                  defaultPivotFilters: tabStructural.defaultPivotFilters
                    ? {[userActiveTabIdx]: tabStructural.defaultPivotFilters}
                    : (topCfg.defaultPivotFilters || {}),
                  colExcluded: tabStructural.colExcluded || [],
                  _reportId: currentMeta.id,
                };
                return merged;
              }
              return {...topCfg,_reportId:currentMeta.id};
            })()}
            data={currentData.rows}
            fields={currentData.fields}
            numFields={currentData.numFields}
            showExport
            cardFields={currentMeta.config&&currentMeta.config.tabs&&currentMeta.config.tabs[userActiveTabIdx]
              ? (currentMeta.config.tabs[userActiveTabIdx].cardFields||[])
              : (currentMeta.cardFields||[])}
            tabs={currentMeta.config&&currentMeta.config.tabs||null}
            activeTabIdx={userActiveTabIdx}
            onTabChange={setUserActiveTabIdx}
            onDrillHiddenColsChange={(cols,fmts)=>{
              try{
                localStorage.setItem("rh_drill_cols_"+currentMeta.id,JSON.stringify(cols));
                if(fmts)localStorage.setItem("rh_drill_fmts_"+currentMeta.id,JSON.stringify(fmts));
              }catch(e){}
            }}/>
        </div>
      ):(!dataLoading&&<div style={{padding:40,textAlign:"center",fontSize:13,color:T.textMd}}>Select a report above.</div>)}
    </div>
    {collabViewReport&&<CollabDataView report={collabViewReport} currentUser={currentUser} currentRole={currentRole||"user"} onClose={()=>setCollabViewReport(null)}/>}
  </>);
}


// ── Settings / User Management ────────────────────────────────────────────────

// ── Report Access Manager — admin assigns users to a report ──────────────────
function ReportAccessPanel({reportId, reportName, onClose}) {
  const [users,setUsers]=useState([]);
  const [loading,setLoading]=useState(true);
  const [saving,setSaving]=useState(false);
  const [selected,setSelected]=useState(new Set());
  const [msg,setMsg]=useState("");

  useEffect(()=>{
    getReportAccess(reportId)
      .then(rows=>{
        setUsers(rows);
        setSelected(new Set(rows.filter(u=>u.has_access).map(u=>u.id)));
      })
      .catch(e=>setMsg("Load failed: "+e.message))
      .finally(()=>setLoading(false));
  },[reportId]);

  async function save() {
    setSaving(true); setMsg("");
    try {
      await setReportAccess(reportId,[...selected]);
      setMsg("✅ Access saved");
      setTimeout(()=>onClose(),800);
    } catch(e) { setMsg("❌ "+e.message); }
    finally { setSaving(false); }
  }

  const toggle=(id)=>setSelected(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n;});

  return(
    <div style={{position:"fixed",inset:0,zIndex:900,background:"rgba(44,24,16,0.55)",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:T.bgCard,borderRadius:12,width:"min(480px,92vw)",maxHeight:"80vh",display:"flex",flexDirection:"column",boxShadow:"0 12px 40px rgba(44,24,16,0.3)"}}>
        <div style={{padding:"16px 20px",background:T.bgHeader,borderRadius:"12px 12px 0 0",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div>
            <div style={{fontWeight:700,fontSize:14,color:T.textLt}}>Manage User Access</div>
            <div style={{fontSize:11,color:"rgba(245,239,230,0.65)",marginTop:2}}>{reportName}</div>
          </div>
          <button onClick={onClose} style={{border:"none",background:"rgba(255,255,255,0.15)",color:T.textLt,borderRadius:6,width:28,height:28,cursor:"pointer",fontSize:16}}>×</button>
        </div>
        <div style={{padding:"10px 16px",background:T.bgStat,borderBottom:"0.5px solid "+T.border,fontSize:11,color:T.textMd}}>
          Tick users who can view this report. Unticked users will see "no reports available" if they have no other assigned reports.
        </div>
        <div style={{overflowY:"auto",padding:"10px 14px",flex:1}}>
          {loading && <div style={{textAlign:"center",padding:20,color:T.textMd}}>Loading users…</div>}
          {!loading && users.length===0 && <div style={{textAlign:"center",padding:20,color:T.textMd}}>No regular users yet. Create users in Settings first.</div>}
          {!loading && users.length>0 && (
            <>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                <span style={{fontSize:11,fontWeight:600,color:T.textMd}}>{selected.size} of {users.length} users selected</span>
                <div style={{display:"flex",gap:10}}>
                  <button onClick={()=>setSelected(new Set(users.map(u=>u.id)))}
                    style={{fontSize:11,color:T.primary,background:"none",border:"none",cursor:"pointer",fontWeight:600}}>All</button>
                  <button onClick={()=>setSelected(new Set())}
                    style={{fontSize:11,color:T.textMd,background:"none",border:"none",cursor:"pointer"}}>None</button>
                </div>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:4}}>
                {users.map(u=>(
                  <label key={u.id} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",
                    borderRadius:7,border:"1px solid "+(selected.has(u.id)?T.primary:T.border),
                    cursor:"pointer",background:selected.has(u.id)?"rgba(92,45,26,0.05)":"none"}}>
                    <input type="checkbox" checked={selected.has(u.id)} onChange={()=>toggle(u.id)}
                      style={{width:15,height:15,accentColor:T.primary,cursor:"pointer"}}/>
                    <span style={{fontWeight:600,fontSize:13,color:T.text,flex:1}}>{u.username}</span>
                    {selected.has(u.id)
                      ? <span style={{fontSize:10,color:T.success,fontWeight:600}}>✓ Has access</span>
                      : <span style={{fontSize:10,color:T.textMd}}>No access</span>}
                  </label>
                ))}
              </div>
            </>
          )}
        </div>
        {msg&&<div style={{padding:"8px 16px",fontSize:12,color:msg.startsWith("✅")?T.success:"#A32D2D",
          background:msg.startsWith("✅")?"rgba(45,106,79,0.07)":"rgba(163,45,45,0.07)",
          borderTop:"0.5px solid "+T.border}}>{msg}</div>}
        <div style={{padding:"12px 16px",borderTop:"0.5px solid "+T.border,display:"flex",gap:8,justifyContent:"flex-end"}}>
          <button onClick={onClose} style={{padding:"7px 16px",background:"none",border:"1px solid "+T.border,borderRadius:6,cursor:"pointer",fontSize:13,color:T.text}}>Cancel</button>
          <button onClick={save} disabled={saving||loading}
            style={{padding:"7px 18px",background:T.primary,color:T.textLt,border:"none",borderRadius:6,
              cursor:saving||loading?"not-allowed":"pointer",fontSize:13,fontWeight:700,opacity:saving||loading?0.6:1}}>
            {saving?"Saving…":"Save access"}
          </button>
        </div>
      </div>
    </div>
  );
}

function SettingsPanel({currentUser,currentRole,onClose}) {
  const isSuperAdmin = currentRole==="admin";
  const [users,setUsers]=useState([]);
  const [pwdEdits,setPwdEdits]=useState({}); // {id: newPassword}
  const [roleEdits,setRoleEdits]=useState({}); // {id: newRole}
  const [newUser,setNewUser]=useState({username:"",password:"",role:"user"});
  const [toast,setToast]=useState("");
  const [loading,setLoading]=useState(false);
  const showToast=msg=>{setToast(msg);setTimeout(()=>setToast(""),3000);};

  // Load users from API on mount
  useEffect(()=>{
    getUsers().then(setUsers).catch(e=>showToast("Load failed: "+e.message));
  },[]);

  async function addUser(){
    if (!newUser.username.trim()||!newUser.password.trim()){showToast("Username and password required.");return;}
    setLoading(true);
    try{
      const u=await createUser(newUser.username.trim(),newUser.password,newUser.role);
      setUsers(p=>[...p,u]);
      setNewUser({username:"",password:"",role:"user"});
      showToast("User created!");
    }catch(e){showToast(e.message||"Create failed.");}
    finally{setLoading(false);}
  }

  async function savePwd(id){
    const pwd=pwdEdits[id]||"";
    if (!pwd){showToast("Enter a new password first.");return;}
    setLoading(true);
    try{
      await updatePassword(id,pwd);
      setPwdEdits(p=>{const n={...p};delete n[id];return n;});
      showToast("Password updated!");
    }catch(e){showToast(e.message||"Update failed.");}
    finally{setLoading(false);}
  }

  async function doApprove(id){
    setLoading(true);
    try{
      await approveUser(id);
      setUsers(p=>p.map(u=>u.id===id?{...u,status:"active"}:u));
      showToast("User approved and can now log in!");
    }catch(e){showToast(e.message||"Approve failed.");}
    finally{setLoading(false);}
  }

  async function saveRole(id){
    const role=roleEdits[id];
    if (!role){showToast("Select a new role first.");return;}
    setLoading(true);
    try{
      await updateRole(id,role);
      setUsers(p=>p.map(u=>u.id===id?{...u,role}:u));
      setRoleEdits(p=>{const n={...p};delete n[id];return n;});
      showToast("Role updated!");
    }catch(e){showToast(e.message||"Update failed.");}
    finally{setLoading(false);}
  }

  async function delUser(id){
    if (users.find(u=>u.id===id)?.username===currentUser){showToast("Cannot delete your own account.");return;}
    if (!confirm("Delete this user?")) return;
    setLoading(true);
    try{
      await deleteUser(id);
      setUsers(p=>p.filter(u=>u.id!==id));
      showToast("User deleted.");
    }catch(e){showToast(e.message||"Delete failed.");}
    finally{setLoading(false);}
  }

  const inp={padding:"7px 10px",border:"1px solid "+T.border,borderRadius:6,fontSize:12,background:T.bgCard,color:T.text,outline:"none",width:"100%",boxSizing:"border-box"};
  return(
    <div style={{position:"fixed",inset:0,zIndex:600,background:"rgba(44,24,16,0.55)",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:T.bgCard,borderRadius:14,width:"min(600px,95vw)",maxHeight:"85vh",display:"flex",flexDirection:"column",boxShadow:"0 12px 48px rgba(44,24,16,0.35)"}}>
        <div style={{padding:"16px 20px",background:T.bgHeader,borderRadius:"14px 14px 0 0",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <span style={{fontWeight:700,fontSize:16,color:T.textLt}}>⚙ Settings — User Management</span>
          <button onClick={onClose} style={{border:"none",background:"rgba(255,255,255,0.15)",color:T.textLt,borderRadius:6,width:28,height:28,cursor:"pointer",fontSize:16}}>×</button>
        </div>
        <div style={{padding:20,overflowY:"auto",flex:1}}>
          {toast&&<div style={{padding:"8px 14px",background:"rgba(45,106,79,0.15)",border:"1px solid rgba(45,106,79,0.4)",borderRadius:7,fontSize:12,color:T.success,marginBottom:14}}>{toast}</div>}
          {/* Pending approval — visible to Super Admin */}
          {isSuperAdmin&&users.filter(u=>u.status==="pending").length>0&&(
            <div style={{marginBottom:20}}>
              <div style={{fontWeight:700,fontSize:13,color:T.warning,marginBottom:8}}>
                ⏳ Pending Approval ({users.filter(u=>u.status==="pending").length})
                <span style={{fontWeight:400,fontSize:11,color:T.textMd,marginLeft:8}}>Created by Sub-Admins — approve to allow login</span>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:8}}>
                {users.filter(u=>u.status==="pending").map(u=>(
                  <div key={u.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",
                    background:"rgba(186,117,23,0.07)",borderRadius:8,border:"1px solid rgba(186,117,23,0.3)"}}>
                    <div style={{width:32,height:32,borderRadius:8,background:T.warning,display:"flex",
                      alignItems:"center",justifyContent:"center",fontSize:12,color:"#fff",fontWeight:700,flexShrink:0}}>?</div>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontWeight:600,fontSize:13,color:T.text}}>{u.username}</div>
                      <div style={{fontSize:11,color:T.textMd}}>{u.role} · pending approval</div>
                    </div>
                    <button onClick={()=>doApprove(u.id)} disabled={loading}
                      style={{padding:"6px 14px",background:T.success,color:"#fff",border:"none",borderRadius:6,
                        cursor:"pointer",fontSize:12,fontWeight:700,flexShrink:0}}>
                      ✓ Approve
                    </button>
                    <button onClick={()=>delUser(u.id)} disabled={loading}
                      style={{padding:"6px 10px",border:"1px solid rgba(163,45,45,0.4)",borderRadius:6,
                        background:"none",cursor:"pointer",fontSize:11,color:T.danger,flexShrink:0}}>
                      Reject
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}
          {/* Existing active users */}
          <div style={{fontWeight:700,fontSize:13,color:T.primary,marginBottom:12}}>Users ({users.filter(u=>u.status!=="pending").length} active{!isSuperAdmin&&users.filter(u=>u.status==="pending").length>0?", "+users.filter(u=>u.status==="pending").length+" pending approval":""})</div>
          <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:20}}>
            {(!isSuperAdmin?users:users.filter(u=>u.status!=="pending")).map(u=>(
              <div key={u.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",background:T.bgStat,borderRadius:8,border:"1px solid "+T.border}}>
                <div style={{width:32,height:32,borderRadius:8,
                  background:u.role==="admin"?T.primary:u.role==="subadmin"?T.accent:u.role==="subadmin_user"?"#6B5B95":T.secondary,
                  display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:T.textLt,fontWeight:700,flexShrink:0}}>
                  {u.role==="admin"?"SA":u.role==="subadmin"?"Sub":u.role==="subadmin_user"?"S+U":"U"}
                </div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:600,fontSize:13,color:T.text}}>{u.username} {u.username===currentUser&&<span style={{fontSize:10,color:T.textMd}}>(you)</span>}</div>
                  <div style={{fontSize:11,color:u.status==="pending"?T.warning:T.textMd}}>
                    {u.role}{u.status==="pending"&&<span style={{marginLeft:6,fontWeight:600}}>· ⏳ pending approval</span>}
                  </div>
                </div>
                {isSuperAdmin&&u.status!=="pending"&&<>
                  <input type="password" value={pwdEdits[u.id]||""} onChange={e=>setPwdEdits(p=>({...p,[u.id]:e.target.value}))}
                    placeholder="New password" title="Change password"
                    style={{...inp,width:130,flexShrink:0}}/>
                  {pwdEdits[u.id]&&<button onClick={()=>savePwd(u.id)} disabled={loading}
                    style={{padding:"5px 8px",border:"1px solid "+T.primary,borderRadius:6,background:T.primary,cursor:"pointer",fontSize:11,color:T.textLt,flexShrink:0,fontWeight:600}}>
                    Save pwd
                  </button>}
                  {u.username!==currentUser&&(<>
                    <select value={roleEdits[u.id]||u.role}
                      onChange={e=>setRoleEdits(p=>({...p,[u.id]:e.target.value}))}
                      style={{padding:"4px 6px",border:"1px solid "+T.border,borderRadius:6,fontSize:11,background:T.bgCard,color:T.text,cursor:"pointer",flexShrink:0}}>
                      <option value="user">User</option>
                      <option value="subadmin_user">Sub-Admin + User</option>
                      <option value="subadmin">Sub-Admin</option>
                      <option value="admin">Super Admin</option>
                    </select>
                    {roleEdits[u.id]&&roleEdits[u.id]!==u.role&&(
                      <button onClick={()=>saveRole(u.id)} disabled={loading}
                        style={{padding:"5px 8px",border:"1px solid "+T.primary,borderRadius:6,background:T.primary,cursor:"pointer",fontSize:11,color:T.textLt,flexShrink:0,fontWeight:600}}>
                        Save role
                      </button>
                    )}
                    <button onClick={()=>delUser(u.id)} disabled={loading} style={{padding:"5px 10px",border:"1px solid rgba(163,45,45,0.4)",borderRadius:6,background:"none",cursor:"pointer",fontSize:11,color:T.danger,flexShrink:0}}>
                      Delete
                    </button>
                  </>)}
                </>}
              </div>
            ))}
          </div>
          {/* Add new user */}
          <div style={{fontWeight:700,fontSize:13,color:T.primary,marginBottom:10}}>Add new user{!isSuperAdmin&&<span style={{fontWeight:400,fontSize:11,color:T.textMd,marginLeft:8}}>(requires Super Admin approval)</span>}</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr auto auto",gap:8,alignItems:"end"}}>
            <div>
              <div style={{fontSize:11,color:T.textMd,marginBottom:4}}>Username</div>
              <input value={newUser.username} onChange={e=>setNewUser(p=>({...p,username:e.target.value}))} placeholder="username" style={inp}/>
            </div>
            <div>
              <div style={{fontSize:11,color:T.textMd,marginBottom:4}}>Password</div>
              <input type="password" value={newUser.password} onChange={e=>setNewUser(p=>({...p,password:e.target.value}))} placeholder="password" style={inp}/>
            </div>
            <div>
              <div style={{fontSize:11,color:T.textMd,marginBottom:4}}>Role</div>
              <select value={newUser.role} onChange={e=>setNewUser(p=>({...p,role:e.target.value}))}
                disabled={!isSuperAdmin}
                style={{...inp,width:"auto",cursor:"pointer"}}>
                <option value="user">User</option>
                {isSuperAdmin&&<option value="subadmin_user">Sub-Admin + User</option>}
                {isSuperAdmin&&<option value="subadmin">Sub-Admin</option>}
                {isSuperAdmin&&<option value="admin">Super Admin</option>}
              </select>
            </div>
            <button onClick={addUser} disabled={loading} style={{padding:"8px 16px",background:T.primary,color:T.textLt,border:"none",borderRadius:6,cursor:loading?"wait":"pointer",fontSize:12,fontWeight:700,alignSelf:"end",opacity:loading?0.6:1}}>
              {loading?"…":"Add"}
            </button>
          </div>
        </div>
        <div style={{padding:"12px 20px",borderTop:"0.5px solid "+T.border,display:"flex",justifyContent:"flex-end",gap:10}}>
          <button onClick={onClose} style={{padding:"7px 18px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,cursor:"pointer",fontSize:13,fontWeight:700}}>Done</button>
        </div>
      </div>
    </div>
  );
}

// ── Login ──────────────────────────────────────────────────────────────────────
function Login({onLogin}) {
  const [username,setUsername]=useState("");
  const [password,setPassword]=useState("");
  const [showPwd,setShowPwd]=useState(false);
  const [err,setErr]=useState("");
  const [loading,setLoading]=useState(false);
  const [waking,setWaking]=useState(false); // Railway wake-up ping
  const inp={width:"100%",padding:"10px 12px",border:"1px solid "+T.border,borderRadius:8,fontSize:14,background:T.bgCard,color:T.text,boxSizing:"border-box",outline:"none"};

  // Ping backend on mount to wake Railway from sleep (free tier spins down)
  useEffect(()=>{
    setWaking(true);
    const base=(typeof BACKEND_URL!=="undefined"?BACKEND_URL:"");
    fetch(base+"/health",{method:"GET",mode:"cors"})
      .catch(()=>{}) // silent — just waking the server up
      .finally(()=>setWaking(false));
  },[]);

  function isNetworkErr(msg) {
    // iOS Safari: "Load failed" · Chrome: "Failed to fetch" · Firefox: "NetworkError"
    return msg.includes("Load failed")||msg.includes("Failed to fetch")||
           msg.includes("NetworkError")||msg.includes("ERR_NETWORK")||
           msg.includes("Network request failed")||msg.includes("The network connection was lost")||
           msg.includes("Could not connect")||msg.includes("fetch")||msg.includes("ECONNREFUSED");
  }

  async function tryLogin(){
    if (!username.trim()||!password){setErr("Enter username and password.");return;}
    setLoading(true);setErr("");
    try{
      const data=await apiLogin(username.trim(),password);
      onLogin(data.role,data.username,data.token,data.id);
    }catch(e){
      const msg=e.message||"";
      if (isNetworkErr(msg))
        setErr("Cannot reach the server. This may be a temporary issue — please wait 10 seconds and try again. If the problem persists, check your internet connection.");
      else if (msg.includes("401")||msg.includes("Invalid")||msg.includes("credentials")||msg.includes("password"))
        setErr("Wrong username or password. Please check and try again.");
      else
        setErr(msg||"Login failed. Please try again.");
    }finally{setLoading(false);}
  }
  return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:T.bgPage,fontFamily:"system-ui,sans-serif",padding:"20px"}}>
      <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:16,padding:"40px 32px",width:"100%",maxWidth:380,boxShadow:"0 4px 24px rgba(92,45,26,0.15)"}}>
        <div style={{textAlign:"center",marginBottom:28}}>
          <div style={{width:60,height:60,background:T.primary,borderRadius:14,display:"flex",alignItems:"center",justifyContent:"center",fontSize:28,margin:"0 auto 16px"}}>📊</div>
          <h2 style={{fontSize:24,fontWeight:800,margin:"0",color:T.primary,letterSpacing:"-0.5px"}}>ReportHub</h2>
        </div>
        {waking&&!err&&<div style={{padding:"7px 12px",background:"rgba(200,146,42,0.1)",border:"1px solid rgba(200,146,42,0.3)",borderRadius:8,fontSize:11,color:T.accent,marginBottom:12,textAlign:"center"}}>
          Connecting to server…
        </div>}
        {err&&<div style={{padding:"10px 14px",background:"rgba(163,45,45,0.09)",border:"1px solid rgba(163,45,45,0.3)",borderRadius:8,fontSize:12,color:T.danger,marginBottom:14,lineHeight:1.5}}>{err}</div>}
        <div style={{display:"flex",flexDirection:"column",gap:12,marginBottom:18}}>
          <div>
            <div style={{fontSize:11,color:T.textMd,fontWeight:600,marginBottom:5}}>Username</div>
            <input value={username} onChange={e=>setUsername(e.target.value)} placeholder="Enter username"
              autoComplete="username"
              style={inp} onKeyDown={e=>e.key==="Enter"&&tryLogin()}/>
          </div>
          <div>
            <div style={{fontSize:11,color:T.textMd,fontWeight:600,marginBottom:5}}>Password</div>
            <div style={{position:"relative"}}>
              <input type={showPwd?"text":"password"} value={password} onChange={e=>setPassword(e.target.value)} placeholder="Enter password"
                autoComplete="current-password"
                style={{...inp,paddingRight:44}} onKeyDown={e=>e.key==="Enter"&&tryLogin()}/>
              <button type="button" onClick={()=>setShowPwd(v=>!v)}
                title={showPwd?"Hide password":"Show password"}
                style={{position:"absolute",right:8,top:"50%",transform:"translateY(-50%)",
                  background:"none",border:"none",cursor:"pointer",fontSize:14,color:T.textMd,padding:"4px 6px"}}>
                {showPwd?"🙈":"👁"}
              </button>
            </div>
          </div>
        </div>
        <button onClick={tryLogin} disabled={loading} style={{width:"100%",padding:"11px",background:loading?"rgba(92,45,26,0.5)":T.primary,color:T.textLt,border:"none",borderRadius:8,cursor:loading?"wait":"pointer",fontSize:14,fontWeight:700,letterSpacing:"0.3px"}}>
          {loading?"Signing in…":"Sign in"}
        </button>
      </div>
    </div>
  );
}

// ── Reports Manager (Admin tab) ────────────────────────────────────────────────
function ReportsTab({savedReports,onOpen,onDelete,onPublish,onUnpublish,publishedId,onReload,onAccessPanel}) {
  const [schedulePanel,setSchedulePanel]=useState(null); // {id,name}
  if (!savedReports.length) return(
    <div style={{padding:40,textAlign:"center"}}>
      <div style={{fontSize:40,marginBottom:14}}>📋</div>
      <div style={{fontWeight:700,fontSize:16,color:T.primary,marginBottom:8}}>No saved reports yet</div>
      <div style={{fontSize:13,color:T.textMd}}>Go to the Builder tab, configure your pivot, then click "Save Report".</div>
    </div>
  );
  return(<>
    <div style={{padding:20,maxWidth:900,margin:"0 auto"}}>
      <div style={{fontWeight:700,fontSize:16,color:T.primary,marginBottom:4}}>Saved Reports</div>
      <div style={{fontSize:12,color:T.textMd,marginBottom:18}}>
        {savedReports.length} report{savedReports.length!==1?"s":""} saved · publish one to make it visible to users
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {savedReports.map(r=>(
          <div key={r.id} style={{background:T.bgCard,border:"1px solid "+(r.isPublished?T.primary:T.border),borderRadius:10,padding:"14px 18px",display:"flex",alignItems:"center",gap:14}}>
            <div style={{width:40,height:40,background:r.isPublished?T.primary:T.bgStat,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>
              {r.isPublished?"📤":"📊"}
            </div>
            <div style={{flex:1,minWidth:0}}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:2,flexWrap:"wrap"}}>
                <span style={{fontWeight:700,fontSize:14,color:T.text}}>{r.name}</span>
                {getSourceLinks(r.config).length>0&&(
                  <span title={getSourceLinks(r.config)[0]?.url}
                    style={{fontSize:10,background:"rgba(0,100,200,0.09)",color:"#0064C8",
                      border:"1px solid rgba(0,100,200,0.22)",borderRadius:4,padding:"1px 6px",fontWeight:600,whiteSpace:"nowrap"}}>
                    {getSourceLabel(getSourceLinks(r.config)[0]?.url)}
                    {getSourceLinks(r.config)[0]?.sheet&&" · "+getSourceLinks(r.config)[0]?.sheet}
                    {getSourceLinks(r.config)[0]?.lastRefreshed
                      &&" · "+new Date(getSourceLinks(r.config)[0]?.lastRefreshed).toLocaleDateString()}
                  </span>
                )}
              </div>
              <div style={{fontSize:11,color:T.textMd,display:"flex",gap:12,flexWrap:"wrap"}}>
                <span>{r.rows.toLocaleString()} rows</span>
                <span>{r.fields} fields</span>
                <span>Rows: {(r.config&&r.config.rows||[]).join(", ")||"—"}</span>
                <span>Values: {(r.config&&r.config?.values||[]).map(v=>v.field).join(", ")||"—"}</span>
                <span>Saved: {new Date(r.savedAt).toLocaleDateString()}</span>
                {r.createdBy&&<span style={{color:T.accent,fontWeight:600}}>By: {r.createdBy}</span>}
              </div>
            </div>
            <div style={{display:"flex",gap:8,flexShrink:0}}>
              <button onClick={()=>onOpen(r.id)}
                style={{padding:"5px 13px",border:"1px solid "+T.border,borderRadius:6,background:"none",cursor:"pointer",fontSize:12,color:T.text,fontWeight:500}}>
                Open
              </button>
              {r.isPublished
                ? <button onClick={async()=>await onUnpublish(r.id)}
                    style={{padding:"5px 13px",border:"1px solid "+T.primary,borderRadius:6,
                      background:T.primary,cursor:"pointer",fontSize:12,color:T.textLt,fontWeight:700}}
                    title="Click to unpublish">
                    ✓ Published
                  </button>
                : <button onClick={async()=>await onPublish(r.id)}
                    style={{padding:"5px 13px",border:"1px solid "+T.border,borderRadius:6,
                      background:"none",cursor:"pointer",fontSize:12,color:T.text}}
                    title="Publish to users">
                    Publish
                  </button>
              }
              <button onClick={()=>onAccessPanel&&onAccessPanel(r.id,r.name)}
                title="Manage which users can view this report"
                style={{padding:"5px 10px",border:"1px solid "+T.border,borderRadius:6,background:"none",cursor:"pointer",fontSize:12,color:T.textMd}}>
                👥
              </button>
              {r.config?.sourceLinks?.length>0&&(
                <button onClick={()=>setSchedulePanel({id:r.id,name:r.name})}
                  title="Set auto-refresh schedule"
                  style={{padding:"5px 10px",border:"1px solid "+T.border,borderRadius:6,background:"none",cursor:"pointer",fontSize:12,color:T.textMd}}>
                  ⏱
                </button>
              )}
              <button onClick={async()=>{if(confirm("Delete report '"+r.name+"'?")) await onDelete(r.id);}}
                style={{padding:"5px 10px",border:"1px solid rgba(163,45,45,0.3)",borderRadius:6,background:"none",cursor:"pointer",fontSize:12,color:T.danger}}>
                Delete
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
    {schedulePanel&&<AutoRefreshPanel reportId={schedulePanel.id} reportName={schedulePanel.name} onClose={()=>setSchedulePanel(null)}/>}
  </>);
}

// ── Root ───────────────────────────────────────────────────────────────────────
// ── Helper: parse API report metadata into local shape ────────────────────────
function parseReportMeta(r) {
  const cfg = typeof r.config==="string" ? JSON.parse(r.config) : (r.config||{});
  // Merge top-level collab_enabled DB column into config so WorkflowListTab can read it
  if (r.collab_enabled !== undefined) cfg.collab_enabled = !!r.collab_enabled;
  return {
    id: r.id,
    name: r.name,
    rows: r.row_count||0,
    fields: r.field_count||0,
    row_count: r.row_count||0,
    savedAt: r.created_at ? new Date(r.created_at).getTime() : Date.now(),
    config: cfg,
    cardFields: (()=>{const cf=typeof r.card_fields==="string"?JSON.parse(r.card_fields):(r.card_fields||[]);
      // Normalise: legacy data may have strings, new data has {field,agg} objects
      return cf.map(x=>typeof x==="string"?{field:x,agg:"sum"}:x);
    })(),
    isPublished: !!r.is_published,
    createdBy: r.created_by_username||r.created_by||null,
    dataset: null, // rows loaded lazily on demand
  };
}

// ── Error Boundary — catches render crashes and shows a recovery screen ───────
class ErrorBoundary extends React.Component {
  constructor(props){super(props);this.state={crashed:false,err:null};}
  static getDerivedStateFromError(err){return{crashed:true,err};}
  componentDidCatch(err,info){console.error("ErrorBoundary caught:",err,info);}
  handleReset(){
    localStorage.removeItem("rh_token");
    localStorage.removeItem("rh_role");
    localStorage.removeItem("rh_username");
    this.setState({crashed:false,err:null});
    window.location.reload();
  }
  render(){
    if(!this.state.crashed)return this.props.children;
    return(
      <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"#faf7f2",fontFamily:"system-ui,sans-serif"}}>
        <div style={{textAlign:"center",maxWidth:400,padding:32}}>
          <div style={{fontSize:48,marginBottom:16}}>⚠️</div>
          <div style={{fontWeight:700,fontSize:18,color:"#2d1a0e",marginBottom:8}}>Something went wrong</div>
          <div style={{fontSize:13,color:"#7a6652",marginBottom:8}}>{this.state.err&&this.state.err.message}</div>
          <div style={{fontSize:12,color:"#a08060",marginBottom:24}}>Your session has been cleared. Please log in again.</div>
          <button onClick={()=>this.handleReset()} style={{padding:"10px 28px",background:"#5c2d1a",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontWeight:600,fontSize:14}}>
            Return to Login
          </button>
        </div>
      </div>
    );
  }
}

export default function App() {
  const [screen,setScreen]=useState("loading"); // loading|login|admin|user
  const [savedReports,setSavedReports]=useState([]);
  const [publishedId,setPublishedId]=useState(null);
  const [currentUser,setCurrentUser]=useState(null);
  const [loadErr,setLoadErr]=useState("");

  // dataCache stores {id -> {rows,fields,numFields}} so we don't re-fetch
  const dataCache=useRef({});

  const publishedReport=useMemo(()=>savedReports.find(r=>r.id===publishedId)||null,[savedReports,publishedId]);

  // ── Restore session from localStorage on mount ─────────────────────────────
  useEffect(()=>{
    const token=localStorage.getItem("rh_token");
    const role=localStorage.getItem("rh_role");
    const username=localStorage.getItem("rh_username");

    function clearAndShowLogin(){
      localStorage.removeItem("rh_token");
      localStorage.removeItem("rh_role");
      localStorage.removeItem("rh_username");
      setCurrentUser(null);
      setScreen("login");
    }

    // Any partial/corrupt state → clear and show login immediately
    if (!token||!role||!username) {
      clearAndShowLogin();
      return;
    }

    setCurrentUser({username, id: localStorage.getItem("rh_user_id")||null, role: localStorage.getItem("rh_role")});
    // Timeout: if API hangs for >8s, bail to login
    const timeout=setTimeout(clearAndShowLogin,8000);
    loadAllReports()
      .then(()=>{
        clearTimeout(timeout);
        // subadmin uses the same AdminView as admin
        setScreen(role==="subadmin"||role==="subadmin_user"?"admin":role);
      })
      .catch(()=>{
        clearTimeout(timeout);
        clearAndShowLogin();
      });
  },[]);

  // ── Load report list from API ──────────────────────────────────────────────
  async function loadAllReports() {
    // NOTE: do NOT catch here — let caller handle auth errors
    const list=await getReports();
    const entries=list.map(parseReportMeta);
    setSavedReports(entries);
    const pub=entries.find(r=>r.isPublished);
    setPublishedId(pub?pub.id:null);
  }

  // ── Lazy-load rows for a specific report ───────────────────────────────────
  async function loadReportData(id) {
    if (dataCache.current[id]) return dataCache.current[id];
    const data=await getReportData(id); // {fields, numFields, rows}
    // numFields comes back as array from JSON, convert to Set
    let nf=new Set(Array.isArray(data.numFields)?data.numFields:Object.values(data.numFields||{}));
    // If DB has no numFields stored (null/empty), auto-detect from data
    if (nf.size===0 && data.rows && data.rows.length>0) {
      const sample=data.rows.slice(0,50);
      (data.fields||[]).forEach(f=>{
        const vals=sample.map(r=>r[f]).filter(v=>v!==null&&v!==undefined&&v!=="");
        if(!vals.length) return;
        const numCount=vals.filter(v=>typeof v==="number"||(typeof v==="string"&&!isNaN(parseFloat(v))&&isFinite(v)&&String(v).trim()!=="")).length;
        if (numCount/vals.length>0.5) nf.add(f); // >50% numeric — permissive since Net Due often has zeros
      });
    }
    const result={rows:data.rows, fields:data.fields, numFields:nf};
    dataCache.current[id]=result;
    return result;
  }

  // ── Save report → POST to API, then refresh list ───────────────────────────
  async function handleSaveReport(reportData) {
    const {name,dataset,config,cardFields,updateId}=reportData;
    const nfArr=[...(dataset.numFields instanceof Set?dataset.numFields:new Set(dataset.numFields||[]))];
    const payload={name,config,cardFields:cardFields||[],rows:dataset.rows,fields:dataset.fields,numFields:nfArr};
    let result;
    if (updateId) {
      // In-place update — preserves publish status and access assignments
      result=await updateReport(updateId, payload);
    } else {
      result=await createReport(payload);
    }
    // Write fresh data INTO both caches (not delete them).
    // Deleting caused two problems:
    //  1. Page refresh → cache miss → DB re-fetch returned stale rows if DB write raced
    //  2. UserView useEffect never re-fired (same report ID) so preview stayed stale
    // Writing fresh data means: immediate local read returns the new rows, and
    // UserView picks up fresh data as soon as its effect fires.
    const freshCacheEntry = {
      rows: dataset.rows,
      fields: dataset.fields,
      numFields: new Set(nfArr),
    };
    dataCache.current[result.id] = freshCacheEntry;
    try {
      localStorage.setItem('rh_data_'+result.id, JSON.stringify({
        rows: dataset.rows,
        fields: dataset.fields,
        numFields: nfArr,
        ts: Date.now(),
      }));
    } catch(e) {/* quota — ignore */}
    await loadAllReports();
    return result.id;
  }

  // ── Delete report → DELETE from API ───────────────────────────────────────
  async function handleDeleteReport(id) {
    await apiDeleteReport(id);
    delete dataCache.current[id];
    try { localStorage.removeItem('rh_data_'+id); } catch(e) {}
    setSavedReports(prev=>prev.filter(r=>r.id!==id));
    if (publishedId===id) setPublishedId(null);
    // Clear builder state so the deleted report's config (tabs, pivot layout)
    // cannot bleed into the next fresh upload via existingConfig in confirmLoad.
    setDataset(null);
    setConfig(null);
    setActiveReportId(null);
    setCardFields([]);
  }

  // ── Publish report → always sets published=true ───────────────────────────
  async function handlePublishReport(id) {
    await apiPublishReport(id);
    await loadAllReports();
  }
  // ── Unpublish report → always sets published=false ─────────────────────────
  async function handleUnpublishReport(id) {
    await apiUnpublishReport(id);
    await loadAllReports();
  }

  // ── Login / Logout ─────────────────────────────────────────────────────────
  async function doLogin(role,username,token,id) {
    localStorage.setItem("rh_role",role);
    localStorage.setItem("rh_username",username);
    localStorage.setItem("rh_token",token);
    if(id) localStorage.setItem("rh_user_id",id);
    setCurrentUser({username, id: id||localStorage.getItem("rh_user_id")||null, role});
    try {
      await loadAllReports();
      // subadmin uses the same AdminView as admin
      setScreen(role==="subadmin"||role==="subadmin_user"?"admin":role);
    } catch(err) {
      localStorage.removeItem("rh_token");
      localStorage.removeItem("rh_role");
      localStorage.removeItem("rh_username");
      setCurrentUser(null);
      setLoadErr("Login succeeded but failed to load reports. Please try again.");
      setScreen("login");
    }
  }

  function doLogout() {
    apiLogout();
    localStorage.removeItem("rh_token");
    localStorage.removeItem("rh_role");
    localStorage.removeItem("rh_username");
    localStorage.removeItem("rh_user_id");
    setCurrentUser(null);
    setSavedReports([]);
    setPublishedId(null);
    dataCache.current={};
    setScreen("login");
  }

  return(
    <ErrorBoundary>
      {screen==="loading"
        ?<div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:T.bgPage,flexDirection:"column",gap:12}}>
            <div style={{fontSize:36,animation:"spin 1s linear infinite",display:"inline-block"}}>⚙️</div>
            <div style={{fontWeight:600,color:T.primary}}>Loading ReportHub…</div>
            {loadErr&&<div style={{fontSize:12,color:T.danger,maxWidth:300,textAlign:"center"}}>{loadErr}</div>}
            <style>{"@keyframes spin{to{transform:rotate(360deg)}}"}</style>
          </div>
        :screen==="login"
          ?<Login onLogin={doLogin}/>
          :screen==="admin"
            ?<AdminView
                onLogout={doLogout}
                savedReports={savedReports}
                publishedId={publishedId}
                onSaveReport={handleSaveReport}
                onPublishReport={handlePublishReport}
                onUnpublishReport={handleUnpublishReport}
                onDeleteReport={handleDeleteReport}
                onLoadReportData={loadReportData}
                onReloadReports={loadAllReports}
                currentUser={currentUser}
                currentRole={localStorage.getItem("rh_role")||"subadmin"}/>
            :<UserView
                onLogout={doLogout}
                savedReports={savedReports}
                onLoadReportData={loadReportData}
                currentUser={currentUser}
                currentRole={localStorage.getItem("rh_role")||"user"}
                isGuest={false}/>}
    </ErrorBoundary>
  );
}

// Safe sourceLinks accessor — prevents null crashes
function getSourceLinks(cfg) { return (cfg && cfg.sourceLinks) || []; }
