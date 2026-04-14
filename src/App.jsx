import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import _ from "lodash";
// API module — calls Railway backend. Stripped in standalone HTML (functions are global there).
import { login as apiLogin, logout as apiLogout, getUsers, createUser, updatePassword,
         deleteUser, getReports, createReport, deleteReport as apiDeleteReport,
         publishReport as apiPublishReport, getReportData } from "./api.js";

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

// ── Number formats ─────────────────────────────────────────────────────────────
const NUM_FORMATS = [
  { key:"Cr",    label:"Crores",    div:1e7, suffix:" Cr", dec:2 },
  { key:"L",     label:"Lakhs",     div:1e5, suffix:" L",  dec:2 },
  { key:"M",     label:"Millions",  div:1e6, suffix:" M",  dec:2 },
  { key:"K",     label:"Thousands", div:1e3, suffix:" K",  dec:1 },
  { key:"units", label:"Units",     div:1,   suffix:"",    dec:0 },
];

const AGGS=["sum","avg","count","min","max"];
const MAX_ROWS=100000, DRILL_PAGE=25, SLICER_SEARCH=30, SLICER_MAX=500, BLANK_THRESH=0.70;
const isMoneyField=f=>/sale|revenue|profit|price|amount|cost|income|spend|budget|fee|net|gross|pay|earn|cash|value|due|paid|deduct|bill/i.test(f);

function fmtNum(n, agg, field, fmtKey) {
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
    const filtered=data.filter(row=>config.filters.every(f=>{const s=filters[f]||[];return !s.length||s.includes(String(row[f]||""));}));
    const rFs=config.rows, cF=config.columns[0], vals=config.values;
    if (!rFs.length||!vals.length) return null;
    const compute=sub=>vals.map(v=>doAgg(sub,v.field,v.agg));
    const seenRk=new Map();
    filtered.forEach(r=>{
      const k=rFs.map(f=>String(r[f]||"")).join("\0");
      if (!seenRk.has(k)) seenRk.set(k,rFs.map(f=>String(r[f]||"")));
    });
    const rowKeys=[...seenRk.values()].sort((a,b)=>a.join("\0").localeCompare(b.join("\0")));
    const colVals=cF?_.uniq(filtered.map(r=>String(r[cF]||""))).sort():[];
    const cells={};
    rowKeys.forEach(rk=>{
      const rkStr=rk.join("\0");
      const rd=filtered.filter(r=>rFs.every((f,i)=>String(r[f]||"")===rk[i]));
      cells[rkStr]={};
      colVals.forEach(cv=>{cells[rkStr][cv]=compute(rd.filter(r=>String(r[cF]||"")===cv));});
      cells[rkStr]["__total__"]=compute(rd);
    });
    const colTotals={};
    colVals.forEach(cv=>{colTotals[cv]=compute(filtered.filter(r=>String(r[cF]||"")===cv));});
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
    const rkStr = rk.join(" ");
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
  // Open a print-friendly version in a new window
  const style = `<style>body{font-family:Arial,sans-serif;font-size:11px}table{border-collapse:collapse;width:100%}th{background:#5C2D1A;color:#F5EFE6;padding:7px 10px;text-align:right}th:first-child{text-align:left}td{padding:6px 10px;border-bottom:1px solid #D4BEA0;text-align:right}td:first-child{text-align:left;font-weight:600}tr:nth-child(even){background:#F5EEE4}tfoot td{font-weight:700;background:#EDE0CF}</style>`;
  // Grab the pivot table element from the page and print it
  const tableEl = document.querySelector("table");
  if (!tableEl) { alert("No pivot table found to export."); return; }
  const win = window.open("","_blank","width=900,height=700");
  win.document.write("<html><head><title>"+config.name+"</title>"+style+"</head><body>");
  win.document.write("<h2 style='color:#5C2D1A;margin-bottom:12px'>"+config.name+"</h2>");
  win.document.write(tableEl.outerHTML);
  win.document.write("</body></html>");
  win.document.close();
  setTimeout(()=>win.print(), 400);
}

// ── Drill-down column filter (Excel-style per-column filter in drill-down) ────
function DrillColFilter({field, data, active, onChange, numFields, activeSort, onSort}) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  // sortDir is controlled externally via activeSort/onSort when provided, else internal
  const sortDir = activeSort||"az";
  const setSortDir = dir => { onSort&&onSort(field,dir); };
  const ref = useRef(null);
  const looksNum = numFields.has(field);
  const rawOpts = useMemo(()=>_.uniq(data.map(r=>String(r[field]||""))),[data,field]);
  const sorted = useMemo(()=>{
    const o=[...rawOpts];
    if (sortDir==="az") o.sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
    else if (sortDir==="za") o.sort((a,b)=>b.localeCompare(a,undefined,{numeric:true}));
    else if (sortDir==="09") o.sort((a,b)=>parseFloat(a||0)-parseFloat(b||0));
    else o.sort((a,b)=>parseFloat(b||0)-parseFloat(a||0));
    return o;
  },[rawOpts,sortDir]);
  const vis = search ? sorted.filter(o=>o.toLowerCase().includes(search.toLowerCase())).slice(0,200) : sorted.slice(0,200);
  const toggle = v => onChange(active.includes(v)?active.filter(x=>x!==v):[...active,v]);
  const partial = active.length>0 && active.length<sorted.length;
  useEffect(()=>{
    if (!open) return;
    const h=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};
    const t=setTimeout(()=>document.addEventListener("click",h),10);
    return()=>{clearTimeout(t);document.removeEventListener("click",h);};
  },[open]);
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
        <div style={{position:"absolute",top:"100%",left:0,zIndex:400,background:T.bgCard,border:"1px solid "+T.border,
          borderRadius:8,minWidth:220,maxWidth:300,boxShadow:"0 6px 20px rgba(92,45,26,0.2)",overflow:"hidden"}}>
          <div style={{padding:"7px 10px",background:T.bgStat,borderBottom:"0.5px solid "+T.border}}>
            <div style={{fontSize:10,color:T.textMd,fontWeight:600,marginBottom:4}}>Sort</div>
            <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
              <SortBtn dir="az" label="A→Z"/>
              <SortBtn dir="za" label="Z→A"/>
              {looksNum&&<SortBtn dir="09" label="0→9"/>}
              {looksNum&&<SortBtn dir="90" label="9→0"/>}
            </div>
          </div>
          <div style={{padding:"6px 10px",borderBottom:"0.5px solid "+T.border}}>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search..."
              style={{width:"100%",padding:"4px 8px",border:"0.5px solid "+T.border,borderRadius:4,fontSize:11,background:T.bgCard,color:T.text,boxSizing:"border-box",outline:"none"}}/>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",padding:"5px 10px",borderBottom:"0.5px solid "+T.border}}>
            <button onClick={()=>onChange([])} style={{fontSize:10,background:"none",border:"none",cursor:"pointer",color:T.textMd}}>Clear</button>
            <button onClick={()=>onChange(sorted)} style={{fontSize:10,background:"none",border:"none",cursor:"pointer",color:T.primary,fontWeight:600}}>All</button>
          </div>
          <div style={{maxHeight:200,overflowY:"auto"}}>
            {vis.map(o=>(
              <label key={o} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 10px",cursor:"pointer",fontSize:11,
                background:active.includes(o)?"rgba(92,45,26,0.05)":undefined,color:T.text}}>
                <input type="checkbox" checked={active.includes(o)} onChange={()=>toggle(o)} style={{width:12,height:12,accentColor:T.primary}}/>
                <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{o||"(blank)"}</span>
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Drill-down panel ──────────────────────────────────────────────────────────
function DrillDown({data,target,fields,numFields,onClose,numFmt,savedHiddenCols,onSaveHiddenCols}) {
  const [page,setPage]=useState(0);
  const [pageSize,setPageSize]=useState(25); // 25|50|100|"all"
  const [hiddenCols,setHiddenCols]=useState(()=>new Set(savedHiddenCols||[]));
  const [showColPicker,setShowColPicker]=useState(false);
  const [colFilters,setColFilters]=useState({}); // {field: [selectedValues]}
  const [rowSort,setRowSort]=useState({}); // {field: dir} — only one active at a time
  const [colWidths,startColResize]=useColResize(130);
  const {rowKey,colVal,rFs,cF,metricLabel}=target;
  const baseRows=useMemo(()=>data.filter(row=>
    rFs.every((f,i)=>String(row[f]||"")===rowKey[i])&&
    (!cF||!colVal||colVal==="__total__"||String(row[cF]||"")===colVal)
  ),[data,target]);
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
  const setColFilter=(f,sel)=>{setColFilters(p=>({...p,[f]:sel}));setPage(0);};
  const hasColFilters=Object.values(colFilters).some(v=>v&&v.length);
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
                    {onSaveHiddenCols&&<button onClick={()=>{onSaveHiddenCols([...hiddenCols]);}} style={{fontSize:10,color:T.primary,background:"none",border:"none",cursor:"pointer",fontWeight:700}}>Save layout</button>}
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
        {/* Hint + clear column filters */}
        <div style={{fontSize:11,color:T.textMd,padding:"5px 14px",background:T.bgStat,borderBottom:"0.5px solid "+T.border,flexShrink:0,display:"flex",alignItems:"center",gap:10}}>
          <span>Columns in original Excel order · Click ⋏ on column headers to filter/sort · Scroll right to see all</span>
          {hasColFilters&&<button onClick={()=>setColFilters({})} style={{fontSize:10,color:T.danger,background:"none",border:"none",cursor:"pointer",textDecoration:"underline",flexShrink:0}}>Clear column filters</button>}
        </div>
        {/* Table — full horizontal scroll, all columns, original order */}
        <div style={{overflowX:"auto",flex:1,overflowY:"auto"}}>
          <table style={{borderCollapse:"collapse",fontSize:12,tableLayout:"fixed",minWidth:"100%"}}>
            <thead style={{position:"sticky",top:0,zIndex:5}}><tr style={{background:T.bgTableH}}>
              {visibleCols.map(f=>{
                const fActive=(colFilters[f]||[]).length>0;
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
                        :numFields.has(f)?fmtNum(+row[f],"sum",f,numFmt):String(row[f])}
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
                        isNum?fmtNum(colSums[f]||0,"sum",f,numFmt):""}
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
function QuickFilterCards({field,data,activeFilters,onFilter,primaryVal,numFmt}) {
  const opts=useMemo(()=>_.uniq(data.map(r=>String(r[field]||""))).sort(),[data,field]);
  const active=activeFilters||[];
  const allActive=active.length===0;
  const cardStyle=(on)=>({
    flexShrink:0,padding:"10px 16px",borderRadius:8,textAlign:"left",cursor:"pointer",minWidth:110,
    border:on?"2px solid "+T.primary:"1px solid "+T.border,
    background:on?T.primary:T.bgCard,
    transition:"all 0.15s",
  });
  return(
    <div style={{marginBottom:16}}>
      <div style={{fontSize:10,fontWeight:700,color:T.textMd,textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:8}}>{field}</div>
      <div style={{display:"flex",gap:8,overflowX:"auto",paddingBottom:4}}>
        <button onClick={()=>onFilter([])} style={cardStyle(allActive)}>
          <div style={{fontSize:10,color:allActive?T.textLt:T.textMd,marginBottom:3}}>All</div>
          <div style={{fontSize:15,fontWeight:700,color:allActive?T.textLt:T.numColor}}>
            {fmtNum(doAgg(data,primaryVal.field,primaryVal.agg),primaryVal.agg,primaryVal.field,numFmt)}
          </div>
        </button>
        {opts.map(val=>{
          const on=active.includes(val);
          const rows=data.filter(r=>String(r[field]||"")===val);
          return(
            <button key={val} onClick={()=>on?onFilter(active.filter(x=>x!==val)):onFilter([...active,val])} style={cardStyle(on)}>
              <div style={{fontSize:10,color:on?T.textLt:T.textMd,marginBottom:3,maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{val||"(blank)"}</div>
              <div style={{fontSize:15,fontWeight:700,color:on?T.textLt:T.numColor}}>
                {fmtNum(doAgg(rows,primaryVal.field,primaryVal.agg),primaryVal.agg,primaryVal.field,numFmt)}
              </div>
              <div style={{fontSize:10,color:on?"rgba(245,239,230,0.6)":T.textMd,marginTop:2}}>{rows.length.toLocaleString()} rows</div>
            </button>
          );
        })}
      </div>
    </div>
  );
}

// ── Slicer — Excel-style with sort A-Z / Z-A / numeric ───────────────────────
function Slicer({field,active,onChange,data}) {
  const [open,setOpen]=useState(false);
  const [search,setSearch]=useState("");
  const [sortDir,setSortDir]=useState("az"); // "az"|"za"|"09"|"90"
  const ref=useRef(null);
  const rawOpts=useMemo(()=>_.uniq(data.map(r=>String(r[field]||""))),[field,data]);
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
  const toggle=v=>onChange(active.includes(v)?active.filter(x=>x!==v):[...active,v]);
  const partial=active.length>0&&active.length<sortedOpts.length;
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
        {partial&&<span style={{background:"rgba(255,255,255,0.25)",color:T.textLt,borderRadius:10,padding:"1px 7px",fontSize:11,fontWeight:600}}>{active.length}</span>}
        <span style={{fontSize:9,opacity:0.5}}>{open?"▲":"▼"}</span>
      </button>
      {open&&(
        <div style={{position:"absolute",top:"calc(100% + 5px)",left:0,zIndex:300,background:T.bgCard,border:"1px solid "+T.border,borderRadius:8,minWidth:260,maxWidth:340,boxShadow:"0 8px 28px rgba(92,45,26,0.2)",overflow:"hidden"}}>
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
            <button onClick={()=>onChange([])} style={{fontSize:11,background:"none",border:"none",cursor:"pointer",color:T.textMd}}>Clear all</button>
            <span style={{fontSize:10,color:T.textMd}}>{sortedOpts.length} values</span>
            <button onClick={()=>onChange(sortedOpts)} style={{fontSize:11,background:"none",border:"none",cursor:"pointer",color:T.primary,fontWeight:600}}>Select all</button>
          </div>
          {/* Checkbox list */}
          <div style={{maxHeight:250,overflowY:"auto"}}>
            {visOpts.map(o=>(
              <label key={o} style={{display:"flex",alignItems:"center",gap:9,padding:"6px 12px",cursor:"pointer",fontSize:12,background:active.includes(o)?"rgba(92,45,26,0.05)":undefined,color:T.text}}>
                <input type="checkbox" checked={active.includes(o)} onChange={()=>toggle(o)} style={{width:13,height:13,accentColor:T.primary,flexShrink:0}}/>
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
function PivotTable({result,onDrillDown,numFmt,colOrder,onColReorder,pivotFilters,onPivotFilter,pivotSort,onPivotSort}) {
  // ── ALL hooks must come before any conditional return (Rules of Hooks) ──────
  const [dragOverCol,setDragOverCol]=useState(null);
  const [colWidths,startColResize]=useColResize(120);

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
        const sel=pivotFilters[i]||[];
        return !sel.length||sel.includes(rk[i]);
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
  // Use external colOrder if provided (for column group drag-reorder), else default
  const orderedColVals=colOrder&&colOrder.length===colVals.length?colOrder:colVals;
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
  const getGrand=col=>(col.key==="__total__"?grandTotals:(colTotals[col.key]||effectiveVals.map(()=>0)))[col.vi]||0;
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
                      active={pivotFilters&&pivotFilters[ri]||[]}
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
          {rowKeys.map((rk,ri)=>{
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
function FormatSelector({value,onChange}) {
  return(
    <div style={{display:"flex",alignItems:"center",gap:6,padding:"4px",background:T.bgStat,borderRadius:8,border:"1px solid "+T.border}}>
      <span style={{fontSize:11,color:T.textMd,paddingLeft:6,fontWeight:500,whiteSpace:"nowrap"}}>Show in:</span>
      {NUM_FORMATS.map(f=>(
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
function Report({config,data,fields,numFields,showExport,cardFields}) {
  const [filters,setFilters]=useState({});
  const [drill,setDrill]=useState(null);
  const [numFmt,setNumFmt]=useState("Cr");
  const [colOrder,setColOrder]=useState(null);
  const [adHocFields,setAdHocFields]=useState([]); // extra filters user adds in view mode
  const [drillHiddenCols,setDrillHiddenCols]=useState([]); // saved column layout for drill-down
  const [pivotFilters,setPivotFilters]=useState({}); // {rowFieldIdx: [selectedValues]}
  const [pivotSort,setPivotSort]=useState(null); // {fieldIdx, dir}
  const [showAdHocPicker,setShowAdHocPicker]=useState(false);
  const adHocRef=useRef(null);
  const result=useMemo(()=>runPivot(data,config,filters),[config,data,filters]);
  useEffect(()=>{if(result&&!result.error&&result.colVals)setColOrder(null);},[config]);
  const setF=(f,v)=>setFilters(p=>({...p,[f]:v}));
  const hasActive=Object.values(filters).some(v=>v&&v.length);
  const slicerFields=(config.filters||[]).filter(f=>!(cardFields||[]).includes(f));
  const primaryVal=(config.values||[])[0]||{field:"",agg:"sum"};
  // All dimension fields available for ad-hoc filtering
  const dimFields=useMemo(()=>fields.filter(f=>!numFields.has(f)),[fields,numFields]);
  // Ad-hoc fields not already in configured slicers or card fields
  const addableFields=dimFields.filter(f=>!slicerFields.includes(f)&&!(cardFields||[]).includes(f)&&!adHocFields.includes(f));
  useEffect(()=>{
    if (!showAdHocPicker) return;
    const h=e=>{if(adHocRef.current&&!adHocRef.current.contains(e.target))setShowAdHocPicker(false);};
    const t=setTimeout(()=>document.addEventListener("click",h),10);
    return()=>{clearTimeout(t);document.removeEventListener("click",h);};
  },[showAdHocPicker]);
  function handleColReorder(from,to) {
    // Two modes: (1) column field set - reorder colVals; (2) no column field - reorder value metric names
    const hasColField=result&&result.colVals&&result.colVals.length>0;
    const base=hasColField
      ?(colOrder||[...result.colVals])
      :(colOrder||(config.values||[]).map(v=>v.field));
    const fi=base.indexOf(from),ti=base.indexOf(to);
    if(fi===-1||ti===-1)return;
    const arr=[...base];arr.splice(fi,1);arr.splice(ti,0,from);
    setColOrder(arr);
  }

  return(
    <div>
      {/* Format selector + export row */}
      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:16,flexWrap:"wrap"}}>
        <FormatSelector value={numFmt} onChange={setNumFmt}/>
        {hasActive&&<button onClick={()=>setFilters({})} style={{fontSize:12,color:T.textMd,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>Clear all filters</button>}
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

      {/* KPI stat cards */}
      {result&&!result.error&&(
        <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
          {result.vals.map((v,i)=>(
            <div key={i} style={{background:i===0?T.primary:T.bgCard,borderRadius:8,padding:"12px 16px",flex:1,minWidth:120,
              border:"1px solid "+(i===0?T.primary:T.border),boxShadow:"0 1px 4px rgba(92,45,26,0.1)"}}>
              <div style={{fontSize:10,color:i===0?"rgba(245,239,230,0.7)":T.textMd,marginBottom:4,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.5px"}}>{v.agg} of {v.field}</div>
              <div style={{fontSize:20,fontWeight:700,color:i===0?T.textLt:T.numColor}}>{fmtNum(result.grandTotals[i],v.agg,v.field,numFmt)}</div>
            </div>
          ))}
          <div style={{background:T.bgCard,borderRadius:8,padding:"12px 16px",flex:1,minWidth:120,border:"1px solid "+T.border}}>
            <div style={{fontSize:10,color:T.textMd,marginBottom:4,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.5px"}}>Records</div>
            <div style={{fontSize:20,fontWeight:700,color:T.numColor}}>{result.count.toLocaleString()}</div>
          </div>
        </div>
      )}

      {/* Card filters */}
      {(cardFields||[]).map(f=>(
        <QuickFilterCards key={f} field={f} data={data} activeFilters={filters[f]||[]}
          onFilter={v=>setF(f,v)} primaryVal={primaryVal} numFmt={numFmt}/>
      ))}

      {/* Slicers — configured + ad-hoc */}
      {(slicerFields.length>0||adHocFields.length>0||addableFields.length>0)&&(
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14,flexWrap:"wrap"}}>
          <span style={{fontSize:12,color:T.textMd,fontWeight:600}}>Filters:</span>
          {slicerFields.map(f=><Slicer key={f} field={f} active={filters[f]||[]} onChange={v=>setF(f,v)} data={data}/>)}
          {adHocFields.map(f=>(
            <div key={f} style={{position:"relative",display:"inline-flex",alignItems:"center",gap:2}}>
              <Slicer field={f} active={filters[f]||[]} onChange={v=>setF(f,v)} data={data}/>
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
          {hasActive&&<button onClick={()=>setFilters({})} style={{fontSize:11,color:T.textMd,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>Clear all</button>}
        </div>
      )}

      <PivotTable result={result} numFmt={numFmt}
        colOrder={colOrder&&result&&result.colVals?colOrder:undefined}
        onColReorder={result&&!result.error&&
          ((result.colVals&&result.colVals.length>1)||((!result.cF)&&result.vals&&result.vals.length>1))
          ?handleColReorder:undefined}
        pivotFilters={Object.keys(pivotFilters).length?pivotFilters:null}
        onPivotFilter={(idx,sel)=>setPivotFilters(p=>({...p,[idx]:sel}))}
        pivotSort={pivotSort}
        onPivotSort={setPivotSort}
        onDrillDown={(rowKey,colVal,label)=>setDrill({rowKey,colVal,rFs:result.rFs,cF:result.cF,metricLabel:label})}/>

      {drill&&<DrillDown data={data} target={drill} fields={fields} numFields={numFields} numFmt={numFmt}
        savedHiddenCols={drillHiddenCols}
        onSaveHiddenCols={cols=>{setDrillHiddenCols(cols);}}
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
            extra={<select value={v.agg} onChange={e=>onAggChange&&onAggChange(v.field,e.target.value)}
              style={{fontSize:10,border:"none",background:"transparent",color,cursor:"pointer",padding:"0 2px",marginLeft:3}}>
              {AGGS.map(a=><option key={a} value={a}>{a}</option>)}
            </select>}/>
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
  return(
    <div style={{position:"sticky",top:0,zIndex:50,background:T.bgHeader,borderBottom:"2px solid "+T.borderHd,
      padding:"0 20px",display:"flex",alignItems:"center",gap:12,height:52,
      boxShadow:"0 2px 12px rgba(44,24,16,0.3)"}}>
      <span style={{fontWeight:700,fontSize:15,color:T.textLt,letterSpacing:"-0.3px"}}>
        <span style={{color:T.accent}}>Report</span>Hub
      </span>
      <span style={{color:"rgba(245,239,230,0.3)"}}>|</span>
      <span style={{fontSize:11,color:T.textLt,background:"rgba(255,255,255,0.12)",padding:"2px 10px",borderRadius:4,fontWeight:500}}>{role}</span>
      <div style={{flex:1}}/>{children}
      <button onClick={onLogout} style={{padding:"5px 14px",background:"rgba(255,255,255,0.12)",border:"1px solid rgba(255,255,255,0.2)",borderRadius:6,cursor:"pointer",fontSize:12,color:T.textLt}}>Logout</button>
    </div>
  );
}

// ── Upload Tab ─────────────────────────────────────────────────────────────────
function UploadTab({libs, onDataLoaded, existingConfig}) {
  const [phase,setPhase]=useState("drop");
  const [dragOver,setDragOver]=useState(false);
  const [fileInfo,setFileInfo]=useState(null);
  const [sheetNames,setSheetNames]=useState([]);
  const [workbook,setWorkbook]=useState(null);
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

  function applySchema(rows,fields,name) {
    if (!rows.length){setParseError("No data rows found after cleaning.");setPhase("error");return;}
    const numFields=detectNumFields(rows,fields);
    const scm=fields.map(f=>({
      field:f,type:numFields.has(f)?"num":"dim",
      sample:_.uniq(rows.slice(0,5).map(r=>String(r[f]||"")).filter(Boolean)).slice(0,3),
      nullPct:Math.round(rows.filter(r=>r[f]===""||r[f]===null||r[f]===undefined).length/rows.length*100),
      uniqueCount:_.uniq(rows.map(r=>String(r[f]||""))).length,
    }));
    setAllRows(rows);setAllFields(fields);setPreviewRows(rows.slice(0,8));setSchema(scm);
    setParseStats({rows:rows.length,fields:fields.length,name});setPhase("preview");
  }

  function processRaw(rawRows,name) {
    try{const{rows,fields}=sanitizeRows(rawRows);applySchema(rows,fields,name);}
    catch(e){setParseError("Cleaning error: "+e.message);setPhase("error");}
  }

  function loadSheet(wb,sheetName) {
    setPhase("parsing");
    setTimeout(()=>{
      try{
        const ws=wb.Sheets[sheetName];
        if (!ws){setParseError("Sheet not found: "+sheetName);setPhase("error");return;}
        if (ws["!ref"]){
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
        if (wb.SheetNames.length===1)loadSheet(wb,wb.SheetNames[0]);
        else{setSheetNames(wb.SheetNames);setPhase("sheet");}
      }
    }catch(e){setParseError("Read error: "+e.message);setPhase("error");}
  }

  async function handleUrl() {
    if (!refreshUrl.trim()){setParseError("Enter a URL first.");setPhase("error");return;}
    setPhase("parsing");setParseError("");
    try{
      const resp=await fetch(refreshUrl.trim());
      if (!resp.ok) throw new Error("HTTP "+resp.status);
      const buf=await resp.arrayBuffer();
      const wb=libs.XLSX.read(buf,{type:"array",cellDates:true});
      setWorkbook(wb);setLastRefresh(new Date());
      if (refreshSheet&&wb.SheetNames.includes(refreshSheet)){loadSheet(wb,refreshSheet);}
      else if (wb.SheetNames.length===1){loadSheet(wb,wb.SheetNames[0]);}
      else{setSheetNames(wb.SheetNames);setPhase("sheet");}
    }catch(e){setParseError("URL fetch failed: "+e.message+". Ensure it is a direct download link with CORS enabled.");setPhase("error");}
  }

  const onDrop=useCallback(e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)handleFile(f);},[libs]);

  function toggleType(field){setSchema(s=>s.map(item=>item.field===field?{...item,type:item.type==="num"?"dim":"num"}:item));}

  function confirmLoad() {
    const numFields=new Set(schema.filter(s=>s.type==="num").map(s=>s.field));
    const fields=schema.map(s=>s.field); // preserve original chronological order
    const name=parseStats&&parseStats.name?parseStats.name:"Report";
    const config=existingConfig?{...existingConfig,name}:autoConfig(fields,numFields,name);
    const rows=allRows.map(r=>{
      const out={...r};
      fields.forEach(f=>{if(numFields.has(f)){const v=r[f];if(typeof v!=="number"){const n=parseFloat(String(v||"").replace(/[$,\u20B9]/g,""));out[f]=isNaN(n)?0:n;}}});
      return out;
    });
    onDataLoaded({rows,fields,numFields,config});
  }

  const fmtSize=b=>b>1048576?(b/1048576).toFixed(1)+" MB":(b/1024).toFixed(1)+" KB";
  const inp={width:"100%",padding:"8px 11px",border:"1px solid "+T.border,borderRadius:7,fontSize:13,background:T.bgCard,color:T.text,boxSizing:"border-box",outline:"none"};

  return(
    <div style={{padding:20,maxWidth:960,margin:"0 auto"}}>

      {(phase==="drop"||phase==="error")&&(<>
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
            <span style={{flex:1}}>{parseError}</span>
            <button onClick={()=>{setPhase("drop");setParseError("");}} style={{fontSize:12,color:T.danger,background:"none",border:"none",cursor:"pointer",textDecoration:"underline",flexShrink:0}}>Try again</button>
          </div>
        )}

        <div style={{marginTop:18,padding:"16px 18px",background:T.bgCard,borderRadius:10,border:"1px solid "+T.border}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
            <span style={{fontSize:16}}>🔗</span>
            <span style={{fontWeight:700,fontSize:13,color:T.text}}>Auto-refresh from URL</span>
            {lastRefresh&&<span style={{fontSize:11,color:T.textMd,marginLeft:"auto"}}>Last refreshed: {lastRefresh.toLocaleTimeString()}</span>}
          </div>
          <div style={{fontSize:12,color:T.textMd,marginBottom:10,lineHeight:1.6}}>
            Paste a direct download URL (SharePoint public share, Dropbox, OneDrive, or any CORS-enabled server). Your pivot layout is preserved on refresh.
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            <input value={refreshUrl} onChange={e=>setRefreshUrl(e.target.value)} placeholder="https://your-server.com/report.xlsx"
              style={{...inp,flex:"2 1 300px"}}/>
            <input value={refreshSheet} onChange={e=>setRefreshSheet(e.target.value)} placeholder="Sheet name (optional)"
              style={{...inp,flex:"1 1 150px"}}/>
            <button onClick={handleUrl} disabled={!refreshUrl.trim()||!libsReady}
              style={{padding:"8px 18px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,cursor:refreshUrl.trim()&&libsReady?"pointer":"not-allowed",fontSize:13,fontWeight:600,opacity:refreshUrl.trim()&&libsReady?1:0.5,whiteSpace:"nowrap"}}>
              Refresh Now
            </button>
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
          <button onClick={()=>{setPhase("drop");setWorkbook(null);}} style={{marginTop:14,fontSize:13,color:T.textMd,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>Different file</button>
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
            </div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>{setPhase("drop");setAllRows([]);setSchema([]);}}
                style={{padding:"8px 16px",background:"none",border:"1px solid "+T.border,borderRadius:7,cursor:"pointer",fontSize:13,color:T.text}}>Different file</button>
              <button onClick={confirmLoad}
                style={{padding:"8px 20px",background:T.primary,color:T.textLt,border:"none",borderRadius:7,cursor:"pointer",fontSize:13,fontWeight:600}}>Load into builder</button>
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
function AdminView({onLogout,savedReports,publishedId,onSaveReport,onPublishReport,onDeleteReport,onLoadReportData,onReloadReports,currentUser}) {
  const libs=useLibs();
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

  const effectiveNumFields=useMemo(()=>{
    if (!dataset) return new Set();
    const s=new Set(dataset.numFields);
    Object.entries(typeOverrides).forEach(([f,t])=>{if(t==="num")s.add(f);else s.delete(f);});
    return s;
  },[dataset,typeOverrides]);

  function onDataLoaded(ds){setDataset(ds);setConfig(ds.config);setTypeOverrides({});setCardFields([]);setActiveReportId(null);setTab("builder");}
  async function openSavedReport(id) {
    const r=savedReports.find(x=>x.id===id);
    if (!r) return;
    setApiLoading(true);
    try {
      const data=await onLoadReportData(id);
      const ds={rows:data.rows,fields:data.fields,numFields:data.numFields};
      setDataset(ds);setConfig(r.config);setCardFields(r.cardFields||[]);setTypeOverrides({});
      setActiveReportId(id);
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
      if (overwrite&&activeReportId) {
        // Delete old then save with same name (Railway API doesn't have PATCH for full data)
        await onDeleteReport(activeReportId);
      }
      const id=await onSaveReport({name:config.name,dataset:{...dataset,numFields:effectiveNumFields},config,cardFields});
      setActiveReportId(id);
      showToast(overwrite?"Report updated!":"Report saved as new!");
    }catch(e){showToast("Save failed: "+e.message);}
    finally{setApiLoading(false);}
  }
  async function doPublish(id) {
    setApiLoading(true);
    try{
      await onPublishReport(id);
      showToast("Report published to users!");
    }catch(e){showToast("Publish failed: "+e.message);}
    finally{setApiLoading(false);}
  }

  function toggleFieldType(field) {
    const curNum=effectiveNumFields.has(field);
    setTypeOverrides(p=>({...p,[field]:curNum?"dim":"num"}));
    if (curNum) setConfig(c=>({...c,values:c.values.filter(v=>v.field!==field)}));
  }

  function toggleCard(field){setCardFields(cf=>cf.includes(field)?cf.filter(f=>f!==field):[...cf,field]);}

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

  function setAgg(field,agg){setConfig(c=>({...c,values:c.values.map(v=>v.field===field?{...v,agg}:v)}));}

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

  const preview=useMemo(()=>dataset&&config?runPivot(dataset.rows,config,{}):[],[dataset,config]);
  const fieldStatus=useMemo(()=>{
    if (!dataset||!config) return {};
    const z={};
    dataset.fields.forEach(f=>{z[f]={rows:config.rows.includes(f),cols:config.columns.includes(f),vals:config.values.some(v=>v.field===f),filters:config.filters.includes(f),card:cardFields.includes(f)};});
    return z;
  },[dataset,config,cardFields]);

  const TABS=[["upload","Upload"],["builder","Report Builder",!dataset],["preview","User Preview",!dataset],["data","Raw Data",!dataset],["reports","Reports ("+savedReports.length+")"]];
  const tabBtn=(t,l,disabled)=>(
    <button key={t} onClick={()=>!disabled&&setTab(t)} style={{padding:"11px 16px",background:"none",border:"none",cursor:disabled?"not-allowed":"pointer",fontSize:13,
      borderBottom:tab===t?"2px solid "+T.accent:"2px solid transparent",
      fontWeight:tab===t?700:400,color:disabled?T.textMd:tab===t?T.textLt:"rgba(245,239,230,0.6)",opacity:disabled?0.4:1}}>
      {l}
    </button>
  );

  return(
    <div style={{minHeight:"100vh",background:T.bgPage,fontFamily:"system-ui,sans-serif"}}>
      <AppHeader role="Admin" onLogout={onLogout}>
        {toast&&<span style={{fontSize:12,color:T.textLt,background:"rgba(45,106,79,0.5)",padding:"4px 12px",borderRadius:6,fontWeight:500,border:"1px solid rgba(45,106,79,0.6)"}}>{toast}</span>}
        {dataset&&config&&<button onClick={doSave} disabled={apiLoading} style={{padding:"6px 14px",background:"rgba(255,255,255,0.15)",color:T.textLt,border:"1px solid rgba(255,255,255,0.25)",borderRadius:6,cursor:apiLoading?"wait":"pointer",fontSize:12,fontWeight:600,opacity:apiLoading?0.6:1}}>
          {apiLoading?"Saving…":"Save Report"}
        </button>}
        <button onClick={()=>setShowSettings(true)} title="User management & settings"
          style={{padding:"6px 12px",background:"rgba(255,255,255,0.12)",color:T.textLt,border:"1px solid rgba(255,255,255,0.2)",borderRadius:6,cursor:"pointer",fontSize:12}}>
          ⚙ Settings
        </button>
      </AppHeader>

      <div style={{position:"sticky",top:52,zIndex:40,background:T.bgHeader,borderBottom:"1px solid "+T.borderHd,padding:"0 20px",display:"flex"}}>
        {TABS.map(([t,l,d])=>tabBtn(t,l,d))}
      </div>

      {tab==="upload"&&<UploadTab libs={libs} onDataLoaded={onDataLoaded} existingConfig={config}/>}

      {tab==="builder"&&dataset&&config&&(
        <div style={{padding:20,display:"grid",gridTemplateColumns:"290px 1fr",gap:20,alignItems:"start"}}>

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
                {dataset.fields.map(f=>(
                  <FieldRow key={f} field={f} isNum={effectiveNumFields.has(f)}
                    status={fieldStatus[f]||{}} onToggle={toggleField}
                    onToggleType={()=>toggleFieldType(f)} onToggleCard={()=>toggleCard(f)}/>
                ))}
              </div>
            </div>

            <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:10,padding:14}}>
              <div style={{fontWeight:700,fontSize:13,color:T.primary,marginBottom:8}}>Report Name</div>
              <input value={config.name} onChange={e=>setConfig(c=>({...c,name:e.target.value}))}
                style={{width:"100%",padding:"7px 10px",border:"1px solid "+T.border,borderRadius:6,fontSize:13,background:T.bgStat,color:T.text,boxSizing:"border-box",outline:"none"}}/>
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
                onRemove={f=>setCardFields(cf=>cf.filter(x=>x!==f))} onReorder={(a,b)=>setCardFields(cf=>{const fi=cf.indexOf(a),ti=cf.indexOf(b);if(fi===-1||ti===-1)return cf;const arr=[...cf];arr.splice(fi,1);arr.splice(ti,0,a);return arr;})}
                emptyMsg="Press K on a dimension field"/>
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
          <Report config={config} data={dataset.rows} fields={dataset.fields} numFields={effectiveNumFields} showExport cardFields={cardFields}/>
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
          onPublish={doPublish}/>
      )}
      {showSettings&&<SettingsPanel currentUser={currentUser} onClose={()=>setShowSettings(false)}/>}
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

// ── User view ──────────────────────────────────────────────────────────────────
function UserView({onLogout,savedReports,publishedReport,onLoadReportData}) {
  const [activeId,setActiveId]=useState(null);
  const [loadedData,setLoadedData]=useState({}); // {reportId -> {rows,fields,numFields}}
  const [dataLoading,setDataLoading]=useState(false);

  const currentMeta=useMemo(()=>{
    if (activeId) return savedReports.find(r=>r.id===activeId)||publishedReport;
    return publishedReport;
  },[activeId,savedReports,publishedReport]);

  // Load data whenever the selected report changes
  useEffect(()=>{
    if (!currentMeta) return;
    const id=currentMeta.id;
    if (loadedData[id]) return; // already loaded
    setDataLoading(true);
    onLoadReportData(id)
      .then(data=>setLoadedData(p=>({...p,[id]:data})))
      .catch(e=>console.error("Load error",e))
      .finally(()=>setDataLoading(false));
  },[currentMeta]);

  const currentData=currentMeta?loadedData[currentMeta.id]:null;

  if (!publishedReport&&!savedReports.length) return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:T.bgPage}}>
      <div style={{textAlign:"center"}}>
        <div style={{fontSize:44,marginBottom:14}}>📋</div>
        <div style={{fontWeight:700,fontSize:16,color:T.text,marginBottom:8}}>No reports available yet</div>
        <div style={{fontSize:13,color:T.textMd}}>Ask your admin to upload data and publish a report.</div>
      </div>
    </div>
  );

  return(
    <div style={{minHeight:"100vh",background:T.bgPage,fontFamily:"system-ui,sans-serif"}}>
      <AppHeader role="User" onLogout={onLogout}>
        {savedReports.length>0&&(
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:11,color:"rgba(245,239,230,0.6)"}}>Report:</span>
            <select value={activeId||(publishedReport?"published":"")}
              onChange={e=>setActiveId(e.target.value==="published"?null:e.target.value)}
              style={{padding:"4px 8px",border:"1px solid rgba(255,255,255,0.25)",borderRadius:6,background:"rgba(255,255,255,0.1)",
                color:T.textLt,fontSize:12,cursor:"pointer",outline:"none",maxWidth:200}}>
              {publishedReport&&<option value="published">{publishedReport.name} (published)</option>}
              {savedReports.filter(r=>!r.isPublished).map(r=>(
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
        <div style={{padding:20}}>
          <div style={{display:"flex",alignItems:"baseline",gap:10,marginBottom:4}}>
            <div style={{fontWeight:700,fontSize:18,color:T.primary}}>{currentMeta.config.name}</div>
            {currentMeta.isPublished&&<span style={{fontSize:11,background:T.primary,color:T.textLt,padding:"2px 8px",borderRadius:10,fontWeight:600}}>Published</span>}
          </div>
          <div style={{fontSize:12,color:T.textMd,marginBottom:18}}>
            {currentData.rows.length.toLocaleString()} records · {currentData.fields.length} fields · Click cells to drill down
          </div>
          <Report
            config={currentMeta.config}
            data={currentData.rows}
            fields={currentData.fields}
            numFields={currentData.numFields}
            showExport
            cardFields={currentMeta.cardFields||[]}/>
        </div>
      ):(!dataLoading&&<div style={{padding:40,textAlign:"center",fontSize:13,color:T.textMd}}>Select a report above.</div>)}
    </div>
  );
}


// ── Settings / User Management ────────────────────────────────────────────────
function SettingsPanel({currentUser,onClose}) {
  const [users,setUsers]=useState([]);
  const [pwdEdits,setPwdEdits]=useState({}); // {id: newPassword}
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
          {/* Existing users */}
          <div style={{fontWeight:700,fontSize:13,color:T.primary,marginBottom:12}}>Existing users ({users.length})</div>
          <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:20}}>
            {users.map(u=>(
              <div key={u.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",background:T.bgStat,borderRadius:8,border:"1px solid "+T.border}}>
                <div style={{width:32,height:32,borderRadius:8,background:u.role==="admin"?T.primary:T.secondary,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,color:T.textLt,fontWeight:700,flexShrink:0}}>
                  {u.role==="admin"?"A":"U"}
                </div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:600,fontSize:13,color:T.text}}>{u.username} {u.username===currentUser&&<span style={{fontSize:10,color:T.textMd}}>(you)</span>}</div>
                  <div style={{fontSize:11,color:T.textMd}}>{u.role}</div>
                </div>
                <input type="password" value={pwdEdits[u.id]||""} onChange={e=>setPwdEdits(p=>({...p,[u.id]:e.target.value}))}
                  placeholder="New password" title="Change password"
                  style={{...inp,width:130,flexShrink:0}}/>
                {pwdEdits[u.id]&&<button onClick={()=>savePwd(u.id)} disabled={loading}
                  style={{padding:"5px 8px",border:"1px solid "+T.primary,borderRadius:6,background:T.primary,cursor:"pointer",fontSize:11,color:T.textLt,flexShrink:0,fontWeight:600}}>
                  Save
                </button>}
                {u.username!==currentUser&&(
                  <button onClick={()=>delUser(u.id)} disabled={loading} style={{padding:"5px 10px",border:"1px solid rgba(163,45,45,0.4)",borderRadius:6,background:"none",cursor:"pointer",fontSize:11,color:T.danger,flexShrink:0}}>
                    Delete
                  </button>
                )}
              </div>
            ))}
          </div>
          {/* Add new user */}
          <div style={{fontWeight:700,fontSize:13,color:T.primary,marginBottom:10}}>Add new user</div>
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
                style={{...inp,width:"auto",cursor:"pointer"}}>
                <option value="user">User</option>
                <option value="admin">Admin</option>
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
  const [err,setErr]=useState("");
  const [loading,setLoading]=useState(false);
  const inp={width:"100%",padding:"9px 12px",border:"1px solid "+T.border,borderRadius:7,fontSize:13,background:T.bgCard,color:T.text,boxSizing:"border-box",outline:"none"};
  async function tryLogin(){
    if (!username.trim()||!password){setErr("Enter username and password.");return;}
    setLoading(true);setErr("");
    try{
      const data=await apiLogin(username.trim(),password);
      onLogin(data.role,data.username,data.token);
    }catch(e){
      setErr(e.message||"Login failed. Check credentials.");
    }finally{setLoading(false);}
  }
  return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:T.bgPage,fontFamily:"system-ui,sans-serif"}}>
      <div style={{background:T.bgCard,border:"1px solid "+T.border,borderRadius:16,padding:"40px 36px",width:380,boxShadow:"0 4px 24px rgba(92,45,26,0.15)"}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{width:54,height:54,background:T.primary,borderRadius:14,display:"flex",alignItems:"center",justifyContent:"center",fontSize:26,margin:"0 auto 14px"}}>📊</div>
          <h2 style={{fontSize:22,fontWeight:800,margin:"0 0 4px",color:T.primary,letterSpacing:"-0.5px"}}>ReportHub</h2>
          <p style={{fontSize:12,color:T.textMd,margin:0}}>Upload Excel · Pivot reports · Drill-down · Crore/Lakh</p>
        </div>
        {err&&<div style={{padding:"8px 12px",background:"rgba(163,45,45,0.09)",border:"1px solid rgba(163,45,45,0.3)",borderRadius:6,fontSize:12,color:T.danger,marginBottom:12}}>{err}</div>}
        <div style={{display:"flex",flexDirection:"column",gap:10,marginBottom:14}}>
          <div>
            <div style={{fontSize:11,color:T.textMd,fontWeight:600,marginBottom:4}}>Username</div>
            <input value={username} onChange={e=>setUsername(e.target.value)} placeholder="Enter username"
              style={inp} onKeyDown={e=>e.key==="Enter"&&tryLogin()}/>
          </div>
          <div>
            <div style={{fontSize:11,color:T.textMd,fontWeight:600,marginBottom:4}}>Password</div>
            <input type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="Enter password"
              style={inp} onKeyDown={e=>e.key==="Enter"&&tryLogin()}/>
          </div>
        </div>
        <button onClick={tryLogin} disabled={loading} style={{width:"100%",padding:"10px",background:loading?"rgba(92,45,26,0.5)":T.primary,color:T.textLt,border:"none",borderRadius:8,cursor:loading?"wait":"pointer",fontSize:14,fontWeight:700}}>
          {loading?"Signing in…":"Sign in"}
        </button>

      </div>
    </div>
  );
}

// ── Reports Manager (Admin tab) ────────────────────────────────────────────────
function ReportsTab({savedReports,onOpen,onDelete,onPublish,publishedId,onReload}) {
  if (!savedReports.length) return(
    <div style={{padding:40,textAlign:"center"}}>
      <div style={{fontSize:40,marginBottom:14}}>📋</div>
      <div style={{fontWeight:700,fontSize:16,color:T.primary,marginBottom:8}}>No saved reports yet</div>
      <div style={{fontSize:13,color:T.textMd}}>Go to the Builder tab, configure your pivot, then click "Save Report".</div>
    </div>
  );
  return(
    <div style={{padding:20,maxWidth:900,margin:"0 auto"}}>
      <div style={{fontWeight:700,fontSize:16,color:T.primary,marginBottom:4}}>Saved Reports</div>
      <div style={{fontSize:12,color:T.textMd,marginBottom:18}}>
        {savedReports.length} report{savedReports.length!==1?"s":""} saved · publish one to make it visible to users
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {savedReports.map(r=>(
          <div key={r.id} style={{background:T.bgCard,border:"1px solid "+(r.id===publishedId?T.primary:T.border),borderRadius:10,padding:"14px 18px",display:"flex",alignItems:"center",gap:14}}>
            <div style={{width:40,height:40,background:r.id===publishedId?T.primary:T.bgStat,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>
              {r.id===publishedId?"📤":"📊"}
            </div>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontWeight:700,fontSize:14,color:T.text,marginBottom:2}}>{r.name}</div>
              <div style={{fontSize:11,color:T.textMd,display:"flex",gap:12,flexWrap:"wrap"}}>
                <span>{r.rows.toLocaleString()} rows</span>
                <span>{r.fields} fields</span>
                <span>Rows: {r.config.rows.join(", ")||"—"}</span>
                <span>Values: {r.config.values.map(v=>v.field).join(", ")||"—"}</span>
                <span>Saved: {new Date(r.savedAt).toLocaleDateString()}</span>
              </div>
            </div>
            <div style={{display:"flex",gap:8,flexShrink:0}}>
              <button onClick={()=>onOpen(r.id)}
                style={{padding:"5px 13px",border:"1px solid "+T.border,borderRadius:6,background:"none",cursor:"pointer",fontSize:12,color:T.text,fontWeight:500}}>
                Open
              </button>
              <button onClick={()=>onPublish(r.id)}
                style={{padding:"5px 13px",border:"1px solid "+(r.id===publishedId?T.primary:T.border),borderRadius:6,
                  background:r.id===publishedId?T.primary:"none",cursor:"pointer",fontSize:12,
                  color:r.id===publishedId?T.textLt:T.text,fontWeight:r.id===publishedId?700:400}}>
                {r.id===publishedId?"Published":"Publish"}
              </button>
              <button onClick={async()=>{if(confirm("Delete report \'"+r.name+"\'?")) await onDelete(r.id);}}
                style={{padding:"5px 10px",border:"1px solid rgba(163,45,45,0.3)",borderRadius:6,background:"none",cursor:"pointer",fontSize:12,color:T.danger}}>
                Delete
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Root ───────────────────────────────────────────────────────────────────────
// ── Helper: parse API report metadata into local shape ────────────────────────
function parseReportMeta(r) {
  return {
    id: r.id,
    name: r.name,
    rows: r.row_count||0,
    fields: r.field_count||0,
    savedAt: r.created_at ? new Date(r.created_at).getTime() : Date.now(),
    config: typeof r.config==="string" ? JSON.parse(r.config) : (r.config||{}),
    cardFields: typeof r.card_fields==="string" ? JSON.parse(r.card_fields) : (r.card_fields||[]),
    isPublished: !!r.is_published,
    dataset: null, // rows loaded lazily on demand
  };
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
    if (token&&role&&username) {
      setCurrentUser(username);
      loadAllReports().then(()=>setScreen(role)).catch(()=>setScreen("login"));
    } else {
      setScreen("login");
    }
  },[]);

  // ── Load report list from API ──────────────────────────────────────────────
  async function loadAllReports() {
    try {
      const list=await getReports();
      const entries=list.map(parseReportMeta);
      setSavedReports(entries);
      const pub=entries.find(r=>r.isPublished);
      setPublishedId(pub?pub.id:null);
    } catch(e) {
      setLoadErr("Could not load reports: "+e.message);
    }
  }

  // ── Lazy-load rows for a specific report ───────────────────────────────────
  async function loadReportData(id) {
    if (dataCache.current[id]) return dataCache.current[id];
    const data=await getReportData(id); // {fields, numFields, rows}
    // numFields comes back as array from JSON, convert to Set
    const nf=new Set(Array.isArray(data.numFields)?data.numFields:Object.values(data.numFields||{}));
    const result={rows:data.rows, fields:data.fields, numFields:nf};
    dataCache.current[id]=result;
    return result;
  }

  // ── Save report → POST to API, then refresh list ───────────────────────────
  async function handleSaveReport(reportData) {
    const {name,dataset,config,cardFields}=reportData;
    const nfArr=[...(dataset.numFields instanceof Set?dataset.numFields:new Set(dataset.numFields||[]))];
    const result=await createReport({
      name,config,cardFields:cardFields||[],
      rows:dataset.rows,fields:dataset.fields,numFields:nfArr
    });
    // Cache the data locally so we don't re-fetch immediately
    dataCache.current[result.id]={rows:dataset.rows,fields:dataset.fields,numFields:dataset.numFields};
    await loadAllReports();
    return result.id;
  }

  // ── Delete report → DELETE from API ───────────────────────────────────────
  async function handleDeleteReport(id) {
    await apiDeleteReport(id);
    delete dataCache.current[id];
    setSavedReports(prev=>prev.filter(r=>r.id!==id));
    if (publishedId===id) setPublishedId(null);
  }

  // ── Publish report → PATCH to API ─────────────────────────────────────────
  async function handlePublishReport(id) {
    await apiPublishReport(id);
    setPublishedId(id);
    setSavedReports(prev=>prev.map(r=>({...r,isPublished:r.id===id})));
  }

  // ── Login / Logout ─────────────────────────────────────────────────────────
  async function doLogin(role,username,token) {
    localStorage.setItem("rh_role",role);
    localStorage.setItem("rh_username",username);
    // token already stored by apiLogin() in api.js
    setCurrentUser(username);
    await loadAllReports();
    setScreen(role);
  }

  function doLogout() {
    apiLogout();
    localStorage.removeItem("rh_role");
    localStorage.removeItem("rh_username");
    setCurrentUser(null);
    setSavedReports([]);
    setPublishedId(null);
    dataCache.current={};
    setScreen("login");
  }

  if (screen==="loading") return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:T.bgPage,flexDirection:"column",gap:12}}>
      <div style={{fontSize:36,animation:"spin 1s linear infinite",display:"inline-block"}}>⚙️</div>
      <div style={{fontWeight:600,color:T.primary}}>Loading ReportHub…</div>
      {loadErr&&<div style={{fontSize:12,color:T.danger,maxWidth:300,textAlign:"center"}}>{loadErr}</div>}
      <style>{"@keyframes spin{to{transform:rotate(360deg)}}"}</style>
    </div>
  );

  return screen==="login"
    ?<Login onLogin={doLogin}/>
    :screen==="admin"
      ?<AdminView
          onLogout={doLogout}
          savedReports={savedReports}
          publishedId={publishedId}
          onSaveReport={handleSaveReport}
          onPublishReport={handlePublishReport}
          onDeleteReport={handleDeleteReport}
          onLoadReportData={loadReportData}
          onReloadReports={loadAllReports}
          currentUser={currentUser}/>
      :<UserView
          onLogout={doLogout}
          savedReports={savedReports}
          publishedReport={publishedReport}
          onLoadReportData={loadReportData}/>;
}
