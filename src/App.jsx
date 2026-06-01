import { useState, useCallback, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

/* ─── Google Fonts injection ─────────────────────────────────────────────── */
const fontLink = document.createElement("link");
fontLink.rel = "stylesheet";
fontLink.href = "https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;1,9..40,300&family=DM+Mono:wght@300;400;500&display=swap";
document.head.appendChild(fontLink);

const SEED_EMPLOYEES = [
  [1,"alvaro cuello"],[2,"noel arguell"],[3,"frank"],[4,"urquia pini"],
  [5,"giusto alfredo"],[6,"sosa jose"],[7,"sergio cabre"],[8,"eder martinez"],
  [9,"robert"],[10,"heber busto"],[11,"javier"],[12,"amaya david"],
  [13,"ferreyra seba"],[14,"moya jonathan"],[15,"romero javier"],[16,"rivas gonzalo"],
  [17,"soule marcos"],[18,"francelys"],[19,"iturra leonel"],[20,"matias perotti"],
  [21,"alejo diaz"],[22,"alvarez gustavo"],[23,"marcelo figuero"],[24,"erick perdenera"],
  [25,"alfonso sergio"],[26,"juan astrada"],[27,"marcelo carrizo"],[28,"CASTRO EZEQUIEL"],
  [29,"zarate sebastia"],[30,"marcos matuchin"],[31,"DAVID MONTES"],[32,"leonel p visval"],
  [33,"eduardo ingar"],[34,"miguel ramirez"],[35,"eduardo herrera"],[36,"ivanna santande"],
  [37,"luis jimenez"],[38,"marcos"],[39,"marcosaldain"],[40,"guille balmaced"],
  [41,"emmanuel torres"],[42,"jose e jimenez"],[43,"ramos nerl"],[44,"leo velez"],
  [45,"gonzalez guille"],[46,"quinteros carlo"],[47,"moyano da"],[48,"saire braian"],
  [49,"campos carlos"],[50,"gabi salas"],[51,"jose martln"],[52,"marcos natali"],
  [53,"david medina"],[55,"cristian g"],[56,"sosa"],[57,"arjona rodrigo"],
  [58,"arjona"],[59,"bianciotti"],[60,"javier d la tor"],[61,"ledesma mar"],
  [62,"pena juan"],[63,"peralta crist"],[64,"juan pena"],[65,"santillan fac"],
  [66,"arrieta paulo"],[67,"pringles fabian"],[68,"romero valentin"],[69,"rodriguez jose"],
  [70,"oliva luca"],[71,"campos mauric"],[72,"favre"],[73,"jara leo"],
  [74,"arjona r"],[75,"jorge oliva"],[76,"oliva jorge"],[77,"walter mont"],
  [78,"ibarra mario"],[79,"segundo claudio"],[80,"fernando cabeza"],[81,"ventura guille"],
  [82,"ventura"],[83,"suarez pablo"],[84,"rojas christ"],[85,"suarez david"],
  [86,"alan fonseca"],[87,"almada seba"],[88,"mansilla lucas"],[89,"bellido alexis"],
  [90,"marchetti chris"],[91,"navarlat david"],[92,"reynoso fernand"],[93,"franco caceres"],
  [94,"toschi nico"],[95,"albornoz marcos"],[96,"aimada franco"],[97,"suarez jose"],
  [98,"luna luis"],[99,"aguero lucas"],[100,"lucas guzman"],[101,"almada"],
  [102,"romero"],[103,"eduardo rivero"],[104,"arguello marian"],[105,"gonza pereyra"],
  [106,"omar retamozo"],[107,"javier papalini"],[108,"ale nieto"],[109,"emiliano domin"],
  [110,"yapura danisa"],[111,"rasilla"],[112,"facu"],[113,"matl"],
  [114,"marina"],[115,"adrian"],[116,"guille v"],[117,"guille ded"],
  [118,"diego vera"],[119,"adrian gomez"],[120,"aguilera joaqun"],[121,"diego nuevo"],
  [122,"orlovas fran"],[123,"gustavo"],[124,"lujan"],[125,"gaston"],
  [126,"gallardo"],[127,"emanuel"],[128,"ruben"],[129,"papurello"],
  [130,"joel caro"],[131,"walter fern"],[132,"leonel quint"],[133,"danelutti crist"],
  [134,"jara maxi"],[135,"luis fernandez"],[136,"pinto PABLO"],[137,"michael"],
  [138,"daneluti"],[139,"hernan silva"],[140,"pinero"],[141,"antonio"],
  [142,"juanjo"],[143,"luduena jonat"],[144,"alan"],[145,"sarfatti lea"],
  [146,"leandro"],[147,"noe"],[148,"fuentes pablo"],[149,"facundo moreno"],
  [150,"nicolas"],[151,"navarro ale"],[152,"pintado daniel"],[153,"miguel villalba"],
  [154,"hugo g"],[155,"rojo hernan"],[156,"valen"],[157,"moreno cristian"],
  [158,"rodrigo juarez"],[159,"celina"],[160,"mitter juan"],[161,"torres marcelo"],
  [162,"seba figueroa"],[163,"ale"],[164,"ale wolff"],[165,"fede martinez"],
  [166,"maxi varela"],[167,"damian diaz"],[168,"gonza gonzalez"],[169,"mariano luna"],
  [170,"porcel j"],[171,"mario vargas"],[172,"basualdo luis"],[173,"liendo marcos"],
  [174,"eber noriega"],[175,"david martinez"],[176,"gerardo ferreyr"],[177,"luis basualdo"],
  [178,"rodrigo romero"],[179,"kevin paez"],[180,"nahuel serrano"],[181,"lucas torres"],
  [182,"morena s"],[183,"matias padilla"],[184,"facundo b"],[185,"oscar"],
  [186,"torres lucas"],[187,"tissera gabriel"],[188,"joaquin"],[189,"joaco mourino"],
  [190,"julio"],[920,"abraham javier"],
];

function makeDefaultEmployees() {
  const map = {};
  for (const [no, nombre] of SEED_EMPLOYEES) {
    map[no] = { empNo: no, nombre, depto: "pyg", entrada: "06:00", salida: "16:30", activo: true, autoDetected: false, tipo: "operario" };
  }
  return map;
}

function parseTimeVal(val) {
  if (val == null) return null;
  if (typeof val === "string") {
    const s = val.trim();
    if (!s || s === "Ausente") return null;
    const m = s.match(/^(\d{1,2}):(\d{2})/);
    if (m) return +m[1] * 60 + +m[2];
  }
  if (typeof val === "number" && !isNaN(val)) {
    if (val >= 0 && val < 1) return Math.round(val * 1440) % 1440;
    return Math.round(val) % 1440;
  }
  return null;
}

function minsToHHMM(m) {
  if (m == null || isNaN(m)) return null;
  const a = Math.abs(Math.round(m));
  return `${String(Math.floor(a/60)).padStart(2,"0")}:${String(a%60).padStart(2,"0")}`;
}

function minsToDisplay(m) {
  if (m == null || isNaN(m)) return "—";
  const s = m < 0 ? "−" : "";
  const a = Math.abs(Math.round(m));
  const h = Math.floor(a/60), min = a%60;
  if (h === 0) return `${s}${min}min`;
  return `${s}${h}h${min > 0 ? String(min).padStart(2,"0") : ""}`;
}

function extractEntradaSalida(row) {
  const marks = [4,5,6,7].map(c=>parseTimeVal(row[c])).filter(v=>v!=null).sort((a,b)=>a-b);
  return {
    entrada: marks.length >= 1 ? minsToHHMM(marks[0]) : null,
    salida:  marks.length >= 2 ? minsToHHMM(marks[marks.length-1]) : null,
    soloEntrada: marks.length === 1,
  };
}

function parseDateVal(val) {
  if (!val) return null;
  if (typeof val === "string") return val.replace(/\//g,"-").slice(0,10);
  if (typeof val === "number") return new Date(Math.round((val-25569)*86400000)).toISOString().slice(0,10);
  if (val instanceof Date) return val.toISOString().slice(0,10);
  return null;
}

function parseAnormalSheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
  const records = [];
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    const empNo = row[0], nombre = row[1], fecha = row[3];
    if (empNo == null || nombre == null) continue;
    const n = Number(empNo);
    if (isNaN(n) || n <= 0) continue;
    const fechaStr = parseDateVal(fecha);
    if (!fechaStr) continue;
    const { entrada, salida, soloEntrada } = extractEntradaSalida(row);
    records.push({ id:`${n}_${fechaStr}`, empNo:n, nombre:String(nombre).trim(), depto:row[2]?String(row[2]).trim():"pyg", fecha:fechaStr, entrada, salida, soloEntrada });
  }
  return records;
}

// Horas extra de operarios solo desde las 06:00
const OPERARIO_EXTRA_FROM = 6 * 60; // 360 min

// Colores por tipo
const TIPO_CFG = {
  operario:      { label:"Operario",      bg:"#edf5ff", color:"#1e5fa8", border:"#c3daff" },
  administrativo:{ label:"Administrativo",bg:"#f0faf4", color:"#276749", border:"#c3e6cb" },
};

function calcRecord(rec, empCfg, specialDays) {
  const cfg = empCfg || { entrada:"06:00", salida:"16:30", tipo:"operario" };
  const esOperario = (cfg.tipo || "operario") === "operario";
  const dayType    = specialDays?.[rec.fecha];
  const diaSemana  = new Date(rec.fecha+"T12:00:00").getDay();
  const esSabado   = diaSemana === 6;
  const esDomingo  = diaSemana === 0;
  const esFeriado  = dayType?.tipo === "feriado";
  // Operarios no acumulan horas extra en sábados, domingos ni feriados
  const sinExtra   = esOperario && (esSabado || esDomingo || esFeriado);
  const effectiveCfg = esFeriado
    ? { entrada: cfg.entrada, salida: "14:00" }
    : esSabado
    ? { entrada: cfg.entrada, salida: "13:00" }
    : cfg;
  if (!rec.entrada || !rec.salida) return { trabajado:null, jornada:null, extra:null, demora:null, salTemprana:null };
  const entMin = parseTimeVal(rec.entrada);
  const salMin = parseTimeVal(rec.salida);
  const entRef = parseTimeVal(effectiveCfg.entrada);
  const salRef = parseTimeVal(effectiveCfg.salida);
  const jornada = salRef - entRef;
  // Horas extra
  let extra = null;
  if (!sinExtra) {
    // Operarios: horas extra desde 06:00 am como mínimo
    // Administrativos: horas extra desde su hora de entrada configurada (no antes)
    const limiteEntrada = esOperario ? OPERARIO_EXTRA_FROM : entRef;
    const entEfectiva   = Math.max(entMin, limiteEntrada);
    const adelanto      = Math.max(0, entRef - entEfectiva); // siempre 0 con este límite
    const extension     = Math.max(0, salMin - salRef);
    const total         = adelanto + extension;
    if (total > 0) extra = total;
  }
  const demora      = Math.max(0, entMin - entRef);
  const salTemprana = Math.max(0, salRef - salMin);
  const entDentro   = Math.max(entMin, entRef);
  const salDentro   = Math.min(salMin, salRef);
  const trabajado   = Math.max(0, salDentro - entDentro);
  return { trabajado, jornada, extra, demora, salTemprana };
}

function detectSchedule(recs) {
  const ents=recs.map(r=>r.entrada?parseTimeVal(r.entrada):null).filter(v=>v!=null).sort((a,b)=>a-b);
  const sals=recs.map(r=>r.salida?parseTimeVal(r.salida):null).filter(v=>v!=null).sort((a,b)=>a-b);
  if (!ents.length) return null;
  return { entrada:minsToHHMM(ents[Math.floor(ents.length/2)]), salida:sals.length?minsToHHMM(sals[Math.floor(sals.length/2)]):"16:30" };
}

function loadLS(key, fb) { try { const r=localStorage.getItem(key); if(r) return JSON.parse(r); } catch {} return fb; }
function saveLS(key, v)  { try { localStorage.setItem(key,JSON.stringify(v)); } catch {} }

const cap = s => s ? s.charAt(0).toUpperCase()+s.slice(1).toLowerCase() : s;

const COL = {
  bg:"#f5f6f8", surface:"#ffffff", border:"#e8ecf1", border2:"#d4dbe4",
  text:"#1e2a38", textSub:"#5a6a7a", textFaint:"#96a3b0",
  accent:"#3d6b9e", accentBg:"#edf2f9", accentSoft:"#dce8f5",
};
const SANS = "'DM Sans', system-ui, sans-serif";
const MONO = "'DM Mono', 'Courier New', monospace";

/* ─── PDF export ─────────────────────────────────────────────────────────── */
function exportLiqPDF(d) {
  const { selEmp, periodo, ingreso, desde, hasta, importeSueldo, diasFinde, valorDiaFinde, importeFinde, horasExtra, horasExtraDisplay, valorHoraExt,
    importeExtras, feriados, valorDia, importeFeriados, sac, vacaciones,
    importeVacaciones, totalAdicionales, subtotal, fraccionesDemora, valorHora,
    descDemoras, horasSalTemp, descSalTemp, totalDescuentos, adelanto,
    totalACobrar, diasTrabajados, nombreDisplay, fmt } = d;

  const nombre = selEmp.nombre.toUpperCase();
  const today  = new Date().toLocaleDateString("es-AR");

  // Try to load logo as base64 from /logo.png (served by Vite from /public)
  const renderPDF = (logoSrc) => {
    const GOLD = "#b08a2e";
    const DARK = "#1a1a1a";

    const logoHTML = logoSrc
      ? `<img src="${logoSrc}" style="height:52px;width:auto;object-fit:contain;display:block;" alt="logo"/>`
      : `<span style="font-size:20px;font-weight:900;color:#1a1a1a;font-family:Arial,sans-serif">PyG<span style="color:#b08a2e">SERVICIOS</span></span>`;

    const row = (label, cant, val, imp, type, color, indent) => {
      const styles = {
        section:     { bg:"#f5f0e8", fw:"700", fs:"11px",   cl:color||"#1a1a1a", bd:"border-bottom:1px solid #d9cfa8" },
        sub:         { bg:"#efefef", fw:"700", fs:"11px",   cl:color||"#1a1a1a", bd:"border-bottom:1px solid #ddd" },
        "total-line":{ bg:"#fff",   fw:"700", fs:"11.5px", cl:color||"#1a3a6b", bd:"border-bottom:2px solid #ccc" },
        detail:      { bg:"#fff",   fw:"400", fs:"11px",   cl:color||"#333",    bd:"border-bottom:1px solid #eee" },
        muted:       { bg:"#fff",   fw:"400", fs:"11px",   cl:"#bbb",           bd:"border-bottom:1px solid #eee" },
      };
      const s = styles[type] || styles.detail;
      const pl = indent ? "24px" : "10px";
      return `<tr style="background:${s.bg};${s.bd}">
        <td style="padding:6px ${pl};font-weight:${s.fw};color:${s.cl};font-size:${s.fs}">${label}</td>
        <td style="padding:6px 8px;text-align:center;font-size:10.5px;color:${type==="muted"?"#ccc":"#555"}">${cant||"—"}</td>
        <td style="padding:6px 8px;text-align:center;font-size:10.5px;color:${type==="muted"?"#ccc":"#555"}">${val||"—"}</td>
        <td style="padding:6px 10px;text-align:right;font-weight:${s.fw};font-size:${s.fs};color:${s.cl}">${imp||"—"}</td>
      </tr>`;
    };

    const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
      @page { size:A4 portrait; margin:14mm 10mm; }
      *{box-sizing:border-box;margin:0;padding:0}
      body{font-family:'Helvetica Neue',Arial,sans-serif;background:#fff;color:#1a1a1a;font-size:11.5px}
      .wrap{max-width:76%;margin:0 auto;border:1px solid #ccc;padding:18px 22px 22px}
      table{width:100%;border-collapse:collapse}
      .hdr{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px}
      .hdr-left{display:flex;gap:16px;align-items:flex-start}
      .hdr-mid{font-size:10px;color:#444;line-height:1.9}
      .hdr-right{text-align:right;font-size:10px;color:#444;line-height:1.9;border-left:1px solid #ddd;padding-left:14px}
      .hdr-right b{color:#1a1a1a}
      .gold-line{border:none;border-top:3px solid #b08a2e;margin:10px 0 14px}
      .title-row{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:14px}
      .doc-title{font-size:15px;font-weight:700;color:#1a1a1a;text-transform:uppercase;letter-spacing:0.3px}
      .doc-meta{font-size:10px;color:#555;text-align:right;line-height:1.9}
      .client-box{background:#f7f5ef;border:1px solid #ddd;padding:8px 0;margin-bottom:14px;display:flex}
      .cc{border-right:1px solid #ddd;padding:4px 14px}
      .cc:last-child{border-right:none}
      .cf{font-size:8.5px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:3px}
      .cv{font-size:12px;font-weight:700;color:#1a1a1a}
      .tbl-head tr{background:#1a1a1a;color:#fff}
      .tbl-head th{padding:7px 10px;font-size:9.5px;font-weight:700;letter-spacing:0.07em;text-transform:uppercase}
      .tbl-head th:not(:first-child){text-align:center}
      .tbl-head th:last-child{text-align:right}
      .total-gold td{padding:9px 10px;font-size:12.5px;font-weight:700;color:#fff;background:#b08a2e}
      .total-gold td:last-child{text-align:right;font-size:13px}
      .firma-box{display:flex;justify-content:space-between;margin-top:28px}
      .ff{width:44%}
      .fl{font-size:8.5px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:22px}
      .fline{border-bottom:1px solid #999;margin-bottom:4px}
      .fsub{font-size:9px;color:#888}
      @media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
    </style></head><body>
    <div class="wrap">

      <div class="hdr">
        <div class="hdr-left">
          ${logoHTML}
          <div class="hdr-mid">            
          </div>
        </div>
        <div class="hdr-right">          
        </div>
      </div>

      <hr class="gold-line"/>

      <div class="title-row">
        <div class="doc-title">Planilla Liquidación Sueldos</div>
        <div class="doc-meta">
          Período: <b>${periodo||"—"}</b><br>
          ${(desde||hasta)?`${desde||"inicio"} → ${hasta||"fin"}<br>`:""}
          Emitido: ${today}
        </div>
      </div>

      <div class="client-box">
        <div class="cc"><div class="cf">Empleado</div><div class="cv">${nombreDisplay}</div></div>
        <div class="cc"><div class="cf">Fecha de ingreso</div><div class="cv">${ingreso||"—"}</div></div>
        <div class="cc"><div class="cf">Horario</div><div class="cv">${selEmp.entrada} – ${selEmp.salida}</div></div>
      </div>

      <table>
        <thead class="tbl-head"><tr>
          <th style="text-align:left">Elemento / Descripción</th>
          <th>Cantidad</th>
          <th>Valor unit.</th>
          <th style="text-align:right">Precio Total</th>
        </tr></thead>
        <tbody>
          ${row("SUELDO BÁSICO","—","—",fmt(importeSueldo),"sub")}
          ${row("ADICIONALES","","","","section")}
          ${horasExtra>0  ?row("Horas extra",horasExtraDisplay,fmt(valorHoraExt),fmt(importeExtras),"detail","",true):""}
          ${diasFinde>0   ?row("Días finde/especiales",diasFinde,fmt(valorDiaFinde),fmt(importeFinde),"detail","",true):""}
          ${feriados>0    ?row("Feriados",feriados,fmt(valorDia),fmt(importeFeriados),"detail","",true):row("Feriados","—",fmt(valorDia),"—","muted","",true)}
          ${row("SAC","—","—",sac>0?fmt(sac):"—",sac>0?"detail":"muted","",true)}
          ${vacaciones>0  ?row("Vacaciones",vacaciones,fmt(valorDia),fmt(importeVacaciones),"detail","",true):row("Vacaciones","—",fmt(valorDia),"—","muted","",true)}
          ${row("Subtotal adicionales","","",fmt(totalAdicionales),"sub")}
          ${row("SUELDO + ADICIONALES","","",fmt(subtotal),"total-line","#1a3a6b")}
          ${row("DESCUENTOS","","","","section")}
          ${fraccionesDemora>0?row("Llegadas tarde (fracc. de 15 min)",fraccionesDemora,`${fmt(valorHora)}/4`,fmt(descDemoras),"detail","#c53030",true):row("Llegadas tarde (fracc. de 15 min)","—","—","—","muted","",true)}
          ${horasSalTemp>0?row("Retiros anticipados (x hora)",horasSalTemp,fmt(valorHora),fmt(descSalTemp),"detail","#c53030",true):row("Retiros anticipados (x hora)","—","—","—","muted","",true)}
          ${row("Total descuentos","","",totalDescuentos>0?fmt(totalDescuentos):"—","sub",totalDescuentos>0?"#c53030":"")}
          ${adelanto>0?row("Adelanto","—","—",fmt(adelanto),"sub","#b45309"):row("Adelanto","—","—","—","muted")}
          <tr class="total-gold">
            <td colspan="3" style="text-align:right;padding-right:16px;letter-spacing:0.04em">TOTAL A COBRAR — ${periodo||""}</td>
            <td>${totalACobrar>0?fmt(totalACobrar):"—"}</td>
          </tr>
        </tbody>
      </table>

      <div class="firma-box">
        <div class="ff">
          <div class="fl">Firma empleador</div>
          <div class="fline"></div>          
        </div>
      </div>

    </div>
    </body></html>`;

    const win = window.open("", "_blank", "width=920,height=780");
    win.document.write(html);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); }, 700);
  };

  // Attempt to load logo from /logo.png (place file in /public folder of Vite project)
  fetch("/logo.png")
    .then(r => r.ok ? r.blob() : Promise.reject())
    .then(blob => {
      const reader = new FileReader();
      reader.onload = () => renderPDF(reader.result);
      reader.onerror = () => renderPDF(null);
      reader.readAsDataURL(blob);
    })
    .catch(() => renderPDF(null));
}


/* ─── Supabase ───────────────────────────────────────────────────────────── */
const SB_URL = import.meta.env.VITE_SUPABASE_URL;
const SB_KEY = import.meta.env.VITE_SUPABASE_SERVICE_KEY;
const SB_ANON = import.meta.env.VITE_SUPABASE_ANON_KEY;

const SB_HEADERS = (write=false) => ({
  "Content-Type": "application/json",
  "Authorization": `Bearer ${write ? SB_KEY : (SB_ANON||SB_KEY)}`,
  "apikey": write ? SB_KEY : (SB_ANON||SB_KEY),
  "Accept-Profile": "rrhh",
  ...(write ? {"Content-Profile":"rrhh"} : {}),
});

// ── Read all rows from a table ────────────────────────────────────────────
async function sbFetch(table, params="") {
  if (!SB_URL || !(SB_KEY||SB_ANON)) return null;
  try {
    const res = await fetch(`${SB_URL}/rest/v1/${table}?${params}`, {
      headers: SB_HEADERS(false),
    });
    if (!res.ok) return null;
    return await res.json();
  } catch { return null; }
}

// ── Upsert many rows ─────────────────────────────────────────────────────
async function sbUpsert(table, rows, pk="id") {
  if (!SB_URL || !SB_KEY || !rows?.length) return;
  try {
    await fetch(`${SB_URL}/rest/v1/${table}?on_conflict=${pk}`, {
      method: "POST",
      headers: { ...SB_HEADERS(true), "Prefer": "resolution=merge-duplicates" },
      body: JSON.stringify(rows),
    });
  } catch {}
}

// ── Upsert single row ────────────────────────────────────────────────────
async function sbUpsertSingle(table, row, pk="id") {
  if (!SB_URL || !SB_KEY) return;
  try {
    await fetch(`${SB_URL}/rest/v1/${table}?on_conflict=${pk}`, {
      method: "POST",
      headers: { ...SB_HEADERS(true), "Prefer": "resolution=merge-duplicates" },
      body: JSON.stringify(row),
    });
  } catch {}
}

// ── Delete row ───────────────────────────────────────────────────────────
async function sbDelete(table, id, pk="id") {
  if (!SB_URL || !SB_KEY) return;
  try {
    await fetch(`${SB_URL}/rest/v1/${table}?${pk}=eq.${encodeURIComponent(id)}`, {
      method: "DELETE",
      headers: SB_HEADERS(true),
    });
  } catch {}
}

// ── Realtime subscription via Supabase websocket ─────────────────────────
function sbSubscribe(table, onInsert, onUpdate, onDelete) {
  if (!SB_URL || !(SB_KEY||SB_ANON)) return () => {};
  const wsUrl = SB_URL.replace("https://","wss://").replace("http://","ws://");
  const key   = SB_ANON || SB_KEY;
  let ws, heartbeat;

  const connect = () => {
    ws = new WebSocket(`${wsUrl}/realtime/v1/websocket?apikey=${key}&vsn=1.0.0`);

    ws.onopen = () => {
      ws.send(JSON.stringify({topic:"realtime:public",event:"phx_join",payload:{},ref:"1"}));
      ws.send(JSON.stringify({
        topic:"realtime:public",event:"phx_join",
        payload:{config:{broadcast:{self:false},presence:{key:""},postgres_changes:[
          {event:"INSERT",schema:"rrhh",table},
          {event:"UPDATE",schema:"rrhh",table},
          {event:"DELETE",schema:"rrhh",table},
        ]}},ref:"2"
      }));
      heartbeat = setInterval(()=>ws.readyState===1&&ws.send(JSON.stringify({topic:"phoenix",event:"heartbeat",payload:{},ref:"hb"})),25000);
    };

    ws.onmessage = (e) => {
      try {
        const msg = JSON.parse(e.data);
        const type = msg.payload?.data?.type;
        const rec  = msg.payload?.data?.record;
        const old  = msg.payload?.data?.old_record;
        if (type==="INSERT"&&rec&&onInsert) onInsert(rec);
        if (type==="UPDATE"&&rec&&onUpdate) onUpdate(rec);
        if (type==="DELETE"&&old&&onDelete) onDelete(old);
      } catch {}
    };

    ws.onclose = () => {
      clearInterval(heartbeat);
      setTimeout(connect, 3000); // reconnect
    };
  };

  connect();
  return () => { clearInterval(heartbeat); ws?.close(); };
}

// ── Row mappers ──────────────────────────────────────────────────────────
function recToRow(r, periodo) {
  return {
    id: r.id, emp_no: r.empNo, nombre: r.nombre, depto: r.depto,
    fecha: r.fecha, entrada: r.entrada, salida: r.salida,
    solo_entrada: r.soloEntrada || false,
    manual: r.manual || false,
    periodo: periodo || r.fecha?.slice(0,7),
  };
}

function rowToRec(r) {
  return {
    id: r.id, empNo: r.emp_no, nombre: r.nombre, depto: r.depto,
    fecha: r.fecha, entrada: r.entrada, salida: r.salida,
    soloEntrada: r.solo_entrada, manual: r.manual,
  };
}

function empToRow(e) {
  return {
    emp_no: e.empNo, nombre: e.nombre,
    nombre_display: e.nombreDisplay || null,
    depto: e.depto, entrada_ref: e.entrada, salida_ref: e.salida,
    activo: e.activo !== false,
    auto_detected: e.autoDetected || false,
  };
}

function rowToEmp(r) {
  return {
    empNo: r.emp_no, nombre: r.nombre,
    nombreDisplay: r.nombre_display,
    depto: r.depto, entrada: r.entrada_ref, salida: r.salida_ref,
    activo: r.activo, autoDetected: r.auto_detected,
  };
}

function diaToRow(fecha, tipo) { return { fecha, tipo }; }


const TABS = ["Importar","Registros","Empleados","Resumen","Por empleado","Calendario","Circular","Liquidación"];

export default function App() {
  const [tab, setTab]             = useState(0);
  const [records, setRecords]     = useState(()=>loadLS("ar3",[]));
  const [employees, setEmployees] = useState(()=>{
    const saved = loadLS("ae3", null);
    if (saved) {
      const migrated = {};
      for (const [k,v] of Object.entries(saved)) migrated[k] = { tipo:"operario", ...v };
      return migrated;
    }
    return makeDefaultEmployees();
  });
  const [empTipoF, setEmpTipoF] = useState("todos");
  const [sbLoading, setSbLoading] = useState(false);
  const [sbStatus,  setSbStatus]  = useState(""); // "synced" | "error" | ""
  const [sbLastSync,setSbLastSync]= useState(null);
  const [importing, setImporting] = useState(false);
  const [importMsg, setImportMsg] = useState({text:"",ok:true});
  const [recF, setRecF]           = useState({emp:"",fecha:""});
  const [empF, setEmpF]           = useState("");
  const [detalleEmp, setDetalleEmp] = useState(null); // empNo selected in Por empleado tab
  const [liqEmp, setLiqEmp]         = useState(null);   // empNo selected in Liquidación tab
  const [circularF, setCircularF]   = useState("");      // filter in Circular tab
  const [circularEdit, setCircularEdit] = useState(null); // empNo being edited in Circular
  const [circularDraft, setCircularDraft] = useState({});
  const [specialDays, setSpecialDays] = useState(()=>loadLS("sp_days",{})); // fecha→{tipo}
  const [manualSalidas, setManualSalidas] = useState(()=>loadLS("man_sal",{})); // id→salida manual
  const [manualRecords, setManualRecords] = useState(()=>loadLS("man_rec",[]));   // fully manual records
  const [editingCell, setEditingCell]     = useState(null); // {id, field}
  const [addingRec, setAddingRec]         = useState(false);
  const [newRec, setNewRec]               = useState({empNo:"",fecha:"",entrada:"",salida:""});
  const [liqParams, setLiqParams]   = useState(() => {
    try { const r = localStorage.getItem("liq_params"); if (r) return JSON.parse(r); } catch {}
    return {};
  });
  const [editingEmp, setEditingEmp] = useState(null);
  const [editDraft, setEditDraft]   = useState({});
  const [bulkH, setBulkH]         = useState({entrada:"06:00",salida:"16:30"});
  const [bulkSel, setBulkSel]     = useState(new Set());
  const fileRef = useRef();

  // ── Persist to localStorage as cache ────────────────────────────────────
  useEffect(()=>saveLS("ar3",records),[records]);
  useEffect(()=>saveLS("ae3",employees),[employees]);
  useEffect(()=>saveLS("liq_params",liqParams),[liqParams]);
  useEffect(()=>saveLS("sp_days",specialDays),[specialDays]);
  useEffect(()=>saveLS("man_sal",manualSalidas),[manualSalidas]);
  useEffect(()=>saveLS("man_rec",manualRecords),[manualRecords]);

  // ── Load from Supabase on mount ──────────────────────────────────────────
  useEffect(()=>{
    if (!SB_URL || !(SB_KEY||SB_ANON)) return;
    setSbLoading(true);
    Promise.all([
      sbFetch("registros","select=*&limit=10000"),
      sbFetch("empleados","select=*"),
      sbFetch("dias_especiales","select=*"),
      sbFetch("correcciones","select=*"),
    ]).then(([regs,emps,dias,corrs])=>{
      if (regs?.length) {
        const byId={};
        for(const r of regs) byId[r.id]=rowToRec(r);
        setRecords(Object.values(byId));
      }
      if (emps?.length) {
        const map={...makeDefaultEmployees()};
        for(const e of emps) map[e.emp_no]={...map[e.emp_no],...rowToEmp(e)};
        setEmployees(map);
      }
      if (dias?.length) {
        const map={};
        for(const d of dias) map[d.fecha]={tipo:d.tipo};
        setSpecialDays(map);
      }
      if (corrs?.length) {
        const salMap={}, entMap={};
        for(const c of corrs){
          if(c.salida_corr) salMap[c.id]=c.salida_corr;
          if(c.entrada_corr) entMap[c.id+"_ent"]=c.entrada_corr;
        }
        setManualSalidas(prev=>({...prev,...salMap,...entMap}));
      }
      setSbLoading(false);
      setSbStatus("synced");
      setSbLastSync(new Date());
    }).catch(()=>{ setSbLoading(false); setSbStatus("error"); });
  },[]);

  // ── Realtime subscriptions ───────────────────────────────────────────────
  useEffect(()=>{
    if (!SB_URL || !(SB_KEY||SB_ANON)) return;

    const unsubRegs = sbSubscribe("registros",
      (r)=>setRecords(p=>{const m={};for(const x of p)m[x.id]=x;m[r.id]=rowToRec(r);return Object.values(m);}),
      (r)=>setRecords(p=>p.map(x=>x.id===r.id?{...x,...rowToRec(r)}:x)),
      (r)=>setRecords(p=>p.filter(x=>x.id!==r.id))
    );

    const unsubEmps = sbSubscribe("empleados",
      (r)=>setEmployees(p=>({...p,[r.emp_no]:{...(p[r.emp_no]||{}),...rowToEmp(r)}})),
      (r)=>setEmployees(p=>({...p,[r.emp_no]:{...(p[r.emp_no]||{}),...rowToEmp(r)}})),
      null
    );

    const unsubDias = sbSubscribe("dias_especiales",
      (r)=>setSpecialDays(p=>({...p,[r.fecha]:{tipo:r.tipo}})),
      (r)=>setSpecialDays(p=>({...p,[r.fecha]:{tipo:r.tipo}})),
      (r)=>setSpecialDays(p=>{const u={...p};delete u[r.fecha];return u;})
    );

    return ()=>{ unsubRegs(); unsubEmps(); unsubDias(); };
  },[]);

  const handleFile = useCallback(async(e)=>{
    const file=e.target.files[0]; if(!file) return;
    setImporting(true); setImportMsg({text:"Procesando...",ok:true});
    try {
      const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:false});
      const wsName=wb.SheetNames.find(n=>n==="Anormal")||wb.SheetNames[0];
      const parsed=parseAnormalSheet(wb.Sheets[wsName]);
      setRecords(prev=>{const m={};for(const r of prev)m[r.id]=r;for(const r of parsed)m[r.id]=r;return Object.values(m);});
      setEmployees(prev=>{
        const u={...prev},byEmp={};
        for(const r of parsed){if(!byEmp[r.empNo])byEmp[r.empNo]=[];byEmp[r.empNo].push(r);}
        for(const[noStr,recs]of Object.entries(byEmp)){
          const no=Number(noStr);
          if(!u[no]){const d=detectSchedule(recs);u[no]={empNo:no,nombre:recs[0].nombre,depto:recs[0].depto,entrada:d?.entrada||"06:00",salida:d?.salida||"16:30",activo:true,autoDetected:true,tipo:"operario"};}
        }
        return u;
      });
      setImportMsg({text:`${parsed.length} registros importados desde "${wsName}"`,ok:true});
      // Sync employees to Supabase silently
      setEmployees(prev => {
        const snap = Object.values(prev);
        sbUpsert("empleados", snap.map(empToRow));
        return prev;
      });
      // Sync to Supabase silently
      const periodo = parsed[0]?.fecha?.slice(0,7);
      sbUpsert("registros", parsed.map(r=>recToRow(r,periodo)));
    } catch(err){setImportMsg({text:`Error: ${err.message}`,ok:false});}
    setImporting(false);
    if(fileRef.current)fileRef.current.value="";
  },[]);

  const startEdit=emp=>{setEditingEmp(emp.empNo);setEditDraft({nombre:emp.nombre,depto:emp.depto,entrada:emp.entrada,salida:emp.salida,tipo:emp.tipo||"operario"});};
  const saveEdit=no=>{
    const updated = {...employees[no],...editDraft,autoDetected:false};
    setEmployees(p=>({...p,[no]:updated}));
    setEditingEmp(null);
    sbUpsertSingle("empleados", empToRow(updated), "emp_no");
  };
  const cancelEdit=()=>setEditingEmp(null);
  const applyBulk=()=>{
    if(!bulkSel.size)return;
    const updated={};
    setEmployees(p=>{
      const u={...p};
      for(const no of bulkSel) if(u[no]){u[no]={...u[no],...bulkH,autoDetected:false};updated[no]=u[no];}
      return u;
    });
    setBulkSel(new Set());
    sbUpsert("empleados", Object.values(updated).map(empToRow));
  };

  const toggleTipo = (empNo) => {
    setEmployees(p => ({
      ...p,
      [empNo]: { ...p[empNo], tipo: (p[empNo].tipo||"operario")==="operario" ? "administrativo" : "operario" }
    }));
  };

  const empList=Object.values(employees).sort((a,b)=>a.empNo-b.empNo);
  const filteredEmps=empList.filter(e=>{
    const matchSearch = !empF||e.nombre.toLowerCase().includes(empF.toLowerCase())||String(e.empNo).includes(empF);
    const matchTipo   = empTipoF==="todos"||(e.tipo||"operario")===empTipoF;
    return matchSearch && matchTipo;
  });
  const allRecs = [...records, ...manualRecords];
  const filteredRecs=allRecs
    .filter(r=>(!recF.emp||r.nombre.toLowerCase().includes(recF.emp.toLowerCase())||String(r.empNo).includes(recF.emp))&&(!recF.fecha||r.fecha.includes(recF.fecha)))
    .sort((a,b)=>a.fecha.localeCompare(b.fecha)||a.empNo-b.empNo);

  const empSummary=empList.map(emp=>{
    const recs=allRecs.filter(r=>r.empNo===emp.empNo);
    const calcs=recs.map(r=>{
      const actualEntrada = r.manual ? r.entrada : (manualSalidas[r.id+"_ent"] || r.entrada);
      const actualSalida  = r.manual ? r.salida  : (manualSalidas[r.id]        || r.salida);
      const rr={...r, entrada:actualEntrada, salida:actualSalida, soloEntrada:!actualSalida&&!actualEntrada};
      return {...rr, ...calcRecord(rr, emp, specialDays)};
    });
    const trabajados=calcs.filter(r=>r.trabajado!=null);
    const totalMin=trabajados.reduce((s,r)=>s+r.trabajado,0);
    const demoraDias=calcs.filter(r=>r.demora>0);
    const totalDemora=demoraDias.reduce((s,r)=>s+r.demora,0);
    const stDias=calcs.filter(r=>r.salTemprana>0);
    const peorDemora=demoraDias.reduce((m,r)=>Math.max(m,r.demora),0);
    return{emp,calcs,dias:trabajados.length,totalMin,demoraDias:demoraDias.length,totalDemora,stDias:stDias.length,peorDemora};
  }).filter(s=>s.dias>0);

  const S = {
    root:{fontFamily:SANS,background:COL.bg,minHeight:"100vh",color:COL.text},
    header:{background:COL.surface,borderBottom:`1px solid ${COL.border}`,padding:"14px 28px",display:"flex",alignItems:"center",gap:12},
    logo:{fontSize:15,fontWeight:600,color:COL.text,display:"flex",alignItems:"center",gap:9,letterSpacing:"-0.2px"},
    dot:{width:8,height:8,borderRadius:"50%",background:COL.accent},
    hRight:{marginLeft:"auto",display:"flex",gap:8},
    chip:{background:COL.accentBg,color:COL.accent,padding:"4px 12px",borderRadius:20,fontSize:12,fontWeight:500},
    tabBar:{background:COL.surface,borderBottom:`1px solid ${COL.border}`,padding:"0 24px",display:"flex",gap:0},
    tab:{padding:"13px 18px",background:"none",border:"none",cursor:"pointer",color:COL.textFaint,fontSize:13,fontFamily:SANS,fontWeight:400,position:"relative",transition:"color .15s"},
    tabOn:{color:COL.text,fontWeight:600},
    tabLine:{position:"absolute",bottom:0,left:18,right:18,height:2,background:COL.accent,borderRadius:2},
    main:{padding:"28px 28px 60px"},
    body:{color:COL.textSub,fontSize:14,lineHeight:1.8,marginBottom:20,marginTop:0},
    drop:{border:`2px dashed ${COL.border2}`,borderRadius:12,padding:"36px 24px",textAlign:"center",display:"flex",flexDirection:"column",alignItems:"center",gap:10,marginBottom:16,cursor:"pointer",background:COL.surface,transition:"all .15s"},
    alertOk:{background:"#f0faf4",color:"#276749",border:"1px solid #c3e6cb",padding:"11px 16px",borderRadius:8,fontSize:13,marginBottom:14},
    alertErr:{background:"#fff5f5",color:"#c53030",border:"1px solid #fed7d7",padding:"11px 16px",borderRadius:8,fontSize:13,marginBottom:14},
    infoBox:{background:COL.surface,border:`1px solid ${COL.border}`,borderRadius:10,padding:"16px 20px",marginTop:16},
    filterBar:{display:"flex",gap:10,marginBottom:14,alignItems:"center",flexWrap:"wrap"},
    sInput:{background:COL.surface,border:`1px solid ${COL.border2}`,borderRadius:8,color:COL.text,padding:"8px 12px",fontFamily:SANS,fontSize:13,outline:"none",width:220},
    dateInput:{background:COL.surface,border:`1px solid ${COL.border2}`,borderRadius:8,color:COL.text,padding:"8px 12px",fontFamily:SANS,fontSize:13,outline:"none"},
    table:{width:"100%",borderCollapse:"collapse",fontSize:13,background:COL.surface},
    thr:{},
    th:{padding:"10px 14px",color:COL.textFaint,fontSize:11,fontWeight:600,textAlign:"center",whiteSpace:"nowrap",letterSpacing:"0.04em",background:"#f7f8fa",borderBottom:`1px solid ${COL.border}`},
    td:{padding:"9px 14px",color:COL.textSub,textAlign:"center",borderBottom:`1px solid ${COL.border}`,whiteSpace:"nowrap"},
    tblWrap:{borderRadius:10,border:`1px solid ${COL.border}`,overflow:"hidden"},
    bulkBar:{display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,marginBottom:14,background:COL.surface,padding:"12px 16px",borderRadius:10,border:`1px solid ${COL.border}`,flexWrap:"wrap"},
    tInp:{background:COL.accentBg,border:`1px solid ${COL.accentSoft}`,borderRadius:6,color:COL.accent,padding:"5px 8px",fontFamily:MONO,fontSize:13,outline:"none",width:96},
    inlineInput:{background:COL.accentBg,border:`1px solid ${COL.accentSoft}`,borderRadius:6,color:COL.text,padding:"4px 8px",fontFamily:SANS,fontSize:13,outline:"none",width:"100%"},
    saveBtn:{background:COL.accent,color:"#fff",border:"none",borderRadius:6,padding:"5px 14px",cursor:"pointer",fontFamily:SANS,fontSize:12,fontWeight:600},
    cancelBtn:{background:"transparent",color:COL.textFaint,border:`1px solid ${COL.border2}`,borderRadius:6,padding:"5px 12px",cursor:"pointer",fontFamily:SANS,fontSize:12},
    editBtn:{background:"transparent",color:COL.textSub,border:`1px solid ${COL.border}`,borderRadius:6,padding:"5px 12px",cursor:"pointer",fontFamily:SANS,fontSize:12},
    btnP:{background:COL.accent,color:"#fff",border:"none",borderRadius:8,padding:"9px 20px",cursor:"pointer",fontFamily:SANS,fontWeight:600,fontSize:13},
    btnD:{background:"#fff5f5",color:"#c53030",border:"1px solid #fed7d7",borderRadius:8,padding:"9px 18px",cursor:"pointer",fontFamily:SANS,fontWeight:600,fontSize:13},
    btnS:{background:COL.accent,color:"#fff",border:"none",borderRadius:7,padding:"6px 14px",cursor:"pointer",fontFamily:SANS,fontWeight:600,fontSize:12},
    sug:{background:COL.surface,border:`1px solid ${COL.border}`,borderRadius:8,color:COL.textSub,padding:"10px 14px",fontSize:12,cursor:"pointer",fontFamily:SANS,textAlign:"left",lineHeight:1.5},
  };

  const H2 = ({children})=><h2 style={{fontFamily:SANS,fontSize:19,fontWeight:600,color:COL.text,margin:"0 0 6px",letterSpacing:"-0.3px"}}>{children}</h2>;

  const TipoBadge = ({tipo, onClick, small}) => {
    const cfg = TIPO_CFG[tipo||"operario"];
    return (
      <span onClick={onClick} title={onClick?"Click para cambiar tipo":undefined}
        style={{display:"inline-flex",alignItems:"center",background:cfg.bg,color:cfg.color,
          border:`1px solid ${cfg.border}`,borderRadius:20,padding:small?"2px 8px":"3px 10px",
          fontSize:small?10:11,fontWeight:600,fontFamily:SANS,cursor:onClick?"pointer":"default",
          letterSpacing:"0.02em",whiteSpace:"nowrap",transition:"opacity .15s",userSelect:"none"}}>
        {cfg.label}
      </span>
    );
  };
  const Code = ({children})=><code style={{fontFamily:MONO,background:COL.accentBg,color:COL.accent,padding:"1px 6px",borderRadius:4,fontSize:12}}>{children}</code>;
  const Num = ({children})=><span style={{fontFamily:MONO,fontSize:11,color:COL.textFaint,background:COL.accentBg,padding:"1px 6px",borderRadius:4}}>{children}</span>;
  const TimeTag = ({children,color})=><span style={{fontFamily:MONO,fontSize:13,fontWeight:500,color:color==="g"?"#276749":color==="b"?"#1e5fa8":COL.textSub,background:color==="g"?"#f0faf4":color==="b"?"#ebf4ff":"transparent",padding:"2px 8px",borderRadius:5}}>{children}</span>;

  const THead = ({cols})=><thead><tr>{cols.map((h,i)=><th key={i} style={{...S.th,textAlign:i<=1?"left":undefined}}>{h}</th>)}</tr></thead>;

  const KpiCard = ({label,value,color})=>(
    <div style={{background:COL.surface,border:`1px solid ${COL.border}`,borderRadius:12,padding:"18px 20px",borderTop:`3px solid ${color}`}}>
      <div style={{fontSize:11,color:COL.textFaint,fontWeight:500,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:8}}>{label}</div>
      <div style={{fontSize:30,fontWeight:700,color,letterSpacing:"-0.5px"}}>{value}</div>
    </div>
  );

  const Section = ({title,children})=>(
    <div style={{background:COL.surface,border:`1px solid ${COL.border}`,borderRadius:12,overflow:"hidden",marginTop:16}}>
      <div style={{padding:"13px 18px",borderBottom:`1px solid ${COL.border}`,fontSize:13,fontWeight:600,color:COL.textSub}}>{title}</div>
      {children}
    </div>
  );

  return (
    <div style={S.root}>
      <style>{`*{box-sizing:border-box}input::placeholder{color:${COL.textFaint}}tbody tr:hover{background:#f2f5f9!important}.hov-edit:hover{border-color:${COL.border2}!important;color:${COL.text}!important}.import-drop:hover{border-color:#9aa5b4!important;background:#fafbfc!important}input[type="time"]::-webkit-calendar-picker-indicator{opacity:.35;cursor:pointer}.chip-auto{display:inline-block;padding:1px 7px;border-radius:20px;background:${COL.accentBg};color:${COL.textFaint};font-size:10px;margin-left:7px;font-family:${MONO}}`}</style>

      {/* HEADER */}
      <header style={S.header}>
        <div style={S.logo}><div style={S.dot}/> Control de Asistencia</div>
        <div style={S.hRight}>
          {sbLoading && (
            <span style={{fontSize:11,color:COL.textFaint,display:"flex",alignItems:"center",gap:5}}>
              <span style={{display:"inline-block",width:8,height:8,borderRadius:"50%",background:"#f59e0b",animation:"pulse 1s infinite"}}/>
              Sincronizando…
            </span>
          )}
          {!sbLoading && sbStatus==="synced" && sbLastSync && (
            <span style={{fontSize:11,color:"#276749",display:"flex",alignItems:"center",gap:5}}>
              <span style={{display:"inline-block",width:8,height:8,borderRadius:"50%",background:"#4ade80"}}/>
              Sincronizado {sbLastSync.toLocaleTimeString("es-AR",{hour:"2-digit",minute:"2-digit"})}
            </span>
          )}
          {!sbLoading && sbStatus==="error" && (
            <span style={{fontSize:11,color:"#c53030",display:"flex",alignItems:"center",gap:5}}>
              <span style={{display:"inline-block",width:8,height:8,borderRadius:"50%",background:"#f87171"}}/>
              Sin conexión a Supabase
            </span>
          )}
          {!SB_URL && (
            <span style={{fontSize:11,color:"#b45309",display:"flex",alignItems:"center",gap:5}}>
              <span style={{display:"inline-block",width:8,height:8,borderRadius:"50%",background:"#f59e0b"}}/>
              Solo local
            </span>
          )}
          <span style={S.chip}>{records.length} registros</span>
          <span style={S.chip}>{empSummary.length} activos</span>
          <span style={{...S.chip,background:"#edf5ff",color:"#1e5fa8"}}>{empList.filter(e=>(e.tipo||"operario")==="operario").length} op.</span>
          <span style={{...S.chip,background:"#f0faf4",color:"#276749"}}>{empList.filter(e=>e.tipo==="administrativo").length} adm.</span>
        </div>
      </header>

      {/* TABS */}
      <div style={S.tabBar}>
        {TABS.map((t,i)=>(
          <button key={i} onClick={()=>setTab(i)} style={{...S.tab,...(tab===i?S.tabOn:{})}}>
            {t}{tab===i&&<span style={S.tabLine}/>}
          </button>
        ))}
      </div>

      <main style={S.main}>

        {/* ── 0 IMPORTAR ── */}
        {tab===0&&(
          <div style={{maxWidth:620}}>
            <H2>Importar desde el reloj</H2>
            <p style={S.body}>Cargá el archivo <Code>.xls</Code> o <Code>.xlsx</Code> tal cual sale del pendrive. La app lee la hoja <Code>Anormal</Code> y toma la primera marca del día como entrada y la última como salida.</p>

            <label htmlFor="fi" className="import-drop" style={S.drop}>
              <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke={COL.textFaint} strokeWidth="1.5" strokeLinecap="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
              <span style={{color:COL.textSub,fontSize:14,fontWeight:500}}>{importing?"Procesando…":"Seleccioná el archivo Excel"}</span>
              <span style={{color:COL.textFaint,fontSize:12}}>Podés cargar múltiples archivos de distintos meses</span>
              <input ref={fileRef} type="file" accept=".xls,.xlsx" onChange={handleFile} id="fi" style={{display:"none"}}/>
            </label>

            {importMsg.text&&<div style={importMsg.ok?S.alertOk:S.alertErr}><span style={{marginRight:8}}>{importMsg.ok?"✓":"✕"}</span>{importMsg.text}</div>}

            {records.length>0&&(
              <div style={{display:"flex",gap:10,marginTop:16}}>
                <button onClick={()=>setTab(1)} style={S.btnP}>Ver registros →</button>
                <button onClick={()=>{if(confirm("¿Borrar todos los registros?"))setRecords([])}} style={S.btnD}>Borrar registros</button>
              </div>
            )}

            <div style={S.infoBox}>
              <p style={{margin:"0 0 10px",fontWeight:600,color:COL.text,fontSize:13}}>Cómo se interpreta el reloj</p>
              {[["Marcas brutas","Las 4 columnas AM/PM se tratan como marcas de tiempo independientes."],["Orden cronológico","Primera marca = entrada · Última marca = salida. Las marcas intermedias se desestiman."],["Jornada","Muestra las horas configuradas para cada empleado (salida ref. − entrada ref.)."],["Hs. extra","Horas trabajadas fuera del horario: adelanto de entrada + extensión de salida."],["Horarios","Configurá la entrada/salida esperada por empleado en la pestaña Empleados."]].map(([k,v])=>(
                <div key={k} style={{display:"flex",gap:12,marginBottom:7}}>
                  <span style={{color:COL.accent,fontSize:12,fontWeight:600,minWidth:130}}>{k}</span>
                  <span style={{color:COL.textSub,fontSize:12}}>{v}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── 1 REGISTROS ── */}
        {tab===1&&(
          <div>
            <H2>Registros de asistencia</H2>
            <div style={{...S.filterBar,justifyContent:"space-between"}}>
              <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
                <input placeholder="Buscar empleado…" value={recF.emp} onChange={e=>setRecF(p=>({...p,emp:e.target.value}))} style={S.sInput}/>
                <input type="date" value={recF.fecha} onChange={e=>setRecF(p=>({...p,fecha:e.target.value}))} style={S.dateInput}/>
                <span style={{color:COL.textFaint,fontSize:12}}>{filteredRecs.length} registros</span>
              </div>
              <button onClick={()=>{setAddingRec(true);setNewRec({empNo:"",fecha:recF.fecha||"",entrada:"",salida:""});}}
                style={{...S.btnS,display:"flex",alignItems:"center",gap:6}}>
                <span style={{fontSize:15,lineHeight:1}}>+</span> Agregar registro
              </button>
            </div>

            {/* ── Agregar registro manual form ── */}
            {addingRec&&(
              <div style={{background:"#fef9f0",border:"1px solid #f6d860",borderRadius:10,padding:"16px 18px",marginBottom:14,display:"flex",gap:12,alignItems:"flex-end",flexWrap:"wrap"}}>
                <div>
                  <div style={{fontSize:11,color:COL.textFaint,marginBottom:4,fontWeight:500}}>Empleado</div>
                  <select value={newRec.empNo} onChange={e=>setNewRec(p=>({...p,empNo:e.target.value}))}
                    style={{...S.sInput,width:220,padding:"7px 10px"}}>
                    <option value="">— Seleccioná —</option>
                    {empList.map(e=><option key={e.empNo} value={e.empNo}>{e.empNo} · {cap(e.nombre)}</option>)}
                  </select>
                </div>
                <div>
                  <div style={{fontSize:11,color:COL.textFaint,marginBottom:4,fontWeight:500}}>Fecha</div>
                  <input type="date" value={newRec.fecha} onChange={e=>setNewRec(p=>({...p,fecha:e.target.value}))} style={{...S.dateInput}}/>
                </div>
                <div>
                  <div style={{fontSize:11,color:COL.textFaint,marginBottom:4,fontWeight:500}}>Entrada</div>
                  <input type="time" value={newRec.entrada} onChange={e=>setNewRec(p=>({...p,entrada:e.target.value}))} style={S.tInp}/>
                </div>
                <div>
                  <div style={{fontSize:11,color:COL.textFaint,marginBottom:4,fontWeight:500}}>Salida</div>
                  <input type="time" value={newRec.salida} onChange={e=>setNewRec(p=>({...p,salida:e.target.value}))} style={S.tInp}/>
                </div>
                <div style={{display:"flex",gap:8}}>
                  <button onClick={()=>{
                    if(!newRec.empNo||!newRec.fecha||!newRec.entrada) return;
                    const emp=employees[Number(newRec.empNo)];
                    if(!emp) return;
                    const id=`${newRec.empNo}_${newRec.fecha}_manual`;
                    const manRec={
                      id, empNo:Number(newRec.empNo), nombre:emp.nombre, depto:emp.depto,
                      fecha:newRec.fecha, entrada:newRec.entrada, salida:newRec.salida||null,
                      soloEntrada:!newRec.salida, manual:true
                    };
                    setManualRecords(p=>[...p.filter(r=>r.id!==id),manRec]);
                    setAddingRec(false);
                    sbUpsertSingle("registros", recToRow(manRec));
                  }} style={S.saveBtn}>Guardar</button>
                  <button onClick={()=>setAddingRec(false)} style={S.cancelBtn}>Cancelar</button>
                </div>
              </div>
            )}

            <div style={S.tblWrap}>
              <table style={S.table}>
                <THead cols={["N°","Nombre","Tipo","Fecha","Entrada","Salida","En jornada","Hs. extra","Demora","Sal. temprana",""]}/>
                <tbody>
                  {filteredRecs.slice(0,300).map((r,i)=>{
                    // Merge manual overrides
                    const manEnt = r.manual ? r.entrada  : (editingCell?.id===r.id&&editingCell?.field==="entrada" ? null : null);
                    const actualEntrada = r.manual ? r.entrada  : (manualSalidas[r.id+"_ent"] || r.entrada);
                    const actualSalida  = r.manual ? r.salida   : (manualSalidas[r.id]        || r.salida);
                    const rr={...r, entrada:actualEntrada, salida:actualSalida, soloEntrada:!actualSalida&&!actualEntrada};
                    const c=calcRecord(rr,employees[r.empNo],specialDays);
                    const isEditing=(field)=>editingCell?.id===r.id&&editingCell?.field===field;

                    const TimeCell=({field,value,color})=>(
                      <td style={{...S.td,fontFamily:MONO,padding:"5px 8px"}}>
                        {isEditing(field)
                          ? <input type="time" autoFocus
                              defaultValue={value||""}
                              onBlur={e=>{
                                const v=e.target.value;
                                const key=field==="salida"?r.id:r.id+"_ent";
                                if(v){
                                  setManualSalidas(p=>({...p,[key]:v}));
                                  // sync correction to Supabase
                                  const corrRow={id:r.id};
                                  if(field==="salida") corrRow.salida_corr=v;
                                  else corrRow.entrada_corr=v;
                                  sbUpsertSingle("correcciones", corrRow);
                                } else {
                                  setManualSalidas(p=>{const u={...p};delete u[key];return u;});
                                }
                                setEditingCell(null);
                              }}
                              onKeyDown={e=>{if(e.key==="Escape")setEditingCell(null);}}
                              style={{...S.tInp,width:88,fontSize:12,padding:"3px 6px"}}
                            />
                          : <span
                              onClick={()=>!r.manual&&setEditingCell({id:r.id,field})}
                              title={r.manual?"":"Click para editar"}
                              style={{color:value?color:"#c0c8d2",cursor:r.manual?"default":"pointer",
                                borderBottom:!r.manual&&value?"1px dashed #c0c8d2":"none",
                                paddingBottom:!r.manual&&value?1:0}}>
                              {value||<span style={{color:"#d1a054",fontSize:11,fontStyle:"italic"}}>sin marca</span>}
                            </span>
                        }
                      </td>
                    );

                    return(
                      <tr key={r.id} style={{background:r.manual?"#fefaf2":i%2===0?"#fff":"#fafbfc"}}>
                        <td style={{...S.td,fontFamily:MONO,fontSize:12,color:COL.accent}}>{r.empNo}</td>
                        <td style={{...S.td,textAlign:"left",fontWeight:500,color:COL.text}}>
                          {cap(r.nombre)}
                          {r.manual&&<span style={{marginLeft:7,fontSize:10,background:"#fde68a",color:"#92400e",borderRadius:4,padding:"1px 6px",fontWeight:600}}>manual</span>}
                        </td>
                        <td style={{...S.td,padding:"5px 8px"}}>
                          {employees[r.empNo]&&<TipoBadge tipo={employees[r.empNo].tipo||"operario"} small/>}
                        </td>
                        <td style={{...S.td,fontFamily:MONO,color:COL.textFaint,fontSize:12}}>{r.fecha}</td>
                        <TimeCell field="entrada" value={rr.entrada} color="#276749"/>
                        <TimeCell field="salida"  value={rr.salida}  color="#1e5fa8"/>
                        <td style={{...S.td,fontFamily:MONO,color:rr.soloEntrada?"#c0c8d2":c.trabajado!=null?(c.trabajado>=c.jornada?"#276749":"#b45309"):COL.textFaint}}>
                          {rr.soloEntrada ? "—" : c.trabajado!=null ? minsToDisplay(c.trabajado) : "—"}
                        </td>
                        <td style={{...S.td,fontFamily:MONO,color:c.extra!=null?"#276749":"#c0c8d2",fontWeight:c.extra!=null?600:400}}>
                          {c.extra!=null ? `+${minsToDisplay(c.extra)}` : "—"}
                        </td>
                        <td style={{...S.td,fontFamily:MONO,color:c.demora>0?"#c53030":"#c0c8d2"}}>{c.demora>0?minsToDisplay(c.demora):"—"}</td>
                        <td style={{...S.td,fontFamily:MONO,color:c.salTemprana>0?"#b45309":"#c0c8d2"}}>{c.salTemprana>0?minsToDisplay(c.salTemprana):"—"}</td>
                        <td style={{...S.td,padding:"5px 8px",width:28}}>
                          {r.manual&&(
                            <button onClick={()=>{setManualRecords(p=>p.filter(x=>x.id!==r.id));sbDelete("registros",r.id);}}
                              title="Eliminar registro manual"
                              style={{background:"none",border:"none",cursor:"pointer",color:"#fca5a5",fontSize:15,lineHeight:1,padding:2}}>×</button>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {filteredRecs.length===0&&<div style={{padding:"36px",textAlign:"center",color:COL.textFaint,fontSize:13}}>Sin registros. Importá un archivo primero o agregá uno manual.</div>}
            </div>
            {filteredRecs.length>300&&<p style={{color:COL.textFaint,fontSize:12,marginTop:8}}>Mostrando 300 de {filteredRecs.length} — usá los filtros para acotar.</p>}
          </div>
        )}

        {/* ── 2 EMPLEADOS ── */}
        {tab===2&&(
          <div>
            <H2>Empleados — Horarios de referencia</H2>
            <p style={S.body}>Estos horarios se usan para calcular demoras y salidas tempranas. Editá individualmente o seleccioná varios para asignar un horario en bloque.</p>

            {/* Tipo filter tabs */}
            <div style={{display:"flex",gap:6,marginBottom:14}}>
              {[["todos","Todos"],["operario","Operarios"],["administrativo","Administrativos"]].map(([val,label])=>(
                <button key={val} onClick={()=>setEmpTipoF(val)}
                  style={{padding:"5px 14px",borderRadius:20,border:`1px solid ${empTipoF===val?COL.accent:COL.border2}`,background:empTipoF===val?COL.accentBg:"#fff",color:empTipoF===val?COL.accent:COL.textSub,fontFamily:SANS,fontSize:12,cursor:"pointer",fontWeight:empTipoF===val?600:400}}>
                  {label} <span style={{fontFamily:MONO,fontSize:11,color:empTipoF===val?COL.accent:COL.textFaint,marginLeft:4}}>
                    {val==="todos"?empList.length:empList.filter(e=>(e.tipo||"operario")===val).length}
                  </span>
                </button>
              ))}
            </div>

            <div style={S.bulkBar}>
              <div style={{display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontSize:13,color:COL.textSub,minWidth:140}}>
                  {bulkSel.size>0?<><strong style={{color:COL.text}}>{bulkSel.size}</strong> seleccionados</>:"Seleccioná con ☑"}
                </span>
                <label style={{display:"flex",alignItems:"center",gap:7,fontSize:12,color:COL.textSub}}>
                  Entrada <input type="time" value={bulkH.entrada} onChange={e=>setBulkH(p=>({...p,entrada:e.target.value}))} style={S.tInp}/>
                </label>
                <label style={{display:"flex",alignItems:"center",gap:7,fontSize:12,color:COL.textSub}}>
                  Salida <input type="time" value={bulkH.salida} onChange={e=>setBulkH(p=>({...p,salida:e.target.value}))} style={S.tInp}/>
                </label>
                <button onClick={applyBulk} disabled={!bulkSel.size} style={{...S.btnS,opacity:bulkSel.size?1:0.4}}>Aplicar al grupo</button>
                {bulkSel.size>0&&<button onClick={()=>setBulkSel(new Set())} style={{...S.btnS,background:"transparent",color:COL.textSub,border:`1px solid ${COL.border2}`}}>Limpiar</button>}
              </div>
              <input placeholder="Buscar empleado…" value={empF} onChange={e=>setEmpF(e.target.value)} style={S.sInput}/>
            </div>

            <div style={S.tblWrap}>
              <table style={S.table}>
                <thead><tr>
                  <th style={{...S.th,width:36}}></th>
                  {["N°","Nombre","Tipo","Depto","Entrada ref.","Salida ref.","Activo","Días reg.",""].map((h,i)=><th key={i} style={{...S.th,textAlign:i===1?"left":undefined}}>{h}</th>)}
                </tr></thead>
                <tbody>
                  {filteredEmps.map((emp,i)=>{
                    const isEd=editingEmp===emp.empNo, isSel=bulkSel.has(emp.empNo);
                    const dias=allRecs.filter(r=>r.empNo===emp.empNo).length;
                    return(
                      <tr key={emp.empNo} style={{background:isSel?"#edf2f9":i%2===0?"#fff":"#fafbfc",transition:"background .1s"}}>
                        <td style={{...S.td,padding:"6px 10px"}}>
                          <input type="checkbox" checked={isSel} onChange={()=>setBulkSel(p=>{const s=new Set(p);s.has(emp.empNo)?s.delete(emp.empNo):s.add(emp.empNo);return s;})} style={{cursor:"pointer",accentColor:COL.accent,width:15,height:15}}/>
                        </td>
                        <td style={{...S.td,fontFamily:MONO,fontSize:12,color:COL.accent}}>{emp.empNo}</td>
                        <td style={{...S.td,textAlign:"left",fontWeight:500,color:COL.text}}>
                          {isEd?<input value={editDraft.nombre} onChange={e=>setEditDraft(p=>({...p,nombre:e.target.value}))} style={S.inlineInput}/>
                            :<>{cap(emp.nombre)}{emp.autoDetected&&<span className="chip-auto">auto</span>}</>}
                        </td>
                        <td style={{...S.td,padding:"6px 10px"}}>
                          {isEd
                            ? <select value={editDraft.tipo||"operario"} onChange={e=>setEditDraft(p=>({...p,tipo:e.target.value}))}
                                style={{...S.sInput,width:150,padding:"5px 8px",fontSize:12}}>
                                <option value="operario">Operario</option>
                                <option value="administrativo">Administrativo</option>
                              </select>
                            : <TipoBadge tipo={emp.tipo||"operario"} onClick={()=>toggleTipo(emp.empNo)} small/>
                          }
                        </td>
                        <td style={S.td}>
                          {isEd?<input value={editDraft.depto} onChange={e=>setEditDraft(p=>({...p,depto:e.target.value}))} style={{...S.inlineInput,width:70,textAlign:"center"}}/>
                            :<span style={{color:COL.textFaint,fontSize:12}}>{emp.depto}</span>}
                        </td>
                        <td style={S.td}>
                          {isEd?<input type="time" value={editDraft.entrada} onChange={e=>setEditDraft(p=>({...p,entrada:e.target.value}))} style={S.tInp}/>
                            :<TimeTag color="g">{emp.entrada}</TimeTag>}
                        </td>
                        <td style={S.td}>
                          {isEd?<input type="time" value={editDraft.salida} onChange={e=>setEditDraft(p=>({...p,salida:e.target.value}))} style={S.tInp}/>
                            :<TimeTag color="b">{emp.salida}</TimeTag>}
                        </td>
                        <td style={{...S.td}}>
                          <button onClick={()=>{
                            const updated={...employees[emp.empNo],activo:!emp.activo};
                            setEmployees(p=>({...p,[emp.empNo]:updated}));
                            sbUpsertSingle("empleados",empToRow(updated),"emp_no");
                          }} style={{
                            background:emp.activo?"#f0faf4":"#f5f5f5",
                            color:emp.activo?"#276749":"#9aa5b4",
                            border:`1px solid ${emp.activo?"#c3e6cb":"#e0e0e0"}`,
                            borderRadius:20,padding:"3px 12px",cursor:"pointer",
                            fontFamily:SANS,fontSize:11,fontWeight:600
                          }}>{emp.activo?"Activo":"Inactivo"}</button>
                        </td>
                        <td style={{...S.td,color:dias>0?COL.textSub:"#c0c8d2",fontSize:13}}>{dias||"—"}</td>
                        <td style={{...S.td,padding:"6px 14px"}}>
                          {isEd
                            ?<span style={{display:"flex",gap:6}}>
                                <button onClick={()=>saveEdit(emp.empNo)} style={S.saveBtn}>Guardar</button>
                                <button onClick={cancelEdit} style={S.cancelBtn}>Cancelar</button>
                              </span>
                            :<button onClick={()=>startEdit(emp)} className="hov-edit" style={S.editBtn}>Editar</button>}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ── 3 RESUMEN ── */}
        {tab===3&&(
          <div style={{maxWidth:960}}>
            <H2>Resumen del período</H2>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:26}}>
              <KpiCard label="Empleados activos"  value={empSummary.length}                             color={COL.accent}/>
              <KpiCard label="Días registrados"   value={records.length}                               color={COL.textSub}/>
              <KpiCard label="Con demoras"        value={empSummary.filter(s=>s.demoraDias>0).length} color="#c53030"/>
              <KpiCard label="Con sal. temprana"  value={empSummary.filter(s=>s.stDias>0).length}     color="#b45309"/>
            </div>

            <Section title="Top 15 — mayor demora acumulada">
              <table style={S.table}>
                <THead cols={["N°","Nombre","Horario","Días con demora","Total demora","Peor día","Total trabajado"]}/>
                <tbody>
                  {[...empSummary].filter(s=>s.totalDemora>0).sort((a,b)=>b.totalDemora-a.totalDemora).slice(0,15).map((s,i)=>(
                    <tr key={s.emp.empNo} style={{background:i%2?"#fafbfc":"#fff"}}>
                      <td style={{...S.td,fontFamily:MONO,fontSize:12,color:COL.accent}}>{s.emp.empNo}</td>
                      <td style={{...S.td,textAlign:"left",fontWeight:500,color:COL.text}}>{cap(s.emp.nombre)}</td>
                      <td style={{...S.td,fontFamily:MONO,color:COL.textFaint,fontSize:12}}>{s.emp.entrada}–{s.emp.salida}</td>
                      <td style={S.td}>{s.demoraDias}</td>
                      <td style={{...S.td,fontFamily:MONO,color:"#c53030",fontWeight:600}}>{minsToDisplay(s.totalDemora)}</td>
                      <td style={{...S.td,fontFamily:MONO,color:"#b45309"}}>{minsToDisplay(s.peorDemora)}</td>
                      <td style={{...S.td,fontFamily:MONO,color:COL.textFaint}}>{minsToDisplay(s.totalMin)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </Section>

            <Section title="Todos los empleados">
              <table style={S.table}>
                <THead cols={["N°","Nombre","Horario","Días","Total hs","Dem. días","Min demora","Sal. temp"]}/>
                <tbody>
                  {empSummary.map((s,i)=>(
                    <tr key={s.emp.empNo} style={{background:i%2?"#fafbfc":"#fff"}}>
                      <td style={{...S.td,fontFamily:MONO,fontSize:12,color:COL.accent}}>{s.emp.empNo}</td>
                      <td style={{...S.td,textAlign:"left",fontWeight:500,color:COL.text}}>{cap(s.emp.nombre)}</td>
                      <td style={{...S.td,fontFamily:MONO,color:COL.textFaint,fontSize:12}}>{s.emp.entrada}–{s.emp.salida}</td>
                      <td style={S.td}>{s.dias}</td>
                      <td style={{...S.td,fontFamily:MONO,color:COL.textSub}}>{minsToDisplay(s.totalMin)}</td>
                      <td style={{...S.td,fontFamily:MONO,color:s.demoraDias>0?"#c53030":"#c0c8d2"}}>{s.demoraDias||"—"}</td>
                      <td style={{...S.td,fontFamily:MONO,color:s.totalDemora>0?"#c53030":"#c0c8d2"}}>{s.totalDemora>0?minsToDisplay(s.totalDemora):"—"}</td>
                      <td style={{...S.td,fontFamily:MONO,color:s.stDias>0?"#b45309":"#c0c8d2"}}>{s.stDias||"—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </Section>
          </div>
        )}


        {/* ── 4 POR EMPLEADO ── */}
        {tab===4&&(()=>{
          const empOptions = empList.filter(e=>e.activo!==false&&records.some(r=>r.empNo===e.empNo));
          const selEmp     = detalleEmp ? employees[detalleEmp] : null;
          const selSummary = selEmp ? empSummary.find(s=>s.emp.empNo===selEmp.empNo) : null;
          const selCalcs   = selSummary ? [...selSummary.calcs].sort((a,b)=>a.fecha.localeCompare(b.fecha)) : [];

          const totTrabajado = selCalcs.reduce((s,r)=>s+(r.trabajado||0),0);
          const totJornada   = selCalcs.reduce((s,r)=>s+(r.jornada||0),0);
          const totExtra     = selCalcs.reduce((s,r)=>s+(r.extra||0),0);
          const totDemora    = selCalcs.reduce((s,r)=>s+(r.demora||0),0);
          const totSalTemp   = selCalcs.reduce((s,r)=>s+(r.salTemprana||0),0);
          const diasConDemora = selCalcs.filter(r=>r.demora>0).length;
          const diasConExtra  = selCalcs.filter(r=>r.extra>0).length;
          const diasConST     = selCalcs.filter(r=>r.salTemprana>0).length;

          return(
            <div style={{maxWidth:860}}>
              <H2>Detalle por empleado</H2>
              <p style={S.body}>Seleccioná un empleado para ver su historial completo del período con todos los totales acumulados.</p>

              <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:24,flexWrap:"wrap"}}>
                <label style={{fontSize:13,color:COL.textSub,fontWeight:500}}>Empleado</label>
                <select value={detalleEmp||""} onChange={e=>setDetalleEmp(e.target.value?Number(e.target.value):null)}
                  style={{background:COL.surface,border:`1px solid ${COL.border2}`,borderRadius:8,color:COL.text,padding:"8px 14px",fontFamily:SANS,fontSize:13,outline:"none",minWidth:280,cursor:"pointer"}}>
                  <option value="">— Elegí un empleado —</option>
                  {empOptions.map(e=>(
                    <option key={e.empNo} value={e.empNo}>{e.empNo} · {cap(e.nombre)}</option>
                  ))}
                </select>
                {selEmp&&<span style={{...S.chip,background:"#f0faf4",color:"#276749"}}>Horario: {selEmp.entrada} – {selEmp.salida}</span>}
              </div>

              {!selEmp&&(
                <div style={{padding:"60px 20px",textAlign:"center",color:COL.textFaint,fontSize:13,background:COL.surface,borderRadius:12,border:`1px solid ${COL.border}`}}>
                  Elegí un empleado en el selector de arriba para ver su detalle.
                </div>
              )}

              {selEmp&&selCalcs.length>0&&(<>
                <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,marginBottom:18}}>
                  {[
                    {l:"Días presentes",  v:selCalcs.length,                                    c:COL.accent},
                    {l:"En jornada total", v:minsToDisplay(totTrabajado),                        c:COL.textSub},
                    {l:"Hs. extra",       v:totExtra>0?`+${minsToDisplay(totExtra)}`:"—",       c:"#276749"},
                    {l:"Total demoras",   v:totDemora>0?minsToDisplay(totDemora):"—",           c:"#c53030"},
                    {l:"Sal. tempranas",  v:totSalTemp>0?minsToDisplay(totSalTemp):"—",         c:"#b45309"},
                  ].map(({l,v,c})=>(
                    <div key={l} style={{background:COL.surface,border:`1px solid ${COL.border}`,borderRadius:10,padding:"13px 14px",borderTop:`3px solid ${c}`}}>
                      <div style={{fontSize:10,color:COL.textFaint,fontWeight:500,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:6}}>{l}</div>
                      <div style={{fontSize:20,fontWeight:700,color:c,letterSpacing:"-0.3px"}}>{v}</div>
                    </div>
                  ))}
                </div>

                <div style={{display:"flex",gap:10,marginBottom:18,flexWrap:"wrap"}}>
                  {[
                    {l:"Días con demora",        v:`${diasConDemora} día${diasConDemora!==1?"s":""}`,  c:"#c53030"},
                    {l:"Días con hs. extra",     v:`${diasConExtra} día${diasConExtra!==1?"s":""}`,    c:"#276749"},
                    {l:"Días con sal. temprana", v:`${diasConST} día${diasConST!==1?"s":""}`,          c:"#b45309"},
                    {l:"Jornada esperada total", v:minsToDisplay(totJornada),                          c:COL.textSub},
                  ].map(({l,v,c})=>(
                    <div key={l} style={{background:COL.accentBg,borderRadius:8,padding:"8px 14px",display:"flex",gap:8,alignItems:"center"}}>
                      <span style={{fontSize:11,color:COL.textFaint}}>{l}</span>
                      <span style={{fontFamily:MONO,fontSize:13,fontWeight:600,color:c}}>{v}</span>
                    </div>
                  ))}
                </div>

                <div style={S.tblWrap}>
                  <table style={S.table}>
                    <thead><tr>
                      {["Fecha","Día","Entrada","Salida","En jornada","Real total","Hs. extra","Demora","Sal. temprana"].map((h,i)=>(
                        <th key={h} style={{...S.th,textAlign:i<=1?"left":undefined}}>{h}</th>
                      ))}
                    </tr></thead>
                    <tbody>
                      {selCalcs.map((r,i)=>{
                        const fecha = new Date(r.fecha+"T12:00:00");
                        const dia   = ["Dom","Lun","Mar","Mié","Jue","Vie","Sáb"][fecha.getDay()];
                        const esFinde = fecha.getDay()===0||fecha.getDay()===6;
                        return(
                          <tr key={r.id} style={{background:esFinde?"#fdf9f3":i%2===0?"#fff":"#fafbfc"}}>
                            <td style={{...S.td,textAlign:"left",fontFamily:MONO,fontSize:12,color:esFinde?"#b45309":COL.textFaint}}>{r.fecha}</td>
                            <td style={{...S.td,textAlign:"left",fontSize:12,fontWeight:500,color:esFinde?"#b45309":COL.textSub,width:40}}>{dia}</td>
                            <td style={{...S.td,fontFamily:MONO,color:r.entrada?"#276749":"#c0c8d2"}}>{r.entrada||"—"}</td>
                            <td style={{...S.td,fontFamily:MONO,color:r.salida?"#1e5fa8":"#c0c8d2"}}>{r.salida||"—"}</td>
                            <td style={{...S.td,fontFamily:MONO,color:r.soloEntrada?"#b45309":r.trabajado!=null?(r.trabajado>=r.jornada?"#276749":"#b45309"):COL.textFaint}}>
                              {r.soloEntrada?"solo entrada":r.trabajado!=null?minsToDisplay(r.trabajado):"—"}
                            </td>
                            <td style={{...S.td,fontFamily:MONO,color:COL.textSub}}>
                              {(()=>{
                                const e=parseTimeVal(r.entrada),s=parseTimeVal(r.salida);
                                return (e!=null&&s!=null)?minsToDisplay(s-e):"—";
                              })()}
                            </td>
                            <td style={{...S.td,fontFamily:MONO,color:r.extra!=null?"#276749":"#c0c8d2",fontWeight:r.extra!=null?600:400}}>
                              {r.extra!=null?`+${minsToDisplay(r.extra)}`:"—"}
                            </td>
                            <td style={{...S.td,fontFamily:MONO,color:r.demora>0?"#c53030":"#c0c8d2"}}>
                              {r.demora>0?minsToDisplay(r.demora):"—"}
                            </td>
                            <td style={{...S.td,fontFamily:MONO,color:r.salTemprana>0?"#b45309":"#c0c8d2"}}>
                              {r.salTemprana>0?minsToDisplay(r.salTemprana):"—"}
                            </td>
                          </tr>
                        );
                      })}
                      <tr style={{background:"#f0f4fa",borderTop:`2px solid ${COL.border2}`}}>
                        <td colSpan={2} style={{...S.td,textAlign:"left",fontWeight:700,color:COL.text,fontSize:11,letterSpacing:"0.05em",textTransform:"uppercase"}}>TOTAL</td>
                        <td style={S.td}></td><td style={S.td}></td>
                        <td style={{...S.td,fontFamily:MONO,fontWeight:600,color:COL.textSub}}>{minsToDisplay(totJornada)}</td>
                        <td style={{...S.td,fontFamily:MONO,fontWeight:600,color:COL.textSub}}>{minsToDisplay(totTrabajado)}</td>
                        <td style={{...S.td,fontFamily:MONO,fontWeight:700,color:"#276749"}}>{totExtra>0?`+${minsToDisplay(totExtra)}`:"—"}</td>
                        <td style={{...S.td,fontFamily:MONO,fontWeight:700,color:"#c53030"}}>{totDemora>0?minsToDisplay(totDemora):"—"}</td>
                        <td style={{...S.td,fontFamily:MONO,fontWeight:700,color:"#b45309"}}>{totSalTemp>0?minsToDisplay(totSalTemp):"—"}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </>)}

              {selEmp&&selCalcs.length===0&&(
                <div style={{padding:"48px 20px",textAlign:"center",color:COL.textFaint,fontSize:13,background:COL.surface,borderRadius:12,border:`1px solid ${COL.border}`}}>
                  No hay registros para este empleado en el período cargado.
                </div>
              )}
            </div>
          );
        })()}


        {/* ── 5 CALENDARIO ── */}
        {tab===5&&(()=>{
          const allDates = [...new Set(records.map(r=>r.fecha))].sort();

          const toggleDay = (fecha, tipo) => {
            setSpecialDays(prev => {
              const cur = prev[fecha];
              if (cur?.tipo === tipo) {
                const u = {...prev}; delete u[fecha];
                sbDelete("dias_especiales", fecha, "fecha");
                return u;
              }
              sbUpsertSingle("dias_especiales", {fecha, tipo}, "fecha");
              return {...prev, [fecha]: {tipo}};
            });
          };

          const byMonth = {};
          for (const f of allDates) {
            const m = f.slice(0,7);
            if (!byMonth[m]) byMonth[m] = [];
            byMonth[m].push(f);
          }

          const DIAS  = ["Dom","Lun","Mar","Mié","Jue","Vie","Sáb"];
          const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
          const feriadosCount = Object.values(specialDays).filter(v=>v.tipo==="feriado").length;

          return (
            <div style={{maxWidth:900}}>
              <H2>Calendario — Días especiales</H2>
              <p style={S.body}>
                Marcá los días como <strong style={{color:"#b45309"}}>Feriado</strong> (operarios trabajan hasta las 14:00, el horario de referencia cambia automáticamente) o como{" "}
                <strong style={{color:"#c53030"}}>Día libre</strong> (sin registros esperados).
                Los cambios afectan los cálculos en todas las pestañas.
              </p>

              <div style={{display:"flex",gap:10,marginBottom:22,flexWrap:"wrap"}}>
                <div style={{background:"#fef3e2",border:"1px solid #f6d860",borderRadius:8,padding:"8px 16px",fontSize:13}}>
                  <span style={{color:"#b45309",fontWeight:600}}>{feriadosCount}</span>
                  <span style={{color:"#92400e",marginLeft:6}}>feriado{feriadosCount!==1?"s":""} marcado{feriadosCount!==1?"s":""}</span>
                </div>
                <div style={{background:COL.accentBg,border:`1px solid ${COL.accentSoft}`,borderRadius:8,padding:"8px 16px",fontSize:13}}>
                  <span style={{color:COL.accent,fontWeight:600}}>{allDates.length}</span>
                  <span style={{color:COL.textSub,marginLeft:6}}>días con registros</span>
                </div>
                {feriadosCount > 0 && (
                  <button onClick={()=>setSpecialDays({})} style={{background:"#fff5f5",border:"1px solid #fed7d7",borderRadius:8,padding:"8px 14px",fontSize:12,color:"#c53030",cursor:"pointer",fontFamily:SANS}}>
                    Limpiar todos
                  </button>
                )}
              </div>

              {allDates.length === 0 && (
                <div style={{padding:"60px 20px",textAlign:"center",color:COL.textFaint,fontSize:13,background:COL.surface,borderRadius:12,border:`1px solid ${COL.border}`}}>
                  Importá un archivo primero para ver el calendario.
                </div>
              )}

              {Object.entries(byMonth).map(([month, dates]) => {
                const [y,m] = month.split("-");
                const monthLabel = `${MESES[parseInt(m)-1]} ${y}`;
                return (
                  <div key={month} style={{marginBottom:28}}>
                    <div style={{fontSize:12,fontWeight:600,color:COL.textSub,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:12,paddingBottom:8,borderBottom:`2px solid ${COL.border}`}}>
                      {monthLabel}
                    </div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:10}}>
                      {dates.map(fecha => {
                        const d         = new Date(fecha+"T12:00:00");
                        const dia       = DIAS[d.getDay()];
                        const num       = parseInt(fecha.slice(8));
                        const sp        = specialDays[fecha];
                        const isFeriado = sp?.tipo === "feriado";
                        const isLibre   = sp?.tipo === "libre";
                        const esFinde   = d.getDay()===0||d.getDay()===6;
                        const recsCount = records.filter(r=>r.fecha===fecha).length;

                        let bg=COL.surface, borderC=COL.border, headC=COL.text, subC=COL.textFaint;
                        if (isFeriado) { bg="#fef9ee"; borderC="#f6d860"; headC="#92400e"; subC="#b45309"; }
                        else if (isLibre) { bg="#fff8f8"; borderC="#fca5a5"; headC="#c53030"; subC="#c53030"; }
                        else if (esFinde) { bg="#fafbfc"; headC=COL.textFaint; }

                        return (
                          <div key={fecha} style={{background:bg,border:`1px solid ${borderC}`,borderRadius:12,padding:"12px 14px",minWidth:130,maxWidth:150,transition:"all .15s",boxShadow:isFeriado?"0 1px 4px rgba(246,216,96,.3)":isLibre?"0 1px 4px rgba(252,165,165,.2)":"none"}}>
                            <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:4}}>
                              <span style={{fontFamily:MONO,fontSize:22,fontWeight:700,color:headC,lineHeight:1}}>{String(num).padStart(2,"0")}</span>
                              <span style={{fontSize:12,color:subC,fontWeight:500}}>{dia}</span>
                            </div>
                            <div style={{fontSize:11,color:COL.textFaint,marginBottom:10,fontFamily:MONO}}>
                              {recsCount} reg.
                            </div>
                            {isFeriado && (
                              <div style={{fontSize:10,background:"#fde68a",color:"#92400e",borderRadius:4,padding:"2px 7px",marginBottom:8,fontWeight:700,display:"inline-block",letterSpacing:"0.03em"}}>
                                ★ FERIADO
                              </div>
                            )}
                            {isFeriado && (
                              <div style={{fontSize:10,color:"#b45309",marginBottom:8}}>hasta 14:00</div>
                            )}
                            {isLibre && (
                              <div style={{fontSize:10,background:"#fecaca",color:"#c53030",borderRadius:4,padding:"2px 7px",marginBottom:8,fontWeight:700,display:"inline-block"}}>
                                DÍA LIBRE
                              </div>
                            )}
                            <div style={{display:"flex",flexDirection:"column",gap:5}}>
                              <button onClick={()=>toggleDay(fecha,"feriado")}
                                style={{fontSize:11,padding:"4px 8px",border:`1px solid ${isFeriado?"#f6d860":COL.border}`,borderRadius:6,cursor:"pointer",background:isFeriado?"#fde68a":"#fff",color:isFeriado?"#92400e":COL.textSub,fontFamily:SANS,fontWeight:isFeriado?700:400,textAlign:"left"}}>
                                {isFeriado ? "✓ Feriado" : "+ Feriado"}
                              </button>
                              <button onClick={()=>toggleDay(fecha,"libre")}
                                style={{fontSize:11,padding:"4px 8px",border:`1px solid ${isLibre?"#fca5a5":COL.border}`,borderRadius:6,cursor:"pointer",background:isLibre?"#fecaca":"#fff",color:isLibre?"#c53030":COL.textSub,fontFamily:SANS,fontWeight:isLibre?700:400,textAlign:"left"}}>
                                {isLibre ? "✓ Día libre" : "+ Día libre"}
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          );
        })()}


        {/* ── 6 CIRCULAR — Escala salarial ── */}
        {tab===6&&(()=>{
          const activeEmps = empList.filter(e=>e.activo!==false);
          const filteredCircular = activeEmps.filter(e=>
            !circularF ||
            e.nombre.toLowerCase().includes(circularF.toLowerCase()) ||
            String(e.empNo).includes(circularF)
          );

          const getP = (empNo) => liqParams[String(empNo)] || {};

          const setVal = (empNo, key, val) => {
            setLiqParams(prev => ({
              ...prev,
              [String(empNo)]: { ...(prev[String(empNo)]||{}), [key]: val }
            }));
          };

          const fmt$ = (v) => v ? `$ ${Number(v).toLocaleString("es-AR",{minimumFractionDigits:2,maximumFractionDigits:2})}` : "—";
          const completos = activeEmps.filter(e=>{ const p=getP(e.empNo); return p.sueldoBasico&&p.valorHora; }).length;

          const OBSERVACIONES_FIJAS = [
            "La carga horaria de los sábados será de 6 hs.",
            "Todas las horas extras se liquidarán al 50%",
            "El pago de los sueldos será mensual",
            "Los días y horas a trabajar se decidirán y aprobarán por parte de la Empresa.",
          ];

          return (
            <div style={{maxWidth:980}}>
              <H2>Circular — Escala salarial</H2>
              <p style={S.body}>Cada tarjeta refleja la circular informativa de cada empleado. Los valores se usan automáticamente en Liquidación.</p>

              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18,flexWrap:"wrap",gap:10}}>
                <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                  <div style={{background:COL.accentBg,border:`1px solid ${COL.accentSoft}`,borderRadius:8,padding:"7px 14px",fontSize:13}}>
                    <span style={{color:COL.accent,fontWeight:600}}>{completos}</span>
                    <span style={{color:COL.textSub,marginLeft:6}}>/ {activeEmps.length} con datos</span>
                  </div>
                  {completos < activeEmps.length && (
                    <div style={{background:"#fef9f0",border:"1px solid #f6d860",borderRadius:8,padding:"7px 14px",fontSize:13,color:"#92400e"}}>
                      ⚠ {activeEmps.length-completos} sin completar
                    </div>
                  )}
                </div>
                <input placeholder="Buscar empleado…" value={circularF}
                  onChange={e=>setCircularF(e.target.value)} style={S.sInput}/>
              </div>

              <div style={{display:"flex",flexDirection:"column",gap:20}}>
                {filteredCircular.map(emp=>{
                  const p       = getP(emp.empNo);
                  const isEd    = circularEdit === emp.empNo;
                  const draft   = isEd ? circularDraft : p;
                  const periodo = p.periodo || "";

                  const numVal  = (k) => draft[k] ? Number(draft[k]) : 0;
                  const dispVal = (k) => draft[k] ? fmt$(draft[k]) : "—";

                  // Derived display values
                  const horarioDisplay = `${emp.entrada||"06:00"} a ${emp.salida||"16:30"} hs`;

                  const InputOrVal = ({field, width=140}) => isEd
                    ? <input type="number" min="0" step="any"
                        value={circularDraft[field]||""}
                        onChange={e=>setCircularDraft(p=>({...p,[field]:e.target.value}))}
                        placeholder="0"
                        style={{...S.sInput,width,padding:"4px 8px",fontFamily:MONO,fontSize:12,textAlign:"right"}}
                      />
                    : <span style={{fontFamily:MONO,fontWeight:600,color:p[field]?COL.text:"#c0c8d2"}}>{dispVal(field)}</span>;

                  return (
                    <div key={emp.empNo} style={{background:COL.surface,border:`1px solid ${COL.border}`,borderRadius:12,overflow:"hidden"}}>

                      {/* Card header */}
                      <div style={{background:"#f7f8fa",borderBottom:`2px solid ${COL.accent}`,padding:"10px 20px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <div style={{display:"flex",alignItems:"center",gap:14}}>
                          <span style={{fontFamily:MONO,fontSize:12,color:COL.textFaint,background:COL.accentBg,padding:"2px 8px",borderRadius:4}}>N°{emp.empNo}</span>
                          <span style={{fontSize:15,fontWeight:700,color:COL.text,textTransform:"uppercase",letterSpacing:"0.5px"}}>{cap(emp.nombre)}</span>
                          {!p.sueldoBasico&&<span style={{fontSize:10,background:"#fef9f0",color:"#b45309",border:"1px solid #f6d860",borderRadius:4,padding:"1px 7px",fontWeight:600}}>sin datos</span>}
                        </div>
                        <div style={{display:"flex",gap:8,alignItems:"center"}}>
                          {isEd
                            ? <>
                                <button onClick={()=>{
                                  setLiqParams(prev=>({...prev,[String(emp.empNo)]:{...(prev[String(emp.empNo)]||{}),...circularDraft}}));
                                  setCircularEdit(null);
                                }} style={S.saveBtn}>Guardar</button>
                                <button onClick={()=>setCircularEdit(null)} style={S.cancelBtn}>Cancelar</button>
                              </>
                            : <button onClick={()=>{
                                setCircularEdit(emp.empNo);
                                setCircularDraft({
                                  sueldoBasico:  p.sueldoBasico  ||"",
                                  valorDia:      p.valorDia      ||"",
                                  valorHora:     p.valorHora     ||"",
                                  valorHoraExt:  p.valorHoraExt  ||"",
                                  valorDiaFinde: p.valorDiaFinde ||"",
                                  observaciones: p.observaciones ||"",
                                  periodoCircular: p.periodoCircular||"",
                                });
                              }} className="hov-edit" style={S.editBtn}>✏ Editar</button>
                          }
                        </div>
                      </div>

                      <div style={{padding:"16px 20px"}}>

                        {/* Top info row */}
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16,marginBottom:18}}>
                          <div>
                            <div style={{fontSize:10,color:COL.textFaint,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Sueldo bruto</div>
                            {isEd
                              ? <input type="number" min="0" step="any" value={circularDraft.sueldoBasico||""} onChange={e=>setCircularDraft(p=>({...p,sueldoBasico:e.target.value}))} placeholder="0"
                                  style={{...S.sInput,width:"100%",padding:"6px 10px",fontFamily:MONO,fontSize:14,fontWeight:600}}/>
                              : <span style={{fontFamily:MONO,fontSize:15,fontWeight:700,color:p.sueldoBasico?COL.text:"#c0c8d2"}}>{p.sueldoBasico?fmt$(p.sueldoBasico):"—"}</span>
                            }
                          </div>
                          <div>
                            <div style={{fontSize:10,color:COL.textFaint,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Horario</div>
                            <span style={{fontSize:13,color:COL.textSub}}>Lunes a Sábados {horarioDisplay}</span>
                          </div>
                          <div>
                            <div style={{fontSize:10,color:COL.textFaint,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Período circular</div>
                            {isEd
                              ? <input value={circularDraft.periodoCircular||""} onChange={e=>setCircularDraft(p=>({...p,periodoCircular:e.target.value}))} placeholder="ej: ago-25"
                                  style={{...S.sInput,width:"100%",padding:"6px 10px",fontSize:13}}/>
                              : <span style={{fontSize:13,color:p.periodoCircular?COL.text:"#c0c8d2"}}>{p.periodoCircular||"—"}</span>
                            }
                          </div>
                        </div>

                        {/* Adicionales y descuentos table */}
                        <div style={{marginBottom:18}}>
                          <div style={{fontSize:11,fontWeight:700,color:COL.text,textTransform:"uppercase",letterSpacing:"0.05em",padding:"7px 12px",background:"#f0f4fa",borderBottom:`1px solid ${COL.border}`,borderTop:`1px solid ${COL.border}`,marginBottom:0}}>
                            Adicionales y descuentos
                          </div>
                          <table style={{...S.table,border:`1px solid ${COL.border}`}}>
                            <tbody>
                              {[
                                {label:"Valor a descontar por día de falta", field:"valorDia",      note:""},
                                {label:"Valor hora de referencia",            field:"valorHora",     note:"por hora"},
                                {label:"Feriados trabajados",                 field:"valorDia",      note:`hasta las ${emp.salida||"16:30"} hs`, readOnly:true},
                                {label:"Sábados",                             field:"valorDiaFinde", note:"08 a 14 hs"},
                                {label:"Horas extras",                        field:"valorHoraExt",  note:`a partir de las ${emp.salida||"16:30"} hs`},
                              ].map(({label,field,note,readOnly},idx)=>(
                                <tr key={idx} style={{background:idx%2===0?"#fff":"#fafbfc",borderBottom:`1px solid ${COL.border}`}}>
                                  <td style={{...S.td,textAlign:"left",color:COL.textSub,fontSize:12,width:"45%"}}>{label}</td>
                                  <td style={{...S.td,fontFamily:MONO,fontSize:12,width:30,color:COL.textFaint}}>$</td>
                                  <td style={{...S.td,textAlign:"right",width:"25%"}}>
                                    {isEd && !readOnly
                                      ? <input type="number" min="0" step="any"
                                          value={circularDraft[field]||""}
                                          onChange={e=>setCircularDraft(p=>({...p,[field]:e.target.value}))}
                                          placeholder="0"
                                          style={{...S.sInput,width:130,padding:"3px 8px",fontFamily:MONO,fontSize:12,textAlign:"right"}}
                                        />
                                      : <span style={{fontFamily:MONO,fontWeight:500,color:p[field]?COL.text:"#c0c8d2",fontSize:12}}>
                                          {p[field]?Number(p[field]).toLocaleString("es-AR",{minimumFractionDigits:2}):"—"}
                                        </span>
                                    }
                                  </td>
                                  <td style={{...S.td,textAlign:"left",fontSize:11,color:COL.textFaint}}>{note}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>

                        {/* Observaciones */}
                        <div>
                          <div style={{fontSize:11,fontWeight:700,color:COL.text,textTransform:"uppercase",letterSpacing:"0.05em",padding:"7px 12px",background:"#f0f4fa",borderBottom:`1px solid ${COL.border}`,borderTop:`1px solid ${COL.border}`,marginBottom:8}}>
                            Observaciones
                          </div>
                          <div style={{padding:"0 4px"}}>
                            {OBSERVACIONES_FIJAS.map((obs,i)=>(
                              <div key={i} style={{fontSize:12,color:COL.textSub,padding:"4px 8px",display:"flex",gap:8}}>
                                <span style={{color:COL.textFaint}}>·</span>{obs}
                              </div>
                            ))}
                            {/* Custom observations per employee */}
                            <div style={{marginTop:8}}>
                              {isEd
                                ? <textarea
                                    value={circularDraft.observaciones||""}
                                    onChange={e=>setCircularDraft(p=>({...p,observaciones:e.target.value}))}
                                    placeholder="Observaciones adicionales para este empleado…"
                                    rows={3}
                                    style={{...S.sInput,width:"100%",resize:"vertical",padding:"8px 10px",fontSize:12,fontFamily:SANS,lineHeight:1.6}}
                                  />
                                : p.observaciones
                                  ? <div style={{fontSize:12,color:COL.textSub,padding:"6px 8px",background:"#f7f8fa",borderRadius:6,marginTop:4,lineHeight:1.7,whiteSpace:"pre-wrap"}}>
                                      {p.observaciones}
                                    </div>
                                  : null
                              }
                            </div>
                          </div>
                        </div>

                      </div>
                    </div>
                  );
                })}
                {filteredCircular.length===0&&(
                  <div style={{padding:"48px",textAlign:"center",color:COL.textFaint,fontSize:13,background:COL.surface,borderRadius:12,border:`1px solid ${COL.border}`}}>
                    {activeEmps.length===0?"No hay empleados activos.":"Sin resultados para la búsqueda."}
                  </div>
                )}
              </div>
            </div>
          );
        })()}


        {/* ── 7 LIQUIDACIÓN ── */}
        {tab===7&&(()=>{
          const empOptions = empList.filter(e=>e.activo!==false&&records.some(r=>r.empNo===e.empNo));
          const selEmp     = liqEmp ? employees[liqEmp] : null;
          const selSummary = selEmp ? empSummary.find(s=>s.emp.empNo===selEmp.empNo) : null;
          const selCalcs   = selSummary ? selSummary.calcs : [];

          // Get params for this employee (or empty object)
          const empKey   = liqEmp ? String(liqEmp) : "__none__";
          const p        = liqParams[empKey] || {};
          const setP     = (field, val) => setLiqParams(prev => ({
            ...prev,
            [empKey]: { ...(prev[empKey]||{}), [field]: val }
          }));

          // Date range filter
          const desde  = p.desde || "";
          const hasta  = p.hasta || "";
          const rangeCalcs = selCalcs.filter(r =>
            (!desde || r.fecha >= desde) && (!hasta || r.fecha <= hasta)
          );

          // Values from attendance data (filtered by range)
          const diasTrabajados   = rangeCalcs.filter(r=>r.trabajado!=null&&r.trabajado>0).length;
          const totalDemoraMin   = rangeCalcs.reduce((s,r)=>s+(r.demora||0),0);
          const totalSalTempMin  = rangeCalcs.reduce((s,r)=>s+(r.salTemprana||0),0);
          const totalExtraMin    = rangeCalcs.reduce((s,r)=>s+(r.extra||0),0);
          // Finde dates: all sat/sun in range (with or without record)
          const allFindeInRange = (()=>{
            if (!desde && !hasta && rangeCalcs.length === 0) return [];
            // collect all sat/sun dates from range records + generate from date range
            const fechasSet = new Set(rangeCalcs
              .filter(r=>{ const d=new Date(r.fecha+"T12:00:00").getDay(); return d===0||d===6; })
              .map(r=>r.fecha)
            );
            return [...fechasSet].sort();
          })();
          // Which ones RRHH checked (stored per employee in liqParams)
          const findeSel = new Set(p.findeSel || allFindeInRange); // default: all pre-selected
          const diasFinde = findeSel.size;
          const toggleFinde = (fecha) => {
            const cur = new Set(p.findeSel || allFindeInRange);
            cur.has(fecha) ? cur.delete(fecha) : cur.add(fecha);
            setP("findeSel", [...cur]);
          };
          const DIAS_CORTO = ["Dom","Lun","Mar","Mié","Jue","Vie","Sáb"];

          // Fracciones de 15 min de demoras (redondeo hacia arriba)
          const fraccionesDemora = Math.ceil(totalDemoraMin / 15);
          // Horas de retiro anticipado (en horas, 2 decimales)
          const horasSalTemp      = parseFloat((totalSalTempMin / 60).toFixed(2));
          // Horas extra — decimal para cálculos, HH:MM para mostrar
          const horasExtra        = parseFloat((totalExtraMin / 60).toFixed(10)); // full precision for math
          const horasExtraDisplay = minsToDisplay(totalExtraMin); // ej: 38h26

          // Manual inputs (stored per employee)
          const sueldoBasico  = parseFloat(p.sueldoBasico  || 0);
          const valorDia      = parseFloat(p.valorDia      || 0);
          const valorHora     = parseFloat(p.valorHora     || 0);
          const valorHoraExt  = parseFloat(p.valorHoraExt  || 0);
          const valorDiaFinde = parseFloat(p.valorDiaFinde || 0);
          const adelanto      = parseFloat(p.adelanto      || 0);
          const feriados      = parseFloat(p.feriados      || 0);
          const sac           = parseFloat(p.sac           || 0);
          const vacaciones    = parseFloat(p.vacaciones    || 0);
          const periodo       = p.periodo || "";
          const ingreso       = p.ingreso || "";

          // Calculations — básico manda, adicionales son extras sobre él
          const importeSueldo    = sueldoBasico;
          const importeExtras    = valorHoraExt * horasExtra;
          const importeFeriados  = valorDia     * feriados;
          const importeVacaciones= valorDia     * vacaciones;
          const importeFinde     = valorDiaFinde * diasFinde;
          const totalAdicionales = importeExtras + importeFeriados + sac + importeVacaciones + importeFinde;
          const descDemoras      = (valorHora / 4) * fraccionesDemora;
          const descSalTemp      = valorHora    * horasSalTemp;
          const totalDescuentos  = descDemoras + descSalTemp;
          const subtotal         = importeSueldo + totalAdicionales;
          const totalACobrar     = subtotal - totalDescuentos - adelanto;

          const fmt = (n) => n === 0 ? "—" : `$${Math.round(n).toLocaleString("es-AR")}`;
          const fmtN = (n) => n === 0 ? "—" : Math.round(n).toLocaleString("es-AR");

          const Field = ({label, field, prefix="$", note=""}) => (
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
              <label style={{fontSize:12,color:COL.textSub,minWidth:200,flexShrink:0}}>{label}</label>
              <div style={{display:"flex",alignItems:"center",gap:4}}>
                {prefix&&<span style={{fontSize:12,color:COL.textFaint}}>{prefix}</span>}
                <input
                  type="number" min="0" step="any"
                  value={p[field]||""}
                  onChange={e=>setP(field, e.target.value)}
                  placeholder="0"
                  style={{...S.sInput,width:140,padding:"6px 10px",fontFamily:MONO,fontSize:13}}
                />
              </div>
              {note&&<span style={{fontSize:11,color:COL.textFaint}}>{note}</span>}
            </div>
          );

          const LiqRow = ({label,cantidad,valor,importe,bold,color,indent,separator}) => (
            <>
              {separator&&<tr><td colSpan={4} style={{height:1,background:COL.border,padding:0}}></td></tr>}
              <tr style={{background: bold?"#f0f4fa":"transparent"}}>
                <td style={{...S.td,textAlign:"left",paddingLeft:indent?28:14,color:bold?COL.text:COL.textSub,fontWeight:bold?600:400,fontSize:13}}>{label}</td>
                <td style={{...S.td,fontFamily:MONO,fontSize:12,color:COL.textFaint}}>{cantidad}</td>
                <td style={{...S.td,fontFamily:MONO,fontSize:12,color:COL.textFaint}}>{valor}</td>
                <td style={{...S.td,fontFamily:MONO,fontSize:13,color:color||(bold?COL.text:COL.textSub),fontWeight:bold?700:400,textAlign:"right",paddingRight:18}}>{importe}</td>
              </tr>
            </>
          );

          return (
            <div style={{maxWidth:860}}>
              <H2>Liquidación de sueldos</H2>
              <p style={S.body}>Seleccioná el empleado, completá los valores de referencia y el sistema calcula la liquidación usando los datos de asistencia del período.</p>

              {/* Employee + period selector */}
              <div style={{display:"flex",gap:12,marginBottom:24,flexWrap:"wrap",alignItems:"flex-end"}}>
                <div>
                  <div style={{fontSize:11,color:COL.textFaint,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:5}}>Empleado</div>
                  <select value={liqEmp||""} onChange={e=>setLiqEmp(e.target.value?Number(e.target.value):null)}
                    style={{background:COL.surface,border:`1px solid ${COL.border2}`,borderRadius:8,color:COL.text,padding:"8px 14px",fontFamily:SANS,fontSize:13,outline:"none",minWidth:240,cursor:"pointer"}}>
                    <option value="">— Elegí un empleado —</option>
                    {empOptions.map(e=>(
                      <option key={e.empNo} value={e.empNo}>{e.empNo} · {cap(e.nombre)}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <div style={{fontSize:11,color:COL.textFaint,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:5}}>Período</div>
                  <input value={periodo} onChange={e=>setP("periodo",e.target.value)} placeholder="ej: MARZO 2026"
                    style={{...S.sInput,width:160}} />
                </div>
                <div>
                  <div style={{fontSize:11,color:COL.textFaint,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:5}}>Fecha ingreso</div>
                  <input type="date" value={ingreso} onChange={e=>setP("ingreso",e.target.value)}
                    style={{...S.dateInput}} />
                </div>
                <div style={{display:"flex",gap:8,alignItems:"flex-end"}}>
                  <div>
                    <div style={{fontSize:11,color:COL.textFaint,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:5}}>Desde</div>
                    <input type="date" value={desde} onChange={e=>setP("desde",e.target.value)} style={{...S.dateInput}}/>
                  </div>
                  <div style={{paddingBottom:9,color:COL.textFaint,fontSize:13}}>—</div>
                  <div>
                    <div style={{fontSize:11,color:COL.textFaint,fontWeight:500,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:5}}>Hasta</div>
                    <input type="date" value={hasta} onChange={e=>setP("hasta",e.target.value)} style={{...S.dateInput}}/>
                  </div>
                </div>
                {selEmp&&(
                  <button onClick={()=>exportLiqPDF({selEmp,periodo,ingreso,desde,hasta,importeSueldo,diasFinde,valorDiaFinde,importeFinde,horasExtra,horasExtraDisplay,valorHoraExt,importeExtras,feriados,valorDia,importeFeriados,sac,vacaciones,importeVacaciones,totalAdicionales,subtotal,fraccionesDemora,valorHora,descDemoras,horasSalTemp,descSalTemp,totalDescuentos,adelanto,totalACobrar,diasTrabajados,nombreDisplay:p.nombreDisplay||(cap(selEmp.nombre)),fmt})}
                    style={{alignSelf:"flex-end",background:"#276749",color:"#fff",border:"none",borderRadius:8,padding:"9px 20px",cursor:"pointer",fontFamily:SANS,fontWeight:600,fontSize:13,display:"flex",alignItems:"center",gap:8,whiteSpace:"nowrap"}}>
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><path d="M6 9V2h12v7"/><path d="M6 18H4a2 2 0 01-2-2v-5a2 2 0 012-2h16a2 2 0 012 2v5a2 2 0 01-2 2h-2"/><path d="M6 14h12v8H6z"/></svg>
                    Exportar PDF
                  </button>
                )}
              </div>

              {!selEmp&&(
                <div style={{padding:"60px 20px",textAlign:"center",color:COL.textFaint,fontSize:13,background:COL.surface,borderRadius:12,border:`1px solid ${COL.border}`}}>
                  Elegí un empleado para generar la liquidación.
                </div>
              )}

              {selEmp&&(
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,alignItems:"start"}}>

                  {/* LEFT: manual inputs */}
                  <div >
                    <div style={{background:COL.surface,border:`1px solid ${COL.border}`,borderRadius:12,padding:"18px 20px",marginBottom:16}}>
                      <div style={{fontSize:11,color:COL.textFaint,fontWeight:600,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:14}}>Valores de referencia</div>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12,paddingBottom:12,borderBottom:`1px solid ${COL.border}`}}>
                        <label style={{fontSize:12,color:COL.textSub,minWidth:200,flexShrink:0}}>Nombre en liquidación</label>
                        <input
                          value={p.nombreDisplay||""}
                          onChange={e=>setP("nombreDisplay",e.target.value)}
                          placeholder={cap(selEmp.nombre)}
                          style={{...S.sInput,flex:1,padding:"6px 10px",fontSize:13}}
                        />
                      </div>
                      <Field label="Sueldo básico"            field="sueldoBasico" />
                      <Field label="Valor día"                field="valorDia" />
                      <Field label="Valor hora"               field="valorHora" />
                      <Field label="Valor hora extra"         field="valorHoraExt" />
                      <Field label="Adicional día finde/especial" field="valorDiaFinde" note="por día trabajado en finde" />
                    </div>

                    <div style={{background:COL.surface,border:`1px solid ${COL.border}`,borderRadius:12,padding:"18px 20px",marginBottom:16}}>
                      <div style={{fontSize:11,color:COL.textFaint,fontWeight:600,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:14}}>Adicionales manuales</div>
                      <Field label="SAC"                      field="sac"       note="importe directo" />
                      <Field label="Vacaciones (días)"        field="vacaciones" prefix="" note="× valor día" />
                      <Field label="Feriados (días)"          field="feriados"  prefix="" note="× valor día" />
                      <Field label="Adelanto"                 field="adelanto" />

                      {/* Días finde checkboxes */}
                      {allFindeInRange.length > 0 && (
                        <div style={{marginTop:14,paddingTop:14,borderTop:`1px solid ${COL.border}`}}>
                          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                            <span style={{fontSize:12,color:COL.textSub,fontWeight:500}}>
                              Días finde/especiales
                              <span style={{marginLeft:8,fontFamily:MONO,color:COL.accent,fontWeight:700}}>{diasFinde}</span>
                              <span style={{color:COL.textFaint}}> seleccionados</span>
                            </span>
                            <div style={{display:"flex",gap:6}}>
                              <button onClick={()=>setP("findeSel",[...allFindeInRange])}
                                style={{fontSize:11,padding:"2px 8px",border:`1px solid ${COL.border2}`,borderRadius:5,cursor:"pointer",background:"#fff",color:COL.textSub,fontFamily:SANS}}>
                                Todos
                              </button>
                              <button onClick={()=>setP("findeSel",[])}
                                style={{fontSize:11,padding:"2px 8px",border:`1px solid ${COL.border2}`,borderRadius:5,cursor:"pointer",background:"#fff",color:COL.textSub,fontFamily:SANS}}>
                                Ninguno
                              </button>
                            </div>
                          </div>
                          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                            {allFindeInRange.map(fecha=>{
                              const checked = findeSel.has(fecha);
                              const d = new Date(fecha+"T12:00:00");
                              const dia = DIAS_CORTO[d.getDay()];
                              const num = fecha.slice(8);
                              const hasRec = rangeCalcs.some(r=>r.fecha===fecha&&r.trabajado!=null);
                              return (
                                <label key={fecha} style={{display:"flex",alignItems:"center",gap:5,padding:"5px 10px",border:`1px solid ${checked?COL.accent:COL.border}`,borderRadius:7,cursor:"pointer",background:checked?COL.accentBg:"#fff",transition:"all .12s"}}>
                                  <input type="checkbox" checked={checked} onChange={()=>toggleFinde(fecha)}
                                    style={{cursor:"pointer",accentColor:COL.accent,width:13,height:13}}/>
                                  <span style={{fontFamily:MONO,fontSize:12,color:checked?COL.accent:COL.textSub,fontWeight:checked?600:400}}>
                                    {dia} {num}
                                  </span>
                                  {!hasRec && <span style={{fontSize:10,color:"#b45309",marginLeft:2}}>sin reg.</span>}
                                </label>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Attendance data summary */}
                    <div style={{background:COL.accentBg,border:`1px solid ${COL.accentSoft}`,borderRadius:12,padding:"16px 20px"}}>
                      <div style={{fontSize:11,color:COL.accent,fontWeight:600,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:4}}>Datos de asistencia del período</div>
                      {(desde||hasta)&&<div style={{fontSize:11,color:COL.textFaint,marginBottom:10}}>{desde||"inicio"} → {hasta||"fin"} · {rangeCalcs.length} días en rango</div>}
                      {[
                        ["Días trabajados",        diasTrabajados,       "días (referencia)"],
                        ["Días finde trabajados",  diasFinde,            "días"],
                        ["Horas extra",            horasExtraDisplay,    ""],
                        ["Demoras (fracc. 15 min)",fraccionesDemora,     "fracciones"],
                        ["Retiros anticipados",    horasSalTemp,         "hs"],
                      ].map(([l,v,u])=>(
                        <div key={l} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:7}}>
                          <span style={{fontSize:12,color:COL.textSub}}>{l}</span>
                          <span style={{fontFamily:MONO,fontSize:13,fontWeight:600,color:COL.accent}}>{v} <span style={{fontSize:11,fontWeight:400,color:COL.textFaint}}>{u}</span></span>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* RIGHT: liquidacion sheet */}
                  <div >
                    <div style={{background:COL.surface,border:`1px solid ${COL.border}`,borderRadius:12,overflow:"hidden"}}>

                      {/* Header */}
                      <div style={{background:COL.accent,color:"#fff",padding:"14px 18px"}}>
                        <div style={{fontWeight:700,fontSize:14,letterSpacing:"0.03em"}}>PLANILLA LIQUIDACIÓN SUELDOS</div>
                        <div style={{fontSize:12,opacity:0.85,marginTop:3,display:"flex",gap:20}}>
                          <span>Período: <strong>{periodo||"—"}</strong></span>
                          <span>Ingreso: <strong>{ingreso||"—"}</strong></span>
                        </div>
                        <div style={{fontSize:13,marginTop:6,fontWeight:600}}>{cap(selEmp.nombre)}</div>
                      </div>

                      <table style={{...S.table,fontSize:12}}>
                        <thead><tr>
                          <th style={{...S.th,textAlign:"left"}}>Concepto</th>
                          <th style={S.th}>Cantidad</th>
                          <th style={S.th}>Valor unit.</th>
                          <th style={{...S.th,textAlign:"right",paddingRight:18}}>Importe</th>
                        </tr></thead>
                        <tbody>
                          <LiqRow label="SUELDO BÁSICO" cantidad="" valor="" importe={fmt(importeSueldo)} bold />

                          <LiqRow label="ADICIONALES" cantidad="" valor="" importe="" bold separator />
                          <LiqRow label="Horas extra"       indent cantidad={horasExtraDisplay} valor={fmt(valorHoraExt)} importe={fmt(importeExtras)} />
                          <LiqRow label="Días finde/especiales" indent cantidad={diasFinde||"—"} valor={diasFinde?fmt(valorDiaFinde):"—"} importe={diasFinde&&valorDiaFinde?fmt(importeFinde):"—"} />
                          <LiqRow label="Feriados"          indent cantidad={feriados||"—"}  valor={fmt(valorDia)}     importe={feriados?fmt(importeFeriados):"—"} />
                          <LiqRow label="SAC"               indent cantidad=""              valor=""                  importe={sac?fmt(sac):"—"} />
                          <LiqRow label="Vacaciones"        indent cantidad={vacaciones||"—"} valor={fmt(valorDia)}   importe={vacaciones?fmt(importeVacaciones):"—"} />
                          <LiqRow label="Subtotal adicionales" cantidad="" valor="" importe={fmt(totalAdicionales)} bold separator />

                          <LiqRow label="Sueldo + adicionales" cantidad="" valor="" importe={fmt(subtotal)} bold color={COL.accent} separator />

                          <LiqRow label="DESCUENTOS" cantidad="" valor="" importe="" bold separator />
                          <LiqRow label="Llegadas tarde (fracc. 15 min)" indent cantidad={fraccionesDemora} valor={`${fmt(valorHora)}/4`} importe={descDemoras?fmt(descDemoras):"—"} color={descDemoras?"#c53030":undefined} />
                          <LiqRow label="Retiros anticipados (x hora)"   indent cantidad={horasSalTemp}     valor={fmt(valorHora)}      importe={descSalTemp?fmt(descSalTemp):"—"} color={descSalTemp?"#c53030":undefined} />
                          <LiqRow label="Total descuentos" cantidad="" valor="" importe={totalDescuentos?fmt(totalDescuentos):"—"} bold color="#c53030" separator />

                          <LiqRow label="Adelanto" cantidad="" valor="" importe={adelanto?fmt(adelanto):"—"} bold color="#b45309" separator />

                          {/* Total */}
                          <tr style={{background:"#f0f4fa"}}>
                            <td colSpan={3} style={{...S.td,textAlign:"right",fontWeight:700,fontSize:13,color:COL.text,borderTop:`2px solid ${COL.border2}`,paddingRight:12}}>TOTAL A COBRAR</td>
                            <td style={{...S.td,fontFamily:MONO,fontSize:15,fontWeight:700,color:"#276749",textAlign:"right",paddingRight:18,borderTop:`2px solid ${COL.border2}`}}>
                              {totalACobrar!==0?`$${Math.round(totalACobrar).toLocaleString("es-AR")}`:"—"}
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>

                </div>
              )}
            </div>
          );
        })()}

      </main>
    </div>
  );
}