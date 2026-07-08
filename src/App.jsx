import { useState, useCallback, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import LoginScreen from "./LoginScreen";

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

// Tolerancia de 15 min para llegadas tarde (solo afecta el descuento en plata).
// La tolerancia se evalúa POR DÍA: cada día con ≤15 min se perdona; los días
// que pasan los 15 cuentan por bloques de 15 desde cero (16-30=1, 31-45=2...).
const TOLERANCIA_DEMORA = 15;
function fraccionesDeUnDia(demoraMin) {
  const d = Math.round(demoraMin || 0);
  if (d <= TOLERANCIA_DEMORA) return 0;
  return Math.ceil(d / 15) - (d % 15 === 0 ? 0 : 1);
}
function fraccionesDemoraCalc(calcsDelRango) {
  if (!Array.isArray(calcsDelRango)) return 0;
  return calcsDelRango.reduce((s, r) => s + fraccionesDeUnDia(r.demora), 0);
}
// Retiros anticipados: misma lógica que las tardanzas (por día, 15 de tolerancia,
// fracciones de 15 min). Cada fracción vale valorHora/4.
function fraccionesSalTempCalc(calcsDelRango) {
  if (!Array.isArray(calcsDelRango)) return 0;
  return calcsDelRango.reduce((s, r) => s + fraccionesDeUnDia(r.salTemprana), 0);
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
    ? { entrada: cfg.entrada, salida: dayType?.salida || "14:00" }
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

const cap = s => s ? s.charAt(0).toUpperCase()+s.slice(1).toLowerCase() : s;

const COL = {
  bg:"#f5f6f8", surface:"#ffffff", border:"#e8ecf1", border2:"#d4dbe4",
  text:"#1e2a38", textSub:"#5a6a7a", textFaint:"#96a3b0",
  accent:"#3d6b9e", accentBg:"#edf2f9", accentSoft:"#dce8f5",
};
const SANS = "'DM Sans', system-ui, sans-serif";
const MONO = "'DM Mono', 'Courier New', monospace";

/* ─── Inputs estables (definidos a nivel módulo para no perder el foco) ─────── */
const INPUT_STYLE = {
  border:`1px solid ${COL.border2}`, borderRadius:8, outline:"none",
  background:"#fff", color:COL.text,
};

function FieldInput({ label, prefix="$", note="", value, onChange }) {
  return (
    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
      <label style={{fontSize:12,color:COL.textSub,minWidth:200,flexShrink:0}}>{label}</label>
      <div style={{display:"flex",alignItems:"center",gap:4}}>
        {prefix&&<span style={{fontSize:12,color:COL.textFaint}}>{prefix}</span>}
        <input
          type="number" min="0" step="any" inputMode="decimal"
          value={value}
          onChange={e=>onChange(e.target.value)}
          placeholder="0"
          style={{...INPUT_STYLE,width:140,padding:"6px 10px",fontFamily:MONO,fontSize:13}}
        />
      </div>
      {note&&<span style={{fontSize:11,color:COL.textFaint}}>{note}</span>}
    </div>
  );
}

// Formulario para iniciar un nuevo período de liquidación para todos los empleados.
function NuevoPeriodoForm({ MESES, defMonth, empCount, onCancel, onConfirm }) {
  const [mes, setMes] = useState(defMonth); // formato "YYYY-MM"
  const [y, m] = mes.split("-").map(Number);
  const ultimoDia = new Date(y, m, 0).getDate();
  const desde = `${mes}-01`;
  const hasta = `${mes}-${String(ultimoDia).padStart(2,"0")}`;
  const nombrePeriodo = `${MESES[m-1].toUpperCase()} ${y}`;

  return (
    <div style={{background:"#fef9f0",border:"1px solid #f6d860",borderRadius:12,padding:"18px 20px",marginBottom:20}}>
      <div style={{fontSize:13,fontWeight:700,color:"#92400e",marginBottom:6}}>Iniciar nuevo período para todos los empleados</div>
      <div style={{fontSize:12,color:"#a16207",marginBottom:16,lineHeight:1.6}}>
        Se aplicará <strong>{nombrePeriodo}</strong> ({desde} → {hasta}) a los {empCount} empleados con datos.
        Se <strong>conservan</strong> los adelantos y los datos de la circular (sueldo, valores). Se <strong>resetean</strong> las
        llegadas tarde y retiros editados a mano, las horas extra manuales, SAC, vacaciones y feriados del mes anterior.
      </div>
      <div style={{display:"flex",gap:14,alignItems:"flex-end",flexWrap:"wrap"}}>
        <div>
          <div style={{fontSize:11,color:COL.textSub,fontWeight:500,marginBottom:5}}>Mes a liquidar</div>
          <input type="month" value={mes} onChange={e=>setMes(e.target.value)}
            style={{...INPUT_STYLE,padding:"8px 12px",fontFamily:SANS,fontSize:13}}/>
        </div>
        <div style={{fontSize:12,color:COL.textSub,paddingBottom:9}}>
          → Período: <strong style={{color:"#92400e"}}>{nombrePeriodo}</strong>
        </div>
        <div style={{display:"flex",gap:8,marginLeft:"auto"}}>
          <button onClick={()=>{
            if(confirm(`¿Iniciar ${nombrePeriodo} para todos los empleados? Se borrarán los descuentos manuales, hs extra manuales, SAC, vacaciones y feriados del período anterior. Los adelantos y la circular se conservan.`)){
              onConfirm(nombrePeriodo, desde, hasta);
            }
          }}
            style={{background:"#b45309",color:"#fff",border:"none",borderRadius:8,padding:"9px 18px",cursor:"pointer",fontFamily:SANS,fontWeight:600,fontSize:13}}>
            Aplicar a todos
          </button>
          <button onClick={onCancel}
            style={{background:"#fff",color:COL.textSub,border:`1px solid ${COL.border2}`,borderRadius:8,padding:"9px 16px",cursor:"pointer",fontFamily:SANS,fontSize:13}}>
            Cancelar
          </button>
        </div>
      </div>
    </div>
  );
}

// Convierte horas decimales (ej 2.75) a {h:2, m:45}
function decimalToHM(dec) {
  const v = parseFloat(dec);
  if (!v || isNaN(v)) return { h:"", m:"" };
  const totalMin = Math.round(v * 60);
  return { h:String(Math.floor(totalMin/60)), m:String(totalMin%60) };
}
// Convierte horas y minutos a horas decimales (ej 2,45 -> 2.75)
function hmToDecimal(h, m) {
  const hh = parseInt(h,10) || 0;
  const mm = parseInt(m,10) || 0;
  if (!hh && !mm) return "";
  return ((hh*60 + mm) / 60).toString();
}

/* ─── PDF export ─────────────────────────────────────────────────────────── */
function exportLiqPDF(d) {
  const { selEmp, periodo, ingreso, desde, hasta, importeSueldo, diasFinde, valorDiaFinde, importeFinde, horasExtra, horasExtraDisplay, valorHoraExt, impExtrasReloj,
    importeExtras, importeExtraManual, horasExtraManualDisplay, feriados, valorDia, importeFeriados, sac, vacaciones,
    importeVacaciones, totalAdicionales, subtotal, fraccionesDemora, valorHora,
    descDemoras, fraccionesSalTemp, descSalTemp, totalDescuentos, adelanto, adelantos,
    totalACobrar, diasTrabajados, nombreDisplay, fmt } = d;

  const nombre = selEmp.nombre.toUpperCase();
  const today  = new Date().toLocaleDateString("es-AR");

  // Try to load logo as base64 from /logo.png (served by Vite from /public)
  const renderPDF = () => {
    const GOLD = "#b08a2e";

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
      .gold-line{border:none;border-top:3px solid #b08a2e;margin:0 0 14px}
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
      @media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
    </style></head><body>
    <div class="wrap">

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
          ${(impExtrasReloj??valorHoraExt*horasExtra)>0 ?row("Horas extra (reloj)",horasExtraDisplay,fmt(valorHoraExt),fmt(impExtrasReloj??valorHoraExt*horasExtra),"detail","",true):""}
          ${importeExtraManual>0?row("Horas extra (manual)",horasExtraManualDisplay,fmt(valorHoraExt),fmt(importeExtraManual),"detail","",true):""}
          ${importeFinde>0 ?row("Días finde/especiales",diasFinde||"—",fmt(valorDiaFinde),fmt(importeFinde),"detail","",true):""}
          ${importeFeriados>0 ?row("Feriados",feriados||"—",fmt(valorDia),fmt(importeFeriados),"detail","",true):row("Feriados","—",fmt(valorDia),"—","muted","",true)}
          ${row("SAC","—","—",sac>0?fmt(sac):"—",sac>0?"detail":"muted","",true)}
          ${importeVacaciones>0 ?row("Vacaciones",vacaciones||"—",fmt(valorDia),fmt(importeVacaciones),"detail","",true):row("Vacaciones","—",fmt(valorDia),"—","muted","",true)}
          ${row("Subtotal adicionales","","",fmt(totalAdicionales),"sub")}
          ${row("SUELDO + ADICIONALES","","",fmt(subtotal),"total-line","#1a3a6b")}
          ${row("DESCUENTOS","","","","section")}
          ${descDemoras>0?row("Llegadas tarde (fracc. de 15 min)",fraccionesDemora,`${fmt(valorHora)}/4`,fmt(descDemoras),"detail","#c53030",true):row("Llegadas tarde (fracc. de 15 min)","—","—","—","muted","",true)}
          ${descSalTemp>0?row("Retiros anticipados (fracc. de 15 min)",fraccionesSalTemp,`${fmt(valorHora)}/4`,fmt(descSalTemp),"detail","#c53030",true):row("Retiros anticipados (fracc. de 15 min)","—","—","—","muted","",true)}
          ${row("Total descuentos","","",totalDescuentos>0?fmt(totalDescuentos):"—","sub",totalDescuentos>0?"#c53030":"")}
          ${(adelantos||[]).filter(a=>parseFloat(a.monto)>0).length
            ? row("ADELANTOS","","","","section","#b45309")
              + (adelantos||[]).filter(a=>parseFloat(a.monto)>0)
                  .map(a=>row(a.desc||"Adelanto","","",fmt(parseFloat(a.monto)),"detail","#b45309",true)).join("")
              + row("Total adelantos","","",fmt(adelanto),"sub","#b45309")
            : row("Adelantos","—","—","—","muted")}
          <tr class="total-gold">
            <td colspan="3" style="text-align:right;padding-right:16px;letter-spacing:0.04em">TOTAL A COBRAR — ${periodo||""}</td>
            <td>${totalACobrar>0?fmt(totalACobrar):"—"}</td>
          </tr>
        </tbody>
      </table>


    </div>
    </body></html>`;

    const win = window.open("", "_blank", "width=920,height=780");
    win.document.write(html);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); }, 700);
  };

  renderPDF();
}

/* ─── Exportar Novedades a PDF ───────────────────────────────────────────── */
// Abre una ventana de impresión con la tabla de novedades (mismo estilo sobrio
// que el PDF de liquidación). Usa window.print() para que el usuario guarde
// como PDF o imprima directamente.
function exportNovedadesPDF({ conMarca, sinMarca, ultimoDia, fmtFecha }) {
  const today = new Date().toLocaleDateString("es-AR");
  const esc = s => String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");

  const filaHTML = (f, faded) => `
    <tr${faded?' style="color:#888"':''}>
      <td style="padding:5px 8px;font-family:monospace;color:#888;border-bottom:1px solid #eee">${esc(f.emp.empNo)}</td>
      <td style="padding:5px 10px;border-bottom:1px solid #eee">${esc(cap(f.emp.nombre))}</td>
      <td style="padding:5px 8px;text-align:center;font-size:10px;color:#777;border-bottom:1px solid #eee">${esc(fmtFecha(f.fechaIng))}</td>
      <td style="padding:5px 8px;text-align:center;font-family:monospace;font-weight:700;color:${f.ingreso?'#1a6b3f':'#bbb'};border-bottom:1px solid #eee">${esc(f.ingreso||"—")}</td>
      <td style="padding:5px 8px;text-align:center;font-size:10px;color:#777;border-bottom:1px solid #eee">${esc(fmtFecha(f.fechaSal))}</td>
      <td style="padding:5px 8px;text-align:center;font-family:monospace;font-weight:700;color:${f.salida?'#1e5fa8':'#bbb'};border-bottom:1px solid #eee">${esc(f.salida||"—")}</td>
      <td style="padding:5px 10px;font-size:10.5px;color:#444;border-bottom:1px solid #eee">${esc(f.observacion||"")}</td>
    </tr>`;

  const seccion = (titulo, color, filas, fade) => filas.length ? `
    <div style="font-size:11px;font-weight:700;color:${color};margin:16px 0 6px">${esc(titulo)} · ${filas.length}</div>
    <table style="width:100%;border-collapse:collapse;font-size:11px">
      <thead><tr style="background:#1a1a1a;color:#fff">
        <th style="padding:6px 8px;text-align:left;font-size:9px;letter-spacing:0.05em">N°</th>
        <th style="padding:6px 10px;text-align:left;font-size:9px;letter-spacing:0.05em">NOMBRE</th>
        <th style="padding:6px 8px;font-size:9px;letter-spacing:0.05em">FECHA ING.</th>
        <th style="padding:6px 8px;font-size:9px;letter-spacing:0.05em">INGRESO</th>
        <th style="padding:6px 8px;font-size:9px;letter-spacing:0.05em">FECHA SAL.</th>
        <th style="padding:6px 8px;font-size:9px;letter-spacing:0.05em">SALIDA</th>
        <th style="padding:6px 10px;text-align:left;font-size:9px;letter-spacing:0.05em">OBSERVACIONES</th>
      </tr></thead>
      <tbody>${filas.map(f=>filaHTML(f,fade)).join("")}</tbody>
    </table>` : "";

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
  <style>
    @page { size:A4 portrait; margin:14mm 12mm; }
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Helvetica Neue',Arial,sans-serif;background:#fff;color:#1a1a1a;font-size:11px}
    .wrap{padding:6px 4px}
    .gold-line{border:none;border-top:3px solid #b08a2e;margin:0 0 12px}
    .title-row{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:10px}
    .doc-title{font-size:15px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px}
    .doc-meta{font-size:10px;color:#555;text-align:right;line-height:1.8}
    @media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
  </style></head><body>
  <div class="wrap">
    <hr class="gold-line"/>
    <div class="title-row">
      <div class="doc-title">Novedades</div>
      <div class="doc-meta">
        Último día: <b>${esc(fmtFecha(ultimoDia))}</b><br>
        Emitido: ${esc(today)}
      </div>
    </div>
    ${seccion(`Presentes el ${fmtFecha(ultimoDia)}`, "#276749", conMarca, false)}
    ${seccion(`Sin marca el ${fmtFecha(ultimoDia)} (última marca registrada)`, "#b45309", sinMarca, true)}
  </div>
  </body></html>`;

  const win = window.open("", "_blank", "width=920,height=780");
  if (!win) { alert("El navegador bloqueó la ventana emergente. Permití pop-ups para exportar el PDF."); return; }
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(() => { win.print(); }, 500);
}

/* ─── Exportar a Excel ───────────────────────────────────────────────────── */
// rows: array de objetos; el orden de columnas lo da `headers` (array de [key,label])
function exportXLSX(rows, headers, sheetName, fileName) {
  const aoa = [headers.map(h=>h[1])];
  for (const r of rows) aoa.push(headers.map(h=>{
    const v = r[h[0]];
    return v === null || v === undefined ? "" : v;
  }));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  // ancho de columnas automático según contenido
  ws["!cols"] = headers.map((h,i)=>{
    const maxLen = Math.max(h[1].length, ...aoa.slice(1).map(row=>String(row[i]??"").length));
    return { wch: Math.min(Math.max(maxLen+2, 8), 40) };
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0,31));
  XLSX.writeFile(wb, fileName);
}

/* ─── Supabase ───────────────────────────────────────────────────────────── */
const SB_URL = import.meta.env.VITE_SUPABASE_URL;
const SB_KEY = import.meta.env.VITE_SUPABASE_SERVICE_KEY;
const SB_ANON = import.meta.env.VITE_SUPABASE_ANON_KEY;

const SB_HEADERS = (write=false) => ({
  "Content-Type": "application/json",
  "Authorization": `Bearer ${SB_KEY}`,
  "apikey": SB_KEY,
  "Accept-Profile": "rrhh",
  ...(write ? {"Content-Profile":"rrhh"} : {}),
});

// ── Read all rows from a table ────────────────────────────────────────────
// Devuelve un array si la lectura fue OK (puede estar vacío), o null si hubo
// un error real de red/permiso. Esto permite distinguir "no hay datos todavía"
// de "no me pude conectar".
async function sbFetch(table, params="") {
  if (!SB_URL || !SB_KEY) return null;
  try {
    const res = await fetch(`${SB_URL}/rest/v1/${table}?${params}`, {
      headers: SB_HEADERS(false),
    });
    if (!res.ok) {
      let body = ""; try { body = await res.text(); } catch {}
      console.error(`[Supabase] Error al leer "${table}": HTTP ${res.status} ${body}`.trim());
      return null;
    }
    return await res.json();
  } catch (err) {
    console.error(`[Supabase] Error de red al leer "${table}":`, err?.message || err);
    return null;
  }
}

// ── Hook opcional para reportar fallos de escritura a la UI ───────────────
// Se setea desde el componente con setSbWriteError. Si una escritura falla,
// se llama con un objeto {table, detail} para que RRHH lo vea.
let _sbOnWriteError = null;
function setSbWriteErrorHandler(fn) { _sbOnWriteError = fn; }
function _reportWriteError(table, detail) {
  console.error(`[Supabase] Error al guardar en "${table}":`, detail);
  if (_sbOnWriteError) _sbOnWriteError({ table, detail: String(detail), at: Date.now() });
}

// Intenta una petición POST/DELETE; reintenta 1 vez ante error de red.
// Devuelve true si Supabase confirmó la escritura, false si falló.
async function _sbWrite(url, options, table, attempt = 0) {
  try {
    const res = await fetch(url, options);
    if (!res.ok) {
      // 4xx/5xx: leer el cuerpo para diagnóstico (constraint, permisos, etc.)
      let body = "";
      try { body = await res.text(); } catch {}
      _reportWriteError(table, `HTTP ${res.status} ${res.statusText} ${body}`.trim());
      return false;
    }
    return true;
  } catch (err) {
    // Error de red: reintentar una sola vez tras breve espera
    if (attempt < 1) {
      await new Promise(r => setTimeout(r, 800));
      return _sbWrite(url, options, table, attempt + 1);
    }
    _reportWriteError(table, err?.message || err);
    return false;
  }
}

// ── Upsert many rows ─────────────────────────────────────────────────────
async function sbUpsert(table, rows, pk="id") {
  if (!SB_URL || !SB_KEY) { _reportWriteError(table, "Supabase no configurado (faltan VITE_SUPABASE_*)"); return false; }
  if (!rows?.length) return true;
  return _sbWrite(`${SB_URL}/rest/v1/${table}?on_conflict=${pk}`, {
    method: "POST",
    headers: { ...SB_HEADERS(true), "Prefer": "resolution=merge-duplicates" },
    body: JSON.stringify(rows),
  }, table);
}

// ── Upsert single row ────────────────────────────────────────────────────
async function sbUpsertSingle(table, row, pk="id") {
  if (!SB_URL || !SB_KEY) { _reportWriteError(table, "Supabase no configurado (faltan VITE_SUPABASE_*)"); return false; }
  return _sbWrite(`${SB_URL}/rest/v1/${table}?on_conflict=${pk}`, {
    method: "POST",
    headers: { ...SB_HEADERS(true), "Prefer": "resolution=merge-duplicates" },
    body: JSON.stringify(row),
  }, table);
}

// ── Update parcial de una fila existente (PATCH) ──────────────────────────
// A diferencia del upsert, PATCH actualiza SOLO los campos enviados y deja
// intactos los demás. Se usa para editar un campo (ej. observación) de un
// registro que ya existe, sin riesgo de violar constraints not-null de las
// columnas que no se mandan (emp_no, fecha, etc.).
async function sbUpdate(table, id, patch, pk="id") {
  if (!SB_URL || !SB_KEY) { _reportWriteError(table, "Supabase no configurado (faltan VITE_SUPABASE_*)"); return false; }
  return _sbWrite(`${SB_URL}/rest/v1/${table}?${pk}=eq.${encodeURIComponent(id)}`, {
    method: "PATCH",
    headers: SB_HEADERS(true),
    body: JSON.stringify(patch),
  }, table);
}

// ── Delete row ───────────────────────────────────────────────────────────
async function sbDelete(table, id, pk="id") {
  if (!SB_URL || !SB_KEY) { _reportWriteError(table, "Supabase no configurado (faltan VITE_SUPABASE_*)"); return false; }
  return _sbWrite(`${SB_URL}/rest/v1/${table}?${pk}=eq.${encodeURIComponent(id)}`, {
    method: "DELETE",
    headers: SB_HEADERS(true),
  }, table);
}

// ── Realtime subscription via Supabase websocket ─────────────────────────
function sbSubscribe(table, onInsert, onUpdate, onDelete) {
  if (!SB_URL || !SB_KEY) return () => {};
  const wsUrl = SB_URL.replace("https://","wss://").replace("http://","ws://");
  const key   = SB_KEY;
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
    observacion: r.observacion || null,
    ausencia: r.ausencia || null,
  };
}

function rowToRec(r) {
  return {
    id: r.id, empNo: r.emp_no, nombre: r.nombre, depto: r.depto,
    fecha: r.fecha, entrada: r.entrada, salida: r.salida,
    soloEntrada: r.solo_entrada, manual: r.manual,
    observacion: r.observacion || null,
    ausencia: r.ausencia || null,
  };
}

function empToRow(e) {
  return {
    emp_no: e.empNo, nombre: e.nombre,
    nombre_display: e.nombreDisplay || null,
    depto: e.depto, entrada_ref: e.entrada, salida_ref: e.salida,
    activo: e.activo !== false,
    auto_detected: e.autoDetected || false,
    tipo: e.tipo || "operario",
  };
}

function rowToEmp(r) {
  return {
    empNo: r.emp_no, nombre: r.nombre,
    nombreDisplay: r.nombre_display,
    depto: r.depto, entrada: r.entrada_ref, salida: r.salida_ref,
    activo: r.activo, autoDetected: r.auto_detected,
    tipo: r.tipo || "operario",
  };
}

  function diaToRow(fecha, tipo) { return { fecha, tipo }; }

// ── Histórico de circulares ───────────────────────────────────────────────
function circHistToRow(empNo, nombre, datos, motivo, archivedAt) {
  return {
    id: `${empNo}_${archivedAt}`,
    emp_no: empNo,
    nombre: nombre || null,
    archived_at: archivedAt,
    motivo: motivo || "manual",
    datos,
  };
}

async function sbInsertCircHist(row) {
  if (!SB_URL || !SB_KEY) { _reportWriteError("circular_historico", "Supabase no configurado"); return false; }
  return _sbWrite(`${SB_URL}/rest/v1/circular_historico`, {
    method: "POST",
    headers: { ...SB_HEADERS(true), "Prefer": "return=minimal" },
    body: JSON.stringify(row),
  }, "circular_historico");
}


const TABS = ["Importar","Registros","Empleados","Por empleado","Calendario","Circular","Resumen","Liquidación","Novedades"];

function AppMain({ session }) {
  const [tab, setTab]             = useState(0);
  const [records, setRecords]     = useState([]);
  const [employees, setEmployees] = useState(()=>makeDefaultEmployees());
  // Ref para leer el estado actual de empleados desde callbacks estables
  const employeesRef = useRef(employees);
  useEffect(()=>{ employeesRef.current = employees; },[employees]);
  const [empTipoF, setEmpTipoF] = useState("todos");
  const [empActivoF, setEmpActivoF] = useState("todos"); // "todos" | "activo" | "inactivo"
  const [sbLoading, setSbLoading] = useState(false);
  const [sbStatus,  setSbStatus]  = useState(""); // "synced" | "error" | ""
  const [sbLastSync,setSbLastSync]= useState(null);
  const [sbWriteError, setSbWriteError] = useState(null); // {table, detail, at} último fallo de guardado
  const [importing, setImporting] = useState(false);
  const [importMsg, setImportMsg] = useState({text:"",ok:true});
  const [recF, setRecF]           = useState({emp:"",desde:"",hasta:""});
  const [empF, setEmpF]           = useState("");
  const [detalleEmp, setDetalleEmp] = useState(null); // empNo selected in Por empleado tab
  const [liqEmp, setLiqEmp]         = useState(null);   // empNo selected in Liquidación tab
  const [circularF, setCircularF]   = useState("");      // filter in Circular tab
  const [circularEdit, setCircularEdit] = useState(null); // empNo being edited in Circular
  const [circularDraft, setCircularDraft] = useState({});
  const [specialDays, setSpecialDays] = useState({});       // fecha→{tipo}
  const [manualSalidas, setManualSalidas] = useState({});   // id→salida manual
  const [manualRecords, setManualRecords] = useState([]);   // registros 100% manuales
  const [editingCell, setEditingCell]     = useState(null); // {id, field}
  const [addingRec, setAddingRec]         = useState(false);
  const [newRec, setNewRec]               = useState({empNo:"",fecha:"",entrada:"",salida:"",observacion:"",ausencia:""});
  const [liqParams, setLiqParams]   = useState({});
  const [editingEmp, setEditingEmp] = useState(null);
  const [editDraft, setEditDraft]   = useState({});
  const [bulkH, setBulkH]         = useState({entrada:"06:00",salida:"16:30"});
  const [bulkSel, setBulkSel]     = useState(new Set());
  const [circHist, setCircHist]   = useState([]);        // versiones archivadas de circulares
  const [histEmp, setHistEmp]     = useState(null);      // empNo cuyo histórico se ve abierto
  const [verPlanillaHist, setVerPlanillaHist] = useState(false); // planilla global de histórico
  const [nuevoPeriodoOpen, setNuevoPeriodoOpen] = useState(false); // modal de nuevo período
  const [resumenMes, setResumenMes] = useState(""); // "YYYY-MM" filtro de mes en Resumen
  const fileRef = useRef();

  // ── Sincronización de liqParams con Supabase (única persistencia) ────────
  // No hay cache local. La cola de pendientes vive solo en memoria: si un
  // upsert falla se reintenta cada 5s mientras la pestaña esté abierta, y
  // si intentan cerrar/refrescar con pendientes, el navegador avisa.
  const prevLiqRef    = useRef({});   // último estado confirmado por el servidor
  const liqSyncTimer  = useRef(null);
  const liqParamsRef  = useRef(liqParams);
  const liqPendingRef = useRef({});   // { empNo: true } upserts sin confirmar
  const liqReadyRef   = useRef(!SB_URL || !SB_KEY);

  const flushLiqPending = useCallback(async () => {
    if (!liqReadyRef.current) return; // esperar la carga inicial de Supabase
    const current = liqParamsRef.current;
    let huboFallo = false;
    for (const key of Object.keys({ ...liqPendingRef.current })) {
      const n = parseInt(key);
      const datos = current[key];
      if (isNaN(n) || n <= 0 || datos === undefined) { delete liqPendingRef.current[key]; continue; }
      const ok = await sbUpsertSingle("liq_params", { emp_no: n, datos }, "emp_no");
      if (ok) {
        prevLiqRef.current[key] = JSON.parse(JSON.stringify(datos));
        delete liqPendingRef.current[key];
      } else {
        huboFallo = true;
      }
    }
    if (huboFallo) {
      clearTimeout(liqSyncTimer.current);
      liqSyncTimer.current = setTimeout(flushLiqPending, 5000);
    }
  }, []);

  useEffect(() => {
    liqParamsRef.current = liqParams;
    for (const [key, datos] of Object.entries(liqParams)) {
      if (JSON.stringify(prevLiqRef.current[key]) !== JSON.stringify(datos)) {
        liqPendingRef.current[key] = true;
      }
    }
    if (!Object.keys(liqPendingRef.current).length) return;
    clearTimeout(liqSyncTimer.current);
    liqSyncTimer.current = setTimeout(flushLiqPending, 700);
  }, [liqParams, flushLiqPending]);

  // Aviso del navegador si cierran/refrescan con cambios aún no confirmados
  useEffect(() => {
    const warn = (e) => {
      if (Object.keys(liqPendingRef.current).length) { e.preventDefault(); e.returnValue = ""; }
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, []);

  // ── Registrar handler global de errores de escritura ─────────────────────
  useEffect(()=>{
    setSbWriteErrorHandler((e)=>{
      setSbWriteError(e);
      setSbStatus("error");
    });
    return ()=>setSbWriteErrorHandler(null);
  },[]);

  // Limpieza única: borrar restos de la persistencia local anterior
  useEffect(() => {
    for (const k of ["ar3","ae3","liq_params","sp_days","man_sal","man_rec","liq_pending"]) {
      try { localStorage.removeItem(k); } catch {}
    }
  }, []);


  // ── Load from Supabase on mount ──────────────────────────────────────────
  useEffect(()=>{
    if (!SB_URL || !SB_KEY) return;
    setSbLoading(true);
    Promise.all([
      sbFetch("registros","select=*&limit=10000"),
      sbFetch("empleados","select=*"),
      sbFetch("dias_especiales","select=*"),
      sbFetch("correcciones","select=*"),
      sbFetch("liq_params","select=*"),
      sbFetch("circular_historico","select=*&order=archived_at.desc"),
    ]).then(([regs,emps,dias,corrs,liqRows,histRows])=>{
      // Solo sobreescribir si Supabase trae datos — nunca borrar con array vacío
      if (regs?.length) {
        const byId={};
        for(const r of regs) byId[r.id]=rowToRec(r);
        const todos = Object.values(byId);
        // Los registros manuales viven en la misma tabla con manual=true.
        // Antes venían de localStorage y se duplicaban con los del server.
        setRecords(todos.filter(r=>!r.manual));
        setManualRecords(todos.filter(r=>r.manual));
      }
      if (emps?.length) {
        const map={...makeDefaultEmployees()};
        for(const e of emps) map[e.emp_no]={...map[e.emp_no],...rowToEmp(e)};
        setEmployees(map);
      }
      if (dias?.length) {
        const map={};
        for(const d of dias) map[d.fecha]={tipo:d.tipo, salida:d.salida||undefined};
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
      {
        // SERVIDOR GANA: el estado se reemplaza por lo de Supabase. Solo
        // sobreviven las ediciones propias en vuelo (cola de pendientes).
        const serverMap = {};
        for (const row of (liqRows || [])) serverMap[String(row.emp_no)] = row.datos;
        setLiqParams(local => {
          const merged = { ...serverMap };
          for (const key of Object.keys(liqPendingRef.current)) {
            if (local[key] !== undefined) merged[key] = local[key];
          }
          return merged;
        });
        prevLiqRef.current = JSON.parse(JSON.stringify(serverMap));
        liqReadyRef.current = true;
        if (Object.keys(liqPendingRef.current).length) {
          clearTimeout(liqSyncTimer.current);
          liqSyncTimer.current = setTimeout(flushLiqPending, 700);
        }
      }
      if (histRows?.length) setCircHist(histRows);
      setSbLoading(false); setSbStatus("synced"); setSbLastSync(new Date());
    }).catch((err)=>{ console.error("Supabase load error:", err); liqReadyRef.current = true; setSbLoading(false); setSbStatus("error"); });
  },[]);

  // ── Realtime subscriptions ───────────────────────────────────────────────
  useEffect(()=>{
    if (!SB_URL || !SB_KEY) return;

    const applyReg = (row) => {
      const rec = rowToRec(row);
      if (rec.manual) {
        setManualRecords(p=>[...p.filter(x=>x.id!==rec.id), rec]);
        setRecords(p=>p.filter(x=>x.id!==rec.id));
      } else {
        setRecords(p=>{const m={};for(const x of p)m[x.id]=x;m[rec.id]=rec;return Object.values(m);});
      }
    };
    const unsubRegs = sbSubscribe("registros",
      applyReg,
      applyReg,
      (r)=>{ setRecords(p=>p.filter(x=>x.id!==r.id)); setManualRecords(p=>p.filter(x=>x.id!==r.id)); }
    );

    const unsubEmps = sbSubscribe("empleados",
      (r)=>setEmployees(p=>({...p,[r.emp_no]:{...(p[r.emp_no]||{}),...rowToEmp(r)}})),
      (r)=>setEmployees(p=>({...p,[r.emp_no]:{...(p[r.emp_no]||{}),...rowToEmp(r)}})),
      null
    );

    const unsubDias = sbSubscribe("dias_especiales",
      (r)=>setSpecialDays(p=>({...p,[r.fecha]:{tipo:r.tipo, salida:r.salida||undefined}})),
      (r)=>setSpecialDays(p=>({...p,[r.fecha]:{tipo:r.tipo, salida:r.salida||undefined}})),
      (r)=>setSpecialDays(p=>{const u={...p};delete u[r.fecha];return u;})
    );

    return ()=>{ unsubRegs(); unsubEmps(); unsubDias(); };
  },[]);

  // Archiva la versión actual de la circular de un empleado como copia histórica
  const archivarCircular = useCallback((empNo, motivo="manual") => {
    const datos = liqParams[String(empNo)];
    if (!datos || !datos.sueldoBasico) return false; // nada que archivar
    const emp = employees[empNo];
    const archivedAt = new Date().toISOString();
    const row = circHistToRow(empNo, emp?.nombre, datos, motivo, archivedAt);
    setCircHist(prev => [row, ...prev]);   // optimista
    sbInsertCircHist(row);
    return true;
  }, [liqParams, employees]);

  const handleFile = useCallback(async(e)=>{
    const file=e.target.files[0]; if(!file) return;
    setImporting(true); setImportMsg({text:"Procesando...",ok:true});
    try {
      const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:false});
      const wsName=wb.SheetNames.find(n=>n==="Anormal")||wb.SheetNames[0];
      const parsed=parseAnormalSheet(wb.Sheets[wsName]);
      setRecords(prev=>{const m={};for(const r of prev)m[r.id]=r;for(const r of parsed)m[r.id]=r;return Object.values(m);});
      // ── Detección de empleados nuevos y de números de reloj REASIGNADOS ──
      // Si un N° ya existe pero la planilla trae OTRO nombre, el reloj
      // recicló el número: se pisa la ficha con el empleado nuevo y se
      // limpia su circular (archivándola antes, si tenía datos).
      const norm = s => (s||"").toLowerCase().normalize("NFD")
        .replace(/[\u0300-\u036f]/g,"").replace(/\s+/g," ").trim();
      const byEmp = {};
      for (const r of parsed){ if(!byEmp[r.empNo]) byEmp[r.empNo]=[]; byEmp[r.empNo].push(r); }

      const nuevos = [], reasignados = [];
      const empsAct = employeesRef.current;
      for (const [noStr, recs] of Object.entries(byEmp)) {
        const no = Number(noStr), exist = empsAct[no];
        const d = detectSchedule(recs);
        const ficha = {empNo:no,nombre:recs[0].nombre,depto:recs[0].depto,
          entrada:d?.entrada||"06:00",salida:d?.salida||"16:30",
          activo:true,autoDetected:true,tipo:"operario"};
        if (!exist) nuevos.push(ficha);
        else if (norm(exist.nombre) !== norm(recs[0].nombre)) {
          reasignados.push({ viejo: exist.nombre, ficha });
        }
      }

      // Reasignados: archivar la circular vieja (si había) y borrarla
      for (const { ficha } of reasignados) {
        archivarCircular(ficha.empNo, "reasignacion");
        sbDelete("liq_params", ficha.empNo, "emp_no");
        setLiqParams(p=>{const u={...p}; delete u[String(ficha.empNo)]; return u;});
        delete prevLiqRef.current[String(ficha.empNo)];
      }

      // Altas y pisadas: al estado y a Supabase de a una
      const altas = [...nuevos, ...reasignados.map(r=>r.ficha)];
      if (altas.length) {
        setEmployees(prev=>{const u={...prev}; for(const e of altas) u[e.empNo]=e; return u;});
        for (const e of altas) sbUpsertSingle("empleados", empToRow(e), "emp_no");
      }

      const notas = reasignados.map(r=>`N° ${r.ficha.empNo}: ${cap(r.viejo)} → ${cap(r.ficha.nombre)}`);
      setImportMsg({
        text:`${parsed.length} registros importados desde "${wsName}"` +
          (notas.length ? ` — ⚠ Números reasignados (ficha y circular reiniciadas): ${notas.join("; ")}` : ""),
        ok:true
      });
      // Sync to Supabase silently
      const periodo = parsed[0]?.fecha?.slice(0,7);
      sbUpsert("registros", parsed.map(r=>recToRow(r,periodo)));
    } catch(err){setImportMsg({text:`Error: ${err.message}`,ok:false});}
    setImporting(false);
    if(fileRef.current)fileRef.current.value="";
  },[archivarCircular]);

  const startEdit=emp=>{setEditingEmp(emp.empNo);setEditDraft({nombre:emp.nombre,depto:emp.depto,entrada:emp.entrada,salida:emp.salida,tipo:emp.tipo||"operario"});};
  // Elimina definitivamente un empleado inactivo y todo su rastro,
  // liberando el número de reloj para un futuro empleado.
  const eliminarEmpleado = useCallback((emp) => {
    const nRecs = [...records, ...manualRecords].filter(r=>r.empNo===emp.empNo).length;
    const ok = window.confirm(
      `¿Eliminar definitivamente a ${cap(emp.nombre)} (N° ${emp.empNo})?\n\n` +
      `Se borrarán: su ficha, ${nRecs} registros de asistencia, sus correcciones ` +
      `de horario, su circular y el histórico de circulares.\n\n` +
      `Esta acción NO se puede deshacer.`
    );
    if (!ok) return;

    // 1) Registros de asistencia (todos, incluidos manuales)
    sbDelete("registros", emp.empNo, "emp_no");
    setRecords(p=>p.filter(r=>r.empNo!==emp.empNo));
    setManualRecords(p=>p.filter(r=>r.empNo!==emp.empNo));

    // 2) Correcciones de hora (sus claves empiezan con "empNo_")
    const pref = `${emp.empNo}_`;
    const keys = Object.keys(manualSalidas).filter(k=>k.startsWith(pref));
    const idsCorr = [...new Set(keys.map(k=>k.endsWith("_ent") ? k.slice(0,-4) : k))];
    for (const id of idsCorr) sbDelete("correcciones", id);
    if (keys.length) setManualSalidas(p=>{const u={...p}; for(const k of keys) delete u[k]; return u;});

    // 3) Circular + histórico de circulares
    sbDelete("liq_params", emp.empNo, "emp_no");
    setLiqParams(p=>{const u={...p}; delete u[String(emp.empNo)]; return u;});
    delete prevLiqRef.current[String(emp.empNo)];
    sbDelete("circular_historico", emp.empNo, "emp_no");
    setCircHist(p=>p.filter(h=>h.emp_no!==emp.empNo));

    // 4) La ficha
    sbDelete("empleados", emp.empNo, "emp_no");
    setEmployees(p=>{const u={...p}; delete u[emp.empNo]; return u;});
  }, [records, manualRecords, manualSalidas]);
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
    sbUpsert("empleados", Object.values(updated).map(empToRow), "emp_no");
  };

  const toggleTipo = (empNo) => {
    setEmployees(p => {
      const nuevoTipo = (p[empNo].tipo||"operario")==="operario" ? "administrativo" : "operario";
      const updated = { ...p[empNo], tipo: nuevoTipo };
      sbUpsertSingle("empleados", empToRow(updated), "emp_no");
      return { ...p, [empNo]: updated };
    });
  };


  // Inicia un nuevo período de liquidación para TODOS los empleados:
  // - setea periodo (ej "JUNIO 2026") y desde/hasta
  // - resetea overrides del mes anterior (llegadas tarde, retiros, hs extra manual,
  //   SAC, vacaciones, feriados, findeSel)
  // - CONSERVA adelantos y todos los datos de la circular (sueldo, valores, etc.)
  const iniciarPeriodo = useCallback((nombrePeriodo, desde, hasta) => {
    // Campos a borrar de cada empleado al arrancar el período
    const RESET = [
      "descDemorasManual", "descSalTempManual",   // overrides de descuentos
      "horasExtraManualHs", "horasExtraManualImp", // hs extra fuera del reloj
      "hsExtraRelojManual",                        // corrección de hs del reloj
      "impExtrasManual", "impFindeManual",         // importes manuales de
      "impFeriadosManual", "impVacacionesManual",  // adicionales
      "sac", "vacaciones", "feriados",             // adicionales manuales
      "findeSel",                                  // selección de findes
    ];
    setLiqParams(prev => {
      const next = {};
      for (const [key, datos] of Object.entries(prev)) {
        const limpio = { ...datos };
        for (const f of RESET) delete limpio[f];
        // setear el período nuevo (adelantos y circular quedan intactos)
        limpio.periodo = nombrePeriodo;
        limpio.desde   = desde;
        limpio.hasta   = hasta;
        next[key] = limpio;
      }
      return next;
    });
  }, []);

  const empList=Object.values(employees).sort((a,b)=>a.empNo-b.empNo);
  const filteredEmps=empList.filter(e=>{
    const matchSearch  = !empF||e.nombre.toLowerCase().includes(empF.toLowerCase())||String(e.empNo).includes(empF);
    const matchTipo    = empTipoF==="todos"||(e.tipo||"operario")===empTipoF;
    const matchActivo  = empActivoF==="todos"||(empActivoF==="activo"?e.activo!==false:e.activo===false);
    return matchSearch && matchTipo && matchActivo;
  });
  const allRecs = [...records, ...manualRecords];
  const filteredRecs=allRecs
    .filter(r=>(!recF.emp||r.nombre.toLowerCase().includes(recF.emp.toLowerCase())||String(r.empNo).includes(recF.emp))&&(!recF.desde||r.fecha>=recF.desde)&&(!recF.hasta||r.fecha<=recF.hasta))
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
        <div style={S.logo}><img src="/PyG-logo.png" alt="PYG S.R.L." style={{height:36,width:"auto",objectFit:"contain",display:"block"}}/></div>
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
              {sbWriteError ? `No se pudo guardar (${sbWriteError.table})` : "Sin conexión a Supabase"}
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

      {/* BANNER DE ERROR DE GUARDADO — visible para que RRHH no pierda datos sin enterarse */}
      {sbWriteError && (
        <div style={{background:"#fff5f5",borderBottom:"1px solid #fed7d7",padding:"10px 28px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
          <span style={{fontSize:13,color:"#c53030",fontWeight:600,display:"flex",alignItems:"center",gap:7}}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#c53030" strokeWidth="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
            Atención: el último cambio no se pudo guardar en la base de datos.
          </span>
          <span style={{fontSize:12,color:COL.textSub}}>
            Tabla <code style={{fontFamily:MONO,background:"#fff",padding:"1px 6px",borderRadius:4,color:"#c53030"}}>{sbWriteError.table}</code>. Quedó guardado solo en este equipo. Revisá la conexión y volvé a editarlo, o descartá este aviso.
          </span>
          <button onClick={()=>{ setSbWriteError(null); setSbStatus(sbLastSync?"synced":""); }}
            style={{marginLeft:"auto",background:"#fff",border:"1px solid #fed7d7",borderRadius:6,padding:"5px 12px",fontSize:12,color:"#c53030",cursor:"pointer",fontFamily:SANS,fontWeight:600}}>
            Descartar
          </button>
        </div>
      )}

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
                <label style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:COL.textSub}}>
                  Desde <input type="date" value={recF.desde} onChange={e=>setRecF(p=>({...p,desde:e.target.value}))} style={S.dateInput}/>
                </label>
                <label style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:COL.textSub}}>
                  Hasta <input type="date" value={recF.hasta} onChange={e=>setRecF(p=>({...p,hasta:e.target.value}))} style={S.dateInput}/>
                </label>
                {(recF.desde||recF.hasta)&&<button onClick={()=>setRecF(p=>({...p,desde:"",hasta:""}))} style={S.cancelBtn}>Limpiar fechas</button>}
                <span style={{color:COL.textFaint,fontSize:12}}>{filteredRecs.length} registros</span>
              </div>
              <button onClick={()=>{setAddingRec(true);setNewRec({empNo:"",fecha:recF.desde||"",entrada:"",salida:""});}}
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
                  <div style={{fontSize:11,color:COL.textFaint,marginBottom:4,fontWeight:500}}>Tipo</div>
                  <select value={newRec.ausencia||""} onChange={e=>setNewRec(p=>({...p,ausencia:e.target.value}))}
                    style={{...S.sInput,width:160,padding:"7px 10px"}}>
                    <option value="">— Normal —</option>
                    <option value="aus_just">Ausencia justificada</option>
                    <option value="aus_injust">Ausencia injustificada</option>
                  </select>
                </div>
                {!newRec.ausencia&&<>
                  <div>
                    <div style={{fontSize:11,color:COL.textFaint,marginBottom:4,fontWeight:500}}>Entrada</div>
                    <input type="time" value={newRec.entrada} onChange={e=>setNewRec(p=>({...p,entrada:e.target.value}))} style={S.tInp}/>
                  </div>
                  <div>
                    <div style={{fontSize:11,color:COL.textFaint,marginBottom:4,fontWeight:500}}>Salida</div>
                    <input type="time" value={newRec.salida} onChange={e=>setNewRec(p=>({...p,salida:e.target.value}))} style={S.tInp}/>
                  </div>
                </>}
                <div>
                  <div style={{fontSize:11,color:COL.textFaint,marginBottom:4,fontWeight:500}}>Observación</div>
                  <input value={newRec.observacion||""} onChange={e=>setNewRec(p=>({...p,observacion:e.target.value}))}
                    placeholder="Opcional..." style={{...S.sInput,width:200,padding:"5px 8px",fontSize:12}}/>
                </div>
                <div style={{display:"flex",gap:8,flexDirection:"column"}}>
                  {(!newRec.empNo||!newRec.fecha)&&<span style={{fontSize:11,color:"#c53030"}}>* Empleado y fecha requeridos</span>}
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>{
                      if(!newRec.empNo||!newRec.fecha) return;
                      const emp=employees[Number(newRec.empNo)];
                      if(!emp) return;
                      const sufijo = newRec.ausencia ? `_${newRec.ausencia}` : "_manual";
                      const id=`${newRec.empNo}_${newRec.fecha}${sufijo}`;
                      const manRec={
                        id, empNo:Number(newRec.empNo), nombre:emp.nombre, depto:emp.depto,
                        fecha:newRec.fecha,
                        entrada: newRec.ausencia ? null : (newRec.entrada||null),
                        salida:  newRec.ausencia ? null : (newRec.salida||null),
                        soloEntrada: false, manual:true,
                        ausencia: newRec.ausencia||null,
                        observacion:newRec.observacion||null
                      };
                      setManualRecords(p=>[...p.filter(r=>r.id!==id),manRec]);
                      setAddingRec(false);
                      sbUpsertSingle("registros", recToRow(manRec));
                    }} style={S.saveBtn}>Guardar</button>
                    <button onClick={()=>setAddingRec(false)} style={S.cancelBtn}>Cancelar</button>
                  </div>
                </div>
              </div>
            )}

            <div style={S.tblWrap}>
              <table style={S.table}>
                <THead cols={["N°","Nombre","Tipo","Fecha","Entrada","Salida","En jornada","Hs. extra","Demora","Sal. temprana","Obs.",""]}/>
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
                      <tr key={r.id} style={{background:r.ausencia==="aus_injust"?"#fff7ed":r.ausencia==="aus_just"?"#f5f3ff":r.manual?"#fefaf2":i%2===0?"#fff":"#fafbfc"}}>
                        <td style={{...S.td,fontFamily:MONO,fontSize:12,color:COL.accent}}>{r.empNo}</td>
                        <td style={{...S.td,textAlign:"left",fontWeight:500,color:COL.text}}>
                          {cap(r.nombre)}
                          {r.manual&&!r.ausencia&&<span style={{marginLeft:7,fontSize:10,background:"#fde68a",color:"#92400e",borderRadius:4,padding:"1px 6px",fontWeight:600}}>manual</span>}
                          {r.ausencia==="aus_just"&&<span style={{marginLeft:7,fontSize:10,background:"#ede9fe",color:"#7c3aed",borderRadius:4,padding:"1px 6px",fontWeight:600}}>AUS. JUST.</span>}
                          {r.ausencia==="aus_injust"&&<span style={{marginLeft:7,fontSize:10,background:"#ffedd5",color:"#dc6b19",borderRadius:4,padding:"1px 6px",fontWeight:600}}>AUS. INJUST.</span>}
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
                        <td style={{...S.td,fontSize:11,padding:"4px 8px",maxWidth:140}}>
                          {editingCell?.id===r.id&&editingCell?.field==="observacion"
                            ? <input autoFocus
                                defaultValue={r.observacion||""}
                                onBlur={e=>{
                                  const v=e.target.value.trim();
                                  // Actualizar en records o manualRecords
                                  if(r.manual){
                                    setManualRecords(p=>p.map(x=>x.id===r.id?{...x,observacion:v||null}:x));
                                  } else {
                                    setRecords(p=>p.map(x=>x.id===r.id?{...x,observacion:v||null}:x));
                                  }
                                  sbUpdate("registros",r.id,{observacion:v||null},"id");
                                  setEditingCell(null);
                                }}
                                onKeyDown={e=>{if(e.key==="Escape")setEditingCell(null);if(e.key==="Enter")e.target.blur();}}
                                style={{...S.inlineInput,fontSize:11,padding:"3px 6px",width:"100%"}}
                              />
                            : <span
                                onClick={()=>setEditingCell({id:r.id,field:"observacion"})}
                                title="Click para editar observación"
                                style={{color:r.observacion?COL.textSub:"#d1d5db",cursor:"pointer",
                                  borderBottom:"1px dashed #d1d5db",display:"block",
                                  overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:130}}>
                                {r.observacion||<span style={{fontSize:10,fontStyle:"italic"}}>agregar obs.</span>}
                              </span>
                          }
                        </td>
                        <td style={{...S.td,padding:"5px 8px",width:28}}>
                          <button onClick={()=>{
                            if(r.manual){setManualRecords(p=>p.filter(x=>x.id!==r.id));}
                            else{setRecords(p=>p.filter(x=>x.id!==r.id));}
                            sbDelete("registros",r.id);
                          }}
                            title="Eliminar registro"
                            style={{background:"none",border:"none",cursor:"pointer",color:"#fca5a5",fontSize:15,lineHeight:1,padding:2}}>×</button>
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
            <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:10}}>
              {[["todos","Todos"],["operario","Operarios"],["administrativo","Administrativos"]].map(([val,label])=>(
                <button key={val} onClick={()=>setEmpTipoF(val)}
                  style={{padding:"5px 14px",borderRadius:20,border:`1px solid ${empTipoF===val?COL.accent:COL.border2}`,background:empTipoF===val?COL.accentBg:"#fff",color:empTipoF===val?COL.accent:COL.textSub,fontFamily:SANS,fontSize:12,cursor:"pointer",fontWeight:empTipoF===val?600:400}}>
                  {label} <span style={{fontFamily:MONO,fontSize:11,color:empTipoF===val?COL.accent:COL.textFaint,marginLeft:4}}>
                    {val==="todos"?empList.length:empList.filter(e=>(e.tipo||"operario")===val).length}
                  </span>
                </button>
              ))}
            </div>

            {/* Activo/Inactivo filter */}
            <div style={{display:"flex",gap:6,marginBottom:14,alignItems:"center"}}>
              <span style={{fontSize:11,color:COL.textFaint,fontWeight:500,letterSpacing:"0.05em",textTransform:"uppercase",marginRight:2}}>Estado:</span>
              {[["todos","Todos"],["activo","Activos"],["inactivo","Inactivos"]].map(([val,label])=>{
                const count = val==="todos" ? empList.length : val==="activo" ? empList.filter(e=>e.activo!==false).length : empList.filter(e=>e.activo===false).length;
                const isOn  = empActivoF===val;
                const colOn = val==="activo" ? "#276749" : val==="inactivo" ? "#9aa5b4" : COL.accent;
                const bgOn  = val==="activo" ? "#f0faf4"  : val==="inactivo" ? "#f5f5f5"  : COL.accentBg;
                const bdOn  = val==="activo" ? "#c3e6cb"  : val==="inactivo" ? "#e0e0e0"  : COL.accent;
                return (
                  <button key={val} onClick={()=>setEmpActivoF(val)}
                    style={{padding:"5px 14px",borderRadius:20,border:`1px solid ${isOn?bdOn:COL.border2}`,background:isOn?bgOn:"#fff",color:isOn?colOn:COL.textSub,fontFamily:SANS,fontSize:12,cursor:"pointer",fontWeight:isOn?600:400}}>
                    {label} <span style={{fontFamily:MONO,fontSize:11,marginLeft:4,color:isOn?colOn:COL.textFaint}}>{count}</span>
                  </button>
                );
              })}
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
                            :<span style={{display:"flex",gap:6,justifyContent:"center"}}>
                                <button onClick={()=>startEdit(emp)} className="hov-edit" style={S.editBtn}>Editar</button>
                                {!emp.activo && (
                                  <button onClick={()=>eliminarEmpleado(emp)}
                                    title="Eliminar definitivamente (libera el N° de reloj)"
                                    style={{...S.cancelBtn,color:"#c53030",borderColor:"#f5c6c6",background:"#fff5f5"}}>
                                    Eliminar
                                  </button>
                                )}
                              </span>}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ── RESUMEN (oculto) ── */}
        {tab===-1&&(
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
        {tab===3&&(()=>{
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
        {tab===4&&(()=>{
          const allDates = [...new Set(records.map(r=>r.fecha))].sort();

          const toggleDay = (fecha, tipo) => {
            setSpecialDays(prev => {
              const cur = prev[fecha];
              if (cur?.tipo === tipo) {
                const u = {...prev}; delete u[fecha];
                sbDelete("dias_especiales", fecha, "fecha");
                return u;
              }
              const salida = tipo === "feriado" ? "14:00" : undefined;
              sbUpsertSingle("dias_especiales", {fecha, tipo, salida: salida||null}, "fecha");
              return {...prev, [fecha]: {tipo, salida}};
            });
          };

          const setSalidaFeriado = (fecha, salida) => {
            setSpecialDays(prev => {
              if (prev[fecha]?.tipo !== "feriado") return prev;
              sbUpsertSingle("dias_especiales", {fecha, tipo:"feriado", salida}, "fecha");
              return {...prev, [fecha]: {tipo:"feriado", salida}};
            });
          };

          // Generar TODOS los días de cada mes que aparece en registros
          // así se pueden marcar ausencias en días sin registro del reloj
          const mesesConDatos = [...new Set(allDates.map(f=>f.slice(0,7)))].sort();
          const byMonth = {};
          for (const month of mesesConDatos) {
            const [y, m] = month.split("-").map(Number);
            const diasEnMes = new Date(y, m, 0).getDate(); // último día del mes
            byMonth[month] = [];
            for (let d = 1; d <= diasEnMes; d++) {
              byMonth[month].push(`${month}-${String(d).padStart(2,"0")}`);
            }
          }

          const DIAS  = ["Dom","Lun","Mar","Mié","Jue","Vie","Sáb"];
          const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
          const feriadosCount = Object.values(specialDays).filter(v=>v.tipo==="feriado").length;

          return (
            <div style={{maxWidth:900}}>
              <H2>Calendario — Días especiales</H2>
              <p style={S.body}>
                Marcá los días como <strong style={{color:"#b45309"}}>Feriado</strong> (operarios trabajan hasta las 14:00, el horario de referencia cambia automáticamente) o como{" "}
                <strong style={{color:"#c53030"}}>Día libre</strong> (sin registros esperados),{" "}
                <strong style={{color:"#7c3aed"}}>Ausencia justificada</strong> o{" "}
                <strong style={{color:"#dc6b19"}}>Ausencia injustificada</strong>.
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
                              <div style={{fontSize:10,color:"#b45309",marginBottom:8}}>hasta {sp?.salida || "14:00"}</div>
                            )}
                            {isLibre && (
                              <div style={{fontSize:10,background:"#fecaca",color:"#c53030",borderRadius:4,padding:"2px 7px",marginBottom:4,fontWeight:700,display:"inline-block"}}>
                                DÍA LIBRE
                              </div>
                            )}
                            {sp?.tipo==="aus_just" && (
                              <div style={{fontSize:10,background:"#ede9fe",color:"#7c3aed",borderRadius:4,padding:"2px 7px",marginBottom:4,fontWeight:700,display:"inline-block"}}>
                                AUS. JUST.
                              </div>
                            )}
                            {sp?.tipo==="aus_injust" && (
                              <div style={{fontSize:10,background:"#ffedd5",color:"#dc6b19",borderRadius:4,padding:"2px 7px",marginBottom:4,fontWeight:700,display:"inline-block"}}>
                                AUS. INJUST.
                              </div>
                            )}
                            <div style={{display:"flex",flexDirection:"column",gap:6}}>
                              <button onClick={()=>toggleDay(fecha,"feriado")}
                                style={{fontSize:10,padding:"3px 7px",border:`1px solid ${isFeriado?"#f6d860":COL.border}`,borderRadius:5,cursor:"pointer",background:isFeriado?"#fde68a":"#fff",color:isFeriado?"#92400e":COL.textSub,fontFamily:SANS,fontWeight:isFeriado?700:400,textAlign:"left"}}>
                                {isFeriado ? "✓ Feriado" : "+ Feriado"}
                              </button>
                              {isFeriado && (
                                <label style={{display:"flex",alignItems:"center",gap:5,fontSize:10,color:"#92400e"}}>
                                  <span style={{whiteSpace:"nowrap"}}>Hasta</span>
                                  <input type="time"
                                    value={sp?.salida || "14:00"}
                                    onChange={e=>setSalidaFeriado(fecha, e.target.value)}
                                    style={{border:"1px solid #f6d860",borderRadius:5,padding:"2px 4px",fontFamily:MONO,fontSize:11,color:"#92400e",background:"#fffdf5",outline:"none",width:"100%"}}/>
                                </label>
                              )}
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
        {tab===5&&(()=>{
          const activeEmps = empList.filter(e=>e.activo!==false);
          const filteredCircular = activeEmps.filter(e=>
            !circularF ||
            e.nombre.toLowerCase().includes(circularF.toLowerCase()) ||
            String(e.empNo).includes(circularF)
          );

          const getP = (empNo) => liqParams[String(empNo)] || {};

          // Deriva valor día, hora, feriado y sábado a partir del sueldo bruto.
          //   día      = sueldo / 22
          //   hora     = día / 8
          //   feriado  = sueldo / 22  (igual que el día de falta)
          //   sábado   = sueldo / 22  (día completo)
          //   hora ext = hora * 1.5
          const derivar = (basico) => {
            const b = parseFloat(basico) || 0;
            if (!b) return { valorDia:"", valorHora:"", valorDiaFinde:"", valorHoraExt:"" };
            const valorDia  = b / 22;
            const valorHora = valorDia / 7;
            return {
              valorDia:      valorDia.toFixed(2),
              valorHora:     valorHora.toFixed(2),
              valorDiaFinde: valorDia.toFixed(2),         // sábado = sueldo/22
              valorHoraExt:  (valorHora * 1.5).toFixed(2), // hora extra = hora*1.5
            };
          };

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
                <div style={{display:"flex",gap:8,alignItems:"center"}}>
                  {circHist.length>0 && (
                    <button onClick={()=>setVerPlanillaHist(v=>!v)}
                      style={{...S.editBtn,display:"flex",alignItems:"center",gap:6,
                        background: verPlanillaHist?COL.accentBg:"transparent",
                        color: verPlanillaHist?COL.accent:COL.textSub,
                        borderColor: verPlanillaHist?COL.accent:COL.border}}>
                      🕓 Ver histórico ({circHist.length})
                    </button>
                  )}
                  <input placeholder="Buscar empleado…" value={circularF}
                    onChange={e=>setCircularF(e.target.value)} style={S.sInput}/>
                </div>
              </div>

              {/* ── Planilla global del histórico de circulares ── */}
              {verPlanillaHist && (()=>{
                const fmtNum = v => v ? Number(v).toLocaleString("es-AR",{minimumFractionDigits:2,maximumFractionDigits:2}) : "—";
                const fmtFechaIng = f => f ? new Date(f+"T12:00:00").toLocaleDateString("es-AR") : "—";
                // ordenar por fecha de archivado desc, y dentro por nº de empleado
                const filasHist = [...circHist]
                  .sort((a,b)=> b.archived_at.localeCompare(a.archived_at) || a.emp_no-b.emp_no);

                const thH = {padding:"9px 12px",color:COL.textFaint,fontSize:11,fontWeight:600,
                  textAlign:"right",whiteSpace:"nowrap",letterSpacing:"0.03em",
                  background:"#f7f8fa",borderBottom:`1px solid ${COL.border}`};
                const tdH = (color=COL.textSub) => ({padding:"8px 12px",color,textAlign:"right",
                  borderBottom:`1px solid ${COL.border}`,whiteSpace:"nowrap",fontFamily:MONO,fontSize:12});

                const exportarHist = ()=>{
                  const headers = [
                    ["empNo","N°"],["nombre","Empleado"],["fechaArch","Fecha archivado"],["motivo","Motivo"],
                    ["sueldo","Sueldo bruto"],["valorDia","Valor día"],["valorHora","Valor hora"],
                    ["valorSab","Sábado"],["valorHExt","Hora extra"],["area","Área"],
                    ["ingreso","Ingreso"],["obs","Observaciones"],
                  ];
                  const rows = filasHist.map(h=>{
                    const dd = h.datos||{};
                    const emp = employees[h.emp_no];
                    return {
                      empNo: h.emp_no,
                      nombre: dd.nombreDisplay || (emp?cap(emp.nombre):h.nombre||""),
                      fechaArch: new Date(h.archived_at).toLocaleString("es-AR"),
                      motivo: h.motivo==="cambio_sueldo"?"Por aumento":"Manual",
                      sueldo: dd.sueldoBasico?Math.round(Number(dd.sueldoBasico)):"",
                      valorDia: dd.valorDia?Number(Number(dd.valorDia).toFixed(2)):"",
                      valorHora: dd.valorHora?Number(Number(dd.valorHora).toFixed(2)):"",
                      valorSab: dd.valorDiaFinde?Number(Number(dd.valorDiaFinde).toFixed(2)):"",
                      valorHExt: dd.valorHoraExt?Number(Number(dd.valorHoraExt).toFixed(2)):"",
                      area: dd.area||"",
                      ingreso: dd.ingreso?new Date(dd.ingreso+"T12:00:00").toLocaleDateString("es-AR"):"",
                      obs: dd.observaciones||"",
                    };
                  });
                  exportXLSX(rows, headers, "Histórico circulares", "historico_circulares.xlsx");
                };

                return (
                  <div style={{marginBottom:22,background:COL.surface,border:`1px solid ${COL.border}`,borderRadius:12,overflow:"hidden"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 18px",background:"#f7f8fa",borderBottom:`2px solid ${COL.accent}`,flexWrap:"wrap",gap:10}}>
                      <span style={{fontSize:14,fontWeight:700,color:COL.text,display:"flex",alignItems:"center",gap:8}}>
                        🕓 Histórico de circulares
                        <span style={{fontFamily:MONO,fontSize:12,fontWeight:400,color:COL.textFaint}}>· {filasHist.length} versiones archivadas</span>
                      </span>
                      <div style={{display:"flex",gap:8}}>
                        <button onClick={exportarHist} style={{...S.btnS,display:"flex",alignItems:"center",gap:6}}>
                          <span style={{fontSize:14,lineHeight:1}}>↓</span> Exportar Excel
                        </button>
                        <button onClick={()=>setVerPlanillaHist(false)} style={S.cancelBtn}>Cerrar</button>
                      </div>
                    </div>
                    <div style={{overflowX:"auto"}}>
                      <table style={{...S.table,fontSize:12}}>
                        <thead>
                          <tr>
                            <th style={{...S.th,textAlign:"left"}}>Empleado</th>
                            <th style={{...S.th,textAlign:"left"}}>Archivado</th>
                            <th style={{...S.th,textAlign:"left"}}>Motivo</th>
                            <th style={thH}>Sueldo bruto</th>
                            <th style={thH}>Día</th>
                            <th style={thH}>Hora</th>
                            <th style={thH}>Sábado</th>
                            <th style={thH}>Hora extra</th>
                            <th style={{...S.th,textAlign:"left"}}>Área</th>
                            <th style={{...S.th,textAlign:"left"}}>Ingreso</th>
                            <th style={{...S.th,textAlign:"left"}}>Observaciones</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filasHist.map((h,i)=>{
                            const dd = h.datos||{};
                            const emp = employees[h.emp_no];
                            const nombre = dd.nombreDisplay || (emp?cap(emp.nombre):h.nombre||"—");
                            const fechaArch = new Date(h.archived_at).toLocaleString("es-AR",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"});
                            return (
                              <tr key={h.id} style={{background:i%2===0?"#fff":"#fafbfc"}}>
                                <td style={{...S.td,textAlign:"left",fontWeight:600,color:COL.text}}>
                                  <span style={{fontFamily:MONO,fontSize:11,color:COL.textFaint,marginRight:8}}>{h.emp_no}</span>
                                  {nombre}
                                </td>
                                <td style={{...S.td,textAlign:"left",fontSize:12,color:COL.textSub,fontFamily:MONO}}>{fechaArch}</td>
                                <td style={{...S.td,textAlign:"left"}}>
                                  <span style={{fontSize:10,padding:"1px 8px",borderRadius:20,fontWeight:600,
                                    background: h.motivo==="cambio_sueldo"?"#fef3e2":COL.accentBg,
                                    color: h.motivo==="cambio_sueldo"?"#b45309":COL.accent}}>
                                    {h.motivo==="cambio_sueldo"?"por aumento":"manual"}
                                  </span>
                                </td>
                                <td style={{...tdH(COL.text),fontWeight:700}}>{dd.sueldoBasico?`$ ${fmtNum(dd.sueldoBasico)}`:"—"}</td>
                                <td style={tdH()}>{fmtNum(dd.valorDia)}</td>
                                <td style={tdH()}>{fmtNum(dd.valorHora)}</td>
                                <td style={tdH()}>{fmtNum(dd.valorDiaFinde)}</td>
                                <td style={tdH("#276749")}>{fmtNum(dd.valorHoraExt)}</td>
                                <td style={{...S.td,textAlign:"left",fontSize:12,color:dd.area?COL.textSub:"#c0c8d2"}}>{dd.area||"—"}</td>
                                <td style={{...S.td,textAlign:"left",fontSize:12,color:dd.ingreso?COL.textSub:"#c0c8d2"}}>{fmtFechaIng(dd.ingreso)}</td>
                                <td style={{...S.td,textAlign:"left",fontSize:11,color:dd.observaciones?COL.textSub:"#c0c8d2",maxWidth:220,whiteSpace:"normal"}}>
                                  {dd.observaciones||"—"}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                );
              })()}

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
                          {!isEd && p.sueldoBasico && (
                            <button onClick={()=>{
                              if(confirm(`¿Archivar la circular actual de ${cap(emp.nombre)}? Quedará guardada como versión histórica con la fecha de hoy.`)){
                                archivarCircular(emp.empNo,"manual");
                              }
                            }}
                              title="Guardar una copia histórica de esta circular"
                              style={{...S.editBtn,display:"flex",alignItems:"center",gap:5}}>
                              🗄 Archivar
                            </button>
                          )}
                          {!isEd && circHist.some(h=>h.emp_no===emp.empNo) && (
                            <button onClick={()=>setHistEmp(histEmp===emp.empNo?null:emp.empNo)}
                              title="Ver versiones anteriores"
                              style={{...S.editBtn,
                                background: histEmp===emp.empNo?COL.accentBg:"transparent",
                                color: histEmp===emp.empNo?COL.accent:COL.textSub}}>
                              🕓 Histórico ({circHist.filter(h=>h.emp_no===emp.empNo).length})
                            </button>
                          )}
                          {isEd
                            ? <>
                                <button onClick={()=>{
                                  const anterior = liqParams[String(emp.empNo)] || {};
                                  const sueldoCambio =
                                    anterior.sueldoBasico &&
                                    String(anterior.sueldoBasico) !== String(circularDraft.sueldoBasico);
                                  if (sueldoCambio) archivarCircular(emp.empNo, "cambio_sueldo");
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
                                  area:          p.area          ||"",
                                  nombreDisplay: p.nombreDisplay ||"",
                                  ingreso:       p.ingreso       ||"",
                                });
                              }} className="hov-edit" style={S.editBtn}>✏ Editar</button>
                          }
                        </div>
                      </div>

                      <div style={{padding:"16px 20px"}}>

                        {/* Top info row */}
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16,marginBottom:12}}>
                          <div>
                            <div style={{fontSize:10,color:COL.textFaint,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Nombre en liquidación</div>
                            {isEd
                              ? <input value={circularDraft.nombreDisplay||""} onChange={e=>setCircularDraft(p=>({...p,nombreDisplay:e.target.value}))} placeholder={cap(emp.nombre)}
                                  style={{...S.sInput,width:"100%",padding:"6px 10px",fontSize:13}}/>
                              : <span style={{fontSize:13,fontWeight:600,color:p.nombreDisplay?COL.text:"#c0c8d2"}}>{p.nombreDisplay||cap(emp.nombre)}</span>
                            }
                          </div>
                          <div>
                            <div style={{fontSize:10,color:COL.textFaint,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Área</div>
                            {isEd
                              ? <input value={circularDraft.area||""} onChange={e=>setCircularDraft(p=>({...p,area:e.target.value}))} placeholder="ej: Producción"
                                  style={{...S.sInput,width:"100%",padding:"6px 10px",fontSize:13}}/>
                              : <span style={{fontSize:13,color:p.area?COL.text:"#c0c8d2"}}>{p.area||"—"}</span>
                            }
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
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16,marginBottom:18}}>
                          <div>
                            <div style={{fontSize:10,color:COL.textFaint,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Sueldo bruto</div>
                            {isEd
                              ? <input type="number" min="0" step="any" value={circularDraft.sueldoBasico||""}
                                  onChange={e=>{
                                    const v = e.target.value;
                                    setCircularDraft(p=>({...p, sueldoBasico:v, ...derivar(v)}));
                                  }}
                                  placeholder="0"
                                  style={{...S.sInput,width:"100%",padding:"6px 10px",fontFamily:MONO,fontSize:14,fontWeight:600}}/>
                              : <span style={{fontFamily:MONO,fontSize:15,fontWeight:700,color:p.sueldoBasico?COL.text:"#c0c8d2"}}>{p.sueldoBasico?fmt$(p.sueldoBasico):"—"}</span>
                            }
                          </div>
                          <div>
                            <div style={{fontSize:10,color:COL.textFaint,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Horario</div>
                            <span style={{fontSize:13,color:COL.textSub}}>Lunes a Sábados {horarioDisplay}</span>
                          </div>
                          <div>
                            <div style={{fontSize:10,color:COL.textFaint,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>Fecha ingreso</div>
                            {isEd
                              ? <input type="date" value={circularDraft.ingreso||""} onChange={e=>setCircularDraft(p=>({...p,ingreso:e.target.value}))}
                                  style={{...S.sInput,width:"100%",padding:"6px 10px",fontSize:13}}/>
                              : <span style={{fontSize:13,color:p.ingreso?COL.text:"#c0c8d2"}}>{p.ingreso?new Date(p.ingreso+"T12:00:00").toLocaleDateString("es-AR"):"—"}</span>
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
                                {label:"Valor a descontar por día de falta", field:"valorDia",      note:"sueldo / 22"},
                                {label:"Valor hora de referencia",            field:"valorHora",     note:"día / 7 · por hora"},
                                {label:"Feriados trabajados",                 field:"valorDia",      note:`sueldo / 22 · hasta las ${emp.salida||"16:30"} hs`},
                                {label:"Sábados",                             field:"valorDiaFinde", note:"sueldo / 22 · 08 a 14 hs"},
                                {label:"Horas extras",                        field:"valorHoraExt",  note:`hora × 1.5 · a partir de las ${emp.salida||"16:30"} hs`},
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

                        {/* Histórico de versiones */}
                        {histEmp===emp.empNo && (
                          <div style={{marginTop:18,paddingTop:16,borderTop:`2px solid ${COL.border}`}}>
                            <div style={{fontSize:11,fontWeight:700,color:COL.accent,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:10}}>
                              Versiones archivadas
                            </div>
                            <div style={{display:"flex",flexDirection:"column",gap:8}}>
                              {circHist.filter(h=>h.emp_no===emp.empNo)
                                .sort((a,b)=>b.archived_at.localeCompare(a.archived_at))
                                .map(h=>{
                                  const dd = h.datos||{};
                                  const fechaH = new Date(h.archived_at).toLocaleString("es-AR",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"});
                                  return (
                                    <div key={h.id} style={{background:"#fafbfc",border:`1px solid ${COL.border}`,borderRadius:8,padding:"10px 14px"}}>
                                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6,flexWrap:"wrap",gap:8}}>
                                        <span style={{fontSize:12,fontWeight:600,color:COL.text}}>
                                          {fechaH}
                                          <span style={{marginLeft:8,fontSize:10,padding:"1px 7px",borderRadius:20,
                                            background: h.motivo==="cambio_sueldo"?"#fef3e2":COL.accentBg,
                                            color: h.motivo==="cambio_sueldo"?"#b45309":COL.accent,fontWeight:600}}>
                                            {h.motivo==="cambio_sueldo"?"por aumento":"manual"}
                                          </span>
                                        </span>
                                        <span style={{fontFamily:MONO,fontSize:13,fontWeight:700,color:COL.text}}>
                                          {dd.sueldoBasico?`$ ${Number(dd.sueldoBasico).toLocaleString("es-AR",{minimumFractionDigits:2})}`:"—"}
                                        </span>
                                      </div>
                                      <div style={{display:"flex",gap:14,flexWrap:"wrap",fontSize:11,color:COL.textSub,fontFamily:MONO}}>
                                        <span>Día: {dd.valorDia?Number(dd.valorDia).toLocaleString("es-AR",{minimumFractionDigits:2}):"—"}</span>
                                        <span>Hora: {dd.valorHora?Number(dd.valorHora).toLocaleString("es-AR",{minimumFractionDigits:2}):"—"}</span>
                                        <span>Sáb: {dd.valorDiaFinde?Number(dd.valorDiaFinde).toLocaleString("es-AR",{minimumFractionDigits:2}):"—"}</span>
                                        <span>H.ext: {dd.valorHoraExt?Number(dd.valorHoraExt).toLocaleString("es-AR",{minimumFractionDigits:2}):"—"}</span>
                                      </div>
                                    </div>
                                  );
                                })}
                            </div>
                          </div>
                        )}

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


        {/* ── 6 RESUMEN LIQUIDACIONES ── */}
        {tab===6&&(()=>{
          const activeEmps = empList.filter(e=>e.activo!==false);
          const getP = (empNo) => liqParams[String(empNo)] || {};

          const fmt$ = n => n ? `$${Math.round(Number(n)).toLocaleString("es-AR")}` : "—";

          const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
          const mesDesde = resumenMes ? `${resumenMes}-01` : null;
          const mesHasta = resumenMes
            ? `${resumenMes}-${String(new Date(+resumenMes.slice(0,4), +resumenMes.slice(5,7), 0).getDate()).padStart(2,"0")}`
            : null;
          const mesLabel = resumenMes ? `${MESES[+resumenMes.slice(5,7)-1]} ${resumenMes.slice(0,4)}` : null;
          const mesesDisponibles = [...new Set([...records,...manualRecords].map(r=>r.fecha?.slice(0,7)).filter(Boolean))].sort().reverse();

          // Calcular totales de liquidación para cada empleado activo con datos
          const filas = activeEmps
            .map(emp => {
              const p = getP(emp.empNo);
              if (!p.sueldoBasico) return null;

              const sueldoBasico  = parseFloat(p.sueldoBasico  || 0);
              const valorHoraExt  = parseFloat(p.valorHoraExt  || 0);
              const valorDia      = parseFloat(p.valorDia      || 0);
              const valorHora     = parseFloat(p.valorHora     || 0);
              const valorDiaFinde = parseFloat(p.valorDiaFinde || 0);
              const adelanto      = (Array.isArray(p.adelantos)
                ? p.adelantos.reduce((s,a)=>s+(parseFloat(a.monto)||0),0)
                : parseFloat(p.adelanto || 0));
              const feriados      = parseFloat(p.feriados      || 0);
              const sac           = parseFloat(p.sac           || 0);
              const vacaciones    = parseFloat(p.vacaciones    || 0);

              const desde = mesDesde || p.desde || "";
              const hasta = mesHasta || p.hasta || "";

              const empCalcs = (empSummary.find(s=>s.emp.empNo===emp.empNo)?.calcs || [])
                .filter(r => (!desde||r.fecha>=desde) && (!hasta||r.fecha<=hasta));

              const totalExtraMin   = empCalcs.reduce((s,r)=>s+(r.extra||0),0);
              const hsRelojOverride = p.hsExtraRelojManual !== undefined && p.hsExtraRelojManual !== ""
                ? parseFloat(p.hsExtraRelojManual) : null;
              const horasExtra      = (hsRelojOverride !== null ? Math.round(hsRelojOverride*60) : totalExtraMin) / 60;
              const fraccionesDem   = fraccionesDemoraCalc(empCalcs);
              const fraccionesSt    = fraccionesSalTempCalc(empCalcs);

              const allFindeInRange = empCalcs.filter(r=>{const d=new Date(r.fecha+"T12:00:00").getDay();return d===0||d===6;});
              const findeSel        = new Set(p.findeSel || allFindeInRange.map(r=>r.fecha));
              const diasFinde       = findeSel.size;

              const horasExtraManualHs  = parseFloat(p.horasExtraManualHs  || 0);
              const horasExtraManualImp = parseFloat(p.horasExtraManualImp || 0);
              const importeExtraManual  = horasExtraManualImp > 0
                ? horasExtraManualImp
                : horasExtraManualHs * valorHoraExt;

              const ovr = (v) => v !== undefined && v !== "" ? parseFloat(v) : null;
              const importeExtras    = (ovr(p.impExtrasManual)     ?? valorHoraExt * horasExtra) + importeExtraManual;
              const importeFeriados  =  ovr(p.impFeriadosManual)   ?? valorDia      * feriados;
              const importeVacacion  =  ovr(p.impVacacionesManual) ?? valorDia      * vacaciones;
              const importeFinde     =  ovr(p.impFindeManual)      ?? valorDiaFinde * diasFinde;
              const totalAdicionales = importeExtras + importeFeriados + sac + importeVacacion + importeFinde;

              const descDemorasCalc = (valorHora / 4) * fraccionesDem;
              const descSalTempCalc = (valorHora / 4) * fraccionesSt;
              const descDemorasManual = p.descDemorasManual !== undefined && p.descDemorasManual !== ""
                ? parseFloat(p.descDemorasManual) : null;
              const descSalTempManual = p.descSalTempManual !== undefined && p.descSalTempManual !== ""
                ? parseFloat(p.descSalTempManual) : null;
              const descDemoras = descDemorasManual !== null ? descDemorasManual : descDemorasCalc;
              const descSalTemp = descSalTempManual !== null ? descSalTempManual : descSalTempCalc;
              const totalDesc        = descDemoras + descSalTemp;
              const subtotal         = sueldoBasico + totalAdicionales - totalDesc - adelanto;

              return {
                emp,
                nombre:    p.nombreDisplay || cap(emp.nombre),
                area:      p.area || "—",
                ingreso:   p.ingreso || emp.ingreso || "",
                sueldoBasico,
                totalAdicionales,
                sac,
                totalDesc,
                adelanto,
                subtotal,
              };
            })
            .filter(Boolean);

          const totBasico     = filas.reduce((s,f)=>s+f.sueldoBasico,0);
          const totAdicional  = filas.reduce((s,f)=>s+f.totalAdicionales,0);
          const totSac        = filas.reduce((s,f)=>s+f.sac,0);
          const totDesc       = filas.reduce((s,f)=>s+f.totalDesc,0);
          const totAdelanto   = filas.reduce((s,f)=>s+f.adelanto,0);
          const totSubtotal   = filas.reduce((s,f)=>s+f.subtotal,0);

          const thR = {padding:"9px 14px",color:COL.textFaint,fontSize:11,fontWeight:600,
            textAlign:"right",whiteSpace:"nowrap",letterSpacing:"0.04em",
            background:"#f7f8fa",borderBottom:`1px solid ${COL.border}`};
          const tdR = (color=COL.textSub) => ({padding:"9px 14px",color,textAlign:"right",
            borderBottom:`1px solid ${COL.border}`,whiteSpace:"nowrap",fontFamily:MONO,fontSize:12});

          return (
            <div style={{maxWidth:1200}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:12,flexWrap:"wrap"}}>
                <div>
                  <H2>Resumen de liquidaciones</H2>
                  <p style={S.body}>Resumen calculado a partir de los datos de Circular y la asistencia del período seleccionado.</p>
                </div>
                {filas.length>0&&(
                  <button onClick={()=>{
                    const headers = [
                      ["empNo","N°"],["nombre","Empleado"],["area","Área"],["ingreso","Ingreso"],
                      ["sueldoBasico","Sueldo básico"],["totalAdicionales","Adicionales"],["sac","SAC"],
                      ["totalDesc","Desc. hs/días"],["adelanto","Adelantos"],["subtotal","Subtotal"],
                    ];
                    const rows = filas.map(f=>({
                      empNo:f.emp.empNo, nombre:f.nombre, area:f.area==="—"?"":f.area,
                      ingreso:f.ingreso?new Date(f.ingreso+"T12:00:00").toLocaleDateString("es-AR"):"",
                      sueldoBasico:Math.round(f.sueldoBasico),
                      totalAdicionales:Math.round(f.totalAdicionales),
                      sac:Math.round(f.sac),
                      totalDesc:Math.round(f.totalDesc),
                      adelanto:Math.round(f.adelanto),
                      subtotal:Math.round(f.subtotal),
                    }));
                    rows.push({empNo:"",nombre:`TOTAL (${filas.length})`,area:"",ingreso:"",
                      sueldoBasico:Math.round(totBasico),totalAdicionales:Math.round(totAdicional),sac:Math.round(totSac),
                      totalDesc:Math.round(totDesc),adelanto:Math.round(totAdelanto),subtotal:Math.round(totSubtotal)});
                    const sufijo = resumenMes ? `_${resumenMes}` : "";
                    exportXLSX(rows, headers, "Resumen", `resumen_liquidaciones${sufijo}.xlsx`);
                  }} style={{...S.btnS,display:"flex",alignItems:"center",gap:6}}>
                    <span style={{fontSize:14,lineHeight:1}}>↓</span> Exportar Excel
                  </button>
                )}
              </div>

              {/* Selector de mes */}
              <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap",marginBottom:18,background:COL.surface,padding:"12px 16px",borderRadius:10,border:`1px solid ${COL.border}`}}>
                <span style={{fontSize:12,color:COL.textSub,fontWeight:600,letterSpacing:"0.04em",textTransform:"uppercase"}}>Período</span>
                <input type="month" value={resumenMes} onChange={e=>setResumenMes(e.target.value)}
                  style={{...INPUT_STYLE,padding:"7px 11px",fontFamily:SANS,fontSize:13}}/>
                {mesLabel
                  ? <span style={{fontSize:13,color:COL.accent,fontWeight:600}}>{mesLabel}</span>
                  : <span style={{fontSize:12,color:"#b45309"}}>Sin filtro — usa el rango individual de cada liquidación (puede mezclar meses)</span>}
                {resumenMes && <button onClick={()=>setResumenMes("")} style={{...S.cancelBtn,marginLeft:4}}>Quitar filtro</button>}
                {mesesDisponibles.length>0 && (
                  <div style={{display:"flex",gap:6,marginLeft:"auto",flexWrap:"wrap"}}>
                    {mesesDisponibles.slice(0,6).map(m=>{
                      const lbl = `${MESES[+m.slice(5,7)-1].slice(0,3)} ${m.slice(2,4)}`;
                      const on = resumenMes===m;
                      return (
                        <button key={m} onClick={()=>setResumenMes(m)}
                          style={{fontSize:11,padding:"4px 10px",borderRadius:20,cursor:"pointer",fontFamily:SANS,
                            border:`1px solid ${on?COL.accent:COL.border2}`,background:on?COL.accentBg:"#fff",
                            color:on?COL.accent:COL.textSub,fontWeight:on?600:400}}>
                          {lbl}
                        </button>
                      );
                    })}
                  </div>
                )}
              </div>

              {filas.length === 0 && (
                <div style={{padding:"60px 20px",textAlign:"center",color:COL.textFaint,fontSize:13,
                  background:COL.surface,borderRadius:12,border:`1px solid ${COL.border}`}}>
                  No hay empleados con datos en Circular. Completá los valores en la pestaña Circular primero.
                </div>
              )}

              {filas.length > 0 && (
                <div style={S.tblWrap}>
                  <table style={{...S.table,fontSize:13}}>
                    <thead>
                      <tr>
                        <th style={{...S.th,textAlign:"left",minWidth:180}}>Empleado</th>
                        <th style={{...S.th,textAlign:"left"}}>Área</th>
                        <th style={{...S.th,textAlign:"left"}}>Ingreso</th>
                        <th style={thR}>Sueldo básico</th>
                        <th style={thR}>Adicionales</th>
                        <th style={thR}>SAC</th>
                        <th style={{...thR,color:"#c53030"}}>Desc. hs/días</th>
                        <th style={{...thR,color:"#b45309"}}>Adelantos</th>
                        <th style={{...thR,color:"#276749",fontWeight:700}}>Subtotal</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filas.map((f,i)=>(
                        <tr key={f.emp.empNo}
                          style={{background:i%2===0?"#fff":"#fafbfc"}}
                          onMouseEnter={e=>e.currentTarget.style.background="#f0f4fa"}
                          onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"#fff":"#fafbfc"}>
                          <td style={{...S.td,textAlign:"left",fontWeight:600,color:COL.text}}>
                            <span style={{fontFamily:MONO,fontSize:11,color:COL.textFaint,marginRight:8}}>{f.emp.empNo}</span>
                            {f.nombre}
                          </td>
                          <td style={{...S.td,textAlign:"left",fontSize:12,color:COL.textFaint}}>{f.area}</td>
                          <td style={{...S.td,textAlign:"left",fontSize:12,color:COL.textFaint}}>
                            {f.ingreso ? new Date(f.ingreso+"T12:00:00").toLocaleDateString("es-AR") : "—"}
                          </td>
                          <td style={tdR()}>{fmt$(f.sueldoBasico)}</td>
                          <td style={tdR("#276749")}>{f.totalAdicionales>0?fmt$(f.totalAdicionales):"—"}</td>
                          <td style={tdR()}>{f.sac>0?fmt$(f.sac):"—"}</td>
                          <td style={tdR("#c53030")}>{f.totalDesc>0?fmt$(f.totalDesc):"—"}</td>
                          <td style={tdR("#b45309")}>{f.adelanto>0?fmt$(f.adelanto):"—"}</td>
                          <td style={{...tdR("#276749"),fontWeight:700,fontSize:13}}>{fmt$(f.subtotal)}</td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr style={{background:"#f0f4fa",borderTop:`2px solid ${COL.border2}`}}>
                        <td colSpan={3} style={{padding:"10px 14px",fontWeight:700,fontSize:13,color:COL.text}}>
                          TOTAL — {filas.length} empleados{mesLabel?` · ${mesLabel}`:""}
                        </td>
                        <td style={{...tdR(),fontWeight:700,fontSize:13}}>{fmt$(totBasico)}</td>
                        <td style={{...tdR("#276749"),fontWeight:700}}>{totAdicional>0?fmt$(totAdicional):"—"}</td>
                        <td style={{...tdR(),fontWeight:700}}>{totSac>0?fmt$(totSac):"—"}</td>
                        <td style={{...tdR("#c53030"),fontWeight:700}}>{totDesc>0?fmt$(totDesc):"—"}</td>
                        <td style={{...tdR("#b45309"),fontWeight:700}}>{totAdelanto>0?fmt$(totAdelanto):"—"}</td>
                        <td style={{...tdR("#276749"),fontWeight:700,fontSize:14}}>{fmt$(totSubtotal)}</td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              )}
            </div>
          );
        })()}

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

          // Fracciones de demora con tolerancia POR DÍA
          const fraccionesDemora = fraccionesDemoraCalc(rangeCalcs);
          // Retiros anticipados: misma lógica que tardanzas (fracciones de 15 min)
          const fraccionesSalTemp = fraccionesSalTempCalc(rangeCalcs);
          // (se mantiene el total en horas solo por compatibilidad de display)
          const horasSalTemp      = parseFloat((totalSalTempMin / 60).toFixed(2));
          // Horas extra — decimal para cálculos, HH:MM para mostrar.
          // Si RRHH cargó una corrección manual (hsExtraRelojManual), ESA manda.
          const hsRelojOverride = p.hsExtraRelojManual !== undefined && p.hsExtraRelojManual !== ""
            ? parseFloat(p.hsExtraRelojManual) : null;
          const extraMinFinal   = hsRelojOverride !== null ? Math.round(hsRelojOverride * 60) : totalExtraMin;
          const horasExtraRelojDisplay = minsToDisplay(totalExtraMin); // lo que marcó el reloj
          const horasExtra        = parseFloat((extraMinFinal / 60).toFixed(10)); // full precision for math
          const horasExtraDisplay = minsToDisplay(extraMinFinal); // ej: 38h26

          // Manual inputs (stored per employee)
          const sueldoBasico  = parseFloat(p.sueldoBasico  || 0);
          const valorDia      = parseFloat(p.valorDia      || 0);
          const valorHora     = parseFloat(p.valorHora     || 0);
          const valorHoraExt  = parseFloat(p.valorHoraExt  || 0);
          const valorDiaFinde = parseFloat(p.valorDiaFinde || 0);
          // Adelantos: lista de líneas {desc, monto}. Compat: si existe el viejo p.adelanto numérico y no hay lista, lo migra.
          const adelantos = Array.isArray(p.adelantos)
            ? p.adelantos
            : (parseFloat(p.adelanto||0) > 0 ? [{desc:"Adelanto", monto:String(parseFloat(p.adelanto))}] : []);
          const adelanto      = adelantos.reduce((s,a)=>s+(parseFloat(a.monto)||0),0);
          const feriados      = parseFloat(p.feriados      || 0);
          const sac           = parseFloat(p.sac           || 0);
          const vacaciones    = parseFloat(p.vacaciones    || 0);
          const periodo       = p.periodo || "";
          const ingreso       = p.ingreso || "";

          // Horas extra manuales (fuera del reloj) — se suman a las del reloj
          const horasExtraManualHs  = parseFloat(p.horasExtraManualHs  || 0); // en horas
          const horasExtraManualImp = parseFloat(p.horasExtraManualImp || 0); // en importe
          // Si pusieron importe directo, toma ese; si pusieron horas, calcula importe
          const importeExtraManual  = horasExtraManualImp > 0
            ? horasExtraManualImp
            : horasExtraManualHs * valorHoraExt;
          const horasExtraManualDisplay = horasExtraManualHs > 0
            ? minsToDisplay(Math.round(horasExtraManualHs*60))
            : horasExtraManualImp > 0 && valorHoraExt > 0
            ? minsToDisplay(Math.round((horasExtraManualImp/valorHoraExt)*60))
            : "—";

          // Descuentos — pueden sobreescribirse manualmente
          const descDemorasCalc   = (valorHora / 4) * fraccionesDemora;
          const descSalTempCalc   = (valorHora / 4) * fraccionesSalTemp;
          const descDemorasManual = p.descDemorasManual !== undefined && p.descDemorasManual !== ""
            ? parseFloat(p.descDemorasManual)
            : null;
          const descSalTempManual = p.descSalTempManual !== undefined && p.descSalTempManual !== ""
            ? parseFloat(p.descSalTempManual)
            : null;
          const descDemoras = descDemorasManual !== null ? descDemorasManual : descDemorasCalc;
          const descSalTemp = descSalTempManual !== null ? descSalTempManual : descSalTempCalc;

          // Cantidades a mostrar: si el descuento se borró (importe 0), no mostrar unidades
          const fraccionesDemoraDisp = descDemoras > 0 ? fraccionesDemora  : "—";
          const fraccionesSalTempDisp= descSalTemp > 0 ? fraccionesSalTemp : "—";

          // Calculations — básico manda, adicionales son extras sobre él
          const importeSueldo    = sueldoBasico;
          // Importes calculados de cada adicional + override manual por importe.
          // Vacío = usa el calculado; con valor cargado, ese importe manda.
          const ovr = (v) => v !== undefined && v !== "" ? parseFloat(v) : null;
          const impExtrasCalc     = valorHoraExt * horasExtra;
          const impFeriadosCalc   = valorDia     * feriados;
          const impVacacionesCalc = valorDia     * vacaciones;
          const impFindeCalc      = valorDiaFinde * diasFinde;
          const impExtrasManual     = ovr(p.impExtrasManual);
          const impFeriadosManual   = ovr(p.impFeriadosManual);
          const impVacacionesManual = ovr(p.impVacacionesManual);
          const impFindeManual      = ovr(p.impFindeManual);
          const impExtrasReloj   = impExtrasManual     !== null ? impExtrasManual     : impExtrasCalc;
          const importeFeriados  = impFeriadosManual   !== null ? impFeriadosManual   : impFeriadosCalc;
          const importeVacaciones= impVacacionesManual !== null ? impVacacionesManual : impVacacionesCalc;
          const importeFinde     = impFindeManual      !== null ? impFindeManual      : impFindeCalc;
          const importeExtras    = impExtrasReloj + importeExtraManual;
          const totalAdicionales = importeExtras + importeFeriados + sac + importeVacaciones + importeFinde;
          const totalDescuentos  = descDemoras + descSalTemp;
          const subtotal         = importeSueldo + totalAdicionales;
          const totalACobrar     = subtotal - totalDescuentos - adelanto;

          const fmt = (n) => n === 0 ? "—" : `$${Math.round(n).toLocaleString("es-AR")}`;
          const fmtN = (n) => n === 0 ? "—" : Math.round(n).toLocaleString("es-AR");

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
                <button onClick={()=>setNuevoPeriodoOpen(o=>!o)}
                  style={{alignSelf:"flex-end",background:nuevoPeriodoOpen?COL.accentBg:"#fff",color:COL.accent,border:`1px solid ${COL.accent}`,borderRadius:8,padding:"9px 18px",cursor:"pointer",fontFamily:SANS,fontWeight:600,fontSize:13,display:"flex",alignItems:"center",gap:8,whiteSpace:"nowrap"}}>
                  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round"><path d="M21 2v6h-6"/><path d="M3 12a9 9 0 0115-6.7L21 8"/><path d="M3 22v-6h6"/><path d="M21 12a9 9 0 01-15 6.7L3 16"/></svg>
                  Nuevo período
                </button>
                {selEmp&&(
                  <button onClick={()=>exportLiqPDF({selEmp,periodo,ingreso,desde,hasta,importeSueldo,diasFinde,valorDiaFinde,importeFinde,horasExtra,horasExtraDisplay,valorHoraExt,importeExtras,impExtrasReloj,importeExtraManual,horasExtraManualDisplay,feriados,valorDia,importeFeriados,sac,vacaciones,importeVacaciones,totalAdicionales,subtotal,fraccionesDemora,valorHora,descDemoras,fraccionesSalTemp,descSalTemp,totalDescuentos,adelanto,adelantos,totalACobrar,diasTrabajados,nombreDisplay:p.nombreDisplay||(cap(selEmp.nombre)),fmt})}
                    style={{alignSelf:"flex-end",background:"#276749",color:"#fff",border:"none",borderRadius:8,padding:"9px 20px",cursor:"pointer",fontFamily:SANS,fontWeight:600,fontSize:13,display:"flex",alignItems:"center",gap:8,whiteSpace:"nowrap"}}>
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><path d="M6 9V2h12v7"/><path d="M6 18H4a2 2 0 01-2-2v-5a2 2 0 012-2h16a2 2 0 012 2v5a2 2 0 01-2 2h-2"/><path d="M6 14h12v8H6z"/></svg>
                    Exportar PDF
                  </button>
                )}
              </div>

              {/* ── Formulario "Nuevo período" ── */}
              {nuevoPeriodoOpen && (()=>{
                const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
                const hoy = new Date();
                const defMonth = `${hoy.getFullYear()}-${String(hoy.getMonth()+1).padStart(2,"0")}`;
                const empConDatos = Object.keys(liqParams).filter(k=>liqParams[k]?.sueldoBasico).length;
                return (
                  <NuevoPeriodoForm
                    MESES={MESES}
                    defMonth={defMonth}
                    empCount={empConDatos}
                    onCancel={()=>setNuevoPeriodoOpen(false)}
                    onConfirm={(nombre, desde, hasta)=>{
                      iniciarPeriodo(nombre, desde, hasta);
                      setNuevoPeriodoOpen(false);
                    }}
                  />
                );
              })()}

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
                      <FieldInput label="Sueldo básico"            value={p.sueldoBasico||""}  onChange={v=>setP("sueldoBasico",v)} />
                      <FieldInput label="Valor día"                value={p.valorDia||""}      onChange={v=>setP("valorDia",v)} />
                      <FieldInput label="Valor hora"               value={p.valorHora||""}     onChange={v=>setP("valorHora",v)} />
                      <FieldInput label="Valor hora extra"         value={p.valorHoraExt||""}  onChange={v=>setP("valorHoraExt",v)} />
                      <FieldInput label="Adicional día finde/especial" value={p.valorDiaFinde||""} onChange={v=>setP("valorDiaFinde",v)} note="por día trabajado en finde" />
                    </div>

                    <div style={{background:COL.surface,border:`1px solid ${COL.border}`,borderRadius:12,padding:"18px 20px",marginBottom:16}}>
                      <div style={{fontSize:11,color:COL.textFaint,fontWeight:600,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:14}}>Adicionales manuales</div>
                      <FieldInput label="SAC"                      value={p.sac||""}        onChange={v=>setP("sac",v)}        note="importe directo" />
                      <FieldInput label="Vacaciones (días)"        value={p.vacaciones||""} onChange={v=>setP("vacaciones",v)} prefix="" note="× valor día" />
                      <FieldInput label="Feriados (días)"          value={p.feriados||""}   onChange={v=>setP("feriados",v)}   prefix="" note="× valor día" />

                      {/* Adelantos — líneas de detalle */}
                      <div style={{marginTop:14,paddingTop:14,borderTop:`1px solid ${COL.border}`}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                          <div style={{fontSize:11,color:"#b45309",fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase"}}>
                            Adelantos
                          </div>
                          <button
                            onClick={()=>setP("adelantos",[...adelantos,{desc:"",monto:""}])}
                            style={{...S.btnS,background:"#b45309",display:"flex",alignItems:"center",gap:5,padding:"4px 11px"}}>
                            <span style={{fontSize:14,lineHeight:1}}>+</span> Agregar
                          </button>
                        </div>
                        {adelantos.length===0 ? (
                          <div style={{fontSize:12,color:COL.textFaint,padding:"4px 0 2px"}}>Sin adelantos cargados.</div>
                        ) : (
                          <div style={{display:"flex",flexDirection:"column",gap:7}}>
                            {adelantos.map((a,idx)=>(
                              <div key={idx} style={{display:"flex",alignItems:"center",gap:8}}>
                                <input
                                  type="text"
                                  value={a.desc}
                                  onChange={e=>{const arr=adelantos.map((x,i)=>i===idx?{...x,desc:e.target.value}:x);setP("adelantos",arr);}}
                                  placeholder="Descripción (ej: 3/7 auto)"
                                  style={{...INPUT_STYLE,flex:1,padding:"6px 10px",fontFamily:SANS,fontSize:13}}
                                />
                                <span style={{fontSize:12,color:COL.textFaint}}>$</span>
                                <input
                                  type="number" min="0" step="any" inputMode="decimal"
                                  value={a.monto}
                                  onChange={e=>{const arr=adelantos.map((x,i)=>i===idx?{...x,monto:e.target.value}:x);setP("adelantos",arr);}}
                                  placeholder="0"
                                  style={{...INPUT_STYLE,width:130,padding:"6px 10px",fontFamily:MONO,fontSize:13}}
                                />
                                <button
                                  onClick={()=>setP("adelantos",adelantos.filter((_,i)=>i!==idx))}
                                  title="Quitar"
                                  style={{background:"none",border:"none",cursor:"pointer",color:COL.textFaint,fontSize:16,lineHeight:1,padding:"0 4px"}}>×</button>
                              </div>
                            ))}
                          </div>
                        )}
                        {adelanto>0&&(
                          <div style={{marginTop:9,fontSize:12,color:"#b45309",fontFamily:MONO,fontWeight:600,textAlign:"right"}}>
                            Total adelantos: ${Math.round(adelanto).toLocaleString("es-AR")}
                          </div>
                        )}
                      </div>

                      {/* Corrección manual de las hs. extra del reloj */}
                      <div style={{marginTop:14,paddingTop:14,borderTop:`1px solid ${COL.border}`}}>
                        <div style={{fontSize:11,color:"#9c4221",fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:10}}>
                          Hs. extra del reloj — corrección manual
                        </div>
                        {(()=>{
                          const { h: curH, m: curM } = decimalToHM(p.hsExtraRelojManual);
                          const setHM = (h, m) => setP("hsExtraRelojManual", hmToDecimal(h, m));
                          return (
                            <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap"}}>
                              <div style={{minWidth:80}}>
                                <div style={{fontSize:11,color:COL.textFaint,marginBottom:4}}>Horas</div>
                                <div style={{display:"flex",alignItems:"center",gap:4}}>
                                  <input type="number" min="0" step="1" inputMode="numeric"
                                    value={curH}
                                    onChange={e=>setHM(e.target.value, curM)}
                                    placeholder="—"
                                    style={{...S.sInput,width:64,padding:"6px 10px",fontFamily:MONO,fontSize:13}}
                                  />
                                  <span style={{fontSize:12,color:COL.textFaint,whiteSpace:"nowrap"}}>h</span>
                                </div>
                              </div>
                              <div style={{minWidth:80}}>
                                <div style={{fontSize:11,color:COL.textFaint,marginBottom:4}}>Minutos</div>
                                <div style={{display:"flex",alignItems:"center",gap:4}}>
                                  <input type="number" min="0" max="59" step="1" inputMode="numeric"
                                    value={curM}
                                    onChange={e=>{
                                      let m = e.target.value;
                                      if(m !== "" && parseInt(m,10) > 59) m = "59";
                                      setHM(curH, m);
                                    }}
                                    placeholder="—"
                                    style={{...S.sInput,width:64,padding:"6px 10px",fontFamily:MONO,fontSize:13}}
                                  />
                                  <span style={{fontSize:12,color:COL.textFaint,whiteSpace:"nowrap"}}>min</span>
                                </div>
                              </div>
                              <button onClick={()=>setP("hsExtraRelojManual","0")}
                                title="Liquidar 0 horas extra aunque el reloj haya marcado"
                                style={{...S.btnS,background:"#9c4221",padding:"6px 11px"}}>Anular (0h)</button>
                              {hsRelojOverride !== null && (
                                <button onClick={()=>setP("hsExtraRelojManual","")}
                                  style={{...S.btnS,background:"#64748b",padding:"6px 11px"}}>Usar reloj</button>
                              )}
                            </div>
                          );
                        })()}
                        <div style={{marginTop:8,fontSize:12,color: hsRelojOverride !== null ? "#9c4221" : COL.textFaint}}>
                          {hsRelojOverride !== null
                            ? <>Usando corrección manual: <b style={{fontFamily:MONO}}>{horasExtraDisplay}</b> · el reloj marcó <b style={{fontFamily:MONO}}>{horasExtraRelojDisplay}</b></>
                            : <>Vacío = se usa el total del reloj (<b style={{fontFamily:MONO}}>{horasExtraRelojDisplay}</b>)</>}
                        </div>
                      </div>

                      {/* Horas extra manuales — fuera del reloj */}
                      <div style={{marginTop:14,paddingTop:14,borderTop:`1px solid ${COL.border}`}}>
                        <div style={{fontSize:11,color:"#276749",fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:10}}>
                          Hs. extra fuera del reloj
                        </div>
                        {(()=>{
                          const { h: curH, m: curM } = decimalToHM(p.horasExtraManualHs);
                          const setHM = (h, m) => {
                            const dec = hmToDecimal(h, m);
                            setP("horasExtraManualHs", dec);
                            setP("horasExtraManualImp", dec && valorHoraExt ? (parseFloat(dec)*valorHoraExt).toFixed(2) : "");
                          };
                          return (
                        <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap"}}>
                          <div style={{minWidth:80}}>
                            <div style={{fontSize:11,color:COL.textFaint,marginBottom:4}}>Horas</div>
                            <div style={{display:"flex",alignItems:"center",gap:4}}>
                              <input type="number" min="0" step="1" inputMode="numeric"
                                value={curH}
                                onChange={e=>setHM(e.target.value, curM)}
                                placeholder="0"
                                style={{...S.sInput,width:64,padding:"6px 10px",fontFamily:MONO,fontSize:13}}
                              />
                              <span style={{fontSize:12,color:COL.textFaint,whiteSpace:"nowrap"}}>h</span>
                            </div>
                          </div>
                          <div style={{minWidth:80}}>
                            <div style={{fontSize:11,color:COL.textFaint,marginBottom:4}}>Minutos</div>
                            <div style={{display:"flex",alignItems:"center",gap:4}}>
                              <input type="number" min="0" max="59" step="1" inputMode="numeric"
                                value={curM}
                                onChange={e=>{
                                  let m = e.target.value;
                                  if(m !== "" && parseInt(m,10) > 59) m = "59";
                                  setHM(curH, m);
                                }}
                                placeholder="0"
                                style={{...S.sInput,width:64,padding:"6px 10px",fontFamily:MONO,fontSize:13}}
                              />
                              <span style={{fontSize:12,color:COL.textFaint,whiteSpace:"nowrap"}}>min</span>
                            </div>
                          </div>
                          <div style={{fontSize:13,color:COL.textFaint,paddingBottom:8}}>↔</div>
                          <div style={{flex:1,minWidth:120}}>
                            <div style={{fontSize:11,color:COL.textFaint,marginBottom:4}}>Importe</div>
                            <div style={{display:"flex",alignItems:"center",gap:4}}>
                              <span style={{fontSize:12,color:COL.textFaint}}>$</span>
                              <input type="number" min="0" step="any" inputMode="decimal"
                                value={p.horasExtraManualImp||""}
                                onChange={e=>{
                                  const imp = e.target.value;
                                  const hs = imp && valorHoraExt ? (parseFloat(imp)/valorHoraExt).toString() : "";
                                  setP("horasExtraManualImp", imp);
                                  setP("horasExtraManualHs", hs);
                                }}
                                placeholder="0"
                                style={{...S.sInput,width:"100%",padding:"6px 10px",fontFamily:MONO,fontSize:13}}
                              />
                            </div>
                          </div>
                        </div>
                          );
                        })()}
                        {importeExtraManual > 0 && (
                          <div style={{marginTop:8,fontSize:11,color:"#276749",fontFamily:MONO}}>
                            + {horasExtraManualDisplay} = ${Math.round(importeExtraManual).toLocaleString("es-AR")}
                          </div>
                        )}
                      </div>

                      {/* Adicionales con importe manual */}
                      <div style={{marginTop:14,paddingTop:14,borderTop:`1px solid ${COL.border}`}}>
                        <div style={{fontSize:11,color:"#276749",fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:10}}>
                          Adicionales — importe manual
                        </div>
                        <div style={{fontSize:11,color:COL.textFaint,marginBottom:10}}>
                          Vacío = usa el importe calculado. Con valor, ese importe manda en la planilla.
                        </div>
                        {[
                          {label:"Horas extra (reloj)",   field:"impExtrasManual",     calc:impExtrasCalc},
                          {label:"Días finde/especiales", field:"impFindeManual",      calc:impFindeCalc},
                          {label:"Feriados",              field:"impFeriadosManual",   calc:impFeriadosCalc},
                          {label:"Vacaciones",            field:"impVacacionesManual", calc:impVacacionesCalc},
                        ].map(({label,field,calc})=>{
                          const activo = p[field] !== undefined && p[field] !== "";
                          return (
                            <div key={field} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                              <label style={{fontSize:12,color:COL.textSub,minWidth:200,flexShrink:0}}>
                                {label}
                                <span style={{marginLeft:6,fontSize:10,color:COL.textFaint}}>
                                  (calc: ${Math.round(calc||0).toLocaleString("es-AR")})
                                </span>
                              </label>
                              <div style={{display:"flex",alignItems:"center",gap:4}}>
                                <span style={{fontSize:12,color:COL.textFaint}}>$</span>
                                <input type="number" min="0" step="any"
                                  value={p[field] !== undefined ? p[field] : ""}
                                  onChange={e=>setP(field, e.target.value)}
                                  placeholder={Math.round(calc||0)||"0"}
                                  style={{...S.sInput,width:140,padding:"6px 10px",fontFamily:MONO,fontSize:13,
                                    borderColor: activo?"#f59e0b":undefined}}
                                />
                                {activo&&<button onClick={()=>setP(field,"")}
                                  style={{fontSize:10,color:COL.textFaint,background:"none",border:"none",cursor:"pointer"}}>↺</button>}
                              </div>
                            </div>
                          );
                        })}
                      </div>

                      {/* Descuentos editables */}
                      <div style={{marginTop:14,paddingTop:14,borderTop:`1px solid ${COL.border}`}}>
                        <div style={{fontSize:11,color:"#c53030",fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:10}}>
                          Descuentos (editables)
                        </div>
                        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                          <label style={{fontSize:12,color:COL.textSub,minWidth:200,flexShrink:0}}>
                            Llegadas tarde
                            <span style={{marginLeft:6,fontSize:10,color:COL.textFaint}}>
                              (calc: ${Math.round(descDemorasCalc).toLocaleString("es-AR")})
                            </span>
                          </label>
                          <div style={{display:"flex",alignItems:"center",gap:4}}>
                            <span style={{fontSize:12,color:COL.textFaint}}>$</span>
                            <input type="number" min="0" step="any"
                              value={p.descDemorasManual !== undefined ? p.descDemorasManual : ""}
                              onChange={e=>setP("descDemorasManual", e.target.value)}
                              placeholder={Math.round(descDemorasCalc)||"0"}
                              style={{...S.sInput,width:140,padding:"6px 10px",fontFamily:MONO,fontSize:13,
                                borderColor: descDemorasManual!==null?"#f59e0b":undefined}}
                            />
                            {descDemorasManual!==null&&<button onClick={()=>setP("descDemorasManual","")}
                              style={{fontSize:10,color:COL.textFaint,background:"none",border:"none",cursor:"pointer"}}>↺</button>}
                          </div>
                        </div>
                        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                          <label style={{fontSize:12,color:COL.textSub,minWidth:200,flexShrink:0}}>
                            Retiros anticipados
                            <span style={{marginLeft:6,fontSize:10,color:COL.textFaint}}>
                              (calc: ${Math.round(descSalTempCalc).toLocaleString("es-AR")})
                            </span>
                          </label>
                          <div style={{display:"flex",alignItems:"center",gap:4}}>
                            <span style={{fontSize:12,color:COL.textFaint}}>$</span>
                            <input type="number" min="0" step="any"
                              value={p.descSalTempManual !== undefined ? p.descSalTempManual : ""}
                              onChange={e=>setP("descSalTempManual", e.target.value)}
                              placeholder={Math.round(descSalTempCalc)||"0"}
                              style={{...S.sInput,width:140,padding:"6px 10px",fontFamily:MONO,fontSize:13,
                                borderColor: descSalTempManual!==null?"#f59e0b":undefined}}
                            />
                            {descSalTempManual!==null&&<button onClick={()=>setP("descSalTempManual","")}
                              style={{fontSize:10,color:COL.textFaint,background:"none",border:"none",cursor:"pointer"}}>↺</button>}
                          </div>
                        </div>
                      </div>

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
                      {(()=>{
                        const ausInjust = [...records,...manualRecords].filter(r=>
                          r.empNo===liqEmp && r.ausencia==="aus_injust" &&
                          (!desde||r.fecha>=desde) && (!hasta||r.fecha<=hasta)
                        );
                        const ausJust = [...records,...manualRecords].filter(r=>
                          r.empNo===liqEmp && r.ausencia==="aus_just" &&
                          (!desde||r.fecha>=desde) && (!hasta||r.fecha<=hasta)
                        );
                        return [
                          ["Días trabajados",           diasTrabajados,       "días (referencia)"],
                          ["Días finde trabajados",     diasFinde,            "días"],
                          ["Horas extra",               horasExtraDisplay,    ""],
                          ["Demoras (fracc. 15 min)",   fraccionesDemora,     "fracciones"],
                          ["Retiros (fracc. 15 min)",   fraccionesSalTemp,    "fracciones"],
                          ...(ausJust.length   ? [["Ausencias justificadas",   ausJust.length,   "días — no descuenta"]] : []),
                          ...(ausInjust.length ? [["⚠ Ausencias injustificadas", ausInjust.length, "días — revisar descuento"]] : []),
                        ];
                      })().map(([l,v,u])=>(
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
                          <LiqRow label={`Horas extra (reloj)${impExtrasManual!==null?" ✎":""}`} indent cantidad={horasExtraDisplay} valor={impExtrasManual!==null?"—":fmt(valorHoraExt)} importe={impExtrasReloj?fmt(impExtrasReloj):"—"} />
                          {importeExtraManual>0&&<LiqRow label="Horas extra (manual)" indent cantidad={horasExtraManualDisplay} valor={fmt(valorHoraExt)} importe={fmt(importeExtraManual)} />}
                          <LiqRow label={`Días finde/especiales${impFindeManual!==null?" ✎":""}`} indent cantidad={diasFinde||"—"} valor={impFindeManual!==null?"—":(diasFinde?fmt(valorDiaFinde):"—")} importe={importeFinde?fmt(importeFinde):"—"} />
                          <LiqRow label={`Feriados${impFeriadosManual!==null?" ✎":""}`}          indent cantidad={feriados||"—"}  valor={impFeriadosManual!==null?"—":fmt(valorDia)}     importe={importeFeriados?fmt(importeFeriados):"—"} />
                          <LiqRow label="SAC"               indent cantidad=""              valor=""                  importe={sac?fmt(sac):"—"} />
                          <LiqRow label={`Vacaciones${impVacacionesManual!==null?" ✎":""}`}        indent cantidad={vacaciones||"—"} valor={impVacacionesManual!==null?"—":fmt(valorDia)}   importe={importeVacaciones?fmt(importeVacaciones):"—"} />
                          <LiqRow label="Subtotal adicionales" cantidad="" valor="" importe={fmt(totalAdicionales)} bold separator />

                          <LiqRow label="Sueldo + adicionales" cantidad="" valor="" importe={fmt(subtotal)} bold color={COL.accent} separator />

                          <LiqRow label="DESCUENTOS" cantidad="" valor="" importe="" bold separator />
                          <LiqRow label={`Llegadas tarde${descDemorasManual!==null?" ✎":""}`} indent cantidad={fraccionesDemoraDisp} valor={descDemoras?`${fmt(valorHora)}/4`:"—"} importe={descDemoras?fmt(descDemoras):"—"} color={descDemoras?"#c53030":undefined} />
                          <LiqRow label={`Retiros anticipados${descSalTempManual!==null?" ✎":""}`} indent cantidad={fraccionesSalTempDisp} valor={descSalTemp?`${fmt(valorHora)}/4`:"—"} importe={descSalTemp?fmt(descSalTemp):"—"} color={descSalTemp?"#c53030":undefined} />
                          <LiqRow label="Total descuentos" cantidad="" valor="" importe={totalDescuentos?fmt(totalDescuentos):"—"} bold color="#c53030" separator />

                          <LiqRow label="Adelantos" cantidad="" valor="" importe="" bold color="#b45309" separator />
                          {adelantos.filter(a=>parseFloat(a.monto)>0).map((a,i)=>(
                            <LiqRow key={i} indent label={a.desc||"Adelanto"} cantidad="" valor="" importe={fmt(parseFloat(a.monto))} color="#b45309" />
                          ))}
                          <LiqRow label="Total adelantos" cantidad="" valor="" importe={adelanto?fmt(adelanto):"—"} bold color="#b45309" />

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

        {tab===8&&(()=>{
          const activeEmps = empList.filter(e=>e.activo!==false);

          // Último día cargado en el sistema = fecha máxima entre todos los registros
          const ultimoDia = allRecs.length
            ? allRecs.reduce((max,r)=> r.fecha>max ? r.fecha : max, allRecs[0].fecha)
            : null;

          // helper: registros de un empleado, sin ausencias, con overrides manuales
          const recsDe = (empNo)=> allRecs
            .filter(r=>r.empNo===empNo && !r.ausencia)
            .map(r=>{
              const entrada = r.manual ? r.entrada : (manualSalidas[r.id+"_ent"] || r.entrada);
              const salida  = r.manual ? r.salida  : (manualSalidas[r.id]        || r.salida);
              return {...r, entrada, salida};
            });

          // CON marca en el último día
          const conMarca = [];
          // SIN marca ese día (mostramos su última marca histórica)
          const sinMarca = [];

          for (const emp of activeEmps) {
            const recs = recsDe(emp.empNo);
            const recDelDia = recs.find(r=>r.fecha===ultimoDia && (r.entrada||r.salida));
            if (recDelDia) {
              // Salida = del registro ANTERIOR con salida (día hábil previo: 28, 26, etc.)
              const salPrev = recs
                .filter(r=>r.fecha < ultimoDia && r.salida)
                .sort((a,b)=>b.fecha.localeCompare(a.fecha))[0] || null;
              conMarca.push({
                emp,
                ingreso:     recDelDia.entrada || null,
                fechaIng:    ultimoDia,
                salida:      salPrev?.salida || null,
                fechaSal:    salPrev?.fecha  || null,
                recId:       recDelDia.id,
                recManual:   recDelDia.manual || false,
                observacion: recDelDia.observacion || null,
              });
            } else {
              // su registro más reciente con algo cargado
              const ult = recs.filter(r=>r.entrada||r.salida)
                .sort((a,b)=>b.fecha.localeCompare(a.fecha))[0] || null;
              // y la salida del registro anterior a ese
              const salPrev = ult ? (recs
                .filter(r=>r.fecha < ult.fecha && r.salida)
                .sort((a,b)=>b.fecha.localeCompare(a.fecha))[0] || null) : null;
              sinMarca.push({
                emp,
                ingreso:  ult?.entrada || null,
                fechaIng: ult?.fecha   || null,
                salida:   (ult?.salida) || (salPrev?.salida) || null,
                fechaSal: (ult?.salida ? ult.fecha : salPrev?.fecha) || null,
                recId:       ult?.id || null,
                recManual:   ult?.manual || false,
                observacion: ult?.observacion || null,
              });
            }
          }

          conMarca.sort((a,b)=>a.emp.empNo-b.emp.empNo);
          sinMarca.sort((a,b)=>(b.fechaIng||"").localeCompare(a.fechaIng||"")||a.emp.empNo-b.emp.empNo);

          const fmtFecha = (f)=>{
            if(!f) return "—";
            const d = new Date(f+"T12:00:00");
            const dd = ["Dom","Lun","Mar","Mié","Jue","Vie","Sáb"][d.getDay()];
            return `${dd} ${f.slice(8,10)}/${f.slice(5,7)}`;
          };

          const exportarNovedades = ()=>{
            const headers = [
              ["empNo","N°"],["nombre","Nombre"],["estado","Estado"],
              ["fechaIng","Fecha ingreso"],["ingreso","Ingreso"],
              ["fechaSal","Fecha salida"],["salida","Salida"],
              ["observacion","Observaciones"],
            ];
            const rows = [
              ...conMarca.map(f=>({empNo:f.emp.empNo,nombre:cap(f.emp.nombre),estado:"Presente",
                fechaIng:f.fechaIng||"",ingreso:f.ingreso||"",fechaSal:f.fechaSal||"",salida:f.salida||"",observacion:f.observacion||""})),
              ...sinMarca.map(f=>({empNo:f.emp.empNo,nombre:cap(f.emp.nombre),estado:"Sin marca",
                fechaIng:f.fechaIng||"",ingreso:f.ingreso||"",fechaSal:f.fechaSal||"",salida:f.salida||"",observacion:f.observacion||""})),
            ];
            exportXLSX(rows, headers, "Novedades", `novedades_${ultimoDia||"sd"}.xlsx`);
          };

          const exportarNovedadesPDF = ()=>{
            exportNovedadesPDF({ conMarca, sinMarca, ultimoDia, fmtFecha });
          };

          // Guarda la observación contra el registro de ingreso de esa fila
          const guardarObsNov = (f, valor)=>{
            const v = (valor||"").trim();
            if(!f.recId) return; // sin registro asociado no hay dónde guardar
            if(f.recManual){
              setManualRecords(p=>p.map(x=>x.id===f.recId?{...x,observacion:v||null}:x));
            } else {
              setRecords(p=>p.map(x=>x.id===f.recId?{...x,observacion:v||null}:x));
            }
            sbUpdate("registros",f.recId,{observacion:v||null},"id");
          };

          const RowNov = (f, faded)=>{
            const TIPO = TIPO_CFG[f.emp.tipo||"operario"]||TIPO_CFG.operario;
            const editKey = `nov_${f.emp.empNo}`;
            const editing = editingCell?.id===editKey && editingCell?.field==="observacion";
            return (
              <tr key={f.emp.empNo}>
                <td style={{...S.td,textAlign:"left",fontFamily:MONO,color:COL.textFaint}}>{f.emp.empNo}</td>
                <td style={{...S.td,textAlign:"left"}}>
                  <span style={{display:"inline-flex",alignItems:"center",gap:7}}>
                    <span style={{width:7,height:7,borderRadius:"50%",background:TIPO.color,flexShrink:0}}/>
                    {cap(f.emp.nombre)}
                  </span>
                </td>
                <td style={{...S.td,fontSize:12,color:COL.textFaint}}>{fmtFecha(f.fechaIng)}</td>
                <td style={{...S.td,fontFamily:MONO,fontWeight:600,color:f.ingreso?(faded?COL.textFaint:"#276749"):COL.textFaint}}>{f.ingreso||"—"}</td>
                <td style={{...S.td,fontSize:12,color:COL.textFaint}}>{fmtFecha(f.fechaSal)}</td>
                <td style={{...S.td,fontFamily:MONO,fontWeight:600,color:f.salida?(faded?COL.textFaint:"#1e5fa8"):COL.textFaint}}>{f.salida||"—"}</td>
                <td style={{...S.td,fontSize:11,padding:"4px 8px",maxWidth:200,textAlign:"left"}}>
                  {editing
                    ? <input autoFocus
                        defaultValue={f.observacion||""}
                        onBlur={e=>{ guardarObsNov(f, e.target.value); setEditingCell(null); }}
                        onKeyDown={e=>{ if(e.key==="Escape")setEditingCell(null); if(e.key==="Enter")e.target.blur(); }}
                        placeholder="Escribir observación…"
                        style={{...S.inlineInput,fontSize:11,padding:"3px 6px",width:"100%"}}
                      />
                    : <span
                        onClick={()=>f.recId&&setEditingCell({id:editKey,field:"observacion"})}
                        title={f.recId?"Click para editar observación":"Sin registro asociado"}
                        style={{color:f.observacion?COL.textSub:"#d1d5db",
                          cursor:f.recId?"pointer":"default",
                          borderBottom:f.recId?"1px dashed #d1d5db":"none",display:"block",
                          overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:190}}>
                        {f.observacion||(f.recId?<span style={{fontSize:10,fontStyle:"italic"}}>agregar obs.</span>:"—")}
                      </span>
                  }
                </td>
              </tr>
            );
          };

          return (
            <div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:12,flexWrap:"wrap"}}>
                <div>
                  <H2>Novedades</H2>
                  <p style={S.body}>
                    Ingreso del último día cargado{ultimoDia?<> (<strong>{fmtFecha(ultimoDia)}</strong>)</>:""} y la última salida registrada antes de ese día.
                    Como la planilla se descarga al mediodía, el ingreso es de ese día y la salida corresponde al día hábil anterior.
                  </p>
                </div>
                {records.length>0&&(
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={exportarNovedades} style={{...S.btnS,display:"flex",alignItems:"center",gap:6}}>
                      <span style={{fontSize:14,lineHeight:1}}>↓</span> Exportar Excel
                    </button>
                    <button onClick={exportarNovedadesPDF} style={{...S.btnS,background:"#fff",color:"#c53030",border:"1px solid #fed7d7",display:"flex",alignItems:"center",gap:6}}>
                      <span style={{fontSize:14,lineHeight:1}}>↓</span> Exportar PDF
                    </button>
                  </div>
                )}
              </div>

              {records.length===0 ? (
                <div style={S.infoBox}>
                  <span style={{color:COL.textFaint,fontSize:13}}>No hay registros cargados todavía. Importá una planilla en la pestaña Importar.</span>
                </div>
              ) : (
                <>
                  <div style={{fontSize:12,color:COL.textSub,fontWeight:600,margin:"4px 0 8px"}}>
                    Presentes el {fmtFecha(ultimoDia)} · {conMarca.length}
                  </div>
                  <div style={S.tblWrap}>
                    <table style={S.table}>
                      <THead cols={["N°","Nombre","Fecha ingreso","Ingreso","Fecha salida","Salida","Observaciones"]}/>
                      <tbody>
                        {conMarca.length>0
                          ? conMarca.map(f=>RowNov(f,false))
                          : <tr><td colSpan={7} style={{...S.td,color:COL.textFaint,padding:"18px"}}>Nadie marcó ese día.</td></tr>}
                      </tbody>
                    </table>
                  </div>

                  {sinMarca.length>0&&(
                    <>
                      <div style={{fontSize:12,color:"#b45309",fontWeight:600,margin:"22px 0 8px"}}>
                        Sin marca el {fmtFecha(ultimoDia)} · {sinMarca.length} (se muestra su última marca registrada)
                      </div>
                      <div style={{...S.tblWrap,opacity:0.92}}>
                        <table style={S.table}>
                          <THead cols={["N°","Nombre","Fecha ingreso","Ingreso","Fecha salida","Salida","Observaciones"]}/>
                          <tbody>
                            {sinMarca.map(f=>RowNov(f,true))}
                          </tbody>
                        </table>
                      </div>
                    </>
                  )}
                </>
              )}
            </div>
          );
        })()}

      </main>
    </div>
  );
}

/* ─── Auth wrapper ───────────────────────────────────────────────────────── */
export default function App() {
  const [session, setSession] = useState(() => {
    try {
      const s = sessionStorage.getItem("pyg_session");
      return s ? JSON.parse(s) : null;
    } catch { return null; }
  });

  const handleLogin = (sessionData) => {
    sessionStorage.setItem("pyg_session", JSON.stringify(sessionData));
    setSession(sessionData);
  };

  const handleLogout = () => {
    sessionStorage.removeItem("pyg_session");
    setSession(null);
  };

  if (!session) {
    return <LoginScreen onLogin={handleLogin} />;
  }

  return (
    <div style={{ position: "relative" }}>
      {/* Badge de sesión + botón logout */}
      <div style={{
        position: "fixed", bottom: 18, right: 18, zIndex: 9999,
        display: "flex", alignItems: "center", gap: 8,
        background: "#ffffff", border: "1px solid #dde3ec",
        borderRadius: 40, padding: "6px 14px 6px 10px",
        boxShadow: "0 2px 12px rgba(26,58,107,0.12)",
        fontFamily: "'DM Sans', system-ui, sans-serif",
        fontSize: 12,
      }}>
        <span style={{
          width: 8, height: 8, borderRadius: "50%",
          background: "#22c55e", display: "inline-block",
          boxShadow: "0 0 0 2px rgba(34,197,94,0.25)",
        }} />
        <span style={{ color: "#4a5a6a", fontWeight: 500 }}>
          {session.usuario}
        </span>
        <span style={{
          background: "#edf2f9", color: "#3d6b9e",
          borderRadius: 20, padding: "1px 7px",
          fontSize: 10.5, fontWeight: 600, letterSpacing: "0.04em",
        }}>
          {session.rol}
        </span>
        <button
          onClick={handleLogout}
          title="Cerrar sesión"
          style={{
            marginLeft: 4, background: "none", border: "none", cursor: "pointer",
            color: "#96a3b0", padding: 2, display: "flex", alignItems: "center",
          }}
        >
          <svg width="14" height="14" fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
            <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/>
            <polyline points="16 17 21 12 16 7"/>
            <line x1="21" y1="12" x2="9" y2="12"/>
          </svg>
        </button>
      </div>
      <AppMain session={session} />
    </div>
  );
}