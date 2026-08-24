/* ════════════════════════════════════════════════════════════════════
   calculos.js — Lógica pura de asistencia y liquidación (sin React,
   sin Supabase, sin DOM). Todo lo que está acá es testeable con
   `node --test calculos.test.mjs` y es la ÚNICA fuente de verdad del
   cálculo: el tab Liquidación, el Resumen y el snapshot de cierre
   usan calcularLiquidacion().
   ════════════════════════════════════════════════════════════════════ */

export function parseTimeVal(val) {
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

export function minsToHHMM(m) {
  if (m == null || isNaN(m)) return null;
  const a = Math.abs(Math.round(m));
  return `${String(Math.floor(a/60)).padStart(2,"0")}:${String(a%60).padStart(2,"0")}`;
}

export function minsToDisplay(m) {
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
export const TOLERANCIA_DEMORA = 15;
export function fraccionesDeUnDia(demoraMin) {
  const d = Math.round(demoraMin || 0);
  if (d <= TOLERANCIA_DEMORA) return 0;
  return Math.ceil(d / 15) - (d % 15 === 0 ? 0 : 1);
}
export function fraccionesDemoraCalc(calcsDelRango) {
  if (!Array.isArray(calcsDelRango)) return 0;
  return calcsDelRango.reduce((s, r) => s + fraccionesDeUnDia(r.demora), 0);
}
// Retiros anticipados: misma lógica que las tardanzas (por día, 15 de tolerancia,
// fracciones de 15 min). Cada fracción vale valorHora/4.
export function fraccionesSalTempCalc(calcsDelRango) {
  if (!Array.isArray(calcsDelRango)) return 0;
  return calcsDelRango.reduce((s, r) => s + fraccionesDeUnDia(r.salTemprana), 0);
}

// ── Recuperación de horas ─────────────────────────────────────────────────
// Cuando un empleado falta o se retira antes, RRHH puede cargar un registro
// manual con "horas a recuperar" (siempre en múltiplos de 30 min). Las horas
// extra que haga DESPUÉS de esa fecha se aplican a saldar la deuda en vez de
// pagarse, contando solo bloques de 30 min por día (un día con +7min no
// descuenta nada; uno con +1h04 descuenta 1h y los 4min sobrantes se pagan).
// Devuelve { deudaTotal, recuperado, pendiente, saldados } en minutos;
// saldados = Set de ids de registros cuya deuda quedó totalmente cubierta
// (se asigna FIFO: las extras van saldando la deuda más vieja primero).
export function calcRecuperacion(calcs) {
  const orden = [...(calcs || [])].sort((a, b) => a.fecha.localeCompare(b.fecha));
  const pendientes = []; // {id, resta}
  const saldados = new Set();
  let deudaTotal = 0, recuperado = 0;
  for (const r of orden) {
    if (r.recuperar && r.recuperarMin > 0) { pendientes.push({ id: r.id, resta: r.recuperarMin }); deudaTotal += r.recuperarMin; }
    if (r.extra > 0 && !r.recuperar && pendientes.length) {
      let aplicable = Math.floor(r.extra / 30) * 30; // solo bloques de 30 min
      while (aplicable > 0 && pendientes.length) {
        const d = pendientes[0];
        const usa = Math.min(d.resta, aplicable);
        d.resta -= usa; aplicable -= usa; recuperado += usa;
        if (d.resta === 0) { saldados.add(d.id); pendientes.shift(); }
      }
    }
  }
  const pendiente = pendientes.reduce((s, d) => s + d.resta, 0);
  return { deudaTotal, recuperado, pendiente, saldados };
}

export function extractEntradaSalida(row) {
  const marks = [4,5,6,7].map(c=>parseTimeVal(row[c])).filter(v=>v!=null).sort((a,b)=>a-b);
  return {
    entrada: marks.length >= 1 ? minsToHHMM(marks[0]) : null,
    salida:  marks.length >= 2 ? minsToHHMM(marks[marks.length-1]) : null,
    soloEntrada: marks.length === 1,
  };
}

export function parseDateVal(val) {
  if (!val) return null;
  if (typeof val === "string") return val.replace(/\//g,"-").slice(0,10);
  if (typeof val === "number") return new Date(Math.round((val-25569)*86400000)).toISOString().slice(0,10);
  if (val instanceof Date) return val.toISOString().slice(0,10);
  return null;
}

// Slug del nombre del reloj para que el id sea único por persona, no solo por
// número: si el reloj reusa un N° para otra persona, los registros no se pisan.
export const slugNombre = s => String(s).trim().toLowerCase()
  .replace(/[^a-z0-9ñ ]/g, "")
  .replace(/ +/g, "-");


// Horas extra de operarios solo desde las 06:00
export const OPERARIO_EXTRA_FROM = 6 * 60; // 360 min

export function calcRecord(rec, empCfg, specialDays) {
  const cfg = empCfg || { entrada:"06:00", salida:"16:30", tipo:"operario" };
  const esOperario = (cfg.tipo || "operario") === "operario";
  const dayType    = specialDays?.[rec.fecha];
  const diaSemana  = new Date(rec.fecha+"T12:00:00").getDay();
  const esSabado   = diaSemana === 6;
  const esDomingo  = diaSemana === 0;
  const esFeriado  = dayType?.tipo === "feriado";

  // extraCorr = horas extra editadas a mano (minutos). null = cálculo automático.
  const extraManual = rec.extraCorr != null ? (rec.extraCorr > 0 ? rec.extraCorr : null) : undefined;

  if (!rec.entrada || !rec.salida) return { trabajado:null, jornada:null, extra: extraManual !== undefined ? extraManual : null, demora:null, salTemprana:null };

  const entMin = parseTimeVal(rec.entrada);
  const salMin = parseTimeVal(rec.salida);

  // ── Administrativos: sábados y domingos NO son obligatorios ──
  // Todo lo que trabajen esos días se paga como hora extra.
  if (!esOperario && (esSabado || esDomingo)) {
    const total = Math.max(0, salMin - entMin);
    return { trabajado:0, jornada:0, extra: extraManual !== undefined ? extraManual : (total>0 ? total : null), demora:0, salTemprana:0 };
  }

  // ── Administrativos en FERIADO: se computa la jornada normal (el feriado
  // se paga aparte en Liquidación) y solo la diferencia por encima de la
  // jornada de referencia va como hora extra. Sin demora ni salida temprana.
  if (!esOperario && esFeriado) {
    const total      = Math.max(0, salMin - entMin);
    const jornadaRef = parseTimeVal(cfg.salida) - parseTimeVal(cfg.entrada);
    let extra = total > jornadaRef ? total - jornadaRef : null;
    if (extraManual !== undefined) extra = extraManual;
    return { trabajado: Math.min(total, jornadaRef), jornada: jornadaRef, extra, demora:0, salTemprana:0 };
  }

  // Horario de referencia efectivo del día (feriado y sábado acortan la salida)
  const effectiveCfg = esFeriado
    ? { entrada: cfg.entrada, salida: dayType?.salida || "14:00" }
    : esSabado
    ? { entrada: cfg.entrada, salida: "13:00" }
    : cfg;

  // ── Operarios ──
  // Sábados y domingos no acumulan horas extra.
  // Los FERIADOS sí: la salida reducida (ej. 14:00) existe para que irse a esa
  // hora no cuente como retiro anticipado, y lo trabajado después son extras.
  const sinExtra = esOperario && (esSabado || esDomingo);

  const entRef = parseTimeVal(effectiveCfg.entrada);
  const salRef = parseTimeVal(effectiveCfg.salida);
  const jornada = salRef - entRef;
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
  if (extraManual !== undefined) extra = extraManual; // edición manual pisa el cálculo
  const demora      = Math.max(0, entMin - entRef);
  const salTemprana = Math.max(0, salRef - salMin);
  const entDentro   = Math.max(entMin, entRef);
  const salDentro   = Math.min(salMin, salRef);
  const trabajado   = Math.max(0, salDentro - entDentro);
  return { trabajado, jornada, extra, demora, salTemprana };
}

export function detectSchedule(recs) {
  const ents=recs.map(r=>r.entrada?parseTimeVal(r.entrada):null).filter(v=>v!=null).sort((a,b)=>a-b);
  const sals=recs.map(r=>r.salida?parseTimeVal(r.salida):null).filter(v=>v!=null).sort((a,b)=>a-b);
  if (!ents.length) return null;
  return { entrada:minsToHHMM(ents[Math.floor(ents.length/2)]), salida:sals.length?minsToHHMM(sals[Math.floor(sals.length/2)]):"16:30" };
}

/* ────────────────────────────────────────────────────────────────────
   calcularLiquidacion(p, calcs, opts)
   p     = parámetros del empleado para el período (liqParams[empNo])
   calcs = registros del empleado ya pasados por calcRecord (empSummary)
   opts  = { desde, hasta } para forzar un rango distinto de p.desde/p.hasta
   Devuelve TODOS los valores de la liquidación. Es determinística y
   serializable (ver snapshotLiquidacion) salvo por rangeCalcs.
   ──────────────────────────────────────────────────────────────────── */
export function calcularLiquidacion(p, calcs, opts = {}) {
  p = p || {};
  calcs = calcs || [];
  const desde = opts.desde != null ? opts.desde : (p.desde || "");
  const hasta = opts.hasta != null ? opts.hasta : (p.hasta || "");
  const rangeCalcs = calcs.filter(r =>
    (!desde || r.fecha >= desde) && (!hasta || r.fecha <= hasta)
  );

  // Values from attendance data (filtered by range)
  const diasTrabajados   = rangeCalcs.filter(r=>r.trabajado!=null&&r.trabajado>0).length;
  const totalDemoraMin   = rangeCalcs.reduce((s,r)=>s+(r.demora||0),0);
  const totalSalTempMin  = rangeCalcs.reduce((s,r)=>s+(r.salTemprana||0),0);
  const totalExtraMin    = rangeCalcs.reduce((s,r)=>s+(r.extra||0),0);
  // Recuperación de horas: la deuda cargada se descuenta de las extras
  // posteriores (bloques de 30 min); solo se paga el neto.
  const recu0            = calcRecuperacion(rangeCalcs);
  const totalExtraNetoMin= Math.max(0, totalExtraMin - recu0.recuperado);
  // Finde dates: all sat/sun in range (with or without record)
  const allFindeInRange = (()=>{
    if (!desde && !hasta && rangeCalcs.length === 0) return [];
    const fechasSet = new Set(rangeCalcs
      .filter(r=>{ const d=new Date(r.fecha+"T12:00:00").getDay(); return d===0||d===6; })
      .map(r=>r.fecha)
    );
    return [...fechasSet].sort();
  })();
  // Which ones RRHH checked (stored per employee in liqParams)
  const findeSelSet = new Set(p.findeSel || allFindeInRange); // default: all pre-selected
  const diasFinde = findeSelSet.size;

  // Fracciones de demora con tolerancia POR DÍA
  const fraccionesDemora = fraccionesDemoraCalc(rangeCalcs);
  // Retiros anticipados: misma lógica que tardanzas (fracciones de 15 min),
  // pero los retiros con recupero de horas YA SALDADO no se descuentan en plata
  const fraccionesSalTemp = fraccionesSalTempCalc(rangeCalcs.filter(r=>!recu0.saldados.has(r.id)));
  const horasSalTemp      = parseFloat((totalSalTempMin / 60).toFixed(2));
  // Horas extra — decimal para cálculos, HH:MM para mostrar.
  // Si RRHH cargó una corrección manual (hsExtraRelojManual), ESA manda.
  const hsRelojOverride = p.hsExtraRelojManual !== undefined && p.hsExtraRelojManual !== ""
    ? parseFloat(p.hsExtraRelojManual) : null;
  const extraMinFinal   = hsRelojOverride !== null ? Math.round(hsRelojOverride * 60) : totalExtraNetoMin;
  const horasExtraRelojDisplay = minsToDisplay(totalExtraMin); // lo que marcó el reloj (bruto)
  const horasExtra        = parseFloat((extraMinFinal / 60).toFixed(10));
  const horasExtraDisplay = minsToDisplay(extraMinFinal);

  // Manual inputs (stored per employee)
  const sueldoBasico  = parseFloat(p.sueldoBasico  || 0);
  const valorDia      = parseFloat(p.valorDia      || 0);
  const valorHora     = parseFloat(p.valorHora     || 0);
  const valorHoraExt  = parseFloat(p.valorHoraExt  || 0);
  const valorDiaFinde = parseFloat(p.valorDiaFinde || 0);
  // Adelantos: lista de líneas {desc, monto}. Compat con el viejo p.adelanto numérico.
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
  const horasExtraManualHs  = parseFloat(p.horasExtraManualHs  || 0);
  const horasExtraManualImp = parseFloat(p.horasExtraManualImp || 0);
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
    ? parseFloat(p.descDemorasManual) : null;
  const descSalTempManual = p.descSalTempManual !== undefined && p.descSalTempManual !== ""
    ? parseFloat(p.descSalTempManual) : null;
  const descDemoras = descDemorasManual !== null ? descDemorasManual : descDemorasCalc;
  const descSalTemp = descSalTempManual !== null ? descSalTempManual : descSalTempCalc;

  // Ausencias (días de falta) — se descuenta valor día × cantidad de días
  const ausencias     = parseFloat(p.ausencias || 0);
  const descAusencias = valorDia * ausencias;

  // Cantidades a mostrar: si el descuento se borró (importe 0), no mostrar unidades
  const fraccionesDemoraDisp = descDemoras > 0 ? fraccionesDemora  : "—";
  const fraccionesSalTempDisp= descSalTemp > 0 ? fraccionesSalTemp : "—";

  // Calculations — básico manda, adicionales son extras sobre él
  const importeSueldo    = sueldoBasico;
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
  // Premios (importes directos)
  const premioIndividual  = parseFloat(p.premioIndividual  || 0);
  const premioArea        = parseFloat(p.premioArea        || 0);
  const premioPresentismo = parseFloat(p.premioPresentismo || 0);
  const monotributo       = parseFloat(p.monotributo       || 0);
  const totalAdicionales = importeExtras + importeFeriados + sac + importeVacaciones + importeFinde + premioIndividual + premioArea + premioPresentismo + monotributo;
  const totalDescuentos  = descDemoras + descSalTemp + descAusencias;
  const subtotal         = importeSueldo + totalAdicionales;
  const totalACobrar     = subtotal - totalDescuentos - adelanto;
  const reciboA          = parseFloat(p.reciboA || 0);
  const enMano           = totalACobrar - reciboA;

  return {
    desde, hasta, rangeCalcs,
    diasTrabajados, totalDemoraMin, totalSalTempMin, totalExtraMin, totalExtraNetoMin,
    recu: { deudaTotal: recu0.deudaTotal, recuperado: recu0.recuperado, pendiente: recu0.pendiente, saldados: [...recu0.saldados] },
    allFindeInRange, findeSel: [...findeSelSet], diasFinde,
    fraccionesDemora, fraccionesSalTemp, horasSalTemp,
    hsRelojOverride, extraMinFinal, horasExtraRelojDisplay, horasExtra, horasExtraDisplay,
    sueldoBasico, valorDia, valorHora, valorHoraExt, valorDiaFinde,
    adelantos, adelanto, feriados, sac, vacaciones, periodo, ingreso,
    horasExtraManualHs, horasExtraManualImp, importeExtraManual, horasExtraManualDisplay,
    descDemorasCalc, descSalTempCalc, descDemorasManual, descSalTempManual, descDemoras, descSalTemp,
    ausencias, descAusencias, fraccionesDemoraDisp, fraccionesSalTempDisp,
    importeSueldo, impExtrasCalc, impFeriadosCalc, impVacacionesCalc, impFindeCalc,
    impExtrasManual, impFeriadosManual, impVacacionesManual, impFindeManual,
    impExtrasReloj, importeFeriados, importeVacaciones, importeFinde, importeExtras,
    premioIndividual, premioArea, premioPresentismo, monotributo,
    totalAdicionales, totalDescuentos, subtotal, totalACobrar, reciboA, enMano,
  };
}

// Versión serializable para guardar como snapshot al cerrar un período:
// congela todos los importes; rangeCalcs (objetos vivos) queda afuera.
export function snapshotLiquidacion(d) {
  const { rangeCalcs, ...rest } = d;
  return JSON.parse(JSON.stringify(rest));
}