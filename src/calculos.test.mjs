// Tests de la lógica pura de asistencia y liquidación.
// Correr con:  node --test calculos.test.mjs
// (requiere "type": "module" en package.json, que Vite ya trae)
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  parseTimeVal, minsToHHMM, minsToDisplay,
  fraccionesDeUnDia, fraccionesDemoraCalc, fraccionesSalTempCalc,
  calcRecuperacion, calcRecord, calcularLiquidacion, snapshotLiquidacion,
} from "./calculos.js";

/* ── Helpers de tiempo ─────────────────────────────────────────────── */

test("parseTimeVal: formatos de hora", () => {
  assert.equal(parseTimeVal("07:30"), 450);
  assert.equal(parseTimeVal("6:05"), 365);
  assert.equal(parseTimeVal("Ausente"), null);
  assert.equal(parseTimeVal(""), null);
  assert.equal(parseTimeVal(null), null);
  assert.equal(parseTimeVal(0.5), 720);   // fracción de día de Excel = 12:00
  assert.equal(parseTimeVal(450), 450);   // minutos directos
});

test("minsToHHMM / minsToDisplay", () => {
  assert.equal(minsToHHMM(450), "07:30");
  assert.equal(minsToDisplay(90), "1h30");
  assert.equal(minsToDisplay(60), "1h");
  assert.equal(minsToDisplay(45), "45min");
  assert.equal(minsToDisplay(null), "—");
});

/* ── Fracciones de demora (tolerancia 15 por día) ──────────────────── */

test("fraccionesDeUnDia: tolerancia y bloques de 15", () => {
  assert.equal(fraccionesDeUnDia(0), 0);
  assert.equal(fraccionesDeUnDia(15), 0);  // dentro de la tolerancia
  assert.equal(fraccionesDeUnDia(16), 1);
  assert.equal(fraccionesDeUnDia(29), 1);
  // OJO: en múltiplos exactos de 15 el código cuenta el bloque completo
  // (30 → 2, 45 → 3). El comentario del código decía "16-30 = 1";
  // estos tests documentan lo que el código HACE. Ver nota en el chat.
  assert.equal(fraccionesDeUnDia(30), 2);
  assert.equal(fraccionesDeUnDia(31), 2);
  assert.equal(fraccionesDeUnDia(45), 3);
});

test("fraccionesDemoraCalc: la tolerancia es POR DÍA", () => {
  // Tres días con 10 min de demora = 0 fracciones (cada día se perdona);
  // un solo día con 30 min = 2 fracciones.
  assert.equal(fraccionesDemoraCalc([{demora:10},{demora:10},{demora:10}]), 0);
  assert.equal(fraccionesDemoraCalc([{demora:30}]), 2);
});

/* ── Recuperación de horas ─────────────────────────────────────────── */

test("calcRecuperacion: extras saldan deuda en bloques de 30, FIFO", () => {
  const calcs = [
    { id:"a", fecha:"2026-07-01", recuperar:true, recuperarMin:60 },
    { id:"b", fecha:"2026-07-02", extra:100 },  // aplicables: 90 (3 bloques de 30)
  ];
  const r = calcRecuperacion(calcs);
  assert.equal(r.deudaTotal, 60);
  assert.equal(r.recuperado, 60);
  assert.equal(r.pendiente, 0);
  assert.ok(r.saldados.has("a"));
});

test("calcRecuperacion: un día con +7min no descuenta nada", () => {
  const calcs = [
    { id:"a", fecha:"2026-07-01", recuperar:true, recuperarMin:30 },
    { id:"b", fecha:"2026-07-02", extra:7 },
  ];
  const r = calcRecuperacion(calcs);
  assert.equal(r.recuperado, 0);
  assert.equal(r.pendiente, 30);
  assert.equal(r.saldados.size, 0);
});

/* ── calcRecord ────────────────────────────────────────────────────── */

const OPERARIO = { entrada:"06:00", salida:"16:30", tipo:"operario" };
const ADMIN    = { entrada:"08:00", salida:"17:00", tipo:"administrativo" };
// 2026-07-06 = lunes, 2026-07-11 = sábado, 2026-07-12 = domingo

test("calcRecord: día normal de operario con extra y sin demora", () => {
  const c = calcRecord({ fecha:"2026-07-06", entrada:"06:00", salida:"17:30" }, OPERARIO, {});
  assert.equal(c.trabajado, 630);   // 06:00→16:30
  assert.equal(c.jornada, 630);
  assert.equal(c.extra, 60);        // 16:30→17:30
  assert.equal(c.demora, 0);
  assert.equal(c.salTemprana, 0);
});

test("calcRecord: demora y salida temprana", () => {
  const c = calcRecord({ fecha:"2026-07-06", entrada:"06:20", salida:"16:00" }, OPERARIO, {});
  assert.equal(c.demora, 20);
  assert.equal(c.salTemprana, 30);
  assert.equal(c.extra, null);
});

test("calcRecord: operario NO acumula extra antes de las 06:00", () => {
  const c = calcRecord({ fecha:"2026-07-06", entrada:"05:30", salida:"16:30" }, OPERARIO, {});
  assert.equal(c.extra, null);
  assert.equal(c.demora, 0);
});

test("calcRecord: sábado de operario — salida ref 13:00, sin extras", () => {
  const c = calcRecord({ fecha:"2026-07-11", entrada:"06:00", salida:"14:00" }, OPERARIO, {});
  assert.equal(c.jornada, 420);     // 06:00→13:00
  assert.equal(c.salTemprana, 0);   // se fue después de la ref
  assert.equal(c.extra, null);      // sábado no genera extra para operarios
});

test("calcRecord: feriado de operario — salida ref 14:00 y lo demás es extra", () => {
  const sp = { "2026-07-06": { tipo:"feriado" } };
  const c = calcRecord({ fecha:"2026-07-06", entrada:"06:00", salida:"16:00" }, OPERARIO, sp);
  assert.equal(c.extra, 120);       // 14:00→16:00
  assert.equal(c.salTemprana, 0);
});

test("calcRecord: administrativo en finde — todo es extra", () => {
  const c = calcRecord({ fecha:"2026-07-11", entrada:"08:00", salida:"12:00" }, ADMIN, {});
  assert.equal(c.trabajado, 0);
  assert.equal(c.extra, 240);
});

test("calcRecord: extraCorr manual pisa el cálculo", () => {
  const c = calcRecord({ fecha:"2026-07-06", entrada:"06:00", salida:"17:30", extraCorr:30 }, OPERARIO, {});
  assert.equal(c.extra, 30);        // manda la corrección, no los 60 del reloj
});

/* ── calcularLiquidacion ───────────────────────────────────────────── */

const P_BASE = {
  sueldoBasico: "1000000", valorDia: "30000", valorHora: "4000", valorHoraExt: "5000",
  desde: "2026-07-01", hasta: "2026-07-31",
};

test("calcularLiquidacion: caso base — extras suman, demoras descuentan", () => {
  const calcs = [
    { id:"1", fecha:"2026-07-06", trabajado:630, extra:120, demora:0,  salTemprana:0 },
    { id:"2", fecha:"2026-07-07", trabajado:610, extra:0,   demora:20, salTemprana:0 },
  ];
  const d = calcularLiquidacion(P_BASE, calcs);
  assert.equal(d.horasExtra, 2);
  assert.equal(d.impExtrasReloj, 10000);        // 2h × 5000
  assert.equal(d.fraccionesDemora, 1);          // 20 min → 1 fracción
  assert.equal(d.descDemoras, 1000);            // 4000/4 × 1
  assert.equal(d.totalACobrar, 1000000 + 10000 - 1000);
  assert.equal(d.diasTrabajados, 2);
});

test("calcularLiquidacion: el rango desde/hasta filtra registros", () => {
  const calcs = [
    { id:"1", fecha:"2026-06-30", trabajado:630, extra:600, demora:0, salTemprana:0 }, // fuera
    { id:"2", fecha:"2026-07-06", trabajado:630, extra:60,  demora:0, salTemprana:0 }, // dentro
  ];
  const d = calcularLiquidacion(P_BASE, calcs);
  assert.equal(d.horasExtra, 1);
  // y opts pisa el rango de p:
  const d2 = calcularLiquidacion(P_BASE, calcs, { desde:"2026-06-01", hasta:"2026-06-30" });
  assert.equal(d2.horasExtra, 10);
});

test("calcularLiquidacion: overrides manuales mandan", () => {
  const calcs = [{ id:"1", fecha:"2026-07-07", trabajado:610, extra:0, demora:60, salTemprana:0 }];
  const d = calcularLiquidacion({ ...P_BASE, descDemorasManual:"0", impExtrasManual:"99999" }, calcs);
  assert.equal(d.descDemoras, 0);               // descuento borrado a mano
  assert.equal(d.impExtrasReloj, 99999);        // importe manual
});

test("calcularLiquidacion: adelantos en lista y compat con p.adelanto viejo", () => {
  const dLista = calcularLiquidacion({ ...P_BASE, adelantos:[{desc:"a",monto:"100000"},{desc:"b",monto:"50000"}] }, []);
  assert.equal(dLista.adelanto, 150000);
  const dViejo = calcularLiquidacion({ ...P_BASE, adelanto:"80000" }, []);
  assert.equal(dViejo.adelanto, 80000);
});

test("calcularLiquidacion: retiro con recupero SALDADO no se descuenta en plata", () => {
  const calcs = [
    { id:"deuda", fecha:"2026-07-06", trabajado:550, extra:0,  demora:0, salTemprana:60, recuperar:true, recuperarMin:60 },
    { id:"paga",  fecha:"2026-07-07", trabajado:630, extra:70, demora:0, salTemprana:0 },
  ];
  const d = calcularLiquidacion(P_BASE, calcs);
  assert.equal(d.recu.recuperado, 60);
  assert.equal(d.fraccionesSalTemp, 0);          // la deuda se saldó → sin descuento
  assert.equal(d.totalExtraNetoMin, 10);         // 70 − 60 recuperados
});

test("snapshotLiquidacion: serializable y sin rangeCalcs", () => {
  const d = calcularLiquidacion(P_BASE, [{ id:"1", fecha:"2026-07-06", trabajado:630, extra:60, demora:0, salTemprana:0 }]);
  const s = snapshotLiquidacion(d);
  assert.equal(s.rangeCalcs, undefined);
  assert.ok(Array.isArray(s.recu.saldados));
  assert.equal(JSON.parse(JSON.stringify(s)).totalACobrar, d.totalACobrar);
});
