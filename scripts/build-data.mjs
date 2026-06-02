import fs from "fs";
import path from "path";
import XLSX from "xlsx";

const RAW_DIR = "data/raw";
const OUT_DIR = "data";
const DETAIL_DIR = "data/detail";

const ABSENCE_CODES = new Set(["SICK", "VAC", "OOF"]);
const ADMIN_CODES = new Set(["ADM","ASCCBT","CLA","CRM","CRS","DIT","EMG","IET","MCK","MTC","MTE","MTI","MTU","SVC"]);
const REST_CODES = new Set(["B", "BB", "BL", "OFF", "DO"]);

function ensureDir(dir) { if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true }); }
function s(v) { return String(v ?? "").trim(); }
function u(v) { return s(v).toUpperCase(); }
function pad(n) { return String(n).padStart(2, "0"); }
function localYMD(d) { return d ? `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}` : ""; }
function localHM(d) { return d ? `${pad(d.getHours())}:${pad(d.getMinutes())}` : ""; }
function ym(d) { return d ? `${d.getFullYear()}-${pad(d.getMonth()+1)}` : ""; }
function round(n, dec=2) { const p = 10 ** dec; return Math.round(Number(n || 0) * p) / p; }
function normalizeCrewId(v) { const raw = s(v); if (!raw) return ""; return raw.replace(/^0+(?=\d)/, ""); }
function get(row, names) { for (const name of names) if (row[name] !== undefined && row[name] !== null && row[name] !== "") return row[name]; return ""; }

function toDate(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date && !isNaN(value)) return value;
  if (typeof value === "number") {
    const p = XLSX.SSF.parse_date_code(value);
    return p ? new Date(p.y, p.m - 1, p.d) : null;
  }
  let txt = u(value);
  const monthMap = { ENE:"JAN", ABR:"APR", AGO:"AUG", DIC:"DEC" };
  for (const [es, en] of Object.entries(monthMap)) txt = txt.replace(es, en);
  const d = new Date(txt);
  return isNaN(d) ? null : d;
}

function toTime(value) {
  if (!value && value !== 0) return "";
  if (typeof value === "number") {
    const total = Math.round(value * 24 * 60);
    return `${pad(Math.floor(total / 60) % 24)}:${pad(total % 60)}`;
  }
  const txt = s(value);
  const m = txt.match(/^(\d{1,2}):(\d{2})/);
  return m ? `${pad(m[1])}:${m[2]}` : txt;
}

function combineDateTime(dateValue, timeValue) {
  const d = toDate(dateValue);
  if (!d) return null;
  const t = toTime(timeValue);
  const [hh, mm] = t.includes(":") ? t.split(":").map(Number) : [0, 0];
  const out = new Date(d);
  out.setHours(hh || 0, mm || 0, 0, 0);
  return out;
}

function blockHours(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return round(v);
  const txt = s(v);
  if (txt.includes(":")) {
    const [h, m] = txt.split(":").map(Number);
    return round((h || 0) + (m || 0) / 60);
  }
  const n = Number(txt.replace(",", "."));
  return Number.isFinite(n) ? round(n) : 0;
}

function fleet(v) {
  const code = u(v);
  if (code === "787") return "Wide Body";
  if (code === "320" || code === "321" || code.startsWith("32")) return "Narrow Body";
  return code;
}

function roleType(row, fileName) {
  const direct = u(get(row, ["tipo_rol", "Tipo Rol", "tipo rol"]));
  if (direct.includes("EJEC")) return "Ejecutado";
  if (direct.includes("PUB")) return "Publicado";
  const f = u(fileName);
  if (f.includes("EFECT") || f.includes("EJEC")) return "Ejecutado";
  return "Publicado";
}

function isFlight(code) { return /^LA\d+/i.test(code); }
function isASB(code) { return /^ASB\d*$/i.test(code); }
function isHSB(code) { return /^HSB\d*$/i.test(code); }
function isRest(code) { return REST_CODES.has(code); }

function normalizeRow(row, fileName) {
  const start = combineDateTime(get(row, ["Str Dt", "str_dt", "start_date", "fecha_inicio"]), get(row, ["Str Tm", "str_tm", "start_time", "hora_inicio"]));
  let end = combineDateTime(get(row, ["End Dt", "end_dt", "end_date", "fecha_fin"]), get(row, ["End Tm", "end_tm", "end_time", "hora_fin"]));
  if (start && end && end < start) end = new Date(end.getTime() + 86400000);

  const periodoDate = toDate(get(row, ["periodo", "Periodo", "PERIODO"]));
  const actividad = u(get(row, ["Activity", "activity_code", "Code IFN", "code_ifn", "actividad"]));
  const cargo = u(get(row, ["Rank", "rank_code", "cargo"]));
  const tipo = roleType(row, fileName);

  return {
    crew_id: normalizeCrewId(get(row, ["Staff Num", "crew_id", "Crew ID", "id"])),
    nombre: s(get(row, ["Nombre completo", "nombre_completo", "Nombre", "name"])),
    cargo,
    periodo: periodoDate ? ym(periodoDate) : ym(start),
    tipo_rol: tipo,
    actividad,
    fecha_inicio: localYMD(start),
    hora_inicio: localHM(start),
    fecha_fin: localYMD(end),
    hora_fin: localHM(end),
    start_ms: start ? start.getTime() : null,
    end_ms: end ? end.getTime() : null,
    fleet: fleet(get(row, ["Fleet", "aircraft_type_desc", "fleet_raw"])),
    block_hours: blockHours(get(row, ["Block Time", "block_time", "block_hours"])),
    es_vuelo: isFlight(actividad),
    es_descanso: isRest(actividad),
    es_turno_aeropuerto: isASB(actividad),
    es_turno_casa: isHSB(actividad),
    es_ausencia: ABSENCE_CODES.has(actividad),
    es_admin_capacitacion: ADMIN_CODES.has(actividad),
    source_file: fileName
  };
}

function readAllRows() {
  const files = fs.existsSync(RAW_DIR) ? fs.readdirSync(RAW_DIR).filter(f => f.toLowerCase().endsWith(".xlsx")) : [];
  const rows = [];
  for (const file of files) {
    const full = path.join(RAW_DIR, file);
    const workbook = XLSX.readFile(full, { cellDates: true });
    for (const sheetName of workbook.SheetNames) {
      const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
      if (!json.length) continue;
      for (const row of json) {
        const r = normalizeRow(row, file);
        if ((!r.crew_id && !r.nombre) || !r.periodo || !r.cargo) continue;
        rows.push(r);
      }
    }
  }
  return rows;
}

function groupBy(rows, fn) { const m = new Map(); for (const r of rows) { const k = fn(r); if (!m.has(k)) m.set(k, []); m.get(k).push(r); } return m; }
function sum(rows, fn) { return round(rows.reduce((a, r) => a + Number(fn(r) || 0), 0)); }
function count(rows, fn) { return rows.filter(fn).length; }
function avg(rows, field) { return rows.length ? round(rows.reduce((a, r) => a + Number(r[field] || 0), 0) / rows.length) : 0; }
function uniq(arr) { return [...new Set(arr.filter(Boolean))]; }
function daysWith(rows, fn) { return new Set(rows.filter(fn).map(r => r.fecha_inicio).filter(Boolean)).size; }
function sameDay(a, b) { return a.fecha_inicio && b.fecha_inicio && a.fecha_inicio === b.fecha_inicio; }

function countActivations(rows, turnoFn) {
  const published = rows.filter(r => r.tipo_rol === "Publicado" && turnoFn(r));
  const executed = rows.filter(r => r.tipo_rol === "Ejecutado");
  let activations = 0;
  for (const p of published) {
    const converted = executed.some(e => sameDay(p, e) && e.actividad !== p.actividad && !turnoFn(e));
    if (converted) activations++;
  }
  return activations;
}

function touchesNightWindow(row) {
  if (!row.es_vuelo || !row.start_ms || !row.end_ms) return false;
  const start = new Date(row.start_ms);
  const end = new Date(row.end_ms);
  const cursor = new Date(start);
  cursor.setHours(0,0,0,0);
  while (cursor <= end) {
    const winStart = new Date(cursor); winStart.setHours(0,30,0,0);
    const winEnd = new Date(cursor); winEnd.setHours(5,30,0,0);
    if (start < winEnd && end > winStart) return true;
    cursor.setDate(cursor.getDate() + 1);
  }
  return false;
}

function countPSVNC(rows, tipoRol) {
  const days = uniq(rows.filter(r => r.tipo_rol === tipoRol && touchesNightWindow(r)).map(r => r.fecha_inicio)).sort();
  let instances = 0;
  for (let i = 1; i < days.length; i++) {
    const prev = new Date(`${days[i - 1]}T00:00:00`);
    const curr = new Date(`${days[i]}T00:00:00`);
    if (Math.round((curr - prev) / 86400000) === 1) instances++;
  }
  return instances;
}

function buildKpis(rows) {
  const groups = groupBy(rows, r => `${r.periodo}|${r.cargo}|${r.crew_id}`);
  const out = [];
  for (const [key, g] of groups) {
    const [periodo, cargo, crew_id] = key.split("|");
    const nombre = g.find(r => r.nombre)?.nombre || "";
    const pub = g.filter(r => r.tipo_rol === "Publicado");
    const eje = g.filter(r => r.tipo_rol === "Ejecutado");
    const ausenciaDias = daysWith(g, r => r.es_ausencia);
    const diasTrabPub = daysWith(pub, r => !r.es_descanso && !r.es_ausencia);
    const diasTrabEje = daysWith(eje, r => !r.es_descanso && !r.es_ausencia);
    const asb = count(pub, r => r.es_turno_aeropuerto);
    const hsb = count(pub, r => r.es_turno_casa);
    const actASB = countActivations(g, r => r.es_turno_aeropuerto);
    const actHSB = countActivations(g, r => r.es_turno_casa);
    const hp = sum(pub, r => r.block_hours);
    const he = sum(eje, r => r.block_hours);
    out.push({
      periodo, cargo, crew_id, nombre,
      horas_programadas: hp,
      horas_efectuadas: he,
      utilizacion_pct: hp ? round(he / hp * 100, 1) : 0,
      dias_trabajados_programado: diasTrabPub,
      dias_trabajados_ejecutado: diasTrabEje,
      horas_promedio_dia_programado: diasTrabPub ? round(hp / diasTrabPub) : 0,
      horas_promedio_dia_ejecutado: diasTrabEje ? round(he / diasTrabEje) : 0,
      descansos: count(pub, r => r.es_descanso),
      blancos: count(pub, r => r.actividad === "B"),
      turnos_asb: asb,
      turnos_hsb: hsb,
      turnos_total: asb + hsb,
      activaciones_asb: actASB,
      activaciones_hsb: actHSB,
      activaciones_total: actASB + actHSB,
      ratio_activacion_asb: asb ? round(actASB / asb * 100, 1) : 0,
      ratio_activacion_hsb: hsb ? round(actHSB / hsb * 100, 1) : 0,
      ratio_activacion_total: (asb + hsb) ? round((actASB + actHSB) / (asb + hsb) * 100, 1) : 0,
      psvnc_programado: countPSVNC(g, "Publicado"),
      psvnc_ejecutado: countPSVNC(g, "Ejecutado"),
      ausencia_dias: ausenciaDias,
      ausencia_7dias: ausenciaDias >= 7
    });
  }
  const cargoMonth = groupBy(out.filter(r => !r.ausencia_7dias), r => `${r.periodo}|${r.cargo}`);
  for (const r of out) {
    const cg = cargoMonth.get(`${r.periodo}|${r.cargo}`) || [];
    r.promedio_cargo_programado = avg(cg, "horas_programadas");
    r.promedio_cargo_efectuado = avg(cg, "horas_efectuadas");
    r.promedio_cargo_descansos = avg(cg, "descansos");
    r.promedio_cargo_blancos = avg(cg, "blancos");
    r.promedio_cargo_asb = avg(cg, "turnos_asb");
    r.promedio_cargo_hsb = avg(cg, "turnos_hsb");
    r.promedio_cargo_turnos = avg(cg, "turnos_total");
    r.promedio_cargo_activacion = avg(cg, "ratio_activacion_total");
    r.promedio_cargo_psvnc_programado = avg(cg, "psvnc_programado");
    r.promedio_cargo_psvnc_ejecutado = avg(cg, "psvnc_ejecutado");
  }
  return out.sort((a,b) => a.periodo.localeCompare(b.periodo) || a.cargo.localeCompare(b.cargo) || a.nombre.localeCompare(b.nombre));
}

function buildAlerts(kpis) {
  const alerts = [];
  for (const k of kpis) {
    if (k.psvnc_programado > 1) alerts.push({ tipo:"PSVNC_PROGRAMADO_EXCESIVO", periodo:k.periodo, cargo:k.cargo, crew_id:k.crew_id, nombre:k.nombre, valor:k.psvnc_programado });
    if (k.psvnc_ejecutado > 2) alerts.push({ tipo:"PSVNC_EJECUTADO_EXCESIVO", periodo:k.periodo, cargo:k.cargo, crew_id:k.crew_id, nombre:k.nombre, valor:k.psvnc_ejecutado });
    if (k.ausencia_7dias) alerts.push({ tipo:"AUSENCIA_MAYOR_7_DIAS", periodo:k.periodo, cargo:k.cargo, crew_id:k.crew_id, nombre:k.nombre, valor:k.ausencia_dias });
  }
  return alerts;
}

function buildAnalytics(rows, kpis) {
  const active = kpis.filter(k => !k.ausencia_7dias);
  const byPeriodCargo = groupBy(active, k => `${k.periodo}|${k.cargo}`);
  const sindicato_periodos = [];
  for (const [key, g] of byPeriodCargo) {
    const [periodo, cargo] = key.split("|");
    sindicato_periodos.push({
      periodo, cargo,
      dotacion_activa: g.length,
      horas_programadas_total: sum(g, r => r.horas_programadas),
      horas_efectuadas_total: sum(g, r => r.horas_efectuadas),
      horas_programadas_promedio: avg(g, "horas_programadas"),
      horas_efectuadas_promedio: avg(g, "horas_efectuadas"),
      utilizacion_promedio_pct: avg(g, "utilizacion_pct"),
      descansos_promedio: avg(g, "descansos"),
      blancos_promedio: avg(g, "blancos"),
      turnos_promedio: avg(g, "turnos_total"),
      activacion_promedio_pct: avg(g, "ratio_activacion_total"),
      personas_psvnc_programado: g.filter(r => r.psvnc_programado > 0).length,
      personas_psvnc_ejecutado: g.filter(r => r.psvnc_ejecutado > 0).length,
      psvnc_programado_total: sum(g, r => r.psvnc_programado),
      psvnc_ejecutado_total: sum(g, r => r.psvnc_ejecutado)
    });
  }
  sindicato_periodos.sort((a,b) => a.periodo.localeCompare(b.periodo) || a.cargo.localeCompare(b.cargo));

  const ranking_psvnc = [...active].filter(k => k.psvnc_programado || k.psvnc_ejecutado)
    .sort((a,b) => (b.psvnc_ejecutado + b.psvnc_programado) - (a.psvnc_ejecutado + a.psvnc_programado))
    .slice(0, 100)
    .map(k => ({ periodo:k.periodo, cargo:k.cargo, crew_id:k.crew_id, nombre:k.nombre, psvnc_programado:k.psvnc_programado, psvnc_ejecutado:k.psvnc_ejecutado }));

  const top_turnos = [...active].sort((a,b) => b.turnos_total - a.turnos_total).slice(0,100)
    .map(k => ({ periodo:k.periodo, cargo:k.cargo, crew_id:k.crew_id, nombre:k.nombre, turnos_asb:k.turnos_asb, turnos_hsb:k.turnos_hsb, turnos_total:k.turnos_total }));

  const blancos = [...active].sort((a,b) => b.descansos - a.descansos).slice(0,100)
    .map(k => ({ periodo:k.periodo, cargo:k.cargo, crew_id:k.crew_id, nombre:k.nombre, descansos:k.descansos, blancos:k.blancos }));

  const heat = new Map();
  for (const r of rows.filter(r => r.tipo_rol === "Ejecutado" && !r.es_descanso && !r.es_ausencia && r.fecha_inicio)) {
    const key = `${r.periodo}|${r.cargo}|${r.fecha_inicio}`;
    heat.set(key, (heat.get(key) || 0) + 1);
  }
  const heatmap = [...heat.entries()].map(([key, valor]) => { const [periodo,cargo,fecha] = key.split("|"); return { periodo, cargo, fecha, valor }; })
    .sort((a,b) => a.fecha.localeCompare(b.fecha));

  return { sindicato_periodos, ranking_psvnc, top_turnos, distribucion_descansos: blancos, heatmap };
}

function buildMetadata(rows, kpis, alerts) {
  const periodos = uniq(rows.map(r => r.periodo)).sort();
  const cargos = uniq(rows.map(r => r.cargo)).sort();
  const personas = uniq(kpis.map(k => `${k.crew_id}|${k.nombre}|${k.cargo}`)).map(x => { const [crew_id,nombre,cargo] = x.split("|"); return { crew_id, nombre, cargo }; }).sort((a,b) => a.nombre.localeCompare(b.nombre));
  return {
    ultima_actualizacion: new Date().toISOString(),
    periodos,
    anios: uniq(periodos.map(p => p.slice(0,4))).sort(),
    meses: uniq(periodos.map(p => p.slice(5,7))).sort(),
    cargos,
    personas,
    total_tripulantes: uniq(rows.map(r => r.crew_id)).length,
    registros_fuente: rows.length,
    registros_kpi: kpis.length,
    total_alertas: alerts.length,
    modelo: "compact-v3-sin-dashboard-json"
  };
}

function writeDetailByPeriod(rows) {
  ensureDir(DETAIL_DIR);
  for (const f of fs.readdirSync(DETAIL_DIR).filter(f => f.endsWith(".json"))) fs.unlinkSync(path.join(DETAIL_DIR, f));
  const byPeriod = groupBy(rows, r => r.periodo);
  for (const [periodo, g] of byPeriod) {
    const compact = g.map(r => ({
      id:r.crew_id, n:r.nombre, c:r.cargo, p:r.periodo, t:r.tipo_rol, a:r.actividad,
      sd:r.fecha_inicio, st:r.hora_inicio, ed:r.fecha_fin, et:r.hora_fin, b:r.block_hours,
      fl:r.fleet, v:r.es_vuelo, d:r.es_descanso, asb:r.es_turno_aeropuerto, hsb:r.es_turno_casa, abs:r.es_ausencia
    }));
    fs.writeFileSync(path.join(DETAIL_DIR, `${periodo}.json`), JSON.stringify(compact));
  }
}

ensureDir(OUT_DIR); ensureDir(DETAIL_DIR);
const rows = readAllRows();
const kpis = buildKpis(rows);
const alerts = buildAlerts(kpis);
const analytics = buildAnalytics(rows, kpis);
const metadata = buildMetadata(rows, kpis, alerts);
if (process.env.GENERATE_DETAIL_JSON === "true") {
  writeDetailByPeriod(rows);
} else {
  ensureDir(DETAIL_DIR);
  for (const f of fs.readdirSync(DETAIL_DIR).filter(f => f.endsWith(".json"))) fs.unlinkSync(path.join(DETAIL_DIR, f));
}
fs.writeFileSync(path.join(OUT_DIR, "kpis.json"), JSON.stringify(kpis));
fs.writeFileSync(path.join(OUT_DIR, "alerts.json"), JSON.stringify(alerts));
fs.writeFileSync(path.join(OUT_DIR, "analytics.json"), JSON.stringify(analytics));
fs.writeFileSync(path.join(OUT_DIR, "metadata.json"), JSON.stringify(metadata));
console.log("JSON generados correctamente sin dashboard.json pesado");
console.log(`rows fuente: ${rows.length}`);
console.log(`kpis: ${kpis.length}`);
console.log(`alerts: ${alerts.length}`);
console.log(`detail JSON: ${process.env.GENERATE_DETAIL_JSON === "true" ? metadata.periodos.length + " periodos" : "omitido"}`);
