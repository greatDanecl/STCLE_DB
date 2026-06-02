const state = { kpis: [], meta: {}, alerts: [], analytics: {}, filters: { year: 'Todos', month: 'Todos', rank: 'Todos', person: 'Todos' } };
const $ = id => document.getElementById(id);
const fmt = n => Number(n || 0).toLocaleString('es-CL', { maximumFractionDigits: 1 });
const mmName = m => ({'01':'Enero','02':'Febrero','03':'Marzo','04':'Abril','05':'Mayo','06':'Junio','07':'Julio','08':'Agosto','09':'Septiembre','10':'Octubre','11':'Noviembre','12':'Diciembre'}[m] || m);

async function loadJSON(url, fallback) { try { const r = await fetch(url); if (!r.ok) throw new Error(url); return await r.json(); } catch { return fallback; } }
async function init(){
  [state.kpis, state.meta, state.alerts, state.analytics] = await Promise.all([
    loadJSON('data/kpis.json', []), loadJSON('data/metadata.json', {}), loadJSON('data/alerts.json', []), loadJSON('data/analytics.json', {})
  ]);
  $('updated').textContent = `Última actualización: ${state.meta.ultima_actualizacion ? new Date(state.meta.ultima_actualizacion).toLocaleString('es-CL') : 'sin datos'}`;
  setupFilters(); setupTabs(); render();
}
function setupFilters(){
  setOptions('yearFilter', ['Todos', ...(state.meta.anios || [])]);
  setOptions('monthFilter', ['Todos', ...(state.meta.meses || []).map(m => `${m} · ${mmName(m)}`)]);
  setOptions('rankFilter', ['Todos', ...(state.meta.cargos || [])]);
  updatePeople();
  ['yearFilter','monthFilter','rankFilter','personFilter'].forEach(id => $(id).addEventListener('change', () => {
    state.filters.year = $('yearFilter').value;
    state.filters.month = $('monthFilter').value.split(' ')[0];
    state.filters.rank = $('rankFilter').value;
    if(id !== 'personFilter') updatePeople();
    state.filters.person = $('personFilter').value;
    render();
  }));
}
function setOptions(id, arr){ $(id).innerHTML = arr.map(x => `<option value="${escapeAttr(x)}">${escapeHtml(x)}</option>`).join(''); }
function updatePeople(){
  const rows = filteredKpis(false);
  const people = [...new Map(rows.map(r => [r.crew_id, r.nombre])).entries()].sort((a,b) => a[1].localeCompare(b[1]));
  $('personFilter').innerHTML = ['Todos', ...people.map(([id, nombre]) => ({ id, nombre }))]
    .map(x => typeof x === 'string'
      ? `<option value="Todos">Todos</option>`
      : `<option value="${escapeAttr(x.id)}">${escapeHtml(x.nombre)}</option>`
    ).join('');
}
function setupTabs(){ document.querySelectorAll('.tab').forEach(b => b.addEventListener('click', () => { document.querySelectorAll('.tab,.tabpane').forEach(x => x.classList.remove('active')); b.classList.add('active'); $(b.dataset.tab).classList.add('active'); })); }
function filteredKpis(includePerson=true){
  return state.kpis.filter(r => {
    const [y,m] = (r.periodo || '').split('-');
    if(state.filters.year !== 'Todos' && y !== state.filters.year) return false;
    if(state.filters.month !== 'Todos' && m !== state.filters.month) return false;
    if(state.filters.rank !== 'Todos' && r.cargo !== state.filters.rank) return false;
    if(includePerson && state.filters.person !== 'Todos' && r.crew_id !== state.filters.person) return false;
    return true;
  });
}
function activeKpis(){ return filteredKpis(true).filter(r => !r.ausencia_7dias); }
function render(){ renderPersonal(); renderUnion(); renderAlerts(); }
function kpiCard(label, value, sub='', alert=false){ return `<div class="card ${alert?'alert':''}"><div class="label">${label}</div><div class="value">${value}</div><div class="sub">${sub}</div></div>`; }
function renderPersonal(){
  const rows = activeKpis();
  const selected = state.filters.person !== 'Todos';
  const base = selected ? rows : rows.slice(0,1);
  const r = base[0];
  if(!r){ $('personalKpis').innerHTML = `<div class="empty">No hay datos para los filtros seleccionados.</div>`; ['compareBars','activationGauge','psvncRanking','hoursTrend'].forEach(id => $(id).innerHTML=''); return; }
  $('personalKpis').innerHTML = [
    kpiCard('Horas programadas', `${fmt(r.horas_programadas)} h`, `Prom cargo: ${fmt(r.promedio_cargo_programado)} h`),
    kpiCard('Horas efectuadas', `${fmt(r.horas_efectuadas)} h`, `Prom cargo: ${fmt(r.promedio_cargo_efectuado)} h`),
    kpiCard('Utilización', `${fmt(r.utilizacion_pct)}%`, 'Efectuadas / programadas'),
    kpiCard('Descansos', fmt(r.descansos), `Prom cargo: ${fmt(r.promedio_cargo_descansos)}`),
    kpiCard('PSVNC ejecutado', fmt(r.psvnc_ejecutado), 'Máximo mensual: 2', r.psvnc_ejecutado > 2)
  ].join('');
  renderCompare(r); renderGauge(r); renderPSVNCRanking(); renderTrend();
}
function renderCompare(r){
  const items = [
    ['Programadas', r.horas_programadas, r.promedio_cargo_programado], ['Efectuadas', r.horas_efectuadas, r.promedio_cargo_efectuado], ['ASB', r.turnos_asb, r.promedio_cargo_asb], ['HSB', r.turnos_hsb, r.promedio_cargo_hsb], ['Descansos', r.descansos, r.promedio_cargo_descansos]
  ];
  $('compareBars').innerHTML = items.map(([label, val, avg]) => {
    const max = Math.max(val, avg, 1), p1 = Math.min(100, val/max*100), p2 = Math.min(100, avg/max*100);
    return `<div><div class="barrow"><b>${label}</b><div class="track"><div class="fill" style="width:${p1}%"></div></div><span>${fmt(val)}</span></div><div class="barrow mini"><span>Prom cargo</span><div class="track"><div class="fill" style="width:${p2}%;opacity:.45"></div></div><span>${fmt(avg)}</span></div></div>`;
  }).join('');
}
function renderGauge(r){ $('activationGauge').innerHTML = `<div class="num">${fmt(r.ratio_activacion_total)}%</div><div class="arc"></div><p>Activaciones totales: <b>${r.activaciones_total}</b> de <b>${r.turnos_total}</b> turnos programados</p><p class="mini">ASB: ${r.activaciones_asb}/${r.turnos_asb} · HSB: ${r.activaciones_hsb}/${r.turnos_hsb}</p>`; }
function renderPSVNCRanking(){
  const rows = (state.analytics.ranking_psvnc || []).filter(x => matchesFilter(x)).slice(0,10);
  $('psvncRanking').innerHTML = table(rows, ['Nombre','Cargo','Periodo','Prog.','Ejec.'], r => [r.nombre,r.cargo,r.periodo,r.psvnc_programado,r.psvnc_ejecutado]);
}
function renderTrend(){
  const rows = (state.analytics.sindicato_periodos || []).filter(x => (state.filters.rank==='Todos'||x.cargo===state.filters.rank));
  const byP = new Map(); rows.forEach(r => { const o = byP.get(r.periodo) || {periodo:r.periodo,p:0,e:0}; o.p += r.horas_programadas_promedio; o.e += r.horas_efectuadas_promedio; byP.set(r.periodo,o); });
  const data = [...byP.values()].sort((a,b)=>a.periodo.localeCompare(b.periodo));
  $('hoursTrend').innerHTML = lineChart(data.map(d => d.periodo), data.map(d => d.p), data.map(d => d.e));
}
function renderUnion(){
  const rows = activeKpis();
  const union = (state.analytics.sindicato_periodos || []).filter(matchesFilter);
  const dot = union.reduce((a,r)=>a+r.dotacion_activa,0);
  const hp = union.reduce((a,r)=>a+r.horas_programadas_total,0), he = union.reduce((a,r)=>a+r.horas_efectuadas_total,0);
  $('unionKpis').innerHTML = [
    kpiCard('Dotación activa', fmt(dot || new Set(rows.map(r=>r.crew_id)).size), 'Excluye ausencias ≥7 días'),
    kpiCard('Horas programadas', `${fmt(hp)} h`, 'Total filtros'),
    kpiCard('Horas efectuadas', `${fmt(he)} h`, 'Total filtros'),
    kpiCard('Activación promedio', `${fmt(avgOf(union,'activacion_promedio_pct'))}%`, 'ASB + HSB'),
    kpiCard('PSVNC ejecutado', fmt(union.reduce((a,r)=>a+r.psvnc_ejecutado_total,0)), 'Total instancias')
  ].join('');
  $('unionTable').innerHTML = table(union, ['Periodo','Cargo','Dotación','H prog prom','H ejec prom','Act %','PSVNC ejec'], r => [r.periodo,r.cargo,r.dotacion_activa,fmt(r.horas_programadas_promedio),fmt(r.horas_efectuadas_promedio),fmt(r.activacion_promedio_pct)+'%',r.psvnc_ejecutado_total]);
  renderHeatmap(); renderTopTurns(); renderRests();
}
function renderHeatmap(){
  const rows = (state.analytics.heatmap || []).filter(matchesFilter).slice(-35);
  if(!rows.length){ $('heatmap').innerHTML = '<div class="empty">Sin datos de heatmap para el filtro.</div>'; return; }
  const max = Math.max(...rows.map(r=>r.valor),1);
  $('heatmap').innerHTML = `<div class="heatgrid">${rows.map(r => `<div class="heatcell" title="${r.fecha}: ${r.valor}" style="background:rgba(20,85,163,${.12 + .75*r.valor/max})">${r.fecha.slice(8)}</div>`).join('')}</div>`;
}
function renderTopTurns(){ const rows = (state.analytics.top_turnos || []).filter(matchesFilter).slice(0,10); $('topTurns').innerHTML = table(rows, ['Nombre','Cargo','Periodo','ASB','HSB','Total'], r => [r.nombre,r.cargo,r.periodo,r.turnos_asb,r.turnos_hsb,r.turnos_total]); }
function renderRests(){ const rows = (state.analytics.distribucion_descansos || []).filter(matchesFilter).slice(0,10); $('restsDistribution').innerHTML = table(rows, ['Nombre','Cargo','Periodo','Descansos','B'], r => [r.nombre,r.cargo,r.periodo,r.descansos,r.blancos]); }
function renderAlerts(){
  const rows = state.alerts.filter(matchesFilter);
  $('alertsList').innerHTML = table(rows, ['Tipo','Periodo','Nombre','Cargo','Valor'], r => [`<span class="pill ${r.tipo.includes('PSVNC')?'red':'blue'}">${r.tipo}</span>`, r.periodo, r.nombre, r.cargo, r.valor ?? '']);
}
function matchesFilter(r){
  const [y,m] = (r.periodo || '').split('-');
  if(state.filters.year !== 'Todos' && y !== state.filters.year) return false;
  if(state.filters.month !== 'Todos' && m !== state.filters.month) return false;
  if(state.filters.rank !== 'Todos' && r.cargo !== state.filters.rank) return false;
  if(state.filters.person !== 'Todos' && r.crew_id && r.crew_id !== state.filters.person) return false;
  return true;
}
function avgOf(rows, field){ return rows.length ? rows.reduce((a,r)=>a+Number(r[field]||0),0)/rows.length : 0; }
function table(rows, headers, mapper){ if(!rows.length) return '<div class="empty">Sin datos para mostrar.</div>'; return `<table><thead><tr>${headers.map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${mapper(r).map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table>`; }
function lineChart(labels, a, b){
  if(!labels.length) return '<div class="empty">Sin datos de evolución.</div>';
  const W=640,H=240,P=35,max=Math.max(...a,...b,1); const x=i=>P+i*((W-2*P)/Math.max(labels.length-1,1)); const y=v=>H-P-(v/max)*(H-2*P);
  const path = vals => vals.map((v,i)=>`${i?'L':'M'}${x(i)},${y(v)}`).join(' ');
  return `<svg class="svgchart" viewBox="0 0 ${W} ${H}"><path d="${path(a)}" fill="none" stroke="#1455a3" stroke-width="4"/><path d="${path(b)}" fill="none" stroke="#c71f37" stroke-width="4"/><text x="35" y="20" fill="#1455a3">Programado</text><text x="150" y="20" fill="#c71f37">Ejecutado</text>${labels.map((l,i)=>`<text x="${x(i)-15}" y="${H-8}" font-size="10" fill="#6b7280">${l.slice(5)}</text>`).join('')}</svg>`;
}
function escapeHtml(x){ return String(x).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }
function escapeAttr(x){ return escapeHtml(x); }
init();
