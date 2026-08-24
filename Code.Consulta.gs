// ══════════════════════════════════════════════════════════════════
// CONSULTA·QM — Code.gs
//
// Portal público de consulta de notas (sin contraseña, por código de
// estudiante). Este archivo se guarda aquí para tener control de
// versiones del backend real — hasta ahora solo vivía pegado en el
// editor de Apps Script, sin historial ni forma de revertir un cambio.
//
// Para desplegar un cambio de este archivo: pégalo completo en
// Extensiones → Apps Script del proyecto de Consulta y vuelve a
// "Implementar → Gestionar implementaciones → editar → Nueva versión".
// ══════════════════════════════════════════════════════════════════

const SPREADSHEET_ID_QRY = '1FzUF-Rnu_6RB3exJXbtTS3QHyMGz-59vhomP3EXSyGE';

const _CURSOS = ['1004','1005','1006','1101','1102','1103','1104'];

const _SH_STU  = 'Estudiantes';
const _SH_GRD  = 'Notas';
const _SH_SP   = 'Especiales';
const _SH_CFG  = 'Config';

function doGet(e) {
  return HtmlService
    .createHtmlOutput(buildHTML())
    .setTitle('Consulta de Notas · Química')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _calcPesos(specials) {
  let pesoAct = null, pesoFin = null;
  for (let i = specials.length-1; i >= 0; i--) {
    if (specials[i].pesoAct != null && specials[i].pesoFinal != null) {
      pesoAct = specials[i].pesoAct;
      pesoFin = specials[i].pesoFinal;
      break;
    }
  }
  if (pesoAct === null || pesoFin === null) {
    const sp  = Math.round(specials.reduce((s,e)=>s+e.weight,0)*10)/10;
    const fA  = Math.min(sp,60), fF = Math.max(0,sp-60);
    pesoAct = Math.round((60-fA)*10)/10;
    pesoFin = Math.round((30-fF)*10)/10;
  }
  return { actividades: pesoAct, autoeval: 5, coeval: 5, final: pesoFin };
}

function consultarEstudiante(code) {
  if (!code || !code.toString().trim()) return { error: 'empty' };
  code = code.toString().trim();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_QRY);

  const stuSh = ss.getSheetByName(_SH_STU);
  if (!stuSh) return { error: 'config' };
  const stuRows = stuSh.getDataRange().getValues().slice(1);

  // ✅ FIX: busca en TODOS los cursos válidos, no en el derivado del código.
  // Esto permite que estudiantes trasladados conserven su código original.
  const match = stuRows.find(r =>
    _CURSOS.includes(r[0].toString()) &&
    r[4] == true &&
    r[3].toString().trim().toLowerCase() === code.toLowerCase()
  );
  if (!match) return { error: 'notfound' };

  const curso     = match[0].toString(); // curso real según el registro
  const studentId = match[1];
  const nombre    = match[2];

  const cfgSh   = ss.getSheetByName(_SH_CFG);
  const cfgRows = cfgSh ? cfgSh.getDataRange().getValues().slice(1) : [];
  const cfgRow  = cfgRows.find(r => r[0].toString() === curso);
  const activePeriod = cfgRow ? +cfgRow[1] : 1;

  const spSh  = ss.getSheetByName(_SH_SP);
  const grdSh = ss.getSheetByName(_SH_GRD);
  const allGrd = grdSh ? grdSh.getDataRange().getValues().slice(1)
    .filter(r => r[0].toString()===curso) : [];

  const periodosSet = new Set(allGrd.filter(r=>r[4]===studentId).map(r=>+r[1]));
  periodosSet.add(activePeriod);

  const periodosData = [];

  [1,2,3].forEach(period => {
    if (!periodosSet.has(period)) return;

    const spRows = spSh ? spSh.getDataRange().getValues().slice(1)
      .filter(r => r[0].toString()===curso && +r[1]===period)
      .map(r => ({
        id:String(r[2]), name:String(r[3]), weight:+r[4],
        pesoAct:   (r[5]!==''&&r[5]!==undefined&&r[5]!==null) ? +r[5] : null,
        pesoFinal: (r[6]!==''&&r[6]!==undefined&&r[6]!==null) ? +r[6] : null
      })) : [];
    const weights = _calcPesos(spRows);

    const grdRows = allGrd.filter(r => +r[1]===period);

    const actNames = [];
    const actSeen = new Set();
    grdRows.filter(r=>r[2]==='actividad').forEach(r=>{
      const n = String(r[3]);
      if (!actSeen.has(n) && n !== '__col__') { actSeen.add(n); actNames.push(n); }
    });

    // ✅ CORREGIDO: distingue celda vacía de nota 0
    const gmap = {};
    grdRows.filter(r => String(r[4])===studentId)
           .forEach(r => {
             const val = r[5];
             if (val !== '' && val !== null && val !== undefined) {
               gmap[r[2]+'__'+r[3]] = +val;
             }
           });

    const actItems = actNames.map(n => ({
      name: n,
      grade: gmap['actividad__'+n] !== undefined ? gmap['actividad__'+n] : null
    }));
    const actVals = actItems.map(a=>a.grade).filter(g=>g!==null);
    const actAvg  = actVals.length
      ? Math.round(actVals.reduce((a,b)=>a+b,0)/actVals.length*10)/10 : null;

    const componentes = [
      { key:'actividades', label:'Actividades de clase', weight:weights.actividades,
        grade:actAvg, actItems },
      { key:'autoeval', label:'Autoevaluación', weight:5,
        grade: gmap['autoeval___']!==undefined ? gmap['autoeval___'] : null },
      { key:'coeval',   label:'Coevaluación',   weight:5,
        grade: gmap['coeval___']!==undefined   ? gmap['coeval___']   : null },
      { key:'final',    label:'Evaluación final', weight:weights.final,
        grade: gmap['final___']!==undefined    ? gmap['final___']    : null }
    ];
    spRows.forEach(sp => {
      componentes.push({ key:'especial_'+sp.id, label:sp.name, weight:sp.weight,
        grade: gmap['especial__'+sp.id]!==undefined ? gmap['especial__'+sp.id] : null });
    });

    const totalPct = componentes.reduce((s,c)=>s+c.weight, 0);
    let sumW = 0;
    // ✅ CORREGIDO: el 0 real sí cuenta para la definitiva
    componentes.forEach(c => { sumW += (c.grade !== null ? c.grade : 0) * (c.weight / 100); });
    let definitiva = totalPct>0 ? Math.round(sumW/(totalPct/100)*100)/100 : null;

    // ✅ NUEVO: nivelación — mismo mecanismo genérico de saveGrade que usa el
    // resto de la app (component:'nivelacion', itemId:'_'), sin sheet nueva.
    // Si el docente diligenció una nota de nivelación para este periodo, la
    // definitiva que ve el estudiante/acudiente se fija en 3.0 y se marca el
    // periodo como nivelado — el valor real de la nivelación queda solo para
    // registro interno del docente, no se muestra aquí.
    const nivelacionGrade = gmap['nivelacion___'] !== undefined ? gmap['nivelacion___'] : null;
    const nivelado = nivelacionGrade !== null;
    if (nivelado) definitiva = 3.0;

    periodosData.push({
      period, componentes, definitiva, totalPct,
      actCount: actNames.length, isActive: period===activePeriod,
      nivelado
    });
  });

  const definitivasAno = periodosData
    .filter(p => p.definitiva !== null)
    .map(p => p.definitiva);
  const definitivaAnual = definitivasAno.length > 0
    ? Math.round(definitivasAno.reduce((a,b)=>a+b,0) / definitivasAno.length * 100) / 100
    : null;

  return { ok:true, nombre, course:curso, activePeriod, code, periodos:periodosData, definitivaAnual };
}


// ════════════════════════════════════════════════════════════════
// HTML
// ════════════════════════════════════════════════════════════════
function buildHTML() {
  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Consulta de Notas</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{
  font-family:'Inter',system-ui,sans-serif;
  background:#f0f4f8;
  color:#1a2332;
  min-height:100vh;
  padding-bottom:32px;
}
.header{
  background:#fff;
  border-bottom:3px solid #1a6b5a;
  padding:14px 20px;
  display:flex;align-items:center;gap:12px;
  margin-bottom:24px;
}
.header-icon{font-size:1.6rem}
.header-title{font-size:1rem;font-weight:700;color:#1a2332;letter-spacing:-.3px}
.header-sub{font-size:.7rem;color:#6b7a8d;margin-top:1px}
.card{
  background:#fff;
  border-radius:12px;
  box-shadow:0 2px 12px rgba(0,0,0,.08);
  padding:24px;
  max-width:560px;
  margin:0 auto 20px;
}
.card-title{
  font-size:1.15rem;font-weight:700;color:#1a6b5a;
  text-align:center;margin-bottom:20px;
  padding-bottom:12px;
  border-bottom:2px solid #1a6b5a;
  display:flex;align-items:center;justify-content:center;gap:8px;
}
.form-row{
  display:flex;align-items:center;gap:10px;
  margin-bottom:6px;flex-wrap:wrap;
}
.form-label{
  font-size:.85rem;font-weight:600;color:#1a2332;
  white-space:nowrap;min-width:140px;
}
.form-input{
  flex:1;min-width:140px;
  border:1.5px solid #c8d5e0;border-radius:8px;
  padding:10px 13px;font-size:.9rem;font-family:'Inter',sans-serif;
  outline:none;transition:border-color .2s;
  background:#f8fafc;color:#1a2332;
}
.form-input:focus{border-color:#1a6b5a;background:#fff}
.form-input::placeholder{color:#9baab8}
.hint-text{font-size:.72rem;color:#6b7a8d;margin-bottom:14px;padding-left:150px}
.search-btn{
  display:block;width:100%;
  background:#1a6b5a;color:#fff;border:none;border-radius:8px;
  padding:12px;font-family:'Inter',sans-serif;font-size:.9rem;font-weight:600;
  cursor:pointer;transition:background .2s;letter-spacing:.2px;
}
.search-btn:hover{background:#145548}
.search-btn:active{background:#0f3f35}
.err-msg{
  display:none;margin-top:12px;
  background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;
  padding:10px 13px;font-size:.82rem;color:#b91c1c;
}
.err-msg.show{display:flex;align-items:center;gap:7px}
.loading{
  text-align:center;padding:28px;color:#6b7a8d;font-size:.85rem;
}
.spinner{
  width:20px;height:20px;
  border:2.5px solid #d1e0d8;border-top-color:#1a6b5a;
  border-radius:50%;animation:spin .7s linear infinite;
  margin:0 auto 10px;
}
@keyframes spin{to{transform:rotate(360deg)}}
.result-wrap{max-width:900px;margin:0 auto;padding:0 12px}
.result-header{
  background:#fff;border-radius:12px;
  box-shadow:0 2px 12px rgba(0,0,0,.08);
  padding:16px 20px;margin-bottom:14px;
  border-left:4px solid #1a6b5a;
}
.found-label{
  font-size:.68rem;font-weight:600;letter-spacing:1.2px;
  text-transform:uppercase;color:#1a6b5a;margin-bottom:4px;
  display:flex;align-items:center;gap:5px;
}
.student-name{font-size:1.15rem;font-weight:700;color:#1a2332;letter-spacing:-.3px;}
.student-meta{font-size:.76rem;color:#6b7a8d;margin-top:3px;}
.course-block{
  background:#fff;border-radius:12px;
  box-shadow:0 2px 12px rgba(0,0,0,.08);
  overflow:hidden;margin-bottom:16px;
  padding:16px 16px 0;
}
.course-head{
  padding:12px 18px;border-bottom:2px solid #1a6b5a;
  display:flex;align-items:center;justify-content:space-between;
  flex-wrap:wrap;gap:8px;
}
.course-num{
  font-size:1rem;font-weight:700;color:#1a6b5a;
  display:flex;align-items:center;gap:7px;
}
.definitiva-badge{
  font-size:.82rem;font-weight:700;padding:5px 14px;
  border-radius:20px;border:2px solid;
}
.niv-badge{
  font-size:.68rem;font-weight:700;padding:3px 10px;
  border-radius:10px;background:#f3e8ff;color:#7c3aed;
  border:1px solid #d8b4fe;white-space:nowrap;
}
.table-wrap{overflow-x:auto;padding:0}
table{width:100%;border-collapse:collapse;min-width:400px;}
thead tr{background:#e8f5f1}
th{
  padding:10px 12px;
  font-size:.7rem;font-weight:700;letter-spacing:.3px;
  text-transform:uppercase;color:#1a6b5a;
  text-align:center;border-bottom:2px solid #b2dbd1;
  white-space:normal;line-height:1.3;vertical-align:bottom;
}
th.left{text-align:left}
td{padding:11px 12px;text-align:center;border-bottom:1px solid #edf2f7;font-size:.88rem;}
td.name-col{text-align:left;font-weight:500;font-size:.82rem;color:#4a5568;}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover td{background:#f7faf9}
.nota{
  display:inline-block;
  font-family:'Inter',monospace;font-weight:700;
  font-size:.88rem;padding:3px 10px;
  border-radius:20px;min-width:44px;
}
.nota.hi{background:#dcfce7;color:#15803d}
.nota.ok{background:#dbeafe;color:#1e40af}
.nota.md{background:#fef9c3;color:#92400e}
.nota.fl{background:#fee2e2;color:#b91c1c}
.nota.nd{background:#f1f5f9;color:#94a3b8;font-weight:400;font-style:italic;font-size:.76rem}
.escala-leyenda{
  display:flex;gap:8px;flex-wrap:wrap;
  padding:8px 16px 12px;font-size:.68rem;
}
.esc-item{
  display:flex;align-items:center;gap:4px;
  padding:2px 8px;border-radius:10px;font-weight:600;
}
.pct-warn{
  margin:0 16px 14px;padding:9px 12px;
  background:#fffbeb;border:1px solid #fbbf24;border-radius:8px;
  font-size:.76rem;color:#92400e;
  display:flex;align-items:center;gap:6px;
}
.niv-notice{
  margin:0 16px 14px;padding:9px 12px;
  background:#f3e8ff;border:1px solid #d8b4fe;border-radius:8px;
  font-size:.76rem;color:#6b21a8;
  display:flex;align-items:center;gap:6px;
}
.page-footer{
  text-align:center;margin-top:24px;
  font-size:.68rem;color:#9baab8;
}
@media(max-width:480px){
  .form-row{flex-direction:column;align-items:stretch}
  .form-label{min-width:auto}
  .hint-text{padding-left:0}
  .card{padding:18px}
  th,td{padding:8px 8px;font-size:.75rem}
  .nota{font-size:.8rem;padding:2px 7px;min-width:36px}
}
</style>
</head>
<body>

<div class="header">
  <div class="header-icon">⚗️</div>
  <div>
    <div class="header-title">Laboratorio de Harvey · Química</div>
    <div class="header-sub">Colegio Rufino José Cuervo IED · Bogotá</div>
  </div>
</div>

<div class="card">
  <div class="card-title">🎓 Consulta tus Notas</div>
  <div class="form-row">
    <label class="form-label" for="codeInput">Ingresa tu Código:</label>
    <input class="form-input" id="codeInput" type="text"
      placeholder="Ej: 100401"
      maxlength="12"
      onkeydown="if(event.key==='Enter')buscar()">
  </div>
  <div class="hint-text">Tu código = curso + número de lista · Ejemplo: 1004<strong>01</strong></div>
  <button class="search-btn" onclick="buscar()">Buscar Notas</button>
  <div class="err-msg" id="errMsg">
    <span>⚠</span><span id="errTxt"></span>
  </div>
</div>

<div id="resultado"></div>

<div class="page-footer">
  Aquí aparecerán tus notas agrupadas por periodo.
</div>

<script>
// ✅ CORREGIDO: null/undefined = sin nota (gris), 0 real = rojo
function nc(n){
  if(n===null||n===undefined) return 'nd';
  if(n>=4.6) return 'hi';
  if(n>=4.0) return 'ok';
  if(n>=3.0) return 'md';
  return 'fl';
}
function nf(n){
  return (n!==null && n!==undefined) ? n.toFixed(1) : '—';
}
function nivelLabel(n){
  if(n===null||n===undefined) return '';
  if(n>=4.6) return 'Superior';
  if(n>=4.0) return 'Alto';
  if(n>=3.0) return 'Básico';
  return 'Bajo';
}
function defColor(n){
  if(n===null) return '#6b7a8d';
  if(n>=4.6) return '#15803d';
  if(n>=4.0) return '#1e40af';
  if(n>=3.0) return '#92400e';
  return '#b91c1c';
}
function defBg(n){
  if(n===null) return '#f1f5f9';
  if(n>=4.6) return '#dcfce7';
  if(n>=4.0) return '#dbeafe';
  if(n>=3.0) return '#fef9c3';
  return '#fee2e2';
}

function buscar(){
  const code = document.getElementById('codeInput').value.trim();
  const errEl = document.getElementById('errMsg');
  const errTxt = document.getElementById('errTxt');
  errEl.classList.remove('show');

  if(!code){
    errTxt.textContent = 'Ingresa tu código';
    errEl.classList.add('show');
    return;
  }

  document.getElementById('resultado').innerHTML =
    '<div class="loading"><div class="spinner"></div>Buscando tus notas…</div>';

  google.script.run
    .withSuccessHandler(function(d){
      if(d.error){
        document.getElementById('resultado').innerHTML = '';
        const msgs = {
          notfound: 'Código no encontrado. Verifica el número o consulta con tu docente.',
          curso_invalido: 'El código no corresponde a un curso válido (1004–1006, 1101–1104).',
          empty: 'Ingresa tu código.',
          config: 'Error de configuración. Contacta al docente.'
        };
        errTxt.textContent = msgs[d.error] || 'Error inesperado. Intenta de nuevo.';
        errEl.classList.add('show');
      } else {
        render(d);
      }
    })
    .withFailureHandler(function(err){
      document.getElementById('resultado').innerHTML = '';
      errTxt.textContent = 'Error: ' + err.message;
      errEl.classList.add('show');
    })
    .consultarEstudiante(code);
}

function render(d){
  const nombre=d.nombre, activePeriod=d.activePeriod, periodos=d.periodos,
        code=d.code, definitivaAnual=d.definitivaAnual;
  const curso=d.course||d.curso;
  if(!periodos||!periodos.length){
    document.getElementById('resultado').innerHTML=
      '<div style="text-align:center;padding:20px;color:#6b7a8d">Sin notas registradas aún.</div>';
    return;
  }

  const tabsHtml='<div style="display:flex;gap:6px;margin-bottom:14px;flex-wrap:wrap">'
    +periodos.map(p=>
      '<button onclick="showPer('+p.period+')" id="ptab_'+p.period+'"'
      +' style="padding:5px 16px;border-radius:6px;border:1.5px solid #1a2e5a;'
      +'font-size:.76rem;font-weight:700;cursor:pointer;font-family:inherit;transition:all .15s;'
      +(p.isActive?'background:#1a2e5a;color:white':'background:white;color:#1a2e5a')
      +'">P'+p.period+(p.isActive?' ★':'')+'</button>'
    ).join('')
    +'<button onclick="showPer(0)" id="ptab_0"'
    +' style="padding:5px 16px;border-radius:6px;border:1.5px solid #c9952a;'
    +'font-size:.76rem;font-weight:700;cursor:pointer;font-family:inherit;transition:all .15s;'
    +'background:white;color:#c9952a">🏆 Definitiva Anual</button>'
    +'</div>';

  const periodBlocks=periodos.map(p=>buildBlock({...p,nombre})).join('')
    +buildAnoBlock(definitivaAnual, periodos, nombre);

  document.getElementById('resultado').innerHTML=
    '<div class="result-wrap">'
    +'<div class="result-header">'
    +'<div class="found-label">✅ Resultados encontrados</div>'
    +'<div class="student-name">'+nombre+'</div>'
    +'<div class="student-meta">Código: <strong>'+code+'</strong> · Curso: <strong>'+curso+'</strong></div>'
    +'</div>'
    +'<div class="course-block">'
    +tabsHtml
    +periodBlocks
    +'</div></div>';

  showPer(activePeriod);
}

function showPer(period){
  document.querySelectorAll('[id^="pblk_"]').forEach(el=>el.style.display='none');
  document.querySelectorAll('[id^="ptab_"]').forEach(el=>{
    el.style.background='white';
    el.style.color=el.id==='ptab_0'?'#c9952a':'#1a2e5a';
  });
  const blk=document.getElementById('pblk_'+period);
  const tab=document.getElementById('ptab_'+period);
  if(blk)blk.style.display='block';
  if(tab){
    tab.style.background=period===0?'#c9952a':'#1a2e5a';
    tab.style.color='white';
  }
}

function buildAnoBlock(definitivaAnual, periodos, nombre){
  const def = definitivaAnual;
  const nivel = nivelLabel(def);
  let rows = periodos.map(p=>
    '<tr>'
    +'<td class="name-col">Periodo '+p.period+(p.nivelado?' <span style="font-size:.65rem;color:#7c3aed;font-weight:700">(nivelado)</span>':'')+'</td>'
    +'<td><span class="nota '+nc(p.definitiva)+'">'+nf(p.definitiva)+'</span></td>'
    +'</tr>'
  ).join('');
  rows += '<tr style="background:#fef3d0;font-weight:700;border-top:2px solid #c9952a">'
    +'<td class="name-col" style="color:#92400e;font-weight:700">🏆 Definitiva Anual</td>'
    +'<td><span class="nota" style="font-size:1rem;font-weight:800;color:'+defColor(def)+';background:'+defBg(def)+'">'+(def!==null?def.toFixed(1):'—')+'</span></td>'
    +'</tr>';

  return '<div id="pblk_0" style="display:none">'
    +'<div class="course-head" style="border-bottom-color:#c9952a">'
    +'<div class="course-num" style="color:#c9952a">🏆 Definitiva Anual</div>'
    +(def!==null?'<div class="definitiva-badge" style="color:'+defColor(def)+';border-color:'+defColor(def)+';background:'+defBg(def)+'">'
      +def.toFixed(1)+' · '+nivel+'</div>':'')
    +'</div>'
    +'<div class="table-wrap"><table>'
    +'<thead><tr>'
    +'<th class="left">Periodo</th>'
    +'<th>Definitiva</th>'
    +'</tr></thead>'
    +'<tbody>'+rows+'</tbody>'
    +'</table></div>'
    +'<div style="padding:10px 16px 14px;font-size:.72rem;color:#6b7a8d">'
    +'Promedio de las definitivas de los 3 periodos académicos.'
    +'</div>'
    +'</div>';
}

function buildBlock(p){
  const{period,componentes,definitiva,totalPct,nivelado}=p;
  const def=definitiva;
  const nivel=nivelLabel(def);

  let html='<div id="pblk_'+period+'" style="display:none">'
    +'<div class="course-head">'
    +'<div class="course-num"><span>Periodo '+period+'</span>'
    +(nivel?'<span style="font-size:.68rem;padding:2px 8px;border-radius:10px;font-weight:600;background:'+defBg(def)+';color:'+defColor(def)+'">'+nivel+'</span>':'')
    +(nivelado?'<span class="niv-badge">📘 Periodo nivelado</span>':'')
    +'</div>'
    +'<div class="definitiva-badge" style="color:'+defColor(def)+';border-color:'+defColor(def)+';background:'+defBg(def)+'">'
    +(def!==null?'Definitiva: '+def.toFixed(1):'Sin definitiva aún')
    +'</div></div>';

  if(nivelado){
    html+='<div class="niv-notice">📘 Este periodo se niveló: la definitiva mostrada corresponde a la nota mínima de aprobación (3.0), independiente del detalle por componente de abajo.</div>';
  }

  const actComp=componentes.find(c=>c.key==='actividades');
  const otherComps=componentes.filter(c=>c.key!=='actividades');

  let thd='<th class="left">Estudiante</th>';
  (actComp?.actItems||[]).forEach((a)=>{
    thd+='<th title="'+a.name+'">'
      +'<div style="font-size:.65rem;max-width:60px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+a.name+'</div>'
      +'</th>';
  });
  otherComps.forEach(c=>{
    thd+='<th>'+c.label+'<br><span style="font-weight:400;font-size:.65rem">'+c.weight+'%</span></th>';
  });
  thd+='<th style="background:#1a3d6b;color:#f5d06a">DEFINITIVA</th>';

  let trow='<td class="name-col">'+p.nombre+'</td>';
  (actComp?.actItems||[]).forEach(a=>{
    trow+='<td><span class="nota '+nc(a.grade)+'">'+nf(a.grade)+'</span></td>';
  });
  otherComps.forEach(c=>{
    trow+='<td><span class="nota '+nc(c.grade)+'">'+nf(c.grade)+'</span></td>';
  });
  trow+='<td style="background:#fef3d0"><span class="nota" style="font-size:.95rem;font-weight:700;color:'+defColor(def)+'">'
    +(def!==null?def.toFixed(1):'—')+'</span></td>';

  html+='<div class="table-wrap"><table>'
    +'<thead><tr>'+thd+'</tr></thead>'
    +'<tbody><tr>'+trow+'</tr></tbody>'
    +'</table></div>';

  html+='<div class="escala-leyenda">'
    +'<span class="esc-item" style="background:#dcfce7;color:#15803d">Superior 4.6–5.0</span>'
    +'<span class="esc-item" style="background:#dbeafe;color:#1e40af">Alto 4.0–4.5</span>'
    +'<span class="esc-item" style="background:#fef9c3;color:#92400e">Básico 3.0–3.9</span>'
    +'<span class="esc-item" style="background:#fee2e2;color:#b91c1c">Bajo 0.0–2.9</span>'
    +'</div>';

  if(totalPct<100 && !nivelado){
    html+='<div class="pct-warn">⚠ Periodo en curso — la nota puede cambiar a medida que se registren más actividades.</div>';
  }

  html+='</div>';
  return html;
}
<\/script>
</body>
</html>`;
}
