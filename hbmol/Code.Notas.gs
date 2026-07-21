// ══════════════════════════════════════════════════════════════════
// HBMOL · Code.Notas.gs
// Sucesor de "NOTAS·QM". Google Apps Script standalone (Web App).
//
// DECISIÓN DE ARQUITECTURA (ver hbmol/ANALISIS_Y_DISEÑO.md §d):
//   Se consolida TODO en UN SOLO Google Sheet ("HBMOL"), compartido entre
//   este script y Code.Asistencia.gs. Ambos scripts son proyectos Apps
//   Script INDEPENDIENTES (cada uno su propio Web App / URL / password,
//   igual que hoy), pero los DOS abren el mismo SPREADSHEET_ID_HBMOL.
//   Ganancias de esto:
//     - La hoja "Config" (cursos, color, activo, periodo activo) es UNA
//       sola fuente de verdad: ya no hay que sincronizar dos catálogos
//       de cursos a mano.
//     - La hoja "Estudiantes" también es compartida: Asistencia deja de
//       mantener su propio roster por nombre y referencia el mismo ID
//       de estudiante que Notas. Esto elimina el desacople roster
//       notas/asistencia que tenía GestiónQM.
//     - La hoja "Log" (auditoría) es compartida: se puede ver en un solo
//       lugar todo lo que pasó, venga de Notas o de Asistencia.
//   Por eso este archivo es la fuente de verdad para el CRUD de
//   estudiantes y de cursos. Code.Asistencia.gs solo LEE esas hojas.
//
// Reemplaza a: gestionqm/Code.gs (backend de notas, NO TOCADO).
// ══════════════════════════════════════════════════════════════════

const PASSWORD = '951220'; // TODO: cámbiala antes de desplegar. Ver notas de
                            // seguridad en ANALISIS_Y_DISEÑO.md (sección de
                            // fragilidades) — sigue siendo un secreto único
                            // compartido, HBMOL no resuelve eso por sí solo,
                            // ver ahí las opciones (enlaces firmados, etc.)

// TODO: pega aquí el ID del Google Sheet "HBMOL" (el MISMO ID debe ir en
// Code.Asistencia.gs). Créalo una sola vez siguiendo hbmol/README.md.
const SPREADSHEET_ID_HBMOL = 'PEGA_AQUI_EL_ID_DEL_SHEET_HBMOL';
function getSS() { return SpreadsheetApp.openById(SPREADSHEET_ID_HBMOL); }

// Cursos por defecto usados SOLO la primera vez que se corre setup()
// sobre una hoja "Config" vacía (semilla inicial, heredada de GestiónQM).
// Después de eso, los cursos viven en la hoja Config y se administran
// con addCourse/removeCourse — este arreglo deja de leerse en runtime.
const DEFAULT_COURSES = [
  { course: '1004', color: '#1a6b5a' },
  { course: '1005', color: '#1a3d6b' },
  { course: '1006', color: '#c8960c' },
  { course: '1101', color: '#7c3aed' },
  { course: '1102', color: '#16a34a' },
  { course: '1103', color: '#c2410c' },
  { course: '1104', color: '#0369a1' },
];
const PERIODS = [1, 2, 3];

const SH_STUDENTS  = 'Estudiantes';  // COMPARTIDA con Code.Asistencia.gs
const SH_GRADES    = 'Notas';
const SH_SPECIALS  = 'Especiales';
const SH_CONFIG    = 'Config';       // COMPARTIDA con Code.Asistencia.gs
const SH_WEIGHTS   = 'Pesos';
const SH_LOG       = 'Log';          // COMPARTIDA con Code.Asistencia.gs

const BACKUP_FOLDER_NAME = 'HBMOL_Backups';
const BACKUP_KEEP_VERSIONS = 30; // conserva ~1 mes de backups fechados
const BASE_WEIGHTS = { actividades: 60, autoeval: 5, coeval: 5, final: 30 };

var _cache = {};

function _sheetData(name) {
  if (!_cache[name]) {
    const sh = getSS().getSheetByName(name);
    _cache[name] = { sh, rows: sh ? sh.getDataRange().getValues().slice(1) : [] };
  }
  return _cache[name];
}
function sheetRows(name) { return _sheetData(name); }
function _invalidate(name) { delete _cache[name]; }

// ──────────────────────────────────────────────────────────────────
// SETUP
// ──────────────────────────────────────────────────────────────────
function setup() {
  const ss = getSS();
  function ensure(name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a2e5a').setFontColor('white');
    }
    return sh;
  }
  ensure(SH_STUDENTS, ['Curso','ID','Nombre','Código','Activo']);
  ensure(SH_GRADES,   ['Curso','Periodo','Componente','ItemID','EstudianteID','Nota']);
  ensure(SH_SPECIALS, ['Curso','Periodo','ID','Nombre','Peso','PesoAct','PesoFinal']);
  ensure(SH_WEIGHTS,  ['Curso','Periodo','PesoAct','PesoAuto','PesoCoeval','PesoFinal']);
  ensure(SH_LOG,      ['Fecha','Módulo','Acción','Curso','Detalle']);

  // Config: Curso | PeriodoActivo | Color | Activo  (semilla solo si está vacía)
  const cfg = ensure(SH_CONFIG, ['Curso','PeriodoActivo','Color','Activo']);
  const existingRows = cfg.getDataRange().getValues().slice(1);
  if (existingRows.length === 0) {
    DEFAULT_COURSES.forEach(c => cfg.appendRow([c.course, 1, c.color, true]));
  }
  _invalidate(SH_CONFIG);

  // Hojas-vista legibles por curso/periodo (solo para cursos activos)
  getAllCourseCodes().forEach(course => {
    PERIODS.forEach(period => {
      const name = `${course}-P${period}`;
      if (!ss.getSheetByName(name)) {
        const sh = ss.insertSheet(name);
        sh.getRange(1,1).setValue(`Hoja de ${course} · Periodo ${period}`).setFontColor('#888888').setFontStyle('italic');
      }
    });
  });
  Logger.log('✅ Setup de HBMOL (Notas) completado.');
}

function json(d) { return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON); }
function uid() { return 'x' + Date.now().toString(36) + Math.random().toString(36).slice(2,6); }
function validGrade(v) {
  if (v === null || v === '') return null;
  const n = Math.round(parseFloat(v) * 10) / 10;
  if (isNaN(n) || n < 0 || n > 5) return false;
  return n;
}
function computeWeights(specials, baseWeights) {
  const totalSpecial = Math.round(specials.reduce((s,e) => s+e.weight, 0) * 10) / 10;
  if (baseWeights && baseWeights.actividades != null) {
    return { actividades: baseWeights.actividades, autoeval: baseWeights.autoeval, coeval: baseWeights.coeval, final: baseWeights.final, especiales: specials, base: baseWeights, totalSpecial };
  }
  let pesoAct = null, pesoFinal = null;
  for (let i = specials.length - 1; i >= 0; i--) {
    if (specials[i].pesoAct !== null && specials[i].pesoFinal !== null) { pesoAct = specials[i].pesoAct; pesoFinal = specials[i].pesoFinal; break; }
  }
  if (pesoAct === null || pesoFinal === null) {
    const fromAct = Math.min(totalSpecial, 60), fromFinal = Math.max(0, totalSpecial - 60);
    pesoAct = Math.round((60 - fromAct) * 10) / 10; pesoFinal = Math.round((30 - fromFinal) * 10) / 10;
  }
  return { actividades: pesoAct, autoeval: 5, coeval: 5, final: pesoFinal, especiales: specials, totalSpecial };
}

// ──────────────────────────────────────────────────────────────────
// AUDITORÍA — hoja "Log" compartida con Asistencia
// ──────────────────────────────────────────────────────────────────
function logAudit(action, course, detail) {
  try {
    const sh = getSS().getSheetByName(SH_LOG);
    if (!sh) return;
    sh.appendRow([new Date(), 'Notas', action, course || '', detail || '']);
  } catch (e) { /* la auditoría nunca debe tumbar la operación principal */ }
}

// ──────────────────────────────────────────────────────────────────
// ROUTER
// ──────────────────────────────────────────────────────────────────
function doGet(e) {
  const p = e.parameter;
  if (p.action === 'ping') return json({ ok:true });
  if (p.action === 'queryStudent') return json(queryStudent(p.code));
  if (p.pwd !== PASSWORD) return json({ error:'auth' });
  try {
    if (p.action === 'getStudents') return json(getStudents(p.course));
    if (p.action === 'getAll') return json(getAll(p.course, +p.period));
    if (p.action === 'getConfig') return json(getConfig());
    if (p.action === 'getCourses') return json(getCourses());
    if (p.action === 'downloadBackup') return json({ csv: generateFullCSV() });
    if (p.action === 'saveWeights') { const esp = p.especiales ? JSON.parse(p.especiales) : []; return json(saveWeightsFn(p.course, +p.period, +p.pesoAct, +p.pesoAuto, +p.pesoCoeval, +p.pesoFinal, esp)); }
    return json({ error:'unknown' });
  } catch(e) { return json({ error: e.message }); }
}
function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.pwd !== PASSWORD) return json({ error:'auth' });
  try {
    if (b.action === 'addStudent') return json(addStudent(b.course, b.name, b.code));
    if (b.action === 'addStudents') return json(addStudents(b.course, b.students));
    if (b.action === 'removeStudent') return json(removeStudent(b.course, b.studentId));
    if (b.action === 'updateStudent') return json(updateStudent(b.course, b.studentId, b.name, b.code));
    if (b.action === 'transferStudent') return json(transferStudent(b.course, b.studentId, b.destCourse));
    if (b.action === 'saveGrade') return json(saveGrade(b));
    if (b.action === 'addSpecial') return json(addSpecial(b.course, b.period, b.name, b.weight, b.pesoAct, b.pesoFinal));
    if (b.action === 'saveWeights') return json(saveWeightsFn(b.course, +b.period, +b.pesoAct, +b.pesoAuto, +b.pesoCoeval, +b.pesoFinal, b.especiales));
    if (b.action === 'removeSpecial') return json(removeSpecial(b.course, b.period, b.specialId));
    if (b.action === 'setActivePeriod') return json(setActivePeriod(b.course, b.period));
    if (b.action === 'addActivity') return json(addActivity(b.course, b.period, b.studentId, b.grade));
    if (b.action === 'removeActivity') return json(removeActivity(b.course, b.period, b.itemId));
    if (b.action === 'addCourse') return json(addCourse(b.course, b.color, b.copyPrevious));
    if (b.action === 'removeCourse') return json(removeCourse(b.course));
    return json({ error:'unknown' });
  } catch(e) { return json({ error: e.message }); }
}

// ──────────────────────────────────────────────────────────────────
// REBUILD DIFERIDO (hojas legibles por curso/periodo)
// ──────────────────────────────────────────────────────────────────
function scheduleRebuild(course, period) {
  const key = 'rebuild_' + course + '_' + period;
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(key)) return;
    props.setProperty(key, '1');
    ScriptApp.newTrigger('runPendingRebuilds').timeBased().after(60 * 1000).create();
  } catch(e) { rebuildCourseSheet(course, period); }
}
function runPendingRebuilds() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const rebuilt = [];
  Object.keys(all).forEach(key => {
    if (key.startsWith('rebuild_')) {
      const parts = key.replace('rebuild_', '').split('_');
      const course = parts[0], period = +parts[1];
      props.deleteProperty(key);
      try { rebuildCourseSheet(course, period); rebuilt.push(key); } catch(e) {}
    }
  });
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'runPendingRebuilds') ScriptApp.deleteTrigger(t); });
  Logger.log('Rebuilds completados: ' + rebuilt.join(', '));
}

// ──────────────────────────────────────────────────────────────────
// CURSOS (Config dinámica) — NUEVO en HBMOL
// ──────────────────────────────────────────────────────────────────
function _configRows() { return sheetRows(SH_CONFIG); }

/** Todos los cursos que alguna vez existieron (activos e inactivos). */
function getAllCourseCodes() { return _configRows().rows.map(r => String(r[0])); }

/** Solo cursos activos — es el equivalente al viejo COURSES hardcodeado. */
function getActiveCourseCodes() {
  return _configRows().rows.filter(r => r[3] == true).map(r => String(r[0]));
}

function getConfig() {
  const { rows } = sheetRows(SH_CONFIG);
  const r = {};
  rows.forEach(row => r[row[0]] = +row[1]);
  return r;
}
function getCourses() {
  const { rows } = sheetRows(SH_CONFIG);
  return rows.map(r => ({ course: String(r[0]), activePeriod: +r[1] || 1, color: r[2] || '#1a3d6b', active: r[3] == true }));
}
function setActivePeriod(course, period) {
  const { sh, rows } = sheetRows(SH_CONFIG);
  for (let i=0; i<rows.length; i++) { if (rows[i][0]==course) { sh.getRange(i+2,2).setValue(period); _invalidate(SH_CONFIG); return {ok:true}; } }
  sh.appendRow([course, period, '#1a3d6b', true]); _invalidate(SH_CONFIG);
  return { ok:true };
}
function addCourse(course, color, copyPrevious) {
  course = String(course || '').trim();
  if (!course) return { ok:false, error:'curso_vacio' };
  const { sh, rows } = sheetRows(SH_CONFIG);
  const existing = rows.find(r => String(r[0]) === course);
  if (existing) {
    // Si ya existía pero estaba inactivo, esto lo reactiva en vez de duplicar.
    if (existing[3] != true) return reactivateCourse(course);
    return { ok:false, error:'curso_existente' };
  }
  sh.appendRow([course, 1, color || '#1a3d6b', true]);
  _invalidate(SH_CONFIG);
  // Crea las hojas-vista de los 3 periodos para que el curso nuevo se vea
  // de inmediato, igual que los cursos originales.
  PERIODS.forEach(p => scheduleRebuild(course, p));
  logAudit('addCourse', course, 'Curso creado' + (color ? ' · color ' + color : ''));
  return { ok:true, course, color: color || '#1a3d6b' };
}
function reactivateCourse(course) {
  const { sh, rows } = sheetRows(SH_CONFIG);
  for (let i=0;i<rows.length;i++) { if (String(rows[i][0])===String(course)) { sh.getRange(i+2,4).setValue(true); _invalidate(SH_CONFIG); logAudit('reactivateCourse', course, 'Curso reactivado'); return { ok:true, course, reactivated:true }; } }
  return { ok:false, error:'no_encontrado' };
}
/**
 * "Elimina" un curso SIN BORRAR datos históricos: solo lo marca inactivo en
 * Config. Estudiantes, notas y asistencia del curso quedan intactos y
 * consultables (p.ej. desde el backup o reactivando el curso más tarde).
 */
function removeCourse(course) {
  course = String(course || '').trim();
  const { sh, rows } = sheetRows(SH_CONFIG);
  for (let i=0;i<rows.length;i++) {
    if (String(rows[i][0]) === course) {
      sh.getRange(i+2,4).setValue(false);
      _invalidate(SH_CONFIG);
      logAudit('removeCourse', course, 'Curso marcado inactivo (datos conservados)');
      return { ok:true, course };
    }
  }
  return { ok:false, error:'no_encontrado' };
}

// ──────────────────────────────────────────────────────────────────
// ESTUDIANTES (hoja compartida con Asistencia)
// ──────────────────────────────────────────────────────────────────
function getStudents(course) { const { rows } = sheetRows(SH_STUDENTS); return rows.filter(r=>r[0]==course&&r[4]==true).map(r=>({ id:r[1], name:r[2], code:r[3]||'' })); }
function addStudent(course, name, code) {
  const sh = getSS().getSheetByName(SH_STUDENTS);
  if (code) { const { rows } = sheetRows(SH_STUDENTS); const dup = rows.find(r => r[4]==true && r[3].toString().trim().toLowerCase()==code.toString().trim().toLowerCase()); if (dup) return { ok:false, error:'duplicate_code', existing: dup[2] }; }
  const id = uid(); sh.appendRow([course, id, name.trim(), (code||'').trim(), true]); _invalidate(SH_STUDENTS);
  logAudit('addStudent', course, name.trim() + (code ? ' (código ' + code + ')' : ''));
  return { ok:true, id, name:name.trim(), code:(code||'').trim() };
}
function addStudents(course, students) {
  const sh = getSS().getSheetByName(SH_STUDENTS);
  const { rows } = sheetRows(SH_STUDENTS);
  const existingCodes = new Set(rows.filter(r=>r[4]==true&&r[3]).map(r=>r[3].toString().trim().toLowerCase()));
  const added=[], skipped=[];
  students.forEach(s => {
    const name=(s.name||'').trim(), code=(s.code||'').trim();
    if (!name) return;
    if (code && existingCodes.has(code.toLowerCase())) { skipped.push(name+' (código duplicado)'); return; }
    const id=uid(); sh.appendRow([course, id, name, code, true]);
    if (code) existingCodes.add(code.toLowerCase());
    added.push({ id, name, code }); Utilities.sleep(5);
  });
  _invalidate(SH_STUDENTS);
  logAudit('addStudents', course, added.length + ' estudiante(s) agregado(s)' + (skipped.length ? ', ' + skipped.length + ' omitido(s)' : ''));
  return { ok:true, added, skipped };
}
function removeStudent(course, studentId) {
  const { sh, rows } = sheetRows(SH_STUDENTS);
  for (let i=0;i<rows.length;i++) { if (rows[i][0]==course&&rows[i][1]==studentId) { sh.getRange(i+2,5).setValue(false); _invalidate(SH_STUDENTS); logAudit('removeStudent', course, 'ID ' + studentId); return {ok:true}; } }
  return { ok:false };
}
function updateStudent(course, studentId, name, code) {
  if (code) { const { rows } = sheetRows(SH_STUDENTS); const dup = rows.find(r => r[4]==true && r[1]!=studentId && r[3].toString().trim().toLowerCase()==code.toString().trim().toLowerCase()); if (dup) return { ok:false, error:'duplicate_code', existing: dup[2] }; }
  const { sh, rows } = sheetRows(SH_STUDENTS);
  for (let i=0;i<rows.length;i++) { if (rows[i][0]==course&&rows[i][1]==studentId) { sh.getRange(i+2,3,1,2).setValues([[name.trim(),(code||'').trim()]]); _invalidate(SH_STUDENTS); return { ok:true }; } }
  return { ok:false };
}

function transferStudent(fromCourse, studentId, toCourse) {
  if (!fromCourse || !studentId || !toCourse) return { ok:false, error:'parametros_incompletos' };
  if (!getActiveCourseCodes().includes(String(toCourse))) return { ok:false, error:'curso_destino_invalido' };
  const { sh: stuSh, rows: stuRows } = sheetRows(SH_STUDENTS);
  let found = false;
  for (let i = 0; i < stuRows.length; i++) {
    if (String(stuRows[i][0]) === String(fromCourse) && String(stuRows[i][1]) === String(studentId) && stuRows[i][4] == true) { stuSh.getRange(i + 2, 1).setValue(toCourse); found = true; break; }
  }
  if (!found) return { ok:false, error:'estudiante_no_encontrado' };
  _invalidate(SH_STUDENTS);
  const { sh: grdSh, rows: grdRows } = sheetRows(SH_GRADES);
  let notasMigradas = 0;
  for (let i = 0; i < grdRows.length; i++) { if (String(grdRows[i][0]) === String(fromCourse) && String(grdRows[i][4]) === String(studentId)) { grdSh.getRange(i + 2, 1).setValue(toCourse); notasMigradas++; } }
  _invalidate(SH_GRADES);
  PERIODS.forEach(p => { scheduleRebuild(fromCourse, p); scheduleRebuild(toCourse, p); });
  logAudit('transferStudent', fromCourse + ' → ' + toCourse, 'ID ' + studentId + ' · ' + notasMigradas + ' nota(s) migradas. NOTA: la asistencia histórica de este estudiante queda registrada bajo el curso de origen (' + fromCourse + ') porque la hoja Asistencia no re-etiqueta filas pasadas; solo las notas futuras/actuales migran de curso.');
  return { ok:true, studentId, from:fromCourse, to:toCourse, notasMigradas };
}

// ──────────────────────────────────────────────────────────────────
// ESPECIALES / PESOS
// ──────────────────────────────────────────────────────────────────
function getSpecials(course, period) {
  const { rows } = sheetRows(SH_SPECIALS);
  return rows.filter(r=>r[0]==course&&r[1]==period).map(r=>({ id:r[2], name:r[3], weight:+r[4], pesoAct: r[5]!==''&&r[5]!==undefined ? +r[5] : null, pesoFinal: r[6]!==''&&r[6]!==undefined ? +r[6] : null }));
}
function getWeights(course, period) {
  const sh = getSS().getSheetByName(SH_WEIGHTS);
  if (!sh) return null;
  const rows = sh.getDataRange().getValues().slice(1);
  const row = rows.find(r => r[0]==course && r[1]==period);
  if (!row) return null;
  return { actividades: +row[2], autoeval: +row[3], coeval: +row[4], final: +row[5] };
}
function saveWeightsFn(course, period, pesoAct, pesoAuto, pesoCoeval, pesoFinal, especiales) {
  const sh = getSS().getSheetByName(SH_WEIGHTS);
  if (!sh) return { ok:false, error:'no_sheet' };
  const rows = sh.getDataRange().getValues();
  const applyEspeciales = () => {
    if (especiales && especiales.length) {
      const spSh = getSS().getSheetByName(SH_SPECIALS); const spRows = spSh.getDataRange().getValues();
      especiales.forEach(esp => { for (let j = 1; j < spRows.length; j++) { if (spRows[j][0]==course && spRows[j][1]==period && spRows[j][2]==esp.id) { spSh.getRange(j+1, 5, 1, 3).setValues([[esp.weight, pesoAct, pesoFinal]]); break; } } });
      _invalidate(SH_SPECIALS);
    }
  };
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]==course && rows[i][1]==period) {
      sh.getRange(i+1, 3, 1, 4).setValues([[pesoAct, pesoAuto, pesoCoeval, pesoFinal]]);
      applyEspeciales();
      logAudit('saveWeights', course, 'Periodo ' + period + ' · Act=' + pesoAct + '% Final=' + pesoFinal + '%');
      scheduleRebuild(course, +period);
      return { ok:true };
    }
  }
  sh.appendRow([course, period, pesoAct, pesoAuto, pesoCoeval, pesoFinal]);
  applyEspeciales();
  logAudit('saveWeights', course, 'Periodo ' + period + ' · Act=' + pesoAct + '% Final=' + pesoFinal + '% (nuevo)');
  scheduleRebuild(course, +period);
  return { ok:true };
}
function addSpecial(course, period, name, weight, pesoAct, pesoFinal) {
  weight = Math.round(+weight * 10) / 10;
  pesoAct = pesoAct !== undefined ? Math.round(+pesoAct * 10) / 10 : '';
  pesoFinal = pesoFinal !== undefined ? Math.round(+pesoFinal * 10) / 10 : '';
  if (!name || isNaN(weight) || weight <= 0 || weight > 90) return { ok:false, error:'invalid' };
  if (pesoAct !== '' && pesoFinal !== '') { const total = Math.round((pesoAct + 5 + 5 + pesoFinal + weight) * 10) / 10; if (Math.abs(total - 100) > 0.05) return { ok:false, error:'invalid_total', total }; }
  const specials = getSpecials(course, period);
  const totalSp = specials.reduce((s,e) => s+e.weight, 0);
  if (totalSp + weight > 90) return { ok:false, error:'overflow', available: Math.round((90-totalSp)*10)/10 };
  const sh = getSS().getSheetByName(SH_SPECIALS); const id = uid();
  sh.appendRow([course, period, id, name.trim(), weight, pesoAct, pesoFinal]); _invalidate(SH_SPECIALS); scheduleRebuild(course, +period);
  logAudit('addSpecial', course, 'Periodo ' + period + ' · "' + name.trim() + '" (' + weight + '%)');
  return { ok:true, id, name:name.trim(), weight, pesoAct, pesoFinal };
}
function removeSpecial(course, period, specialId) {
  const { sh, rows } = sheetRows(SH_SPECIALS);
  for (let i=rows.length-1;i>=0;i--) { if (rows[i][0]==course&&rows[i][1]==period&&rows[i][2]==specialId) sh.deleteRow(i+2); }
  _invalidate(SH_SPECIALS);
  const { sh:sg, rows:gr } = sheetRows(SH_GRADES);
  for (let i=gr.length-1;i>=0;i--) { if (gr[i][0]==course&&gr[i][1]==period&&gr[i][2]=='especial'&&gr[i][3]==specialId) sg.deleteRow(i+2); }
  _invalidate(SH_GRADES); scheduleRebuild(course, +period);
  logAudit('removeSpecial', course, 'Periodo ' + period + ' · ID ' + specialId);
  return { ok:true };
}

function addActivity(course, period, studentId, grade) { return { ok:true }; }
function removeActivity(course, period, itemId) {
  const { sh, rows } = sheetRows(SH_GRADES);
  for (let i=rows.length-1;i>=0;i--) { if (rows[i][0]==course&&rows[i][1]==period&&rows[i][2]=='actividad'&&rows[i][3]==itemId) sh.deleteRow(i+2); }
  _invalidate(SH_GRADES); scheduleRebuild(course, +period);
  logAudit('removeActivity', course, 'Periodo ' + period + ' · itemId ' + itemId);
  return { ok:true };
}

function saveGrade(b) {
  const g = validGrade(b.grade);
  if (g === false) return { ok:false, error:'invalid_grade' };
  const { sh, rows } = sheetRows(SH_GRADES);
  for (let i=0;i<rows.length;i++) {
    if (rows[i][0]==b.course && rows[i][1]==b.period && rows[i][2]==b.component && rows[i][3]==b.itemId && rows[i][4]==b.studentId) {
      if (g===null) sh.deleteRow(i+2); else sh.getRange(i+2,6).setValue(g);
      _invalidate(SH_GRADES); scheduleRebuild(b.course, +b.period);
      logAudit('saveGrade', b.course, 'P' + b.period + ' · ' + b.component + '/' + b.itemId + ' · estudiante ' + b.studentId + ' → ' + (g===null?'(borrada)':g));
      return { ok:true };
    }
  }
  if (g!==null) {
    sh.appendRow([b.course, b.period, b.component, b.itemId, b.studentId, g]);
    _invalidate(SH_GRADES); scheduleRebuild(b.course, +b.period);
    logAudit('saveGrade', b.course, 'P' + b.period + ' · ' + b.component + '/' + b.itemId + ' · estudiante ' + b.studentId + ' → ' + g + ' (nueva)');
  }
  else if (b.studentId === '__col__') { sh.appendRow([b.course, b.period, b.component, b.itemId, b.studentId, '']); _invalidate(SH_GRADES); }
  return { ok:true };
}
function getGrades(course, period) {
  const { rows } = sheetRows(SH_GRADES);
  const result = {};
  rows.filter(r=>r[0]==course&&r[1]==period).forEach(r=>{ const [,, comp, item, sid, nota] = r; if (sid === '__col__') return; if (!result[comp]) result[comp]={}; if (!result[comp][item]) result[comp][item]={}; result[comp][item][sid] = +nota; });
  return result;
}
function getActivityNames(course, period) {
  const { rows } = sheetRows(SH_GRADES);
  const seen = new Set(), names = [];
  rows.filter(r=>r[0]==course&&r[1]==period&&r[2]=='actividad').forEach(r=>{ if(!seen.has(r[3])){ seen.add(r[3]); names.push(r[3]); } });
  return names;
}
function getAll(course, period) {
  const specials = getSpecials(course, period);
  const baseW = getWeights(course, period);
  const weights = computeWeights(specials, baseW);
  const grades = getGrades(course, period);
  const actNames = getActivityNames(course, period);
  return { students: getStudents(course), weights, actNames, grades };
}

// ──────────────────────────────────────────────────────────────────
// BACKUP — ahora CON VERSIONES (uno de los huecos de GestiónQM: el
// backup anterior sobrescribía siempre el mismo archivo). Aquí cada
// corrida crea un archivo fechado y se podan los más antiguos.
// ──────────────────────────────────────────────────────────────────
function generateFullCSV() {
  const lines = [];
  const ts = new Date().toLocaleString('es-CO');
  lines.push(`"BACKUP COMPLETO · HBMOL · ${ts}"`); lines.push('');
  getAllCourseCodes().forEach(course => {
    PERIODS.forEach(period => {
      const students = getStudents(course);
      const specials = getSpecials(course, period);
      const weights = computeWeights(specials);
      const grades = getGrades(course, period);
      const actNames = getActivityNames(course, period);
      lines.push(`"=== CURSO ${course} · PERIODO ${period} ==="`);
      const headers = ['#', 'Código', 'Nombre'];
      actNames.forEach((n, i) => headers.push(`Act.${i+1}: ${n}`));
      headers.push(`Prom.Act (${weights.actividades}%)`, 'Autoeval (5%)', 'Coeval (5%)', `Final (${weights.final}%)`);
      specials.forEach(sp => headers.push(`${sp.name} (${sp.weight}%)`));
      headers.push('DEFINITIVA');
      lines.push(headers.map(h => `"${h}"`).join(','));
      if (students.length === 0) { lines.push('"(sin estudiantes)"'); }
      else {
        students.forEach((s, i) => {
          const row = [i+1, s.code||'', s.name];
          const actVals = actNames.map(n => { const v=grades['actividad']?.[n]?.[s.id]; return (v!==undefined&&v!==null)?v:''; });
          actVals.forEach(v => row.push(v));
          const filled = actVals.filter(v=>v!=='');
          const actAvg = filled.length ? Math.round(filled.reduce((a,b)=>a+b,0)/filled.length*10)/10 : '';
          row.push(actAvg);
          const ae=grades['autoeval']?.['_']?.[s.id]??'', ce=grades['coeval']?.['_']?.[s.id]??'', fi=grades['final']?.['_']?.[s.id]??'';
          row.push(ae,ce,fi);
          specials.forEach(sp => row.push(grades['especial']?.[sp.id]?.[s.id]??''));
          const totPct=weights.actividades+5+5+weights.final+specials.reduce((s,e)=>s+e.weight,0);
          let sumW2=0;
          sumW2+=(actAvg!==''?actAvg:0)*(weights.actividades/100);
          sumW2+=(ae!==''?ae:0)*(5/100); sumW2+=(ce!==''?ce:0)*(5/100); sumW2+=(fi!==''?fi:0)*(weights.final/100);
          specials.forEach((sp,idx)=>{ const v=row[3+actNames.length+4+idx]??''; sumW2+=(v!==''?v:0)*(sp.weight/100); });
          const def=totPct>0?Math.round(sumW2/(totPct/100)*100)/100:0;
          row.push(def);
          lines.push(row.map(v=>`"${v}"`).join(','));
        });
      }
      lines.push(''); _invalidate(SH_GRADES);
    });
  });
  return lines.join('\r\n');
}
function _backupFolder() {
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
}
function dailyBackup() {
  const csv = generateFullCSV();
  const stamp = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd_HHmm');
  const filename = `HBMOL_Backup_${stamp}.csv`;
  const folder = _backupFolder();
  folder.createFile(Utilities.newBlob('﻿' + csv, 'text/csv;charset=utf-8', filename));
  // Poda: conserva solo los BACKUP_KEEP_VERSIONS más recientes.
  const files = [];
  const it = folder.getFilesByType('text/csv');
  while (it.hasNext()) files.push(it.next());
  files.sort((a,b) => b.getDateCreated() - a.getDateCreated());
  files.slice(BACKUP_KEEP_VERSIONS).forEach(f => f.setTrashed(true));
  logAudit('dailyBackup', '', filename + ' (' + files.length + ' versiones conservadas)');
  Logger.log('Backup: ' + filename + ' · ' + new Date().toLocaleString('es-CO'));
}
function createBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyBackup').timeBased().atHour(2).everyDays(1).create();
  Logger.log('Trigger de backup creado.');
}

// ──────────────────────────────────────────────────────────────────
// HOJAS-VISTA LEGIBLES
// ──────────────────────────────────────────────────────────────────
function rebuildCourseSheet(course, period) {
  const ss = getSS();
  const shName = `${course}-P${period}`;
  let sh = ss.getSheetByName(shName);
  if (!sh) sh = ss.insertSheet(shName);
  sh.clearContents(); sh.clearFormats();
  const students = getStudents(course);
  const specials = getSpecials(course, period);
  const weights = computeWeights(specials);
  const grades = getGrades(course, period);
  const actNames = getActivityNames(course, period);
  const headers = ['#', 'Código', 'Nombre'];
  actNames.forEach((n,i) => headers.push(`Act.${i+1}: ${n}`));
  headers.push(`Prom.Act (${weights.actividades}%)`, `Autoeval (5%)`, `Coeval (5%)`, `Final (${weights.final}%)`);
  specials.forEach(sp => headers.push(`${sp.name} (${sp.weight}%)`));
  headers.push('DEFINITIVA');
  const dataRows = students.map((s,i) => {
    const row = [i+1, s.code||'', s.name];
    const actVals = actNames.map(n => { const v=grades['actividad']?.[n]?.[s.id]; return (v!==undefined&&v!==null)?v:''; });
    actVals.forEach(v => row.push(v));
    const filled = actVals.filter(v=>v!=='');
    const actAvg = filled.length ? Math.round(filled.reduce((a,b)=>a+b,0)/filled.length*10)/10 : '';
    row.push(actAvg);
    const ae=grades['autoeval']?.['_']?.[s.id]??'', ce=grades['coeval']?.['_']?.[s.id]??'', fi=grades['final']?.['_']?.[s.id]??'';
    row.push(ae,ce,fi);
    specials.forEach(sp => row.push(grades['especial']?.[sp.id]?.[s.id]??''));
    const totalPctR=weights.actividades+5+5+weights.final+specials.reduce((s,e)=>s+e.weight,0);
    let sumWR=0;
    sumWR+=(actAvg!==''?actAvg:0)*(weights.actividades/100);
    sumWR+=(ae!==''?ae:0)*(5/100); sumWR+=(ce!==''?ce:0)*(5/100); sumWR+=(fi!==''?fi:0)*(weights.final/100);
    specials.forEach((sp,idx)=>{ const v=row[3+actNames.length+4+idx]??''; sumWR+=(v!==''?v:0)*(sp.weight/100); });
    const definitiva=totalPctR>0?Math.round(sumWR/(totalPctR/100)*100)/100:0;
    row.push(definitiva);
    return row;
  });
  const allRows = [headers, ...dataRows];
  if (allRows.length > 0) sh.getRange(1,1,allRows.length,headers.length).setValues(allRows);
  const totalCols = headers.length, totalRows = allRows.length;
  sh.getRange(1,1,1,totalCols).setBackground('#1a2e5a').setFontColor('white').setFontWeight('bold').setWrap(true);
  sh.setFrozenRows(1); sh.setFrozenColumns(3);
  const defCol = totalCols;
  if (dataRows.length > 0) {
    sh.getRange(2,defCol,dataRows.length,1).setFontWeight('bold').setBackground('#e8f5e9');
    dataRows.forEach((row,i) => { const val = row[defCol-1]; if (val==='') return; let bg = val>=4.6?'#dcfce7':val>=4.0?'#dbeafe':val>=3.0?'#fef9c3':val>=2.0?'#fce5cd':'#fcd2d2'; sh.getRange(i+2,defCol).setBackground(bg); });
    const promActCol = 3+actNames.length+1;
    sh.getRange(2,promActCol,dataRows.length,1).setBackground('#fff8e1').setFontWeight('bold');
  }
  sh.setColumnWidth(1,35); sh.setColumnWidth(2,90); sh.setColumnWidth(3,200);
  for (let c=4;c<=totalCols;c++) sh.setColumnWidth(c,70);
  sh.setColumnWidth(defCol,90);
  if (totalRows>1) sh.getRange(1,1,totalRows,totalCols).setBorder(true,true,true,true,true,true,'#cccccc',SpreadsheetApp.BorderStyle.SOLID);
  for (let r=2;r<=totalRows;r++) { if(r%2===0) sh.getRange(r,1,1,totalCols-1).setBackground('#f8f9fa'); }
  if (totalCols>3) sh.getRange(1,4,totalRows,totalCols-3).setHorizontalAlignment('center');
  sh.getRange(totalRows+2,1).setValue(`Actualizado: ${new Date().toLocaleString('es-CO')}`).setFontColor('#888888').setFontSize(9).setFontStyle('italic');
  _invalidate(SH_GRADES);
}
function rebuildAllSheets() {
  getAllCourseCodes().forEach(course => { PERIODS.forEach(period => { rebuildCourseSheet(course, period); }); });
  Logger.log('✅ Hojas de todos los cursos regeneradas.');
}

// ──────────────────────────────────────────────────────────────────
// PORTAL PÚBLICO DEL ESTUDIANTE (sin contraseña, por diseño)
// ──────────────────────────────────────────────────────────────────
function queryStudent(code) {
  if (!code||code.toString().trim()==='') return { error:'empty' };
  code = code.toString().trim();
  const { rows } = sheetRows(SH_STUDENTS);
  const match = rows.find(r=>r[4]==true&&r[3].toString().trim().toLowerCase()===code.toLowerCase());
  if (!match) return { error:'notfound' };
  const course=match[0], studentId=match[1], name=match[2];
  const config=getConfig(), period=config[course]||1;
  const specials=getSpecials(course,period), weights=computeWeights(specials);
  const grades=getGrades(course,period), actNames=getActivityNames(course,period);
  const actGrades=actNames.map(n=>(grades['actividad']?.[n]?.[studentId]??null));
  const validActs=actGrades.filter(g=>g!==null);
  const actAvg=validActs.length>0?Math.round((validActs.reduce((a,b)=>a+b,0)/validActs.length)*10)/10:null;
  const components=[
    { key:'actividades', label:'Actividades de clase', weight:weights.actividades, detail: actNames.length>0?`${validActs.length} de ${actNames.length} notas · promedio ${actAvg!==null?actAvg.toFixed(1):'—'}`:'Sin actividades aún', grade: actAvg },
    { key:'autoeval', label:'Autoevaluación', weight:weights.autoeval, grade: grades['autoeval']?.['_']?.[studentId]??null },
    { key:'coeval', label:'Coevaluación', weight:weights.coeval, grade: grades['coeval']?.['_']?.[studentId]??null },
    { key:'final', label:'Evaluación final', weight:weights.final, grade: grades['final']?.['_']?.[studentId]??null },
  ];
  weights.especiales.forEach(sp=>{ components.push({ key:'especial_'+sp.id, label:sp.name, weight:sp.weight, grade: grades['especial']?.[sp.id]?.[studentId]??null }); });
  const totalPct=components.reduce((s,c)=>s+c.weight,0);
  let sumW=0; components.forEach(c=>{ sumW+=(c.grade!==null?c.grade:0)*(c.weight/100); });
  const avg=totalPct>0?Math.round(sumW/(totalPct/100)*100)/100:null;
  return { ok:true, name, course, period, components, avg, totalPct, actCount:actNames.length };
}

// ──────────────────────────────────────────────────────────────────
// ALERTAS AUTOMÁTICAS A ACUDIENTES (NUEVO en HBMOL, opcional)
// Requiere que el estudiante tenga un correo de acudiente en una hoja
// "Acudientes" (Curso|EstudianteID|EmailAcudiente) que debes crear tú
// mismo si quieres usar esta función — no se crea en setup() porque
// implica manejar datos de contacto de terceros (ver ANALISIS_Y_DISEÑO.md).
// Se deja aquí como plantilla lista para activar, no como acción del router.
// ──────────────────────────────────────────────────────────────────
function checkLowGradeAlerts(course, period, threshold) {
  threshold = threshold || 3.0;
  const sh = getSS().getSheetByName('Acudientes');
  if (!sh) { Logger.log('No existe la hoja "Acudientes" — omitiendo alertas.'); return; }
  const contactos = sh.getDataRange().getValues().slice(1); // Curso|EstudianteID|Email
  const data = getAll(course, period);
  data.students.forEach(s => {
    const contacto = contactos.find(c => String(c[0])===String(course) && String(c[1])===String(s.id));
    if (!contacto || !contacto[2]) return;
    // (cálculo de definitiva simplificado; en producción reutiliza queryStudent)
    const acts = data.actNames.map(n => data.grades['actividad']?.[n]?.[s.id]).filter(v=>v!=null && v!==undefined);
    if (!acts.length) return;
    const prom = acts.reduce((a,b)=>a+b,0)/acts.length;
    if (prom < threshold) {
      MailApp.sendEmail({
        to: contacto[2],
        subject: `HBMOL · Alerta académica de ${s.name} — Química`,
        htmlBody: `<p>Estimado acudiente,</p><p>El promedio actual de actividades de <b>${s.name}</b> (curso ${course}, periodo ${period}) es <b>${prom.toFixed(1)}</b>, por debajo del umbral de seguimiento (${threshold.toFixed(1)}).</p><p>Se sugiere agendar un espacio de acompañamiento en casa. Cualquier duda, contactar al docente.</p><p>— Harvey Barcenas Morales, Ciencias Naturales · Química</p>`
      });
    }
  });
}
