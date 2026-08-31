// ══════════════════════════════════════════════════════════════════
// NOTAS·QM — Code.gs
// Google Sheet: https://docs.google.com/spreadsheets/d/1FzUF-Rnu_6RB3exJXbtTS3QHyMGz-59vhomP3EXSyGE
//
// Este archivo es la copia de referencia versionada en git del Code.gs real,
// que vive en el proyecto de Apps Script vinculado al Sheet de arriba.
// Después de editar aquí, hay que pegar el contenido completo en el editor
// de Apps Script y desplegar una NUEVA VERSIÓN (Implementar → Administrar
// implementaciones → Editar → Nueva versión) — guardar el script NO
// actualiza la URL en producción por sí solo.
//
// Cambios respecto a la versión anterior:
//   - Cursos dinámicos: la hoja Config gana columnas Color y Activo, y el
//     router gana las acciones getCourses / addCourse / removeCourse.
//   - Log de auditoría simple (hoja Log) escrito en cada operación de
//     escritura relevante.
//   - Backup diario con versiones fechadas + poda automática, en vez de
//     sobrescribir siempre el mismo archivo.
// ══════════════════════════════════════════════════════════════════

const PASSWORD = '951220';

const SPREADSHEET_ID_NOTAS = '1FzUF-Rnu_6RB3exJXbtTS3QHyMGz-59vhomP3EXSyGE';
function getSS() { return SpreadsheetApp.openById(SPREADSHEET_ID_NOTAS); }

// Semilla usada solo por setup() la primera vez que se crea la hoja Config.
// Después de la primera ejecución, la lista real de cursos vive en la hoja
// Config (columna Activo) y se administra con getCourses/addCourse/removeCourse.
const COURSES_SEED = ['1004','1005','1006','1101','1102','1103','1104'];
const COLORS_SEED = {
  '1004':'#1a6b5a','1005':'#1a3d6b','1006':'#c8960c',
  '1101':'#7c3aed','1102':'#16a34a','1103':'#c2410c','1104':'#0369a1'
};
const PERIODS = [1, 2, 3];

const SH_STUDENTS  = 'Estudiantes';
const SH_GRADES    = 'Notas';
const SH_SPECIALS  = 'Especiales';
const SH_CONFIG    = 'Config';
const SH_WEIGHTS   = 'Pesos';
const SH_LOG       = 'Log';

const BACKUP_FOLDER_NAME    = 'NotasQM_Backups';
const BACKUP_KEEP_VERSIONS  = 30;
const BASE_WEIGHTS = { actividades: 60, autoeval: 5, coeval: 5, final: 30 };

// ════════════════════════════════════════════════════════════════
// CACHÉ DE LECTURA — evita leer el Sheet múltiples veces por request
// ════════════════════════════════════════════════════════════════
var _cache = {};

function _sheetData(name) {
  if (!_cache[name]) {
    const sh = getSS().getSheetByName(name);
    _cache[name] = { sh, rows: sh ? sh.getDataRange().getValues().slice(1) : [] };
  }
  return _cache[name];
}

function sheetRows(name) {
  return _sheetData(name);
}

function _invalidate(name) {
  delete _cache[name];
}

// ════════════════════════════════════════════════════════════════
// SETUP
// ════════════════════════════════════════════════════════════════
function setup() {
  const ss = getSS();
  function ensure(name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,headers.length)
        .setFontWeight('bold').setBackground('#0f2027').setFontColor('white');
    }
    return sh;
  }
  ensure(SH_STUDENTS, ['Curso','ID','Nombre','Código','Activo']);
  ensure(SH_GRADES,   ['Curso','Periodo','Componente','ItemID','EstudianteID','Nota']);
  ensure(SH_SPECIALS, ['Curso','Periodo','ID','Nombre','Peso','PesoAct','PesoFinal']);
  const cfg = ensure(SH_CONFIG, ['Curso','PeriodoActivo','Color','Activo']);
  ensure(SH_WEIGHTS, ['Curso','Periodo','PesoAct','PesoAuto','PesoCoeval','PesoFinal']);
  ensure(SH_LOG, ['Fecha','Accion','Curso','Detalle']);
  ensureConfigColumns();

  const existing = cfg.getDataRange().getValues().slice(1).map(r=>String(r[0]));
  COURSES_SEED.forEach(c => {
    if(!existing.includes(c)) cfg.appendRow([c, 1, COLORS_SEED[c] || '#1a3d6b', true]);
  });

  getActiveCourseCodes().forEach(course => {
    PERIODS.forEach(period => {
      const name = `${course}-P${period}`;
      if (!ss.getSheetByName(name)) {
        const sh = ss.insertSheet(name);
        sh.getRange(1,1).setValue(`Hoja de ${course} · Periodo ${period}`)
          .setFontColor('#888888').setFontStyle('italic');
      }
    });
  });

  Logger.log('✅ Setup completado.');
}

// Migra la hoja Config si venía de una versión anterior sin Color/Activo.
function ensureConfigColumns() {
  const sh = getSS().getSheetByName(SH_CONFIG);
  if (!sh) return;
  const lastCol = Math.max(sh.getLastColumn(), 2);
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  if (headers[2] !== 'Color' || headers[3] !== 'Activo') {
    sh.getRange(1,3,1,2).setValues([['Color','Activo']]);
    // Backfill: filas existentes sin Activo se consideran activas por defecto.
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][3] === '' || rows[i][3] === undefined) {
        sh.getRange(i+1, 4).setValue(true);
      }
      if (rows[i][2] === '' || rows[i][2] === undefined) {
        sh.getRange(i+1, 3).setValue(COLORS_SEED[String(rows[i][0])] || '#1a3d6b');
      }
    }
  }
  _invalidate(SH_CONFIG);
}

// ════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════
function json(d) {
  return ContentService.createTextOutput(JSON.stringify(d))
    .setMimeType(ContentService.MimeType.JSON);
}

function uid() {
  return 'x' + Date.now().toString(36) + Math.random().toString(36).slice(2,6);
}

function validGrade(v) {
  if (v === null || v === '') return null;
  const n = Math.round(parseFloat(v) * 10) / 10;
  if (isNaN(n) || n < 0 || n > 5) return false;
  return n;
}

function computeWeights(specials, baseWeights) {
  const totalSpecial = Math.round(specials.reduce((s,e) => s+e.weight, 0) * 10) / 10;
  if (baseWeights && baseWeights.actividades != null) {
    return {
      actividades: baseWeights.actividades,
      autoeval:    baseWeights.autoeval,
      coeval:      baseWeights.coeval,
      final:       baseWeights.final,
      especiales:  specials,
      base:        baseWeights,
      totalSpecial
    };
  }
  let pesoAct = null, pesoFinal = null;
  for (let i = specials.length - 1; i >= 0; i--) {
    if (specials[i].pesoAct !== null && specials[i].pesoFinal !== null) {
      pesoAct   = specials[i].pesoAct;
      pesoFinal = specials[i].pesoFinal;
      break;
    }
  }
  if (pesoAct === null || pesoFinal === null) {
    const fromAct   = Math.min(totalSpecial, 60);
    const fromFinal = Math.max(0, totalSpecial - 60);
    pesoAct   = Math.round((60 - fromAct)   * 10) / 10;
    pesoFinal = Math.round((30 - fromFinal) * 10) / 10;
  }
  return { actividades: pesoAct, autoeval: 5, coeval: 5, final: pesoFinal, especiales: specials, totalSpecial };
}

// ── Log de auditoría ────────────────────────────────────────────
// Deliberadamente simple: no guarda el valor anterior de cada celda, solo
// qué acción ocurrió, en qué curso y cuándo — suficiente para responder
// "¿qué pasó el martes a las 8pm en el curso 1101?" sin bloquear la
// operación principal si el log falla por cualquier razón.
function writeLog(action, course, detail) {
  try {
    const sh = getSS().getSheetByName(SH_LOG);
    if (!sh) return;
    sh.appendRow([new Date(), action, course || '', detail || '']);
  } catch(e) { /* no interrumpir la operación principal */ }
}

// ── Cursos dinámicos ─────────────────────────────────────────────
function getCourses() {
  ensureConfigColumns();
  const { rows } = sheetRows(SH_CONFIG);
  // Deduplica por curso (p.ej. si la hoja Config quedó con filas repetidas
  // por una edición manual): se conserva el orden de primera aparición,
  // pero con los datos de la ÚLTIMA fila de cada curso, que es la más
  // probable de reflejar el estado actual.
  const order = [];
  const byCourse = {};
  rows.forEach(r => {
    const course = String(r[0]).trim();
    if (!course) return;
    if (!byCourse[course]) order.push(course);
    byCourse[course] = {
      course: course,
      periodoActivo: +r[1] || 1,
      color: r[2] || '#1a3d6b',
      active: r[3] !== false
    };
  });
  return order.map(c => byCourse[c]);
}
function getActiveCourseCodes() {
  return getCourses().filter(c => c.active).map(c => c.course);
}

function addCourse(course, color) {
  course = String(course || '').trim();
  if (!course) return { ok:false, error:'curso_vacio' };
  ensureConfigColumns();
  const { sh, rows } = sheetRows(SH_CONFIG);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === course) {
      // Ya existe (posiblemente inactivo) — reactivar en vez de duplicar.
      sh.getRange(i+2, 4).setValue(true);
      if (color) sh.getRange(i+2, 3).setValue(color);
      _invalidate(SH_CONFIG);
      writeLog('addCourse', course, 'Curso reactivado');
      ensureCoursePeriodSheets(course);
      return { ok:true, reactivated:true };
    }
  }
  sh.appendRow([course, 1, color || '#1a3d6b', true]);
  _invalidate(SH_CONFIG);
  ensureCoursePeriodSheets(course);
  writeLog('addCourse', course, 'Curso creado · color ' + (color || '#1a3d6b'));
  return { ok:true };
}

function removeCourse(course) {
  course = String(course || '').trim();
  ensureConfigColumns();
  const { sh, rows } = sheetRows(SH_CONFIG);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === course) {
      sh.getRange(i+2, 4).setValue(false);
      _invalidate(SH_CONFIG);
      writeLog('removeCourse', course, 'Curso desactivado (estudiantes y notas se conservan)');
      return { ok:true };
    }
  }
  return { ok:false, error:'no_encontrado' };
}

function ensureCoursePeriodSheets(course) {
  const ss = getSS();
  PERIODS.forEach(period => {
    const name = `${course}-P${period}`;
    if (!ss.getSheetByName(name)) {
      const sh = ss.insertSheet(name);
      sh.getRange(1,1).setValue(`Hoja de ${course} · Periodo ${period}`)
        .setFontColor('#888888').setFontStyle('italic');
    }
  });
}

// ════════════════════════════════════════════════════════════════
// ROUTER
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  const p = e.parameter;
  if (p.action === 'ping')         return json({ ok:true });
  if (p.action === 'queryStudent') return json(queryStudent(p.code));
  if (p.pwd !== PASSWORD)          return json({ error:'auth' });
  try {
    if (p.action === 'getStudents') return json(getStudents(p.course));
    if (p.action === 'getAll')      return json(getAll(p.course, +p.period));
    if (p.action === 'getConfig')   return json(getConfig());
    if (p.action === 'getCourses')  return json(getCourses());
    if (p.action === 'downloadBackup') return json({ csv: generateFullCSV() });
    if (p.action === 'saveWeights') {
      const esp = p.especiales ? JSON.parse(p.especiales) : [];
      return json(saveWeightsFn(p.course, +p.period, +p.pesoAct, +p.pesoAuto, +p.pesoCoeval, +p.pesoFinal, esp));
    }
    return json({ error:'unknown' });
  } catch(e) { return json({ error: e.message }); }
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.pwd !== PASSWORD) return json({ error:'auth' });
  try {
    if (b.action === 'addStudent')        return json(addStudent(b.course, b.name, b.code));
    if (b.action === 'addStudents')       return json(addStudents(b.course, b.students));
    if (b.action === 'removeStudent')     return json(removeStudent(b.course, b.studentId));
    if (b.action === 'updateStudent')     return json(updateStudent(b.course, b.studentId, b.name, b.code));
    if (b.action === 'transferStudent')   return json(transferStudent(b.course, b.studentId, b.destCourse));
    if (b.action === 'saveGrade')         return json(saveGrade(b));
    if (b.action === 'addSpecial')        return json(addSpecial(b.course, b.period, b.name, b.weight, b.pesoAct, b.pesoFinal));
    if (b.action === 'saveWeights')       return json(saveWeightsFn(b.course, +b.period, +b.pesoAct, +b.pesoAuto, +b.pesoCoeval, +b.pesoFinal, b.especiales));
    if (b.action === 'removeSpecial')     return json(removeSpecial(b.course, b.period, b.specialId));
    if (b.action === 'setActivePeriod')   return json(setActivePeriod(b.course, b.period));
    if (b.action === 'removeActivity')    return json(removeActivity(b.course, b.period, b.itemId));
    if (b.action === 'addCourse')         return json(addCourse(b.course, b.color));
    if (b.action === 'removeCourse')      return json(removeCourse(b.course));
    return json({ error:'unknown' });
  } catch(e) { return json({ error: e.message }); }
}

// ════════════════════════════════════════════════════════════════
// REBUILD DIFERIDO
// ════════════════════════════════════════════════════════════════
function scheduleRebuild(course, period) {
  const key = 'rebuild_' + course + '_' + period;
  const props = PropertiesService.getScriptProperties();
  try {
    const existing = props.getProperty(key);
    if (existing) return;
    props.setProperty(key, '1');
    ScriptApp.newTrigger('runPendingRebuilds')
      .timeBased().after(60 * 1000).create();
  } catch(e) {
    // Si falla la creación del trigger (p.ej. cuota de triggers agotada), la
    // bandera ya quedó puesta arriba pero nunca se va a limpiar sola —
    // ningún trigger va a correr runPendingRebuilds() para borrarla. Sin
    // este catch, cada guardado de nota siguiente para este curso/periodo
    // vería la bandera y no reintentaría nada, dejando la hoja bonita
    // congelada hasta borrar la propiedad a mano.
    try { props.deleteProperty(key); } catch(e2) {}
    rebuildCourseSheet(course, period);
  }
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
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runPendingRebuilds') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Rebuilds completados: ' + rebuilt.join(', '));
}

// ════════════════════════════════════════════════════════════════
// CONFIG
// ════════════════════════════════════════════════════════════════
function getConfig() {
  const { rows } = sheetRows(SH_CONFIG);
  const r = {};
  rows.forEach(row => r[row[0]] = +row[1]);
  return r;
}

function setActivePeriod(course, period) {
  const { sh, rows } = sheetRows(SH_CONFIG);
  for (let i=0; i<rows.length; i++) {
    if (rows[i][0]==course) {
      sh.getRange(i+2,2).setValue(period); _invalidate(SH_CONFIG);
      writeLog('setActivePeriod', course, 'Periodo activo → P' + period);
      return {ok:true};
    }
  }
  sh.appendRow([course, period, COLORS_SEED[course] || '#1a3d6b', true]); _invalidate(SH_CONFIG);
  writeLog('setActivePeriod', course, 'Periodo activo → P' + period);
  return { ok:true };
}

// ════════════════════════════════════════════════════════════════
// STUDENTS
// ════════════════════════════════════════════════════════════════

// Trim que también quita espacios "invisibles" (NBSP y afines) que
// String.trim() no elimina — una línea pegada desde Excel/Word que se ve
// vacía pero trae uno de estos caracteres pasaba el chequeo !name y
// quedaba como estudiante "activo" sin nombre visible en la tabla de notas.
function cleanName(s) {
  return String(s == null ? '' : s).replace(/^[\s\u00A0\u200B\uFEFF]+|[\s\u00A0\u200B\uFEFF]+$/g, '');
}

function getStudents(course) {
  const { sh, rows } = sheetRows(SH_STUDENTS);
  const result = [];
  rows.forEach((r, i) => {
    if (r[0] != course || r[4] != true) return;
    const name = cleanName(r[2]);
    if (!name) {
      // Fila fantasma (activa pero sin nombre real): se autodesactiva al
      // detectarla para que deje de "existir" en vez de seguir apareciendo
      // en cada lectura futura.
      try { sh.getRange(i + 2, 5).setValue(false); _invalidate(SH_STUDENTS); } catch (e) {}
      return;
    }
    result.push({ id: r[1], name, code: r[3] || '' });
  });
  return result;
}

// LockService: addStudent/addStudents/updateStudent primero LEEN todas las
// filas para chequear código duplicado y luego ESCRIBEN — sin un lock, dos
// requests casi simultáneos (dos pestañas, o una alta manual mientras corre
// una importación masiva) pueden leer el mismo estado "sin duplicado" antes
// de que el otro escriba, y ambos terminan creando el mismo código dos veces.
function addStudent(course, name, code) {
  name = cleanName(name);
  if (!name) return { ok:false, error:'invalid_name' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSS().getSheetByName(SH_STUDENTS);
    if (code) {
      const { rows } = sheetRows(SH_STUDENTS);
      const dup = rows.find(r => r[4]==true && r[3].toString().trim().toLowerCase()==code.toString().trim().toLowerCase());
      if (dup) return { ok:false, error:'duplicate_code', existing: dup[2] };
    }
    const id = uid();
    sh.appendRow([course, id, name, (code||'').trim(), true]);
    _invalidate(SH_STUDENTS);
    writeLog('addStudent', course, name + (code?' · código ' + code:''));
    return { ok:true, id, name, code:(code||'').trim() };
  } finally {
    lock.releaseLock();
  }
}

function addStudents(course, students) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSS().getSheetByName(SH_STUDENTS);
    const { rows } = sheetRows(SH_STUDENTS);
    const existingCodes = new Set(rows.filter(r=>r[4]==true&&r[3]).map(r=>r[3].toString().trim().toLowerCase()));
    const added=[], skipped=[];
    students.forEach(s => {
      const name=cleanName(s.name), code=(s.code||'').trim();
      if (!name) return;
      if (code && existingCodes.has(code.toLowerCase())) { skipped.push(name+' (código duplicado)'); return; }
      const id=uid();
      sh.appendRow([course, id, name, code, true]);
      if (code) existingCodes.add(code.toLowerCase());
      added.push({ id, name, code });
      Utilities.sleep(5);
    });
    _invalidate(SH_STUDENTS);
    writeLog('addStudents', course, added.length + ' estudiante(s) importado(s)');
    return { ok:true, added, skipped };
  } finally {
    lock.releaseLock();
  }
}

function removeStudent(course, studentId) {
  const { sh, rows } = sheetRows(SH_STUDENTS);
  for (let i=0;i<rows.length;i++) {
    if (rows[i][0]==course&&rows[i][1]==studentId) {
      sh.getRange(i+2,5).setValue(false); _invalidate(SH_STUDENTS);
      writeLog('removeStudent', course, 'Estudiante ' + rows[i][2] + ' retirado');
      return {ok:true};
    }
  }
  return { ok:false };
}

function updateStudent(course, studentId, name, code) {
  name = cleanName(name);
  if (!name) return { ok:false, error:'invalid_name' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (code) {
      const { rows } = sheetRows(SH_STUDENTS);
      const dup = rows.find(r => r[4]==true && r[1]!=studentId && r[3].toString().trim().toLowerCase()==code.toString().trim().toLowerCase());
      if (dup) return { ok:false, error:'duplicate_code', existing: dup[2] };
    }
    const { sh, rows } = sheetRows(SH_STUDENTS);
    for (let i=0;i<rows.length;i++) {
      if (rows[i][0]==course&&rows[i][1]==studentId) {
        sh.getRange(i+2,3,1,2).setValues([[name,(code||'').trim()]]);
        _invalidate(SH_STUDENTS);
        writeLog('updateStudent', course, 'Estudiante actualizado: ' + name);
        return { ok:true };
      }
    }
    return { ok:false };
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════
// TRASLADO DE ESTUDIANTE ENTRE CURSOS
// ════════════════════════════════════════════════════════════════
function transferStudent(fromCourse, studentId, toCourse) {
  if (!fromCourse || !studentId || !toCourse) return { ok:false, error:'parametros_incompletos' };
  if (!getActiveCourseCodes().includes(String(toCourse))) return { ok:false, error:'curso_destino_invalido' };

  const { sh: stuSh, rows: stuRows } = sheetRows(SH_STUDENTS);
  let found = false;
  for (let i = 0; i < stuRows.length; i++) {
    if (String(stuRows[i][0]) === String(fromCourse) &&
        String(stuRows[i][1]) === String(studentId)  &&
        stuRows[i][4] == true) {
      stuSh.getRange(i + 2, 1).setValue(toCourse);
      found = true;
      break;
    }
  }
  if (!found) return { ok:false, error:'estudiante_no_encontrado' };
  _invalidate(SH_STUDENTS);

  const { sh: grdSh, rows: grdRows } = sheetRows(SH_GRADES);
  let notasMigradas = 0;
  for (let i = 0; i < grdRows.length; i++) {
    if (String(grdRows[i][0]) === String(fromCourse) &&
        String(grdRows[i][4]) === String(studentId)) {
      grdSh.getRange(i + 2, 1).setValue(toCourse);
      notasMigradas++;
    }
  }
  _invalidate(SH_GRADES);

  PERIODS.forEach(p => {
    scheduleRebuild(fromCourse, p);
    scheduleRebuild(toCourse,   p);
  });

  writeLog('transferStudent', fromCourse, 'Estudiante ' + studentId + ' → ' + toCourse + ' (' + notasMigradas + ' notas migradas)');
  return { ok:true, studentId, from:fromCourse, to:toCourse, notasMigradas };
}

// ════════════════════════════════════════════════════════════════
// ESPECIALES
// ════════════════════════════════════════════════════════════════
function getSpecials(course, period) {
  const { rows } = sheetRows(SH_SPECIALS);
  return rows.filter(r=>r[0]==course&&r[1]==period)
             .map(r=>({
               id:r[2], name:r[3], weight:+r[4],
               pesoAct:   r[5]!==''&&r[5]!==undefined ? +r[5] : null,
               pesoFinal: r[6]!==''&&r[6]!==undefined ? +r[6] : null
             }));
}

function getWeights(course, period) {
  const sh = getSS().getSheetByName(SH_WEIGHTS);
  if (!sh) return null;
  const rows = sh.getDataRange().getValues().slice(1);
  const row  = rows.find(r => r[0]==course && r[1]==period);
  if (!row) return null;
  return { actividades: +row[2], autoeval: +row[3], coeval: +row[4], final: +row[5] };
}

function saveWeightsFn(course, period, pesoAct, pesoAuto, pesoCoeval, pesoFinal, especiales) {
  const sh = getSS().getSheetByName(SH_WEIGHTS);
  if (!sh) return { ok:false, error:'no_sheet' };
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]==course && rows[i][1]==period) {
      sh.getRange(i+1, 3, 1, 4).setValues([[pesoAct, pesoAuto, pesoCoeval, pesoFinal]]);
      if (especiales && especiales.length) {
        const spSh = getSS().getSheetByName(SH_SPECIALS);
        const spRows = spSh.getDataRange().getValues();
        especiales.forEach(esp => {
          for (let j = 1; j < spRows.length; j++) {
            if (spRows[j][0]==course && spRows[j][1]==period && spRows[j][2]==esp.id) {
              spSh.getRange(j+1, 5, 1, 3).setValues([[esp.weight, pesoAct, pesoFinal]]); break;
            }
          }
        });
        _invalidate(SH_SPECIALS);
      }
      scheduleRebuild(course, +period);
      writeLog('saveWeights', course, 'Pesos P' + period + ' actualizados');
      return { ok:true };
    }
  }
  sh.appendRow([course, period, pesoAct, pesoAuto, pesoCoeval, pesoFinal]);
  if (especiales && especiales.length) {
    const spSh = getSS().getSheetByName(SH_SPECIALS);
    const spRows = spSh.getDataRange().getValues();
    especiales.forEach(esp => {
      for (let j = 1; j < spRows.length; j++) {
        if (spRows[j][0]==course && spRows[j][1]==period && spRows[j][2]==esp.id) {
          spSh.getRange(j+1, 5, 1, 3).setValues([[esp.weight, pesoAct, pesoFinal]]); break;
        }
      }
    });
    _invalidate(SH_SPECIALS);
  }
  scheduleRebuild(course, +period);
  writeLog('saveWeights', course, 'Pesos P' + period + ' creados');
  return { ok:true };
}

function addSpecial(course, period, name, weight, pesoAct, pesoFinal) {
  weight    = Math.round(+weight * 10) / 10;
  pesoAct   = pesoAct   !== undefined ? Math.round(+pesoAct   * 10) / 10 : '';
  pesoFinal = pesoFinal !== undefined ? Math.round(+pesoFinal * 10) / 10 : '';
  if (!name || isNaN(weight) || weight <= 0 || weight > 90) return { ok:false, error:'invalid' };
  if (pesoAct !== '' && pesoFinal !== '') {
    const total = Math.round((pesoAct + 5 + 5 + pesoFinal + weight) * 10) / 10;
    if (Math.abs(total - 100) > 0.05) return { ok:false, error:'invalid_total', total };
  }
  const specials = getSpecials(course, period);
  const totalSp = specials.reduce((s,e) => s+e.weight, 0);
  if (totalSp + weight > 90) return { ok:false, error:'overflow', available: Math.round((90-totalSp)*10)/10 };
  const sh = getSS().getSheetByName(SH_SPECIALS);
  const id = uid();
  sh.appendRow([course, period, id, name.trim(), weight, pesoAct, pesoFinal]);
  _invalidate(SH_SPECIALS);
  scheduleRebuild(course, +period);
  writeLog('addSpecial', course, 'P' + period + ' · ' + name.trim() + ' (' + weight + '%)');
  return { ok:true, id, name:name.trim(), weight, pesoAct, pesoFinal };
}

function removeSpecial(course, period, specialId) {
  const { sh, rows } = sheetRows(SH_SPECIALS);
  for (let i=rows.length-1;i>=0;i--) {
    if (rows[i][0]==course&&rows[i][1]==period&&rows[i][2]==specialId) sh.deleteRow(i+2);
  }
  _invalidate(SH_SPECIALS);
  const { sh:sg, rows:gr } = sheetRows(SH_GRADES);
  for (let i=gr.length-1;i>=0;i--) {
    if (gr[i][0]==course&&gr[i][1]==period&&gr[i][2]=='especial'&&gr[i][3]==specialId) sg.deleteRow(i+2);
  }
  _invalidate(SH_GRADES);
  scheduleRebuild(course, +period);
  writeLog('removeSpecial', course, 'P' + period + ' · especial ' + specialId + ' eliminado');
  return { ok:true };
}

// ════════════════════════════════════════════════════════════════
// ACTIVITIES
// ════════════════════════════════════════════════════════════════
function removeActivity(course, period, itemId) {
  const { sh, rows } = sheetRows(SH_GRADES);
  for (let i=rows.length-1;i>=0;i--) {
    if (rows[i][0]==course&&rows[i][1]==period&&rows[i][2]=='actividad'&&rows[i][3]==itemId) sh.deleteRow(i+2);
  }
  _invalidate(SH_GRADES);
  scheduleRebuild(course, +period);
  writeLog('removeActivity', course, 'P' + period + ' · actividad "' + itemId + '" eliminada');
  return { ok:true };
}

// ════════════════════════════════════════════════════════════════
// GRADES
// ════════════════════════════════════════════════════════════════
function saveGrade(b) {
  const g = validGrade(b.grade);
  if (g === false) return { ok:false, error:'invalid_grade' };

  const { sh, rows } = sheetRows(SH_GRADES);
  for (let i=0;i<rows.length;i++) {
    if (rows[i][0]==b.course && rows[i][1]==b.period &&
        rows[i][2]==b.component && rows[i][3]==b.itemId &&
        rows[i][4]==b.studentId) {
      if (g===null) sh.deleteRow(i+2);
      else sh.getRange(i+2,6).setValue(g);
      _invalidate(SH_GRADES);
      scheduleRebuild(b.course, +b.period);
      if (b.studentId !== '__col__') writeLog('saveGrade', b.course, 'P' + b.period + ' · ' + b.component + '/' + b.itemId + ' · estudiante ' + b.studentId + ' → ' + (g===null?'(borrada)':g));
      return { ok:true };
    }
  }
  if (g!==null) {
    sh.appendRow([b.course, b.period, b.component, b.itemId, b.studentId, g]);
    _invalidate(SH_GRADES);
    scheduleRebuild(b.course, +b.period);
    if (b.studentId !== '__col__') writeLog('saveGrade', b.course, 'P' + b.period + ' · ' + b.component + '/' + b.itemId + ' · estudiante ' + b.studentId + ' → ' + g);
  } else if (b.studentId === '__col__') {
    sh.appendRow([b.course, b.period, b.component, b.itemId, b.studentId, '']);
    _invalidate(SH_GRADES);
  }
  return { ok:true };
}

function getGrades(course, period) {
  const { rows } = sheetRows(SH_GRADES);
  const result = {};
  rows.filter(r=>r[0]==course&&r[1]==period).forEach(r=>{
    const [,, comp, item, sid, nota] = r;
    if (sid === '__col__') return;
    if (!result[comp]) result[comp]={};
    if (!result[comp][item]) result[comp][item]={};
    result[comp][item][sid] = +nota;
  });
  return result;
}

function getActivityNames(course, period) {
  const { rows } = sheetRows(SH_GRADES);
  const seen = new Set(), names = [];
  rows.filter(r=>r[0]==course&&r[1]==period&&r[2]=='actividad')
      .forEach(r=>{ if(!seen.has(r[3])){ seen.add(r[3]); names.push(r[3]); } });
  return names;
}

// ════════════════════════════════════════════════════════════════
// GET ALL
// ════════════════════════════════════════════════════════════════
function getAll(course, period) {
  const specials  = getSpecials(course, period);
  const baseW     = getWeights(course, period);
  const weights   = computeWeights(specials, baseW);
  const grades    = getGrades(course, period);
  const actNames  = getActivityNames(course, period);
  return {
    students: getStudents(course),
    weights,
    actNames,
    grades
  };
}

// ════════════════════════════════════════════════════════════════
// BACKUP
// ════════════════════════════════════════════════════════════════
// Celda CSV segura: comillas dobladas (una comilla suelta en un nombre
// rompía la estructura de columnas del backup) + protección contra
// "CSV/formula injection" — Excel/Sheets ejecutan como fórmula una celda
// que empieza con =,+,-,@ si el archivo se abre directo.
function csvField(v) {
  let s = (v === null || v === undefined) ? '' : String(v);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return '"' + s.replace(/"/g, '""') + '"';
}
function generateFullCSV() {
  const lines = [];
  const ts = new Date().toLocaleString('es-CO');
  lines.push(`"BACKUP COMPLETO · NotasQM · ${ts}"`);
  lines.push('');
  getActiveCourseCodes().forEach(course => {
    PERIODS.forEach(period => {
      const students = getStudents(course);
      const specials = getSpecials(course, period);
      // ✅ FIX: faltaba pasar baseW — sin esto, un curso/periodo con pesos
      // personalizados (Guardar pesos) exportaba con el split 60/30 por
      // defecto en vez del real, igual que hacía getAll() correctamente.
      const baseW    = getWeights(course, period);
      const weights  = computeWeights(specials, baseW);
      const grades   = getGrades(course, period);
      const actNames = getActivityNames(course, period);
      lines.push(`"=== CURSO ${course} · PERIODO ${period} ==="`);
      const headers = ['#', 'Código', 'Nombre'];
      actNames.forEach((n, i) => headers.push(`Act.${i+1}: ${n}`));
      headers.push(`Prom.Act (${weights.actividades}%)`, 'Autoeval (5%)', 'Coeval (5%)', `Final (${weights.final}%)`);
      specials.forEach(sp => headers.push(`${sp.name} (${sp.weight}%)`));
      headers.push('DEFINITIVA');
      lines.push(headers.map(h => csvField(h)).join(','));
      if (students.length === 0) {
        lines.push('"(sin estudiantes)"');
      } else {
        students.forEach((s, i) => {
          const row = [i+1, s.code||'', s.name];
          const actVals = actNames.map(n => { const v=grades['actividad']?.[n]?.[s.id]; return (v!==undefined&&v!==null)?v:''; });
          actVals.forEach(v => row.push(v));
          const filled = actVals.filter(v=>v!=='');
          // actAvgRaw (sin redondear) es lo que entra a la suma ponderada,
          // igual que grdCalcFinal en index.html — actAvg (redondeado a 1
          // decimal) es solo para la columna "Prom.Act" visible en el CSV.
          // Redondear el intermedio ANTES de sumar es justo lo que hacía
          // que este backup mostrara una definitiva distinta a la app.
          const actAvgRaw = filled.length ? filled.reduce((a,b)=>a+b,0)/filled.length : 0;
          const actAvg = filled.length ? Math.round(actAvgRaw*10)/10 : '';
          row.push(actAvg);
          const ae=grades['autoeval']?.['_']?.[s.id]??'', ce=grades['coeval']?.['_']?.[s.id]??'', fi=grades['final']?.['_']?.[s.id]??'';
          row.push(ae,ce,fi);
          specials.forEach(sp => row.push(grades['especial']?.[sp.id]?.[s.id]??''));
          const totPct=weights.actividades+5+5+weights.final+specials.reduce((s,e)=>s+e.weight,0);
          let sumW2=0;
          sumW2+=actAvgRaw*(weights.actividades/100);
          sumW2+=(ae!==''?ae:0)*(5/100); sumW2+=(ce!==''?ce:0)*(5/100); sumW2+=(fi!==''?fi:0)*(weights.final/100);
          specials.forEach((sp,idx)=>{ const v=row[3+actNames.length+4+idx]??''; sumW2+=(v!==''?v:0)*(sp.weight/100); });
          // Nivelación reemplaza la definitiva calculada, igual que en
          // Acumulado/informes/sistematizadora de index.html — este backup
          // es el registro "oficial" del docente. OJO: NO coincide con lo
          // que ve el estudiante en Consulta·QM/queryStudent, que por
          // política institucional siempre muestra 3.0 en vez del valor
          // real de nivelación — son dos audiencias distintas a propósito.
          const nivelacionGrade = grades['nivelacion']?.['_']?.[s.id] ?? null;
          const def = nivelacionGrade !== null ? nivelacionGrade
            : (totPct>0?Math.round(sumW2/(totPct/100)*10)/10:0);
          row.push(def);
          lines.push(row.map(v => csvField(v)).join(','));
        });
      }
      lines.push(''); _invalidate(SH_GRADES);
    });
  });
  return lines.join('\r\n');
}

// Backup con versiones fechadas + poda automática.
// Antes: siempre sobrescribía el mismo NotasQM_Backup.csv (si un bug corría
// esa noche, el backup "bueno" de ayer se perdía). Ahora: un archivo nuevo
// por corrida, dentro de una carpeta dedicada, conservando solo las últimas
// BACKUP_KEEP_VERSIONS copias.
function getBackupFolder() {
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

function dailyBackup() {
  const csv = generateFullCSV();
  const ts = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd_HHmm');
  const filename = `NotasQM_Backup_${ts}.csv`;
  const folder = getBackupFolder();
  folder.createFile(Utilities.newBlob('﻿'+csv, 'text/csv;charset=utf-8', filename));
  pruneOldBackups(folder);
  writeLog('dailyBackup', '', filename);
  Logger.log('Backup: ' + filename + ' · ' + new Date().toLocaleString('es-CO'));
}

function pruneOldBackups(folder) {
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf('NotasQM_Backup_') === 0) files.push(f);
  }
  files.sort((a, b) => b.getDateCreated().getTime() - a.getDateCreated().getTime());
  files.slice(BACKUP_KEEP_VERSIONS).forEach(f => f.setTrashed(true));
}

function createBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBackup').timeBased().atHour(2).everyDays(1).create();
  Logger.log('Trigger de backup creado.');
}

// ════════════════════════════════════════════════════════════════
// REBUILD COURSE SHEET
// ════════════════════════════════════════════════════════════════
function rebuildCourseSheet(course, period) {
  const ss = getSS();
  const shName = `${course}-P${period}`;
  let sh = ss.getSheetByName(shName);
  if (!sh) sh = ss.insertSheet(shName);
  sh.clearContents(); sh.clearFormats();

  const students = getStudents(course);
  const specials = getSpecials(course, period);
  // ✅ FIX: mismo bug que generateFullCSV() — faltaba baseW.
  const baseW    = getWeights(course, period);
  const weights  = computeWeights(specials, baseW);
  const grades   = getGrades(course, period);
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
    // Ver nota equivalente en generateFullCSV(): actAvgRaw sin redondear
    // entra a la suma ponderada (igual que grdCalcFinal en index.html);
    // actAvg redondeado es solo para la columna "Prom.Act" visible.
    const actAvgRaw = filled.length ? filled.reduce((a,b)=>a+b,0)/filled.length : 0;
    const actAvg = filled.length ? Math.round(actAvgRaw*10)/10 : '';
    row.push(actAvg);
    const ae=grades['autoeval']?.['_']?.[s.id]??'', ce=grades['coeval']?.['_']?.[s.id]??'', fi=grades['final']?.['_']?.[s.id]??'';
    row.push(ae,ce,fi);
    specials.forEach(sp => row.push(grades['especial']?.[sp.id]?.[s.id]??''));
    const totalPctR=weights.actividades+5+5+weights.final+specials.reduce((s,e)=>s+e.weight,0);
    let sumWR=0;
    sumWR+=actAvgRaw*(weights.actividades/100);
    sumWR+=(ae!==''?ae:0)*(5/100); sumWR+=(ce!==''?ce:0)*(5/100); sumWR+=(fi!==''?fi:0)*(weights.final/100);
    specials.forEach((sp,idx)=>{ const v=row[3+actNames.length+4+idx]??''; sumWR+=(v!==''?v:0)*(sp.weight/100); });
    // Nivelación reemplaza la definitiva calculada — esta hoja legible es
    // el registro que queda visible dentro del propio Google Sheet, así
    // que se trata como registro "oficial" igual que generateFullCSV(),
    // no como la tabla de edición en vivo de index.html (esa sí se deja
    // sin nivelación a propósito, para que el docente vea el cálculo
    // original mientras edita).
    const nivelacionGrade = grades['nivelacion']?.['_']?.[s.id] ?? null;
    const definitiva = nivelacionGrade !== null ? nivelacionGrade
      : (totalPctR>0?Math.round(sumWR/(totalPctR/100)*10)/10:0);
    row.push(definitiva);
    return row;
  });

  const allRows = [headers, ...dataRows];
  if (allRows.length > 0) sh.getRange(1,1,allRows.length,headers.length).setValues(allRows);

  const totalCols = headers.length, totalRows = allRows.length;
  sh.getRange(1,1,1,totalCols).setBackground('#0f2027').setFontColor('white').setFontWeight('bold').setWrap(true);
  sh.setFrozenRows(1); sh.setFrozenColumns(3);

  const defCol = totalCols;
  if (dataRows.length > 0) {
    sh.getRange(2,defCol,dataRows.length,1).setFontWeight('bold').setBackground('#e8f5e9');
    dataRows.forEach((row,i) => {
      const val = row[defCol-1]; if (val==='') return;
      let bg = val>=4.6?'#dcfce7':val>=4.0?'#dbeafe':val>=3.0?'#fef9c3':val>=2.0?'#fce5cd':'#fcd2d2';
      sh.getRange(i+2,defCol).setBackground(bg);
    });
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
  const active = getActiveCourseCodes();
  active.forEach(course => { PERIODS.forEach(period => { rebuildCourseSheet(course, period); }); });
  Logger.log(`✅ ${active.length * PERIODS.length} hojas regeneradas.`);
}

// ════════════════════════════════════════════════════════════════
// QUERY STUDENT (público)
// ════════════════════════════════════════════════════════════════
function queryStudent(code) {
  if (!code||code.toString().trim()==='') return { error:'empty' };
  code = code.toString().trim();
  const { rows } = sheetRows(SH_STUDENTS);
  const match = rows.find(r=>r[4]==true&&r[3].toString().trim().toLowerCase()===code.toLowerCase());
  if (!match) return { error:'notfound' };
  const course=match[0], studentId=match[1], name=match[2];
  const config=getConfig(), period=config[course]||1;
  // ✅ FIX: faltaba baseW — mismo bug que generateFullCSV/rebuildCourseSheet;
  // un curso/periodo con pesos personalizados (Guardar pesos) mostraba acá
  // el split 60/30 por defecto en vez del real que ya usa getAll().
  const specials=getSpecials(course,period);
  const baseW=getWeights(course,period);
  const weights=computeWeights(specials,baseW);
  const grades=getGrades(course,period), actNames=getActivityNames(course,period);
  const actGrades=actNames.map(n=>(grades['actividad']?.[n]?.[studentId]??null));
  const validActs=actGrades.filter(g=>g!==null);
  // actAvgRaw (sin redondear) es lo único que debe entrar a la suma ponderada
  // de abajo, igual que grdCalcFinal en index.html. actAvg (redondeado) es
  // para el texto "detail" y para el campo "grade" que se devuelve — este
  // endpoint es una API pública, cualquier consumidor que lea grade directo
  // espera un número limpio de 1 decimal como los demás componentes, no la
  // cola de flotantes de actAvgRaw. Por eso el resumen usa actAvgRaw
  // explícitamente en vez de leer components[0].grade.
  const actAvgRaw=validActs.length>0?validActs.reduce((a,b)=>a+b,0)/validActs.length:null;
  const actAvg=actAvgRaw!==null?Math.round(actAvgRaw*10)/10:null;
  const components=[
    { key:'actividades', label:'Actividades de clase', weight:weights.actividades,
      detail: actNames.length>0?`${validActs.length} de ${actNames.length} notas · promedio ${actAvg!==null?actAvg.toFixed(1):'—'}`:'Sin actividades aún',
      grade: actAvg },
    { key:'autoeval', label:'Autoevaluación', weight:weights.autoeval, grade: grades['autoeval']?.['_']?.[studentId]??null },
    { key:'coeval',   label:'Coevaluación',   weight:weights.coeval,   grade: grades['coeval']?.['_']?.[studentId]??null },
    { key:'final',    label:'Evaluación final', weight:weights.final,  grade: grades['final']?.['_']?.[studentId]??null },
  ];
  weights.especiales.forEach(sp=>{ components.push({ key:'especial_'+sp.id, label:sp.name, weight:sp.weight, grade: grades['especial']?.[sp.id]?.[studentId]??null }); });
  const totalPct=components.reduce((s,c)=>s+c.weight,0);
  let sumW=0;
  components.forEach(c=>{
    const g = c.key==='actividades' ? actAvgRaw : c.grade;
    sumW+=(g!==null?g:0)*(c.weight/100);
  });
  let avg=totalPct>0?Math.round(sumW/(totalPct/100)*10)/10:null;
  // Nivelación: mismo mecanismo genérico de saveGrade (component:'nivelacion',
  // itemId:'_') que usa Code.Consulta.gs — si el docente la diligenció, la
  // definitiva pública se fija en 3.0. Sin esto, este endpoint (que también
  // es público, vía doGet?action=queryStudent) mostraría la nota reprobada
  // original mientras Consulta·QM ya muestra 3.0 para el mismo estudiante.
  const nivelacionGrade = grades['nivelacion']?.['_']?.[studentId] ?? null;
  const nivelado = nivelacionGrade !== null;
  if (nivelado) avg = 3.0;
  return { ok:true, name, course, period, components, avg, totalPct, actCount:actNames.length, nivelado };
}
