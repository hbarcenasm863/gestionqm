// ══════════════════════════════════════════════════════════════════
// ASISTENCIA·QM — Code.gs
// Google Sheet: https://docs.google.com/spreadsheets/d/1_unsXdtPoe3_lqXlV3X3U4VbS9os5_hQMhntIumk-CQ
//
// Copia de referencia versionada en git del Code.gs real, que vive en el
// proyecto de Apps Script vinculado al Sheet de arriba. Después de editar
// aquí, pega el contenido completo en el editor de Apps Script y despliega
// una NUEVA VERSIÓN (Implementar → Administrar implementaciones → Editar →
// Nueva versión) — guardar el script no actualiza la URL en producción.
//
// Cambio respecto a la versión anterior: log de auditoría simple (hoja Log),
// igual que en Code.Notas.gs, escrito en altas/bajas de estudiante y en
// cada marcado de asistencia.
//
// Nota: este backend no necesita una lista dinámica de cursos porque nunca
// validó COURSES_ATT contra nada — ya acepta cualquier código de curso que
// reciba. Los cursos se administran únicamente desde Code.Notas.gs.
// ══════════════════════════════════════════════════════════════════

const PASSWORD_ATT = '951220';  // ← MISMA que en Notas

const SPREADSHEET_ID_ATT = '1_unsXdtPoe3_lqXlV3X3U4VbS9os5_hQMhntIumk-CQ';
function getSS_ATT() { return SpreadsheetApp.openById(SPREADSHEET_ID_ATT); }

const SH_ATT_STUDENTS   = 'Estudiantes';  // Curso | ID | Nombre | Activo
const SH_ATT_ATTENDANCE = 'Asistencia';   // Curso | Fecha | EstudianteID | Estado
const SH_ATT_LOG        = 'Log';          // Fecha | Accion | Curso | Detalle

// ════════════════════════════════════════════════════════════════
// SETUP — ejecutar solo una vez
// ════════════════════════════════════════════════════════════════
function setup() {
  const ss = getSS_ATT();

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

  ensure(SH_ATT_STUDENTS,   ['Curso','ID','Nombre','Activo']);
  ensure(SH_ATT_ATTENDANCE, ['Curso','Fecha','EstudianteID','Estado']);
  ensure(SH_ATT_LOG,        ['Fecha','Accion','Curso','Detalle']);

  Logger.log('Setup de Asistencia completado.');
}

// ════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════
function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function sheetRows(name) {
  const sh = getSS_ATT().getSheetByName(name);
  return { sh, rows: sh.getDataRange().getValues().slice(1) };
}
function uid() {
  return 'a' + Date.now().toString(36) + Math.random().toString(36).slice(2,6);
}

function writeLog(action, course, detail) {
  try {
    const sh = getSS_ATT().getSheetByName(SH_ATT_LOG);
    if (!sh) return;
    sh.appendRow([new Date(), action, course || '', detail || '']);
  } catch(e) { /* no interrumpir la operación principal */ }
}

// ════════════════════════════════════════════════════════════════
// ROUTER
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  const p = e.parameter;
  if (p.pwd !== PASSWORD_ATT) return jsonOut({ error: 'auth' });
  try {
    if (p.action === 'getStudents')        return jsonOut(getStudents(p.course));
    if (p.action === 'getAllAttendance')    return jsonOut(getAllAttendance(p.course));
    if (p.action === 'getAttendance')      return jsonOut(getAttendance(p.course, p.date));
    if (p.action === 'getAttendanceDate')  return jsonOut(getAttendanceDate(p.course, p.date));
    if (p.action === 'getAttInit')         return jsonOut(getAttInit(p.course, p.date));
    if (p.action === 'debug')              return jsonOut(debugSheet(p.course));
    return jsonOut({ error: 'unknown' });
  } catch(e) { return jsonOut({ error: e.message }); }
}

// Diagnóstico: primeras 5 filas crudas de asistencia para un curso
function debugSheet(course) {
  const { rows } = sheetRows(SH_ATT_ATTENDANCE);
  const courseStr = String(course).trim();
  const all = rows.filter(r => String(r[0]).trim() === courseStr);
  const sample = all.slice(0, 3).map(r => ({
    col0: String(r[0]),
    col0type: typeof r[0],
    col1isDate: r[1] instanceof Date,
    col1type: typeof r[1],
    col1raw: JSON.stringify(r[1]),
    col1norm: normDate(r[1]),
    col2: String(r[2]),
    col3: r[3]
  }));
  const attResult = getAllAttendance(course);
  const attKeys = Object.keys(attResult).slice(0, 5);
  return {
    totalRows: rows.length,
    matching: all.length,
    sample,
    courseStr,
    allAttKeys: attKeys,
    allAttCount: Object.keys(attResult).length
  };
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.pwd !== PASSWORD_ATT) return jsonOut({ error: 'auth' });
  try {
    if (b.action === 'addStudent')      return jsonOut(addStudent(b.course, b.name));
    if (b.action === 'addStudents')     return jsonOut(addStudents(b.course, b.names));
    if (b.action === 'removeStudent')   return jsonOut(removeStudent(b.course, b.studentId));
    if (b.action === 'markAttendance')  return jsonOut(markAttendance(b.course, b.date, b.studentId, b.status));
    if (b.action === 'markAll')         return jsonOut(markAll(b.course, b.date, b.records));
    return jsonOut({ error: 'unknown' });
  } catch(e) { return jsonOut({ error: e.message }); }
}

// ════════════════════════════════════════════════════════════════
// STUDENTS
// ════════════════════════════════════════════════════════════════

// Normaliza fecha a string YYYY-MM-DD — versión infalible
function normDate(d) {
  if (!d && d !== 0) return '';
  if (typeof d === 'object') {
    try {
      return Utilities.formatDate(d, 'America/Bogota', 'yyyy-MM-dd');
    } catch(e1) {
      try {
        var iso = JSON.stringify(d).replace(/"/g, '');
        if (iso && iso.indexOf('T') !== -1) return iso.substring(0, 10);
        if (/^\d{4}-\d{2}-\d{2}/.test(iso)) return iso.substring(0, 10);
      } catch(e2) {}
    }
  }
  if (typeof d === 'number') {
    try {
      var dt = new Date(Math.round((d - 25569) * 86400 * 1000));
      return Utilities.formatDate(dt, 'America/Bogota', 'yyyy-MM-dd');
    } catch(e) {}
  }
  var s = String(d).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  if (s.indexOf('T') !== -1) return s.split('T')[0].substring(0, 10);
  var meses = {Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',
               Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
  var m1 = s.match(/(\w{3})\s+(\d{1,2})\s+(\d{4})/);
  if (m1 && meses[m1[1]]) return m1[3] + '-' + meses[m1[1]] + '-' + m1[2].padStart(2,'0');
  var m2 = s.match(/(\w{3})\s+(\d{1,2})/);
  if (m2 && meses[m2[1]]) {
    var y = new Date().getFullYear();
    return y + '-' + meses[m2[1]] + '-' + m2[2].padStart(2,'0');
  }
  return '';
}

function getStudents(course) {
  const { rows } = sheetRows(SH_ATT_STUDENTS);
  const courseStr = String(course).trim();
  return rows
    .filter(r => String(r[0]).trim() === courseStr && r[3] == true)
    .map(r => ({ id: r[1], name: r[2] }));
}

function addStudent(course, name) {
  const sh = getSS_ATT().getSheetByName(SH_ATT_STUDENTS);
  const id = uid();
  sh.appendRow([course, id, name.trim(), true]);
  writeLog('addStudent', course, name.trim());
  return { ok: true, id, name: name.trim() };
}

function addStudents(course, names) {
  const sh = getSS_ATT().getSheetByName(SH_ATT_STUDENTS);
  const added = [];
  names.forEach(name => {
    name = name.trim(); if (!name) return;
    const id = uid();
    sh.appendRow([course, id, name, true]);
    added.push({ id, name });
    Utilities.sleep(5);
  });
  writeLog('addStudents', course, added.length + ' estudiante(s) importado(s)');
  return { ok: true, added };
}

function removeStudent(course, studentId) {
  const { sh, rows } = sheetRows(SH_ATT_STUDENTS);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] == course && rows[i][1] == studentId) {
      sh.getRange(i+2, 4).setValue(false);
      writeLog('removeStudent', course, 'Estudiante ' + rows[i][2] + ' retirado');
      return { ok: true };
    }
  }
  return { ok: false };
}

// ════════════════════════════════════════════════════════════════
// ATTENDANCE
// ════════════════════════════════════════════════════════════════
function getAttendance(course, date) {
  const { rows } = sheetRows(SH_ATT_ATTENDANCE);
  const result = {};
  rows.filter(r => r[0]==course && r[1]==date)
      .forEach(r => result[r[2]] = r[3]);
  return result;
}

function getAttendanceDate(course, date) {
  date = normDate(date);
  if (!course || !date) return {};
  const { rows } = sheetRows(SH_ATT_ATTENDANCE);
  const result = {};
  rows
    .filter(r => String(r[0]) === String(course) && normDate(r[1]) === date)
    .forEach(r => { if (r[2] && r[3]) result[String(r[2])] = String(r[3]); });
  return result;
}

function getAttInit(course, date) {
  date = normDate(date);
  return {
    students: getStudents(course),
    dayAtt:   getAttendanceDate(course, date)
  };
}

function getAllAttendance(course) {
  const { rows } = sheetRows(SH_ATT_ATTENDANCE);
  const result = {};
  const courseStr = String(course).trim();
  rows.filter(r => String(r[0]).trim() === courseStr).forEach(r => {
    const date = normDate(r[1]);
    if (!date) return;
    if (!result[date]) result[date] = {};
    result[date][String(r[2])] = r[3];
  });
  return result;
}

function markAttendance(course, date, studentId, status) {
  date = normDate(date);
  const { sh, rows } = sheetRows(SH_ATT_ATTENDANCE);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0]==course && normDate(rows[i][1])==date && rows[i][2]==studentId) {
      if (!status) sh.deleteRow(i+2);
      else sh.getRange(i+2, 4).setValue(status);
      return { ok: true };
    }
  }
  if (status) sh.appendRow([course, date, studentId, status]);
  return { ok: true };
}

function markAll(course, date, records) {
  records.forEach(r => markAttendance(course, date, r.studentId, r.status));
  writeLog('markAll', course, records.length + ' registro(s) · ' + normDate(date));
  return { ok: true };
}
