// ══════════════════════════════════════════════════════════════════
// HBMOL · Code.Asistencia.gs
// Sucesor de "ASISTENCIA·QM". Google Apps Script standalone (Web App)
// INDEPENDIENTE de Code.Notas.gs (su propio deployment/URL/password),
// pero apuntando al MISMO Google Sheet ("HBMOL").
//
// DECISIÓN CLAVE (ver ANALISIS_Y_DISEÑO.md §d y el encabezado de
// Code.Notas.gs): la hoja "Estudiantes" YA NO es propia de este script.
// Es la misma hoja que administra Code.Notas.gs. Esto significa:
//   - Los IDs de estudiante son los MISMOS en notas y en asistencia
//     (arregla el defecto de GestiónQM: rosters independientes por
//     nombre, sin ID compartido — un estudiante podía existir con
//     nombres ligeramente distintos en cada sistema).
//   - Este archivo YA NO define addStudent/addStudents/removeStudent:
//     el alta/baja/edición de estudiantes se hace SOLO desde el módulo
//     de Notas (Code.Notas.gs), para tener una única fuente de verdad
//     y evitar que las dos apps escriban el roster por caminos distintos.
//     Si igual necesitas gestionar estudiantes sin abrir Notas, hazlo
//     directo en la hoja "Estudiantes" del Sheet — es la misma hoja.
//   - "Config" (cursos/color/activo) también es la misma hoja: getCourses()
//     aquí es de SOLO LECTURA; addCourse/removeCourse viven únicamente
//     en Code.Notas.gs.
//   - "Log" (auditoría) también es compartida: lo que se marca aquí
//     aparece en la misma hoja que usa Notas, con Módulo='Asistencia'.
//
// Reemplaza a: gestionqm/Code.gs (backend de asistencia, NO TOCADO).
// ══════════════════════════════════════════════════════════════════

const PASSWORD_ATT = '951220'; // TODO: cámbiala. Ideal: distinta de PASSWORD
                                // de Notas si quieres separar quién puede
                                // tocar qué módulo (ver ANALISIS_Y_DISEÑO.md).

// TODO: EL MISMO ID que pegaste en Code.Notas.gs → SPREADSHEET_ID_HBMOL.
const SPREADSHEET_ID_HBMOL = 'PEGA_AQUI_EL_ID_DEL_SHEET_HBMOL';
function getSS_ATT() { return SpreadsheetApp.openById(SPREADSHEET_ID_HBMOL); }

const SH_ATT_STUDENTS   = 'Estudiantes'; // COMPARTIDA — la misma que Notas usa. Curso|ID|Nombre|Código|Activo
const SH_ATT_ATTENDANCE = 'Asistencia';  // Curso | Fecha | EstudianteID | Estado
const SH_ATT_CONFIG     = 'Config';      // COMPARTIDA — solo lectura desde aquí
const SH_ATT_LOG        = 'Log';         // COMPARTIDA

function setup() {
  const ss = getSS_ATT();
  function ensure(name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); sh.getRange(1,1,1,headers.length).setValues([headers]); sh.setFrozenRows(1); sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a2e5a').setFontColor('white'); }
    return sh;
  }
  // Esta función NO crea "Estudiantes", "Config" ni "Log": esas hojas las
  // crea Code.Notas.gs → setup(). Ejecuta ese setup() PRIMERO. Aquí solo
  // se asegura la hoja propia de este módulo.
  ensure(SH_ATT_ATTENDANCE, ['Curso','Fecha','EstudianteID','Estado']);
  if (!ss.getSheetByName(SH_ATT_STUDENTS) || !ss.getSheetByName(SH_ATT_CONFIG)) {
    Logger.log('⚠ Faltan las hojas compartidas "Estudiantes"/"Config". Corre primero setup() en Code.Notas.gs sobre este mismo Sheet.');
  } else {
    Logger.log('✅ Setup de HBMOL (Asistencia) completado.');
  }
}

function jsonOut(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function sheetRows(name) { const sh = getSS_ATT().getSheetByName(name); return { sh, rows: sh ? sh.getDataRange().getValues().slice(1) : [] }; }

function logAudit(action, course, detail) {
  try {
    const sh = getSS_ATT().getSheetByName(SH_ATT_LOG);
    if (!sh) return;
    sh.appendRow([new Date(), 'Asistencia', action, course || '', detail || '']);
  } catch (e) { /* no tumbar la operación principal por un fallo de log */ }
}

function doGet(e) {
  const p = e.parameter;
  if (p.pwd !== PASSWORD_ATT) return jsonOut({ error: 'auth' });
  try {
    if (p.action === 'getStudents') return jsonOut(getStudents(p.course));
    if (p.action === 'getAllAttendance') return jsonOut(getAllAttendance(p.course));
    if (p.action === 'getAttendance') return jsonOut(getAttendance(p.course, p.date));
    if (p.action === 'getAttendanceDate') return jsonOut(getAttendanceDate(p.course, p.date));
    if (p.action === 'getAttInit') return jsonOut(getAttInit(p.course, p.date));
    if (p.action === 'getCourses') return jsonOut(getCourses()); // solo lectura, ver nota de arriba
    if (p.action === 'debug') return jsonOut(debugSheet(p.course));
    return jsonOut({ error: 'unknown' });
  } catch(e) { return jsonOut({ error: e.message }); }
}
function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.pwd !== PASSWORD_ATT) return jsonOut({ error: 'auth' });
  try {
    // NOTA: ya NO hay addStudent/addStudents/removeStudent aquí a propósito.
    // Gestiona el roster desde el módulo de Notas (misma hoja "Estudiantes").
    if (b.action === 'addStudent' || b.action === 'addStudents' || b.action === 'removeStudent') {
      return jsonOut({ error: 'gestiona_estudiantes_desde_notas', detail: 'El roster es compartido: usa el módulo de Notas (Code.Notas.gs) para altas/bajas/ediciones de estudiantes. Aquí solo se registra asistencia.' });
    }
    if (b.action === 'markAttendance') return jsonOut(markAttendance(b.course, b.date, b.studentId, b.status));
    if (b.action === 'markAll') return jsonOut(markAll(b.course, b.date, b.records));
    return jsonOut({ error: 'unknown' });
  } catch(e) { return jsonOut({ error: e.message }); }
}

function normDate(d) {
  if (!d && d !== 0) return '';
  if (typeof d === 'object') {
    try { return Utilities.formatDate(d, 'America/Bogota', 'yyyy-MM-dd'); }
    catch(e1) { try { var iso = JSON.stringify(d).replace(/"/g, ''); if (iso && iso.indexOf('T') !== -1) return iso.substring(0, 10); if (/^\d{4}-\d{2}-\d{2}/.test(iso)) return iso.substring(0, 10); } catch(e2) {} }
  }
  if (typeof d === 'number') { try { var dt = new Date(Math.round((d - 25569) * 86400 * 1000)); return Utilities.formatDate(dt, 'America/Bogota', 'yyyy-MM-dd'); } catch(e) {} }
  var s = String(d).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  if (s.indexOf('T') !== -1) return s.split('T')[0].substring(0, 10);
  var meses = {Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
  var m1 = s.match(/(\w{3})\s+(\d{1,2})\s+(\d{4})/); if (m1 && meses[m1[1]]) return m1[3] + '-' + meses[m1[1]] + '-' + m1[2].padStart(2,'0');
  var m2 = s.match(/(\w{3})\s+(\d{1,2})/); if (m2 && meses[m2[1]]) { var y = new Date().getFullYear(); return y + '-' + meses[m2[1]] + '-' + m2[2].padStart(2,'0'); }
  return '';
}

// ── Lectura del roster compartido (SIN escritura — ver nota de cabecera) ──
function getStudents(course) { const { rows } = sheetRows(SH_ATT_STUDENTS); const courseStr = String(course).trim(); return rows.filter(r => String(r[0]).trim() === courseStr && r[4] == true).map(r => ({ id: r[1], name: r[2] })); }

// ── Lectura del catálogo de cursos compartido (SIN escritura) ──
function getCourses() {
  const { rows } = sheetRows(SH_ATT_CONFIG);
  return rows.map(r => ({ course: String(r[0]), activePeriod: +r[1] || 1, color: r[2] || '#1a3d6b', active: r[3] == true }));
}

function getAttendance(course, date) { const { rows } = sheetRows(SH_ATT_ATTENDANCE); const result = {}; rows.filter(r => r[0]==course && r[1]==date).forEach(r => result[r[2]] = r[3]); return result; }
function getAttendanceDate(course, date) {
  date = normDate(date); if (!course || !date) return {};
  const { rows } = sheetRows(SH_ATT_ATTENDANCE); const result = {};
  rows.filter(r => String(r[0]) === String(course) && normDate(r[1]) === date).forEach(r => { if (r[2] && r[3]) result[String(r[2])] = String(r[3]); });
  return result;
}
function getAttInit(course, date) { date = normDate(date); return { students: getStudents(course), dayAtt: getAttendanceDate(course, date) }; }
function getAllAttendance(course) {
  const { rows } = sheetRows(SH_ATT_ATTENDANCE); const result = {}; const courseStr = String(course).trim();
  rows.filter(r => String(r[0]).trim() === courseStr).forEach(r => { const date = normDate(r[1]); if (!date) return; if (!result[date]) result[date] = {}; result[date][String(r[2])] = r[3]; });
  return result;
}
function markAttendance(course, date, studentId, status) {
  date = normDate(date);
  const { sh, rows } = sheetRows(SH_ATT_ATTENDANCE);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0]==course && normDate(rows[i][1])==date && rows[i][2]==studentId) {
      if (!status) { sh.deleteRow(i+2); logAudit('markAttendance', course, date + ' · estudiante ' + studentId + ' → (borrado)'); }
      else { sh.getRange(i+2, 4).setValue(status); logAudit('markAttendance', course, date + ' · estudiante ' + studentId + ' → ' + status); }
      return { ok: true };
    }
  }
  if (status) { sh.appendRow([course, date, studentId, status]); logAudit('markAttendance', course, date + ' · estudiante ' + studentId + ' → ' + status + ' (nuevo)'); }
  return { ok: true };
}
function markAll(course, date, records) {
  records.forEach(r => markAttendance(course, date, r.studentId, r.status));
  logAudit('markAll', course, normDate(date) + ' · ' + records.length + ' registro(s)');
  return { ok: true };
}

function debugSheet(course) {
  const { rows } = sheetRows(SH_ATT_ATTENDANCE);
  return { totalRows: rows.length, sample: rows.filter(r => String(r[0]).trim() === String(course).trim()).slice(0, 5) };
}
