# HBMOL — Instalación y despliegue

HBMOL es el sucesor de **GestiónQM**, para la gestión de notas y asistencia de Harvey Barcenas Morales (Ciencias Naturales — Química, Colegio Rufino José Cuervo IED). Sigue usando **Google Sheets + Google Apps Script** (sin servidor propio), pero corrige varias debilidades de GestiónQM y agrega funciones nuevas. Ver el análisis completo en [`ANALISIS_Y_DISEÑO.md`](./ANALISIS_Y_DISEÑO.md).

Esta carpeta (`hbmol/`) es **independiente** del resto del repositorio (`index.html`, `Code.gs`, `sw.js`, `manifest.json` en la raíz siguen siendo GestiónQM, sin tocar). Todo lo nuevo vive aquí.

## Qué se heredó vs. qué es nuevo

| | GestiónQM (raíz del repo) | HBMOL (esta carpeta) |
|---|---|---|
| Google Sheets | 2 hojas de cálculo separadas (Notas, Asistencia) | **1 sola hoja de cálculo** ("HBMOL") con pestañas compartidas |
| Roster de estudiantes | Duplicado, un roster por hoja, sin ID compartido | **Un solo roster** (`Estudiantes`), administrado solo desde Notas, leído por Asistencia |
| Cursos | Arreglo `COURSES` hardcodeado en 2-3 archivos | Hoja `Config` dinámica + acciones `getCourses`/`addCourse`/`removeCourse` + panel de administración |
| Auditoría | No existe | Hoja `Log` compartida (fecha, módulo, acción, curso, detalle) |
| Backup | 1 archivo CSV sobrescrito cada noche | Carpeta de Drive con backups fechados y poda automática (últimas 30 versiones) |
| Informe individual | Sí, con análisis narrativo | Igual, conservado (versión simplificada en este scaffold) |
| Informe de grupo/curso | No existe | **Nuevo**: promedio del curso, distribución de desempeños, ranking |
| Modo oscuro | No | **Nuevo** |
| Alertas a acudientes | No | Plantilla lista (`checkLowGradeAlerts` en `Code.Notas.gs`), inactiva por defecto — requiere que tú decidas activarla y llenar los correos |
| Portal de consulta del estudiante | Sí (`queryStudent`, sin contraseña) | Igual, conservado |
| Pesos flexibles / "especiales" | Sí | Igual, conservado sin cambios de lógica |
| Cola de sincronización offline | Sí | Igual, conservado |

## Archivos de esta carpeta

- `Code.Notas.gs` — backend de notas (Apps Script). Fuente de verdad de estudiantes y cursos.
- `Code.Asistencia.gs` — backend de asistencia (Apps Script). Lee el roster y el catálogo de cursos del mismo Sheet; ya no los duplica.
- `index.html` — front-end (PWA). Requiere las URLs de los dos Web Apps de arriba.
- `ANALISIS_Y_DISEÑO.md` — documento de análisis y decisiones de arquitectura.

## Paso a paso para desplegar (todo manual, no hay CI/CD)

### 1. Crear el Google Sheet único

1. Ve a [sheets.google.com](https://sheets.google.com) y crea una hoja de cálculo nueva. Nómbrala, por ejemplo, **"HBMOL"**.
2. Copia su ID (la parte de la URL entre `/d/` y `/edit`), algo como:
   `https://docs.google.com/spreadsheets/d/`**`1AbCdEfGhIjKlMnOpQrStUvWxYz`**`/edit`

### 2. Desplegar el backend de Notas

1. En el mismo Google Sheet, ve a **Extensiones → Apps Script**.
2. Borra el contenido del archivo `Code.gs` por defecto y pega **todo** el contenido de `hbmol/Code.Notas.gs`.
3. Reemplaza la constante `SPREADSHEET_ID_HBMOL` con el ID que copiaste en el paso 1.
4. (Opcional pero recomendado) Cambia la constante `PASSWORD` por una contraseña propia, distinta de `951220`.
5. En el editor, selecciona la función `setup` en el desplegable de funciones y ejecútala (▶). La primera vez te pedirá autorizar permisos — acéptalos (son tuyos, sobre tu propio Sheet).
6. Verifica en el Sheet que se crearon las pestañas: `Config`, `Estudiantes`, `Notas`, `Especiales`, `Pesos`, `Log`, y las hojas-vista `1004-P1`, `1004-P2`, etc. (los 7 cursos por defecto). Si quieres otros cursos desde el arranque, edita `DEFAULT_COURSES` en el código ANTES de correr `setup()` por primera vez (después, usa el panel de administración).
7. Despliega como aplicación web: **Implementar → Nueva implementación → Aplicación web**.
   - Ejecutar como: **Yo (tu cuenta)**.
   - Quién tiene acceso: **Cualquier usuario** (o **Cualquier usuario con una cuenta de Google**, si quieres una capa extra de fricción para desconocidos — ver `ANALISIS_Y_DISEÑO.md`, sección de seguridad).
8. Copia la URL que te entrega (termina en `/exec`). Esa es tu `URL_GRD`.

### 3. Desplegar el backend de Asistencia — MISMO Sheet, proyecto Apps Script aparte

Este es el punto más importante de no confundir: **NO es una hoja de cálculo nueva**. Es un **segundo proyecto de Apps Script** que abre el **mismo** Google Sheet por ID.

1. Todavía en el mismo Sheet, o desde [script.google.com](https://script.google.com) → "Proyecto nuevo" (standalone), crea un proyecto Apps Script nuevo, separado del de Notas.
2. Pega todo el contenido de `hbmol/Code.Asistencia.gs`.
3. Reemplaza `SPREADSHEET_ID_HBMOL` con el **mismo ID** del paso 1 (idéntico al de Notas).
4. Corre `setup()` de este proyecto. Si te sale la advertencia de "faltan las hojas Estudiantes/Config", significa que olvidaste correr primero el `setup()` de Notas — hazlo y vuelve a correr este.
5. Despliega como aplicación web igual que en el paso 2.7. Copia la URL — esa es tu `URL_ATT`.

### 4. Configurar el front-end

1. Abre `hbmol/index.html` en un editor de texto.
2. Reemplaza:
   ```js
   var URL_ATT = 'PEGA_AQUI_LA_URL_DEL_WEBAPP_DE_ASISTENCIA';
   var URL_GRD = 'PEGA_AQUI_LA_URL_DEL_WEBAPP_DE_NOTAS';
   ```
   con las dos URLs `/exec` que obtuviste.
3. Copia `manifest.json`, `sw.js` y `hbmol.png` (están en la raíz del repo) dentro de esta carpeta `hbmol/`, o ajusta las rutas en el `<head>` de `index.html` si prefieres que sigan sirviéndose desde la raíz. Esto es necesario para que HBMOL funcione como PWA instalable independiente de GestiónQM.
4. Sube `hbmol/` (o solo su contenido) al hosting que uses hoy para GestiónQM (GitHub Pages, Netlify, lo que sea) — puede convivir en una subcarpeta o en un repositorio/sitio aparte, según prefieras.

### 5. Activar el backup con versiones (recomendado, ya viene en el código)

1. En el proyecto Apps Script de **Notas**, selecciona la función `createBackupTrigger` y ejecútala una vez. Esto crea un disparador diario a las 2am que corre `dailyBackup()`.
2. Los backups aparecerán en una carpeta de tu Drive llamada `HBMOL_Backups`, con nombre `HBMOL_Backup_YYYY-MM-DD_HHMM.csv`. Se conservan automáticamente las últimas 30 versiones.
3. **Complemento gratuito sin código:** en el propio Google Sheet, usa **Archivo → Historial de versiones → Nombrar la versión actual** al cerrar cada periodo. Es una segunda red de seguridad nativa de Google, independiente de estos backups CSV.

### 6. (Opcional) Activar alertas automáticas a acudientes

Esta función viene como **plantilla desactivada** en `Code.Notas.gs` (`checkLowGradeAlerts`) porque implica manejar correos de contacto de terceros (acudientes de menores de edad) — decisión que debes tomar tú, no algo que active el código por defecto.

Para activarla:
1. Crea manualmente una hoja llamada `Acudientes` en el Sheet, con columnas `Curso | EstudianteID | EmailAcudiente` (el `EstudianteID` es el mismo ID que aparece en la hoja `Estudiantes`).
2. Llénala con los correos que quieras notificar (puedes empezar solo con los estudiantes en riesgo).
3. Crea un disparador de tiempo (`Triggers` → añadir disparador) para `checkLowGradeAlerts`, ejecutándose, por ejemplo, cada semana. Ajusta `threshold` según el umbral de nota que consideres relevante.

### 7. Primer uso

1. Abre `index.html` desplegado, ingresa con la contraseña de Notas.
2. Ve al módulo **Cursos** y confirma que aparecen los 7 cursos por defecto (o los que hayas configurado). Desde ahí puedes crear/desactivar cursos sin tocar código nunca más.
3. Agrega estudiantes desde el módulo de Notas (no hay pantalla dedicada de alta masiva en este scaffold — usa la acción `addStudents` vía la consola de Apps Script, o extiende `index.html` con un formulario; ver "Qué falta por construir" abajo).

## Qué queda pendiente / responsabilidad manual del usuario

Este scaffold se generó **sin acceso a un entorno real de Apps Script** para desplegar y probar en vivo (no hay forma de ejecutar Google Apps Script fuera de Google). Por lo tanto:

- **Revisa el código por inspección antes de usarlo con datos reales.** Se conservó la lógica de negocio de GestiónQM casi intacta (cálculo de pesos, definitiva, backup), pero no se ejecutó contra un Sheet real.
- **No hay pantalla de alta masiva de estudiantes ni de traslado entre cursos en este `index.html`** — el backend ya expone `addStudents`/`transferStudent`/`updateStudent`, pero el front-end de este scaffold no incluye esas pantallas (a diferencia de GestiónQM, que sí las tiene completas). Si las necesitas de inmediato, se pueden portar directamente desde `gestionqm/index.html` (los flujos `addStudent`/`transferStudent` de ese archivo llaman a la misma forma de API).
- **La generación masiva de PDFs de boletines, la integración con Google Classroom y el archivo de año lectivo** están documentados como ideas evaluadas en `ANALISIS_Y_DISEÑO.md`, pero no están implementadas — quedan como siguientes pasos.
- **Las contraseñas siguen siendo un secreto compartido en texto plano** dentro del código fuente de cada Apps Script — HBMOL no resuelve esto de raíz (ver limitaciones documentadas), solo mitiga con la opción de restringir el despliegue a "cualquier cuenta de Google".
- Verifica los límites de cuota de Apps Script (ejecución de 6 minutos, cuotas diarias de `MailApp`/`UrlFetch`) si en algún momento creces por encima de ~250-300 estudiantes o activas las alertas por correo para todos los cursos a la vez.
