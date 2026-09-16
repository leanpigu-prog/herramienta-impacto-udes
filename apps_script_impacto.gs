// ================================================================
// Apps Script — Seguimiento Impacto Institucional UDES 2024-2028
// ================================================================
// PASOS PARA ACTIVAR:
// 1. En la Google Sheet "SeguimientoImpacto_UDES":
//    - Hoja "Datos" (ya existe): id | funcion | lb | va26 | va28 | estado | observaciones | timestamp
//    - Crear hoja nueva "Indicadores_Programa" con fila 1 (esquema completo, igual al catálogo
//      institucional — permite que el director coloque su propia línea base y metas):
//      id | id_padre | funcion | niv | programa | sede | director | nombre | desc | und | lb | m26 | va26 | m28 | va28 | evidencia | estado | lb2021 | timestamp
//      MIGRACIÓN (si la hoja ya existía con el esquema viejo id|id_padre|funcion|programa|sede|
//      director|nombre|desc|und|meta|valor_actual|evidencia|estado|timestamp): agregar las columnas
//      nuevas (niv, lb, m26, va26, m28, va28) y, para cada fila existente, mover manualmente
//      meta→m26 y valor_actual→va26, fijando niv según a qué indicador institucional aporta
//      (columna id_padre). Ej. el registro piloto P_BACVLL_001 (id_padre=I01, nivel Insumos de esa
//      matriz) queda niv="Insumos".
//      MIGRACIÓN 2 (columna lb2021, "Línea Base 2021", agregada aparte de la "LB 2023" existente):
//      si la hoja ya existía sin esta columna, agregarla manualmente en la fila 1 antes de
//      "timestamp" (encabezado exacto "lb2021"); las filas existentes quedan con ese valor vacío
//      hasta que el director lo capture desde la tabla de Mi Programa.
//    - Crear hoja nueva "Programas" con fila 1:
//      codigo | nombre_programa | sede | director | funciones | clave
//      (piloto: 3 filas para Bacteriología — BAC-BGA/BAC-CUC/BAC-VLL, funciones="INV,ENS,EXT")
//    Las hojas nuevas se crean solas como listas vacías si aún no existen — readSheet() no falla,
//    pero doPost sí necesita que la hoja de destino ya exista para poder escribir en ella.
// 2. Abre Extensiones > Apps Script y pega este código.
// 3. Guarda (Ctrl+S).
// 4. Implementar > Nueva implementación (o "Gestionar implementaciones" > editar la existente)
//    Tipo: Aplicación web | Ejecutar como: Yo | Acceso: Cualquier persona
// 5. Autoriza los permisos cuando los pida.
// 6. Copia la URL de la implementación (empieza con https://script.google.com/macros/s/...)
//    y pégala en herramienta_impacto.html como valor de APPS_SCRIPT_URL (si cambia).
// ================================================================

const SHEET_ID = '15HpVcXgHatswxIAj62v8Hi7xlFsse09xaR9p6yNxozY';

const SHEET_NAME          = 'Datos';
const HEADERS             = ['id','funcion','lb','va26','va28','estado','observaciones','timestamp'];

const SHEET_NAME_PROGRAMA = 'Indicadores_Programa';
const HEADERS_PROGRAMA    = ['id','id_padre','funcion','niv','programa','sede','director','nombre','desc','und','lb','m26','va26','m28','va28','evidencia','estado','lb2021','timestamp'];

const SHEET_NAME_CATALOGO = 'Programas';
const HEADERS_CATALOGO    = ['codigo','nombre_programa','sede','director','funciones','clave'];

// Lee una hoja completa y la devuelve como array de objetos, usando la fila 1 como cabeceras.
// Si la hoja no existe todavía, devuelve [] en vez de fallar (permite desplegar el backend
// antes de crear manualmente las hojas nuevas).
// Todas las columnas son texto libre (aceptan cualquier valor) salvo 'timestamp', que sí es
// fecha real. Si Sheets llegó a interpretar una celda de otra columna como fecha (autoformato
// al escribir un valor que "parece" fecha), se devuelve vacío en vez de propagar el objeto Date.
function readSheet(ss, sheetName, fallbackHeaders) {
  const ws = ss.getSheetByName(sheetName);
  if (!ws) return [];
  const rows = ws.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const hdr = rows[0][0] ? rows[0] : fallbackHeaders;
  return rows.slice(1)
    .filter(r => r.some(v => v !== ''))
    .map(r => Object.fromEntries(hdr.map((h, i) => {
      const v = r[i];
      return [h, (h !== 'timestamp' && v instanceof Date) ? '' : v];
    })));
}

// Hace upsert por id (columna 1) sobre una hoja dada. Crea la fila de cabeceras si hace falta.
function upsertRows(ws, headers, rows, idField) {
  if (!ws) throw new Error('La hoja de destino no existe. Créala primero en el Spreadsheet.');
  const all = ws.getDataRange().getValues();
  if (all.length === 0 || all[0][0] !== headers[0]) {
    ws.getRange(1, 1, 1, headers.length).setValues([headers]);
    all.length = 0; // forzar reindexación
  }
  rows.forEach(row => {
    const ts = new Date().toISOString();
    const fila = headers.map(h => h === 'timestamp' ? ts : (row[h] ?? ''));
    const idx = all.findIndex((r, i) => i > 0 && String(r[0]) === String(row[idField]));
    if (idx > 0) {
      ws.getRange(idx + 1, 1, 1, fila.length).setValues([fila]);
    } else {
      ws.appendRow(fila);
    }
  });
}

// GET → devuelve los 3 datasets combinados: indicadores institucionales (Datos), indicadores
// de programa (Indicadores_Programa) y el catálogo de programas/directores (Programas).
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    return resp({
      ok: true,
      data: readSheet(ss, SHEET_NAME, HEADERS),
      data_programa: readSheet(ss, SHEET_NAME_PROGRAMA, HEADERS_PROGRAMA),
      programas: readSheet(ss, SHEET_NAME_CATALOGO, HEADERS_CATALOGO)
    });
  } catch (err) {
    return resp({ ok: false, error: err.message });
  }
}

// POST → guarda / actualiza registros (upsert por id). payload.entity decide la hoja destino:
//   'indicador_programa' → Indicadores_Programa (id_padre vincula con un id de IND del front)
//   default / 'indicador_vr' → Datos (comportamiento original, sin cambios)
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action !== 'upsert') return resp({ ok: false, error: 'acción desconocida' });

    const ss = SpreadsheetApp.openById(SHEET_ID);
    if (payload.entity === 'indicador_programa') {
      upsertRows(ss.getSheetByName(SHEET_NAME_PROGRAMA), HEADERS_PROGRAMA, payload.rows, 'id');
    } else {
      // renombrar 'obs' -> 'observaciones' para que coincida con HEADERS, igual que antes
      const rows = payload.rows.map(r => ({ ...r, observaciones: r.obs ?? r.observaciones }));
      upsertRows(ss.getSheetByName(SHEET_NAME), HEADERS, rows, 'id');
    }

    return resp({ ok: true });
  } catch (err) {
    return resp({ ok: false, error: err.message });
  }
}

function resp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
// FUNCIÓN TEMPORAL DE MIGRACIÓN — ejecutar UNA SOLA VEZ desde el editor y luego borrarla.
// IMPORTANTE: seleccionar "migrarEsquemaIndicadoresPrograma" en el desplegable de funciones
// (junto al botón ▶) antes de ejecutar — si queda seleccionada otra función esto no corre.
//
// La hoja "Indicadores_Programa" real todavía tiene el esquema viejo (sin niv/lb/m26/va26/m28/
// va28, con meta/valor_actual en su lugar), aunque HEADERS_PROGRAMA ya espera el esquema nuevo.
// Esta función reescribe la hoja al esquema nuevo sin perder datos existentes:
//   - id/id_padre/funcion/programa/sede/director/nombre/desc/und/evidencia/estado/timestamp: se
//     copian igual.
//   - meta -> m26, valor_actual -> va26 (mismo valor, solo renombrado de columna).
//   - niv/lb/m28/va28: quedan vacíos (no existían antes).
//   - lb2021: se limpia a vacío — los valores actuales son fechas corruptas producidas por el
//     bug de autoformato de Sheets, no datos reales capturados por un director.
// ================================================================
function migrarEsquemaIndicadoresPrograma() {
  const ESQUEMA_VIEJO = ['id','id_padre','funcion','programa','sede','director','nombre','desc','und','meta','valor_actual','evidencia','estado','lb2021','timestamp'];
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ws = ss.getSheetByName(SHEET_NAME_PROGRAMA);
  if (!ws) throw new Error('No existe la hoja ' + SHEET_NAME_PROGRAMA);
  const all = ws.getDataRange().getValues();
  const hdrActual = all[0].map(String);
  if (JSON.stringify(hdrActual.slice(0, ESQUEMA_VIEJO.length)) !== JSON.stringify(ESQUEMA_VIEJO)) {
    throw new Error('El encabezado actual no coincide con el esquema viejo esperado, no se migra. Encabezado actual: ' + hdrActual.join('|'));
  }
  const filas = all.slice(1).filter(r => r.some(v => v !== ''));
  const nuevasFilas = filas.map(r => {
    const o = Object.fromEntries(ESQUEMA_VIEJO.map((h, i) => [h, r[i]]));
    return HEADERS_PROGRAMA.map(h => {
      if (h === 'niv' || h === 'lb' || h === 'm28' || h === 'va28' || h === 'lb2021') return '';
      if (h === 'm26') return o.meta;
      if (h === 'va26') return o.valor_actual;
      return o[h] ?? '';
    });
  });
  ws.clearContents();
  ws.getRange(1, 1, 1, HEADERS_PROGRAMA.length).setValues([HEADERS_PROGRAMA]);
  if (nuevasFilas.length) {
    ws.getRange(2, 1, nuevasFilas.length, HEADERS_PROGRAMA.length).setValues(nuevasFilas);
  }
  Logger.log('Migradas ' + nuevasFilas.length + ' filas. Nuevo encabezado: ' + HEADERS_PROGRAMA.join('|'));
}
