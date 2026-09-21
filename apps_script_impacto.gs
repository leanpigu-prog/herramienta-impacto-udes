// ================================================================
// Apps Script — Seguimiento Impacto Institucional UDES 2024-2028
// ================================================================
// PASOS PARA ACTIVAR:
// 1. En la Google Sheet "SeguimientoImpacto_UDES":
//    - Hoja "Datos" (ya existe): id | funcion | lb | va26 | va28 | m26 | m28 | estado | observaciones | timestamp
//      MIGRACIÓN (2026-09-17, columnas m26/m28 nuevas): si la hoja ya existía con el esquema viejo
//      de 8 columnas (id|funcion|lb|va26|va28|estado|observaciones|timestamp), correr UNA VEZ
//      migrarEsquemaDatosM26M28() desde el editor de Apps Script (ver función al final de este
//      archivo) antes de publicar esta versión.
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
const HEADERS             = ['id','funcion','lb','va26','va28','m26','m28','estado','observaciones','timestamp'];

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

// Calcula el siguiente id consecutivo con un prefijo dado (ej. 'P_BACVLL_' -> 'P_BACVLL_004'),
// leyendo el snapshot `all` ya tomado bajo LockService (ver doPost) — nunca se calcula solo con
// datos del navegador, así dos dispositivos nunca pueden calcular el mismo id (2026-09-21,
// corrige colisión reportada: crear un indicador propio desde 2 equipos a la vez podía pisar uno).
function siguienteIdConPrefijo(all, prefijo) {
  let n = 1;
  for (let i = 1; i < all.length; i++) {
    const id = String(all[i][0] || '');
    if (id.startsWith(prefijo)) {
      const num = parseInt(id.slice(prefijo.length), 10);
      if (!isNaN(num) && num >= n) n = num + 1;
    }
  }
  return prefijo + String(n).padStart(3, '0');
}

// Hace upsert por id (columna 1) sobre una hoja dada. Crea la fila de cabeceras si hace falta.
// Debe llamarse siempre con el LockService ya tomado (ver doPost) — lee y escribe el rango
// completo de la hoja de una sola vez (no una llamada a la API de Sheets por fila) para que:
// (a) sea seguro frente a escrituras concurrentes (todo el ciclo lectura-modificación-escritura
//     ocurre dentro de una sola ejecución con el lock tomado), y
// (b) sea rápido incluso con payloads de varias filas (antes: 1 llamada a Sheets por fila).
//
// Campos de control que puede traer cada `row` (no se escriben en el Sheet, `headers` no los
// incluye):
//   _baseTimestamp: timestamp que el cliente tenía cargado para este id. Si no coincide con el
//     timestamp actual de la fila en el Sheet, alguien más escribió después — no se sobreescribe,
//     se reporta en `conflictos` con el valor vigente del servidor (2026-09-21, a pedido de la
//     usuaria: "avisar y no sobrescribir a ciegas" en vez de "gana el último que guarda").
//   _prefijoId / _tempId: si `row[idField]` viene vacío, se asigna un id nuevo con
//     siguienteIdConPrefijo() y se reporta en `asignaciones[_tempId] = idNuevo`.
function upsertRows(ws, headers, rows, idField) {
  if (!ws) throw new Error('La hoja de destino no existe. Créala primero en el Spreadsheet.');
  let all = ws.getDataRange().getValues();
  if (all.length === 0 || all[0][0] !== headers[0]) {
    ws.getRange(1, 1, 1, headers.length).setValues([headers]);
    all = [headers];
  }
  const idxTimestamp = headers.indexOf('timestamp');
  const conflictos = [];
  const asignaciones = {};
  const guardados = [];
  const nuevasFilas = [];
  let huboActualizacionEnSitio = false;
  const totalFilasOriginal = all.length;

  rows.forEach(row => {
    let idActual = row[idField];
    if ((!idActual || String(idActual).trim() === '') && row._prefijoId) {
      idActual = siguienteIdConPrefijo(all, row._prefijoId);
      if (row._tempId) asignaciones[row._tempId] = idActual;
    }

    const idx = all.findIndex((r, i) => i > 0 && String(r[0]) === String(idActual));
    const filaExistente = idx > 0 ? all[idx] : null;

    if (filaExistente && row._baseTimestamp && idxTimestamp >= 0) {
      const valorTsServidor = filaExistente[idxTimestamp];
      const tsServidor = valorTsServidor ? new Date(valorTsServidor).toISOString() : '';
      if (tsServidor && tsServidor !== row._baseTimestamp) {
        conflictos.push({
          id: idActual,
          servidor: Object.fromEntries(headers.map((h, i) => [h, filaExistente[i]]))
        });
        return; // no se aplica este cambio puntual — el resto del payload sigue su curso
      }
    }

    const ts = new Date().toISOString();
    const fila = headers.map(h => h === 'timestamp' ? ts : (h === idField ? idActual : (row[h] ?? '')));
    if (idx > 0) {
      all[idx] = fila;
      huboActualizacionEnSitio = true;
    } else if (idx === -1) {
      nuevasFilas.push(fila);
      // Se añade también a `all` (no solo a nuevasFilas) para que si el MISMO lote trae más de un
      // indicador nuevo con el mismo `_prefijoId` (ej. el usuario crea dos indicadores propios
      // seguidos, antes de que el autosave de 3s del primero alcance a viajar), el siguiente
      // siguienteIdConPrefijo() de este mismo bucle ya vea este id como ocupado y no lo repita.
      all.push(fila);
    }
    guardados.push({ id: idActual, timestamp: ts });
  });

  if (huboActualizacionEnSitio) {
    ws.getRange(1, 1, totalFilasOriginal, headers.length).setValues(all.slice(0, totalFilasOriginal));
  }
  if (nuevasFilas.length) {
    ws.getRange(totalFilasOriginal + 1, 1, nuevasFilas.length, headers.length).setValues(nuevasFilas);
  }
  return { conflictos, asignaciones, guardados };
}

const CACHE_KEY_DATOS = 'doGet_datos_v1';
const CACHE_TTL_SEGUNDOS = 20; // corto a propósito: prioriza frescura sobre ahorro de cuota

function invalidarCache() {
  CacheService.getScriptCache().remove(CACHE_KEY_DATOS);
}

// GET → devuelve los 3 datasets combinados: indicadores institucionales (Datos), indicadores
// de programa (Indicadores_Programa) y el catálogo de programas/directores (Programas).
// Cacheado 20s (CacheService, por script, compartido entre todos los dispositivos/usuarios) para
// que varios equipos entrando casi al mismo tiempo no disparen una relectura completa de las 3
// hojas cada uno — se invalida en cuanto cualquier doPost escribe algo (ver invalidarCache()).
function doGet(e) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheado = cache.get(CACHE_KEY_DATOS);
    if (cacheado) return resp(JSON.parse(cacheado));

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const payload = {
      ok: true,
      data: readSheet(ss, SHEET_NAME, HEADERS),
      data_programa: readSheet(ss, SHEET_NAME_PROGRAMA, HEADERS_PROGRAMA),
      programas: readSheet(ss, SHEET_NAME_CATALOGO, HEADERS_CATALOGO)
    };
    try {
      cache.put(CACHE_KEY_DATOS, JSON.stringify(payload), CACHE_TTL_SEGUNDOS);
    } catch (errCache) {
      // Si el payload supera el límite de tamaño de CacheService (100KB), simplemente no se
      // cachea esta vez — no debe romper la respuesta al cliente.
    }
    return resp(payload);
  } catch (err) {
    return resp({ ok: false, error: err.message });
  }
}

// POST → guarda / actualiza registros (upsert por id). payload.entity decide la hoja destino:
//   'indicador_programa' → Indicadores_Programa (id_padre vincula con un id de IND del front)
//   default / 'indicador_vr' → Datos (comportamiento original, sin cambios)
// Protegido con LockService (2026-09-21): antes, dos doPost casi simultáneos podían leer la hoja
// en el mismo instante y el que terminara de escribir último pisaba silenciosamente al primero
// (carrera lectura-modificación-escritura clásica). Con el lock, las escrituras quedan
// serializadas — la segunda espera a que la primera termine y lee el estado ya actualizado.
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (errLock) {
    return resp({ ok: false, error: 'No se pudo obtener el bloqueo de escritura (demasiadas solicitudes simultáneas). Intenta guardar de nuevo en unos segundos.' });
  }
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action !== 'upsert') return resp({ ok: false, error: 'acción desconocida' });

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let resultado;
    if (payload.entity === 'indicador_programa') {
      resultado = upsertRows(ss.getSheetByName(SHEET_NAME_PROGRAMA), HEADERS_PROGRAMA, payload.rows, 'id');
    } else {
      // renombrar 'obs' -> 'observaciones' para que coincida con HEADERS, igual que antes
      const rows = payload.rows.map(r => ({ ...r, observaciones: r.obs ?? r.observaciones }));
      resultado = upsertRows(ss.getSheetByName(SHEET_NAME), HEADERS, rows, 'id');
    }
    invalidarCache();

    return resp({ ok: true, conflictos: resultado.conflictos, asignaciones: resultado.asignaciones, guardados: resultado.guardados });
  } catch (err) {
    return resp({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
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

// ================================================================
// FUNCIÓN TEMPORAL DE MIGRACIÓN (2026-09-17) — ejecutar UNA SOLA VEZ desde el editor y luego
// borrarla. IMPORTANTE: seleccionar "migrarEsquemaDatosM26M28" en el desplegable de funciones
// (junto al botón ▶) antes de ejecutar.
//
// La hoja "Datos" todavía tiene el esquema viejo de 8 columnas (sin m26/m28), aunque HEADERS ya
// espera el esquema nuevo de 10. Esta función reescribe la hoja al esquema nuevo sin perder datos
// existentes: id/funcion/lb/va26/va28/estado/observaciones/timestamp se copian igual; m26/m28
// quedan vacíos (no existían antes, el VR los completa desde la herramienta).
// ================================================================
function migrarEsquemaDatosM26M28() {
  const ESQUEMA_VIEJO = ['id','funcion','lb','va26','va28','estado','observaciones','timestamp'];
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ws = ss.getSheetByName(SHEET_NAME);
  if (!ws) throw new Error('No existe la hoja ' + SHEET_NAME);
  const all = ws.getDataRange().getValues();
  const hdrActual = all[0].map(String);
  if (JSON.stringify(hdrActual) !== JSON.stringify(ESQUEMA_VIEJO)) {
    throw new Error('El encabezado actual no coincide con el esquema viejo esperado, no se migra. Encabezado actual: ' + hdrActual.join('|'));
  }
  const filas = all.slice(1).filter(r => r.some(v => v !== ''));
  const nuevasFilas = filas.map(r => {
    const o = Object.fromEntries(ESQUEMA_VIEJO.map((h, i) => [h, r[i]]));
    return HEADERS.map(h => (h === 'm26' || h === 'm28') ? '' : (o[h] ?? ''));
  });
  ws.clearContents();
  ws.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (nuevasFilas.length) {
    ws.getRange(2, 1, nuevasFilas.length, HEADERS.length).setValues(nuevasFilas);
  }
  Logger.log('Migradas ' + nuevasFilas.length + ' filas. Nuevo encabezado: ' + HEADERS.join('|'));
}
