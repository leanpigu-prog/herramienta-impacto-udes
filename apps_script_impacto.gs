// ================================================================
// Apps Script — Seguimiento Impacto Institucional UDES 2024-2028
// ================================================================
// PASOS PARA ACTIVAR:
// 1. En la Google Sheet "SeguimientoImpacto_UDES":
//    - Renombra la primera hoja como "Datos"
//    - Fila 1: id | funcion | lb | va26 | va28 | estado | observaciones | timestamp
// 2. Abre Extensiones > Apps Script y pega este código.
// 3. Guarda (Ctrl+S).
// 4. Implementar > Nueva implementación
//    Tipo: Aplicación web | Ejecutar como: Yo | Acceso: Cualquier persona
// 5. Autoriza los permisos cuando los pida.
// 6. Copia la URL de la implementación (empieza con https://script.google.com/macros/s/...)
//    y pégala en herramienta_impacto.html como valor de APPS_SCRIPT_URL.
// ================================================================

const SHEET_ID   = '15HpVcXgHatswxIAj62v8Hi7xlFsse09xaR9p6yNxozY';
const SHEET_NAME = 'Datos';
const HEADERS    = ['id','funcion','lb','va26','va28','estado','observaciones','timestamp'];

// GET → devuelve todos los registros guardados
function doGet(e) {
  try {
    const ws   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const rows = ws.getDataRange().getValues();
    if (rows.length <= 1) return resp({ ok: true, data: [] });
    const hdr  = rows[0];
    const data = rows.slice(1).map(r =>
      Object.fromEntries(hdr.map((h, i) => [h, r[i]]))
    );
    return resp({ ok: true, data });
  } catch (err) {
    return resp({ ok: false, error: err.message });
  }
}

// POST → guarda / actualiza registros (upsert por id)
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action !== 'upsert') return resp({ ok: false, error: 'acción desconocida' });

    const ws  = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const all = ws.getDataRange().getValues();

    // Asegurar cabeceras en fila 1
    if (all.length === 0 || all[0][0] !== 'id') {
      ws.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      all.length = 0; // forzar reindexación
    }

    payload.rows.forEach(row => {
      const ts = new Date().toISOString();
      const fila = [
        row.id       ?? '',
        row.funcion  ?? '',
        row.lb       ?? '',
        row.va26     ?? '',
        row.va28     ?? '',
        row.estado   ?? '',
        row.obs      ?? '',
        ts
      ];
      const idx = all.findIndex((r, i) => i > 0 && String(r[0]) === String(row.id));
      if (idx > 0) {
        ws.getRange(idx + 1, 1, 1, fila.length).setValues([fila]);
      } else {
        ws.appendRow(fila);
      }
    });

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
