/**
 * Atenea Web App - Transferencias + NO MATCH + Dashboard + Carga CSV
 * Estados: Pendiente | Match | Sin match | Estafador
 * Matching: ESTRICTO (nombre exacto normalizado + (si hay CUIL) CUIL exacto + monto exacto + ventana tiempo)
 */

const TZ = 'America/Argentina/Buenos_Aires';
const SCRIPT_PROPS = PropertiesService.getScriptProperties();

const SHEET = {
  USUARIOS: 'Usuarios',
  TRANSF: 'Transferencias_Operativas',
  MOVS: 'Movimientos_Billetera',
  CONFIG: 'Config',
  LOGS: 'Logs',
  CONTROL: 'Control'
};

const HEADERS = {
  [SHEET.USUARIOS]: ['user_id','nombre_canonico','cuit_cuil','alias_billetera','organizacion','created_at','estado', 'user_id_beast'],
  [SHEET.TRANSF]: ['transfer_id','user_id','fecha_hora_operativa','monto','cajero','comprobante_url','comprobante_file_id','comprobante_uploaded_at','estado','nota'],
  [SHEET.MOVS]: ['Unnamed: 0','raya.cash','Fecha','ID','Monto','Destinario/Origen','CUIL/CUIT','origen_normalizado','imported_at','raw'],
  [SHEET.CONFIG]: ['key','value'],
  [SHEET.LOGS]: ['ts','action','details','actor'],
};

function doGet(e) {
  ensureAllSheets_();
  const t = HtmlService.createTemplateFromFile('app');
  t.appTitle = 'Atenea · Transferencias';
  return t.evaluate()
    .setTitle(t.appTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Ejecutalo 1 vez desde el editor (ideal). */
function setupOnce() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Abrí el Google Sheet vinculado y ejecutá setupOnce().');
  SCRIPT_PROPS.setProperty('SPREADSHEET_ID', ss.getId());
  ensureAllSheets_();
  return { ok: true, spreadsheet_id: ss.getId() };
}

function getSs_() {
  const id = SCRIPT_PROPS.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    SCRIPT_PROPS.setProperty('SPREADSHEET_ID', ss.getId());
    return ss;
  }
  throw new Error('No pude determinar el Spreadsheet. Ejecutá setupOnce() desde el editor.');
}

function ensureSheet_(name) {
  const ss = getSs_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  
  const headers = HEADERS[name];
  if (!headers || headers.length === 0) return sh;

  const needCols = headers.length;
  
  // Esto evita el error "out of bounds" cuando la hoja está vacía (0 datos).
  if (sh.getMaxColumns() < needCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
  }

  // Si la hoja está vacía (porque borramos historial), volvemos a poner los encabezados
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, needCols).setValues([headers]);
    try { sh.setFrozenRows(1); } catch(e){}
  }
  
  return sh;
}

function ensureAllSheets_() {
  ensureSheet_(SHEET.USUARIOS, HEADERS[SHEET.USUARIOS]);
  ensureSheet_(SHEET.TRANSF, HEADERS[SHEET.TRANSF]);
  ensureSheet_(SHEET.MOVS, HEADERS[SHEET.MOVS]);
  ensureSheet_(SHEET.CONFIG, HEADERS[SHEET.CONFIG]);
  ensureSheet_(SHEET.LOGS, HEADERS[SHEET.LOGS]);
  ensureConfigDefaults_();
}

function ensureConfigDefaults_() {
  const shCfg = sh_(SHEET.CONFIG);
  const last = shCfg.getLastRow();
  const existing = new Set();

  if (last >= 2) {
    const vals = shCfg.getRange(2, 1, last - 1, 1).getValues();
    vals.forEach(r => {
      const k = String(r[0] || '').trim();
      if (k) existing.add(k);
    });
  }

  const defaults = [
    ['drive_folder_id', ''],
    ['time_window_minutes', 10],
    ['wallet_monto_factor', 1],
    ['wallet_id_factor', 1],
    ['wallet_scan_tail_rows', 20000],
    ['transfers_scan_tail_rows', 3000],
  ];

  const toAdd = defaults.filter(([k]) => !existing.has(k));
  if (toAdd.length) shCfg.getRange(shCfg.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
}

// -------------------- Helpers base --------------------
function sh_(name) {
  const sh = getSs_().getSheetByName(name);
  if (!sh) throw new Error(`Hoja no encontrada: ${name}`);
  return sh;
}

function log_(action, details) {
  try {
    sh_(SHEET.LOGS).appendRow([new Date(), action, details, Session.getActiveUser().getEmail() || '']);
  } catch (e) {
    Logger.log('log error: ' + e.message);
  }
}

function getConfig_() {
  const v = sh_(SHEET.CONFIG).getDataRange().getValues();
  const m = {};
  for (let i = 1; i < v.length; i++) {
    const k = String(v[i][0] || '').trim();
    if (!k) continue;
    m[k] = v[i][1];
  }
  const asInt = (x, d) => {
    const n = parseInt(String(x ?? '').trim(), 10);
    return Number.isFinite(n) ? n : d;
  };
  const asNum = (x, d) => {
    const n = Number(String(x ?? '').trim());
    return Number.isFinite(n) ? n : d;
  };
  return {
    drive_folder_id: String(m.drive_folder_id || '').trim(),
    time_window_minutes: asInt(m.time_window_minutes, 15),
    wallet_monto_factor: asNum(m.wallet_monto_factor, 1),
    wallet_id_factor: asNum(m.wallet_id_factor, 1),
    wallet_scan_tail_rows: asInt(m.wallet_scan_tail_rows, 20000),
    transfers_scan_tail_rows: asInt(m.transfers_scan_tail_rows, 3000),
  };
}

function genId_(prefix) {
  const ts = Date.now();
  const r = Math.floor(Math.random() * 100000);
  return `${prefix}_${ts}_${r}`;
}

function cleanCuil_(s) {
  if (!s) return '';
  return String(s).replace(/[^\d]/g, '');
}

function cleanWalletId_(v, factor) {
  if (v == null) return '';
  let s = String(v).trim();
  s = s.replace(/\s/g, '');
  s = s.replace(/\.0+$/,'');
  const n = Number(s);
  if (Number.isFinite(n) && factor && factor !== 1) {
    return String(Math.round(n * factor));
  }
  return s;
}

function parseMonto_(v, factor) {
  if (v == null || v === '') return null;
  let s = String(v).trim().replace(/\s/g, '');
  s = s.replace(',', '.');
  let n = Number(s);
  if (!Number.isFinite(n)) return null;
  if (factor && factor !== 1) n = n * factor;
  return n;
}

function normalizeName_(str) {
  if (!str) return '';
  let s = String(str).toLowerCase();
  const map = { 'á':'a','é':'e','í':'i','ó':'o','ú':'u','ñ':'n','ü':'u','à':'a','è':'e','ì':'i','ò':'o','ù':'u' };
  s = s.split('').map(ch => map[ch] || ch).join('');
  s = s.replace(/[^a-z\s]/g, ' ').replace(/\s+/g, ' ').trim();
  return s;
}

function formatDateTimeAR_(d) {
  const dd = (d instanceof Date) ? d : new Date(d);
  return Utilities.formatDate(dd, TZ, 'dd/MM/yyyy HH:mm');
}

function isValidDate_(d) {
  return Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d.getTime());
}

function parseARDateTime_(v) {
  if (isValidDate_(v)) return v;
  if (v == null || v === '') return null;

  if (typeof v === 'number' && Number.isFinite(v)) {
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return isValidDate_(d) ? d : null;
  }

  const txt = String(v).trim();
  if (!txt) return null;

  const m = txt.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const dd = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10) - 1;
    const yyyy = parseInt(m[3], 10);
    const HH = m[4] ? parseInt(m[4], 10) : 0;
    const MM = m[5] ? parseInt(m[5], 10) : 0;
    const SS = m[6] ? parseInt(m[6], 10) : 0;
    const d = new Date(yyyy, mm, dd, HH, MM, SS, 0);
    return isValidDate_(d) ? d : null;
  }

  const iso = txt.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) {
    const yyyy = parseInt(iso[1], 10);
    const mm = parseInt(iso[2], 10) - 1;
    const dd = parseInt(iso[3], 10);
    const d = new Date(yyyy, mm, dd, 0, 0, 0, 0);
    return isValidDate_(d) ? d : null;
  }

  return null;
}

function startOfDay_(d) {
  if (!d) return null;
  const x = new Date(d);
  return new Date(x.getFullYear(), x.getMonth(), x.getDate(), 0, 0, 0, 0);
}
function endOfDay_(d) {
  if (!d) return null;
  const x = new Date(d);
  return new Date(x.getFullYear(), x.getMonth(), x.getDate(), 23, 59, 59, 999);
}
function inRange_(d, minD, maxD) {
  if (!d) return false;
  if (minD && d < minD) return false;
  if (maxD && d > maxD) return false;
  return true;
}

// -------- fast find row by ID (TextFinder)
function findRowById_(sheetName, colIndex1Based, idValue) {
  const sh = sh_(sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const rg = sh.getRange(2, colIndex1Based, lastRow - 1, 1);
  const hit = rg.createTextFinder(String(idValue)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : null;
}

// -------------------- Usuarios --------------------
function createUser(payload) {
  const nombre = String(payload?.nombre_canonico || '').trim();
  const cuilRaw = String(payload?.cuit_cuil || '').trim();
  const beastId = String(payload?.user_id_beast || '').trim(); // Capturamos Beast
  
  if (!nombre) return { error: 'El nombre es obligatorio.' };

  const sh = sh_(SHEET.USUARIOS);
  const data = sh.getDataRange().getValues();
  
  // Validación de duplicados por nombre
  const newNameNorm = normalizeName_(nombre);
  for (let i = 1; i < data.length; i++) {
    if (normalizeName_(String(data[i][1])) === newNameNorm) {
      return { error: `Ya existe un usuario con el nombre "${data[i][1]}".` };
    }
  }

  // CREACIÓN: Generamos el ID aquí para que no falle
  const user_id = genId_('USR'); 
  
  sh.appendRow([
    user_id,        // Col A
    nombre,         // Col B
    cuilRaw,        // Col C
    '',             // Col D (alias)
    'Atenea',       // Col E (org)
    new Date(),     // Col F
    'Activo',       // Col G
    beastId         // Col H (ID BEAST)
  ]);

  log_('CREATE_USER', `${user_id} | ${nombre} | Beast: ${beastId}`);
  return { success: true, user_id: user_id };
}

function searchUsers(query) {
  const sh = sh_(SHEET.USUARIOS);
  const data = sh.getDataRange().getValues();
  const out = [];

  const qRaw = String(query || '').trim(); 
  const qNameClean = normalizeName_(qRaw).trim();
  const qDigits = qRaw.replace(/\D/g, ''); 

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const user_id = String(row[0] || '');
    if (!user_id) continue;

    const nombreOriginal = String(row[1] || '').trim();
    const cuilOriginal = row[2] != null ? String(row[2]).split('.')[0].trim() : '';
    const beastIdOriginal = String(row[7] || '').trim(); // Leemos la Columna H (ID Beast)

    const nombreBaseNorm = normalizeName_(nombreOriginal);
    const cuilBaseDigits = cuilOriginal.replace(/\D/g, ''); 

    let match = false;

    if (qRaw === "") {
      match = true; 
    } else {
      // A. Búsqueda por NOMBRE
      if (qNameClean.length > 0 && nombreBaseNorm.indexOf(qNameClean) !== -1) {
        match = true;
      }

      // B. Búsqueda por CUIL (coincidencia exacta de números)
      if (!match && qDigits.length > 0 && cuilBaseDigits === qDigits) {
        match = true;
      }

      // C. NUEVO: Búsqueda por ID BEAST (que busque si el ID contiene lo que escribiste)
      if (!match && beastIdOriginal.toLowerCase().indexOf(qRaw.toLowerCase()) !== -1) {
        match = true;
      }
    }

    if (match) {
      out.push({ 
        user_id, 
        nombre_canonico: nombreOriginal, 
        cuit_cuil: cuilOriginal, 
        user_id_beast: beastIdOriginal, // Incluimos el dato en el resultado
        estado: row[6] || 'Activo'
      });
    }
  }

  return out.sort((a, b) => a.nombre_canonico.localeCompare(b.nombre_canonico));
}

function getUser(user_id) {
  const row = findRowById_(SHEET.USUARIOS, 1, user_id);
  if (!row) return null;
  const r = sh_(SHEET.USUARIOS).getRange(row, 1, 1, 8).getValues()[0]; // Leemos hasta col 8
  return {
    user_id: r[0],
    nombre_canonico: r[1],
    cuit_cuil: r[2],
    estado: r[6] || 'Activo',
    user_id_beast: r[7] || '-' // Devolvemos el dato
  };
}

function updateUser(user_id, patch) {
  const sh = sh_(SHEET.USUARIOS);
  const row = findRowById_(SHEET.USUARIOS, 1, user_id);
  if (!row) return { error: 'Usuario no encontrado.' };

  const cur = sh.getRange(row, 1, 1, 8).getValues()[0]; // Leemos hasta la columna 8
  const nombre = String(patch?.nombre_canonico ?? cur[1] ?? '').trim();
  const cuil = String(patch?.cuit_cuil ?? cur[2] ?? '').trim();
  const beast = String(patch?.user_id_beast ?? cur[7] ?? '').trim(); // Capturar del patch

  sh.getRange(row, 2).setValue(nombre);
  sh.getRange(row, 3).setValue(cuil);
  sh.getRange(row, 8).setValue(beast); // Escribir en Columna H (8)

  log_('UPDATE_USER', `${user_id} | ${nombre} | CUIL: ${cuil} | Beast: ${beast}`);
  return { success: true };
}

// -------------------- Transferencias --------------------
function createTransfer(payload) {
  const user_id = String(payload?.user_id || '').trim();
  const monto = Number(payload?.monto);
  const turno = String(payload?.turno || '').trim();
  const fecha = String(payload?.fecha || '').trim();
  const hora = String(payload?.hora || '').trim();
  const notaExtra = String(payload?.nota || '').trim();
  
  // NUEVO: Capturamos la billetera. Si no viene, ponemos cadena vacía (el lector luego asumirá OTRA)
  const billetera = String(payload?.billetera || '').trim().toUpperCase();

  if (!user_id) return { error: 'Falta user_id.' };
  if (!Number.isFinite(monto) || monto === 0) return { error: 'Monto inválido (no puede ser 0).' };
  if (!turno) return { error: 'Turno obligatorio.' };
  if (!fecha || !hora) return { error: 'Fecha y hora obligatorias.' };

  const [yyyy, mm, dd] = fecha.split('-').map(x => parseInt(x, 10));
  const [HH, MM] = hora.split(':').map(x => parseInt(x, 10));
  if (![yyyy,mm,dd,HH,MM].every(n => Number.isFinite(n))) return { error: 'Fecha/hora inválidas.' };

  const d = new Date(yyyy, (mm - 1), dd, HH, MM, 0, 0);
  const fecha_hora_operativa_txt = formatDateTimeAR_(d);

  const transfer_id = genId_('TRF');
  const nota = `Turno=${turno}${notaExtra ? ' | ' + notaExtra : ''}`;

  sh_(SHEET.TRANSF).appendRow([
    transfer_id,              // Col A
    user_id,                  // Col B
    fecha_hora_operativa_txt, // Col C
    monto,                    // Col D
    '-',                      // Col E (reservado)
    '',                       // Col F (reservado)
    '',                       // Col G (reservado)
    '',                       // Col H (reservado)
    'Pendiente',              // Col I
    nota,                     // Col J
    billetera                 // Col K <-- AQUÍ SE GUARDA LA BILLETERA
  ]);

  log_('CREATE_TRANSFER', `${transfer_id} | user=${user_id} | $${monto} | ${fecha_hora_operativa_txt} | ${turno} | ${billetera}`);
  return { success: true, transfer_id };
}

function listTransfersByUser(user_id, limit = 50) {
  const sh = sh_(SHEET.TRANSF);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const numRows = lastRow - 1;
  const data = sh.getRange(2, 1, numRows, 10).getValues();

  const out = [];
  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];
    if (String(r[1]) !== String(user_id)) continue;
    const dt = parseARDateTime_(r[2]);
    out.push({
      transfer_id: r[0],
      user_id: r[1],
      fecha_hora_operativa: dt ? formatDateTimeAR_(dt) : (r[2] || ''),
      monto: r[3],
      comprobante_url: r[5],
      estado: r[8],
      nota: r[9] || ''
    });
    if (out.length >= limit) break;
  }
  return out;
}

function getTransferRow_(transfer_id) {
  const sh = sh_(SHEET.TRANSF);
  const row = findRowById_(SHEET.TRANSF, 1, transfer_id);
  if (!row) return null;
  
  // CORRECCIÓN: Cambiamos 10 por 11 para incluir la Columna K (Billetera)
  const r = sh.getRange(row, 1, 1, 11).getValues()[0];
  
  return { rowIndex: row, row: r };
}

function setTransferEstado_(transfer_id, estado) {
  const hit = getTransferRow_(transfer_id);
  if (!hit) return { error: 'Transferencia no encontrada.' };
  sh_(SHEET.TRANSF).getRange(hit.rowIndex, 9).setValue(String(estado || '').trim());
  log_('SET_TRANSFER_ESTADO', `${transfer_id} -> ${estado}`);
  return { success: true };
}

function updateTransfer(transfer_id, patch) {
  const hit = getTransferRow_(transfer_id);
  if (!hit) return { error: 'Transferencia no encontrada.' };

  const sh = sh_(SHEET.TRANSF);
  const row = hit.rowIndex;

  // 1. Actualizar Monto
  if (patch?.monto != null && patch.monto !== '') {
    const monto = Number(patch.monto);
    if (!Number.isFinite(monto) || monto === 0) return { error: 'Monto inválido.' };
    sh.getRange(row, 4).setValue(monto);
  }
  
  // 2. Actualizar Fecha
  if (patch?.fecha_hora_operativa) {
    const d = parseARDateTime_(patch.fecha_hora_operativa);
    if (!d) return { error: 'Fecha inválida.' };
    sh.getRange(row, 3).setValue(formatDateTimeAR_(d));
  }

  // 3. Actualizar Billetera (NUEVO)
  if (patch?.billetera) {
    sh.getRange(row, 11).setValue(patch.billetera); // Col K
  }

  // 4. Construir y Actualizar Nota Completa (Turno + Nota)
  if (patch?.turno || patch?.nota !== undefined) {
    // Si viene en el patch lo usamos, sino leemos lo que ya estaba? 
    // Simplificación: El modal siempre manda todo junto.
    const nuevoTurno = patch.turno || 'TARDE';
    const nuevaNota = patch.nota || '';
    
    // Armamos el string: "Turno=TARDE | Cliente pide factura"
    const finalString = `Turno=${nuevoTurno}${nuevaNota ? ' | ' + nuevaNota : ''}`;
    sh.getRange(row, 10).setValue(finalString); // Col J
  }

  log_('UPDATE_TRANSFER', `${transfer_id} | patch=${JSON.stringify(patch)}`);
  return { success: true };
}

// -------------------- Upload comprobante --------------------
function uploadComprobante(base64Data, mimeType, transferId) {
  try {
    const cfg = getConfig_();
    if (!cfg.drive_folder_id) return { error: 'Falta drive_folder_id en Config.' };

    const hit = getTransferRow_(transferId);
    if (!hit) return { error: 'Transferencia no encontrada: ' + transferId };

    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType);

    const extMap = { 'image/png':'png','image/jpeg':'jpg','image/gif':'gif','image/webp':'webp','application/pdf':'pdf' };
    const ext = extMap[mimeType] || 'bin';

    const ts = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss');
    const fileName = `${transferId}_${ts}.${ext}`;
    blob.setName(fileName);

    const folder = DriveApp.getFolderById(cfg.drive_folder_id);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const url = file.getUrl();
    const id = file.getId();

    const sh = sh_(SHEET.TRANSF);
    const r = hit.rowIndex;
    sh.getRange(r, 6).setValue(url);
    sh.getRange(r, 7).setValue(id);
    sh.getRange(r, 8).setValue(new Date());

    log_('UPLOAD_COMPROBANTE', `${transferId} | file=${id}`);
    return { success: true, url, fileId: id, fileName };
  } catch (e) {
    return { error: e.message };
  }
}

// -------------------- Carga CSV --------------------
function importWalletCsv(base64Data, mimeType, fileName) {
  try {
    ensureAllSheets_();

    // 1. DETERMINAR EL ID DE BILLETERA SEGÚN EL NOMBRE DEL ARCHIVO
    let billeteraIdAsignada = "03"; // Por defecto "OTRA"
    const nameUpper = (fileName || "").toUpperCase();
    
    if (nameUpper.includes("RAYA") || nameUpper.includes("BARROSO")) {
      billeteraIdAsignada = "01";
    } else if (nameUpper.includes("DC")) {
      billeteraIdAsignada = "02";
    } else if (nameUpper.includes("AMFRIS")) {
      billeteraIdAsignada = "04";
    } else if (nameUpper.includes("MANZO")) { // <--- AGREGADO MANZO
      billeteraIdAsignada = "05";
    }

    if (mimeType !== 'text/csv' && mimeType !== 'application/vnd.ms-excel' && mimeType !== 'application/octet-stream') {
      return { error: 'Subí un CSV válido.' };
    }

    const cfg = getConfig_();
    const decoded = Utilities.base64Decode(base64Data);
    const csvText = Utilities.newBlob(decoded, mimeType).getDataAsString('UTF-8');
    const rows = Utilities.parseCsv(csvText);
    if (!rows || rows.length < 2) return { error: 'CSV vacío o inválido.' };

    const norm = (x) => String(x || '').replace(/^\uFEFF/, '').trim().toLowerCase();
    let headerRowIndex = 0;
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const r = rows[i].map(norm);
      if (r.includes('fecha') && r.includes('monto')) { headerRowIndex = i; break; }
    }

    const header = rows[headerRowIndex].map(h => String(h || '').replace(/^\uFEFF/, '').trim());
    const idxAny = (names) => {
      const lower = header.map(h => h.toLowerCase());
      for (const n of names) {
        const j = lower.findIndex(h => h === String(n).toLowerCase());
        if (j >= 0) return j;
      }
      return -1;
    };

    const iFecha = idxAny(['Fecha']);
    const iMonto = idxAny(['Monto']);
    const iNombre = idxAny(['Destinario/Origen','Destinatario/Origen']);
    const iCuil = idxAny(['CUIL/CUIT','CUIT/CUIL']);
    const iId = idxAny(['ID']);

    if ([iFecha,iMonto,iNombre,iCuil,iId].some(x => x < 0)) {
      return { error: 'CSV no tiene las columnas requeridas: Fecha, ID, Monto, Destinario/Origen, CUIL/CUIT.' };
    }

    const sh = sh_(SHEET.MOVS);
    const lastRow = sh.getLastRow();
    const existing = new Set();
    if (lastRow >= 2) {
      const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      vals.forEach(v => { const id = String(v[0] || '').trim(); if (id) existing.add(id); });
    }

    const out = [];
    const now = new Date();

    for (let i = headerRowIndex + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.length === 0) continue;

      const fechaTxt = String(r[iFecha] || '').trim();
      const montoTxt = r[iMonto];
      const nombre = String(r[iNombre] || '').trim();
      const cuilRaw = String(r[iCuil] || '').trim();
      const walletIdRaw = r[iId];

      if (!fechaTxt || !nombre || montoTxt === '' || montoTxt == null) continue;

      const fecha = parseARDateTime_(fechaTxt);
      if (!fecha) continue;

      const walletId = cleanWalletId_(walletIdRaw, cfg.wallet_id_factor);
      const monto = parseMonto_(montoTxt, cfg.wallet_monto_factor);
      if (!walletId || monto == null) continue;

      const movement_id = `MOV_${walletId}`;
      if (existing.has(movement_id)) continue;

      const origen_normalizado = normalizeName_(nombre);
      const cuil_norm = cleanCuil_(cuilRaw);

      const raw = JSON.stringify({
        source_file: fileName || 'wallet.csv',
        wallet_id: walletId,
        billetera_id: billeteraIdAsignada
      });

      // Columna A: movement_id, Columna B: billeteraIdAsignada
      out.push([movement_id, billeteraIdAsignada, formatDateTimeAR_(fecha), walletId, monto, nombre, cuilRaw, origen_normalizado, now, raw]);
      existing.add(movement_id);
    }

    if (out.length) {
      sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
    }
    return { success: true, imported: out.length };
  } catch (e) {
    return { error: e.message };
  }
}

// -------------------- Indexes --------------------
function getUserIndex_() {
  const data = sh_(SHEET.USUARIOS).getDataRange().getValues();
  const idx = {};
  for (let i = 1; i < data.length; i++) {
    const user_id = String(data[i][0] || '').trim();
    if (!user_id) continue;

    const nombre = String(data[i][1] || '');
    const cuilRaw = String(data[i][2] || '');

    idx[user_id] = {
      user_id,

      
      nombre,                 
      nombre_canonico: nombre, 

      nombre_norm: normalizeName_(nombre),

     
      cuil_raw: cuilRaw,
      cuit_cuil: cuilRaw,

      cuil_norm: cleanCuil_(cuilRaw)
    };
  }
  return idx;
}

// -------------------- Lecturas optimizadas (tail + rango) --------------------
// --- VERSIÓN CORREGIDA DE getTransfersTail_ ---
function getTransfersTail_(tailRows) {
  const shT = sh_(SHEET.TRANSF);
  
  // 1. Si la hoja no existe o está vacía (menos de 2 filas), devolvemos vacío
  if (!shT || shT.getLastRow() < 2) return { startRow: 2, data: [] };

  const lastRow = shT.getLastRow();
  const cfg = getConfig_();
  const t = Math.max(50, Math.min(50000, Number(tailRows || cfg.transfers_scan_tail_rows || 3000)));

  // 2. Cálculo seguro de filas
  const startRow = Math.max(2, lastRow - t + 1);
  const numRows = lastRow - startRow + 1;

  // 3. Verificamos que vamos a pedir algo lógico
  if (numRows < 1) return { startRow: 2, data: [] };

  // 4. Verificamos las columnas reales para evitar el error "columns out of bounds"
  const lastCol = shT.getLastColumn();
  if (lastCol < 1) return { startRow: 2, data: [] };

  // 5. Pedimos los datos usando el número real de columnas disponibles
  // (Esto evita que explote si la hoja tiene menos de 10 columnas por algún motivo)
  const data = shT.getRange(startRow, 1, numRows, lastCol).getValues();
  
  return { startRow, data };
}

function getMovsTail_(tailRows) {
  const shM = sh_(SHEET.MOVS);
  
  // Seguridad: si no existe la hoja o está vacía
  if (!shM || shM.getLastRow() < 2) return { startRow: 2, data: [] };

  const lastRow = shM.getLastRow();
  const cfg = getConfig_();
  const t = Math.max(200, Math.min(200000, Number(tailRows || cfg.wallet_scan_tail_rows || 20000)));

  const startRow = Math.max(2, lastRow - t + 1);
  const numRows = lastRow - startRow + 1;
  
  // Seguridad anti-crash
  if (numRows < 1) return { startRow: 2, data: [] };
  
  // Verificamos columnas reales para no pedir de más
  const lastCol = shM.getLastColumn();
  if (lastCol < 1) return { startRow: 2, data: [] };

  const data = shM.getRange(startRow, 1, numRows, lastCol).getValues();
  return { startRow, data };
}

function readTransfersInRange_(fromDate, toDate, limit, tailRows) {
  const { startRow, data } = getTransfersTail_(tailRows);
  const out = [];
  const lim = Math.max(1, Math.min(5000, Number(limit || 300)));

  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];
    const id = String(r[0] || '').trim();
    if (!id) continue;

    const dt = parseARDateTime_(r[2]);
    if (!dt) continue;

    if (!inRange_(dt, fromDate, toDate)) continue;

    out.push({ rowIndex: startRow + i, row: r, tFecha: dt });
    if (out.length >= lim) break;
  }
  return out;
}

function readMovsInRange_(fromDate, toDate, tailRows) {
  const { data } = getMovsTail_(tailRows);
  const movs = [];

  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];
    const movement_id = String(r[0] || '').trim();
    if (!movement_id) continue;

    const dt = parseARDateTime_(r[2]);
    if (!dt) continue;
    if (!inRange_(dt, fromDate, toDate)) continue;

    const monto = Number(r[4]);
    if (!Number.isFinite(monto)) continue;

    const nombre = String(r[5] || '');
    const cuilRaw = String(r[6] || '');
    const cuil_norm = cleanCuil_(cuilRaw);
    const nombre_norm = String(r[7] || normalizeName_(nombre));

    movs.push({ movement_id, fecha: dt, monto, nombre, nombre_norm, cuil_norm });
  }
  return movs;
}

// -------------------- Matching estricto --------------------
function strictFindMatch_(transferData, userIdx, movs, cfg) {
  // 1. OBTENCIÓN DE DATOS
  const r = transferData.row;
  const userId = String(r[1] || '').trim();
  const u = userIdx[userId] || {};
  
  // Datos del Usuario (Normalizados)
  const tNombre = normalizeName_(u.nombre_canonico || ''); 
  const tCuil = cleanCuil_(u.cuit_cuil || '');       
  const tMonto = transferData.tMontoManual || parseFloat(r[3]); 
  
  // Verificación básica
  if (!tMonto) return { matched: false, reason: 'Monto inválido.' };

  // --- FILTRO 1: MONTO (Obligatorio) ---
  // Filtramos de la lista de movimientos (que ya viene filtrada por fecha desde matchTransferAndUpdate)
  const candidatosPorMonto = movs.filter(m => Math.abs(parseFloat(m[4]) - tMonto) < 1.0);
  
  if (candidatosPorMonto.length === 0) {
    return { matched: false, reason: `No existe ingreso de $${tMonto} en ese rango horario.` };
  }

  // --- FILTRO 2: CUIL (PRIORIDAD ABSOLUTA - EL CAMBIO CLAVE) ---
  // Si el perfil del cliente tiene CUIL, buscamos coincidencia EXACTA.
  if (tCuil && tCuil.length > 5) {
    const matchCuil = candidatosPorMonto.find(m => {
      // Columna G (índice 6) es el CUIL/CUIT en la hoja Movimientos
      const mCuil = cleanCuil_(m[6]); 
      return mCuil === tCuil;
    });

    if (matchCuil) {
      // ¡MATCH POR CUIL! Ignoramos si el nombre está al revés.
      return { matched: true, movement_id: matchCuil[0], delta_min: 0 };
    }
  }

  // --- FILTRO 3: NOMBRE (PLAN B) ---
  // Solo si no hubo match por CUIL (o el cliente no tiene CUIL cargado), probamos por nombre.
  if (!tNombre) {
    return { matched: false, reason: 'El perfil no tiene CUIL coincidente ni Nombre válido.' };
  }

  const matchNombre = candidatosPorMonto.find(m => {
    const mNombre = normalizeName_(m[5]);      // Nombre original banco
    const mNombreNorm = normalizeName_(m[7]);  // Nombre normalizado banco
    
    // Búsqueda flexible: ¿El nombre del usuario está en el del banco o viceversa?
    return mNombre.includes(tNombre) || tNombre.includes(mNombre) || 
           mNombreNorm.includes(tNombre) || tNombre.includes(mNombreNorm);
  });

  if (matchNombre) {
    return { matched: true, movement_id: matchNombre[0], delta_min: 0 };
  }

  // Si llegamos acá: Hay plata, pero ni el CUIL ni el Nombre coinciden.
  return { matched: false, reason: 'Monto coincide, pero los datos del titular no.' };
}

function buildStrictNoMatchReason_(t, pool, cfg) {
  const { tCuil, tName, tFecha, monto } = t;

  if (!tFecha) return 'Sin fecha/hora operativa válida';
  if (!tName) return 'Usuario sin nombre (Usuarios.nombre_canonico vacío)';
  if (!Number.isFinite(monto) || monto === 0) return 'Monto inválido o 0';

  if (!pool || pool.length === 0) {
    if (tCuil) return 'No hay movimientos en ese rango con ese CUIL';
    return 'No hay movimientos en ese rango con ese nombre';
  }

  const namePool = pool.filter(m => String(m.nombre_norm || '').trim() === tName);
  if (namePool.length === 0) {
    if (tCuil) return 'Hay movimientos con ese CUIL, pero el nombre NO coincide (estricto)';
    return 'Hay movimientos en rango, pero el nombre NO coincide exacto (estricto)';
  }

  const montoPool = namePool.filter(m => Number(m.monto) === monto);
  if (montoPool.length === 0) return 'Nombre coincide, pero NO hay monto exacto';

  const timePool = montoPool.filter(m => {
    if (!m.fecha) return false;
    const deltaMin = Math.abs(tFecha - m.fecha) / 60000;
    return deltaMin <= cfg.time_window_minutes;
  });
  if (timePool.length === 0) return `Monto y nombre ok, pero fuera de ventana (${cfg.time_window_minutes} min)`;

  return 'No matcheó por regla estricta (revisar normalización o fechas)';
}

function getNoMatchListLive(opts) {
  ensureAllSheets_();
  opts = opts || {};
  const cfg = getConfig_();
  const userIdx = getUserIndex_();

  const mode = String(opts.mode || 'sync').toLowerCase(); 
  
  let fromDate, toDate;

  // 1. DEFINIR RANGO DE FECHAS
  if (opts.from) {
    // Si viene fecha manual del filtro, usamos esa
    fromDate = new Date(Number(opts.from.slice(0, 4)), Number(opts.from.slice(5, 7)) - 1, Number(opts.from.slice(8, 10)), 0, 0, 0, 0);
  } else {
    // CAMBIO AQUÍ: 15 días atrás por defecto (antes 7) para cubrir casos pendientes viejos
    fromDate = new Date(Date.now() - 15 * 24 * 60 * 60 * 1000);
  }

  if (opts.to) {
    toDate = new Date(Number(opts.to.slice(0, 4)), Number(opts.to.slice(5, 7)) - 1, Number(opts.to.slice(8, 10)), 23, 59, 59, 999);
  } else {
    toDate = new Date();
  }

  const cutoff = getLastWalletCutoff_(); 

  // 2. LEER MOVIMIENTOS (Para calcular el motivo del error)
  const spanMin = Math.max(60, cfg.time_window_minutes * 6);
  const movFrom = new Date(fromDate.getTime() - spanMin * 60000);
  const movTo = new Date(toDate.getTime() + spanMin * 60000);
  const movs = readMovsInRange_(movFrom, movTo, cfg.wallet_scan_tail_rows);

  const byCuil = new Map();
  const byName = new Map();
  for (const m of movs) {
    const cn = String(m.cuil_norm || '').trim();
    const nn = String(m.nombre_norm || '').trim();
    if (cn) {
      if (!byCuil.has(cn)) byCuil.set(cn, []);
      byCuil.get(cn).push(m);
    }
    if (nn) {
      if (!byName.has(nn)) byName.set(nn, []);
      byName.get(nn).push(m);
    }
  }

  // 3. LEER TRANSFERENCIAS (Aumentamos límite a 2000 para cubrir los 15 días)
  const transfers = readTransfersInRange_(fromDate, toDate, 2000, cfg.transfers_scan_tail_rows);
  const out = [];

  for (const tr of transfers) {
    const r = tr.row;
    const transfer_id = String(r[0] || '').trim();
    if (!transfer_id) continue;

    const estado = String(r[8] || '').trim() || 'Pendiente';

    // Ignoramos estafadores
    if (estado === 'Estafador') continue;

    // FILTRO PRINCIPAL: Si estamos en modo sync, SOLO mostramos 'Sin match'
    if (mode === 'sync' && estado !== 'Sin match') continue;

    const user_id = String(r[1] || '').trim();
    const u = userIdx[user_id] || {};
    // Usamos u.nombre (que es como lo guarda el índice ahora)
    const nombre = u.nombre || 'Desconocido';
    const cuil = u.cuil_raw || '';
    const dt = tr.tFecha || parseARDateTime_(r[2]);
    const tName = String(u.nombre_norm || '').trim();
    const tCuil = String(u.cuil_norm || '').trim();
    const monto = Number(r[3]);

    let motivo = '';

    if (estado === 'Sin match') {
      const pool = tCuil ? (byCuil.get(tCuil) || []) : (byName.get(tName) || []);
      motivo = buildStrictNoMatchReason_({ tCuil, tName, tFecha: dt, monto }, pool, cfg);
    } else if (mode === 'range') {
      if (estado === 'Match') motivo = 'Matcheada';
      else if (estado === 'Pendiente') {
        if (dt && cutoff && dt > cutoff) motivo = 'Pendiente: fuera de corte billetera';
        else motivo = 'Pendiente: dentro de corte (sincronizá)';
      }
    }

    out.push({
      transfer_id,
      motivo,
      user_id,
      nombre,
      cuil,
      fecha_hora_operativa: dt ? formatDateTimeAR_(dt) : (r[2] || ''),
      monto: r[3] || '',
      comprobante_url: r[5] || '',
      estado_transferencia: estado
    });
  }

  // 4. ORDENAR POR FECHA (Lo más nuevo arriba)
  out.sort((a, b) => {
    const da = parseARDateTime_(a.fecha_hora_operativa);
    const db = parseARDateTime_(b.fecha_hora_operativa);
    return (db || 0) - (da || 0);
  });

  return out;
}

// -------------------- Edit modal data + Matchear + Estafador --------------------
function getTransferEditData(transfer_id) {
  ensureAllSheets_();
  const hit = getTransferRow_(transfer_id);
  if (!hit) return { error: 'Transferencia no encontrada.' };

  const r = hit.row;
  const user_id = String(r[1] || '').trim();
  const u = getUser(user_id) || {};
  const dt = parseARDateTime_(r[2]);
  
  // --- PARSEO INTELIGENTE DE NOTA Y TURNO ---
  const rawNota = String(r[9] || ''); // Columna J
  let turno = 'TARDE'; // Default si no encuentra nada
  let notaTexto = '';

  const matchT = rawNota.match(/Turno=([A-Z]+)/);
  if (matchT) {
    turno = matchT[1];
    // Todo lo que esté después del pipe (|) es la nota
    const parts = rawNota.split('|');
    if (parts.length > 1) {
      notaTexto = parts.slice(1).join('|').trim();
    }
  } else if (rawNota) {
    notaTexto = rawNota; // Si no hay formato Turno=, todo es nota
  }

  // --- OBTENER BILLETERA ---
  const billetera = String(r[10] || 'OTRA'); // Columna K

  return {
    success: true,
    transfer: {
      transfer_id: r[0],
      user_id,
      fecha_hora_operativa: dt ? formatDateTimeAR_(dt) : (r[2] || ''),
      monto: r[3],
      estado: r[8] || 'Pendiente',
      comprobante_url: r[5] || '',
      turno: turno,       // Dato separado
      nota: notaTexto,    // Dato separado
      billetera: billetera // Dato separado
    },
    user: {
      nombre_canonico: u.nombre_canonico || 'Desconocido',
      cuit_cuil: u.cuit_cuil || '-'
    }
  };
}

function getBilleteraId(nombreBilletera) {
  const mapping = {
    'BARROSO': '01',
    'DC': '02',
    'OTRA': '03',
    'AMFRIS': '04',
    'MANZO': '05' // <--- AGREGADO MANZO
  };
  return mapping[nombreBilletera] || '03';
}

function matchTransferAndUpdate(payload) {
  ensureAllSheets_();
  const cfg = getConfig_();

  const transfer_id = String(payload?.transfer_id || '').trim();
  if (!transfer_id) return { error: 'Falta transfer_id.' };

  const hit = getTransferRow_(transfer_id);
  if (!hit) return { error: 'Transferencia no encontrada.' };

  // 1. Parches
  const patchT = payload?.transfer_patch || {};
  if (patchT && (patchT.fecha_hora_operativa || patchT.monto != null)) {
    updateTransfer(transfer_id, patchT);
  }
  
  const hitActualizado = getTransferRow_(transfer_id);
  const r = hitActualizado.row;

  let rawMonto = String(r[3]); 
  let cleanMonto = rawMonto.replace(/[$.]/g, '').replace(',', '.').trim();
  const tMonto = parseFloat(cleanMonto);
  const tFecha = parseARDateTime_(r[2]);
  if (!tFecha) return { success:true, matched:false, estado:'Sin match', reason:'Fecha inválida.' };

  // Filtro de Billetera
  const billeteraNombre = String(r[10] || '').toUpperCase(); 
  let idTargetNum = 3; 

  if (billeteraNombre.includes("RAYA") || billeteraNombre.includes("BARROSO")) idTargetNum = 1;
  else if (billeteraNombre.includes("DC")) idTargetNum = 2;
  else if (billeteraNombre.includes("AMFRIS")) idTargetNum = 4;
  else if (billeteraNombre.includes("MANZO")) idTargetNum = 5; // <--- AGREGADO MANZO

  // 2. OBTENER MOVIMIENTOS
  const { data } = getMovsTail_(cfg.wallet_scan_tail_rows);
  const spanMs = 12 * 60 * 60 * 1000; 
  const minTime = tFecha.getTime() - spanMs;
  const maxTime = tFecha.getTime() + spanMs;

  const movs = data.filter(m => {
      const idEnCelda = m[1]; 
      if (idEnCelda == null || idEnCelda === '') return false;
      if (parseInt(idEnCelda, 10) != idTargetNum) return false;
      const mFecha = parseARDateTime_(m[2]);
      if (!mFecha) return false;
      const mTime = mFecha.getTime();
      return (mTime >= minTime && mTime <= maxTime);
  });

  // 3. MATCHEO
  const userIdx = getUserIndex_();
  
  const res = strictFindMatch_({ 
      rowIndex: hitActualizado.rowIndex, 
      row: r, 
      tFecha: tFecha,
      tMontoManual: tMonto 
  }, userIdx, movs, cfg);

  if (res.matched) {
    // --- VERIFICACIÓN DE DUPLICADOS ---
    const shT = sh_(SHEET.TRANSF);
    const lastRow = shT.getLastRow();
    const usedIdsData = shT.getRange(2, 12, lastRow-1, 1).getValues(); 
    
    const isUsed = usedIdsData.flat().some(id => String(id).trim() === String(res.movement_id).trim());

    if (isUsed) {
      setTransferEstado_(transfer_id, 'Duplicada');
      shT.getRange(hit.rowIndex, 9).setBackground('#fef08a'); 
      return { success: true, matched: false, estado: 'Duplicada', reason: 'Este ingreso ya fue asignado a otra transferencia.' };
    }

    setTransferEstado_(transfer_id, 'Match');
    shT.getRange(hit.rowIndex, 9).setBackground(null);
    shT.getRange(hit.rowIndex, 12).setValue(res.movement_id); 
    
    return { success:true, matched:true, estado:'Match', movement_id: res.movement_id };
  } else {
    const motivoFinal = res.reason || 'No se encontró una coincidencia clara.';
    setTransferEstado_(transfer_id, 'Sin match', motivoFinal); 
    return { success: true, matched: false, estado: 'Sin match', reason: motivoFinal };
  }
}

function markTransferAsEstafador(transfer_id) {
  ensureAllSheets_();
  const hit = getTransferRow_(transfer_id);
  if (!hit) return { error: 'Transferencia no encontrada.' };
  sh_(SHEET.TRANSF).getRange(hit.rowIndex, 9).setValue('Estafador');
  log_('MARK_ESTAFADOR', transfer_id);
  return { success:true };
}

// -------------------- Dashboard --------------------
function getDashboardData() {
  ensureAllSheets_();
  const cfg = getConfig_();
  const userIdx = getUserIndex_();

  const now = new Date();
  const from = new Date(now.getTime() - 24*60*60*1000); // Últimas 24hs para los contadores
  const to = now;

  // 1. PREPARAR EL GRÁFICO (Últimos 7 días)
  const chartMap = {};
  const days = [];
  for (let d=6; d>=0; d--) {
    const dObj = new Date(now.getTime() - d*24*60*60*1000);
    const label = Utilities.formatDate(dObj, TZ, 'dd/MM');
    days.push(label);
    chartMap[label] = 0; 
  }

  // 2. LEER DATOS (Leemos 2000 filas para asegurarnos de agarrar todo)
  const { data } = getTransfersTail_(Math.max(cfg.transfers_scan_tail_rows, 2000)); 

  let total24h = 0;
  let matchCount = 0;
  let pendingCount = 0;
  let noMatchCount = 0;
  let estafadorCount = 0;

  const latest = [];
  const latestLimit = 20; // Cantidad de filas a mostrar en la tabla

  // Recorremos de atrás para adelante (desde lo más nuevo)
  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];
    const transfer_id = String(r[0] || '').trim();
    if (!transfer_id) continue;

    const dt = parseARDateTime_(r[2]); // Fecha hora operativa
    const estado = String(r[8] || '').trim() || 'Pendiente';
    
    // --- LÓGICA DE CONTADORES (Solo suma si es de las últimas 24hs) ---
    if (dt && inRange_(dt, from, to)) {
      total24h++;
      if (estado === 'Match') matchCount++;
      else if (estado === 'Pendiente' || estado === '') pendingCount++;
      else if (estado === 'Sin match') noMatchCount++;
      else if (estado === 'Estafador') estafadorCount++;
    }

    // --- LÓGICA DEL GRÁFICO (Suma al día que corresponda) ---
    if (dt) {
      const lbl = Utilities.formatDate(dt, TZ, 'dd/MM');
      if (chartMap.hasOwnProperty(lbl)) {
        chartMap[lbl] += 1; 
      }
    }

    // --- LÓGICA DE LA TABLA (ACÁ ESTABA EL PROBLEMA) ---
    // Ahora agregamos a la lista SIEMPRE, sin importar la fecha
    if (latest.length < latestLimit) {
      const user_id = String(r[1] || '').trim();
      const u = userIdx[user_id] || {};
      
      latest.push({
        transfer_id,
        nombre: (u.nombre_canonico || u.nombre || 'Desconocido'),
        monto: r[3],
        hora: dt ? Utilities.formatDate(dt, TZ, 'HH:mm') : '',
        fecha_hora_operativa: dt ? formatDateTimeAR_(dt) : (r[2] || ''),
        estado,
        user_id
      });
    }
  }

  const chartValues = days.map(d => chartMap[d]);

  return {
    success:true,
    updated_at: formatDateTimeAR_(now),
    window_from: formatDateTimeAR_(from),
    window_to: formatDateTimeAR_(to),
    totals: { total24h, matchCount, pendingCount, noMatchCount, estafadorCount },
    latest, // Ahora esta lista va llena seguro
    chart: { labels: days, data: chartValues } 
  };
}

function getLastWalletCutoff_() {
  const cfg = getConfig_();
  const { data } = getMovsTail_(cfg.wallet_scan_tail_rows);
  let max = null;

  for (let i = 0; i < data.length; i++) {
    const dt = parseARDateTime_(data[i][2]); // col "Fecha" en Movimientos_Billetera
    if (dt && (!max || dt > max)) max = dt;
  }
  return max; // Date | null
}

function buildMovIndex_(movs) {
  const byNameMonto = new Map(); // Índice para buscar por NOMBRE
  const byCuilMonto = new Map(); // Índice para buscar por CUIL

  for (const m of movs) {
    const nn = String(m.nombre_norm || '').trim();
    const cn = String(m.cuil_norm || '').trim();
    const monto = Number(m.monto);

    if (!Number.isFinite(monto)) continue;

    // A. Guardamos en el índice de NOMBRES
    if (nn) {
      const kName = `${nn}|${monto}`;
      if (!byNameMonto.has(kName)) byNameMonto.set(kName, []);
      byNameMonto.get(kName).push(m);
    }

    // B. Guardamos en el índice de CUILS (Para la prioridad)
    if (cn) {
      const kCuil = `${cn}|${monto}`;
      if (!byCuilMonto.has(kCuil)) byCuilMonto.set(kCuil, []);
      byCuilMonto.get(kCuil).push(m);
    }
  }

  return { byNameMonto, byCuilMonto };
}

function strictFindMatchFast_(tr, userIdx, movIndex, cfg, usedSet) {
  const r = tr.row;
  const transfer_id = String(r[0] || '').trim();
  const user_id = String(r[1] || '').trim();
  const monto = Number(r[3]);

  const u = userIdx[user_id] || {};
  const tFecha = tr.tFecha || parseARDateTime_(r[2]);
  
  // Datos del Cliente (Usuario)
  const tName = String(u.nombre_norm || '').trim();
  const tCuil = String(u.cuil_norm || '').trim();

  if (!tFecha) return { matched: false, duplicate: false, transfer_id, reason: 'Sin fecha válida' };
  if (!Number.isFinite(monto) || monto === 0) return { matched: false, duplicate: false, transfer_id, reason: 'Monto inválido' };

  let foundButUsed = false; 

  // ======================================================
  // PRIORIDAD 1: BÚSQUEDA POR CUIL
  // ======================================================
  if (tCuil) {
    // Buscamos DIRECTO en la lista de CUILs
    const poolCuil = movIndex.byCuilMonto.get(`${tCuil}|${monto}`) || [];

    for (const m of poolCuil) {
      const deltaMin = Math.abs(tFecha - m.fecha) / 60000;
      
      // Chequeo de Fecha (Es estricto según tu configuración)
      if (deltaMin <= cfg.time_window_minutes) {
        
        // Si ya se usó, marcamos bandera pero seguimos buscando otro libre
        if (usedSet && usedSet.has(m.movement_id)) {
          foundButUsed = true;
          continue; 
        }

        // ¡MATCH POR CUIL! (Éxito total, ignoramos el nombre)
        return {
          matched: true,
          duplicate: false,
          transfer_id,
          movement_id: m.movement_id,
          delta_min: Math.round(deltaMin * 10) / 10
        };
      }
    }
  }

  // ======================================================
  // PRIORIDAD 2: BÚSQUEDA POR NOMBRE (Respaldo)
  // ======================================================
  // Si llegamos acá es porque NO hubo match por CUIL (o no tenía).
  if (tName) {
    const poolName = movIndex.byNameMonto.get(`${tName}|${monto}`) || [];

    for (const m of poolName) {
      const deltaMin = Math.abs(tFecha - m.fecha) / 60000;
      
      if (deltaMin <= cfg.time_window_minutes) {
        if (usedSet && usedSet.has(m.movement_id)) {
          foundButUsed = true;
          continue;
        }

        // ¡MATCH POR NOMBRE!
        return {
          matched: true,
          duplicate: false,
          transfer_id,
          movement_id: m.movement_id,
          delta_min: Math.round(deltaMin * 10) / 10
        };
      }
    }
  }

  // Resultado Final
  if (foundButUsed) {
    return { matched: false, duplicate: true, transfer_id, reason: 'Duplicada (ya asignado)' };
  }

  return { matched: false, duplicate: false, transfer_id, reason: 'Sin match' };
}


function syncTransfers24h() {
  ensureAllSheets_();
  const cfg = getConfig_();
  const userIdx = getUserIndex_();

  let cutoff = getLastWalletCutoff_();
  if (!cutoff) cutoff = new Date(2020, 0, 1);

  const shT = sh_(SHEET.TRANSF);
  const lastRow = shT.getLastRow();
  
  if (lastRow < 2) {
    return { success: true, processed: 0, setMatch: 0, setNoMatch: 0, cutoff_at: formatDateTimeAR_(cutoff) };
  }

  // --- 1. RECOLECTAR DATOS ---
  const rowsToProcess = [];
  const usedMovementIds = new Set(); // Conjunto de billetes ya gastados

  let minDt = null;
  let maxDt = null;
  const CHUNK = 2000; 
  
  // LEEMOS TAMBIÉN LA COLUMNA L (Índice 12) DONDE GUARDAMOS EL ID DEL MOVIMIENTO
  for (let end = lastRow; end >= 2; end -= CHUNK) {
    const start = Math.max(2, end - CHUNK + 1);
    const num = end - start + 1;
    // Pedimos hasta la columna 12 (L)
    const cols = Math.max(12, shT.getLastColumn()); 
    const block = shT.getRange(start, 1, num, cols).getValues();

    for (let i = block.length - 1; i >= 0; i--) {
      const r = block[i];
      const transfer_id = String(r[0] || '').trim();
      if (!transfer_id) continue;

      const estado = String(r[8] || '').trim();
      const idMovimientoGuardado = String(r[11] || '').trim(); // Columna L (índice 11 en array)

      // Si ya está matcheada y tiene ID, lo guardamos en la lista negra
      if (estado === 'Match' && idMovimientoGuardado) {
        usedMovementIds.add(idMovimientoGuardado);
        continue; // No la procesamos de nuevo
      }

      if (estado === 'Estafador') continue; 
      
      const dt = parseARDateTime_(r[2]);
      if (!dt) continue;
      if (cutoff && (cutoff.getTime() - dt.getTime()) > (45 * 24 * 60 * 60 * 1000)) continue;
      if (dt > cutoff) continue;

      const rowIndex = start + i;
      rowsToProcess.push({ rowIndex, row: r, tFecha: dt, estado });

      if (!minDt || dt < minDt) minDt = dt;
      if (!maxDt || dt > maxDt) maxDt = dt;
    }
  }

  if (rowsToProcess.length === 0) {
    return { success: true, processed: 0, setMatch: 0, setNoMatch: 0, cutoff_at: formatDateTimeAR_(cutoff) };
  }

  // --- 2. TRAEMOS LA BILLETERA ---
  const spanMin = Math.max(60, cfg.time_window_minutes * 6);
  const movFrom = new Date(minDt.getTime() - spanMin * 60000);
  const movTo = new Date(Math.min(cutoff.getTime(), maxDt.getTime()) + spanMin * 60000);

  const movs = readMovsInRange_(movFrom, movTo, cfg.wallet_scan_tail_rows);
  const movIndex = buildMovIndex_(movs);

  // --- 3. PROCESAR ---
  let processed = 0, setMatch = 0, setNoMatch = 0;

  rowsToProcess.sort((a, b) => a.tFecha - b.tFecha);

  for (const tr of rowsToProcess) {
    // Buscamos match pasándole la lista 'usedMovementIds' que ya tiene los usados
    const res = strictFindMatchFast_(tr, userIdx, movIndex, cfg, usedMovementIds);
    
    if (res.matched) {
      shT.getRange(tr.rowIndex, 9).setValue('Match');
      shT.getRange(tr.rowIndex, 9).setBackground(null); // Limpiar color
      
      // ¡AQUÍ ESTÁ LA CLAVE! Guardamos el ID del movimiento en Columna L (12)
      shT.getRange(tr.rowIndex, 12).setValue(res.movement_id);
      
      usedMovementIds.add(res.movement_id); 
      setMatch++;
    } else if (res.duplicate) {
      shT.getRange(tr.rowIndex, 9).setValue('Duplicada');
      shT.getRange(tr.rowIndex, 9).setBackground('#fef08a'); // Pintamos amarillo suave
      setNoMatch++;
    } else {
      shT.getRange(tr.rowIndex, 9).setValue('Sin match');
      shT.getRange(tr.rowIndex, 9).setBackground(null);
      setNoMatch++;
    }
    processed++;
  }

  return { success: true, processed, setMatch, setNoMatch, cutoff_at: formatDateTimeAR_(cutoff) };
}

// -------------------- Trigger (opcional) --------------------
function setupDashboardMinuteTrigger() {
  // NOTA: Esto NO puede “refrescar” la UI abierta. Solo corre cada minuto.
  // La UI se refresca con polling (setInterval) desde app.html.
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'dashboardMinuteTick_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dashboardMinuteTick_').timeBased().everyMinutes(1).create();
  return { success:true };
}

function dashboardMinuteTick_() {
  // “Tick” de salud / log. Si querés, podrías cachear resultados.
  try {
    const d = getDashboardData();
    log_('DASH_TICK', `ok | ${d.updated_at}`);
  } catch (e) {
    log_('DASH_TICK_ERR', e.message);
  }
}


// -------------------- Lógica Hoja Control (Automatización) --------------------

function updateAndGetControlUrl(password) {
  if (password !== 'admin123') { 
    return { error: 'Contraseña incorrecta.' };
  }

  const ss = getSs_();
  const shControl = ss.getSheetByName(SHEET.CONTROL);
  const shTransf = sh_(SHEET.TRANSF);
  const shMovs = sh_(SHEET.MOVS);
  const userIdx = getUserIndex_(); 
  const cfg = getConfig_();

  const MAP_BILLETERAS = { 
    "01": "BARROSO", "1": "BARROSO",
    "02": "DC",      "2": "DC",
    "03": "OTRA",    "3": "OTRA",
    "04": "AMFRIS",  "4": "AMFRIS",
    "05": "MANZO",   "5": "MANZO" 
  };

  // ----------------------------------------------------------------
  // 1. RECOPILAR TRANSFERENCIAS MANUALES (Base del reporte)
  // ----------------------------------------------------------------
  const tData = shTransf.getDataRange().getValues();
  let allTransfers = [];
  const totalesDeclarados = {}; 
  const matchedPool = {}; 

  for (let i = 1; i < tData.length; i++) {
    const row = tData[i];
    const user_id = String(row[1] || '').trim();
    const fechaRaw = row[2];
    const monto = Number(row[3]) || 0;
    const estado = String(row[8] || 'Pendiente');
    const nota = String(row[9] || '');
    
    let billeteraNombre = String(row[10] || 'OTRA').toUpperCase().trim();
    if (!billeteraNombre) billeteraNombre = 'OTRA';

    const fechaObj = parseARDateTime_(fechaRaw);
    if (!fechaObj) continue; 
    const fechaStr = Utilities.formatDate(fechaObj, TZ, 'dd/MM/yyyy');
    
    const u = userIdx[user_id];
    let nombreCliente = (u && u.nombre) ? u.nombre : 'Cliente Desconocido';

    let turno = '-';
    const turnoMatch = nota.match(/Turno=([a-zA-Z]+)/);
    if (turnoMatch) turno = turnoMatch[1];

    const ingresoBilletera = (estado === 'Match') ? monto : '';

    allTransfers.push({
      timestamp: fechaObj.getTime(),
      rowArray: [fechaStr, turno, nombreCliente, monto, ingresoBilletera, estado],
      nombreCliente: nombreCliente,
      monto: monto,
      fechaStr: fechaStr,
      estado: estado
    });

    // Agrupamos Totales Manuales (Esto define las filas de la derecha)
    const key = fechaStr + "|" + billeteraNombre;
    if (!totalesDeclarados[key]) {
      totalesDeclarados[key] = {
        fecha: fechaStr,
        monto: 0,
        billetera: billeteraNombre,
        timestamp: fechaObj.getTime()
      };
    }
    totalesDeclarados[key].monto += monto;

    if (estado === 'Match') {
      const normName = normalizeName_(nombreCliente);
      const keyPool = `${normName}|${monto}`;
      if (!matchedPool[keyPool]) matchedPool[keyPool] = [];
      matchedPool[keyPool].push({ time: fechaObj.getTime(), used: false });
    }
  }

  allTransfers.sort((a, b) => a.timestamp - b.timestamp);

  const filasParaEscribir = [];
  for (let i = 0; i < allTransfers.length; i++) {
    const item = allTransfers[i];
    if (i % 10 === 0 && i > 0) filasParaEscribir.push(['', '', '', '', '', '']); 
    filasParaEscribir.push(item.rowArray);
  }

  // ----------------------------------------------------------------
  // 2. RECORRER BILLETERA (CALCULAR TOTALES REALES + SOBRANTES)
  // ----------------------------------------------------------------
  const mData = shMovs.getDataRange().getValues();
  const sobrantesList = []; 
  const totalesSobrantes = {}; 
  
  // Acumulador para totales reales (Col L y M)
  const totalesReales = {}; 

  for (let i = 1; i < mData.length; i++) {
    const row = mData[i];
    
    // Identificamos Billetera del CSV
    // Columna B es index 1.
    const idBill = String(row[1] || "3").trim(); 
    const nombreBillMov = MAP_BILLETERAS[idBill] || "OTRA";

    const mFechaRaw = row[2];
    let mMonto = row[4];
    
    if (typeof mMonto === 'string') {
      let s = mMonto.replace(/[$.]/g, '').replace(',', '.').trim();
      mMonto = parseFloat(s);
    } else {
      mMonto = Number(mMonto);
    }

    if (!mFechaRaw || !Number.isFinite(mMonto)) continue;

    const mFechaObj = parseARDateTime_(mFechaRaw);
    if (!mFechaObj) continue;
    const mFechaStr = Utilities.formatDate(mFechaObj, TZ, 'dd/MM/yyyy');

    // --- CÁLCULO DE TOTALES REALES (LO IMPORTANTE) ---
    // Sumamos TODO lo que hay en el CSV para ese día y billetera
    const keyReal = mFechaStr + "|" + nombreBillMov;
    
    if (!totalesReales[keyReal]) {
      totalesReales[keyReal] = { in: 0, out: 0 };
    }
    
    if (mMonto > 0) {
      totalesReales[keyReal].in += mMonto; // Suma ingresos
    } else {
      totalesReales[keyReal].out += mMonto; // Suma retiros (negativos)
    }
    // -------------------------------------------------

    // Filtro para lógica de Match/Sobrantes (sigue ignorando negativos)
    if (mMonto <= 0) continue;

    const mNombreNorm = String(row[7] || '').trim(); 
    const mNombreOriginal = String(row[5] || 'Desconocido'); 
    const keyMatch = `${mNombreNorm}|${mMonto}`;
    let isAccountedFor = false;

    if (matchedPool[keyMatch]) {
      const windowMs = cfg.time_window_minutes * 60 * 1000; 
      for (let k = 0; k < matchedPool[keyMatch].length; k++) {
        const candidate = matchedPool[keyMatch][k];
        if (!candidate.used) {
          const diff = Math.abs(candidate.time - mFechaObj.getTime());
          if (diff <= windowMs) {
            candidate.used = true; 
            isAccountedFor = true;
            break; 
          }
        }
      }
    }

    if (!isAccountedFor) {
      sobrantesList.push([mNombreOriginal, mMonto, nombreBillMov]);
      const keySob = mFechaStr + "|" + nombreBillMov;
      if (!totalesSobrantes[keySob]) {
        totalesSobrantes[keySob] = {
          fecha: mFechaStr,
          monto: 0,
          billetera: nombreBillMov,
          timestamp: mFechaObj.getTime()
        };
      }
      totalesSobrantes[keySob].monto += mMonto;
    }
  }

  // ----------------------------------------------------------------
  // 3. ESCRIBIR EN LA HOJA
  // ----------------------------------------------------------------

  shControl.getRange(2, 1, shControl.getMaxRows(), shControl.getMaxColumns()).clearContent();

  // A-F: Lista Principal
  if (filasParaEscribir.length > 0) {
    shControl.getRange(2, 1, filasParaEscribir.length, 6).setValues(filasParaEscribir);
  }

  // H-J: Totales Manuales
  const listaTotales = Object.values(totalesDeclarados).sort((a,b) => a.timestamp - b.timestamp);
  const arrFinalDecl = listaTotales.map(o => [o.fecha, o.monto, o.billetera]);
  
  if (arrFinalDecl.length > 0) {
    shControl.getRange(2, 8, arrFinalDecl.length, 3).setValues(arrFinalDecl);
  }

  // --- [MODIFICADO] L-M: TOTALES REALES (SIN COLUMNA N) ---
  const arrFinalReales = listaTotales.map(o => {
     // Usamos la fecha y billetera de la carga manual (H y J) para buscar el dato real
     const k = o.fecha + "|" + o.billetera;
     const stats = totalesReales[k] || { in: 0, out: 0 };
     // Devolvemos solo INGRESO y RETIRO
     return [stats.in, stats.out];
  });

  // Encabezados L-M
  shControl.getRange("L1:M1").setValues([["TOTAL RECAUDADO REAL (+)", "TOTAL RETIROS REAL (-)"]]);
  shControl.getRange("L1:M1").setFontWeight("bold").setBackground("#333").setFontColor("#FFF");

  // Escribimos solo 2 columnas de ancho
  if (arrFinalReales.length > 0) {
    shControl.getRange(2, 12, arrFinalReales.length, 2).setValues(arrFinalReales);
  }
  // -------------------------------------------------------------

  // P-R: Sobrantes
  shControl.getRange("P1:R1").setValues([["NO COINCIDEN C/REGISTRO ATENEA", "MONTO", "BILLETERA"]]);
  shControl.getRange("P1:R1").setFontWeight("bold").setBackground("#333").setFontColor("#FFF");
  if (sobrantesList.length > 0) {
    shControl.getRange(2, 16, sobrantesList.length, 3).setValues(sobrantesList);
  }

  // T-V: Totales Sobrantes
  shControl.getRange("T1:V1").setValues([["FECHA", "MONTO TOTAL $", "BILLETERA"]]);
  shControl.getRange("T1:V1").setFontWeight("bold").setBackground("#333").setFontColor("#FFF");
  const listaTotalesSob = Object.values(totalesSobrantes).sort((a,b) => a.timestamp - b.timestamp);
  const arrFinalSob = listaTotalesSob.map(o => [o.fecha, o.monto, o.billetera]);
  
  if (arrFinalSob.length > 0) {
    shControl.getRange(2, 20, arrFinalSob.length, 3).setValues(arrFinalSob);
  }

  return { success: true, url: ss.getUrl() + '#gid=' + shControl.getSheetId() };
}
// -------------------- Eliminar Transferencia --------------------
function deleteTransfer(transfer_id) {
  const hit = getTransferRow_(transfer_id); // Usamos el helper que ya tenías
  if (!hit) return { error: 'Transferencia no encontrada.' };

  const sh = sh_(SHEET.TRANSF);
  sh.deleteRow(hit.rowIndex); // Borra la fila completa
  
  log_('DELETE_TRANSFER', `ID: ${transfer_id} eliminada.`);
  return { success: true };
}

// -------------------- ZONA DE CIERRE (DANGER ZONE) --------------------

function deleteAllComprobantes(password) {
  if (password !== 'admin123') return { error: '⛔ Contraseña incorrecta.' };

  try {
    const cfg = getConfig_();
    if (!cfg.drive_folder_id) return { error: 'No hay ID de carpeta configurado.' };

    const folder = DriveApp.getFolderById(cfg.drive_folder_id);
    const files = folder.getFiles();
    
    let count = 0;
    while (files.hasNext()) {
      const file = files.next();
      file.setTrashed(true); // Los manda a la papelera (no los borra permanente por seguridad)
      count++;
    }
    
    log_('CIERRE_COMPROBANTES', `Se eliminaron ${count} archivos.`);
    return { success: true, count };
    
  } catch (e) {
    return { error: e.message };
  }
}

function deleteAllHistory(password) {
  if (password !== 'admin123') return { error: '⛔ Contraseña incorrecta.' };

  try {
    const ss = getSs_();
    
    // 1. Borrar Transferencias (El sistema)
    const shTransf = ss.getSheetByName(SHEET.TRANSF);
    if (shTransf.getLastRow() > 1) {
      shTransf.getRange(2, 1, shTransf.getMaxRows() - 1, shTransf.getMaxColumns()).clearContent();
    }

    // 2. Borrar Movimientos Billetera (El csv cargado)
    const shMovs = ss.getSheetByName(SHEET.MOVS);
    if (shMovs.getLastRow() > 1) {
      shMovs.getRange(2, 1, shMovs.getMaxRows() - 1, shMovs.getMaxColumns()).clearContent();
    }

    // 3. Borrar Hoja de Control (El reporte excel)
    const shControl = ss.getSheetByName(SHEET.CONTROL);
    if (shControl.getLastRow() > 1) {
      shControl.getRange(2, 1, shControl.getMaxRows() - 1, shControl.getMaxColumns()).clearContent();
    }

    // 4. Borrar Logs (Opcional, para empezar limpio)
    const shLogs = ss.getSheetByName(SHEET.LOGS);
    if (shLogs.getLastRow() > 1) {
      shLogs.getRange(2, 1, shLogs.getMaxRows() - 1, shLogs.getMaxColumns()).clearContent();
    }

    // NO TOCAMOS LA HOJA "USUARIOS" (Los clientes quedan guardados)

    log_('CIERRE_HISTORIAL', 'Se vaciaron Transferencias, Movimientos y Control.');
    return { success: true };

  } catch (e) {
    return { error: e.message };
  }
}

// -------------------- Obtener Sobrantes de Billetera (Huérfanos) --------------------
function getWalletOrphans(query) {
  const ss = getSs_();
  const shTransf = sh_(SHEET.TRANSF);
  const shMovs = sh_(SHEET.MOVS);
  const userIdx = getUserIndex_();
  const cfg = getConfig_();

  // Preparar búsqueda (si el usuario escribió algo)
  const qClean = query ? normalizeName_(String(query)) : '';

  // 1. Mapa de Matches (para descartar lo que ya se usó)
  const tData = shTransf.getDataRange().getValues();
  const matchedPool = {}; 

  for (let i = 1; i < tData.length; i++) {
    const row = tData[i];
    const user_id = String(row[1] || '').trim();
    const fechaRaw = row[2];
    const monto = Number(row[3]);
    const estado = String(row[8] || '');

    if (estado !== 'Match') continue; 

    const fechaObj = parseARDateTime_(fechaRaw);
    if (!fechaObj) continue;

    const u = userIdx[user_id];
    const nombreCliente = (u && u.nombre) ? u.nombre : 'Cliente Desconocido';
    
    const normName = normalizeName_(nombreCliente);
    const key = `${normName}|${monto}`;
    
    if (!matchedPool[key]) matchedPool[key] = [];
    matchedPool[key].push({ time: fechaObj.getTime(), used: false });
  }

  // 2. Definir fecha de corte (Ayer a las 00:00:00 para atrás se ignora)
  const now = new Date();
  const yesterdayStart = new Date(now);
  yesterdayStart.setDate(now.getDate() - 1);
  yesterdayStart.setHours(0, 0, 0, 0);
  const minTime = yesterdayStart.getTime();

  // 3. Buscar en Billetera
  const mData = shMovs.getDataRange().getValues();
  const orphans = [];
  
  // Leemos de atrás para adelante
  for (let i = mData.length - 1; i >= 1; i--) {
    const row = mData[i];
    const mFechaRaw = row[2];
    
    // Limpieza de monto
    let mMonto = row[4];
    if (typeof mMonto === 'string') {
      let s = mMonto.replace(/[$.]/g, '').replace(',', '.').trim();
      mMonto = parseFloat(s);
    } else {
      mMonto = Number(mMonto);
    }

    const mNombreNorm = String(row[7] || '').trim(); 
    const mNombreOriginal = String(row[5] || 'Desconocido'); 

    // Filtros Básicos
    if (!mFechaRaw || !Number.isFinite(mMonto) || mMonto <= 0) continue;

    const mFechaObj = parseARDateTime_(mFechaRaw);
    if (!mFechaObj) continue;

    // --- FILTRO DE FECHA (SOLO AYER Y HOY) ---
    if (mFechaObj.getTime() < minTime) {
      continue; // Es muy viejo, lo saltamos
    }

    // --- FILTRO DE BÚSQUEDA (SI EL USUARIO ESCRIBIÓ ALGO) ---
    if (qClean) {
      const matchName = normalizeName_(mNombreOriginal).includes(qClean);
      const matchMonto = String(mMonto).includes(qClean);
      // Si no coincide ni nombre ni monto, saltamos
      if (!matchName && !matchMonto) continue;
    }

    // Chequeo Match (¿Ya tiene dueño?)
    const key = `${mNombreNorm}|${mMonto}`;
    let isAccountedFor = false;

    if (matchedPool[key]) {
      const windowMs = cfg.time_window_minutes * 60 * 1000; 
      for (let k = 0; k < matchedPool[key].length; k++) {
        const candidate = matchedPool[key][k];
        if (!candidate.used) {
          const diff = Math.abs(candidate.time - mFechaObj.getTime());
          if (diff <= windowMs) {
            candidate.used = true; 
            isAccountedFor = true;
            break; 
          }
        }
      }
    }

    // Si NO tiene dueño -> ES UN INGRESO NO REGISTRADO
    if (!isAccountedFor) {
      orphans.push({
        fecha: formatDateTimeAR_(mFechaObj),
        nombre: mNombreOriginal,
        monto: mMonto
      });
    }
  }

  // Ordenar: Más reciente arriba
  orphans.sort((a, b) => {
    const da = parseARDateTime_(a.fecha);
    const db = parseARDateTime_(b.fecha);
    return (db || 0) - (da || 0);
  });

  return orphans;
}

// -------------------- NUEVAS FUNCIONES: REPORTE Y LOGS --------------------

function getSystemLogs(limit) {
  const sh = sh_(SHEET.LOGS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  
  const start = Math.max(2, lastRow - limit + 1);
  const data = sh.getRange(start, 1, lastRow - start + 1, 4).getValues();
  
  // Devolver invertido (más nuevo primero)
  return data.reverse().map(r => ({
    fecha: r[0] ? Utilities.formatDate(new Date(r[0]), TZ, 'dd/MM HH:mm') : '-',
    accion: r[1],
    detalle: r[2],
    usuario: r[3] || 'Sistema'
  }));
}

function saveTransferChanges(payload) {
  const transfer_id = String(payload?.transfer_id || '').trim();
  if (!transfer_id) return { error: 'Falta transfer_id.' };

  // 1. Actualizar Transferencia (Fecha/Monto)
  const patchT = payload?.transfer_patch || {};
  if (patchT && (patchT.fecha_hora_operativa || patchT.monto != null)) {
    updateTransfer(transfer_id, patchT);
  }

  // 2. Actualizar Usuario (Nombre/CUIL)
  const hit = getTransferRow_(transfer_id);
  if (hit) {
    const userId = hit.row[1]; // Columna B es el UserID
    const patchU = payload?.user_patch || {};
    if (patchU && (patchU.nombre_canonico || patchU.cuit_cuil)) {
      updateUser(userId, patchU);
    }
  }

  return { success: true };
}

function getTransfersWithNotes() {
  const sh = sh_(SHEET.TRANSF);
  // Leemos las últimas 1000 transferencias para no saturar
  const { data, startRow } = getTransfersTail_(1000); 
  const userIdx = getUserIndex_();
  const out = [];

  // Recorremos de la más nueva a la más vieja
  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];
    const nota = String(r[9] || ''); // Columna J (Índice 9) es Nota

    // CRITERIO: Solo si tiene el separador "|" significa que hay texto extra del usuario
    // Ejemplo: "Turno=TARDE | Pide factura A"
    if (nota.includes('|')) {
      const user_id = String(r[1] || '').trim();
      const u = userIdx[user_id] || {};
      const dt = parseARDateTime_(r[2]);
      
      // Limpiamos la nota para mostrar solo la parte del mensaje
      const parts = nota.split('|');
      const mensaje = parts.length > 1 ? parts.slice(1).join('|').trim() : nota;

      out.push({
        transfer_id: r[0],
        nombre: u.nombre_canonico || 'Desconocido',
        fecha: dt ? formatDateTimeAR_(dt) : (r[2] || ''),
        monto: r[3],
        estado: r[8],
        nota_completa: nota,
        mensaje_corto: mensaje
      });
    }
  }
  return out;
}
// Agregá esto al final de tu archivo Código.gs
function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const action = params.action;
  const payload = params.payload;

  // Aquí ruteamos las funciones según lo que pida GitHub
  if (action === 'searchUsers') return jsonResponse(searchUsers(payload.query));
  if (action === 'createUser') return jsonResponse(createUser(payload));
  // ... seguiremos sumando el resto
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
