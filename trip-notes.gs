// ═══════════════════════════════════════════════════════════════════════
// Trip Notes — Google Apps Script Backend
// ═══════════════════════════════════════════════════════════════════════
// 初次設定步驟：
//   1. 執行 initSheets()，從 Logger 複製 Drive Folder ID 填入 IMAGE_FOLDER_ID
//   2. Deploy → New Deployment → Web App
//      Execute as: Me | Who has access: Anyone (including anonymous)
//   3. 複製 Web App URL，貼入旅行規劃站「設定 → Apps Script URL」
// ═══════════════════════════════════════════════════════════════════════

const SS_ID = '1GTnDgfwZVAvEYySHPf8Kzo2pBv4T-81zfBPDU22aQe4';
const IMAGE_FOLDER_ID = '1-y1v6PATDvgafqPQamQgi5YpH81HGywu';

const NOTE_COLS = [
  'noteId','tripId','date','category','text','imageRefs',
  'expenseAmount','expenseCurrency','expensePayer',
  'createdBy','createdByDevice','createdAt','updatedBy','updatedAt','deleted'
];
const IMAGE_COLS = [
  'imageId','noteId','driveFileId','driveThumbnailId',
  'imageUrl','thumbnailUrl','filename','uploadedBy','uploadedAt'
];

// ── HTTP Router ──────────────────────────────────────────────────────────
function doGet(e) {
  const p = e.parameter;
  try {
    if (p.action === 'getShareInfo')     return ok(getShareInfo(p.token));
    if (p.action === 'getNotes')         return ok(getNotes(p.tripId, p.token));
    if (p.action === 'inferDestination') return ok(inferDestination_(p));
    return fail('UNKNOWN_ACTION', '不支援的操作');
  } catch (ex) { return fail(ex.code || 'SERVER_ERROR', ex.message); }
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); }
  catch (_) { return fail('PARSE_ERROR', '請求格式錯誤'); }
  try {
    if (b.action === 'createNote')  return ok(createNote(b));
    if (b.action === 'updateNote')  return ok(updateNote(b));
    if (b.action === 'deleteNote')  return ok(deleteNote(b));
    if (b.action === 'uploadImage')       return ok(uploadImage(b));
    if (b.action === 'deleteImage')       return ok(deleteImage(b));
    if (b.action === 'inferDestination')  return ok(inferDestination_(b));
    return fail('UNKNOWN_ACTION', '不支援的操作');
  } catch (ex) { return fail(ex.code || 'SERVER_ERROR', ex.message); }
}

// ── Response ─────────────────────────────────────────────────────────────
function ok(data) {
  return out({ success: true, data, error: null, ts: new Date().toISOString() });
}
function fail(code, message) {
  return out({ success: false, data: null, error: { code, message }, ts: new Date().toISOString() });
}
function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function AppError(code, msg) {
  const e = new Error(msg); e.code = code; return e;
}

// ── Lock ─────────────────────────────────────────────────────────────────
function withLock(cb) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); return cb(); }
  finally { lock.releaseLock(); }
}

// ── Sheet helpers ─────────────────────────────────────────────────────────
function ss() { return SpreadsheetApp.openById(SS_ID); }
function getSheet(name) {
  const s = ss().getSheetByName(name);
  if (!s) throw AppError('SHEET_NOT_FOUND', '找不到工作表: ' + name);
  return s;
}
function sheetRows(name) {
  const s = getSheet(name);
  const v = s.getDataRange().getValues();
  const h = v[0] || [];
  const data = v.slice(1).map((r, i) => {
    const o = { _row: i + 2 };
    h.forEach((k, j) => { o[k] = (r[j] !== undefined && r[j] !== null) ? r[j] : ''; });
    return o;
  });
  return { s, h, data };
}

// ── _shares （無 header row，直接用欄位索引）────────────────────────────
// A=token, B=name, C=trips(|分隔 or 'all'), D=url, E=nicknameOptions(|分隔)
function getShareInfo(token) {
  if (!token) throw AppError('MISSING_TOKEN', '缺少 token');
  const sheetObj = ss().getSheetByName('_shares');
  if (!sheetObj) throw AppError('SHEET_NOT_FOUND', '找不到 _shares 工作表');
  const vals = sheetObj.getDataRange().getValues();
  for (const row of vals) {
    if ((row[0] || '').toString().trim() !== token) continue;
    const tripsRaw = (row[2] || '').toString().trim();
    if (!tripsRaw) throw AppError('NO_PERMISSION', '分享連結未設定權限');
    const isOwner = tripsRaw === 'all';
    const tripList = isOwner ? ['*'] : tripsRaw.split('|').map(x => x.trim()).filter(Boolean);
    const nicknameOptions = (row[4] || '').toString().trim()
      ? (row[4]).toString().split('|').map(x => x.trim()).filter(Boolean)
      : [];
    return {
      token,
      name: (row[1] || '').toString().trim(),
      trips: tripList,
      isOwner,
      nicknameOptions
    };
  }
  throw AppError('TOKEN_NOT_FOUND', '分享連結無效');
}

function validateToken(token, tripId) {
  const info = getShareInfo(token);
  if (tripId && !info.isOwner && !info.trips.includes(tripId)) {
    throw AppError('PERMISSION_DENIED', '無權存取此旅程');
  }
  return info;
}

// ── getNotes ──────────────────────────────────────────────────────────────
function getNotes(tripId, token) {
  validateToken(token, tripId);
  const { data } = sheetRows('trip_notes');
  const notes = data
    .filter(r => r.tripId === tripId && !isTrue(r.deleted))
    .map(formatNote);

  const { data: imgData } = sheetRows('trip_note_images');
  const imgMap = {};
  for (const img of imgData) {
    if (!imgMap[img.noteId]) imgMap[img.noteId] = [];
    imgMap[img.noteId].push({ imageId: img.imageId, imageUrl: img.imageUrl, thumbnailUrl: img.thumbnailUrl });
  }
  notes.forEach(n => { n.images = imgMap[n.noteId] || []; });

  return { notes };
}

// ── createNote ────────────────────────────────────────────────────────────
function createNote(body) {
  const { token, tripId, note } = body;
  validateToken(token, tripId);
  const noteId = 'note_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6);
  const now = new Date().toISOString();
  withLock(() => {
    getSheet('trip_notes').appendRow([
      noteId, tripId,
      note.date || '', note.category || 'note', note.text || '',
      (note.imageRefs || []).join('|'),
      note.expenseAmount !== null && note.expenseAmount !== undefined ? note.expenseAmount : '',
      note.expenseCurrency || '',
      note.expensePayer || '',
      note.createdBy || '', note.createdByDevice || '', now,
      note.createdBy || '', now, false
    ]);
  });
  return { noteId };
}

// ── updateNote ────────────────────────────────────────────────────────────
function updateNote(body) {
  const { token, noteId, updates, updatedBy } = body;
  withLock(() => {
    const { s, h, data } = sheetRows('trip_notes');
    const row = data.find(r => r.noteId === noteId);
    if (!row) throw AppError('NOTE_NOT_FOUND', '找不到筆記');
    validateToken(token, row.tripId);
    const now = new Date().toISOString();
    const patch = { ...updates, updatedBy: updatedBy || '', updatedAt: now };
    if (Array.isArray(patch.imageRefs)) patch.imageRefs = patch.imageRefs.join('|');
    for (const [k, v] of Object.entries(patch)) {
      const col = h.indexOf(k);
      if (col >= 0) s.getRange(row._row, col + 1).setValue(v === null ? '' : v);
    }
  });
  return { noteId };
}

// ── deleteNote ────────────────────────────────────────────────────────────
function deleteNote(body) {
  const { token, noteId, requestDevice } = body;
  withLock(() => {
    const { s, h, data } = sheetRows('trip_notes');
    const row = data.find(r => r.noteId === noteId);
    if (!row) throw AppError('NOTE_NOT_FOUND', '找不到筆記');
    const shareInfo = validateToken(token, row.tripId);
    if (!shareInfo.isOwner && row.createdByDevice !== requestDevice) {
      throw AppError('PERMISSION_DENIED', '無刪除權限');
    }
    s.getRange(row._row, h.indexOf('deleted') + 1).setValue(true);
    _deleteImagesForNote(noteId);
  });
  return { noteId };
}

function _deleteImagesForNote(noteId) {
  const { s, data } = sheetRows('trip_note_images');
  const targets = data.filter(r => r.noteId === noteId).reverse();
  for (const r of targets) {
    try { DriveApp.getFileById(r.driveFileId).setTrashed(true); } catch (_) {}
    try { DriveApp.getFileById(r.driveThumbnailId).setTrashed(true); } catch (_) {}
    s.deleteRow(r._row);
  }
}

// ── uploadImage ───────────────────────────────────────────────────────────
function uploadImage(body) {
  const { token, tripId, noteId, filename, originalBase64, thumbnailBase64, uploadedBy } = body;
  validateToken(token, tripId);
  const imageId = 'img_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6);
  const root  = getImageRootFolder();
  const trip  = getOrCreateFolder(root, tripId);
  const origF = getOrCreateFolder(trip, 'original');
  const thmF  = getOrCreateFolder(trip, 'thumbnail');

  function saveToDrive(folder, b64, name) {
    const raw  = b64.includes(',') ? b64.split(',')[1] : b64;
    const blob = Utilities.newBlob(Utilities.base64Decode(raw), 'image/jpeg', name);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { id: file.getId(), url: 'https://drive.google.com/uc?id=' + file.getId() };
  }

  const orig  = saveToDrive(origF, originalBase64,  imageId + '.jpg');
  const thumb = saveToDrive(thmF,  thumbnailBase64, imageId + '_t.jpg');
  const now   = new Date().toISOString();

  withLock(() => {
    getSheet('trip_note_images').appendRow([
      imageId, noteId, orig.id, thumb.id, orig.url, thumb.url,
      filename || '', uploadedBy || '', now
    ]);
  });
  return { imageId, imageUrl: orig.url, thumbnailUrl: thumb.url };
}

// ── deleteImage ───────────────────────────────────────────────────────────
function deleteImage(body) {
  const { token, imageId } = body;
  withLock(() => {
    const { s, data } = sheetRows('trip_note_images');
    const row = data.find(r => r.imageId === imageId);
    if (!row) throw AppError('IMAGE_NOT_FOUND', '找不到圖片');
    const { data: noteData } = sheetRows('trip_notes');
    const note = noteData.find(r => r.noteId === row.noteId);
    if (note) validateToken(token, note.tripId);
    try { DriveApp.getFileById(row.driveFileId).setTrashed(true); } catch (_) {}
    try { DriveApp.getFileById(row.driveThumbnailId).setTrashed(true); } catch (_) {}
    s.deleteRow(row._row);
  });
  return { imageId };
}

// ── Helpers ───────────────────────────────────────────────────────────────
function isTrue(v) { return v === true || v === 'TRUE' || v === 'true'; }

function formatNote(r) {
  return {
    noteId: r.noteId, tripId: r.tripId,
    date: r.date ? (r.date instanceof Date ? Utilities.formatDate(r.date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : r.date.toString()) : null,
    category: r.category || 'note',
    text: r.text || '',
    imageRefs: r.imageRefs ? r.imageRefs.toString().split('|').filter(Boolean) : [],
    expenseAmount: (r.expenseAmount !== '' && r.expenseAmount !== null) ? Number(r.expenseAmount) : null,
    expenseCurrency: r.expenseCurrency || null,
    expensePayer: r.expensePayer || null,
    createdBy: r.createdBy || '',
    createdByDevice: r.createdByDevice || '',
    createdAt: r.createdAt ? (r.createdAt instanceof Date ? r.createdAt.toISOString() : r.createdAt.toString()) : '',
    updatedBy: r.updatedBy || '',
    updatedAt: r.updatedAt ? (r.updatedAt instanceof Date ? r.updatedAt.toISOString() : r.updatedAt.toString()) : ''
  };
}

// ── Drive ─────────────────────────────────────────────────────────────────
function getImageRootFolder() {
  if (IMAGE_FOLDER_ID) return DriveApp.getFolderById(IMAGE_FOLDER_ID);
  const iter = DriveApp.getFoldersByName('travel-planner-images');
  return iter.hasNext() ? iter.next() : DriveApp.createFolder('travel-planner-images');
}
function getOrCreateFolder(parent, name) {
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}

// ── inferDestination ─────────────────────────────────────────────────────
function inferDestination_(p) {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw AppError('NO_API_KEY', 'GAS 尚未設定 OPENAI_API_KEY');
  const prompt = '你是旅遊資料推論助理。根據旅程名稱判斷目的地國家和主要城市。\n旅程名稱：' + p.title + '\n日期：' + p.startDate + ' ～ ' + p.endDate + '\n只回傳 JSON：{"countries":[{"code":"JP","name_zh":"日本"}],"cities":[{"name_zh":"東京","country_code":"JP"}],"confidence":0.95}\n規則：台灣旅程 code 用 "TW"；無法判斷時 confidence<0.3，countries/cities 留空陣列。';
  const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post', contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + key },
    payload: JSON.stringify({
      model: 'gpt-4o-mini',
      messages: [{ role: 'user', content: prompt }],
      response_format: { type: 'json_object' },
      temperature: 0.1, max_tokens: 300
    }),
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200)
    throw AppError('OPENAI_ERROR', 'OpenAI 回應錯誤 ' + resp.getResponseCode());
  return JSON.parse(JSON.parse(resp.getContentText()).choices[0].message.content);
}

// ── One-time setup ────────────────────────────────────────────────────────
function initSheets() {
  const sp = ss();

  if (!sp.getSheetByName('trip_notes')) {
    const s = sp.insertSheet('trip_notes');
    s.appendRow(NOTE_COLS);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, NOTE_COLS.length).setFontWeight('bold');
    Logger.log('✅ 建立 trip_notes');
  } else {
    Logger.log('⏭ trip_notes 已存在，略過');
  }

  if (!sp.getSheetByName('trip_note_images')) {
    const s = sp.insertSheet('trip_note_images');
    s.appendRow(IMAGE_COLS);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, IMAGE_COLS.length).setFontWeight('bold');
    Logger.log('✅ 建立 trip_note_images');
  } else {
    Logger.log('⏭ trip_note_images 已存在，略過');
  }

  const folder = getImageRootFolder();
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📁 Drive 圖片資料夾 ID:');
  Logger.log(folder.getId());
  Logger.log('請將此 ID 填入 IMAGE_FOLDER_ID 後重新部署');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
}

// ── 一次性工具：寫入 _shares public_token（執行一次即可，完成後可刪除）──────────
function fillSharesPublicTokens() {
  const SS_ID = '1GTnDgfwZVAvEYySHPf8Kzo2pBv4T-81zfBPDU22aQe4';
  const tokens = {
    'astravel': 'N6epWWbPuLNi',
    'family':   '37UlCAS4EqGt',
    'nitta':    'RrgUy8YW2t5L',
    'sandy':    'EjdVNGhFVelh',
    'mars':     'C7F391OwVJFF',
  };
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('_shares');
  const data  = sheet.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const id = (data[i][0] || '').trim();
    if (tokens[id] && !data[i][5]) {
      sheet.getRange(i + 1, 6).setValue(tokens[id]);
      Logger.log(`✅ ${id} → ${tokens[id]}`);
      updated++;
    }
  }
  Logger.log(`共更新 ${updated} 列`);
}

