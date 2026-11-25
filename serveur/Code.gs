// Stockage des comptes MapPrep dans ce Google Sheet (onglet « Comptes »).
// À coller dans Extensions › Apps Script, puis Déployer › Nouveau déploiement › Application Web
// (Exécuter en tant que : moi ; Qui a accès : Tout le monde).
const SHEET = 'Comptes';

function doGet() { return out({ ok: true, app: 'mapprep' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return out(handle(JSON.parse(e.postData.contents))); }
  catch (err) { return out({ ok: false, error: 'server', detail: String(err) }); }
  finally { lock.releaseLock(); }
}

function handle(r) {
  const sh = sheet();
  const name = String(r.name || '').trim().toLowerCase();
  if (!name || !r.hash) return { ok: false, error: 'missing' };
  const vals = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === name) { row = i + 1; break; }
  const data = JSON.stringify(r.data || {});
  if (data.length > 49000) return { ok: false, error: 'too_big' };

  if (r.action === 'register') {
    if (row > 0) return { ok: false, error: 'exists' };
    sh.appendRow([name, r.hash, data, Number(r.updated) || Date.now(), new Date()]);
    return { ok: true, updated: Number(r.updated) || Date.now() };
  }
  if (row < 0) return { ok: false, error: 'unknown' };
  const v = vals[row - 1];
  if (String(v[1]) !== r.hash) return { ok: false, error: 'password' };
  if (r.action === 'login') return { ok: true, data: JSON.parse(v[2] || '{}'), updated: Number(v[3]) || 0 };
  if (r.action === 'save') {
    sh.getRange(row, 3, 1, 3).setValues([[data, Number(r.updated) || Date.now(), new Date()]]);
    return { ok: true, updated: Number(r.updated) || Date.now() };
  }
  return { ok: false, error: 'action' };
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(SHEET); }
  if (sh.getLastRow() === 0) sh.appendRow(['nom', 'mdp_hash', 'donnees', 'maj', 'date_maj']);
  return sh;
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
