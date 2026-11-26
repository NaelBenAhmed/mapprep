// Stockage des comptes et des signalements MapPrep dans ce Google Sheet
// (onglets « Comptes » et « Signalements »).
// À coller dans Extensions › Apps Script, puis Déployer › Gérer les déploiements › Modifier (crayon)
// › Version : Nouvelle version › Déployer. L'adresse de l'appli Web ne change pas.
// (Exécuter en tant que : moi ; Qui a accès : Tout le monde) : les préparateurs n'ont pas besoin de compte Google.
const SHEET = 'Comptes';
const REPORTS = 'Signalements';
const R_HEAD = ['id', 'type', 'x', 'y', 'lieu', 'note', 'auteur', 'cree', 'date', 'expire', 'statut', 'confirmations', 'maj', 'regle_par', 'date_regle'];
const R_TYPES = ['filmeuse_hs', 'imprimante_hs', 'flaque', 'obstacle', 'allee_bloquee', 'casse', 'eclairage', 'autre'];
const MAX_TTL = 72 * 3600 * 1000;

function doGet() { return out({ ok: true, app: 'mapprep' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return out(handle(JSON.parse(e.postData.contents))); }
  catch (err) { return out({ ok: false, error: 'server', detail: String(err) }); }
  finally { lock.releaseLock(); }
}

function handle(r) {
  // Lecture des signalements : ouverte, et très légère si rien n'a changé depuis la dernière fois.
  if (r.action === 'reports') return listReports(Number(r.since) || 0);

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
  // Signalements : il faut un compte de l'appli (nom + mot de passe) pour écrire.
  if (r.action === 'report') return addReport(r, name);
  if (r.action === 'confirm' || r.action === 'resolve') return updateReport(r, name);
  return { ok: false, error: 'action' };
}

function addReport(r, name) {
  const sh = reportSheet();
  const type = R_TYPES.indexOf(r.type) >= 0 ? r.type : 'autre';
  const now = Date.now();
  const ttl = Math.min(MAX_TTL, Math.max(15 * 60 * 1000, Number(r.ttl) || 8 * 3600 * 1000));
  const id = String(r.id || ('s' + now.toString(36))).slice(0, 24);
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
  if (!ids.some(x => String(x[0]) === id)) {
    sh.appendRow([id, type, Math.round(Number(r.x)) || 0, Math.round(Number(r.y)) || 0, String(r.label || '').slice(0, 60),
      String(r.note || '').slice(0, 140), name, now, new Date(), now + ttl, 'actif', 0, now, '', '']);
  }
  return listReports(0);
}

function updateReport(r, name) {
  const sh = reportSheet();
  const vals = sh.getDataRange().getValues();
  const now = Date.now();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) !== String(r.id)) continue;
    if (r.action === 'confirm') {
      // « Toujours là » : on prolonge la durée de vie du signalement.
      const ttl = Math.min(MAX_TTL, Math.max(15 * 60 * 1000, Number(r.ttl) || 8 * 3600 * 1000));
      sh.getRange(i + 1, 10, 1, 4).setValues([[Math.max(Number(vals[i][9]) || 0, now + ttl), 'actif', (Number(vals[i][11]) || 0) + 1, now]]);
    } else {
      sh.getRange(i + 1, 11, 1, 5).setValues([['regle', vals[i][11], now, name, new Date()]]);
    }
    break;
  }
  return listReports(0);
}

function listReports(since) {
  const sh = reportSheet();
  const vals = sh.getDataRange().getValues();
  const now = Date.now();
  let ver = 0;
  const list = [];
  for (let i = 1; i < vals.length; i++) {
    const v = vals[i];
    const maj = Number(v[12]) || 0, exp = Number(v[9]) || 0;
    ver = Math.max(ver, maj, exp < now ? exp : 0);
    if (v[10] !== 'actif' || exp < now) continue;
    list.push({ id: String(v[0]), type: v[1], x: Number(v[2]), y: Number(v[3]), label: v[4], note: v[5], by: v[6], at: Number(v[7]), exp, conf: Number(v[11]) || 0 });
  }
  if (since && since === ver) return { ok: true, same: true, ver };
  return { ok: true, ver, reports: list };
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(SHEET); }
  if (sh.getLastRow() === 0) sh.appendRow(['nom', 'mdp_hash', 'donnees', 'maj', 'date_maj']);
  return sh;
}

function reportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(REPORTS);
  if (!sh) sh = ss.insertSheet(REPORTS);
  if (sh.getLastRow() === 0) { sh.appendRow(R_HEAD); sh.setFrozenRows(1); }
  return sh;
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
