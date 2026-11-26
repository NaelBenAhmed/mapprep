// Stockage MapPrep dans ce Google Sheet : onglets « Comptes », « Signalements » et « Plan » (repères communs).
// À coller dans Extensions › Apps Script, puis Déployer › Gérer les déploiements › Modifier (crayon)
// › Version : Nouvelle version › Déployer. L'adresse de l'appli Web ne change pas.
// (Exécuter en tant que : moi ; Qui a accès : Tout le monde) : les préparateurs n'ont pas besoin de compte Google.
const SHEET = 'Comptes';
const REPORTS = 'Signalements';
const PLAN = 'Plan';
// Administrateurs : voient tous les comptes, posent des signalements sans limite de durée, modifient le plan.
// Noms comparés sans accents ni majuscules.
const ADMINS = ['nael ben ahmed', 'mickael'];
const R_HEAD = ['id', 'type', 'x', 'y', 'lieu', 'note', 'auteur', 'cree', 'date', 'expire', 'statut', 'confirmations', 'maj', 'regle_par', 'date_regle'];
const P_HEAD = ['id', 'type', 'nom', 'detail', 'x', 'y', 'supprime', 'maj', 'par', 'date'];
const R_TYPES = ['filmeuse_hs', 'imprimante_hs', 'flaque', 'obstacle', 'allee_bloquee', 'casse', 'eclairage', 'autre'];
const MAX_TTL = 72 * 3600 * 1000, ADMIN_TTL = 30 * 24 * 3600 * 1000, FOREVER = 10 * 365 * 24 * 3600 * 1000;

function norm(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function isAdmin(name) { return ADMINS.indexOf(norm(name)) >= 0; }

function doGet() { return out({ ok: true, app: 'mapprep' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return out(handle(JSON.parse(e.postData.contents))); }
  catch (err) { return out({ ok: false, error: 'server', detail: String(err) }); }
  finally { lock.releaseLock(); }
}

function handle(r) {
  // Lecture des signalements et du plan commun : ouverte, et très légère si rien n'a changé.
  if (r.action === 'reports') return listReports(Number(r.since) || 0);

  const sh = sheet();
  const vals = sh.getDataRange().getValues();

  // Connexion par badge (code-barre) : le code scanné remplace le nom + mot de passe.
  if (r.action === 'badge_login') {
    const b = String(r.badge || '');
    if (!b) return { ok: false, error: 'missing' };
    for (let i = 1; i < vals.length; i++) if (String(vals[i][5]) === b) {
      return { ok: true, name: String(vals[i][0]), hash: String(vals[i][1]), data: JSON.parse(vals[i][2] || '{}'), updated: Number(vals[i][3]) || 0, admin: isAdmin(vals[i][0]) };
    }
    return { ok: false, error: 'no_badge' };
  }

  const name = String(r.name || '').trim().toLowerCase();
  if (!name || !r.hash) return { ok: false, error: 'missing' };
  let row = -1;
  for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === name) { row = i + 1; break; }
  const data = JSON.stringify(r.data || {});
  if (data.length > 49000) return { ok: false, error: 'too_big' };
  const badgeUsed = b => !!b && vals.some((v, i) => i > 0 && String(v[5]) === b && String(v[0]) !== name);

  if (r.action === 'register') {
    if (row > 0) return { ok: false, error: 'exists' };
    const b = String(r.badge || '');
    if (badgeUsed(b)) return { ok: false, error: 'badge_used' };
    sh.appendRow([name, r.hash, data, Number(r.updated) || Date.now(), new Date(), b]);
    return { ok: true, updated: Number(r.updated) || Date.now(), admin: isAdmin(name) };
  }
  if (row < 0) return { ok: false, error: 'unknown' };
  const v = vals[row - 1];
  if (String(v[1]) !== r.hash) return { ok: false, error: 'password' };
  const admin = isAdmin(name);
  if (r.action === 'login') return { ok: true, data: JSON.parse(v[2] || '{}'), updated: Number(v[3]) || 0, admin, badge: !!v[5] };
  if (r.action === 'save') {
    sh.getRange(row, 3, 1, 3).setValues([[data, Number(r.updated) || Date.now(), new Date()]]);
    return { ok: true, updated: Number(r.updated) || Date.now() };
  }
  if (r.action === 'link_badge') {
    const b = String(r.badge || '');
    if (!b) return { ok: false, error: 'missing' };
    if (badgeUsed(b)) return { ok: false, error: 'badge_used' };
    sh.getRange(row, 6).setValue(b);
    return { ok: true };
  }
  // Signalements : il faut un compte de l'appli (nom + mot de passe) pour écrire.
  if (r.action === 'report') return addReport(r, name, admin);
  if (r.action === 'confirm' || r.action === 'resolve') return updateReport(r, name, admin);
  // Actions réservées aux administrateurs
  if (r.action === 'admin_accounts' || r.action === 'plan_set') {
    if (!admin) return { ok: false, error: 'forbidden' };
    if (r.action === 'admin_accounts') return adminAccounts(vals);
    return setPlan(r, name);
  }
  return { ok: false, error: 'action' };
}

// Tous les comptes, avec seulement les 45 derniers jours (pour rester léger sur le réseau).
function adminAccounts(vals) {
  const cut = Utilities.formatDate(new Date(Date.now() - 45 * 864e5), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const keep = o => { const x = {}; Object.keys(o || {}).forEach(k => { if (k >= cut) x[k] = o[k]; }); return x; };
  const list = [];
  for (let i = 1; i < vals.length; i++) {
    let d = {}; try { d = JSON.parse(vals[i][2] || '{}'); } catch (e) {}
    const q = d.quota || {};
    list.push({ name: String(vals[i][0]), updated: Number(vals[i][3]) || 0, badge: !!vals[i][5], admin: isAdmin(vals[i][0]),
      quota: { set: q.set || null, data: keep(q.data), hrs: keep(q.hrs), plan: keep(q.plan) } });
  }
  return { ok: true, accounts: list };
}

function ttlOf(r, admin) {
  const t = Number(r.ttl);
  if (admin && t === 0) return FOREVER; // permanent jusqu'à « C'est réglé »
  return Math.min(admin ? ADMIN_TTL : MAX_TTL, Math.max(15 * 60 * 1000, t || 8 * 3600 * 1000));
}

function addReport(r, name, admin) {
  const sh = reportSheet();
  const type = R_TYPES.indexOf(r.type) >= 0 ? r.type : 'autre';
  const now = Date.now();
  const id = String(r.id || ('s' + now.toString(36))).slice(0, 24);
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
  if (!ids.some(x => String(x[0]) === id)) {
    sh.appendRow([id, type, Math.round(Number(r.x)) || 0, Math.round(Number(r.y)) || 0, String(r.label || '').slice(0, 60),
      String(r.note || '').slice(0, 140), name, now, new Date(), now + ttlOf(r, admin), 'actif', 0, now, '', '']);
  }
  return listReports(0);
}

function updateReport(r, name, admin) {
  const sh = reportSheet();
  const vals = sh.getDataRange().getValues();
  const now = Date.now();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) !== String(r.id)) continue;
    if (r.action === 'confirm') {
      // « Toujours là » : on prolonge la durée de vie du signalement.
      sh.getRange(i + 1, 10, 1, 4).setValues([[Math.max(Number(vals[i][9]) || 0, now + ttlOf(r, admin)), 'actif', (Number(vals[i][11]) || 0) + 1, now]]);
    } else {
      sh.getRange(i + 1, 11, 1, 5).setValues([['regle', vals[i][11], now, name, new Date()]]);
    }
    break;
  }
  return listReports(0);
}

// Repère commun du plan (ajout, déplacement, renommage, suppression), visible par tout le monde.
function setPlan(r, name) {
  const sh = planSheet();
  const vals = sh.getDataRange().getValues();
  const now = Date.now();
  const id = String(r.id || ('c' + now.toString(36))).slice(0, 24);
  const rowv = [id, String(r.type || 'autre').slice(0, 20), String(r.pname || '').slice(0, 60), String(r.sub || '').slice(0, 80),
    Math.round(Number(r.x)) || 0, Math.round(Number(r.y)) || 0, r.deleted ? 1 : 0, now, name, new Date()];
  let found = false;
  for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === id) { sh.getRange(i + 1, 1, 1, rowv.length).setValues([rowv]); found = true; break; }
  if (!found) sh.appendRow(rowv);
  return listReports(0);
}

function listReports(since) {
  const now = Date.now();
  let ver = 0;
  const list = [];
  const rv = reportSheet().getDataRange().getValues();
  for (let i = 1; i < rv.length; i++) {
    const v = rv[i];
    const maj = Number(v[12]) || 0, exp = Number(v[9]) || 0;
    ver = Math.max(ver, maj, exp < now ? exp : 0);
    if (v[10] !== 'actif' || exp < now) continue;
    list.push({ id: String(v[0]), type: v[1], x: Number(v[2]), y: Number(v[3]), label: v[4], note: v[5], by: v[6], at: Number(v[7]), exp, conf: Number(v[11]) || 0 });
  }
  const plan = [];
  const pv = planSheet().getDataRange().getValues();
  for (let i = 1; i < pv.length; i++) {
    const v = pv[i];
    ver = Math.max(ver, Number(v[7]) || 0);
    plan.push({ id: String(v[0]), type: v[1], name: v[2], sub: v[3], x: Number(v[4]), y: Number(v[5]), deleted: !!Number(v[6]) });
  }
  if (since && since === ver) return { ok: true, same: true, ver };
  return { ok: true, ver, reports: list, plan };
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(SHEET); }
  if (sh.getLastRow() === 0) sh.appendRow(['nom', 'mdp_hash', 'donnees', 'maj', 'date_maj', 'badge']);
  else if (sh.getRange(1, 6).getValue() === '') sh.getRange(1, 6).setValue('badge');
  return sh;
}

function namedSheet(n, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(n);
  if (!sh) sh = ss.insertSheet(n);
  if (sh.getLastRow() === 0) { sh.appendRow(head); sh.setFrozenRows(1); }
  return sh;
}
function reportSheet() { return namedSheet(REPORTS, R_HEAD); }
function planSheet() { return namedSheet(PLAN, P_HEAD); }

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
