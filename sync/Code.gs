// Synchro de l'appli Chantiers
// 1. Remplace CHANGE-MOI par un mot de passe à toi (le même que dans l'appli, Réglages).
// 2. Déployer > Nouveau déploiement > Application Web, Exécuter en tant que : Moi, Accès : Tout le monde.
var KEY = 'CHANGE-MOI';
var VERSION = 2;
var CAL_NAME = 'Chantiers';

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function props_() { return PropertiesService.getScriptProperties(); }

function dataFile_() {
  var p = props_(), id = p.getProperty('dataFile');
  if (id) { try { return DriveApp.getFileById(id); } catch (e) {} }
  var f = DriveApp.createFile('chantiers-donnees.json', '{}', MimeType.PLAIN_TEXT);
  p.setProperty('dataFile', f.getId());
  return f;
}
function photoFolder_() {
  var p = props_(), id = p.getProperty('photoFolder');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('Chantiers photos');
  p.setProperty('photoFolder', f.getId());
  return f;
}
function rev_() { return Number(props_().getProperty('rev') || 0); }

function doGet(e) {
  var q = (e && e.parameter) || {};
  if (q.key !== KEY) return out_({ ok: false, error: 'cle' });
  if (q.action === 'pull') {
    return out_({ ok: true, v: VERSION, rev: rev_(), data: dataFile_().getBlob().getDataAsString() });
  }
  if (q.action === 'photo') {
    var it = photoFolder_().getFilesByName(q.id + '.jpg');
    if (!it.hasNext()) return out_({ ok: false, error: 'absente' });
    var b = it.next().getBlob();
    return out_({ ok: true, data: 'data:image/jpeg;base64,' + Utilities.base64Encode(b.getBytes()) });
  }
  return out_({ ok: true, v: VERSION, pong: true, rev: rev_() });
}

function doPost(e) {
  var b = JSON.parse(e.postData.contents);
  if (b.key !== KEY) return out_({ ok: false, error: 'cle' });
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (b.action === 'push') {
      if (String(b.rev) !== String(rev_())) return out_({ ok: false, conflict: true, rev: rev_() });
      dataFile_().setContent(b.data);
      var r = rev_() + 1;
      props_().setProperty('rev', String(r));
      if (b.summary) { try { writeSheets_(b.summary); } catch (err) {} }
      var ag = '';
      if (b.agenda) { try { syncAgenda_(b.agenda); ag = 'ok'; } catch (err) { ag = String(err); } }
      return out_({ ok: true, v: VERSION, rev: r, agenda: ag });
    }
    if (b.action === 'photo') {
      var fo = photoFolder_();
      if (!fo.getFilesByName(b.id + '.jpg').hasNext()) {
        var bytes = Utilities.base64Decode(String(b.data).split(',')[1]);
        fo.createFile(Utilities.newBlob(bytes, 'image/jpeg', b.id + '.jpg'));
      }
      return out_({ ok: true });
    }
  } finally {
    lock.releaseLock();
  }
  return out_({ ok: false, error: 'action' });
}

// Copie lisible des données dans les onglets du Google Sheet
function writeSheets_(sum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  Object.keys(sum).forEach(function (name) {
    var rows = sum[name];
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clearContents();
    if (rows.length && rows[0].length) {
      sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
}

// À exécuter une fois depuis l'éditeur pour autoriser l'accès à Google Agenda
function autoriserAgenda() {
  var cal = calendar_();
  Logger.log('Agenda prêt : ' + cal.getName());
}

// Agenda Google "Chantiers" : rendez vous et rappels envoyés par l'appli
function calendar_() {
  var l = CalendarApp.getCalendarsByName(CAL_NAME);
  if (l.length) return l[0];
  return CalendarApp.createCalendar(CAL_NAME, { color: CalendarApp.Color.ORANGE });
}
function parseDay_(s) { var p = String(s).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
function syncAgenda_(items) {
  var cal = calendar_();
  var from = new Date(); from.setDate(from.getDate() - 3); from.setHours(0, 0, 0, 0);
  var to = new Date(); to.setFullYear(to.getFullYear() + 1);
  var existing = {};
  cal.getEvents(from, to).forEach(function (ev) {
    var k = ev.getTag('ck');
    if (k) { if (existing[k]) ev.deleteEvent(); else existing[k] = ev; }
  });
  var wanted = {};
  items.forEach(function (it) {
    wanted[it.k] = 1;
    var sig = JSON.stringify([it.t, it.d, it.f || '', it.h || '', it.n || '']);
    var ev = existing[it.k];
    if (ev && ev.getTag('cs') === sig) return;
    if (ev) ev.deleteEvent();
    var start = parseDay_(it.d), ne;
    if (it.h) {
      var p = String(it.h).split(':');
      start.setHours(+p[0], +p[1] || 0, 0, 0);
      ne = cal.createEvent(it.t, start, new Date(start.getTime() + 3600000), { description: it.n || '' });
      ne.removeAllReminders();
      ne.addPopupReminder(60);
      ne.addPopupReminder(1440);
    } else {
      var end = parseDay_(it.f || it.d); end.setDate(end.getDate() + 1);
      ne = cal.createAllDayEvent(it.t, start, end, { description: it.n || '' });
      ne.removeAllReminders();
      ne.addPopupReminder(360);   // la veille à 18 h
    }
    ne.setTag('ck', it.k);
    ne.setTag('cs', sig);
  });
  Object.keys(existing).forEach(function (k) { if (!wanted[k]) existing[k].deleteEvent(); });
}
