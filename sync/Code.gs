// Synchro de l'appli Chantiers
// 1. Remplace CHANGE-MOI par un mot de passe à toi (le même que dans l'appli, Réglages).
// 2. Déployer > Nouveau déploiement > Application Web, Exécuter en tant que : Moi, Accès : Tout le monde.
var KEY = 'CHANGE-MOI';

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
    return out_({ ok: true, rev: rev_(), data: dataFile_().getBlob().getDataAsString() });
  }
  if (q.action === 'photo') {
    var it = photoFolder_().getFilesByName(q.id + '.jpg');
    if (!it.hasNext()) return out_({ ok: false, error: 'absente' });
    var b = it.next().getBlob();
    return out_({ ok: true, data: 'data:image/jpeg;base64,' + Utilities.base64Encode(b.getBytes()) });
  }
  return out_({ ok: true, pong: true, rev: rev_() });
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
      return out_({ ok: true, rev: r });
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
