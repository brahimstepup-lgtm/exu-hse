// ================================================================
//  HSE TAGS — Code.gs  v10 FINAL (Profils User/Manager/Admin)
//  ✅ Auto-increment ID basé sur la colonne A de "plan"
//  ✅ Anti-blocage : return SUCCESS explicite
//  ✅ Statut "Ouvert" par défaut
//  ✅ Strict same-row photo mapping
//  ✅ Endpoints batch + lazy pour images
// ================================================================

var SHEET_ID         = '1UPfWAHzIKBVvECIPqjg3ccgIwBlX1KzyflJ73jfPRc0';
var SHEET_NAME       = 'plan';
var PHOTOS_FOLDER_ID = '1AE6QUkG0hLSmnyzKAVOqWadtjVzUEUQ7';

var C = {
  NUM:0, DATE_CR:1, DATE_CI:2, DANGER:3, GRAVITE:4,
  EMPLACE:5, ZONE:6, DESC:7, RISQUE:8, ACTION:9,
  PROPO:10, PHOTO_AV:11, STATUT:12, RESP:13, PHOTO_AP:14,
  DATE_FERME:15, AUTEUR:16, ANON:17, LAT:18, LNG:19, QR:20
};

var HEADER_ROW = 7;
var DATA_START = 8;

// ================================================================
//  ENTRY POINT (Web App + Image endpoint)
// ================================================================
function doGet(e) {
  if (e && e.parameter && e.parameter.img) {
    return serveImage_(e.parameter.img);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('HSE Tags 2025/2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ================================================================
//  ORDRE 3 + 4 + 5 : INSERT TAG WITH AUTO-INCREMENT
//  Garantit ID unique +1, statut "Ouvert", retour SUCCESS rapide
// ================================================================
function insertTagWithAutoIncrement(payload) {
  try {
    var sheet = getSheet_();
    if (!sheet) return 'ERROR: La feuille "plan" est introuvable.';

    // Ordre 3 : balayage de la colonne A pour le max ID
    var lastRow = sheet.getLastRow();
    var nextId = 1;
    if (lastRow >= DATA_START) {
      var vals = sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, 1).getValues();
      var maxId = 0;
      for (var i = 0; i < vals.length; i++) {
        var n = parseInt(vals[i][0], 10);
        if (!isNaN(n) && n > 0 && n < 99999 && String(vals[i][0]).trim().length <= 6) {
          if (n > maxId) maxId = n;
        }
      }
      nextId = maxId + 1;
    }

    var ref = 'Photo_' + String(nextId).padStart(3, '0');

    // Ordre 4 : appendRow avec ordre strict des colonnes
    sheet.appendRow([
      nextId,                                                            // A
      payload.dateCreation ? new Date(payload.dateCreation) : new Date(),// B
      payload.dateCible ? new Date(payload.dateCible) : '',              // C
      payload.dangerType || payload.danger || '',                        // D
      toGraviteEmoji_(payload.gravity || payload.gravite),               // E
      payload.emplacement || '',                                         // F
      payload.zone || '',                                                // G
      payload.description || '',                                         // H
      payload.risk || payload.risque || '',                              // I
      payload.actionUrgente || payload.action || '',                     // J
      payload.propositions || '',                                        // K
      ref,                                                                // L (sera remplacé par URL si photo)
      '🟥 Ouvert',                                                      // M — Ordre 5 défaut "Ouvert"
      payload.responsable || '',                                         // N
      ''                                                                  // O
    ]);

    var ir = sheet.getLastRow();

    // Sauvegarde des photos (avant/après) sur la même ligne
    if (payload.photoAvant && payload.photoAvant.length > 10) {
      var fnAv  = nextId + '.Référence Photo.' + nowHHMMSS_() + '.jpg';
      var urlAv = savePhotoToFolder_(payload.photoAvant, fnAv, payload.mime || 'image/jpeg');
      if (urlAv) sheet.getRange(ir, C.PHOTO_AV + 1).setValue(urlAv);
    }
    if (payload.photoApres && payload.photoApres.length > 10) {
      var fnAp  = nextId + '.Photo après.' + nowHHMMSS_() + '.jpg';
      var urlAp = savePhotoToFolder_(payload.photoApres, fnAp, payload.mime || 'image/jpeg');
      if (urlAp) sheet.getRange(ir, C.PHOTO_AP + 1).setValue(urlAp);
    }

    routeNotification_(payload, nextId);
    return { success: true, id: nextId, status: 'SUCCESS' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Alias historique (utilisé par la version Quick Add du HTML)
function addTag(d) { return insertTagWithAutoIncrement(d); }

// ================================================================
//  IMAGE ENDPOINTS
// ================================================================
function serveImage_(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    return ContentService.createTextOutput(
      'data:' + blob.getContentType() + ';base64,' +
      Utilities.base64Encode(blob.getBytes())
    ).setMimeType(ContentService.MimeType.TEXT);
  } catch(e) {
    return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
  }
}

function getImageBase64(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    return {
      success: true,
      data: 'data:' + blob.getContentType() + ';base64,' +
            Utilities.base64Encode(blob.getBytes())
    };
  } catch(e) { return { success: false, error: e.message }; }
}

function getImagesBatch(fileIds) {
  var results = {};
  for (var i = 0; i < fileIds.length; i++) {
    var id = fileIds[i];
    if (!id) continue;
    try {
      var file = DriveApp.getFileById(id);
      var blob = file.getBlob();
      results[id] = 'data:' + blob.getContentType() + ';base64,' +
                    Utilities.base64Encode(blob.getBytes());
    } catch(e) { results[id] = ''; }
  }
  return { success: true, data: results };
}

// ================================================================
//  GET SHEET
// ================================================================
function getSheet_() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    var all = ss.getSheets();
    for (var i = 0; i < all.length; i++) {
      try {
        var v = String(all[i].getRange(HEADER_ROW, 1).getValue()).toLowerCase();
        if (v.indexOf('cas') !== -1) { sheet = all[i]; break; }
      } catch(e) {}
    }
    if (!sheet) sheet = ss.getSheets()[0];
  }
  return sheet;
}

// ================================================================
//  EXTRACT FILE ID FROM ANY CELL CONTENT
// ================================================================
function extractFileId_(cellValue, formulaValue) {
  if (formulaValue) {
    var f = String(formulaValue);
    var m = f.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
    if (m) return m[1];
    m = f.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
    if (m) return m[1];
    m = f.match(/googleusercontent\.com\/d\/([a-zA-Z0-9_-]{20,})/);
    if (m) return m[1];
  }
  if (cellValue === null || cellValue === undefined) return '';
  var s = String(cellValue).trim();
  if (!s) return '';
  var m2 = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m2) return m2[1];
  m2 = s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m2) return m2[1];
  m2 = s.match(/googleusercontent\.com\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m2) return m2[1];
  if (/^[a-zA-Z0-9_-]{25,50}$/.test(s)) return s;
  return '';
}

// ================================================================
//  FOLDER CACHE (fallback Photo_NNN)
// ================================================================
var _folderCache = null;

function loadFolderCache_() {
  if (_folderCache) return;
  _folderCache = {};
  try {
    var folder = DriveApp.getFolderById(PHOTOS_FOLDER_ID);
    var files  = folder.getFiles();
    while (files.hasNext()) {
      var f      = files.next();
      var name   = f.getName();
      var fileId = f.getId();
      var m = name.match(/^(\d+)\.(Référence Photo|Reference Photo|Photo apr[èe]s|Photo avant)\b/i);
      if (m) {
        var id   = m[1];
        var type = m[2].toLowerCase();
        if (!_folderCache[id]) _folderCache[id] = { avantId:'', apresId:'' };
        if (type.indexOf('apr') !== -1) {
          if (!_folderCache[id].apresId) _folderCache[id].apresId = fileId;
        } else {
          if (!_folderCache[id].avantId) _folderCache[id].avantId = fileId;
        }
        continue;
      }
      var m2 = name.match(/^[Pp]hoto[_ ]?(\d+)/);
      if (m2) {
        var id2 = m2[1].replace(/^0+/, '') || '0';
        if (!_folderCache[id2]) _folderCache[id2] = { avantId:'', apresId:'' };
        if (!_folderCache[id2].avantId) _folderCache[id2].avantId = fileId;
      }
    }
  } catch(e) { Logger.log('loadFolderCache_: ' + e.message); }
}

function getFolderPhotoFor_(num) {
  if (!_folderCache) loadFolderCache_();
  return _folderCache[String(num)] || { avantId:'', apresId:'' };
}

// ================================================================
//  GET TAGS — same-row strict matching
// ================================================================
function getTags() {
  try {
    var sheet = getSheet_();
    var lr    = sheet.getLastRow();
    var lc    = Math.max(sheet.getLastColumn(), 15);
    if (lr < DATA_START) return { success: true, data: [] };

    var numRows = lr - DATA_START + 1;
    var range   = sheet.getRange(DATA_START, 1, numRows, Math.min(lc, 21));
    var vals     = range.getValues();
    var formulas = range.getFormulas();

    var richValuesAv = null, richValuesAp = null;
    try { richValuesAv = sheet.getRange(DATA_START, C.PHOTO_AV+1, numRows, 1).getRichTextValues(); } catch(e) {}
    try { if (lc > 14) richValuesAp = sheet.getRange(DATA_START, C.PHOTO_AP+1, numRows, 1).getRichTextValues(); } catch(e) {}

    loadFolderCache_();

    var tags = [];
    for (var i = 0; i < vals.length; i++) {
      var row    = vals[i];
      var rawNum = row[C.NUM];
      var num    = parseInt(rawNum, 10);
      if (isNaN(num) || num < 1 || num > 99999) continue;
      if (String(rawNum).trim().length > 6)     continue;

      var richAv = '', richAp = '';
      if (richValuesAv && richValuesAv[i] && richValuesAv[i][0]) {
        try {
          var runs = richValuesAv[i][0].getRuns();
          for (var rj = 0; rj < runs.length; rj++) {
            var lk = runs[rj].getLinkUrl();
            if (lk) { richAv = lk; break; }
          }
        } catch(e) {}
      }
      if (richValuesAp && richValuesAp[i] && richValuesAp[i][0]) {
        try {
          var runs2 = richValuesAp[i][0].getRuns();
          for (var rk = 0; rk < runs2.length; rk++) {
            var lk2 = runs2[rk].getLinkUrl();
            if (lk2) { richAp = lk2; break; }
          }
        } catch(e) {}
      }

      var avId = extractFileId_(row[C.PHOTO_AV], formulas[i][C.PHOTO_AV])
              || extractFileId_(richAv, '');
      var apId = extractFileId_(lc > 14 ? row[C.PHOTO_AP] : '', lc > 14 ? formulas[i][C.PHOTO_AP] : '')
              || extractFileId_(richAp, '');

      if (!avId || !apId) {
        var fp = getFolderPhotoFor_(num);
        if (!avId && fp.avantId) avId = fp.avantId;
        if (!apId && fp.apresId) apId = fp.apresId;
      }

      tags.push({
        id:           num,
        rowIndex:     DATA_START + i,
        dateCreation: fmtDate_(row[C.DATE_CR]),
        dateCible:    fmtDate_(row[C.DATE_CI]),
        dateFerme:    fmtDate_(lc > 15 ? (row[C.DATE_FERME]||'') : ''),
        danger:       noFormula_(row[C.DANGER]),
        gravite:      parseGravite_(row[C.GRAVITE]),
        emplacement:  txt_(row[C.EMPLACE]),
        zone:         txt_(row[C.ZONE]),
        description:  txt_(row[C.DESC]),
        risque:       txt_(row[C.RISQUE]),
        action:       noFormula_(row[C.ACTION]),
        propositions: noFormula_(row[C.PROPO]),
        photoAvantId: avId,
        photoApresId: apId,
        statut:       parseStatut_(row[C.STATUT]),
        responsable:  txt_(row[C.RESP]),
        auteur:       lc > 16 ? txt_(row[C.AUTEUR]||'') : '',
        anonymous:    lc > 17 ? String(row[C.ANON]||'').toLowerCase() === 'true' : false,
        lat:          lc > 18 ? (parseFloat(row[C.LAT])||0) : 0,
        lng:          lc > 19 ? (parseFloat(row[C.LNG])||0) : 0,
        qrZone:       lc > 20 ? txt_(row[C.QR]||'') : ''
      });
    }
    tags.sort(function(a, b) { return b.id - a.id; });
    return { success: true, data: tags };
  } catch(e) { return { success: false, error: e.message }; }
}

// ================================================================
//  GET ZONES + RESPONSABLES UNIQUES (pour filtres mobile)
// ================================================================
function getFilterOptions() {
  try {
    var res = getTags();
    if (!res.success) return res;
    var zones = {}, resps = {};
    res.data.forEach(function(t) {
      if (t.zone)        zones[t.zone] = true;
      if (t.responsable) resps[t.responsable] = true;
    });
    return {
      success: true,
      data: {
        zones:        Object.keys(zones).sort(),
        responsables: Object.keys(resps).sort()
      }
    };
  } catch(e) { return { success: false, error: e.message }; }
}

// ================================================================
//  KPI
// ================================================================
function getKPIs() {
  try {
    var res = getTags();
    if (!res.success) return res;
    var tags  = res.data;
    var today = new Date();
    var total = tags.length, ouverts = 0, fermes = 0, overdue = 0;
    var byZone={}, byDanger={}, byMonth={}, byGravite={Élevée:0,Moyenne:0,Faible:0};
    var durClosed = [], durOpen = [];

    tags.forEach(function(t) {
      if (t.statut === 'Fermé') fermes++;
      else {
        ouverts++;
        if (t.dateCible) {
          var dc = new Date(t.dateCible);
          if (!isNaN(dc.getTime()) && dc < today) overdue++;
        }
      }
      if (t.dateCreation) {
        var d1 = new Date(t.dateCreation);
        if (!isNaN(d1.getTime())) {
          if (t.statut === 'Fermé') {
            var d2 = t.dateFerme ? new Date(t.dateFerme) : today;
            var days = Math.max(0, Math.round((d2 - d1) / 86400000));
            durClosed.push({ id:t.id, danger:t.danger, zone:t.zone, days:days });
          } else {
            var daysSince = Math.round((today - d1) / 86400000);
            durOpen.push({ id:t.id, danger:t.danger, zone:t.zone, days:daysSince });
          }
        }
      }
      var z = t.zone || 'Inconnu';
      byZone[z] = (byZone[z]||0) + 1;
      var d = (t.danger||'Inconnu').substring(0, 32);
      byDanger[d] = (byDanger[d]||0) + 1;
      if (t.dateCreation) {
        var m = t.dateCreation.substring(0, 7);
        byMonth[m] = (byMonth[m]||0) + 1;
      }
      if (byGravite[t.gravite] !== undefined) byGravite[t.gravite]++;
    });

    var avgDays=0, maxDays=0, minDays=0, sla7=0;
    if (durClosed.length) {
      var sum = durClosed.reduce(function(s,d){return s+d.days;}, 0);
      avgDays = Math.round(sum / durClosed.length * 10) / 10;
      maxDays = Math.max.apply(null, durClosed.map(function(d){return d.days;}));
      minDays = Math.min.apply(null, durClosed.map(function(d){return d.days;}));
      sla7    = Math.round(durClosed.filter(function(d){return d.days<=7;}).length / durClosed.length * 100);
    }

    return { success: true, data: {
      total, ouverts, fermes, overdue, avgDays, maxDays, minDays, sla7,
      byZone, byDanger, byMonth, byGravite,
      topSlow: durClosed.sort(function(a,b){return b.days-a.days;}).slice(0,5),
      topOpen: durOpen.sort(function(a,b){return b.days-a.days;}).slice(0,5)
    }};
  } catch(e) { return { success:false, error:e.message }; }
}

// ================================================================
//  UPDATE / DELETE
// ================================================================
function updateTag(d) {
  try {
    var sheet = getSheet_();
    var ri    = parseInt(d.rowIndex, 10);
    if (!ri || ri < DATA_START) return { success:false, error:'rowIndex invalide' };

    var sv = function(col, val) { sheet.getRange(ri, col + 1).setValue(val); };
    var oldStatut = parseStatut_(sheet.getRange(ri, C.STATUT + 1).getValue());
    if (oldStatut === 'Ouvert' && d.statut === 'Fermé') {
      sv(C.DATE_FERME, new Date());
    }

    sv(C.DATE_CI,   d.dateCible ? new Date(d.dateCible) : '');
    sv(C.DANGER,    d.danger || '');
    sv(C.GRAVITE,   toGraviteEmoji_(d.gravite));
    sv(C.EMPLACE,   d.emplacement || '');
    sv(C.ZONE,      d.zone || '');
    sv(C.DESC,      d.description || '');
    sv(C.RISQUE,    d.risque || '');
    sv(C.ACTION,    d.action || '');
    sv(C.PROPO,     d.propositions || '');
    sv(C.STATUT,    toStatutEmoji_(d.statut));
    sv(C.RESP,      d.responsable || '');

    var tagId = parseInt(d.id, 10) || (ri - (HEADER_ROW - 1));
    if (d.photoAvant && d.photoAvant.length > 10) {
      var fnAv  = tagId + '.Référence Photo.' + nowHHMMSS_() + '.jpg';
      var urlAv = savePhotoToFolder_(d.photoAvant, fnAv, d.mime||'image/jpeg');
      if (urlAv) sv(C.PHOTO_AV, urlAv);
    }
    if (d.photoApres && d.photoApres.length > 10) {
      var fnAp  = tagId + '.Photo après.' + nowHHMMSS_() + '.jpg';
      var urlAp = savePhotoToFolder_(d.photoApres, fnAp, d.mime||'image/jpeg');
      if (urlAp) sv(C.PHOTO_AP, urlAp);
    }
    return { success: true, status: 'SUCCESS' };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteTag(rowIndex) {
  try {
    getSheet_().deleteRow(parseInt(rowIndex, 10));
    return { success: true, status: 'SUCCESS' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ================================================================
//  SAVE AFTER-PHOTO ONLY (sans réécrire les autres colonnes)
//  Permet d'ajouter la "Photo après" directement depuis le détail
// ================================================================
function saveAfterPhoto(p) {
  try {
    var sheet = getSheet_();
    var ri = parseInt(p.rowIndex, 10);
    if (!ri || ri < DATA_START) return { success:false, error:'rowIndex invalide' };
    if (!p.photoApres || p.photoApres.length < 10) return { success:false, error:'Photo manquante' };
    var tagId = parseInt(p.id, 10) || ri;
    var fn  = tagId + '.Photo après.' + nowHHMMSS_() + '.jpg';
    var url = savePhotoToFolder_(p.photoApres, fn, p.mime || 'image/jpeg');
    if (!url) return { success:false, error:'Échec de la sauvegarde' };
    sheet.getRange(ri, C.PHOTO_AP + 1).setValue(url);
    return { success:true, status:'SUCCESS', url:url };
  } catch(e) { return { success:false, error:e.message }; }
}

// ================================================================
//  ANALYSE IA DE LA PHOTO (Gemini Vision) → description suggérée
//  Nécessite une clé API dans Propriétés du script : GEMINI_API_KEY
// ================================================================
function analyzePhoto(p) {
  try {
    if (!p || !p.photo || p.photo.length < 10) return { success:false, error:'Photo manquante' };
    var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) return { success:false, error:'Clé GEMINI_API_KEY non configurée (Propriétés du script)' };

    var clean = String(p.photo).replace(/^data:[^;]+;base64,/, '');
    var model = 'gemini-2.0-flash';
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model +
              ':generateContent?key=' + encodeURIComponent(key);

    var prompt = "Tu es un inspecteur HSE (Hygiène Sécurité Environnement) en milieu industriel. " +
      "Décris en UNE seule phrase concise et factuelle, en français, l'anomalie ou le danger de sécurité " +
      "visible sur la photo (équipement concerné, défaut observé, risque). " +
      "Pas de préambule ni de liste : uniquement la description.";

    var payload = {
      contents: [{ parts: [
        { text: prompt },
        { inline_data: { mime_type: (p.mime || 'image/jpeg'), data: clean } }
      ]}],
      generationConfig: { temperature: 0.4, maxOutputTokens: 150 }
    };

    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    var data = JSON.parse(res.getContentText() || '{}');
    if (code !== 200) {
      return { success:false, error:'Gemini ' + code + ': ' + ((data.error && data.error.message) || '') };
    }
    var text = '';
    try { text = data.candidates[0].content.parts[0].text.trim(); } catch(e) {}
    if (!text) return { success:false, error:'Réponse vide de Gemini' };
    return { success:true, text:text };
  } catch(e) { return { success:false, error:e.toString() }; }
}

// ================================================================
//  HELPERS
// ================================================================
function nowHHMMSS_() {
  var n = new Date();
  return pad2_(n.getHours()) + pad2_(n.getMinutes()) + pad2_(n.getSeconds());
}
function pad2_(n) { return String(n).padStart(2, '0'); }

function savePhotoToFolder_(b64, name, mime) {
  try {
    var clean = b64.replace(/^data:[^;]+;base64,/, '');
    var blob  = Utilities.newBlob(Utilities.base64Decode(clean), mime, name);
    var folder= DriveApp.getFolderById(PHOTOS_FOLDER_ID);
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    _folderCache = null;
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch(e) { Logger.log('savePhotoToFolder_: ' + e.message); return ''; }
}

function routeNotification_(d, id) {
  try {
    var danger = (d.danger||d.dangerType||'').toLowerCase();
    var dept = /[eé]lectr/i.test(danger) ? 'Maintenance Électrique' :
               /fuite|pompe|mécan|hydraul/i.test(danger) ? 'Maintenance Mécanique' :
               /incendie|explos|gaz|feu/i.test(danger) ? 'Sécurité/HSE' :
               /chimiq|corros|dévers/i.test(danger) ? 'HSE/Environnement' : 'HSE';
    Logger.log('[HSE #' + id + '] → ' + dept);
  } catch(e) {}
}

function fmtDate_(v) {
  if (!v) return '';
  try {
    var d = (v instanceof Date) ? v : new Date(v);
    if (isNaN(d.getTime())) return String(v).replace(/\s+0:00:?0?$/, '').trim();
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch(e) { return String(v).replace(/\s+0:00:?0?$/, '').trim(); }
}
function txt_(v) { return (v==null||v===undefined)?'':String(v).trim(); }
function noFormula_(v) { var s=txt_(v); return (s.startsWith('=')||s.startsWith('\\='))?'':s; }
function parseGravite_(v) {
  var s=txt_(v);
  if(/[eé]lev[eé]/i.test(s)||s.indexOf('🟥')!==-1) return 'Élevée';
  if(/moyen/i.test(s)||s.indexOf('🟧')!==-1) return 'Moyenne';
  if(/faib/i.test(s)||s.indexOf('🟢')!==-1) return 'Faible';
  return 'Élevée';
}
function parseStatut_(v) {
  var s=txt_(v);
  return (/ferm/i.test(s)||s.indexOf('✅')!==-1)?'Fermé':'Ouvert';
}
function toGraviteEmoji_(g) {
  if(g==='Moyenne'||g==='🟧 Moyenne') return '🟧 Moyenne';
  if(g==='Faible'||g==='🟢 Faible')   return '🟢 Faible';
  return '🟥 Élevée';
}
function toStatutEmoji_(s) {
  return (s==='Fermé')?'✅ Fermé':'🟥 Ouvert';
}

// ================================================================
//  DEBUG
// ================================================================
function debugPhotoMapping() {
  var res = getTags();
  if (!res.success) { Logger.log('Erreur: ' + res.error); return; }
  Logger.log('=== Diagnostic photos (premiers 20 tags) ===');
  res.data.slice(0, 20).forEach(function(t) {
    Logger.log('#' + t.id + ' (ligne ' + t.rowIndex + ') | ' +
      'avantId=' + (t.photoAvantId || '∅') + ' | ' +
      'apresId=' + (t.photoApresId || '∅'));
  });
}
