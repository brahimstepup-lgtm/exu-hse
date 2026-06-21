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
  DATE_FERME:15, AUTEUR:16, ANON:17, LAT:18, LNG:19, QR:20,
  EMAIL_SENT:21
};

var HEADER_ROW = 7;
var DATA_START = 8;

// ================================================================
//  AUTHENTIFICATION / RÔLES
//  Mot de passe admin stocké dans Propriétés du script : ADMIN_PASSWORD
//  Les utilisateurs ordinaires (anonymes) peuvent uniquement :
//    - ajouter un tag      (addTag)
//    - ajouter une photo   (saveAfterPhoto)
//  Toute action de modification/suppression/dashboard exige l'admin.
// ================================================================
function login(password) {
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!stored) return { success:false, error:'ADMIN_PASSWORD non configuré dans les Propriétés du script' };
  if (String(password) === String(stored)) return { success:true, role:'admin' };
  return { success:false, error:'Mot de passe incorrect' };
}
function isAdmin_(key) {
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return !!stored && key != null && String(key) === String(stored);
}

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
    var dateCreation = payload.dateCreation ? new Date(payload.dateCreation) : new Date();
    var dateCible;
    if (payload.dateCible) {
      dateCible = new Date(payload.dateCible);
    } else {
      dateCible = new Date(dateCreation.getTime());
      dateCible.setDate(dateCible.getDate() + 3);
    }
    sheet.appendRow([
      nextId,           // A
      dateCreation,     // B
      dateCible,        // C — défaut J+3 si non fourni
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
      if (urlAp) {
        sheet.getRange(ir, C.PHOTO_AP + 1).setValue(urlAp);
        sheet.getRange(ir, C.STATUT + 1).setValue(toStatutEmoji_('Fermé'));
        sheet.getRange(ir, C.DATE_FERME + 1).setValue(new Date());
      }
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
    var range   = sheet.getRange(DATA_START, 1, numRows, Math.min(lc, 22));
    var vals     = range.getValues();
    var formulas = range.getFormulas();

    var richValuesAv = null, richValuesAp = null;
    try { richValuesAv = sheet.getRange(DATA_START, C.PHOTO_AV+1, numRows, 1).getRichTextValues(); } catch(e) {}
    try { if (lc > 14) richValuesAp = sheet.getRange(DATA_START, C.PHOTO_AP+1, numRows, 1).getRichTextValues(); } catch(e) {}

    loadFolderCache_();

    var tags = [];
    var cellRepairs = []; // collect {row, col, url} to write back missing links
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

      // Fallback Drive uniquement pour la photo avant.
      // La photo après ne doit jamais être déduite du dossier :
      // c'est une action intentionnelle (fermeture du tag).
      if (!avId) {
        var fp = getFolderPhotoFor_(num);
        if (fp.avantId) {
          avId = fp.avantId;
          cellRepairs.push({ rowIndex: DATA_START + i, col: C.PHOTO_AV + 1, fileId: avId });
        }
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
        qrZone:       lc > 20 ? txt_(row[C.QR]||'') : '',
        emailSent:    lc > 21 ? txt_(row[C.EMAIL_SENT]||'') : ''
      });
    }
    // Write Drive URLs back to sheet for any cells that were missing
    if (cellRepairs.length > 0) {
      try {
        for (var ci = 0; ci < cellRepairs.length; ci++) {
          var rep = cellRepairs[ci];
          var url = 'https://drive.google.com/uc?export=view&id=' + rep.fileId;
          sheet.getRange(rep.rowIndex, rep.col).setValue(url);
        }
      } catch(repErr) { Logger.log('cellRepair error: ' + repErr.message); }
    }

    tags.sort(function(a, b) { return b.id - a.id; });
    return { success: true, data: tags };
  } catch(e) { return { success: false, error: e.message }; }
}

// ================================================================
//  GET ZONES + RESPONSABLES UNIQUES (pour filtres mobile)
// ================================================================
// Liste officielle des responsables (référence pour dédoublonnage)
var RESP_LIST = [
  'Omar Ourihan', 'Amir Mahmoud', 'Kouachi Brahim', 'Chekalil Brahim',
  'Salah Haloui', 'Kerrad Nazim', 'Rabah Seba', 'Hadoune Youcef',
  'Media Amine', 'Mohamed Zerar'
];

// Clé normalisée : minuscules, sans accents, espaces réduits
function normRespKey_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

// Renvoie le nom officiel si correspondance, sinon le nom nettoyé
function canonResp_(s, lookup) {
  var k = normRespKey_(s);
  if (lookup[k]) return lookup[k];
  return String(s || '').replace(/\s+/g, ' ').trim();
}

// ── Annuaire des responsables (noms + emails) ──────────────────
// Modifiable par l'admin, stocké dans la propriété RESP_DIRECTORY (JSON).
// À défaut, construit depuis RESP_LIST + RESP_EMAILS (valeurs par défaut).
function getRespDirectory_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('RESP_DIRECTORY');
    if (raw) {
      var arr = JSON.parse(raw);
      if (arr && arr.length) {
        return arr.map(function(r){ return { name:String(r.name||'').trim(), email:String(r.email||'').trim() }; })
                  .filter(function(r){ return r.name; });
      }
    }
  } catch(e) {}
  return RESP_LIST.map(function(n){ return { name:n, email:(RESP_EMAILS[n]||'') }; });
}

// Public : noms seulement (pour les menus déroulants de l'app)
function getResponsablesList() {
  return { success:true, data: getRespDirectory_().map(function(r){ return r.name; }) };
}

// Admin : noms + emails (pour l'édition)
function getResponsables(adminKey) {
  if (!isAdmin_(adminKey)) return { success:false, error:'Accès réservé à l\'administrateur' };
  return { success:true, data: getRespDirectory_() };
}

// Admin : enregistre l'annuaire (noms + emails)
function saveResponsables(adminKey, list) {
  if (!isAdmin_(adminKey)) return { success:false, error:'Accès réservé à l\'administrateur' };
  if (!Array.isArray(list)) return { success:false, error:'Liste invalide' };
  var clean = list.map(function(r){ return { name:String(r.name||'').trim(), email:String(r.email||'').trim() }; })
                  .filter(function(r){ return r.name; });
  if (!clean.length) return { success:false, error:'Au moins un responsable requis' };
  PropertiesService.getScriptProperties().setProperty('RESP_DIRECTORY', JSON.stringify(clean));
  return { success:true, data:clean };
}

function getFilterOptions() {
  try {
    var res = getTags();
    if (!res.success) return res;
    // Table de correspondance clé normalisée -> nom officiel
    var lookup = {};
    getRespDirectory_().forEach(function(r){ lookup[normRespKey_(r.name)] = r.name; });

    var zones = {}, resps = {};
    res.data.forEach(function(t) {
      if (t.zone) zones[String(t.zone).replace(/\s+/g, ' ').trim()] = true;
      if (t.responsable) {
        var name = canonResp_(t.responsable, lookup);
        if (name) resps[name] = true;
      }
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
function getKPIs(adminKey) {
  if (!isAdmin_(adminKey)) return { success:false, error:'Accès réservé à l\'administrateur' };
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
  if (!isAdmin_(d && d.adminKey)) return { success:false, error:'Accès réservé à l\'administrateur' };
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

function deleteTag(rowIndex, adminKey) {
  if (!isAdmin_(adminKey)) return { success:false, error:'Accès réservé à l\'administrateur' };
  try {
    getSheet_().deleteRow(parseInt(rowIndex, 10));
    return { success: true, status: 'SUCCESS' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ================================================================
//  SUPPRESSION D'UNE PHOTO (avant / après) — Admin uniquement
//  Vide la cellule correspondante et met le fichier Drive à la corbeille.
// ================================================================
function deletePhoto(p) {
  if (!isAdmin_(p && p.adminKey)) return { success:false, error:'Accès réservé à l\'administrateur' };
  try {
    var sheet = getSheet_();
    var ri = parseInt(p.rowIndex, 10);
    if (!ri || ri < DATA_START) return { success:false, error:'rowIndex invalide' };
    var which = String(p.which || '').toLowerCase();
    var col = which === 'apres' ? C.PHOTO_AP : which === 'avant' ? C.PHOTO_AV : -1;
    if (col < 0) return { success:false, error:'Type de photo invalide (avant/apres)' };

    // Récupère l'ID de la ligne pour retrouver le tag
    var tagId = String(sheet.getRange(ri, C.NUM + 1).getValue() || '').trim();

    // Supprime le fichier Drive depuis la cellule (best-effort)
    try {
      var url = String(sheet.getRange(ri, col + 1).getValue() || '').trim();
      var fileId = extractFileId_(url, '');
      if (fileId) DriveApp.getFileById(fileId).setTrashed(true);
    } catch(e) {}

    // Supprime AUSSI tous les fichiers correspondants dans le dossier Drive
    // pour éviter qu'ils soient retrouvés par le cache dossier
    if (tagId) {
      try {
        var isApres = (which === 'apres');
        var folder = DriveApp.getFolderById(PHOTOS_FOLDER_ID);
        var files = folder.getFiles();
        while (files.hasNext()) {
          var f = files.next();
          var name = f.getName();
          var m = name.match(/^(\d+)\.(Référence Photo|Reference Photo|Photo apr[èe]s|Photo avant)\b/i);
          if (m && String(parseInt(m[1], 10)) === String(parseInt(tagId, 10))) {
            var isApresFile = /apr/i.test(m[2]);
            if (isApres === isApresFile) {
              try { f.setTrashed(true); } catch(fe) {}
            }
          }
        }
      } catch(e) {}
    }

    sheet.getRange(ri, col + 1).setValue('');
    return { success:true, status:'SUCCESS', which:which };
  } catch(e) { return { success:false, error:e.toString() }; }
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

    var oldStatut = parseStatut_(sheet.getRange(ri, C.STATUT + 1).getValue());
    if (oldStatut === 'Ouvert') {
      sheet.getRange(ri, C.STATUT + 1).setValue(toStatutEmoji_('Fermé'));
      sheet.getRange(ri, C.DATE_FERME + 1).setValue(new Date());
    }

    return { success:true, status:'SUCCESS', url:url, statut:'Fermé' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ================================================================
//  TEST D'AUTORISATION MAIL — à exécuter UNE FOIS depuis l'éditeur
//  (sélectionner "testMailAuth" puis ▶ Exécuter) pour déclencher
//  l'écran d'autorisation de script.send_mail. Peut être supprimée
//  après l'autorisation accordée.
// ================================================================
function testMailAuth() {
  MailApp.sendEmail('brahimstepup@gmail.com', 'Test HSE Tags', 'Autorisation Mail OK.');}

// ================================================================
//  EMAIL AU RESPONSABLE (bouton manuel dans la vue détail)
//  Correspondance noms → emails. Complétez les adresses ci-dessous.
//  Vous pouvez aussi ajouter/surcharger via Propriétés du script :
//    clé RESP_EMAILS = {"Kouachi Brahim":"BrahimKOUACHI@palmaryfood.com", ...}
// ================================================================
var RESP_EMAILS = {
  'Omar Ourihan':    '',
  'Amir Mahmoud':    '',
  'Kouachi Brahim':  'BrahimKOUACHI@palmaryfood.com',
  'Chekalil Brahim': '',
  'Salah Haloui':    '',
  'Kerrad Nazim':    '',
  'Rabah Seba':      '',
  'Hadoune Youcef':  '',
  'Media Amine':     '',
  'Mohamed Zerar':   ''
};

function respEmailFor_(name) {
  var k = normRespKey_(name);
  // 1) Annuaire (modifiable par l'admin) — source principale
  var dir = getRespDirectory_();
  for (var i = 0; i < dir.length; i++) {
    if (normRespKey_(dir[i].name) === k && dir[i].email) return dir[i].email;
  }
  // 2) Surcharge héritée via propriété RESP_EMAILS (JSON)
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('RESP_EMAILS');
    if (raw) {
      var extra = JSON.parse(raw);
      for (var key in extra) {
        if (normRespKey_(key) === k && extra[key]) return String(extra[key]).trim();
      }
    }
  } catch(e) {}
  return '';
}

function sendTagEmail(p) {
  if (!isAdmin_(p && p.adminKey)) return { success:false, error:'Accès réservé à l\'administrateur' };
  try {
    var sheet = getSheet_();
    var ri = parseInt(p.rowIndex, 10);
    if (!ri || ri < DATA_START) return { success:false, error:'rowIndex invalide' };

    var row = sheet.getRange(ri, 1, 1, 21).getValues()[0];

    // Résoudre les destinataires principaux (toList) envoyés depuis le client
    var toEmails = [];
    if (p.toList && p.toList.length) {
      for (var ti = 0; ti < p.toList.length; ti++) {
        var tEntry = String(p.toList[ti] || '').trim();
        if (!tEntry) continue;
        var tEm = tEntry.indexOf('@') >= 0 ? tEntry : respEmailFor_(tEntry);
        if (tEm && toEmails.indexOf(tEm) < 0) toEmails.push(tEm);
      }
    }
    // Fallback : responsable du tag
    if (!toEmails.length) {
      var resp = txt_(row[C.RESP]);
      if (!resp) return { success:false, error:'Aucun responsable assigné à ce tag' };
      var fallbackTo = respEmailFor_(resp);
      if (!fallbackTo) return { success:false, error:'Email non configuré pour « ' + resp + ' ».' };
      toEmails.push(fallbackTo);
    }
    var to = toEmails[0];
    var extraTo = toEmails.slice(1); // destinataires supplémentaires → ajoutés en CC

    // Destinataires en copie (CC)
    var ccEmails = [].concat(extraTo);
    if (p.cc && p.cc.length) {
      for (var ci = 0; ci < p.cc.length; ci++) {
        var c = String(p.cc[ci] || '').trim();
        if (!c) continue;
        var em = c.indexOf('@') >= 0 ? c : respEmailFor_(c);
        if (em && em.toLowerCase() !== to.toLowerCase() && ccEmails.indexOf(em) < 0) ccEmails.push(em);
      }
    }

    var id       = txt_(row[C.NUM]) || p.id || ri;
    var danger   = noFormula_(row[C.DANGER]);
    var gravite  = parseGravite_(row[C.GRAVITE]);
    var statut   = parseStatut_(row[C.STATUT]);
    var zone     = noFormula_(row[C.ZONE]);
    var emplace  = noFormula_(row[C.EMPLACE]);
    var desc     = noFormula_(row[C.DESC]);
    var risque   = noFormula_(row[C.RISQUE]);
    var action   = noFormula_(row[C.ACTION]);
    var propo    = noFormula_(row[C.PROPO]);
    var dateCr   = fmtDate_(row[C.DATE_CR]);
    var dateCi   = fmtDate_(row[C.DATE_CI]);

    var subject = '[HSE] Tag #' + id + ' qui vous est assigné — ' + (danger || 'Anomalie') + ' (' + gravite + ')';

    var line = function(k, v) {
      if (!v) return '';
      return '<tr><td style="padding:6px 12px;font-weight:600;color:#555;white-space:nowrap;vertical-align:top">' + k +
             '</td><td style="padding:6px 12px;color:#111">' + String(v).replace(/\n/g, '<br>') + '</td></tr>';
    };
    var gravColor = gravite === 'Élevée' ? '#e74c3c' : gravite === 'Moyenne' ? '#e67e22' : '#27ae60';

    // Récupère la photo avant depuis Drive (best-effort)
    var photoBlob = null;
    try {
      var avUrl = String(row[C.PHOTO_AV] || '').trim();
      var avId  = extractFileId_(avUrl, '');
      if (!avId) {
        var fp = getFolderPhotoFor_(parseInt(id, 10));
        avId = fp.avantId || '';
      }
      if (avId) photoBlob = DriveApp.getFileById(avId).getBlob().setName('photo_avant.jpg');
    } catch(e) {}

    var photoHtml = photoBlob
      ? '<div style="padding:0 20px 16px"><img src="cid:photoAvant" style="max-width:100%;border-radius:6px;border:1px solid #eee"></div>'
      : '';

    var html =
      '<div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:auto;border:1px solid #eee;border-radius:8px;overflow:hidden">' +
        '<div style="background:#1f2937;color:#fff;padding:16px 20px">' +
          '<div style="font-size:18px;font-weight:700">⚠️ Tag HSE #' + id + '</div>' +
          '<div style="font-size:13px;opacity:.85;margin-top:2px">Hygiène · Sécurité · Environnement</div>' +
        '</div>' +
        '<div style="padding:16px 20px;color:#111">' +
          '<p style="margin:0 0 12px">Bonjour <b>' + resp + '</b>,</p>' +
          '<p style="margin:0 0 16px">Un tag HSE vous est assigné. Voici les détails :</p>' +
          '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
            line('N° Cas', '#' + id) +
            line('Type de danger', danger) +
            '<tr><td style="padding:6px 12px;font-weight:600;color:#555">Gravité</td>' +
              '<td style="padding:6px 12px"><span style="background:' + gravColor + ';color:#fff;padding:2px 10px;border-radius:999px;font-size:12px;font-weight:700">' + gravite + '</span></td></tr>' +
            line('Statut', statut) +
            line('Zone', zone) +
            line('Emplacement', emplace) +
            line('Description', desc) +
            line('Risque principal', risque) +
            line('Action', action) +
            line('Propositions', propo) +
            line('Date création', dateCr) +
            line('Date cible', dateCi) +
          '</table>' +
          '<p style="margin:18px 0 0;font-size:13px;color:#666">Merci de prendre en charge ce signalement.</p>' +
        '</div>' +
        photoHtml +
        '<div style="background:#f7f7f7;padding:12px 20px;font-size:11px;color:#999;text-align:center">HSE Tags 2025/2026 — message automatique</div>' +
      '</div>';

    var mailOpts = { to:to, subject:subject, htmlBody:html, name:'HSE Tags' };
    if (ccEmails.length) mailOpts.cc = ccEmails.join(',');
    if (photoBlob) mailOpts.inlineImages = { photoAvant: photoBlob };
    MailApp.sendEmail(mailOpts);

    // Trace de l'envoi (colonne V) : "yyyy-MM-dd HH:mm → destinataire"
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var trace = stamp + ' → ' + to + (ccEmails.length ? ' (CC: ' + ccEmails.join(', ') + ')' : '');
    try { sheet.getRange(ri, C.EMAIL_SENT + 1).setValue(trace); } catch(e) {}

    return { success:true, to:to, cc:ccEmails.join(', '), responsable:resp, emailSent:trace };
  } catch(e) { return { success:false, error:e.toString() }; }
}

// ================================================================
//  ANALYSE IA DE LA PHOTO → description suggérée
//  Compatible OpenAI (Groq par défaut, gratuit). Configurable via
//  Propriétés du script :
//    AI_API_KEY  (obligatoire)  ex: clé Groq (gsk_...) ou OpenRouter
//    AI_API_URL  (optionnel)    défaut Groq
//    AI_MODEL    (optionnel)    défaut Llama 4 Scout (vision)
//  Pour OpenRouter : AI_API_URL=https://openrouter.ai/api/v1/chat/completions
//                    AI_MODEL=meta-llama/llama-3.2-11b-vision-instruct:free
// ================================================================
function analyzePhoto(p) {
  try {
    if (!p || !p.photo || p.photo.length < 10) return { success:false, error:'Photo manquante' };
    var props = PropertiesService.getScriptProperties();
    var key   = props.getProperty('AI_API_KEY') || props.getProperty('GROQ_API_KEY');
    if (!key) return { success:false, error:'Clé AI_API_KEY non configurée (Propriétés du script)' };
    var url   = props.getProperty('AI_API_URL') || 'https://api.groq.com/openai/v1/chat/completions';
    var model = props.getProperty('AI_MODEL')   || 'meta-llama/llama-4-scout-17b-16e-instruct';

    var clean   = String(p.photo).replace(/^data:[^;]+;base64,/, '');
    var dataUrl = 'data:' + (p.mime || 'image/jpeg') + ';base64,' + clean;

    var prompt = "Tu es un inspecteur HSE (Hygiène Sécurité Environnement) en milieu industriel. " +
      "Décris en UNE seule phrase concise et factuelle, en français, l'anomalie ou le danger de sécurité " +
      "visible sur la photo (équipement concerné, défaut observé, risque). " +
      "Pas de préambule ni de liste : uniquement la description.";

    var payload = {
      model: model,
      messages: [{
        role: 'user',
        content: [
          { type: 'text', text: prompt },
          { type: 'image_url', image_url: { url: dataUrl } }
        ]
      }],
      temperature: 0.4,
      max_tokens: 200
    };

    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + key, 'HTTP-Referer': 'https://hse-tags.app', 'X-Title': 'HSE Tags' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    var raw  = res.getContentText() || '{}';
    var data = {};
    try { data = JSON.parse(raw); } catch(e) {}
    if (code !== 200) {
      var msg = (data.error && (data.error.message || data.error)) || raw.substring(0, 200);
      return { success:false, error:'IA ' + code + ': ' + msg };
    }
    var text = '';
    try { text = data.choices[0].message.content.trim(); } catch(e) {}
    if (!text) return { success:false, error:'Réponse vide de l\'IA' };
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
