/* ====================================================================
   ITSUPERTECH - Backend Google Sheets pour Briefs
   ====================================================================
   
   INSTRUCTIONS D'INSTALLATION (5 minutes) :
   
   1. Allez sur https://sheets.new (crée un nouveau Google Sheet)
   2. Dans le menu, clique: Extensions > Apps Script
   3. EFFACE tout le code existant
   4. COPIE-COLLE tout ce fichier
   5. Clique sur le bouton "Deploy" > "New deployment"
   6. Clique sur l'engrenage ⚙ > "Web app"
   7. - Execute as: Me
      - Who has access: Anyone
   8. Clique "Deploy", autorise, et COPIE l'URL
   9. Colle cette URL dans 2index.html et admin.html à la place de
      'DEPLOYEZ_LE_SCRIPT_ET_COLLEZ_VOTRE_URL_ICI'
  10. Clique sur le lien "Done" du déploiement pour revenir au script
  
   Pour mettre à jour le script après modification:
   - Clique sur le bouton "Deploy" > Manage deployments
     - Clique sur l'icône crayon ✏ du déploiement existant
     - Clique "Deploy" (la nouvelle version)
   ==================================================================== */

// Clé secrète pour protéger l'accès (changez-la!)
var SECRET = 'itsuptech_brief_2024';

function initSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', 'Données JSON', 'Entreprise', 'Statut', 'Date']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1a5fb4').setFontColor('#fff');
  }
}

function getSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getLastRow() === 0) initSheet();
  return sheet;
}

function generateId() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 20);
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function checkKey(e) {
  var key = '';
  if (e.parameter && e.parameter.key) key = e.parameter.key;
  if (!key && e.postData) {
    try { var body = JSON.parse(e.postData.contents); key = body.key || ''; } catch(ex) {}
  }
  if (key !== SECRET) return false;
  return true;
}

// =========================
//  GET - Lire les briefs
// =========================
function doGet(e) {
  try {
    if (!checkKey(e)) return jsonResponse({ error: 'Accès non autorisé' });
    
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var briefs = [];
    var search = (e.parameter.q || '').toLowerCase();
    var status = e.parameter.status || '';
    
    for (var i = 1; i < data.length; i++) {
      try {
        var brief = JSON.parse(data[i][1]);
        if (status && brief.statut !== status) continue;
        if (search) {
          var haystack = [brief.entreprise, brief.contact, brief.email, brief.telephone, brief._id || ''].join(' ').toLowerCase();
          if (haystack.indexOf(search) === -1) continue;
        }
        briefs.push(brief);
      } catch(ex) {}
    }
    
    briefs.sort(function(a, b) {
      return (b.date_creation_str || '').localeCompare(a.date_creation_str || '');
    });
    
    return jsonResponse({ briefs: briefs, total: briefs.length });
    
  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

// =========================
//  POST - Créer / Mettre à jour / Supprimer
// =========================
function doPost(e) {
  try {
    if (!checkKey(e)) return jsonResponse({ error: 'Accès non autorisé' });
    
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    
    if (action === 'create') return createBrief(body);
    if (action === 'update') return updateBrief(body);
    if (action === 'delete') return deleteBrief(body);
    if (action === 'ping') return jsonResponse({ ok: true });
    
    return jsonResponse({ error: 'Action inconnue: ' + action });
    
  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

function createBrief(body) {
  var sheet = getSheet();
  var brief = body.data || body;
  var id = generateId();
  brief._id = id;
  if (!brief.statut) brief.statut = 'nouveau';
  if (!brief.date_creation) brief.date_creation = new Date().toISOString();
  if (!brief.date_creation_str) brief.date_creation_str = new Date().toLocaleString('fr-FR');
  
  sheet.appendRow([id, JSON.stringify(brief), brief.entreprise || '', brief.statut, brief.date_creation_str]);
  
  return jsonResponse({ success: true, id: id });
}

function updateBrief(body) {
  var sheet = getSheet();
  var id = body.id;
  var updates = body.data || {};
  var values = sheet.getDataRange().getValues();
  
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      var brief = JSON.parse(values[i][1]);
      var keys = Object.keys(updates);
      for (var k = 0; k < keys.length; k++) {
        brief[keys[k]] = updates[keys[k]];
      }
      sheet.getRange(i + 1, 2).setValue(JSON.stringify(brief));
      sheet.getRange(i + 1, 4).setValue(brief.statut || '');
      return jsonResponse({ success: true });
    }
  }
  
  return jsonResponse({ error: 'Brief non trouvé' });
}

function deleteBrief(body) {
  var sheet = getSheet();
  var id = body.id;
  var values = sheet.getDataRange().getValues();
  
  for (var i = values.length - 1; i >= 1; i--) {
    if (values[i][0] === id) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ success: true });
    }
  }
  
  return jsonResponse({ error: 'Brief non trouvé' });
}