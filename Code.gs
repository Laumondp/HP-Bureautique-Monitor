/**
 * HP Bureautique Monitor - Google Apps Script
 * Surveillance des imprimantes HP Bureautique sur un plan de bureau
 *
 * Les imprimantes sont stockées dans la feuille "Printers" du Google Sheets
 * Colonnes: id, name, ip, model, location, serial, mac, x, y, description
 */

// Colonnes de la feuille Printers
const PRINTER_COLUMNS = ['id', 'name', 'ip', 'model', 'location', 'serial', 'mac', 'x', 'y', 'description', 'contractEndDate'];

// ID du Google Sheets (hardcodé pour éviter les problèmes de configuration)
// TODO: Remplacer par l'ID de votre nouveau Google Sheets
const HARDCODED_SPREADSHEET_ID = '';

// Clé API pour sécuriser les requêtes du script PowerShell
// IMPORTANT: Changez cette clé et gardez-la secrète !
const API_SECRET_KEY = 'HPBM-2026-Abc123Xyz789';

// Liste des emails autorisés à accéder à l'application
const AUTHORIZED_EMAILS = [
  'laumondp@gmail.com',
  'philippe.laumond@valeo.com'
];

/**
 * Vérifie si l'utilisateur actuel est autorisé
 */
function isUserAuthorized() {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  return AUTHORIZED_EMAILS.map(e => e.toLowerCase()).includes(userEmail);
}

/**
 * Déploie l'application web
 */
function doGet(e) {
  // Vérifier l'autorisation
  if (!isUserAuthorized()) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family: Arial, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: #f0f4f8;">' +
      '<div style="text-align: center; padding: 40px; background: white; border-radius: 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.1);">' +
      '<h1 style="color: #ef4444;">Accès non autorisé</h1>' +
      '<p style="color: #64748b;">Votre compte n\'est pas autorisé à accéder à cette application.</p>' +
      '<p style="color: #94a3b8; font-size: 0.9em;">Contactez l\'administrateur pour obtenir l\'accès.</p>' +
      '</div></body></html>'
    ).setTitle('Accès refusé');
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Moniteur Imprimantes HP Bureautique')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==================== GESTION DES IMPRIMANTES ====================

/**
 * Récupère l'ID du spreadsheet (hardcodé ou depuis les propriétés)
 */
function getSpreadsheetId() {
  // Utiliser l'ID hardcodé en priorité, sinon fallback sur les propriétés
  return HARDCODED_SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

/**
 * Récupère la configuration des imprimantes depuis Google Sheets
 */
function getPrinterConfig() {
  const spreadsheetId = getSpreadsheetId();

  if (!spreadsheetId) {
    return getDefaultPrinterConfig();
  }

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('Printers');

    if (!sheet) {
      return getDefaultPrinterConfig();
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return [];
    }

    const printers = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[2]) continue; // Ignorer les lignes vides

      printers.push({
        id: row[0] || 'printer_' + Date.now() + '_' + i,
        name: row[1] || '',
        ip: row[2] || '',
        model: row[3] || '',
        location: row[4] || '',
        serial: row[5] || '',
        mac: row[6] || '',
        x: parseInt(row[7]) || 0,
        y: parseInt(row[8]) || 0,
        description: row[9] || '',
        contractEndDate: row[10] ? (row[10] instanceof Date ? row[10].toISOString().split('T')[0] : row[10]) : '',
        status: 'unknown'
      });
    }

    return printers;

  } catch (e) {
    console.error('Erreur lecture feuille Printers:', e);
    const stored = PropertiesService.getScriptProperties().getProperty('PRINTER_CONFIG');
    if (stored) {
      return JSON.parse(stored);
    }
    return getDefaultPrinterConfig();
  }
}

/**
 * Configuration par défaut des imprimantes
 */
function getDefaultPrinterConfig() {
  return [
    {
      id: 'printer_1',
      name: 'ZT410-Expedition-01',
      ip: '8.8.8.8',
      model: 'ZT410',
      location: 'Zone Expédition',
      serial: '',
      mac: '',
      x: 150,
      y: 200,
      status: 'unknown',
      description: 'Google DNS - Test'
    }
  ];
}

/**
 * Sauvegarde la configuration des imprimantes dans Google Sheets
 */
function savePrinterConfig(printers) {
  const spreadsheetId = getSpreadsheetId();

  if (!spreadsheetId) {
    PropertiesService.getScriptProperties().setProperty('PRINTER_CONFIG', JSON.stringify(printers));
    return { success: true, count: printers.length };
  }

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = ss.getSheetByName('Printers');

    if (!sheet) {
      sheet = ss.insertSheet('Printers');
      sheet.getRange('A1:K1').setValues([PRINTER_COLUMNS]);
      sheet.getRange('A1:K1').setFontWeight('bold');
      sheet.getRange('A1:K1').setBackground('#4285f4');
      sheet.getRange('A1:K1').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    // Effacer les données existantes (sauf l'en-tête)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, PRINTER_COLUMNS.length).clearContent();
    }

    // Écrire les nouvelles données
    if (printers.length > 0) {
      const data = printers.map(p => [
        p.id || '',
        p.name || '',
        p.ip || '',
        p.model || '',
        p.location || '',
        p.serial || '',
        p.mac || '',
        p.x || 0,
        p.y || 0,
        p.description || '',
        p.contractEndDate || ''
      ]);

      sheet.getRange(2, 1, data.length, PRINTER_COLUMNS.length).setValues(data);
    }

    return { success: true, count: printers.length };

  } catch (e) {
    console.error('Erreur sauvegarde feuille Printers:', e);
    throw e;
  }
}

/**
 * Récupère la liste des imprimantes avec leur statut
 */
function getPrinters() {
  const printers = getPrinterConfig();
  const spreadsheetId = getSpreadsheetId();

  if (spreadsheetId) {
    try {
      const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Status');
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        const statusMap = {};

        for (let i = 1; i < data.length; i++) {
          statusMap[data[i][0]] = {
            status: data[i][1],
            lastCheck: data[i][2],
            responseTime: data[i][3]
          };
        }

        printers.forEach(printer => {
          if (statusMap[printer.ip]) {
            printer.status = statusMap[printer.ip].status;
            printer.lastCheck = statusMap[printer.ip].lastCheck;
            printer.responseTime = statusMap[printer.ip].responseTime;
          }
        });
      }
    } catch (e) {
      console.error('Erreur lecture statuts:', e);
    }
  }

  return printers;
}

/**
 * Ajoute une nouvelle imprimante
 */
function addPrinter(printer) {
  const printers = getPrinterConfig();

  // Valider le format de l'adresse MAC
  if (printer.mac) {
    const macError = getMACError(printer.mac);
    if (macError) {
      return { success: false, error: macError };
    }
  }

  // Vérifier l'unicité de l'IP
  if (printer.ip && printers.some(p => p.ip === printer.ip)) {
    return { success: false, error: 'Cette adresse IP existe déjà' };
  }

  // Vérifier l'unicité de l'adresse MAC
  if (printer.mac && printers.some(p => p.mac && p.mac.toLowerCase() === printer.mac.toLowerCase())) {
    return { success: false, error: 'Cette adresse MAC existe déjà' };
  }

  // Vérifier l'unicité du numéro de série
  if (printer.serial && printers.some(p => p.serial && p.serial.toLowerCase() === printer.serial.toLowerCase())) {
    return { success: false, error: 'Ce numéro de série existe déjà' };
  }

  printer.id = 'printer_' + Date.now();
  printers.push(printer);

  savePrinterConfig(printers);
  return { success: true, printer: printer };
}

/**
 * Supprime une imprimante
 */
function removePrinter(printerId) {
  let printers = getPrinterConfig();
  const initialCount = printers.length;

  printers = printers.filter(p => p.id !== printerId);

  if (printers.length === initialCount) {
    return { success: false, error: 'Imprimante non trouvée' };
  }

  savePrinterConfig(printers);
  return { success: true };
}

/**
 * Met à jour une imprimante existante
 */
function updatePrinter(printerId, updates) {
  const printers = getPrinterConfig();
  const index = printers.findIndex(p => p.id === printerId);

  if (index === -1) {
    return { success: false, error: 'Imprimante non trouvée' };
  }

  // Valider le format de l'adresse MAC
  if (updates.mac) {
    const macError = getMACError(updates.mac);
    if (macError) {
      return { success: false, error: macError };
    }
  }

  // Vérifier l'unicité de l'IP (si modifiée)
  if (updates.ip && updates.ip !== printers[index].ip) {
    if (printers.some(p => p.ip === updates.ip)) {
      return { success: false, error: 'Cette adresse IP existe déjà' };
    }
  }

  // Vérifier l'unicité de l'adresse MAC (si modifiée)
  if (updates.mac && updates.mac !== printers[index].mac) {
    if (printers.some(p => p.mac && p.mac.toLowerCase() === updates.mac.toLowerCase())) {
      return { success: false, error: 'Cette adresse MAC existe déjà' };
    }
  }

  // Vérifier l'unicité du numéro de série (si modifié)
  if (updates.serial && updates.serial !== printers[index].serial) {
    if (printers.some(p => p.serial && p.serial.toLowerCase() === updates.serial.toLowerCase())) {
      return { success: false, error: 'Ce numéro de série existe déjà' };
    }
  }

  printers[index] = { ...printers[index], ...updates };
  savePrinterConfig(printers);

  return { success: true, printer: printers[index] };
}

// ==================== INITIALISATION ====================

/**
 * Initialise la feuille Google Sheets
 * Crée deux feuilles: "Status" et "Printers"
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.create('HP Bureautique Monitor');

  // Feuille Status
  const statusSheet = ss.getActiveSheet();
  statusSheet.setName('Status');
  statusSheet.getRange('A1:D1').setValues([['IP', 'Status', 'LastCheck', 'ResponseTime']]);
  statusSheet.getRange('A1:D1').setFontWeight('bold');
  statusSheet.getRange('A1:D1').setBackground('#34a853');
  statusSheet.getRange('A1:D1').setFontColor('#ffffff');

  // Feuille Printers
  const printersSheet = ss.insertSheet('Printers');
  printersSheet.getRange('A1:K1').setValues([PRINTER_COLUMNS]);
  printersSheet.getRange('A1:K1').setFontWeight('bold');
  printersSheet.getRange('A1:K1').setBackground('#4285f4');
  printersSheet.getRange('A1:K1').setFontColor('#ffffff');
  printersSheet.setFrozenRows(1);

  // Ajuster la largeur des colonnes
  printersSheet.setColumnWidth(1, 150);
  printersSheet.setColumnWidth(2, 180);
  printersSheet.setColumnWidth(3, 120);
  printersSheet.setColumnWidth(4, 80);
  printersSheet.setColumnWidth(5, 150);
  printersSheet.setColumnWidth(6, 150);
  printersSheet.setColumnWidth(7, 150);
  printersSheet.setColumnWidth(8, 60);
  printersSheet.setColumnWidth(9, 60);
  printersSheet.setColumnWidth(10, 250);

  // Stocker l'ID
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  // Migrer les imprimantes existantes
  const storedConfig = PropertiesService.getScriptProperties().getProperty('PRINTER_CONFIG');
  let printers = storedConfig ? JSON.parse(storedConfig) : getDefaultPrinterConfig();

  // Écrire les imprimantes
  if (printers.length > 0) {
    const data = printers.map(p => [
      p.id || '',
      p.name || '',
      p.ip || '',
      p.model || '',
      p.location || '',
      p.serial || '',
      p.mac || '',
      p.x || 0,
      p.y || 0,
      p.description || '',
      p.contractEndDate || ''
    ]);
    printersSheet.getRange(2, 1, data.length, PRINTER_COLUMNS.length).setValues(data);

    // Ajouter les statuts initiaux
    printers.forEach((printer, index) => {
      statusSheet.getRange(index + 2, 1, 1, 4).setValues([
        [printer.ip, 'unknown', '', 0]
      ]);
    });
  }

  return {
    spreadsheetId: ss.getId(),
    url: ss.getUrl(),
    message: 'Feuille de calcul créée avec succès!'
  };
}

/**
 * Obtient l'URL de la feuille Google Sheets
 */
function getSpreadsheetUrl() {
  const spreadsheetId = getSpreadsheetId();

  if (!spreadsheetId) {
    return { success: false, error: 'SPREADSHEET_ID non configuré' };
  }

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    return {
      success: true,
      url: ss.getUrl(),
      name: ss.getName()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== STATUTS ====================

/**
 * Met à jour le statut d'une imprimante
 */
function updatePrinterStatus(ip, status, responseTime) {
  const spreadsheetId = getSpreadsheetId();
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID non configuré');
  }

  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Status');
  if (!sheet) {
    throw new Error('Feuille "Status" non trouvée');
  }

  const data = sheet.getDataRange().getValues();
  const now = new Date().toISOString();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === ip) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[status, now, responseTime]]);
      return { success: true, updated: true };
    }
  }

  sheet.appendRow([ip, status, now, responseTime]);
  return { success: true, updated: false, created: true };
}

/**
 * Met à jour plusieurs imprimantes en batch
 */
function updatePrinterStatusBatch(statusArray) {
  const results = [];
  statusArray.forEach(item => {
    try {
      const result = updatePrinterStatus(item.ip, item.status, item.responseTime || 0);
      results.push({ ip: item.ip, ...result });
    } catch (e) {
      results.push({ ip: item.ip, success: false, error: e.message });
    }
  });
  return results;
}

// ==================== API EXTERNE ====================

/**
 * Vérifie la clé API
 */
function isValidApiKey(key) {
  return key === API_SECRET_KEY;
}

/**
 * Endpoint API pour le script de ping externe
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Vérifier la clé API pour les actions sensibles
    if (data.action === 'getPrinters') {
      // Vérifier la clé API
      if (!isValidApiKey(data.apiKey)) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'Clé API invalide ou manquante'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const printers = getPrinterConfig();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        printers: printers.map(p => ({ ip: p.ip, name: p.name }))
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'updateStatus') {
      // Vérifier la clé API
      if (!isValidApiKey(data.apiKey)) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'Clé API invalide ou manquante'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const updateResults = updatePrinterStatusBatch(data.statuses);

      let alertResult = { alertsSent: 0 };
      if (data.statusChanges && data.statusChanges.length > 0) {
        alertResult = checkAndSendAlerts(data.statusChanges);
      }

      const successCount = updateResults.filter(r => r.success).length;
      const errorCount = updateResults.filter(r => !r.success).length;

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        updated: successCount,
        errors: errorCount,
        results: updateResults,
        alertsSent: alertResult.alertsSent,
        alerts: alertResult.alerts || []
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ error: 'Action inconnue' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ error: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== PLAN D'USINE ====================

/**
 * Récupère les informations du plan d'usine
 */
function getFloorPlanConfig() {
  const stored = PropertiesService.getScriptProperties().getProperty('FLOOR_PLAN_CONFIG');
  if (stored) {
    return JSON.parse(stored);
  }

  return {
    width: 1156,
    height: 938,
    backgroundImage: '',
    gridSize: 50
  };
}

/**
 * Récupère l'image du plan d'usine en base64
 */
function getFloorPlanImage() {
  const fileId = PropertiesService.getScriptProperties().getProperty('FLOOR_PLAN_FILE_ID');
  if (!fileId) return '';

  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    console.error('Erreur chargement image plan:', e);
    return '';
  }
}

/**
 * Configure l'image du plan d'usine
 */
function setupFloorPlanImage() {
  var files = DriveApp.getFilesByName('Plan usine.png');
  if (files.hasNext()) {
    var file = files.next();
    PropertiesService.getScriptProperties().setProperty('FLOOR_PLAN_FILE_ID', file.getId());
    return { success: true, fileId: file.getId(), message: 'Image configurée avec succès' };
  }
  return { success: false, message: 'Fichier "Plan usine.png" non trouvé dans Google Drive' };
}

/**
 * Sauvegarde la configuration du plan d'usine
 */
function saveFloorPlanConfig(config) {
  PropertiesService.getScriptProperties().setProperty('FLOOR_PLAN_CONFIG', JSON.stringify(config));
  return { success: true };
}

// ==================== IMAGES IMPRIMANTES ====================

/**
 * Mapping des modèles vers les noms de fichiers image
 */
const PRINTER_IMAGE_MAPPING = {
  '110PAX4': '110PAXL4.jpg',
  '110PAXL4': '110PAXL4.jpg',
  'Z4M': 'Z4M.jpg',
  'Z6M': 'Z6M.jpg',
  'Z6MPlus': 'Z6MPlus.jpg',
  'ZE500': 'ZE500L4.jpg',
  'ZE500L4': 'ZE500L4.jpg',
  'ZM400': 'ZM400.jpg',
  'ZM600': 'ZM600.jpg',
  'ZT230': 'Z230.jpg',
  'ZT410': 'Z410.jpg',
  'ZT411': 'Z411.jpg',
  'ZT420': 'Z420.jpg',
  'ZT421': 'Z421.jpg'
};

/**
 * Configure les images des imprimantes depuis Google Drive
 * Exécuter cette fonction une fois après avoir uploadé les images
 */
function setupPrinterImages() {
  const props = PropertiesService.getScriptProperties();
  const results = [];

  for (const [model, filename] of Object.entries(PRINTER_IMAGE_MAPPING)) {
    const files = DriveApp.getFilesByName(filename);
    if (files.hasNext()) {
      const file = files.next();
      props.setProperty('PRINTER_IMAGE_' + model, file.getId());
      results.push({ model: model, status: 'OK', fileId: file.getId() });
    } else {
      results.push({ model: model, status: 'NOT_FOUND', filename: filename });
    }
  }

  console.log('Configuration images imprimantes:', results);
  return results;
}

/**
 * Récupère toutes les images des imprimantes en base64
 */
function getPrinterImages() {
  const props = PropertiesService.getScriptProperties();
  const images = {};

  for (const model of Object.keys(PRINTER_IMAGE_MAPPING)) {
    const fileId = props.getProperty('PRINTER_IMAGE_' + model);
    if (fileId) {
      try {
        const file = DriveApp.getFileById(fileId);
        const blob = file.getBlob();
        images[model] = 'data:image/jpeg;base64,' + Utilities.base64Encode(blob.getBytes());
      } catch (e) {
        console.error('Erreur chargement image ' + model + ':', e);
        images[model] = null;
      }
    } else {
      images[model] = null;
    }
  }

  return images;
}

// ==================== ALERTES ====================

/**
 * Récupère la configuration des alertes
 */
function getAlertConfig() {
  const stored = PropertiesService.getScriptProperties().getProperty('ALERT_CONFIG');
  if (stored) {
    return JSON.parse(stored);
  }

  return {
    enabled: true,
    emailRecipients: [],
    alertOnOffline: true,
    alertOnBackOnline: true,
    cooldownMinutes: 5
  };
}

/**
 * Sauvegarde la configuration des alertes
 */
function saveAlertConfig(config) {
  PropertiesService.getScriptProperties().setProperty('ALERT_CONFIG', JSON.stringify(config));
  return { success: true };
}

/**
 * Récupère l'historique des alertes
 */
function getAlertHistory() {
  const stored = PropertiesService.getScriptProperties().getProperty('ALERT_HISTORY');
  if (stored) {
    return JSON.parse(stored);
  }
  return {};
}

/**
 * Sauvegarde l'historique des alertes
 */
function saveAlertHistory(history) {
  PropertiesService.getScriptProperties().setProperty('ALERT_HISTORY', JSON.stringify(history));
}

/**
 * Vérifie et envoie les alertes
 */
function checkAndSendAlerts(statusChanges) {
  const alertConfig = getAlertConfig();

  if (!alertConfig.enabled || alertConfig.emailRecipients.length === 0) {
    return { alertsSent: 0, message: 'Alertes désactivées ou aucun destinataire' };
  }

  const printers = getPrinterConfig();
  const printerMap = {};
  printers.forEach(p => printerMap[p.ip] = p);

  const alertHistory = getAlertHistory();
  const now = new Date().getTime();
  const cooldownMs = alertConfig.cooldownMinutes * 60 * 1000;

  let alertsSent = 0;
  const alertsToSend = [];

  statusChanges.forEach(change => {
    const printer = printerMap[change.ip];
    if (!printer) return;

    const lastAlert = alertHistory[change.ip] || 0;
    if (now - lastAlert < cooldownMs) {
      return;
    }

    const webAppUrl = ScriptApp.getService().getUrl();

    if (change.previousStatus === 'online' && change.newStatus === 'offline' && alertConfig.alertOnOffline) {
      alertsToSend.push({
        printer: printer,
        type: 'offline',
        message: `⚠️ ALERTE: L'imprimante "${printer.name}" est HORS LIGNE\n\nDétails:\n- Modèle: ${printer.model}\n- IP: ${printer.ip}\n- Emplacement: ${printer.location}\n- Heure: ${new Date().toLocaleString('fr-FR')}\n\n📍 Voir le moniteur:\n${webAppUrl}`
      });
      alertHistory[change.ip] = now;
    }

    if (change.previousStatus === 'offline' && change.newStatus === 'online' && alertConfig.alertOnBackOnline) {
      alertsToSend.push({
        printer: printer,
        type: 'online',
        message: `✅ RETABLIE: L'imprimante "${printer.name}" est de nouveau EN LIGNE\n\nDétails:\n- Modèle: ${printer.model}\n- IP: ${printer.ip}\n- Emplacement: ${printer.location}\n- Heure: ${new Date().toLocaleString('fr-FR')}\n\n📍 Voir le moniteur:\n${webAppUrl}`
      });
      alertHistory[change.ip] = now;
    }
  });

  alertsToSend.forEach(alert => {
    try {
      const subject = alert.type === 'offline'
        ? `[ALERTE] Imprimante ${alert.printer.name} - HORS LIGNE`
        : `[OK] Imprimante ${alert.printer.name} - Rétablie`;

      MailApp.sendEmail({
        to: alertConfig.emailRecipients.join(','),
        subject: subject,
        body: alert.message
      });
      alertsSent++;
    } catch (e) {
      console.error('Erreur envoi email:', e);
    }
  });

  saveAlertHistory(alertHistory);

  return { alertsSent: alertsSent, alerts: alertsToSend.map(a => a.printer.name) };
}

/**
 * Ajoute un email destinataire
 */
function addAlertRecipient(email) {
  const config = getAlertConfig();
  if (!config.emailRecipients.includes(email)) {
    config.emailRecipients.push(email);
    saveAlertConfig(config);
  }
  return { success: true, recipients: config.emailRecipients };
}

/**
 * Supprime un email destinataire
 */
function removeAlertRecipient(email) {
  const config = getAlertConfig();
  config.emailRecipients = config.emailRecipients.filter(e => e !== email);
  saveAlertConfig(config);
  return { success: true, recipients: config.emailRecipients };
}

/**
 * Test d'envoi d'alerte
 */
function testAlert(email) {
  try {
    const webAppUrl = ScriptApp.getService().getUrl();
    MailApp.sendEmail({
      to: email,
      subject: '[TEST] HP Bureautique Monitor - Test d\'alerte',
      body: 'Ceci est un test du système d\'alerte.\n\nSi vous recevez cet email, les alertes fonctionnent correctement.\n\n📍 Accéder au moniteur:\n' + webAppUrl
    });
    return { success: true, message: 'Email de test envoyé' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== IMPORT/EXPORT CSV ====================

/**
 * Importe des imprimantes depuis un CSV
 */
function importPrintersFromCSV(csvContent, replaceAll) {
  try {
    const lines = csvContent.trim().split(/\r?\n/);

    if (lines.length < 2) {
      return { success: false, error: 'Le fichier CSV doit contenir au moins un en-tête et une ligne de données' };
    }

    const header = parseCSVLine(lines[0]);
    const requiredColumns = ['name', 'ip', 'model', 'location', 'serial', 'mac', 'x', 'y'];
    const headerLower = header.map(h => h.toLowerCase().trim());

    for (const col of requiredColumns) {
      if (!headerLower.includes(col)) {
        return { success: false, error: `Colonne manquante: ${col}` };
      }
    }

    const colIndex = {};
    headerLower.forEach((col, idx) => {
      colIndex[col] = idx;
    });

    const newPrinters = [];
    const errors = [];
    const seenIPs = new Set();

    for (let i = 1; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;

      const values = parseCSVLine(line);
      const lineNum = i + 1;

      const name = values[colIndex['name']]?.trim();
      const ip = values[colIndex['ip']]?.trim();
      const model = values[colIndex['model']]?.trim();
      const location = values[colIndex['location']]?.trim();
      const x = parseInt(values[colIndex['x']]);
      const y = parseInt(values[colIndex['y']]);
      const serial = values[colIndex['serial']]?.trim() || '';
      const mac = values[colIndex['mac']]?.trim() || '';
      const description = colIndex['description'] !== undefined ? (values[colIndex['description']]?.trim() || '') : '';

      if (!name || !ip || !isValidIP(ip) || seenIPs.has(ip) || !model || !location || isNaN(x) || isNaN(y)) {
        errors.push(`Ligne ${lineNum}: données invalides`);
        continue;
      }

      seenIPs.add(ip);
      newPrinters.push({
        id: 'printer_' + Date.now() + '_' + i,
        name, ip, model, location, serial, mac, x, y, description,
        status: 'unknown'
      });
    }

    if (newPrinters.length === 0) {
      return { success: false, error: 'Aucune imprimante valide trouvée', details: errors };
    }

    if (replaceAll) {
      savePrinterConfig(newPrinters);
    } else {
      const existingPrinters = getPrinterConfig();
      const existingIPs = new Set(existingPrinters.map(p => p.ip));
      const printersToAdd = newPrinters.filter(p => !existingIPs.has(p.ip));
      savePrinterConfig([...existingPrinters, ...printersToAdd]);
    }

    return { success: true, imported: newPrinters.length, warnings: errors.length > 0 ? errors : null };

  } catch (e) {
    return { success: false, error: 'Erreur: ' + e.message };
  }
}

/**
 * Parse une ligne CSV
 */
function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current);

  return result;
}

/**
 * Valide une adresse IP
 */
function isValidIP(ip) {
  const pattern = /^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$/;
  return pattern.test(ip);
}

/**
 * Valide une adresse MAC
 */
function isValidMAC(mac) {
  if (!mac) return true; // MAC optionnelle
  const pattern = /^([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})$/;
  return pattern.test(mac);
}

/**
 * Vérifie les erreurs courantes dans une adresse MAC et retourne un message d'erreur détaillé
 */
function getMACError(mac) {
  if (!mac) return null;

  // Vérifier confusion O/0 (lettre O au lieu du chiffre 0)
  if (/[oO]/.test(mac)) {
    return 'Confusion O/0 : utilisez le chiffre "0" (zéro) et non la lettre "O"';
  }

  // Vérifier le nombre de segments
  const segments = mac.split(/[:-]/);
  if (segments.length !== 6) {
    return `Format invalide : ${segments.length} segment(s) au lieu de 6. Format attendu : XX:XX:XX:XX:XX:XX`;
  }

  // Vérifier chaque segment
  for (let i = 0; i < segments.length; i++) {
    if (segments[i].length !== 2) {
      return `Segment ${i + 1} invalide : "${segments[i]}" doit avoir 2 caractères`;
    }
    if (!/^[0-9A-Fa-f]{2}$/.test(segments[i])) {
      return `Segment ${i + 1} invalide : "${segments[i]}" contient des caractères non hexadécimaux (0-9, A-F)`;
    }
  }

  // Vérifier le format global
  const pattern = /^([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})$/;
  if (!pattern.test(mac)) {
    return 'Format invalide. Utilisez : XX:XX:XX:XX:XX:XX ou XX-XX-XX-XX-XX-XX';
  }

  return null;
}

/**
 * Génère un template CSV
 */
function getCSVTemplate() {
  return 'name,ip,model,location,serial,mac,x,y,description\nZT410-Example,192.168.1.101,ZT410,Zone A,SN123,00:1A:2B:3C:4D:5E,200,150,Description';
}

/**
 * Exporte les imprimantes en CSV
 */
function exportPrintersToCSV() {
  const printers = getPrinterConfig();
  const header = 'name,ip,model,location,serial,mac,x,y,description';

  const lines = printers.map(p => {
    const desc = (p.description || '').replace(/"/g, '""');
    const serial = (p.serial || '').replace(/"/g, '""');
    const mac = (p.mac || '').replace(/"/g, '""');
    return `${p.name},${p.ip},${p.model},${p.location},"${serial}","${mac}",${p.x},${p.y},"${desc}"`;
  });

  return header + '\n' + lines.join('\n');
}

/**
 * Liste des modèles supportés
 */
function getSupportedModels() {
  return [
    { model: 'ZT410', description: 'Imprimante industrielle 4 pouces' },
    { model: 'ZT411', description: 'Imprimante industrielle 4 pouces (nouvelle génération)' },
    { model: 'ZT420', description: 'Imprimante industrielle 6 pouces' },
    { model: 'ZT421', description: 'Imprimante industrielle 6 pouces (nouvelle génération)' },
    { model: 'ZM600', description: 'Imprimante industrielle haute performance' },
    { model: 'Z6M', description: 'Imprimante industrielle grand format' },
    { model: 'ZM400', description: 'Imprimante industrielle moyenne gamme' },
    { model: 'ZT230', description: 'Imprimante industrielle compacte' }
  ];
}

/**
 * Réinitialise la configuration
 */
function resetToDefaultConfig() {
  const defaultConfig = getDefaultPrinterConfig();
  savePrinterConfig(defaultConfig);
  return { success: true, printers: defaultConfig };
}

// ==================== PROTECTION DOUBLONS GOOGLE SHEETS ====================

/**
 * Trigger onEdit - Vérifie les doublons lors de modifications directes dans Google Sheets
 * Gère les modifications simples ET les copier-coller
 * Pour installer ce trigger: Exécutez installEditTrigger() une seule fois
 */
function onSheetEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();

    // Ne vérifier que la feuille "Printers"
    if (sheet.getName() !== 'Printers') return;

    // Lancer la validation complète de la feuille
    const errors = validatePrintersSheet(sheet);

    if (errors.length > 0) {
      e.source.toast(
        errors.join('\n'),
        '⚠️ Erreur de validation',
        10
      );
    }
  } catch (error) {
    console.error('Erreur onSheetEdit:', error);
  }
}

/**
 * Valide la feuille Printers et retourne les erreurs trouvées
 * Marque les cellules en erreur en rouge
 */
function validatePrintersSheet(sheet) {
  const data = sheet.getDataRange().getValues();
  const errors = [];

  if (data.length <= 1) return errors;

  // Maps pour détecter les doublons
  const idMap = {};      // col 1 (index 0)
  const ipMap = {};      // col 3 (index 2)
  const serialMap = {};  // col 6 (index 5)
  const macMap = {};     // col 7 (index 6)

  // Réinitialiser le fond des colonnes ID, IP, Model, Serial, MAC
  if (data.length > 1) {
    sheet.getRange(2, 1, data.length - 1, 1).setBackground(null); // ID
    sheet.getRange(2, 3, data.length - 1, 1).setBackground(null); // IP
    sheet.getRange(2, 4, data.length - 1, 1).setBackground(null); // Model
    sheet.getRange(2, 6, data.length - 1, 1).setBackground(null); // Serial
    sheet.getRange(2, 7, data.length - 1, 1).setBackground(null); // MAC
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 1;

    const id = row[0] ? row[0].toString().trim() : '';
    const ip = row[2] ? row[2].toString().trim() : '';
    const model = row[3] ? row[3].toString().trim() : '';
    const serial = row[5] ? row[5].toString().trim() : '';
    const mac = row[6] ? row[6].toString().trim() : '';

    // Liste des modèles valides
    const validModels = ['110PAX4', 'Z4M', 'Z6M', 'Z6MPLUS', 'ZE500', 'ZE500L4', 'ZM400', 'ZM600', 'ZT230', 'ZT410', 'ZT411', 'ZT420', 'ZT421'];

    // Vérifier ID
    if (id) {
      if (idMap[id.toLowerCase()]) {
        errors.push(`Ligne ${rowNum}: ID en double "${id}"`);
        sheet.getRange(rowNum, 1).setBackground('#ffcccc');
        sheet.getRange(idMap[id.toLowerCase()], 1).setBackground('#ffcccc');
      } else {
        idMap[id.toLowerCase()] = rowNum;
      }
    }

    // Vérifier IP
    if (ip) {
      if (!isValidIP(ip)) {
        errors.push(`Ligne ${rowNum}: IP invalide "${ip}"`);
        sheet.getRange(rowNum, 3).setBackground('#ffcccc');
      } else if (ipMap[ip.toLowerCase()]) {
        errors.push(`Ligne ${rowNum}: IP en double "${ip}"`);
        sheet.getRange(rowNum, 3).setBackground('#ffcccc');
        sheet.getRange(ipMap[ip.toLowerCase()], 3).setBackground('#ffcccc');
      } else {
        ipMap[ip.toLowerCase()] = rowNum;
      }
    }

    // Vérifier Model
    if (model) {
      if (!validModels.includes(model.toUpperCase())) {
        errors.push(`Ligne ${rowNum}: Modèle invalide "${model}". Valides: ${validModels.join(', ')}`);
        sheet.getRange(rowNum, 4).setBackground('#ffcccc');
      }
    }

    // Vérifier Serial
    if (serial) {
      if (serialMap[serial.toLowerCase()]) {
        errors.push(`Ligne ${rowNum}: Série en double "${serial}"`);
        sheet.getRange(rowNum, 6).setBackground('#ffcccc');
        sheet.getRange(serialMap[serial.toLowerCase()], 6).setBackground('#ffcccc');
      } else {
        serialMap[serial.toLowerCase()] = rowNum;
      }
    }

    // Vérifier MAC
    if (mac) {
      const macError = getMACError(mac);
      if (macError) {
        errors.push(`Ligne ${rowNum}: ${macError}`);
        sheet.getRange(rowNum, 7).setBackground('#ffcccc');
      } else if (macMap[mac.toLowerCase()]) {
        errors.push(`Ligne ${rowNum}: MAC en double "${mac}"`);
        sheet.getRange(rowNum, 7).setBackground('#ffcccc');
        sheet.getRange(macMap[mac.toLowerCase()], 7).setBackground('#ffcccc');
      } else {
        macMap[mac.toLowerCase()] = rowNum;
      }
    }
  }

  return errors;
}

/**
 * Validation manuelle - Exécutez pour vérifier les doublons dans la feuille Printers
 */
function validatePrintersManually() {
  const spreadsheetId = getSpreadsheetId();
  if (!spreadsheetId) {
    return { success: false, error: 'SPREADSHEET_ID non configuré' };
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('Printers');

  if (!sheet) {
    return { success: false, error: 'Feuille Printers non trouvée' };
  }

  const errors = validatePrintersSheet(sheet);

  if (errors.length === 0) {
    ss.toast('Aucune erreur trouvée !', '✅ Validation OK', 5);
    return { success: true, message: 'Aucune erreur' };
  } else {
    ss.toast(`${errors.length} erreur(s) trouvée(s). Les cellules en erreur sont surlignées en rouge.`, '⚠️ Erreurs', 10);
    return { success: false, errors: errors };
  }
}

/**
 * Installe le trigger onEdit pour la protection des doublons
 * Exécutez cette fonction UNE SEULE FOIS depuis l'éditeur Google Apps Script
 */
function installEditTrigger() {
  // Supprimer les anciens triggers onSheetEdit
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onSheetEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Créer le nouveau trigger
  const spreadsheetId = getSpreadsheetId();
  if (!spreadsheetId) {
    return { success: false, error: 'SPREADSHEET_ID non configuré' };
  }

  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(spreadsheetId)
    .onEdit()
    .create();

  return { success: true, message: 'Trigger installé avec succès' };
}

/**
 * Désinstalle le trigger onEdit
 */
function uninstallEditTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onSheetEdit') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  return { success: true, message: `${removed} trigger(s) supprimé(s)` };
}

/**
 * Ajoute la colonne contractEndDate à la feuille Printers existante
 * Exécutez cette fonction UNE SEULE FOIS pour mettre à jour la feuille
 */
function addContractEndDateColumn() {
  const spreadsheetId = getSpreadsheetId();
  if (!spreadsheetId) {
    return { success: false, error: 'SPREADSHEET_ID non configuré' };
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('Printers');

  if (!sheet) {
    return { success: false, error: 'Feuille Printers non trouvée' };
  }

  // Vérifier si la colonne existe déjà
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.includes('contractEndDate')) {
    return { success: true, message: 'La colonne contractEndDate existe déjà' };
  }

  // Ajouter la colonne K (11ème colonne)
  const colIndex = 11;
  sheet.getRange(1, colIndex).setValue('contractEndDate');
  sheet.getRange(1, colIndex).setFontWeight('bold');
  sheet.getRange(1, colIndex).setBackground('#4285f4');
  sheet.getRange(1, colIndex).setFontColor('#ffffff');
  sheet.setColumnWidth(colIndex, 120);

  return { success: true, message: 'Colonne contractEndDate ajoutée avec succès' };
}
