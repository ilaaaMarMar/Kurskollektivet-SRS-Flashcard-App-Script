/**
* ============================================================================
* SERVER-SIDE GOOGLE APPS SCRIPT (SM-2, INTERVAL FUZZING, HOURLY REVIEWS)
* Handles Google Sheets interactions, data processing, and API routing.
* ============================================================================
*/

var TEMPLATE_ID = "1s04BKwYk-s0WVSmsJOfCxCz4X-fqdZzR0O65ANymu9E";
var GEMINI_MODELS = [
  "gemini-3.6-flash",
  "gemini-3.5-flash",
  "gemini-3.1-flash-lite",
  "gemini-2.5-flash"
];
var CARD_ID_HEADER = "SRS Card ID";

// Private helpers end in _ so they cannot be called through google.script.run.
function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 SRS Flashcards')
    .addItem('Launch Sidebar View', 'showFlashcardSidebar')
    .addItem('Launch Pop-Out Window', 'showFlashcardDialog')
    .addToUi();
}

// HTML entry points: all three views share the same evaluated Sidebar template.
function escapeJS(str) {
  if (!str) return "";
  return str.toString()
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/'/g, "\\'")
    .replace(/</g, '\\x3c')
    .replace(/>/g, '\\x3e');
}

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Sidebar');
  var html = template.evaluate()
    .setTitle('📚 SRS Flashcards')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

  html.append('<script>window.IS_WEB_APP = true;</script>');
  if (e && e.parameter && e.parameter.deck) {
    html.append('<script>window.URL_DECK_ID = "' + escapeJS(e.parameter.deck) + '";</script>');
  }
  return html;
}

function showFlashcardSidebar() {
  var html = HtmlService.createTemplateFromFile('Sidebar').evaluate()
    .setTitle('📚 SRS Flashcards');
  html.append('<script>window.IS_WEB_APP = false;</script>');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showFlashcardDialog() {
  var html = HtmlService.createTemplateFromFile('Sidebar').evaluate()
    .setWidth(600)
    .setHeight(720);
  html.append('<script>window.IS_WEB_APP = false;</script>');
  SpreadsheetApp.getUi().showModelessDialog(html, '📚 SRS Flashcards');
}

function resolveSpreadsheet(customDeckId) {
  if (customDeckId && customDeckId.toString().trim().length > 0) {
    return SpreadsheetApp.openById(customDeckId.toString().trim());
  }
  try {
    return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(TEMPLATE_ID);
  } catch (e) {
    return SpreadsheetApp.openById(TEMPLATE_ID);
  }
}

// Spreadsheet writes pass through this guard to prevent formula injection.
function preventInjection(text) {
  var str = text ? text.toString() : '';
  if (/^[=+\-@]/.test(str)) return "'" + str;
  return str;
}

var CACHED_SCRIPT_TIMEZONE = null;

// Resolve the spreadsheet timezone once per Apps Script execution.
function getScriptTimeZone() {
  if (CACHED_SCRIPT_TIMEZONE) return CACHED_SCRIPT_TIMEZONE;

  try {
    var active = SpreadsheetApp.getActive();
    if (active) {
      CACHED_SCRIPT_TIMEZONE = active.getSpreadsheetTimeZone();
      if (CACHED_SCRIPT_TIMEZONE) return CACHED_SCRIPT_TIMEZONE;
    }
  } catch (e) {}

  try {
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs) {
      CACHED_SCRIPT_TIMEZONE = activeSs.getSpreadsheetTimeZone();
      if (CACHED_SCRIPT_TIMEZONE) return CACHED_SCRIPT_TIMEZONE;
    }
  } catch (e) {}

  CACHED_SCRIPT_TIMEZONE = "GMT";
  return CACHED_SCRIPT_TIMEZONE;
}

function resolveClientTimeZone(clientTimeZone) {
  if (!clientTimeZone) return getScriptTimeZone();
  var tz = clientTimeZone.toString().trim();
  if (!tz) return getScriptTimeZone();

  // Basic defensive validation for IANA timezone format.
  if (!/^[A-Za-z_]+(?:\/[A-Za-z0-9_+\-]+)+$/.test(tz)) {
    return getScriptTimeZone();
  }
  return tz;
}

function parseDateToString(val, timeZone) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, timeZone || getScriptTimeZone(), "yyyy-MM-dd");
  }
  return val.toString().trim();
}

var INTERVAL_STAGES = [1 / 24, 1, 3, 7, 14, 30, 60, 90, 180, 360];

// Shared scheduling primitives used by the client-compatible SM-2 model.
function getStageFromInterval(ivl, isNew) {
  if (isNew || ivl < 1) return 1;
  for (var i = 1; i < INTERVAL_STAGES.length; i++) {
    if (ivl <= INTERVAL_STAGES[i]) return i + 1;
  }
  return INTERVAL_STAGES.length;
}

function getIntervalForStage(stage) {
  var index = Math.max(1, Math.min(INTERVAL_STAGES.length, parseInt(stage, 10) || 1)) - 1;
  return INTERVAL_STAGES[index];
}

function cleanString(str) {
  if (!str) return '';
  return str
    .toString()
    .toLowerCase()
    .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()?"'«»]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function countWords(str) {
  if (!str) return 0;
  return str.toString().trim().split(/\s+/).filter(Boolean).length;
}

function getActiveUserId() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email && email.length > 0) return email;

    var effectiveEmail = Session.getEffectiveUser().getEmail();
    if (effectiveEmail && effectiveEmail.length > 0) return effectiveEmail;

    var temporaryUserKey = Session.getTemporaryActiveUserKey();
    if (temporaryUserKey) return 'temporary_user_' + temporaryUserKey;
  } catch (e) {
    console.error("Auth Error: " + e.message);
  }
  throw new Error("Unable to establish the active Google user.");
}

// Mutating card endpoints verify the row still represents the requested card.
function getCardLocation_(sheet, rowIndex, colIdxA, colIdxB, expectedCardId) {
  var r = parseInt(rowIndex, 10);
  var columns = getColumnPair_(sheet, colIdxA, colIdxB);
  var cA = columns.colIdxA;
  var cB = columns.colIdxB;
  var lastRow = sheet.getLastRow();

  if (isNaN(r) || r < 2 || r > lastRow) {
    throw new Error("Invalid card location.");
  }

  var idInfo = ensureCardIdColumn_(sheet);
  var row = sheet.getRange(r, 1, 1, sheet.getMaxColumns()).getValues()[0];
  var front = row[cA] ? row[cA].toString().trim() : "";
  var back = row[cB] ? row[cB].toString().trim() : "";
  var cardId = cleanString(front) + ":::" + cleanString(back);
  var stableId = row[idInfo.idColumn] ? row[idInfo.idColumn].toString().trim() : "";
  if (!front || !back || (expectedCardId && stableId !== expectedCardId)) {
    throw new Error("Card changed or no longer exists.");
  }

  return { rowIndex: r, colIdxA: cA, colIdxB: cB, cardId: stableId, legacyCardId: cardId };
}

function getColumnPair_(sheet, colIdxA, colIdxB) {
  var cA = parseInt(colIdxA, 10);
  var cB = parseInt(colIdxB, 10);
  var maxColumns = sheet.getMaxColumns();
  if (isNaN(cA) || isNaN(cB) || cA < 0 || cB < 0 || cA === cB || cA >= maxColumns || cB >= maxColumns) {
    throw new Error("Invalid language columns.");
  }
  return { colIdxA: cA, colIdxB: cB };
}

// Normalize progress rows once so data loading and batch persistence share the same mapping.
function readProgressForUser_(sheet, userKey, timeZone, todayStr) {
  var progress = {};
  if (!sheet) return progress;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== userKey) continue;
    var cardId = data[i][1];
    progress[cardId] = {
      interval: data[i][2] !== undefined ? parseFloat(data[i][2]) : 1,
      nextReview: data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], timeZone, "yyyy-MM-dd HH:mm:ss") : (data[i][3] ? data[i][3].toString().trim() : todayStr),
      failCount: data[i][4] !== undefined ? parseFloat(data[i][4]) : 0,
      isLeech: data[i][5] === 1,
      lastReviewed: data[i][6] ? parseDateToString(data[i][6], timeZone) : "",
      ef: data[i][7] !== undefined ? parseFloat(data[i][7]) : 2.5,
      prevInterval: data[i][8] !== undefined && data[i][8] !== "" ? parseFloat(data[i][8]) : null
    };
  }
  return progress;
}

function getProgressId_(cardId, colIdxA, colIdxB) {
  return cardId + ":::pair:" + colIdxA + ":" + colIdxB;
}

function ensureCardIdColumn_(sheet) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var idColumn = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim() === CARD_ID_HEADER) {
      idColumn = i;
      break;
    }
  }

  if (idColumn === -1) {
    idColumn = lastColumn;
    sheet.getRange(1, idColumn + 1).setValue(CARD_ID_HEADER);
    sheet.hideColumns(idColumn + 1);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { idColumn: idColumn, ids: [] };

  var idRange = sheet.getRange(2, idColumn + 1, lastRow - 1, 1);
  var idValues = idRange.getValues();
  var changed = false;
  var ids = [];
  for (var row = 0; row < idValues.length; row++) {
    var id = idValues[row][0] ? idValues[row][0].toString().trim() : "";
    if (!id) {
      id = Utilities.getUuid();
      idValues[row][0] = id;
      changed = true;
    }
    ids.push(id);
  }
  if (changed) idRange.setValues(idValues);
  return { idColumn: idColumn, ids: ids };
}

function enforceRateLimit_(name, limit, windowSeconds) {
  var cache = CacheService.getUserCache();
  var key = "srs_rate_" + name;
  var now = Date.now();
  var state = {};
  try {
    state = JSON.parse(cache.get(key) || "{}");
  } catch (e) {
    state = {};
  }

  if (!state.windowStart || now - state.windowStart >= windowSeconds * 1000) {
    state = { windowStart: now, count: 0 };
  }
  if (state.count >= limit) throw new Error("Rate limit exceeded.");
  state.count++;
  cache.put(key, JSON.stringify(state), windowSeconds);
}

// External API endpoints accept short language identifiers only.
function isValidLanguageCode_(lang) {
  return !lang || /^[a-z]{2,3}(?:-[A-Z]{2})?$/i.test(lang.toString().trim());
}

function getTodayString(timeZone) {
  return Utilities.formatDate(new Date(), timeZone || getScriptTimeZone(), "yyyy-MM-dd");
}

function getFlashcardData(customDeckId, clientTodayStr, clientTimeZone) {
  var activeUserKey = getActiveUserId();
  var timeZone = resolveClientTimeZone(clientTimeZone);
  var todayStr = getTodayString(timeZone);

  var ss = resolveSpreadsheet(customDeckId);
  var ssName = ss.getName();
  var activeDeckId = ss.getId();
  var baseUrl = "";
  try {
    baseUrl = PropertiesService.getScriptProperties().getProperty('SRS_WEBAPP_URL') || "";
  } catch (e) {}
  if (!baseUrl) {
    try {
      baseUrl = ScriptApp.getService().getUrl() || "";
    } catch (e) {}
  }
  var webAppUrl = baseUrl
    ? (baseUrl + "?deck=" + encodeURIComponent(activeDeckId))
    : "";
  var sheet = ss.getSheets()[0];
  var cardIdInfo = ensureCardIdColumn_(sheet);
  var allValues = sheet.getDataRange().getValues();

  var progressSheet = ss.getSheetByName('SRS_Progress');
  var savedProgressRaw = readProgressForUser_(progressSheet, activeUserKey, timeZone, todayStr);

  var studyHistory = {};
  var studiedLog = { date: todayStr, newCount: 0, oldCount: 0 };
  var historySheet = ss.getSheetByName('SRS_History');
  if (historySheet) {
    var histData = historySheet.getDataRange().getValues();
    for (var j = 1; j < histData.length; j++) {
      if (histData[j][0] === activeUserKey) {
        var hDateStr = parseDateToString(histData[j][1], timeZone);
        var hNew = parseInt(histData[j][2], 10) || 0;
        var hOld = parseInt(histData[j][3], 10) || 0;
        var hTotal = parseInt(histData[j][4], 10) || 0;

        studyHistory[hDateStr] = hTotal;
        if (hDateStr === todayStr) studiedLog = { date: todayStr, newCount: hNew, oldCount: hOld };
      }
    }
  }

  if (allValues.length < 2 || allValues[0].length < 2) {
    return {
      headers: [], rawRows: [], savedProgress: {},
      studiedToday: { newCount: 0, oldCount: 0 },
      activeUserKey: activeUserKey,
      todayStr: todayStr,
      studyHistory: studyHistory, ssName: ssName, activeDeckId: activeDeckId,
      webAppUrl: webAppUrl
    };
  }

  var headers = allValues[0].map(function (h) { return h.toString().trim(); });
  var rawData = allValues.slice(1);
  var savedProgress = {};

  rawData.forEach(function (row, rowIndex) {
    var front = row[0] ? row[0].toString().trim() : "";
    for (var c = 1; c < row.length; c++) {
      var back = row[c] ? row[c].toString().trim() : "";
      if (!front || !back || c === cardIdInfo.idColumn) continue;
      var legacyId = cleanString(front) + ":::" + cleanString(back);
      var stableId = getProgressId_(cardIdInfo.ids[rowIndex], 0, c);
      if (savedProgressRaw[stableId]) savedProgress[stableId] = savedProgressRaw[stableId];
      if (savedProgressRaw[legacyId] && !savedProgress[stableId]) savedProgress[stableId] = savedProgressRaw[legacyId];
    }
  });

  return {
    headers: headers, rawRows: rawData, savedProgress: savedProgress,
    cardIds: cardIdInfo.ids, cardIdColumn: cardIdInfo.idColumn,
    studiedToday: { newCount: studiedLog.newCount || 0, oldCount: studiedLog.oldCount || 0 },
    activeUserKey: activeUserKey,
    todayStr: todayStr,
    studyHistory: studyHistory, ssName: ssName, activeDeckId: activeDeckId,
    webAppUrl: webAppUrl
  };
}

// Translation providers intentionally return partial results so one missing
// provider does not block the built-in Google translation path.
function translateText(text, sourceLang, targetLang) {
  try {
    enforceRateLimit_("translate", 20, 60);
    var cleanPrompt = text ? text.toString().trim() : '';
    if (!cleanPrompt) return { success: false, error: "Empty text prompt" };
    if (cleanPrompt.length > 500) return { success: false, error: "Source text exceeds the character limit." };
    if (countWords(cleanPrompt) > 20) return { success: false, error: "Source text exceeds the limit of 20 words." };
    if (!isValidLanguageCode_(sourceLang) || !isValidLanguageCode_(targetLang)) return { success: false, error: "Invalid language code." };

    var googleTrans = "";
    try { googleTrans = LanguageApp.translate(cleanPrompt, sourceLang || '', targetLang || 'en'); }
    catch (e) { googleTrans = "Google Translation unavailable."; }

    var geminiTrans = "";
    try {
      var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
      if (!apiKey) {
        geminiTrans = "Gemini Error: Script Property 'GEMINI_API_KEY' is missing or empty.";
      } else {
        // Fallback Stack: Cascades to the next model if one times out or errors
        var models = GEMINI_MODELS;
        
        var promptStr = "Translate the following text directly from " + (sourceLang || "auto") + " to " + (targetLang || "en") + ". Output ONLY the raw translated string without quote marks or conversational filler: " + cleanPrompt;
        var payload = { "contents": [{ "parts": [{ "text": promptStr }] }] };
        var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
        
        var success = false;
        var lastError = "";

        for (var i = 0; i < models.length; i++) {
          var url = "https://generativelanguage.googleapis.com/v1beta/models/" + models[i] + ":generateContent?key=" + apiKey.trim();
          var response = UrlFetchApp.fetch(url, options);
          var responseCode = response.getResponseCode();
          var json = JSON.parse(response.getContentText());

          if (responseCode === 200 && json && json.candidates && json.candidates[0] && json.candidates[0].content) {
            geminiTrans = json.candidates[0].content.parts[0].text.trim();
            success = true;
            break; // Stop looping once a model succeeds
          } else if (responseCode === 429 || responseCode >= 500) {
            // Rate limit (15 RPM) or server error - save error but try the next model
            lastError = json && json.error ? json.error.message : "HTTP " + responseCode;
            continue; 
          } else {
            // Other errors (400 Bad Request, invalid key, 404 Not Found) - No point in retrying
            geminiTrans = "Gemini Error (" + responseCode + "): " + (json && json.error ? json.error.message : "Unknown");
            success = true; // Mark as "handled" to avoid the fallback error message
            break;
          }
        }
        
        if (!success) {
           geminiTrans = "Gemini Error (All models failed): " + lastError;
        }
      }
    } catch (e) { geminiTrans = "Gemini Exception: " + e.toString(); }

    var deeplTrans = "";
    try {
      var deeplKey = PropertiesService.getScriptProperties().getProperty('DEEPL_API_KEY');
      if (!deeplKey) {
        deeplTrans = "DeepL Error: Script Property 'DEEPL_API_KEY' is missing or empty.";
      } else {
        var trimmedDeepLKey = deeplKey.trim();
        var isFreeDeepL = trimmedDeepLKey.endsWith(':fx');
        var deeplUrl = isFreeDeepL ? "https://api-free.deepl.com/v2/translate" : "https://api.deepl.com/v2/translate";
        
        var srcDeepL = sourceLang ? sourceLang.toUpperCase() : null;
        if (srcDeepL === 'NO') srcDeepL = 'NB';
        var tgtDeepL = targetLang ? targetLang.toUpperCase() : 'EN';
        if (tgtDeepL === 'NO') tgtDeepL = 'NB';

        var deeplPayload = {
          "text": [cleanPrompt],
          "target_lang": tgtDeepL
        };
        if (srcDeepL && srcDeepL !== 'AUTO') {
          deeplPayload["source_lang"] = srcDeepL;
        }
        
        var deeplOptions = {
          "method": "post",
          "contentType": "application/json",
          "headers": { "Authorization": "DeepL-Auth-Key " + trimmedDeepLKey },
          "payload": JSON.stringify(deeplPayload),
          "muteHttpExceptions": true
        };
        
        var deeplRes = UrlFetchApp.fetch(deeplUrl, deeplOptions);
        var dResponseCode = deeplRes.getResponseCode();
        var dText = deeplRes.getContentText();
        
        try {
          if (!dText) throw new Error("Empty response");
          var dJson = JSON.parse(dText);
          if (dResponseCode === 200 && dJson.translations && dJson.translations[0]) {
            deeplTrans = dJson.translations[0].text.trim();
          } else if (dJson.message) {
            deeplTrans = "DeepL Error (" + dResponseCode + "): " + dJson.message;
          } else {
            deeplTrans = "DeepL Error: HTTP " + dResponseCode;
          }
        } catch (jsonErr) {
          deeplTrans = "DeepL Error (" + dResponseCode + "): Invalid JSON payload.";
        }
      }
    } catch (e) { deeplTrans = "DeepL Exception: " + e.toString(); }

    return { success: true, google: googleTrans, gemini: geminiTrans, deepl: deeplTrans };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function addNewCardToSheet(frontText, backText, customDeckId, colIdxA, colIdxB) {
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    var ss = resolveSpreadsheet(customDeckId);
    var sheet = ss.getSheets()[0];
    var cardIdInfo = ensureCardIdColumn_(sheet);
    var f = preventInjection(frontText ? frontText.toString().trim() : '');
    var b = preventInjection(backText ? backText.toString().trim() : '');
    var cA = parseInt(colIdxA, 10);
    var cB = parseInt(colIdxB, 10);

    if (isNaN(cA) || cA < 0) cA = 0;
    if (isNaN(cB) || cB < 0) cB = 1;
    if (cA === cB || cA >= sheet.getMaxColumns() || cB >= sheet.getMaxColumns()) {
      return { success: false, error: "Invalid column index." };
    }

    if (!f || !b) return { success: false, error: "Both front and back required." };
    if (countWords(f) > 20 || countWords(b) > 20) return { success: false, error: "Text exceeds maximum 20 words limit." };

    var width = Math.max(sheet.getLastColumn(), cA + 1, cB + 1);
    var rowData = [];
    for (var i = 0; i < width; i++) rowData.push('');
    rowData[cA] = f;
    rowData[cB] = b;
    rowData[cardIdInfo.idColumn] = Utilities.getUuid();

    sheet.appendRow(rowData);
    return { success: true, rowIndex: sheet.getLastRow(), front: f, back: b };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// Card edits migrate the current user's progress only after row identity passes.
function updateCardInSheet(rowIndex, colIdxA, colIdxB, newFront, newBack, customDeckId, oldCardId) {
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    var ss = resolveSpreadsheet(customDeckId);
    var sheet = ss.getSheets()[0];
    if (!oldCardId) return { success: false, error: "Missing card identity." };
    var location = getCardLocation_(sheet, rowIndex, colIdxA, colIdxB, oldCardId);
    var r = location.rowIndex;
    var cA = location.colIdxA;
    var cB = location.colIdxB;

    var cleanF = preventInjection(newFront ? newFront.toString().trim() : '');
    var cleanB = preventInjection(newBack ? newBack.toString().trim() : '');
    if (!cleanF || !cleanB) return { success: false, error: "Both front and back required." };
    if (countWords(cleanF) > 20 || countWords(cleanB) > 20) return { success: false, error: "Text exceeds maximum 20 words limit." };

    var oldF = sheet.getRange(r, cA + 1).getValue();
    var oldB = sheet.getRange(r, cB + 1).getValue();
    try {
      sheet.getRange(r, cA + 1).setValue(cleanF);
      sheet.getRange(r, cB + 1).setValue(cleanB);
    } catch (writeError) {
      sheet.getRange(r, cA + 1).setValue(oldF);
      sheet.getRange(r, cB + 1).setValue(oldB);
      throw writeError;
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveProgressToSheet_(activeDeckId, uKey, progressMap, lockAlreadyHeld) {
  var lock = null;
  try {
    if (!lockAlreadyHeld) {
      lock = LockService.getDocumentLock();
      lock.waitLock(15000);
    }
    var ss = SpreadsheetApp.openById(activeDeckId);
    var sheet = ss.getSheetByName('SRS_Progress');

    if (!sheet) {
      sheet = ss.insertSheet('SRS_Progress');
      sheet.appendRow(['UserKey', 'CardId', 'Interval', 'NextReview', 'FailCount', 'IsLeech', 'LastReviewed', 'EF', 'PrevInterval']);
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0] || ['UserKey', 'CardId', 'Interval', 'NextReview', 'FailCount', 'IsLeech', 'LastReviewed', 'EF', 'PrevInterval'];
    if (headers.length < 9) {
      headers = headers.slice();
      while (headers.length < 9) headers.push(headers.length === 8 ? 'PrevInterval' : '');
    }
    var newData = [headers];
    var userProgressByCardId = {};

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === uKey) {
        userProgressByCardId[data[i][1]] = data[i];
      } else {
        newData.push(data[i]);
      }
    }

    for (var cardId in progressMap) {
      var p = progressMap[cardId];
      userProgressByCardId[cardId] = [uKey, cardId, p.interval, "'" + p.nextReview, p.failCount, p.isLeech ? 1 : 0, p.lastReviewed ? "'" + p.lastReviewed : "", p.ef, p.prevInterval || ""];
    }

    for (var cardId in userProgressByCardId) {
      newData.push(userProgressByCardId[cardId]);
    }

    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
    var lastRow = sheet.getLastRow();
    if (lastRow > newData.length) {
      sheet.getRange(newData.length + 1, 1, lastRow - newData.length, newData[0].length).clearContent();
    }
  } catch (e) {
    console.error("Failed to save progress: " + e.message);
  } finally {
    if (lock) lock.releaseLock();
  }
}

// History is aggregated once per user and local calendar day.
function saveHistoryToSheet_(activeDeckId, uKey, todayStr, newCardDelta, oldCardDelta, lockAlreadyHeld, timeZone) {
  var lock = null;
  try {
    if (!lockAlreadyHeld) {
      lock = LockService.getDocumentLock();
      lock.waitLock(15000);
    }
    var ss = SpreadsheetApp.openById(activeDeckId);
    var sheet = ss.getSheetByName('SRS_History');
    var tz = timeZone || getScriptTimeZone();

    if (!sheet) {
      sheet = ss.insertSheet('SRS_History');
      sheet.appendRow(['UserKey', 'Date', 'NewCount', 'OldCount', 'TotalReviews']);
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      var cellDateStr = parseDateToString(data[i][1], tz);
      if (data[i][0] === uKey && cellDateStr === todayStr) {
        rowIndex = i + 1;
        break;
      }
    }

    var totalDelta = newCardDelta + oldCardDelta;
    if (rowIndex > -1) {
      var currentNew = parseInt(data[rowIndex - 1][2] || 0, 10);
      var currentOld = parseInt(data[rowIndex - 1][3] || 0, 10);
      var currentTotal = parseInt(data[rowIndex - 1][4] || 0, 10);
      sheet.getRange(rowIndex, 3, 1, 3).setValues([[currentNew + newCardDelta, currentOld + oldCardDelta, currentTotal + totalDelta]]);
    } else {
      sheet.appendRow([uKey, "'" + todayStr, newCardDelta, oldCardDelta, totalDelta]);
    }
  } catch (e) {
    console.error("Failed to save history: " + e.message);
  } finally {
    if (lock) lock.releaseLock();
  }
}

// Batch results are validated against live sheet card IDs before persistence.
function saveSegmentResults(batchArray, customDeckId, clientTodayStr, clientTimeZone, colIdxA, colIdxB) {
  if (!batchArray || batchArray.length === 0) return { success: true };
  if (!Array.isArray(batchArray) || batchArray.length > 50) return { success: false, error: "Invalid result batch." };

  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);

    var ss = resolveSpreadsheet(customDeckId);
    var activeDeckId = ss.getId();
    var mainSheet = ss.getSheets()[0];
    var columns;
    try {
      columns = getColumnPair_(mainSheet, colIdxA, colIdxB);
    } catch (columnError) {
      return { success: false, error: columnError.message };
    }
    var selectedColA = columns.colIdxA;
    var selectedColB = columns.colIdxB;
    var cardIdInfo = ensureCardIdColumn_(mainSheet);
    var validCardIds = {};
    var mainData = mainSheet.getDataRange().getValues();
    for (var rowIndex = 1; rowIndex < mainData.length; rowIndex++) {
      var front = mainData[rowIndex][selectedColA] ? mainData[rowIndex][selectedColA].toString().trim() : "";
      var back = mainData[rowIndex][selectedColB] ? mainData[rowIndex][selectedColB].toString().trim() : "";
      if (front && back) {
        var stableId = mainData[rowIndex][cardIdInfo.idColumn];
        validCardIds[getProgressId_(stableId, selectedColA, selectedColB)] = true;
      }
    }
    var timeZone = resolveClientTimeZone(clientTimeZone);
    var todayStr = getTodayString(timeZone);
    var uKey = getActiveUserId();
    var progressSheet = ss.getSheetByName('SRS_Progress');
    var savedProgress = readProgressForUser_(progressSheet, uKey, timeZone, todayStr);
    for (var row = 1; row < mainData.length; row++) {
      var rowFront = mainData[row][selectedColA] ? mainData[row][selectedColA].toString().trim() : "";
      var rowBack = mainData[row][selectedColB] ? mainData[row][selectedColB].toString().trim() : "";
      var rowStableId = mainData[row][cardIdInfo.idColumn] ? mainData[row][cardIdInfo.idColumn].toString().trim() : "";
      var rowLegacyId = cleanString(rowFront) + ":::" + cleanString(rowBack);
      var rowProgressId = getProgressId_(rowStableId, selectedColA, selectedColB);
      if (rowStableId && rowFront && rowBack && !savedProgress[rowProgressId] && savedProgress[rowLegacyId]) {
        savedProgress[rowProgressId] = savedProgress[rowLegacyId];
      }
    }

    var newDelta = 0;
    var oldDelta = 0;

    batchArray.forEach(function (item) {
      if (!item || !item.cardId || !validCardIds[item.cardId]) throw new Error("Invalid study result.");
      var cardId = item.cardId.toString();
      var wasNew = !savedProgress[cardId];

      var cardState = savedProgress[cardId] || { interval: 1, failCount: 0, isLeech: false, ef: 2.5, prevInterval: null };
      var newInterval = item.newInterval !== undefined && item.newInterval !== null ? parseFloat(item.newInterval) : (parseFloat(cardState.interval) || 1);
      if (!isFinite(newInterval) || newInterval <= 0) newInterval = parseFloat(cardState.interval) || 1;
      newInterval = Math.min(newInterval, 36500);

      var newEF = item.newEF !== undefined && item.newEF !== null ? parseFloat(item.newEF) : (parseFloat(cardState.ef) || 2.5);
      if (!isFinite(newEF) || newEF < 1.3) newEF = Math.max(1.3, parseFloat(cardState.ef) || 2.5);
      newEF = Math.min(newEF, 10);

      var failCount = item.failCount !== undefined && item.failCount !== null ? parseFloat(item.failCount) : (parseFloat(cardState.failCount) || 0);
      if (!isFinite(failCount) || failCount < 0) failCount = parseFloat(cardState.failCount) || 0;
      failCount = Math.min(failCount, 100);

      var isLeech = item.isLeech !== undefined ? !!item.isLeech : !!cardState.isLeech;
      var prevInterval = item.prevInterval !== undefined ? item.prevInterval : cardState.prevInterval;
      if (prevInterval !== null && prevInterval !== "" && prevInterval !== undefined) {
        var parsedPrev = parseFloat(prevInterval);
        prevInterval = isFinite(parsedPrev) ? parsedPrev : null;
      } else {
        prevInterval = null;
      }

      var lastReviewed = todayStr;

      var nextDateStr = item.nextReview ? item.nextReview.toString().trim() : "";
      if (!nextDateStr) {
        nextDateStr = cardState.nextReview || todayStr;
      }
      if (isNaN(Date.parse(nextDateStr))) nextDateStr = todayStr;

      savedProgress[cardId] = {
        interval: newInterval,
        nextReview: nextDateStr,
        failCount: failCount,
        isLeech: isLeech,
        lastReviewed: lastReviewed,
        ef: newEF,
        prevInterval: prevInterval
      };

      if (wasNew) newDelta++;
      else oldDelta++;
    });

    saveProgressToSheet_(activeDeckId, uKey, savedProgress, true);
    saveHistoryToSheet_(activeDeckId, uKey, todayStr, newDelta, oldDelta, true, timeZone);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// Archive operations are idempotent and restore progress if deletion fails.
function archiveMasteredCards(customDeckId, colIdxA, colIdxB) {
  var lock = LockService.getDocumentLock();
  var progressCleared = false;
  var progressSheet = null;
  var originalProgressData = null;
  try {
    lock.waitLock(15000);
    var ss = resolveSpreadsheet(customDeckId);
    var mainSheet = ss.getSheets()[0];
    var activeUserKey = getActiveUserId();
    var columns;
    try {
      columns = getColumnPair_(mainSheet, colIdxA, colIdxB);
    } catch (columnError) {
      return { success: false, message: columnError.message };
    }
    var selectedColA = columns.colIdxA;
    var selectedColB = columns.colIdxB;
    var cardIdInfo = ensureCardIdColumn_(mainSheet);
    var progSheet = ss.getSheetByName('SRS_Progress');
    progressSheet = progSheet;
    var archiveSheet = ss.getSheetByName('Mastered_Cards');

    if (!progSheet) return { success: false, message: "No progress found." };
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet('Mastered_Cards');
      archiveSheet.appendRow(mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0]);
    }

    var progData = progSheet.getDataRange().getValues();
    originalProgressData = progData;
    var mainData = mainSheet.getDataRange().getValues();
    var masteredIds = new Set();

    for (var i = 1; i < progData.length; i++) {
      if (progData[i][0] === activeUserKey && parseFloat(progData[i][2]) >= 360) {
        masteredIds.add(progData[i][1]);
      }
    }

    if (masteredIds.size === 0) return { success: true, count: 0, message: "No fully mastered (Interval 360+ days) cards to archive." };

    var rowsToDelete = [];
    var rowsToArchive = [];
    var archivedCardIds = {};
    var archiveIdInfo = ensureCardIdColumn_(archiveSheet);
    var archiveData = archiveSheet.getDataRange().getValues();
    for (var a = 1; a < archiveData.length; a++) {
      var archivedId = archiveData[a][archiveIdInfo.idColumn] ? archiveData[a][archiveIdInfo.idColumn].toString().trim() : "";
      if (archivedId) archivedCardIds[archivedId] = true;
    }

    for (var r = mainData.length - 1; r >= 1; r--) {
      var front = mainData[r][selectedColA] ? mainData[r][selectedColA].toString().trim() : "";
      var back = mainData[r][selectedColB] ? mainData[r][selectedColB].toString().trim() : "";
      var stableId = mainData[r][cardIdInfo.idColumn] ? mainData[r][cardIdInfo.idColumn].toString().trim() : "";
      var cardId = getProgressId_(stableId, selectedColA, selectedColB);

      if (masteredIds.has(cardId)) {
        rowsToDelete.push(r + 1);
        if (!archivedCardIds[stableId]) rowsToArchive.push(mainData[r]);
      }
    }

    if (rowsToArchive.length > 0) {
      archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    }

    if (rowsToDelete.length > 0 || rowsToArchive.length > 0) {
      var remainingProgress = [progData[0]];
      for (var p = 1; p < progData.length; p++) {
        if (!(progData[p][0] === activeUserKey && masteredIds.has(progData[p][1]))) {
          remainingProgress.push(progData[p]);
        }
      }
      progSheet.clearContents();
      progSheet.getRange(1, 1, remainingProgress.length, remainingProgress[0].length).setValues(remainingProgress);
      progressCleared = true;
    }

    if (rowsToDelete.length > 0) {
      rowsToDelete.forEach(function (rowNum) { mainSheet.deleteRow(rowNum); });
    }

    return { success: true, count: rowsToArchive.length, message: "Archived " + rowsToArchive.length + " mastered cards." };
  } catch (e) {
    if (progressCleared && progressSheet && originalProgressData && originalProgressData.length > 0) {
      try {
        progressSheet.clearContents();
        progressSheet.getRange(1, 1, originalProgressData.length, originalProgressData[0].length).setValues(originalProgressData);
      } catch (restoreError) {}
    }
    return { success: false, message: "System busy. Please try again later." };
  } finally {
    lock.releaseLock();
  }
}

// Audio and spelling endpoints are separately rate-limited because they call
// external services and can be invoked directly by a web-app client.
function getAudioBase64(text, lang) {
  try {
    enforceRateLimit_("audio", 60, 60);
    var cleanText = text ? text.toString().trim() : "";
    var cleanLang = lang ? lang.toString().trim() : "";
    if (!cleanText || cleanText.length > 500 || countWords(cleanText) > 20 || !isValidLanguageCode_(cleanLang)) {
      return { success: false, error: "Invalid audio request." };
    }
    var url = "https://translate.google.com/translate_tts?ie=UTF-8&q=" + encodeURIComponent(cleanText) + "&tl=" + encodeURIComponent(cleanLang) + "&client=tw-ob";
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      var blob = response.getBlob();
      return { success: true, data: "data:audio/mp3;base64," + Utilities.base64Encode(blob.getBytes()) };
    }
    return { success: false, error: "HTTP " + response.getResponseCode() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function checkSpelling(text, lang) {
  try {
    enforceRateLimit_("spellcheck", 30, 60);
    var cleanPrompt = text ? text.toString().trim() : '';
    if (!cleanPrompt || cleanPrompt.length > 500 || countWords(cleanPrompt) > 20) return { success: false };
    if (!isValidLanguageCode_(lang)) return { success: false };

    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return { success: false };

    var models = GEMINI_MODELS;
    
    // Strict prompt to ensure we only get 'OK' or the corrected text back
    var promptStr = "You are a strict spellchecker for the language code '" + lang + "'. If the following text contains typos, output ONLY the corrected text without any quotes or explanations. If it is perfectly spelled, output exactly the word 'OK'. For Norwegian input (no, nb, nn): if the text is a noun or noun phrase missing the indefinite article, suggest the corrected form with the proper article prepended (en, ei, or et). Text: " + cleanPrompt;
    var payload = { "contents": [{ "parts": [{ "text": promptStr }] }], "generationConfig": { "temperature": 0.1 } };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    
    for (var i = 0; i < models.length; i++) {
      var url = "https://generativelanguage.googleapis.com/v1beta/models/" + models[i] + ":generateContent?key=" + apiKey.trim();
      var response = UrlFetchApp.fetch(url, options);
      var responseCode = response.getResponseCode();
      var json = JSON.parse(response.getContentText());

      if (responseCode === 200 && json && json.candidates && json.candidates[0]) {
        var result = json.candidates[0].content.parts[0].text.trim();
        // Ignore case changes or perfect spelling
        if (result === 'OK' || result.toLowerCase() === cleanPrompt.toLowerCase()) {
           return { success: true, corrected: 'OK' };
        }
        return { success: true, corrected: result };
      }
    }
    return { success: false };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Remove a single verified row while preserving retryability on partial failure.
function archiveSingleCard(rowIndex, customDeckId, cardId, colIdxA, colIdxB) {
  var lock = LockService.getDocumentLock();
  var progressCleared = false;
  var progressSheet = null;
  var originalProgressData = null;
  try {
    lock.waitLock(10000);
    var ss = resolveSpreadsheet(customDeckId);
    var mainSheet = ss.getSheets()[0];
    var archiveSheet = ss.getSheetByName('Mastered_Cards');
    if (!cardId) return { success: false, error: "Missing card identity." };
    
    // Create Mastered_Cards sheet if it doesn't exist
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet('Mastered_Cards');
      archiveSheet.appendRow(mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0]);
    }
    
    var location = getCardLocation_(mainSheet, rowIndex, colIdxA, colIdxB, cardId);
    var r = location.rowIndex;
    var progressId = getProgressId_(cardId, location.colIdxA, location.colIdxB);
    var archiveIdInfo = ensureCardIdColumn_(archiveSheet);
    
    // Copy to archive, then delete from main
    var rowData = mainSheet.getRange(r, 1, 1, mainSheet.getLastColumn()).getValues();
    var archiveData = archiveSheet.getDataRange().getValues();
    var alreadyArchived = false;
    for (var a = 1; a < archiveData.length; a++) {
      var archivedId = archiveData[a][archiveIdInfo.idColumn] ? archiveData[a][archiveIdInfo.idColumn].toString().trim() : "";
      if (archivedId === cardId) {
        alreadyArchived = true;
        break;
      }
    }
    if (!alreadyArchived) archiveSheet.appendRow(rowData[0]);
    progressSheet = ss.getSheetByName('SRS_Progress');
    var activeUserKey = getActiveUserId();
    if (progressSheet && cardId) {
      var progressData = progressSheet.getDataRange().getValues();
      originalProgressData = progressData;
      var remainingProgress = [progressData[0]];
      for (var i = 1; i < progressData.length; i++) {
        if (!(progressData[i][0] === activeUserKey && progressData[i][1] === progressId)) {
          remainingProgress.push(progressData[i]);
        }
      }
      progressSheet.clearContents();
      progressSheet.getRange(1, 1, remainingProgress.length, remainingProgress[0].length).setValues(remainingProgress);
      progressCleared = true;
    }

    mainSheet.deleteRow(r);
    
    return { success: true };
  } catch (e) {
    if (progressCleared && progressSheet && originalProgressData && originalProgressData.length > 0) {
      try {
        progressSheet.clearContents();
        progressSheet.getRange(1, 1, originalProgressData.length, originalProgressData[0].length).setValues(originalProgressData);
      } catch (restoreError) {}
    }
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}