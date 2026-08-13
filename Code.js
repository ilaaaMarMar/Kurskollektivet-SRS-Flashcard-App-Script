var TEMPLATE_ID = "1s04BKwYk-s0WVSmsJOfCxCz4X-fqdZzR0O65ANymu9E";
var GEMINI_MODELS = [
  "gemini-3.6-flash",
  "gemini-3.5-flash",
  "gemini-3.1-flash-lite",
  "gemini-2.5-flash"
];
var CARD_ID_HEADER = "SRS Card ID";
var MAX_INTERVAL_DAYS = 360;
var ARCHIVE_SHEET_NAME = 'Archived cards';
var REVIEW_EVENT_HEADERS = ['UserKey', 'ClientId', 'CardId', 'IsCorrect', 'Confidence', 'WasNew', 'IsTypo', 'InDifficultMode', 'LastReviewed', 'ResultJson', 'Status', 'AppliedAt'];
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
function applyServerFuzz_(intervalInDays) {
  var learningFuzz = Math.min(5 / 1440, Math.max(1 / 1440, intervalInDays * 0.1));
  var fuzzRange = intervalInDays < 1 ? learningFuzz : intervalInDays * 0.05;
  var fuzz = (Math.random() * (fuzzRange * 2)) - fuzzRange;
  return Math.max(0.01, intervalInDays + fuzz);
}
function getBoundedOverdueDays_(nextReview, currentInterval, nowMs) {
  if (!nextReview || currentInterval < 1) return 0;
  var dueMs = Date.parse(nextReview.toString().trim());
  if (isNaN(dueMs)) return 0;
  var overdueDays = Math.max(0, (nowMs - dueMs) / (24 * 60 * 60 * 1000));
  return Math.min(overdueDays, currentInterval * 0.5);
}
// Recompute scheduling from trusted progress and the answer outcome. Client
// interval fields are treated as display hints only and are never persisted.
function scheduleAnswerResult_(cardState, item, wasNew, nowMs, todayStr) {
  var currentInterval = parseFloat(cardState.interval) || 1;
  var currentEF = parseFloat(cardState.ef) || 2.5;
  var currentPrevInterval = cardState.prevInterval || null;
  var currentIsLeech = !!cardState.isLeech;
  var isCorrect = item.isCorrect === true;
  var isTypo = item.isTypo === true;
  var confidence = ['again', 'hard', 'good', 'easy'].indexOf(item.confidence) >= 0 ? item.confidence : (isCorrect ? 'good' : 'again');
  var q = 4;
  if (!isCorrect || confidence === 'again') q = 1;
  else if (isTypo || confidence === 'hard') q = 3;
  else if (confidence === 'good') q = 4;
  else if (confidence === 'easy') q = 5;
  var failCount = parseFloat(cardState.failCount) || 0;
  if (!isCorrect) failCount += 1;
  else if (isTypo) failCount += 0.5;
  var newInterval = 1;
  var newEF = currentEF;
  var inDifficultMode = item.inDifficultMode === true;
  if (inDifficultMode && isCorrect) {
    failCount = 0;
    var baseInterval = currentPrevInterval || currentInterval;
    var baseStage = getStageFromInterval(baseInterval, wasNew);
    var targetStage = Math.max(1, baseStage - 1);
    newInterval = getIntervalForStage(targetStage);
    newEF = Math.max(1.3, currentEF - 0.1);
    currentPrevInterval = null;
  } else {
    if (q < 3) {
      newInterval = 15 / 1440;
    } else if (wasNew) {
      if (q === 3) newInterval = 15 / 1440;
      else if (q === 4) newInterval = 0.25;
      else if (q === 5) newInterval = 1;
    } else if (currentInterval < 1) {
      if (q === 5) newInterval = currentInterval >= 0.5 ? 3 : 1;
      else if (q === 3) newInterval = currentInterval;
      else newInterval = currentInterval <= 0.26 ? 0.5 : 1;
    } else {
      var overdueDays = getBoundedOverdueDays_(cardState.nextReview, currentInterval, nowMs);
      var effectiveInterval = currentInterval + overdueDays;
      if (q === 3) newInterval = effectiveInterval * 1.2;
      else if (q === 4) newInterval = effectiveInterval * currentEF;
      else if (q === 5) newInterval = effectiveInterval * currentEF * 1.3;
    }
    if (!currentIsLeech && failCount >= 1) currentPrevInterval = currentInterval;
    newEF = currentEF + (0.1 - (5 - q) * (0.08 + (5 - q) * 0.02));
    if (newEF < 1.3) newEF = 1.3;
  }
  var finalInterval = Math.min(MAX_INTERVAL_DAYS, applyServerFuzz_(newInterval));
  var nextReview = new Date(nowMs + finalInterval * 24 * 60 * 60 * 1000).toISOString();
  return {
    interval: finalInterval,
    nextReview: nextReview,
    failCount: Math.min(100, Math.max(0, failCount)),
    isLeech: failCount >= 1,
    lastReviewed: todayStr,
    ef: Math.min(10, Math.max(1.3, newEF)),
    prevInterval: currentPrevInterval
  };
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
function getActiveUserEmail_() {
  var email = Session.getActiveUser().getEmail();
  if (email) return email;
  return Session.getEffectiveUser().getEmail() || '';
}
function getStableUserKey_(email) {
  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    ScriptApp.getScriptId() + '|' + email.toString().trim().toLowerCase()
  );
  return 'user_' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/, '');
}
function getActiveUserId() {
  try {
    var email = getActiveUserEmail_();
    if (email && email.length > 0) return getStableUserKey_(email);
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
  var stableId = row[idInfo.idColumn] ? row[idInfo.idColumn].toString().trim() : "";
  var identity = getCardIdentity_(row, cA, cB, stableId);
  var front = identity.front;
  var back = identity.back;
  if (!front || !back || (expectedCardId && stableId !== expectedCardId)) {
    throw new Error("Card changed or no longer exists.");
  }
  return { rowIndex: r, colIdxA: cA, colIdxB: cB, cardId: stableId };
}
function getCardIdentity_(row, colIdxA, colIdxB, stableId) {
  var front = row[colIdxA] ? row[colIdxA].toString().trim() : "";
  var back = row[colIdxB] ? row[colIdxB].toString().trim() : "";
  return {
    front: front,
    back: back,
    stableId: stableId ? stableId.toString().trim() : "",
    progressId: getProgressId_(stableId, colIdxA, colIdxB)
  };
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
function normalizeDateColumn_(sheet, column, startRow, lastRow, includeTime) {
  if (lastRow < startRow) return;
  var range = sheet.getRange(startRow, column, lastRow - startRow + 1, 1);
  var values = range.getValues();
  var changed = false;
  values.forEach(function (row) {
    if (row[0] instanceof Date || !row[0]) return;
    var raw = row[0].toString().trim().replace(/^'/, '');
    var parsed = new Date(raw.replace(' ', 'T') + (/[zZ]|[+\-]\d{2}:?\d{2}$/.test(raw) ? '' : 'Z'));
    if (!isNaN(parsed.getTime())) {
      row[0] = parsed;
      changed = true;
    }
  });
  if (changed) range.setValues(values);
  range.setNumberFormat(includeTime ? 'yyyy-mm-dd HH:mm' : 'yyyy-mm-dd');
}
function formatTrackingSheet_(sheet, type) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var lastColumn = Math.max(sheet.getLastColumn(), type === 'progress' ? 9 : 5);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastColumn)
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  if (type === 'progress') {
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 360);
    sheet.setColumnWidth(3, 90);
    sheet.setColumnWidth(4, 185);
    sheet.setColumnWidth(5, 85);
    sheet.setColumnWidth(6, 75);
    sheet.setColumnWidth(7, 110);
    sheet.setColumnWidth(8, 70);
    sheet.setColumnWidth(9, 95);
    if (lastRow > 1) {
      normalizeDateColumn_(sheet, 4, 2, lastRow, true);
      normalizeDateColumn_(sheet, 7, 2, lastRow, false);
      sheet.getRange(2, 3, lastRow - 1, 1).setNumberFormat('0.000');
      sheet.getRange(2, 5, lastRow - 1, 1).setNumberFormat('0.0');
      sheet.getRange(2, 6, lastRow - 1, 1).setNumberFormat('0');
      sheet.getRange(2, 8, lastRow - 1, 1).setNumberFormat('0.00');
    }
  } else {
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 105);
    sheet.setColumnWidth(3, 95);
    sheet.setColumnWidth(4, 95);
    sheet.setColumnWidth(5, 115);
    if (lastRow > 1) {
      normalizeDateColumn_(sheet, 2, 2, lastRow, false);
      sheet.getRange(2, 3, lastRow - 1, 3).setNumberFormat('0');
    }
  }
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
  if (idColumn === -1) throw new Error("Missing SRS Card ID column.");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { idColumn: idColumn, ids: [] };
  var idRange = sheet.getRange(2, idColumn + 1, lastRow - 1, 1);
  var ids = idRange.getValues().map(function (row) {
    var id = row[0] ? row[0].toString().trim() : "";
    if (!id) throw new Error("Blank SRS Card ID.");
    return id;
  });
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
function isValidLanguageCode_(lang) {
  return !lang || /^[a-z]{2,3}(?:-[A-Z]{2})?$/i.test(lang.toString().trim());
}
function getTodayString(timeZone) {
  return Utilities.formatDate(new Date(), timeZone || getScriptTimeZone(), "yyyy-MM-dd");
}
function getHistoryCacheKey_(ss, userKey) {
  return 'srs_history_' + ss.getId() + '_' + userKey;
}
function getFlashcardData(customDeckId, clientTodayStr, clientTimeZone) {
  var timeZone = resolveClientTimeZone(clientTimeZone);
  var todayStr = getTodayString(timeZone);
  var ss = resolveSpreadsheet(customDeckId);
  var activeUserKey = getActiveUserId();
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
  var historyCache = CacheService.getUserCache();
  var cachedHistory = historyCache.get(getHistoryCacheKey_(ss, activeUserKey));
  if (cachedHistory) {
    try {
      var cachedHistoryData = JSON.parse(cachedHistory);
      studyHistory = cachedHistoryData.studyHistory || {};
      studiedLog = cachedHistoryData.studiedLog || studiedLog;
    } catch (e) {
      cachedHistory = null;
    }
  }
  if (historySheet && !cachedHistory) {
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
    historyCache.put(getHistoryCacheKey_(ss, activeUserKey), JSON.stringify({ studyHistory: studyHistory, studiedLog: studiedLog }), 300);
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
      var stableId = getProgressId_(cardIdInfo.ids[rowIndex], 0, c);
      if (savedProgressRaw[stableId]) savedProgress[stableId] = savedProgressRaw[stableId];
    }
  });
  return {
    headers: headers, rawRows: rawData, savedProgress: savedProgress,
    cardIds: cardIdInfo.ids, cardIdColumn: cardIdInfo.idColumn,
    studiedToday: { newCount: studiedLog.newCount || 0, oldCount: studiedLog.oldCount || 0 },
    activeUserKey: activeUserKey, activeUserEmail: getActiveUserEmail_(),
    todayStr: todayStr,
    studyHistory: studyHistory, ssName: ssName, activeDeckId: activeDeckId,
    webAppUrl: webAppUrl
  };
}
function translateText(text, sourceLang, targetLang, includeGoogle) {
  try {
    var cleanPrompt = text ? text.toString().trim() : '';
    if (!cleanPrompt) return { success: false, error: "Empty text prompt" };
    if (cleanPrompt.length > 500) return { success: false, error: "Source text exceeds the character limit." };
    if (countWords(cleanPrompt) > 20) return { success: false, error: "Source text exceeds the limit of 20 words." };
    if (!isValidLanguageCode_(sourceLang) || !isValidLanguageCode_(targetLang)) return { success: false, error: "Invalid language code." };
    enforceRateLimit_("translate", 20, 60);
    var googleTrans = "";
    if (includeGoogle !== false) {
      try {
        googleTrans = LanguageApp.translate(cleanPrompt, sourceLang || '', targetLang || 'en');
      } catch (e) {
        googleTrans = "Google Translation unavailable.";
      }
    }
    var geminiTrans = "";
    var deeplTrans = "";
    var parallelRequests = [];
    var requestTypes = [];
    var geminiNeedsFallback = false;
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (apiKey) {
      var promptStr = "Translate the following text directly from " + (sourceLang || "auto") + " to " + (targetLang || "en") + ". Output ONLY the raw translated string without quote marks or conversational filler: " + cleanPrompt;
      var payload = { "contents": [{ "parts": [{ "text": promptStr }] }] };
      parallelRequests.push({
        "url": "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODELS[0] + ":generateContent",
        "method": "post", "contentType": "application/json",
        "headers": { "x-goog-api-key": apiKey.trim() },
        "payload": JSON.stringify(payload), "muteHttpExceptions": true
      });
      requestTypes.push('gemini');
    } else {
      geminiTrans = "Gemini Error: Script Property 'GEMINI_API_KEY' is missing or empty.";
    }
    var deeplKey = PropertiesService.getScriptProperties().getProperty('DEEPL_API_KEY');
    if (deeplKey) {
      var trimmedDeepLKey = deeplKey.trim();
      var deeplUrl = trimmedDeepLKey.endsWith(':fx') ? "https://api-free.deepl.com/v2/translate" : "https://api.deepl.com/v2/translate";
      var srcDeepL = sourceLang ? sourceLang.toUpperCase() : null;
      if (srcDeepL === 'NO') srcDeepL = 'NB';
      var tgtDeepL = targetLang ? targetLang.toUpperCase() : 'EN';
      if (tgtDeepL === 'NO') tgtDeepL = 'NB';
      var deeplPayload = { "text": [cleanPrompt], "target_lang": tgtDeepL };
      if (srcDeepL && srcDeepL !== 'AUTO') deeplPayload["source_lang"] = srcDeepL;
      parallelRequests.push({
        "url": deeplUrl, "method": "post", "contentType": "application/json",
        "headers": { "Authorization": "DeepL-Auth-Key " + trimmedDeepLKey },
        "payload": JSON.stringify(deeplPayload), "muteHttpExceptions": true
      });
      requestTypes.push('deepl');
    } else {
      deeplTrans = "DeepL Error: Script Property 'DEEPL_API_KEY' is missing or empty.";
    }
    if (parallelRequests.length > 0) {
      try {
        var parallelResponses = UrlFetchApp.fetchAll(parallelRequests);
        for (var requestIndex = 0; requestIndex < parallelResponses.length; requestIndex++) {
          var response = parallelResponses[requestIndex];
          var responseCode = response.getResponseCode();
          var responseText = response.getContentText();
          var json = responseText ? JSON.parse(responseText) : {};
          if (requestTypes[requestIndex] === 'gemini') {
            if (responseCode === 200 && json.candidates && json.candidates[0] && json.candidates[0].content) {
              geminiTrans = json.candidates[0].content.parts[0].text.trim();
            } else if (responseCode === 429 || responseCode >= 500) {
              geminiNeedsFallback = true;
              geminiTrans = "Gemini Error (" + responseCode + "): " + (json.error ? json.error.message : "Unknown");
            } else {
              geminiTrans = "Gemini Error (" + responseCode + "): " + (json.error ? json.error.message : "Unknown");
            }
          } else if (responseCode === 200 && json.translations && json.translations[0]) {
            deeplTrans = json.translations[0].text.trim();
          } else if (json.message) {
            deeplTrans = "DeepL Error (" + responseCode + "): " + json.message;
          } else {
            deeplTrans = "DeepL Error: HTTP " + responseCode;
          }
        }
      } catch (e) {
        if (!geminiTrans) geminiTrans = "Gemini Exception: " + e.toString();
        if (!deeplTrans) deeplTrans = "DeepL Exception: " + e.toString();
      }
    }
    if (geminiNeedsFallback) {
      for (var modelIndex = 1; modelIndex < GEMINI_MODELS.length; modelIndex++) {
        try {
          var fallbackUrl = "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODELS[modelIndex] + ":generateContent";
          var fallbackResponse = UrlFetchApp.fetch(fallbackUrl, {
            "method": "post", "contentType": "application/json",
            "headers": { "x-goog-api-key": apiKey.trim() },
            "payload": JSON.stringify({ "contents": [{ "parts": [{ "text": promptStr }] }] }),
            "muteHttpExceptions": true
          });
          var fallbackCode = fallbackResponse.getResponseCode();
          var fallbackJson = JSON.parse(fallbackResponse.getContentText() || '{}');
          if (fallbackCode === 200 && fallbackJson.candidates && fallbackJson.candidates[0] && fallbackJson.candidates[0].content) {
            geminiTrans = fallbackJson.candidates[0].content.parts[0].text.trim();
            break;
          }
          if (fallbackCode !== 429 && fallbackCode < 500) break;
        } catch (fallbackError) {
          geminiTrans = "Gemini Exception: " + fallbackError.toString();
          break;
        }
      }
    }
    var translationResult = { success: true, google: googleTrans, gemini: geminiTrans, deepl: deeplTrans };
    return translationResult;
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
    var existingData = sheet.getDataRange().getValues();
    var cleanFront = cleanString(f);
    var cleanBack = cleanString(b);
    for (var existingRowIndex = 1; existingRowIndex < existingData.length; existingRowIndex++) {
      var existingFront = existingData[existingRowIndex][cA] ? existingData[existingRowIndex][cA].toString().trim() : '';
      var existingBack = existingData[existingRowIndex][cB] ? existingData[existingRowIndex][cB].toString().trim() : '';
      if (cleanString(existingFront) === cleanFront && cleanString(existingBack) === cleanBack) {
        return {
          success: false,
          duplicate: true,
          error: "This card already exists in the selected language pair.",
          front: existingFront,
          back: existingBack
        };
      }
    }
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
function saveProgressToSheet_(activeDeckId, uKey, progressMap, lockAlreadyHeld, changedCardIds) {
  var lock = null;
  try {
    if (!lockAlreadyHeld) {
      lock = LockService.getDocumentLock();
      lock.waitLock(15000);
    }
    var ss = SpreadsheetApp.openById(activeDeckId);
    var sheet = ss.getSheetByName('SRS_Progress');
    var headers = ['UserKey', 'CardId', 'Interval', 'NextReview', 'FailCount', 'IsLeech', 'LastReviewed', 'EF', 'PrevInterval'];
    var created = false;
    if (!sheet) {
      sheet = ss.insertSheet('SRS_Progress');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      created = true;
    }
    if (sheet.getMaxColumns() < headers.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    }
    var data = sheet.getDataRange().getValues();
    if (data.length === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      data = [headers];
    } else if (data[0].length < headers.length || data[0][8] !== headers[8]) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    var rowsByCardId = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === uKey) {
        rowsByCardId[data[i][1]] = i + 1;
      }
    }
    var idsToSave = changedCardIds || Object.keys(progressMap);
    var newRows = [];
    for (var idIndex = 0; idIndex < idsToSave.length; idIndex++) {
      var cardId = idsToSave[idIndex];
      var p = progressMap[cardId];
      if (!p) continue;
      var rowValues = [uKey, cardId, p.interval, "'" + p.nextReview, p.failCount, p.isLeech ? 1 : 0, p.lastReviewed ? "'" + p.lastReviewed : "", p.ef, p.prevInterval || ""];
      if (rowsByCardId[cardId]) {
        sheet.getRange(rowsByCardId[cardId], 1, 1, headers.length).setValues([rowValues]);
      } else {
        newRows.push(rowValues);
      }
    }
    if (newRows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    if (created) formatTrackingSheet_(sheet, 'progress');
  } catch (e) {
    console.error("Failed to save progress: " + e.message);
    throw e;
  } finally {
    if (lock) lock.releaseLock();
  }
}
function saveHistoryToSheet_(activeDeckId, uKey, todayStr, newCardDelta, oldCardDelta, lockAlreadyHeld, timeZone, absoluteCounts) {
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
    var cachedNextNew = newCardDelta;
    var cachedNextOld = oldCardDelta;
    if (rowIndex > -1) {
      var currentNew = parseInt(data[rowIndex - 1][2] || 0, 10);
      var currentOld = parseInt(data[rowIndex - 1][3] || 0, 10);
      var nextNew = absoluteCounts ? Math.max(currentNew, newCardDelta) : currentNew + newCardDelta;
      var nextOld = absoluteCounts ? Math.max(currentOld, oldCardDelta) : currentOld + oldCardDelta;
      cachedNextNew = nextNew;
      cachedNextOld = nextOld;
      sheet.getRange(rowIndex, 3, 1, 3).setValues([[nextNew, nextOld, nextNew + nextOld]]);
    } else {
      sheet.appendRow([uKey, "'" + todayStr, newCardDelta, oldCardDelta, absoluteCounts ? newCardDelta + oldCardDelta : totalDelta]);
    }
    var historyCache = CacheService.getUserCache();
    var historyCacheKey = getHistoryCacheKey_(ss, uKey);
    var cachedHistory = {};
    var existingCachedHistory = historyCache.get(historyCacheKey);
    if (existingCachedHistory) {
      try { cachedHistory = JSON.parse(existingCachedHistory) || {}; } catch (e) { cachedHistory = {}; }
      cachedHistory.studyHistory = cachedHistory.studyHistory || {};
      cachedHistory.studiedLog = cachedHistory.studiedLog || {};
      cachedHistory.studyHistory[todayStr] = cachedNextNew + cachedNextOld;
      cachedHistory.studiedLog = { date: todayStr, newCount: cachedNextNew, oldCount: cachedNextOld };
      historyCache.put(historyCacheKey, JSON.stringify(cachedHistory), 300);
    }
  } catch (e) {
    console.error("Failed to save history: " + e.message);
    throw e;
  } finally {
    if (lock) lock.releaseLock();
  }
}
function getReviewEventSheet_(ss) {
  var sheet = ss.getSheetByName('SRS_ReviewEvents');
  if (!sheet) {
    sheet = ss.insertSheet('SRS_ReviewEvents');
    sheet.getRange(1, 1, 1, REVIEW_EVENT_HEADERS.length).setValues([REVIEW_EVENT_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, REVIEW_EVENT_HEADERS.length).setFontWeight('bold');
  } else if (sheet.getMaxColumns() < REVIEW_EVENT_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), REVIEW_EVENT_HEADERS.length - sheet.getMaxColumns());
  }
  return sheet;
}
function parseReviewEventResult_(value) {
  if (!value) return null;
  try { return JSON.parse(value.toString().replace(/^'/, '')); } catch (e) { return null; }
}
function saveHistoryCountsFromEvents_(activeDeckId, uKey, todayStr, eventData, timeZone) {
  var newCount = 0;
  var oldCount = 0;
  for (var i = 1; i < eventData.length; i++) {
    if (eventData[i][0] !== uKey || parseDateToString(eventData[i][8], timeZone) !== todayStr) continue;
    if (eventData[i][5] === true || eventData[i][5] === 1 || eventData[i][5] === 'true') newCount++;
    else oldCount++;
  }
  saveHistoryToSheet_(activeDeckId, uKey, todayStr, newCount, oldCount, true, timeZone, true);
}
function saveSegmentResults(batchArray, customDeckId, clientTodayStr, clientTimeZone, colIdxA, colIdxB) {
  if (!batchArray || batchArray.length === 0) return { success: true };
  if (!Array.isArray(batchArray) || batchArray.length > 50) return { success: false, error: "Invalid result batch." };
  var ss = resolveSpreadsheet(customDeckId);
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);
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
      var identity = getCardIdentity_(mainData[rowIndex], selectedColA, selectedColB, mainData[rowIndex][cardIdInfo.idColumn]);
      if (identity.front && identity.back) validCardIds[identity.progressId] = true;
    }
    var timeZone = resolveClientTimeZone(clientTimeZone);
    var todayStr = getTodayString(timeZone);
    var nowMs = Date.now();
    var uKey = getActiveUserId();
    var progressSheet = ss.getSheetByName('SRS_Progress');
    var savedProgress = readProgressForUser_(progressSheet, uKey, timeZone, todayStr);
    var eventSheet = getReviewEventSheet_(ss);
    var eventData = eventSheet.getDataRange().getValues();
    var originalEventLength = eventData.length;
    var eventsByClientId = {};
    for (var eventRow = 1; eventRow < eventData.length; eventRow++) {
      if (eventData[eventRow][0] === uKey) eventsByClientId[String(eventData[eventRow][1])] = eventRow;
    }
    var pendingEvents = [];
    batchArray.forEach(function (item) {
      if (!item || !item.cardId || !validCardIds[item.cardId] || item.clientId === undefined || item.clientId === null) throw new Error("Invalid study result.");
      var clientId = String(item.clientId);
      var existingRow = eventsByClientId[clientId];
      if (existingRow !== undefined) {
        if (eventData[existingRow][10] === 'APPLIED') return;
        pendingEvents.push({ row: existingRow, item: {
          cardId: eventData[existingRow][2], isCorrect: eventData[existingRow][3] === true || eventData[existingRow][3] === 1,
          confidence: eventData[existingRow][4], isTypo: eventData[existingRow][6] === true || eventData[existingRow][6] === 1,
          inDifficultMode: eventData[existingRow][7] === true || eventData[existingRow][7] === 1
        }});
        return;
      }
      var newRow = eventData.length;
      eventData.push([uKey, clientId, item.cardId.toString(), item.isCorrect === true, item.confidence || '', '', item.isTypo === true, item.inDifficultMode === true, item.lastReviewed || todayStr, '', 'PENDING', '']);
      eventsByClientId[clientId] = newRow;
      pendingEvents.push({ row: newRow, item: item });
    });
    for (var pendingIndex = 0; pendingIndex < pendingEvents.length; pendingIndex++) {
      var event = pendingEvents[pendingIndex];
      var eventRowData = eventData[event.row];
      var cardId = event.item.cardId.toString();
      var result = parseReviewEventResult_(eventRowData[9]);
      if (!result) {
        var wasNew = !savedProgress[cardId];
        var cardState = savedProgress[cardId] || { interval: 1, failCount: 0, isLeech: false, ef: 2.5, prevInterval: null };
        result = scheduleAnswerResult_(cardState, event.item, wasNew, nowMs, todayStr);
        eventRowData[5] = wasNew;
        eventRowData[9] = JSON.stringify(result);
      }
      savedProgress[cardId] = result;
    }
    // Persist computed results before the sheet update. A retry can reuse them
    // after an interruption without scheduling the same answer twice.
    var newEventRows = [];
    for (var eventWriteIndex = 0; eventWriteIndex < pendingEvents.length; eventWriteIndex++) {
      var eventToWrite = pendingEvents[eventWriteIndex];
      if (eventToWrite.row >= originalEventLength) newEventRows.push(eventData[eventToWrite.row]);
      else eventSheet.getRange(eventToWrite.row + 1, 1, 1, REVIEW_EVENT_HEADERS.length).setValues([eventData[eventToWrite.row]]);
    }
    if (newEventRows.length > 0) eventSheet.getRange(originalEventLength + 1, 1, newEventRows.length, REVIEW_EVENT_HEADERS.length).setValues(newEventRows);
    var changedCardIds = pendingEvents.map(function (event) { return event.item.cardId.toString(); });
    saveProgressToSheet_(activeDeckId, uKey, savedProgress, true, changedCardIds);
    saveHistoryCountsFromEvents_(activeDeckId, uKey, todayStr, eventData, timeZone);
    for (var appliedIndex = 0; appliedIndex < pendingEvents.length; appliedIndex++) {
      eventData[pendingEvents[appliedIndex].row][10] = 'APPLIED';
      eventData[pendingEvents[appliedIndex].row][11] = new Date();
    }
    for (var appliedWriteIndex = 0; appliedWriteIndex < pendingEvents.length; appliedWriteIndex++) {
      var appliedEvent = pendingEvents[appliedWriteIndex];
      eventSheet.getRange(appliedEvent.row + 1, 11, 1, 2).setValues([[eventData[appliedEvent.row][10], eventData[appliedEvent.row][11]]]);
    }
    var appliedProgress = {};
    batchArray.forEach(function (item) {
      var itemCardId = item.cardId.toString();
      if (savedProgress[itemCardId]) appliedProgress[itemCardId] = savedProgress[itemCardId];
    });
    return { success: true, appliedProgress: appliedProgress };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
function deleteRowsInRuns_(sheet, rowNumbers) {
  if (!rowNumbers || rowNumbers.length === 0) return;
  var sorted = rowNumbers.slice().sort(function (a, b) { return b - a; });
  var runEnd = sorted[0];
  var runStart = runEnd;
  for (var i = 1; i < sorted.length; i++) {
    if (sorted[i] === runStart - 1) {
      runStart = sorted[i];
    } else {
      sheet.deleteRows(runStart, runEnd - runStart + 1);
      runStart = sorted[i];
      runEnd = runStart;
    }
  }
  sheet.deleteRows(runStart, runEnd - runStart + 1);
}
function rewriteRowsAndTrim_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  var width = rows[0].length;
  sheet.getRange(1, 1, rows.length, width).setValues(rows);
  var lastRow = sheet.getLastRow();
  if (lastRow > rows.length) sheet.deleteRows(rows.length + 1, lastRow - rows.length);
}
function getArchiveSheet_(ss, mainSheet) {
  var archiveSheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(ARCHIVE_SHEET_NAME);
    archiveSheet.appendRow(mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0]);
  }
  return archiveSheet;
}
// Archive operations are idempotent and restore progress if deletion fails.
function archiveMasteredCards(customDeckId, colIdxA, colIdxB) {
  var ss = resolveSpreadsheet(customDeckId);
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);
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
    var archiveSheet = getArchiveSheet_(ss, mainSheet);
    if (!progSheet) return { success: false, message: "No progress found." };
    var progData = progSheet.getDataRange().getValues();
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
    if (rowsToDelete.length > 0) {
      if (rowsToDelete.length > 10) {
        var rowsToDeleteSet = {};
        rowsToDelete.forEach(function (rowNumber) { rowsToDeleteSet[rowNumber] = true; });
        var retainedMainData = [mainData[0]];
        for (var retainedIndex = 1; retainedIndex < mainData.length; retainedIndex++) {
          if (!rowsToDeleteSet[retainedIndex + 1]) retainedMainData.push(mainData[retainedIndex]);
        }
        rewriteRowsAndTrim_(mainSheet, retainedMainData);
      } else {
        deleteRowsInRuns_(mainSheet, rowsToDelete);
      }
    }
    if (rowsToDelete.length > 0 || rowsToArchive.length > 0) {
      var progressRowsToDelete = [];
      for (var p = 1; p < progData.length; p++) {
        if (progData[p][0] === activeUserKey && masteredIds.has(progData[p][1])) progressRowsToDelete.push(p + 1);
      }
      if (progressRowsToDelete.length > 10) {
        var progressDeleteSet = {};
        progressRowsToDelete.forEach(function (rowNumber) { progressDeleteSet[rowNumber] = true; });
        var retainedProgressData = [progData[0]];
        for (var retainedProgressIndex = 1; retainedProgressIndex < progData.length; retainedProgressIndex++) {
          if (!progressDeleteSet[retainedProgressIndex + 1]) retainedProgressData.push(progData[retainedProgressIndex]);
        }
        rewriteRowsAndTrim_(progSheet, retainedProgressData);
      } else {
        deleteRowsInRuns_(progSheet, progressRowsToDelete);
      }
    }
    return { success: true, count: rowsToArchive.length, message: "Archived " + rowsToArchive.length + " mastered cards." };
  } catch (e) {
    return { success: false, message: "System busy. Please try again later." };
  } finally {
    lock.releaseLock();
  }
}
function getAudioBase64(text, lang) {
  try {
    enforceRateLimit_("audio", 60, 60);
    var cleanText = text ? text.toString().trim() : "";
    var cleanLang = lang ? lang.toString().trim() : "";
    if (!cleanText || cleanText.length > 500 || countWords(cleanText) > 20 || !isValidLanguageCode_(cleanLang)) {
      return { success: false, error: "Invalid audio request." };
    }
    var cache = CacheService.getUserCache();
    var cacheKey = 'srs_audio_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, cleanLang + '|' + cleanText));
    var cachedAudio = cache.get(cacheKey);
    if (cachedAudio) return { success: true, data: cachedAudio };
    var url = "https://translate.google.com/translate_tts?ie=UTF-8&q=" + encodeURIComponent(cleanText) + "&tl=" + encodeURIComponent(cleanLang) + "&client=tw-ob";
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      var blob = response.getBlob();
      var audioData = "data:audio/mp3;base64," + Utilities.base64Encode(blob.getBytes());
      cache.put(cacheKey, audioData, 21600);
      return { success: true, data: audioData };
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
    var options = { "method": "post", "contentType": "application/json", "headers": { "x-goog-api-key": apiKey.trim() }, "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    for (var i = 0; i < models.length; i++) {
      var url = "https://generativelanguage.googleapis.com/v1beta/models/" + models[i] + ":generateContent";
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
function archiveSingleCard(rowIndex, customDeckId, cardId, colIdxA, colIdxB) {
  var ss = resolveSpreadsheet(customDeckId);
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    var mainSheet = ss.getSheets()[0];
    var archiveSheet = getArchiveSheet_(ss, mainSheet);
    if (!cardId) return { success: false, error: "Missing card identity." };
    var location = getCardLocation_(mainSheet, rowIndex, colIdxA, colIdxB, cardId);
    var r = location.rowIndex;
    var progressId = getProgressId_(cardId, location.colIdxA, location.colIdxB);
    var archiveIdInfo = ensureCardIdColumn_(archiveSheet);
    // Copy to archive, then delete from main
    var rowData = mainSheet.getRange(r, 1, 1, mainSheet.getLastColumn()).getValues();
    var archiveLastRow = archiveSheet.getLastRow();
    var archiveIdRange = archiveLastRow > 1 ? archiveSheet.getRange(2, archiveIdInfo.idColumn + 1, archiveLastRow - 1, 1) : null;
    var alreadyArchived = !!(archiveIdRange && archiveIdRange.createTextFinder(cardId).matchEntireCell(true).findNext());
    if (!alreadyArchived) archiveSheet.appendRow(rowData[0]);
    var activeUserKey = getActiveUserId();
    var progressSheet = ss.getSheetByName('SRS_Progress');
    if (progressSheet && cardId) {
      var progressData = progressSheet.getDataRange().getValues();
      var singleProgressRowsToDelete = [];
      for (var progressIndex = 1; progressIndex < progressData.length; progressIndex++) {
        if (progressData[progressIndex][0] === activeUserKey && progressData[progressIndex][1] === progressId) singleProgressRowsToDelete.push(progressIndex + 1);
      }
      if (singleProgressRowsToDelete.length > 0) {
        deleteRowsInRuns_(progressSheet, singleProgressRowsToDelete);
      }
    }
    deleteRowsInRuns_(mainSheet, [r]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}