/**
* ============================================================================
* SERVER-SIDE GOOGLE APPS SCRIPT (SM-2, INTERVAL FUZZING, HOURLY REVIEWS)
* Handles Google Sheets interactions, data processing, and API routing.
* ============================================================================
*/

var TEMPLATE_ID = "1s04BKwYk-s0WVSmsJOfCxCz4X-fqdZzR0O65ANymu9E";

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
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('📚 SRS Flashcards')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

  html.append('<script>window.IS_WEB_APP = true;</script>');
  if (e && e.parameter) {
    if (e.parameter.deck) html.append('<script>window.URL_DECK_ID = "' + escapeJS(e.parameter.deck) + '";</script>');
    if (e.parameter.name) html.append('<script>window.URL_SYNC_NAME = "' + escapeJS(e.parameter.name) + '";</script>');
  }
  return html;
}

function showFlashcardSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('📚 SRS Flashcards');
  html.append('<script>window.IS_WEB_APP = false;</script>');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showFlashcardDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(600)
    .setHeight(720);
  html.append('<script>window.IS_WEB_APP = false;</script>');
  SpreadsheetApp.getUi().showModelessDialog(html, '📚 SRS Flashcards');
}

function resolveSpreadsheet(customDeckId) {
  if (customDeckId && customDeckId.toString().trim().length > 0) {
    try {
      return SpreadsheetApp.openById(customDeckId.toString().trim());
    } catch (e) {
      return SpreadsheetApp.openById(TEMPLATE_ID);
    }
  }
  try {
    return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(TEMPLATE_ID);
  } catch (e) {
    return SpreadsheetApp.openById(TEMPLATE_ID);
  }
}

function preventInjection(text) {
  var str = text ? text.toString() : '';
  if (/^[=+\-@]/.test(str)) return "'" + str;
  return str;
}

var CACHED_SCRIPT_TIMEZONE = null;

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

function getStageFromInterval(ivl, isNew) {
  if (isNew || ivl < 1) return 1;
  if (ivl <= 3) return 2;
  if (ivl <= 7) return 3;
  if (ivl <= 14) return 4;
  if (ivl <= 30) return 5;
  if (ivl <= 60) return 6;
  if (ivl <= 90) return 7;
  if (ivl <= 180) return 8;
  if (ivl <= 359) return 9;
  return 10;
}

function getIntervalForStage(stage) {
  switch (stage) {
    case 1: return 1 / 24;
    case 2: return 1;
    case 3: return 3;
    case 4: return 7;
    case 5: return 14;
    case 6: return 30;
    case 7: return 60;
    case 8: return 90;
    case 9: return 180;
    default: return 360;
  }
}

function getWeekKey(timeZone) {
  var today = new Date();
  var tz = timeZone || getScriptTimeZone();
  var year = Utilities.formatDate(today, tz, "yyyy");
  var d = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return year + "_W" + String(weekNo).padStart(2, '0');
}

function getMonthKey(timeZone) {
  var today = new Date();
  return Utilities.formatDate(today, timeZone || getScriptTimeZone(), "yyyy_MM");
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

function sanitizeText(str, maxLen) {
  if (!str) return '';
  var clean = str.toString().trim();
  return clean.substring(0, maxLen || 500);
}

function countWords(str) {
  if (!str) return 0;
  return str.toString().trim().split(/\s+/).filter(Boolean).length;
}

function generateFunName() {
  var adjectives = ["Curious", "Brave", "Clever", "Swift", "Happy", "Eager", "Wise", "Keen", "Bright", "Calm", "Noble", "Nimble", "Gentle", "Jolly", "Mighty", "Cosmic", "Lively"];
  var animals = ["Panda", "Owl", "Falcon", "Tiger", "Fox", "Viking", "Bear", "Dolphin", "Otter", "Wolf", "Lynx", "Eagle", "Moose", "Penguin", "Koala", "Raven", "Puffin"];
  var adj = adjectives[Math.floor(Math.random() * adjectives.length)];
  var anim = animals[Math.floor(Math.random() * animals.length)];
  var num = Math.floor(10 + Math.random() * 90);
  return adj + anim + num;
}

function getActiveUserId() {
  try {
    // Fetches the email of the logged-in Google User
    var email = Session.getActiveUser().getEmail();
    if (email && email.length > 0) return email;
    
    // Fallback if accessed via sidebar instead of web app
    var effectiveEmail = Session.getEffectiveUser().getEmail();
    if (effectiveEmail && effectiveEmail.length > 0) return effectiveEmail;
    
  } catch (e) {
    console.error("Auth Error: " + e.message);
  }
  return "unauthenticated_user";
}

function getTodayString(timeZone) {
  return Utilities.formatDate(new Date(), timeZone || getScriptTimeZone(), "yyyy-MM-dd");
}

/**
* Applies a scaling random variance to the interval to prevent review clumping.
*/
function applyFuzz(intervalInDays) {
  var fuzzRange = 0;

  // Sub-day fuzzing (adds minutes/hours of variance)
  if (intervalInDays <= 0.26) fuzzRange = 0.02;      // +/- ~30 mins for 6h
  else if (intervalInDays <= 0.51) fuzzRange = 0.04; // +/- ~1 hr for 12h
  else if (intervalInDays <= 1) fuzzRange = 0.08;   // +/- ~2 hrs for 1d

  // Standard day fuzzing
  else if (intervalInDays <= 4) fuzzRange = 0.5;    // +/- 12 hours
  else if (intervalInDays <= 7) fuzzRange = 1;      // +/- 1 day
  else if (intervalInDays <= 14) fuzzRange = 2;     // +/- 2 days
  else fuzzRange = intervalInDays * 0.1;            // +/- 10% for mature cards

  var fuzz = (Math.random() * (fuzzRange * 2)) - fuzzRange;
  return Math.max(0.01, intervalInDays + fuzz); // Ensure it never drops to 0
}

function getFlashcardData(clientDisplayName, customDeckId, clientTodayStr, clientTimeZone) {
  var activeUserKey = getActiveUserId();
  var docProps = PropertiesService.getDocumentProperties();
  var timeZone = resolveClientTimeZone(clientTimeZone);
  var todayStr = clientTodayStr || getTodayString(timeZone);

  var storedName = docProps.getProperty('SRS_NAME_' + activeUserKey);
  var activeDisplayName = "";
  var displayNameExists = false;

  if (storedName && storedName !== "undefined" && storedName !== "null") {
    activeDisplayName = storedName;
    displayNameExists = true;
  } else {
    try {
      var email = Session.getActiveUser().getEmail();
      if (email && email.length > 0) activeDisplayName = email.split('@')[0];
    } catch (e) { }
    if (!activeDisplayName || activeDisplayName === "undefined") {
      activeDisplayName = (clientDisplayName && clientDisplayName !== "undefined") ? sanitizeText(clientDisplayName, 20) : generateFunName();
    }
    docProps.setProperty('SRS_NAME_' + activeUserKey, activeDisplayName);
    displayNameExists = true;
  }

  var ss = resolveSpreadsheet(customDeckId);
  var ssName = ss.getName();
  var activeDeckId = ss.getId();
  var baseUrl = "https://script.google.com/macros/s/AKfycbypzH1uT_9sCzAQgw0MZ_MVuGC1NGHIWnAL0ajixNXiEx4ND2UyKASaNLKiWGElpGtM/exec";
  var webAppUrl = baseUrl + "?deck=" + encodeURIComponent(activeDeckId) + "&name=" + encodeURIComponent(activeDisplayName);
  var sheet = ss.getSheets()[0];
  var allValues = sheet.getDataRange().getValues();

  var weekKey = getWeekKey(timeZone);
  var monthKey = getMonthKey(timeZone);
  var leaderboards = getLeaderboards(ss, weekKey, monthKey);

  var savedProgressRaw = {};
  var progressSheet = ss.getSheetByName('SRS_Progress');
  if (progressSheet) {
    var pData = progressSheet.getDataRange().getValues();
    for (var i = 1; i < pData.length; i++) {
      if (pData[i][0] === activeUserKey) {
        var cId = pData[i][1];
        savedProgressRaw[cId] = {
          interval: pData[i][2] !== undefined ? parseFloat(pData[i][2]) : 1,
          nextReview: pData[i][3] instanceof Date ? Utilities.formatDate(pData[i][3], timeZone, "yyyy-MM-dd HH:mm:ss") : (pData[i][3] ? pData[i][3].toString().trim() : todayStr),
          failCount: pData[i][4] !== undefined ? parseFloat(pData[i][4]) : 0, 
          isLeech: pData[i][5] === 1,
          lastReviewed: pData[i][6] ? parseDateToString(pData[i][6], timeZone) : "",
          ef: pData[i][7] !== undefined ? parseFloat(pData[i][7]) : 2.5
        };
      }
    }
  }

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
      activeUserKey: activeUserKey, displayName: activeDisplayName, displayNameExists: displayNameExists,
      todayStr: todayStr, leaderboard: leaderboards.allTime,
      leaderboardWeekly: leaderboards.weekly, leaderboardMonthly: leaderboards.monthly,
      studyHistory: studyHistory, ssName: ssName, activeDeckId: activeDeckId,
      webAppUrl: webAppUrl
    };
  }

  var headers = allValues[0].map(function (h) { return h.toString().trim(); });
  var rawData = allValues.slice(1);
  var savedProgress = {};
  var validCardIds = {};

  rawData.forEach(function (row) {
    var front = row[0] ? row[0].toString().trim() : "";
    for (var c = 1; c < row.length; c++) {
      var back = row[c] ? row[c].toString().trim() : "";
      if (front && back) validCardIds[cleanString(front) + ":::" + cleanString(back)] = true;
    }
  });

  for (var cardKey in savedProgressRaw) {
    if (!validCardIds[cardKey]) continue;
    var entry = savedProgressRaw[cardKey];
    if (entry) {
      savedProgress[cardKey] = {
        interval: entry.interval || 1,
        nextReview: entry.nextReview || todayStr,
        failCount: entry.failCount || 0,
        isLeech: entry.isLeech || false,
        lastReviewed: entry.lastReviewed || "",
        ef: entry.ef || 2.5
      };
    }
  }

  return {
    headers: headers, rawRows: rawData, savedProgress: savedProgress,
    studiedToday: { newCount: studiedLog.newCount || 0, oldCount: studiedLog.oldCount || 0 },
    activeUserKey: activeUserKey, displayName: activeDisplayName, displayNameExists: displayNameExists,
    todayStr: todayStr, leaderboard: leaderboards.allTime,
    leaderboardWeekly: leaderboards.weekly, leaderboardMonthly: leaderboards.monthly,
    studyHistory: studyHistory, ssName: ssName, activeDeckId: activeDeckId,
    webAppUrl: webAppUrl
  };
}

function saveDisplayName(activeUserKey, newDisplayName) {
  if (!newDisplayName || newDisplayName === "undefined") return { success: false, message: "Name cannot be empty." };
  var docProps = PropertiesService.getDocumentProperties();
  var uKey = getActiveUserId();
  var cleanName = sanitizeText(newDisplayName, 20).trim();
  if (!cleanName || cleanName === "undefined") return { success: false, message: "Invalid name." };

  var allProps = docProps.getProperties();
  for (var key in allProps) {
    if (key.startsWith('SRS_NAME_')) {
      if (allProps[key].toLowerCase() === cleanName.toLowerCase() && key !== 'SRS_NAME_' + uKey) {
        return { success: false, message: "⚠️ That name is already taken by another user." };
      }
    }
  }

  docProps.setProperty('SRS_NAME_' + uKey, cleanName);
  return { success: true, name: cleanName };
}

function translateText(text, sourceLang, targetLang) {
  try {
    var cleanPrompt = text ? text.toString().trim() : '';
    if (!cleanPrompt) return { success: false, error: "Empty text prompt" };
    if (countWords(cleanPrompt) > 20) return { success: false, error: "Source text exceeds the limit of 20 words." };

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
        var models = [
          "gemini-3.6-flash",
          "gemini-3.5-flash",
          "gemini-3.1-flash-lite",
          "gemini-2.5-flash"
        ];
        
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

function addNewCardToSheet(frontText, backText, customDeckId) {
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    var ss = resolveSpreadsheet(customDeckId);
    var sheet = ss.getSheets()[0];
    var f = preventInjection(frontText ? frontText.toString().trim() : '');
    var b = preventInjection(backText ? backText.toString().trim() : '');

    if (!f || !b) return { success: false, error: "Both front and back required." };
    if (countWords(f) > 20 || countWords(b) > 20) return { success: false, error: "Text exceeds maximum 20 words limit." };

    sheet.appendRow([f, b]);
    return { success: true, rowIndex: sheet.getLastRow(), front: f, back: b };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateCardInSheet(rowIndex, colIdxA, colIdxB, newFront, newBack, customDeckId) {
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    var ss = resolveSpreadsheet(customDeckId);
    var sheet = ss.getSheets()[0];
    var r = parseInt(rowIndex, 10);
    var cA = parseInt(colIdxA, 10);
    var cB = parseInt(colIdxB, 10);

    if (isNaN(r) || r < 2 || isNaN(cA) || isNaN(cB)) return { success: false, error: "Invalid index" };

    var cleanF = preventInjection(newFront ? newFront.toString().trim() : '');
    var cleanB = preventInjection(newBack ? newBack.toString().trim() : '');
    if (countWords(cleanF) > 20 || countWords(cleanB) > 20) return { success: false, error: "Text exceeds maximum 20 words limit." };

    sheet.getRange(r, cA + 1).setValue(cleanF);
    sheet.getRange(r, cB + 1).setValue(cleanB);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateLeaderboardSheet(ss) {
  var sheet = ss.getSheetByName('SRS_Leaderboard');
  if (!sheet) {
    sheet = ss.insertSheet('SRS_Leaderboard');
    sheet.appendRow(['UserKey', 'PeriodType', 'PeriodKey', 'DisplayName', 'Score']);
  }
  return sheet;
}

function formatLeaderboardRows(data, periodType, periodKey) {
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[1] === periodType && (row[2] || '') === (periodKey || '')) {
      var score = parseInt(row[4], 10) || 0;
      if (score > 0) list.push({ username: sanitizeText(row[3], 20), score: score });
    }
  }
  list.sort(function (a, b) { return b.score - a.score; });
  return list.slice(0, 10);
}

function getLeaderboards(ss, weekKey, monthKey) {
  var sheet = getOrCreateLeaderboardSheet(ss);
  var data = sheet.getDataRange().getValues();
  return {
    allTime: formatLeaderboardRows(data, 'ALL', ''),
    weekly: formatLeaderboardRows(data, 'WEEK', weekKey),
    monthly: formatLeaderboardRows(data, 'MONTH', monthKey)
  };
}

function saveProgressToSheet(activeDeckId, uKey, progressMap, lockAlreadyHeld) {
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

function saveHistoryToSheet(activeDeckId, uKey, todayStr, newCardDelta, oldCardDelta, lockAlreadyHeld, timeZone) {
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

function updateLeaderboardInSheet(activeDeckId, uKey, dName, reviewsCount, weekKey, monthKey, lockAlreadyHeld) {
  var lock = null;
  try {
    if (!lockAlreadyHeld) {
      lock = LockService.getDocumentLock();
      lock.waitLock(15000);
    }
    var ss = SpreadsheetApp.openById(activeDeckId);
    var sheet = getOrCreateLeaderboardSheet(ss);
    var data = sheet.getDataRange().getValues();

    var periods = [{ type: 'ALL', key: '' }, { type: 'WEEK', key: weekKey }, { type: 'MONTH', key: monthKey }];
    var rowIndexMap = {};
    for (var i = 1; i < data.length; i++) {
      rowIndexMap[data[i][0] + '|' + data[i][1] + '|' + data[i][2]] = i;
    }

    var newRows = [];
    periods.forEach(function (p) {
      var mapKey = uKey + '|' + p.type + '|' + p.key;
      if (rowIndexMap.hasOwnProperty(mapKey)) {
        var idx = rowIndexMap[mapKey];
        var currentScore = parseInt(data[idx][4], 10) || 0;
        data[idx][3] = dName;
        data[idx][4] = currentScore + reviewsCount;
      } else {
        newRows.push([uKey, p.type, p.key, dName, reviewsCount]);
      }
    });

    if (data.length > 1) sheet.getRange(2, 1, data.length - 1, 5).setValues(data.slice(1));
    if (newRows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 5).setValues(newRows);
  } catch (e) {
    console.error("Failed to update leaderboard: " + e.message);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function saveSegmentResults(batchArray, activeUserKey, displayName, customDeckId, clientTodayStr, clientTimeZone) {
  if (!batchArray || batchArray.length === 0) return { success: true };

  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);

    var ss = resolveSpreadsheet(customDeckId);
    var activeDeckId = ss.getId();
    var timeZone = resolveClientTimeZone(clientTimeZone);
    var todayStr = clientTodayStr || getTodayString(timeZone);
    var uKey = getActiveUserId();
    var weekKey = getWeekKey(timeZone);
    var monthKey = getMonthKey(timeZone);

    var docProps = PropertiesService.getDocumentProperties();
    var storedName = docProps.getProperty('SRS_NAME_' + uKey);
    var dName = (storedName && storedName !== "undefined") ? storedName : (sanitizeText(displayName, 20) || "Learner");

    var savedProgress = {};
    var progressSheet = ss.getSheetByName('SRS_Progress');
    if (progressSheet) {
      var pData = progressSheet.getDataRange().getValues();
      for (var i = 1; i < pData.length; i++) {
        if (pData[i][0] === uKey) {
          var cId = pData[i][1];
          savedProgress[cId] = {
            interval: pData[i][2] !== undefined ? parseFloat(pData[i][2]) : 1,
            nextReview: pData[i][3] instanceof Date ? Utilities.formatDate(pData[i][3], timeZone, "yyyy-MM-dd HH:mm:ss") : (pData[i][3] ? pData[i][3].toString().trim() : todayStr),
            failCount: pData[i][4] !== undefined ? parseFloat(pData[i][4]) : 0,
            isLeech: pData[i][5] === 1,
            lastReviewed: pData[i][6] ? parseDateToString(pData[i][6], timeZone) : "",
            ef: pData[i][7] !== undefined ? parseFloat(pData[i][7]) : 2.5,
            prevInterval: pData[i][8] !== undefined && pData[i][8] !== "" ? parseFloat(pData[i][8]) : null
          };
        }
      }
    }

    var reviewsCount = 0;
    var newDelta = 0;
    var oldDelta = 0;

    batchArray.forEach(function (item) {
      reviewsCount++;
      var cardId = item.cardId;
      var wasNew = item.wasNew;

      var cardState = savedProgress[cardId] || { interval: 1, failCount: 0, isLeech: false, ef: 2.5, prevInterval: null };
      var newInterval = item.newInterval !== undefined && item.newInterval !== null ? parseFloat(item.newInterval) : (parseFloat(cardState.interval) || 1);
      if (!isFinite(newInterval) || newInterval <= 0) newInterval = parseFloat(cardState.interval) || 1;

      var newEF = item.newEF !== undefined && item.newEF !== null ? parseFloat(item.newEF) : (parseFloat(cardState.ef) || 2.5);
      if (!isFinite(newEF) || newEF < 1.3) newEF = Math.max(1.3, parseFloat(cardState.ef) || 2.5);

      var failCount = item.failCount !== undefined && item.failCount !== null ? parseFloat(item.failCount) : (parseFloat(cardState.failCount) || 0);
      if (!isFinite(failCount) || failCount < 0) failCount = parseFloat(cardState.failCount) || 0;

      var isLeech = item.isLeech !== undefined ? !!item.isLeech : !!cardState.isLeech;
      var prevInterval = item.prevInterval !== undefined ? item.prevInterval : cardState.prevInterval;
      if (prevInterval !== null && prevInterval !== "" && prevInterval !== undefined) {
        var parsedPrev = parseFloat(prevInterval);
        prevInterval = isFinite(parsedPrev) ? parsedPrev : null;
      } else {
        prevInterval = null;
      }

      var lastReviewed = item.lastReviewed ? item.lastReviewed.toString().trim() : todayStr;
      if (!lastReviewed) lastReviewed = todayStr;

      var nextDateStr = item.nextReview ? item.nextReview.toString().trim() : "";
      if (!nextDateStr) {
        nextDateStr = cardState.nextReview || todayStr;
      }

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

    saveProgressToSheet(activeDeckId, uKey, savedProgress, true);
    saveHistoryToSheet(activeDeckId, uKey, todayStr, newDelta, oldDelta, true, timeZone);
    updateLeaderboardInSheet(activeDeckId, uKey, dName, reviewsCount, weekKey, monthKey, true);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function archiveMasteredCards(customDeckId, activeUserKey) {
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);
    var ss = resolveSpreadsheet(customDeckId);
    var mainSheet = ss.getSheets()[0];
    var progSheet = ss.getSheetByName('SRS_Progress');
    var archiveSheet = ss.getSheetByName('Mastered_Cards');

    if (!progSheet) return { success: false, message: "No progress found." };
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet('Mastered_Cards');
      archiveSheet.appendRow(mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0]);
    }

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

    for (var r = mainData.length - 1; r >= 1; r--) {
      var front = mainData[r][0] ? mainData[r][0].toString().trim() : "";
      var back = mainData[r][1] ? mainData[r][1].toString().trim() : "";
      var cardId = cleanString(front) + ":::" + cleanString(back);

      if (masteredIds.has(cardId)) {
        rowsToArchive.push(mainData[r]);
        rowsToDelete.push(r + 1);
      }
    }

    if (rowsToArchive.length > 0) {
      archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
      rowsToDelete.forEach(function (rowNum) { mainSheet.deleteRow(rowNum); });
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
    var url = "https://translate.google.com/translate_tts?ie=UTF-8&q=" + encodeURIComponent(text) + "&tl=" + lang + "&client=tw-ob";
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

function migrateBoxesToIntervals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(TEMPLATE_ID);
  var sheet = ss.getSheetByName('SRS_Progress');
  if (!sheet) return "No SRS_Progress sheet found.";

  var data = sheet.getDataRange().getValues();
  var legacyIntervals = { 1: 1, 2: 3, 3: 7, 4: 14, 5: 30, 6: 60, 7: 90, 8: 180, 9: 270, 10: 360 };

  var header = data[0];
  header[2] = 'Interval';
  if (header.length < 8) header[7] = 'EF';

  for (var i = 1; i < data.length; i++) {
    var oldBox = parseFloat(data[i][2]) || 1;
    if (oldBox >= 1 && oldBox <= 10 && Number.isInteger(oldBox)) {
      data[i][2] = legacyIntervals[oldBox] || 1;
    }
    if (data[i][7] === undefined || data[i][7] === "") {
      data[i][7] = 2.5;
    }
  }

  sheet.clearContents();
  sheet.getRange(1, 1, data.length, 8).setValues(data);
  return "Migration Complete!";
}

function checkSpelling(text, lang) {
  try {
    var cleanPrompt = text ? text.toString().trim() : '';
    if (!cleanPrompt || countWords(cleanPrompt) > 20) return { success: false };

    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return { success: false };

    var models = [
      "gemini-3.6-flash",
      "gemini-3.5-flash",
      "gemini-3.1-flash-lite",
      "gemini-2.5-flash"
    ];
    
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

function archiveSingleCard(rowIndex, customDeckId) {
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    var ss = resolveSpreadsheet(customDeckId);
    var mainSheet = ss.getSheets()[0];
    var archiveSheet = ss.getSheetByName('Mastered_Cards');
    
    // Create Mastered_Cards sheet if it doesn't exist
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet('Mastered_Cards');
      archiveSheet.appendRow(mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0]);
    }
    
    var r = parseInt(rowIndex, 10);
    if (isNaN(r) || r < 2) return { success: false, error: "Invalid row index." };
    
    // Copy to archive, then delete from main
    var rowData = mainSheet.getRange(r, 1, 1, mainSheet.getLastColumn()).getValues();
    archiveSheet.appendRow(rowData[0]);
    mainSheet.deleteRow(r);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}