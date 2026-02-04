/**
 * CDCW NFC App Backend + Natural Language Data Assistant (Installable Trigger)
 * ----------------------------------------------------------------------------
 * Handles requests from the React Native App and internal Sheet queries.
 */

const SHEET_ID = '11r1v6xs4fPOQCEaEpRGGMND_YaGaZW32RK3ZGsIvnFg';
const BUILD_ID = 'v2026-02-03-ai-assistant-installable-v1';

// AI ASSISTANT CONFIG
const DASHBOARD_SHEET = "Dashboard";
const TRIGGER_CELL = "D2"; // Changed to D2 as per user snippet (Wait, strictly follow user snippet but PRD said C2 before. Snippet says D2. Wait... Snippet says: "Trigger checkbox: Dashboard!C2" in comments but code says `TRIGGER_CELL = "D2"`. Wait, user request text says: `TRIGGER_CELL = "D2"`. I will trust the code constant.)
// actually looking closer at user snippet:
// `const TRIGGER_CELL = "D2";`
// `range.getA1Notation() === TRIGGER_CELL`
// `dash.getRange(TRIGGER_CELL).setValue(false);`
// So the checkbox is at D2.

const QUERY_RANGE = "A2:C3";
const OUTPUT_RANGE_TOPLEFT = "A4";
const STATUS_CELL = "D4";
const LOG_SHEET = "Log";
const GUESTS_SHEET = "Guests";
const PIVOT_SHEET = "Pivot Table";
const LOG_LIMIT = 200;
const GUESTS_LIMIT = 200;
const PIVOT_LIMIT = 50;
const GEMINI_MODEL = "gemini-2.5-flash-lite-preview-09-2025";

/**
 * Safely executes SpreadsheetApp.openById or openByUrl based on config.
 */
function getSpreadsheet_() {
  if (!SHEET_ID) throw new Error('SHEET_ID is missing in config.');
  const idOrUrl = SHEET_ID.trim();
  if (idOrUrl.indexOf('http') === 0) {
    return SpreadsheetApp.openByUrl(idOrUrl);
  } else {
    return SpreadsheetApp.openById(idOrUrl);
  }
}

/**
 * Helper to get a map of Header Name -> Column Index (0-based)
 */
function getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, index) => {
    const key = h.toString().toLowerCase().replace(/[^a-z0-9]/g, '');
    map[key] = index;
    map[h.toString().trim().toLowerCase()] = index;
    if (key.includes('firstvisit')) {
      map['firstvisit'] = index;
    }
  });
  return map;
}

function response(data) {
  const finalData = { ...data, buildId: BUILD_ID };
  return ContentService.createTextOutput(JSON.stringify(finalData))
    .setMimeType(ContentService.MimeType.JSON);
}

function getFormattedDateTime_(dateObj) {
  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(dateObj, tz, "MM/dd/yyyy");
  const timeStr = Utilities.formatDate(dateObj, tz, "HH:mm");
  return { dateStr, timeStr, tz };
}

function isFirstVisit_(sheet, guestId, colIndex) {
  if (!guestId) return false;
  if (/^anon/i.test(String(guestId))) return true;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return true;
  const ids = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  const search = String(guestId);
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === search) {
      return false;
    }
  }
  return true;
}

// ------------------------------------------------------------------
// AI ASSISTANT LOGIC (Installable Trigger)
// ------------------------------------------------------------------

/**
 * One-time setup: install an onEdit trigger that runs under the authorizing user.
 * Run this ONCE from the editor: select installOnEditTrigger → Run
 */
function installOnEditTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Delete any existing triggers pointing at our handler to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onEditInstalled") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create installable trigger
  ScriptApp.newTrigger("onEditInstalled")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log("[Trigger] Installed onEditInstalled for spreadsheet: " + ss.getId());
}

/**
 * Installable onEdit handler.
 * Fires when Dashboard!Trigger_Cell is checked TRUE.
 */
function onEditInstalled(e) {
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (sheet.getName() === DASHBOARD_SHEET &&
      range.getA1Notation() === TRIGGER_CELL &&
      range.getValue() === true) {
    askAI_();
  }
}

/**
 * Main flow: collect data → call Gemini → write response.
 */
function askAI_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(DASHBOARD_SHEET);
  if (!dash) throw new Error(`Missing sheet: ${DASHBOARD_SHEET}`);

  // Status
  dash.getRange(STATUS_CELL).setValue("Thinking...");
  SpreadsheetApp.flush();

  try {
    // 1) Read query (Dashboard!A2:C3) via display values
    const qGrid = dash.getRange(QUERY_RANGE).getDisplayValues(); // 2D array
    const userQuery = qGrid.flat().map(s => (s || "").trim()).filter(Boolean).join(" ").trim();

    if (!userQuery) {
      throw new Error(`No question found in ${DASHBOARD_SHEET}!${QUERY_RANGE}.`);
    }

    // 2) Build context from sheets
    const logData = getSheetDataString_(ss, LOG_SHEET, LOG_LIMIT, { tail: true });
    const guestData = getSheetDataString_(ss, GUESTS_SHEET, GUESTS_LIMIT, { tail: false });
    const pivotData = getSheetDataString_(ss, PIVOT_SHEET, PIVOT_LIMIT, { tail: false });

    // 3) Construct prompt
    const context = [
      `BUILD_ID: ${BUILD_ID}`,
      "",
      "DATA SOURCE 1: RECENT LOGS",
      logData,
      "",
      "DATA SOURCE 2: GUEST SUMMARIES",
      guestData,
      "",
      "DATA SOURCE 3: MONTHLY TRENDS (Pivot)",
      pivotData
    ].join("\n");

    const promptText =
`You are a data analyst for a homeless shelter.
Use ONLY the provided data to answer accurately and compassionately.

CONTEXT:
${context}

USER QUESTION:
${userQuery}

INSTRUCTIONS:
- Answer concisely.
- If the data is insufficient, say exactly what is missing.
- Do not invent numbers or guests.`.trim();

    const payload = {
      contents: [{
        parts: [{ text: promptText }]
      }]
    };

    // 4) Call Gemini
    const apiKey = getGeminiApiKey_();
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    const raw = res.getContentText();
    let json;
    try {
      json = JSON.parse(raw);
    } catch (parseErr) {
      throw new Error(`Gemini response was not valid JSON (HTTP ${code}). Raw: ${raw}`);
    }

    if (code !== 200) {
      throw new Error(`Gemini API Error ${code}: ${json.error?.message || raw}`);
    }

    const aiResponse =
      json.candidates?.[0]?.content?.parts?.[0]?.text ||
      "No response generated.";

    // 5) Write output + status
    dash.getRange(OUTPUT_RANGE_TOPLEFT).setValue(aiResponse);
    dash.getRange(STATUS_CELL).setValue("Done");

  } catch (err) {
    dash.getRange(STATUS_CELL).setValue("Error");
    dash.getRange(OUTPUT_RANGE_TOPLEFT).setValue("Error: " + String(err));
  } finally {
    // Always uncheck trigger box to reset
    dash.getRange(TRIGGER_CELL).setValue(false);
  }
}

/**
 * Build CSV-like string from a sheet.
 * - tail=true: include header + last N rows (best for Log)
 * - tail=false: include header + first N rows (best for Guests/Pivot)
 */
function getSheetDataString_(ss, sheetName, limit, { tail } = {}) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return `[${sheetName} Sheet Missing]`;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return `[${sheetName} is Empty]`;

  // Always include header row
  const header = sheet.getRange(1, 1, 1, lastCol).getDisplayValues();

  let rows = [];
  if (lastRow > 1) {
    if (tail) {
      const startRow = Math.max(2, lastRow - limit + 1);
      rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getDisplayValues();
    } else {
      const numRows = Math.min(limit, lastRow - 1);
      rows = sheet.getRange(2, 1, numRows, lastCol).getDisplayValues();
    }
  }

  const values = header.concat(rows);

  // Convert to CSV-like lines (escape commas minimally by leaving as-is; Gemini handles it fine)
  return values.map(r => r.map(cell => (cell === null || cell === undefined) ? "" : String(cell)).join(",")).join("\n");
}

/**
 * Reads GEMINI_API_KEY from Script Properties.
 */
function getGeminiApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!key) {
    throw new Error('Missing Script Property "GEMINI_API_KEY". Add it in Project Settings → Script properties.');
  }
  return key.trim();
}

/**
 * Optional helper to quickly test UrlFetch permission from editor.
 * Run testUrlFetch() manually once after setting manifest scopes.
 */
function testUrlFetch() {
  const r = UrlFetchApp.fetch("https://www.google.com");
  Logger.log("testUrlFetch status=" + r.getResponseCode());
}

// ------------------------------------------------------------------
// APP BACKEND HANDLERS (doPost / doGet)
// ------------------------------------------------------------------

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
        return response({ status: 'error', message: 'No post data' });
    }
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const ss = getSpreadsheet_();
    
    Logger.log(`[${BUILD_ID}] Action: ${action} SS: ${ss.getId()}`);

    if (action === 'LOG_SERVICE') {
      return logService(ss, data.payload);
    } else if (action === 'UPDATE_GUEST') {
      return updateGuest(ss, data.payload);
    } else if (action === 'REPLACE_CARD') {
      return replaceCard(ss, data.payload);
    } else if (action === 'CLOTHING_PURCHASE') {
      return clothingPurchase(ss, data.payload);
    } else if (action === 'ANONYMOUS_ENTRY') {
      return anonymousEntry(ss, data.payload);
    } else if (action === 'GET_BUDGET') {
      return getBudget(ss, data.payload);
    }
    return response({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    Logger.log('Error: ' + err.toString());
    return response({ status: 'error', message: err.toString() });
  }
}

function getBudget(ss, payload) {
  const { guestId } = payload;
  const sheet = ss.getSheetByName('Guests');
  const h = getHeaderMap_(sheet);
  
  let colBucks = h['feltonbucksbudget'];
  if (colBucks === undefined) colBucks = h['feltonbucks'];
  if (colBucks === undefined) colBucks = h['feltonbucks(budget)'];
  
  const diag = { colBucksIndex: colBucks, headersFound: Object.keys(h) };
  Logger.log(`[getBudget] Diag: ${JSON.stringify(diag)}`);

  if (colBucks === undefined) {
      return response({ status: 'error', message: 'Budget Column Not Found', diag });
  }

  const colId = h['guestid'];
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
     if (String(data[i][colId]) === String(guestId)) {
         const rawVal = data[i][colBucks];
         const parsedVal = Number(rawVal);
         const budget = isNaN(parsedVal) ? 0 : parsedVal;
         Logger.log(`[getBudget] Found ${guestId} at row ${i+1}. Raw: "${rawVal}" Parsed: ${budget}`);
         return response({ status: 'success', budget: budget, raw: rawVal, row: i+1 });
     }
  }
  return response({ status: 'error', message: 'Guest Not Found', diag });
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.health === '1') {
        return response({ status: 'ok', health: true, method: 'doGet' });
    }
    const ss = getSpreadsheet_();
    const guestSheet = ss.getSheetByName('Guests');
    const data = guestSheet.getDataRange().getValues();
    const h = getHeaderMap_(guestSheet);

    const guests = {};
    const colId = h['guestid'] !== undefined ? h['guestid'] : 0;
    const colName = h['nameoptional'] !== undefined ? h['nameoptional'] : h['name'];
    const colHealth = h['healthcareprogram'];
    const colSeasonal = h['seasonalnight'];
    const colSustain = h['sustainabilityprogram'];
    
    let colBucks = h['feltonbucksbudget'];
    if (colBucks === undefined) colBucks = h['feltonbucks'];
    if (colBucks === undefined) colBucks = h['feltonbucks(budget)'];
    
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const guestId = row[colId]; 
        if (guestId) {
          guests[guestId] = {
              id: String(guestId),
              name: colName !== undefined ? row[colName] : '',
              programs: {
                  healthcare: colHealth !== undefined ? row[colHealth] === true : false,
                  seasonalNight: colSeasonal !== undefined ? row[colSeasonal] === true : false,
                  sustainability: colSustain !== undefined ? row[colSustain] === true : false
              },
              feltonBucks: colBucks !== undefined ? Number(row[colBucks]) || 0 : 0,
              lastVisit: new Date().toISOString()
          };
        }
    }
    return response({ status: 'success', guests: guests });
  } catch (err) {
      return response({ status: 'error', message: err.toString() });
  }
}

// ------------------------------------------------------------------
// BUSINESS LOGIC: CLOTHING PURCHASE
// ------------------------------------------------------------------
function clothingPurchase(ss, payload) {
  const { guestId, quantity, timestamp, eventId } = payload;
  const guestSheet = ss.getSheetByName('Guests');
  const logSheet = ss.getSheetByName('Log');
  const diag = { targetTab: 'Log', availableTabs: ss.getSheets().map(s => s.getName()) };

  if (!guestSheet) return response({ status: 'error', message: 'Guests sheet missing', diag });
  if (!logSheet) return response({ status: 'error', message: 'Log sheet missing', diag });

  const gh = getHeaderMap_(guestSheet);
  const colId = gh['guestid'];
  let colBucks = gh['feltonbucksbudget'];
  if (colBucks === undefined) colBucks = gh['feltonbucks'];
  if (colBucks === undefined) colBucks = gh['feltonbucks(budget)'];
  
  if (colBucks === undefined) {
    return response({ status: 'error', message: 'No Felton Bucks column found', diag });
  }
  
  const lock = LockService.getScriptLock();
  try {
      lock.waitLock(10000); // 10s
  } catch (e) {
      return response({ status: 'error', message: 'Server busy, try again', diag });
  }

  try {
      const gData = guestSheet.getDataRange().getValues();
      let guestRow = -1;
      let currentBudget = 0;
      for (let i = 1; i < gData.length; i++) {
        if (String(gData[i][colId]) === String(guestId)) {
          guestRow = i + 1;
          currentBudget = Number(gData[i][colBucks]) || 0;
          break;
        }
      }
      
      if (guestRow === -1) {
        lock.releaseLock();
        return response({ status: 'error', message: 'Guest not found' });
      }
      if (currentBudget < quantity) {
        lock.releaseLock();
        return response({ status: 'error', message: `Insufficient budget (Has: ${currentBudget}, Need: ${quantity})` });
      }

      // 4. Log Dedupe
      const lh = getHeaderMap_(logSheet);
      const colLogEventId = lh['eventid'];
      if (eventId && colLogEventId !== undefined) {
         const lastRow = logSheet.getLastRow();
         const startRow = Math.max(2, lastRow - 50); 
         if (lastRow > 1) {
           const checkCol = colLogEventId + 1;
           const logs = logSheet.getRange(startRow, checkCol, lastRow - startRow + 1, 1).getValues();
           for (let i = 0; i < logs.length; i++) {
             if (logs[i][0] === eventId) {
               lock.releaseLock();
               return response({ status: 'success', message: 'Deduped', diag });
             }
           }
         }
      }

      const newBudget = currentBudget - quantity;
      guestSheet.getRange(guestRow, colBucks + 1).setValue(newBudget);

      const lastColIdx = logSheet.getLastColumn();
      const newRow = new Array(lastColIdx).fill('');
      const dateObj = timestamp ? new Date(timestamp) : new Date();
      const { dateStr, timeStr, tz } = getFormattedDateTime_(dateObj);
      diag.datetimeSplit = { date: dateStr, time: timeStr, timezone: tz };
      
      const colLogDate = lh['date'] !== undefined ? lh['date'] : lh['datemmddyy'];
      const colLogTime = lh['time'] !== undefined ? lh['time'] : lh['timehhmm'];
      const colLogId = lh['guestid'];
      const colFirstVisit = lh['firstvisit'];

      let colLogClothing = lh['ofclothingitems'];
      if (colLogClothing === undefined) colLogClothing = lh['ofclothing'];
      if (colLogClothing === undefined) colLogClothing = lh['clothingitems'];
      if (colLogClothing === undefined) colLogClothing = lh['clothing'];

      if (colLogDate !== undefined) newRow[colLogDate] = dateStr;
      if (colLogTime !== undefined) newRow[colLogTime] = timeStr;
      if (colLogId !== undefined) newRow[colLogId] = guestId;
      else {
          lock.releaseLock();
          return response({ status: 'error', message: 'Guest ID column missing in Log', diag });
      }
      
      if (colFirstVisit !== undefined) {
        const isFirst = isFirstVisit_(logSheet, guestId, colLogId);
        newRow[colFirstVisit] = isFirst;
        Logger.log(`[clothingPurchase] Guest ${guestId} First Visit? ${isFirst}`);
      } else {
        lock.releaseLock();
        return response({ status: 'error', message: 'First Visit? column missing in Log', diag });
      }

      if (colLogClothing !== undefined) {
          newRow[colLogClothing] = Number(quantity);
      } else {
          lock.releaseLock();
          return response({ status: 'error', message: 'Clothing column missing in Log', diag });
      }
      if (colLogEventId !== undefined) newRow[colLogEventId] = eventId;
      
      const rowBefore = logSheet.getLastRow();
      logSheet.appendRow(newRow);
      SpreadsheetApp.flush();
      const rowAfter = logSheet.getLastRow();
      lock.releaseLock();

      if (rowAfter <= rowBefore) {
         return response({ status: 'error', message: 'Log append failed', diag });
      }

      const writtenRowVals = logSheet.getRange(rowAfter, 1, 1, lastColIdx).getValues()[0];
      const writtenQty = writtenRowVals[colLogClothing];
      
      const verify = {
          guestId: writtenRowVals[colLogId],
          clothing: writtenQty,
          dateVal: colLogDate !== undefined ? writtenRowVals[colLogDate] : null,
          timeVal: colLogTime !== undefined ? writtenRowVals[colLogTime] : null,
          firstVisit: colFirstVisit !== undefined ? writtenRowVals[colFirstVisit] : null
      };
      
      Logger.log(`[clothingPurchase] Verified Row ${rowAfter}: Date=${verify.dateVal}, Time=${verify.timeVal}, First=${verify.firstVisit}`);
      
      return response({ 
          status: 'success', 
          budget: { old: currentBudget, new: newBudget },
          log: { written: 1, row: rowAfter, clothing: writtenQty },
          verify: verify,
          diag: diag
      });
  } catch (e) {
      lock.releaseLock();
      return response({ status: 'error', message: e.toString() });
  }
}

// ------------------------------------------------------------------
// BUSINESS LOGIC: ANONYMOUS ENTRY
// ------------------------------------------------------------------
function anonymousEntry(ss, payload) {
  const { meals, timestamp } = payload;
  const logSheet = ss.getSheetByName('Log');
  const analyticsSheet = ss.getSheetByName('Analytics');
  
  const lh = getHeaderMap_(logSheet);
  const lastColIdx = logSheet.getLastColumn();
  const newRow = new Array(lastColIdx).fill('');
  const dateObj = timestamp ? new Date(timestamp) : new Date();
  
  const { dateStr, timeStr } = getFormattedDateTime_(dateObj);

  const colDate = lh['date'] !== undefined ? lh['date'] : lh['datemmddyy'];
  const colTime = lh['time'] !== undefined ? lh['time'] : lh['timehhmm'];
  const colId = lh['guestid'];
  const colMeals = lh['mealsaccessed'] !== undefined ? lh['mealsaccessed'] : lh['meals'];
  const colFirstVisit = lh['firstvisit'];

  if (colDate !== undefined) newRow[colDate] = dateStr;
  if (colTime !== undefined) newRow[colTime] = timeStr;
  if (colId !== undefined) newRow[colId] = 'anonymous';
  if (colMeals !== undefined) newRow[colMeals] = meals || 0;
  
  if (colFirstVisit !== undefined) {
      newRow[colFirstVisit] = true;
  } else {
      return response({ status: 'error', message: 'First Visit? column missing in Log' });
  }
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    logSheet.appendRow(newRow);
    SpreadsheetApp.flush();
    if (analyticsSheet) {
      const range = analyticsSheet.getRange('B1'); 
      const currentVal = range.getValue();
      const count = typeof currentVal === 'number' ? currentVal : 0;
      analyticsSheet.getRange('A1').setValue('Anonymous Unique');
      range.setValue(count + 1);
    }
    lock.releaseLock();
    return response({ status: 'success' });
  } catch (e) {
    lock.releaseLock();
    return response({ status: 'error', message: e.toString() });
  }
}

// ------------------------------------------------------------------
// BUSINESS LOGIC: LOG SERVICE (Strict + Verify)
// ------------------------------------------------------------------
function logService(ss, payload) {
  const logSheet = ss.getSheetByName('Log');
  const guestSheet = ss.getSheetByName('Guests');
  const { guestId, services, timestamp, eventId } = payload;
  
  const diag = {
    ssid: ss.getId().slice(-6),
    ssName: ss.getName(),
    targetTab: 'Log',
    availableTabs: ss.getSheets().map(s => s.getName())
  };
  Logger.log(`[logService] Diag: ${JSON.stringify(diag)}`);

  if (!logSheet) {
    Logger.log('[logService] CRITICAL: "Log" tab not found!');
    return response({ status: 'error', step: 'log_sheet_missing', diag });
  }

  const lock = LockService.getScriptLock();
  try {
      lock.waitLock(10000);
  } catch (e) {
      return response({ status: 'error', message: 'Server busy, try again', diag });
  }
  
  try {
      const h = getHeaderMap_(logSheet);
      const colTimestamp = h['date'] !== undefined ? h['date'] : h['datemmddyy']; 
      const colTime = h['time'] !== undefined ? h['time'] : h['timehhmm'];
      const colGuestId = h['guestid'];
      const colShower = h['shower'];
      const colLaundry = h['laundry'];
      const colMeals = h['mealsaccessed'] !== undefined ? h['mealsaccessed'] : h['meals'];
      const colHygiene = h['ofhygienekits'] !== undefined ? h['ofhygienekits'] : h['hygienekits'];
      const colClothing = h['ofclothing'] !== undefined ? h['ofclothing'] : h['clothing'];
      const colEventId = h['eventid'];
      const colNotes = h['staffnotes'];
      const colFirstVisit = h['firstvisit'];

      if (eventId && colEventId !== undefined) {
        const lastRow = logSheet.getLastRow();
        const startRow = Math.max(2, lastRow - 50); 
        if (lastRow > 1) {
           const checkCol = colEventId + 1; 
           const data = logSheet.getRange(startRow, checkCol, lastRow - startRow + 1, 1).getValues();
           for (let i = 0; i < data.length; i++) {
             if (data[i][0] === eventId) {
               lock.releaseLock();
               return response({ status: 'success', message: 'Deduped', diag });
             }
           }
        }
      }

      const lastColIdx = logSheet.getLastColumn();
      const newRow = new Array(lastColIdx).fill('');
      const dateObj = timestamp ? new Date(timestamp) : new Date();

      const toBool = (val) => val === true;
      const toInt = (val) => {
        if (typeof val === 'number') return Math.trunc(val);
        if (typeof val === 'string') return parseInt(val, 10) || 0;
        return 0;
      };

      let hasData = false;
      const { dateStr, timeStr, tz } = getFormattedDateTime_(dateObj);
      diag.datetimeSplit = { date: dateStr, time: timeStr, timezone: tz };

      if (colTimestamp !== undefined) { 
          newRow[colTimestamp] = dateStr;
          hasData = true; 
      }
      if (colTime !== undefined) { 
          newRow[colTime] = timeStr;
      }
      if (colGuestId !== undefined) { 
          newRow[colGuestId] = guestId; 
          hasData = true; 
      } else {
          lock.releaseLock();
          return response({ status: 'error', step: 'header_missing', field: 'Guest ID', diag });
      }

      if (colFirstVisit !== undefined) {
          const isFirst = isFirstVisit_(logSheet, guestId, colGuestId);
          newRow[colFirstVisit] = isFirst;
          Logger.log(`[logService] Guest ${guestId} First Visit? ${isFirst}`);
      } else {
          lock.releaseLock();
          return response({ status: 'error', message: 'First Visit? column missing in Log', diag });
      }

      if (colShower !== undefined) newRow[colShower] = toBool(services.shower);
      if (colLaundry !== undefined) newRow[colLaundry] = toBool(services.laundry);
      if (colMeals !== undefined) newRow[colMeals] = toInt(services.meals);
      if (colHygiene !== undefined) newRow[colHygiene] = toInt(services.hygieneKits);
      if (colClothing !== undefined) newRow[colClothing] = toInt(services.clothing);
      if (colEventId !== undefined) newRow[colEventId] = eventId;
      if (colNotes !== undefined && services.notes) newRow[colNotes] = services.notes;

      if (!hasData) {
          lock.releaseLock();
          return response({ status: 'error', step: 'empty_payload', message: 'Row would be empty. Check headers.', diag });
      }

      const rowBefore = logSheet.getLastRow();
      Logger.log(`[logService] Appending to Row ${rowBefore + 1}`);
      logSheet.appendRow(newRow);
      SpreadsheetApp.flush();
      const rowAfter = logSheet.getLastRow();
      lock.releaseLock();

      if (rowAfter <= rowBefore) {
          return response({ status: 'error', step: 'append_failed', message: 'Row count did not increase', diag });
      }

      const writtenRowVals = logSheet.getRange(rowAfter, 1, 1, lastColIdx).getValues()[0];
      const verify = {
          row: rowAfter,
          guestId: colGuestId !== undefined ? writtenRowVals[colGuestId] : null,
          meals: colMeals !== undefined ? writtenRowVals[colMeals] : null,
          shower: colShower !== undefined ? writtenRowVals[colShower] : null,
          laundry: colLaundry !== undefined ? writtenRowVals[colLaundry] : null,
          dateVal: colTimestamp !== undefined ? writtenRowVals[colTimestamp] : null,
          timeVal: colTime !== undefined ? writtenRowVals[colTime] : null,
          firstVisit: colFirstVisit !== undefined ? writtenRowVals[colFirstVisit] : null
      };
      
      Logger.log(`[logService] Verified Row ${rowAfter}: Date=${verify.dateVal}, Time=${verify.timeVal}, First=${verify.firstVisit}`);
      
      let guestUpdated = 0;
      try {
          if (guestSheet) {
              const gh = getHeaderMap_(guestSheet);
              const gColId = gh['guestid'];
              const gColVisit = gh['lastvisit'] !== undefined ? gh['lastvisit'] : gh['lastvisitdate'];
              
              if (gColId !== undefined && gColVisit !== undefined) {
                 const gData = guestSheet.getDataRange().getValues();
                 for (let i = 1; i < gData.length; i++) {
                     if (String(gData[i][gColId]) === String(guestId)) {
                         guestSheet.getRange(i + 1, gColVisit + 1).setValue(dateObj);
                         guestUpdated = 1;
                         break;
                     }
                 }
              }
          }
      } catch(e) {
          Logger.log('[logService] Guest update warning: ' + e.toString());
      }

      return response({ 
          status: 'success', 
          log: { written: 1, row: rowAfter, sheet: 'Log' }, 
          guests: { updated: guestUpdated },
          verify: verify,
          diag: diag
      });
      
  } catch (e) {
      try { lock.releaseLock(); } catch (err) {}
      return response({ status: 'error', message: e.toString() });
  }
}

function updateGuest(ss, guest) {
  const sheet = ss.getSheetByName('Guests');
  const data = sheet.getDataRange().getValues();
  const h = getHeaderMap_(sheet);

  const colId = h['guestid'];
  const colAltId = h['alternateids'];
  const colName = h['nameoptional'] !== undefined ? h['nameoptional'] : h['name'];
  const colHealth = h['healthcareprogram'];
  const colSeasonal = h['seasonalnight'];
  const colSustain = h['sustainabilityprogram'];
  
  if (colId === undefined) {
     return response({ status: 'error', message: 'Guest ID column not found' });
  }

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][colId]);
    const rowAlts = colAltId !== undefined ? String(data[i][colAltId]) : '';
    
    if (rowId === guest.id) {
      rowIndex = i + 1;
      break;
    }
    if (colAltId !== undefined && rowAlts.includes(guest.id)) {
        rowIndex = i + 1;
        break;
    }
  }
  
  if (rowIndex === -1) {
    const newRow = new Array(sheet.getLastColumn()).fill('');
    if (colId !== undefined) newRow[colId] = guest.id;
    if (colName !== undefined) newRow[colName] = guest.name || '';
    if (colHealth !== undefined) newRow[colHealth] = guest.programs?.healthcare === true;
    if (colSeasonal !== undefined) newRow[colSeasonal] = guest.programs?.seasonalNight === true;
    if (colSustain !== undefined) newRow[colSustain] = guest.programs?.sustainability === true;
    sheet.appendRow(newRow);
  } else {
    if (colName !== undefined && guest.name) sheet.getRange(rowIndex, colName + 1).setValue(guest.name);
  }
  return response({ status: 'success' });
}

function replaceCard(ss, { oldId, newId }) {
  const sheet = ss.getSheetByName('Guests');
  const data = sheet.getDataRange().getValues();
  const h = getHeaderMap_(sheet);
  const colId = h['guestid'];
  const colAlt = h['alternateids'];

  if (colId === undefined) return response({status: 'error', message: 'No Guest ID col'});

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colId]) === oldId) {
      const currentRow = i + 1;
      sheet.getRange(currentRow, colId + 1).setValue(newId);
      if (colAlt !== undefined) {
          const currentAlt = String(data[i][colAlt]);
          const newAlt = currentAlt ? currentAlt + ', ' + oldId : oldId;
          sheet.getRange(currentRow, colAlt + 1).setValue(newAlt);
      }
      return response({ status: 'success' });
    }
  }
  return response({ status: 'error', message: 'Old ID not found' });
}
