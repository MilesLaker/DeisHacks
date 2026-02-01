/**
 * CDCW NFC App Backend
 * --------------------
 * Handles requests from the React Native App.
 */

const SHEET_ID = '11r1v6xs4fPOQCEaEpRGGMND_YaGaZW32RK3ZGsIvnFg';
const BUILD_ID = 'v2026-02-01-first-visit-logic-v5';

/**
 * Safely executes SpreadsheetApp.openById or openByUrl based on config.
 * Trims whitespace and handles errors gracefully.
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
    map[h.toString().trim().toLowerCase()] = index; // Exact match fallback
    
    // Normalize "First Visit?" to "firstvisit"
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

/**
 * Checks if this is the first visit for a user.
 * Anonymous users (starts with 'anon') are ALWAYS treated as first visit (per requirements).
 * Uses targeted column scan for robustness.
 */
function isFirstVisit_(sheet, guestId, colIndex) {
  if (!guestId) return false;
  if (/^anon/i.test(String(guestId))) return true;
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return true; // Only header exists
  
  // Scans just the Guest ID column (1-based index = colIndex + 1)
  const ids = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  const search = String(guestId);
  
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === search) {
      return false;
    }
  }
  return true;
}

function doPost(e) {
  try {
    // Basic validation
    if (!e || !e.postData || !e.postData.contents) {
        return response({ status: 'error', message: 'No post data' });
    }

    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    // Test spreadsheet access immediately to fail fast
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
  
  // Robust Column Finder
  // Try: feltonbucksbudget, feltonbucks, feltonbucks(budget)
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
         
         return response({ 
             status: 'success', 
             budget: budget,
             raw: rawVal,
             row: i+1
         });
     }
  }
  
  return response({ status: 'error', message: 'Guest Not Found', diag });
}

function doGet(e) {
  try {
    // Health Check
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
    
    // Robust Budget Finder for Sync
    let colBucks = h['feltonbucksbudget'];
    if (colBucks === undefined) colBucks = h['feltonbucks'];
    if (colBucks === undefined) colBucks = h['feltonbucks(budget)'];
    
    // ... loop continues ->

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

// ... Handlers (Move these down or keep here, order doesn't matter for function declarations in JS) ...
// But for cleanliness I will include them below.

// ------------------------------------------------------------------
// CLOTHING PURCHASE (Strict + Verify)
// ------------------------------------------------------------------
function clothingPurchase(ss, payload) {
  const { guestId, quantity, timestamp, eventId } = payload;
  const guestSheet = ss.getSheetByName('Guests');
  const logSheet = ss.getSheetByName('Log');
  
  // 1. Diagnostics
  const diag = {
      targetTab: 'Log',
      availableTabs: ss.getSheets().map(s => s.getName())
  };

  // 2. Validate Sheets
  if (!guestSheet) return response({ status: 'error', message: 'Guests sheet missing', diag });
  if (!logSheet) return response({ status: 'error', message: 'Log sheet missing', diag });

  // 3. Guest Budget Lookup & Update (Critical Path)
  const gh = getHeaderMap_(guestSheet);
  const colId = gh['guestid'];
  let colBucks = gh['feltonbucksbudget'];
  if (colBucks === undefined) colBucks = gh['feltonbucks'];
  if (colBucks === undefined) colBucks = gh['feltonbucks(budget)'];
  
  if (colBucks === undefined) {
    return response({ status: 'error', message: 'No Felton Bucks column found', diag });
  }
  
  // PREPARE LOCK - We lock early to cover budget check + update + log
  const lock = LockService.getScriptLock();
  try {
      lock.waitLock(10000); // Wait up to 10s
  } catch (e) {
      return response({ status: 'error', message: 'Server busy, try again', diag });
  }

  try {
      // Re-read data under lock for consistency
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
        lock.releaseLock(); // RELEASE
        return response({ status: 'error', message: 'Guest not found' });
      }
      
      if (currentBudget < quantity) {
        lock.releaseLock(); // RELEASE
        return response({ status: 'error', message: `Insufficient budget (Has: ${currentBudget}, Need: ${quantity})` });
      }

      // 4. Log Dedupe
      const lh = getHeaderMap_(logSheet);
      const colLogEventId = lh['eventid'];
      if (eventId && colLogEventId !== undefined) {
         const lastRow = logSheet.getLastRow();
         // Optimization: Only check last 50
         const startRow = Math.max(2, lastRow - 50); 
         if (lastRow > 1) {
           const checkCol = colLogEventId + 1;
           const logs = logSheet.getRange(startRow, checkCol, lastRow - startRow + 1, 1).getValues();
           for (let i = 0; i < logs.length; i++) {
             if (logs[i][0] === eventId) {
               lock.releaseLock(); // RELEASE
               return response({ status: 'success', message: 'Deduped', diag });
             }
           }
         }
      }

      // 5. Update Budget (First)
      const newBudget = currentBudget - quantity;
      guestSheet.getRange(guestRow, colBucks + 1).setValue(newBudget);

      // 6. Build Log Row (Header Driven)
      const lastColIdx = logSheet.getLastColumn();
      const newRow = new Array(lastColIdx).fill('');
      const dateObj = timestamp ? new Date(timestamp) : new Date();
      
      // Format Date/Time Split
      const { dateStr, timeStr, tz } = getFormattedDateTime_(dateObj);
      diag.datetimeSplit = { date: dateStr, time: timeStr, timezone: tz };
      
      const colLogDate = lh['date'] !== undefined ? lh['date'] : lh['datemmddyy'];
      const colLogTime = lh['time'] !== undefined ? lh['time'] : lh['timehhmm'];
      const colLogId = lh['guestid'];
      const colFirstVisit = lh['firstvisit'];

      // Robust Clothing Header Match
      let colLogClothing = lh['ofclothingitems']; // # of clothing items
      if (colLogClothing === undefined) colLogClothing = lh['ofclothing']; // # of clothing
      if (colLogClothing === undefined) colLogClothing = lh['clothingitems'];
      if (colLogClothing === undefined) colLogClothing = lh['clothing'];

      // Fill Row
      if (colLogDate !== undefined) newRow[colLogDate] = dateStr; // Write STRING
      if (colLogTime !== undefined) newRow[colLogTime] = timeStr; // Write STRING
      if (colLogId !== undefined) newRow[colLogId] = guestId;
      else {
          lock.releaseLock(); // RELEASE
          return response({ status: 'error', message: 'Guest ID column missing in Log', diag });
      }
      
      // -- First Visit Check --
      if (colFirstVisit !== undefined) {
        const isFirst = isFirstVisit_(logSheet, guestId, colLogId);
        newRow[colFirstVisit] = isFirst;
        Logger.log(`[clothingPurchase] Guest ${guestId} First Visit? ${isFirst}`);
      } else {
        lock.releaseLock(); // RELEASE
        return response({ status: 'error', message: 'First Visit? column missing in Log', diag });
      }

      if (colLogClothing !== undefined) {
          newRow[colLogClothing] = Number(quantity);
      } else {
          lock.releaseLock(); // RELEASE
          return response({ status: 'error', message: 'Clothing column missing in Log', diag });
      }

      if (colLogEventId !== undefined) newRow[colLogEventId] = eventId;
      
      // 7. Append & Verify
      const rowBefore = logSheet.getLastRow();
      logSheet.appendRow(newRow);
      SpreadsheetApp.flush();
      const rowAfter = logSheet.getLastRow();
      
      // RELEASE LOCK NOW
      lock.releaseLock();

      if (rowAfter <= rowBefore) {
         return response({ status: 'error', message: 'Log append failed', diag });
      }

      // 8. Read-back
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
  
  // Anonymous is ALWAYS First Visit=TRUE (as per requirements)
  if (colFirstVisit !== undefined) {
      newRow[colFirstVisit] = true;
  } else {
      return response({ status: 'error', message: 'First Visit? column missing in Log' });
  }
  
  // Use Lock for safety even on anon
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
// LOG SERVICE (Strict + Verify)
// ------------------------------------------------------------------
function logService(ss, payload) {
  const logSheet = ss.getSheetByName('Log');
  const guestSheet = ss.getSheetByName('Guests');
  const { guestId, services, timestamp, eventId } = payload;
  
  // 1. Diagnostics (Prove correct Sheet/Tab)
  const diag = {
    ssid: ss.getId().slice(-6),
    ssName: ss.getName(),
    targetTab: 'Log',
    availableTabs: ss.getSheets().map(s => s.getName())
  };
  Logger.log(`[logService] Diag: ${JSON.stringify(diag)}`);

  // 2. Validate Log Sheet Exists
  if (!logSheet) {
    Logger.log('[logService] CRITICAL: "Log" tab not found!');
    return response({ status: 'error', step: 'log_sheet_missing', diag });
  }

  // LOCK EARLY
  const lock = LockService.getScriptLock();
  try {
      lock.waitLock(10000);
  } catch (e) {
      return response({ status: 'error', message: 'Server busy, try again', diag });
  }
  
  try {
      // 3. Header Mapping & Row Construction
      const h = getHeaderMap_(logSheet);
      Logger.log(`[logService] Headers found: ${JSON.stringify(h)}`);

      // Normalize column target indices
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

      // 4. Dedupe (Fast)
      if (eventId && colEventId !== undefined) {
        const lastRow = logSheet.getLastRow();
         // Check last 50 entries only for perf
        const startRow = Math.max(2, lastRow - 50); 
        if (lastRow > 1) {
           const checkCol = colEventId + 1; 
           const data = logSheet.getRange(startRow, checkCol, lastRow - startRow + 1, 1).getValues();
           for (let i = 0; i < data.length; i++) {
             if (data[i][0] === eventId) {
               // Must release lock if returning early!
               lock.releaseLock();
               return response({ status: 'success', message: 'Deduped', diag });
             }
           }
        }
      }

      // 5. Build Row
      const lastColIdx = logSheet.getLastColumn();
      const newRow = new Array(lastColIdx).fill('');
      const dateObj = timestamp ? new Date(timestamp) : new Date();

      // Helper for Types
      const toBool = (val) => val === true; // Strict boolean for checkboxes
      const toInt = (val) => {
        if (typeof val === 'number') return Math.trunc(val);
        if (typeof val === 'string') return parseInt(val, 10) || 0;
        return 0;
      };

      let hasData = false;
      
      // Format Date/Time Split
      const { dateStr, timeStr, tz } = getFormattedDateTime_(dateObj);
      diag.datetimeSplit = { date: dateStr, time: timeStr, timezone: tz };

      // -- Write Date/Time --
      if (colTimestamp !== undefined) { 
          newRow[colTimestamp] = dateStr; // Write STRING
          hasData = true; 
      }
      if (colTime !== undefined) { 
          newRow[colTime] = timeStr; // Write STRING
      }

      // -- Write Guest ID --
      if (colGuestId !== undefined) { 
          newRow[colGuestId] = guestId; 
          hasData = true; 
      } else {
          lock.releaseLock();
          return response({ status: 'error', step: 'header_missing', field: 'Guest ID', diag });
      }

      // -- First Visit Check --
      if (colFirstVisit !== undefined) {
          const isFirst = isFirstVisit_(logSheet, guestId, colGuestId);
          newRow[colFirstVisit] = isFirst;
          Logger.log(`[logService] Guest ${guestId} First Visit? ${isFirst}`);
      } else {
          // Explicit Fail for missing First Visit column
          lock.releaseLock();
          return response({ status: 'error', message: 'First Visit? column missing in Log', diag });
      }

      // -- Write Values --
      if (colShower !== undefined) newRow[colShower] = toBool(services.shower);
      if (colLaundry !== undefined) newRow[colLaundry] = toBool(services.laundry);
      if (colMeals !== undefined) newRow[colMeals] = toInt(services.meals);
      if (colHygiene !== undefined) newRow[colHygiene] = toInt(services.hygieneKits);
      if (colClothing !== undefined) newRow[colClothing] = toInt(services.clothing);
      
      if (colEventId !== undefined) newRow[colEventId] = eventId;
      if (colNotes !== undefined && services.notes) newRow[colNotes] = services.notes;

      // 6. Guard: Blank Write
      if (!hasData) {
          lock.releaseLock();
          return response({ status: 'error', step: 'empty_payload', message: 'Row would be empty. Check headers.', diag });
      }

      // 7. PERFORM WRITE
      const rowBefore = logSheet.getLastRow();
      Logger.log(`[logService] Appending to Row ${rowBefore + 1}`);
      logSheet.appendRow(newRow);
      SpreadsheetApp.flush(); // FORCE WRITE
      const rowAfter = logSheet.getLastRow();
      
      // RELEASE LOCK
      lock.releaseLock();

      if (rowAfter <= rowBefore) {
          return response({ status: 'error', step: 'append_failed', message: 'Row count did not increase', diag });
      }

      // 8. READ-BACK VERIFICATION
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
      
      // 9. Update Guests (Last Visit) - Only if Log succeeded
      let guestUpdated = 0;
      try {
          if (guestSheet) {
              const gh = getHeaderMap_(guestSheet);
              const gColId = gh['guestid'];
              const gColVisit = gh['lastvisit'] !== undefined ? gh['lastvisit'] : gh['lastvisitdate'];
              
              if (gColId !== undefined && gColVisit !== undefined) {
                 const gData = guestSheet.getDataRange().getValues();
                 for (let i = 1; i < gData.length; i++) {
                     // Fuzzy match ID
                     if (String(gData[i][gColId]) === String(guestId)) {
                         // Still write Full Date object to 'Last Visit' as that's usually a single col
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
      // Catch-all unlock
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
    // Limited updates for existing guest
    if (colName !== undefined && guest.name) sheet.getRange(rowIndex, colName + 1).setValue(guest.name);
    // Program updates logic here if needed (skipping for consistency with prompt reqs)
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
