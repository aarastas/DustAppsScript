/**
 * Dust Journal - Modified for Today / Yesterday / Tomorrow sheets
 * Looks in "Today" → "Yesterday" → "Tomorrow"
 * Column A: Date, Column B: Entry
 */


const DISPLAY_SHEETS = ['Today', 'Yesterday', 'Tomorrow'];
const WRITE_SHEET = 'Dust';
const PROPERTIES_KEY = 'DUST_JOURNAL_SPREADSHEET_ID';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dust Journal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSavedSpreadsheetId() {
  return PropertiesService.getUserProperties().getProperty(PROPERTIES_KEY);
}

function saveSpreadsheetId(spreadsheetId) {
  try {
    let cleanId = spreadsheetId.trim();
    if (cleanId.includes('docs.google.com/spreadsheets')) {
      const match = cleanId.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match) cleanId = match[1];
    }
    const ss = SpreadsheetApp.openById(cleanId);
    PropertiesService.getUserProperties().setProperty(PROPERTIES_KEY, cleanId);
    return {
      success: true,
      name: ss.getName(),
      url: ss.getUrl()
    };
  } catch (e) {
    return { success: false, message: 'Invalid spreadsheet: ' + e.message };
  }
}

function clearSpreadsheetId() {
  PropertiesService.getUserProperties().deleteProperty(PROPERTIES_KEY);
  return { success: true };
}

function getSpreadsheet() {
  const savedId = getSavedSpreadsheetId();
  if (!savedId) throw new Error('NO_SPREADSHEET_SELECTED');
  return SpreadsheetApp.openById(savedId);
}

function getActiveSheetWithEntries() {
  const ss = getSpreadsheet();
  for (let name of DISPLAY_SHEETS) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    if (values.some(row => row[1] && String(row[1]).trim())) {
      return { sheet, name };
    }
  }
  return null;
}

function getEntries() {
  const ss = getSpreadsheet();
  const dustSheet = ss.getSheetByName(WRITE_SHEET);
  if (!dustSheet) throw new Error(`Sheet "${WRITE_SHEET}" not found`);

  // Get all Dust data once (for lookup)
  const dustLastRow = dustSheet.getLastRow();
  let dustData = [];
  if (dustLastRow > 1) {
    dustData = dustSheet.getRange(2, 1, dustLastRow - 1, 2).getValues();
  }

  // Find the display sheet with content
  let displaySheet = null;
  let displayName = '';
  for (let name of DISPLAY_SHEETS) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    if (values.some(row => row[1] && String(row[1]).trim())) {
      displaySheet = sheet;
      displayName = name;
      break;
    }
  }

  if (!displaySheet) return [];

  const lastRow = displaySheet.getLastRow();
  const displayData = displaySheet.getRange(2, 1, lastRow - 1, 2).getValues();

  return displayData
    .map((row, index) => {
      const dateRaw = row[0];
      const content = String(row[1] || '').trim();
      if (!content) return null;

      let dateStr = dateRaw instanceof Date 
        ? Utilities.formatDate(dateRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd') 
        : String(dateRaw || '').trim();

      // Try to find matching row in Dust by date + content
      let dustRowNum = null;
      for (let i = 0; i < dustData.length; i++) {
        const dRow = dustData[i];
        const dDate = dRow[0] instanceof Date 
          ? Utilities.formatDate(dRow[0], Session.getScriptTimeZone(), 'yyyy-MM-dd') 
          : String(dRow[0] || '').trim();
        const dContent = String(dRow[1] || '').trim();

        if (dDate === dateStr && dContent === content) {
          dustRowNum = i + 2; // 1-based row number in Dust
          break;
        }
      }

      return {
        id: index + 2,               // display sheet row (for reference)
        dustId: dustRowNum,          // ← the important one: real row in Dust
        date: dateStr || '—',
        content,
        sheetName: displayName
      };
    })
    .filter(Boolean)
    .reverse();
}

function getEntryById(dustId) {
  try {
    const ss = getSpreadsheet();
    const dustSheet = ss.getSheetByName(WRITE_SHEET);
    if (!dustSheet) throw new Error(`Sheet "${WRITE_SHEET}" not found`);

    const rowNum = parseInt(dustId);
    if (rowNum < 2 || rowNum > dustSheet.getLastRow()) {
      throw new Error('Entry not found');
    }

    const row = dustSheet.getRange(rowNum, 1, 1, 2).getValues()[0];
    let dateStr = '';
    if (row[0] instanceof Date) {
      dateStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof row[0] === 'string') {
      dateStr = row[0].trim();
    }

    return {
      id: dustId,                    // now it's the Dust row number
      date: dateStr,
      content: String(row[1] || '').trim()
    };
  } catch (e) {
    throw new Error('Cannot load entry: ' + e.message);
  }
}

function addEntry(content, dateStr) {
  try {
    if (!content?.trim()) throw new Error('Content cannot be empty');

    const ss = getSpreadsheet();
    const dustSheet = ss.getSheetByName(WRITE_SHEET);
    if (!dustSheet) {
      throw new Error(`Sheet "${WRITE_SHEET}" not found. Please create it.`);
    }

    let dateValue = '';
    if (dateStr) {
      const d = new Date(dateStr);
      if (!isNaN(d.getTime())) dateValue = d;
    }

    dustSheet.appendRow([
      dateValue || new Date(),
      content.trim()
    ]);

    return { success: true };
  } catch (e) {
    throw new Error('Add failed: ' + e.message);
  }
}

function updateEntry(dustId, content, dateStr) {
  try {
    if (!content?.trim()) throw new Error('Content cannot be empty');

    const ss = getSpreadsheet();
    const dustSheet = ss.getSheetByName(WRITE_SHEET);
    if (!dustSheet) throw new Error(`Sheet "${WRITE_SHEET}" not found`);

    const rowNum = parseInt(dustId);
    if (rowNum < 2 || rowNum > dustSheet.getLastRow()) {
      throw new Error('Entry not found in Dust tab');
    }

    // Update Column B: entry content (always)
    dustSheet.getRange(rowNum, 2).setValue(content.trim());

    // Update Column E: custom/modified date (only if user provided one)
    if (dateStr) {
      const d = new Date(dateStr);
      if (!isNaN(d.getTime())) {
        dustSheet.getRange(rowNum, 5).setValue(d);  // Column E = 5
      } else {
        // Optional: clear Column E if invalid date provided
        dustSheet.getRange(rowNum, 5).clearContent();
      }
    }
    // If no dateStr provided → leave Column E unchanged (do nothing)

    // Column A is NEVER touched

    return { success: true };
  } catch (e) {
    throw new Error('Update failed: ' + e.message);
  }
}

function deleteEntry(dustId) {
  try {
    const ss = getSpreadsheet();
    const dustSheet = ss.getSheetByName(WRITE_SHEET);
    if (!dustSheet) throw new Error(`Sheet "${WRITE_SHEET}" not found`);

    const rowNum = parseInt(dustId);
    if (rowNum < 2 || rowNum > dustSheet.getLastRow()) {
      throw new Error('Entry not found in Dust tab');
    }

    dustSheet.deleteRow(rowNum);
    return { success: true };
  } catch (e) {
    throw new Error('Delete failed: ' + e.message);
  }
}

function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getUserInfo() {
  const email = Session.getActiveUser().getEmail();
  return {
    email,
    name: email.split('@')[0] || 'User'
  };
}

function logout() {
  if (confirm('Are you sure you want to log out?')) {
    // Attempt to close (works sometimes in popups)
    window.close();
    // Fallback: redirect to blank or Google
    if (window.location.href.indexOf('googleusercontent.com') !== -1) {
      window.location.href = 'https://accounts.google.com/Logout';
    }
  }
}

function getCurrentSpreadsheetInfo() {
  const id = getSavedSpreadsheetId();
  if (!id) return { selected: false };
  try {
    const ss = SpreadsheetApp.openById(id);
    return {
      selected: true,
      id,
      name: ss.getName(),
      url: ss.getUrl()
    };
  } catch (e) {
    PropertiesService.getUserProperties().deleteProperty(PROPERTIES_KEY);
    return { selected: false };
  }
}
