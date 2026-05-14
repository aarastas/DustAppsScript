// Version policy:
// Bump `CODE_VERSION` when Code.gs changes.
// Index.html owns its own version constants and sends them in client calls.
// Keep the summary comments below current so future edits are traceable.
const JOURNAL_SHEET_NAME = 'Dust';
const SPECIAL_DATES_SHEET_NAME = 'SpecialDates';
const DUST_META_SHEET_NAME = 'DustMeta';
const SPECIAL_DATE_DISPLAY_MODE_PROPERTY = 'DUST_SPECIAL_DATE_DISPLAY_MODE';
const SPECIAL_DATE_DISPLAY_MODE_SPECIAL_ONLY = 'special-only';
const SPECIAL_DATE_DISPLAY_MODE_SPECIAL_AND_DEFAULT = 'special-and-default';
const PHOTO_FOLDER_NAME = 'Dust Photos';
const JOURNAL_HEADER = ['Timestamp', 'Content', 'Location', 'Date', 'Modified', 'GPSCoordinate', 'Photo'];
const PHOTO_COLUMN_INDEX = 7;
const CODE_VERSION = '1.18'; // Version 1.18: Added photo attachments stored in a photo column and Drive-backed image formulas.
const CODE_CHANGELOG = 'v1.18 | Code.gs | Added photo attachments stored in a photo column and Drive-backed image formulas.';
const SPECIAL_DATE_HEADER = ['Type', 'Label', 'RepeatAnnually', 'RuleType', 'RuleValue', 'Enabled'];
const DEFAULT_SPECIAL_DATE_ROWS = [
  { type: 'Holiday', label: "New Year's Day", ruleType: 'fixed-month-day', ruleValue: '1/1', repeatAnnually: true },
  { type: 'Holiday', label: "Martin Luther King Jr. Day", ruleType: 'nth-weekday', ruleValue: '3,1,0', repeatAnnually: true },
  { type: 'Holiday', label: "Presidents' Day", ruleType: 'nth-weekday', ruleValue: '3,1,1', repeatAnnually: true },
  { type: 'Holiday', label: 'Memorial Day', ruleType: 'last-weekday', ruleValue: '1,4', repeatAnnually: true },
  { type: 'Holiday', label: 'Independence Day', ruleType: 'fixed-month-day', ruleValue: '7/4', repeatAnnually: true },
  { type: 'Holiday', label: 'Labor Day', ruleType: 'nth-weekday', ruleValue: '1,1,8', repeatAnnually: true },
  { type: 'Holiday', label: 'Columbus Day', ruleType: 'nth-weekday', ruleValue: '2,1,9', repeatAnnually: true },
  { type: 'Holiday', label: 'Veterans Day', ruleType: 'fixed-month-day', ruleValue: '11/11', repeatAnnually: true },
  { type: 'Holiday', label: 'Thanksgiving Day', ruleType: 'nth-weekday', ruleValue: '4,4,10', repeatAnnually: true },
  { type: 'Holiday', label: 'Christmas Eve', ruleType: 'fixed-month-day', ruleValue: '12/24', repeatAnnually: true },
  { type: 'Holiday', label: 'Christmas Day', ruleType: 'fixed-month-day', ruleValue: '12/25', repeatAnnually: true },
  { type: 'Holiday', label: "New Year's Eve", ruleType: 'fixed-month-day', ruleValue: '12/31', repeatAnnually: true },
  { type: 'Holiday', label: 'Easter Sunday', ruleType: 'easter', ruleValue: '', repeatAnnually: true },
  { type: 'Holiday', label: 'General Conference Sunday', ruleType: 'nth-weekday', ruleValue: '1,0,3', repeatAnnually: true },
  { type: 'Holiday', label: 'General Conference Saturday', ruleType: 'relative', ruleValue: 'nth-weekday|1,0,3|-1', repeatAnnually: true },
  { type: 'Holiday', label: 'General Conference Sunday', ruleType: 'nth-weekday', ruleValue: '1,0,9', repeatAnnually: true },
  { type: 'Holiday', label: 'General Conference Saturday', ruleType: 'relative', ruleValue: 'nth-weekday|1,0,9|-1', repeatAnnually: true },
];

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dust Journal')
    .setFaviconUrl('https://raw.githubusercontent.com/aarastas/DustAppsScript/main/favicon-32x32.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getAppData(referenceDateInput, clientMeta) {
  return getAppData_(referenceDateInput, clientMeta);
}

function getAppData_(referenceDateInput, clientMeta) {
  const tz = Session.getScriptTimeZone();
  const referenceDate = parseDateInput_(referenceDateInput) || new Date();
  const hasJournalSheet = !!getSheetIfExists_(JOURNAL_SHEET_NAME);
  const hasSpecialDatesSheet = !!getSheetIfExists_(SPECIAL_DATES_SHEET_NAME);
  if (hasJournalSheet && hasSpecialDatesSheet) {
    const cached = getCachedAppData_(referenceDate, tz);
    if (cached) {
      return cached;
    }
  }

  const base = {
    user: { name: 'Signed in', email: '' },
    today: Utilities.formatDate(referenceDate, tz, 'yyyy-MM-dd'),
    previewDate: Utilities.formatDate(referenceDate, tz, 'yyyy-MM-dd'),
    entries: [],
    specialDates: [],
    view: {
      mode: 'empty',
      referenceDate: dateKey_(referenceDate, tz),
      title: 'No entries',
      labels: [],
      targetKey: getViewKeyText_(referenceDate, tz),
      reason: 'No journal rows were found.',
    },
  };

  try {
    base.user = getUserInfo();
  } catch (error) {}
  base.config = getAppConfig_();
  base.versions = getAppVersions_(clientMeta);

  let specialDates = [];
  try {
    specialDates = getSpecialDates_();
  } catch (error) {
    specialDates = [];
  }

  let allEntries = [];
  try {
    allEntries = getEntries_(specialDates, tz);
  } catch (error) {
    allEntries = [];
  }

  try {
    const view = buildViewContext_(referenceDate, allEntries, specialDates, tz, base.config);
    base.entries = view.entries;
    base.specialDates = specialDates;
    base.view = view.meta;
    if (hasJournalSheet && hasSpecialDatesSheet) {
      setCachedAppData_(referenceDate, tz, base);
    }
    return base;
  } catch (error) {
    base.specialDates = specialDates;
    base.entries = allEntries;
    base.view = {
      mode: 'error',
      referenceDate: dateKey_(referenceDate, tz),
      title: 'Load error',
      labels: [],
      targetKey: getViewKeyText_(referenceDate, tz),
      reason: error && error.message ? error.message : 'Failed to build view.',
    };
    if (hasJournalSheet && hasSpecialDatesSheet) {
      setCachedAppData_(referenceDate, tz, base);
    }
    return base;
  }
}

function getEntries(referenceDateInput) {
  return getAppData(referenceDateInput).entries;
}

function getSpecialDates() {
  return getSpecialDates_();
}

function getAppConfig_() {
  const props = PropertiesService.getUserProperties();
  const mode = String(props.getProperty(SPECIAL_DATE_DISPLAY_MODE_PROPERTY) || SPECIAL_DATE_DISPLAY_MODE_SPECIAL_AND_DEFAULT).trim();
  const userDisplayMode = String(props.getProperty('DUST_USER_DISPLAY_MODE') || 'username-only').trim();
  return {
    specialDateDisplayMode: mode === SPECIAL_DATE_DISPLAY_MODE_SPECIAL_ONLY
      ? SPECIAL_DATE_DISPLAY_MODE_SPECIAL_ONLY
      : SPECIAL_DATE_DISPLAY_MODE_SPECIAL_AND_DEFAULT,
    userDisplayMode: userDisplayMode === 'full' ? 'full' : 'username-only',
  };
}

function getAppVersions_(clientMeta) {
  return {
    codeVersion: CODE_VERSION,
    codeChangelog: CODE_CHANGELOG,
    indexVersion: String(clientMeta && clientMeta.indexVersion ? clientMeta.indexVersion : ''),
    indexChangelog: String(clientMeta && clientMeta.indexChangelog ? clientMeta.indexChangelog : ''),
  };
}

function setAppConfig(config, clientMeta) {
  const incoming = config || {};
  const mode = String(incoming.specialDateDisplayMode || '').trim();
  const userDisplayMode = String(incoming.userDisplayMode || '').trim();
  if (mode !== SPECIAL_DATE_DISPLAY_MODE_SPECIAL_ONLY && mode !== SPECIAL_DATE_DISPLAY_MODE_SPECIAL_AND_DEFAULT) {
    throw new Error('Invalid special date display mode.');
  }
  if (userDisplayMode !== '' && userDisplayMode !== 'full' && userDisplayMode !== 'username-only') {
    throw new Error('Invalid user display mode.');
  }

  const props = PropertiesService.getUserProperties();
  props.setProperty(SPECIAL_DATE_DISPLAY_MODE_PROPERTY, mode);
  props.setProperty('DUST_USER_DISPLAY_MODE', userDisplayMode === 'full' ? 'full' : 'username-only');
  invalidateAppDataCache_();
  syncDustMeta_(clientMeta);
  return getAppConfig_();
}

function syncDustMeta_(clientMeta) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    return;
  }

  let sheet = spreadsheet.getSheetByName(DUST_META_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(DUST_META_SHEET_NAME);
  }

  const indexVersion = String(clientMeta && clientMeta.indexVersion ? clientMeta.indexVersion : '').trim();
  const indexChangelog = String(clientMeta && clientMeta.indexChangelog ? clientMeta.indexChangelog : '').trim();
  const rows = [
    ['Key', 'Value'],
    ['CodeVersion', CODE_VERSION],
    ['CodeSummary', CODE_CHANGELOG],
    ['IndexVersion', indexVersion],
    ['IndexSummary', indexChangelog],
    ['UpdatedAt', new Date().toISOString()],
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.hideSheet();
}

function addEntry(contents, customDate, location, gpsCoordinate, photoDataUrl, clientMeta) {
  const sheet = getOrCreateSheet_(JOURNAL_SHEET_NAME);
  ensureJournalHeader_(sheet);
  const text = String(contents || '').trim();
  const resolvedLocation = String(location || '').trim();
  const resolvedGps = normalizeGpsCoordinate_(gpsCoordinate);
  const photo = savePhotoAttachment_(photoDataUrl);

  if (!text) {
    throw new Error('Entry text is required.');
  }

  const dateValue = parseDateInput_(customDate) || startOfDay_(new Date());
  const timestamp = new Date();
  const place = resolvedLocation || reverseGeocodeLocation_(resolvedGps);
  const row = [timestamp, text, place, dateValue, '', resolvedGps];

  sheet.appendRow(row);
  if (photo.url) {
    sheet.getRange(sheet.getLastRow(), PHOTO_COLUMN_INDEX).setFormula(buildPhotoImageFormula_(photo.url));
  }
  invalidateAppDataCache_();
  syncDustMeta_(clientMeta);
  return buildEntrySnapshot_(sheet.getLastRow(), timestamp, text, place, dateValue, null, resolvedGps, photo.url);
}

function reverseGeocodeLocation_(gpsCoordinate) {
  const coords = parseGpsCoordinate_(gpsCoordinate);
  if (!coords) {
    return '';
  }

  const lat = coords.latitude;
  const lon = coords.longitude;
  const cache = CacheService.getUserCache();
  const key = [
    'reverse-geocode',
    lat.toFixed(3),
    lon.toFixed(3),
  ].join('|');
  const cached = cache.get(key);
  if (cached) {
    return cached;
  }

  const label = resolveLocationDetails_(lat, lon);
  if (label) {
    try {
      cache.put(key, label, 21600);
    } catch (error) {}
  }
  return label;
}

function resolveLocationDetails(latitude, longitude) {
  const lat = Number(latitude);
  const lon = Number(longitude);
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) {
    return '';
  }

  return resolveLocationDetailsFromGeocoder_(lat, lon);
}

function resolveLocationDetailsFromGeocoder_(latitude, longitude) {
  const lat = Number(latitude);
  const lon = Number(longitude);
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) {
    return '';
  }

  try {
    const response = Maps.newGeocoder().reverseGeocode(lat, lon);
    return formatCityStateFromGeocoderResponse_(response);
  } catch (error) {
    return '';
  }
}

function updateEntry(rowNumber, contents, customDate, location, gpsCoordinate, photoDataUrl, clientMeta) {
  const sheet = getOrCreateSheet_(JOURNAL_SHEET_NAME);
  ensureJournalHeader_(sheet);

  const row = Number(rowNumber);
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
    throw new Error('Invalid journal entry row.');
  }

  const text = String(contents || '').trim();
  const resolvedLocation = String(location || '').trim();
  const resolvedGps = normalizeGpsCoordinate_(gpsCoordinate);
  const photo = savePhotoAttachment_(photoDataUrl);
  if (!text) {
    throw new Error('Entry text is required.');
  }

  const current = sheet.getRange(row, 1, 1, PHOTO_COLUMN_INDEX).getValues()[0];
  const currentFormulas = sheet.getRange(row, 1, 1, PHOTO_COLUMN_INDEX).getFormulas()[0];
  const timestamp = coerceDate_(current[0]) || new Date();
  const dateValue = parseDateInput_(customDate) || coerceDate_(current[3]) || startOfDay_(new Date());
  const modified = new Date();

  const place = resolvedLocation || String(current[2] || '').trim() || reverseGeocodeLocation_(resolvedGps || current[5]);
  const gps = resolvedGps || normalizeGpsCoordinate_(current[5]);

  sheet.getRange(row, 1, 1, 6).setValues([[timestamp, text, place, dateValue, modified, gps]]);
  if (photo.url) {
    sheet.getRange(row, PHOTO_COLUMN_INDEX).setFormula(buildPhotoImageFormula_(photo.url));
  }
  const storedPhotoUrl = photo.url || extractPhotoUrlFromCell_(current[6], currentFormulas[6]);
  invalidateAppDataCache_();
  syncDustMeta_(clientMeta);
  return buildEntrySnapshot_(row, timestamp, text, place, dateValue, modified, gps, storedPhotoUrl);
}

function addSpecialDate(labelOrDate, ruleTypeOrLabel, dateValue, ruleValue, clientMeta) {
  const sheet = getOrCreateSpecialDatesSheet_(true);
  ensureSpecialDatesHeader_(sheet);

  const text = String(labelOrDate || '').trim();
  const ruleType = String(ruleTypeOrLabel || '').trim().toLowerCase();
  const date = parseDateInput_(dateValue);
  const value = String(ruleValue || '').trim();

  if (!text) {
    throw new Error('A label is required.');
  }

  if (ruleType === 'fixed-date') {
    if (!date) {
      throw new Error('A valid date is required.');
    }
    const tz = Session.getScriptTimeZone();
    sheet.appendRow(['Holiday', text, true, 'fixed-month-day', Utilities.formatDate(date, tz, 'M/d'), true]);
    invalidateAppDataCache_();
    syncDustMeta_(clientMeta);
    return true;
  }

  if (!ruleType) {
    throw new Error('A rule type is required.');
  }

  sheet.appendRow(['Holiday', text, true, ruleType, value, true]);
  invalidateAppDataCache_();
  syncDustMeta_(clientMeta);
  return true;
}

function getUserInfo() {
  try {
    const user = Session.getActiveUser();
    const email = user && typeof user.getEmail === 'function' ? String(user.getEmail() || '') : '';
    if (!email) {
      return { name: 'Signed in', email: '' };
    }
    return {
      name: email ? email.split('@')[0] : 'Guest',
      email: email,
    };
  } catch (e) {
    return { name: 'Signed in', email: '' };
  }
}

function buildEntrySnapshot_(rowNumber, timestamp, content, location, entryDate, modified, gpsCoordinate, photoUrl) {
  const tz = Session.getScriptTimeZone();
  const date = entryDate ? startOfDay_(entryDate) : null;
  const stamp = timestamp ? new Date(timestamp) : null;
  const change = modified ? new Date(modified) : null;

  return {
    rowNumber: rowNumber,
    id: buildEntryId_(stamp || date, rowNumber),
    timestamp: stamp ? stamp.toISOString() : '',
    dateKey: date ? dateKey_(date, tz) : '',
    displayDate: formatLongDisplayDate_(date || stamp, tz),
    weekday: date ? Utilities.formatDate(date, tz, 'EEEE') : '',
    content: String(content || ''),
    location: String(location || ''),
    gpsCoordinate: normalizeGpsCoordinate_(gpsCoordinate),
    photoUrl: String(photoUrl || ''),
    modified: change ? change.toISOString() : '',
    modifiedDisplay: change ? formatDisplayDate_(change, tz) : '',
    labels: [],
    viewKey: date ? getViewKeyNumber_(date, tz) : null,
    viewKeyText: date ? getViewKeyText_(date, tz) : '',
  };
}

function getEntries_(specialDates, tz) {
  const sheet = getOrCreateSheet_(JOURNAL_SHEET_NAME);
  ensureJournalHeader_(sheet);
  const values = sheet.getDataRange().getValues();
  const formulas = sheet.getDataRange().getFormulas();

  if (!values.length) {
    return [];
  }

  const startRow = isHeaderRow_(values[0], ['timestamp', 'content', 'date']) ? 1 : 0;
  const rows = values
    .map((row, index) => ({ row: row, rowNumber: index + 1 }))
    .slice(startRow)
    .filter(item => item.row.some(cell => cell !== '' && cell !== null));

  return rows.map(item => {
    const row = item.row;
    const timestamp = coerceDate_(row[0]);
    const content = String(row[1] ?? '').trim();
    const location = String(row[2] ?? '').trim();
    const entryDate = resolveJournalEntryDate_(row[3], timestamp);
    const modified = coerceDate_(row[4]);
    const gpsCoordinate = normalizeGpsCoordinate_(row[5]);
    const photoUrl = extractPhotoUrlFromCell_(row[6], formulas[item.rowNumber - 1] ? formulas[item.rowNumber - 1][6] : '');
    const labels = entryDate ? getLabelsForDate_(entryDate, specialDates, tz) : [];
    const viewKey = entryDate ? getViewKeyNumber_(entryDate, tz) : null;
    const viewKeyText = entryDate ? getViewKeyText_(entryDate, tz) : '';

    return {
      rowNumber: item.rowNumber,
      id: buildEntryId_(timestamp || entryDate, item.rowNumber),
      timestamp: timestamp ? timestamp.toISOString() : '',
      dateKey: entryDate ? dateKey_(entryDate, tz) : '',
      displayDate: formatLongDisplayDate_(entryDate || timestamp, tz),
      weekday: entryDate ? Utilities.formatDate(entryDate, tz, 'EEEE') : '',
      content: content,
      location: location,
      gpsCoordinate: gpsCoordinate,
      photoUrl: photoUrl,
      modified: modified ? modified.toISOString() : '',
      modifiedDisplay: modified ? formatDisplayDate_(modified, tz) : '',
      labels: labels,
      viewKey: viewKey,
      viewKeyText: viewKeyText,
    };
  }).sort((a, b) => {
    const aDate = a.timestamp || a.dateKey || '';
    const bDate = b.timestamp || b.dateKey || '';
    return new Date(bDate).getTime() - new Date(aDate).getTime();
  });
}

function buildViewContext_(referenceDate, entries, specialDates, tz, config) {
  const activeLabels = getLabelsForDate_(referenceDate, specialDates, tz);
  const targetKey = getViewKeyNumber_(referenceDate, tz);
  const targetKeyText = getViewKeyText_(referenceDate, tz);
  const displayMode = getSpecialDateDisplayMode_(config);

  if (!entries.length) {
    return {
      entries: [],
      meta: {
        mode: 'empty',
        referenceDate: dateKey_(referenceDate, tz),
        title: 'No entries',
        labels: [],
        targetKey: targetKeyText,
        reason: 'No journal rows were found.',
      },
    };
  }

  if (activeLabels.length) {
    const specialMatches = entries
      .filter(entry => hasAnyLabel_(entry.labels, activeLabels))
      .map(entry => Object.assign({}, entry, { specialMatch: true }));
    if (displayMode === SPECIAL_DATE_DISPLAY_MODE_SPECIAL_ONLY) {
      return {
        entries: specialMatches,
        meta: {
          mode: SPECIAL_DATE_DISPLAY_MODE_SPECIAL_ONLY,
          referenceDate: dateKey_(referenceDate, tz),
          title: activeLabels.join(', '),
          labels: activeLabels,
          targetKey: targetKeyText,
          reason: 'Special date label override.',
        },
      };
    }

    const defaultMatches = getDefaultMatches_(entries, targetKey, targetKeyText)
      .map(entry => Object.assign({}, entry, { specialMatch: false }));
    const merged = [];
    const seen = {};
    specialMatches.concat(defaultMatches).forEach(entry => {
      const key = entry && entry.rowNumber != null ? String(entry.rowNumber) : String(entry.id || '');
      if (seen[key]) {
        return;
      }
      seen[key] = true;
      merged.push(entry);
    });

    return {
      entries: merged,
      meta: {
        mode: 'special',
        referenceDate: dateKey_(referenceDate, tz),
        title: activeLabels.join(', '),
        labels: activeLabels,
        targetKey: targetKeyText,
        reason: 'Special date labels shown first, followed by default matches.',
      },
    };
  }

  const exact = entries.filter(entry => entry.viewKeyText === targetKeyText);
  if (exact.length) {
    return {
      entries: exact,
      meta: {
        mode: 'exact',
        referenceDate: dateKey_(referenceDate, tz),
        title: formatDisplayDate_(referenceDate, tz),
        labels: [],
        targetKey: targetKeyText,
        reason: 'Exact weekday/week match.',
      },
    };
  }

  let minDiff = null;
  entries.forEach(entry => {
    if (typeof entry.viewKey !== 'number') {
      return;
    }
    const diff = Math.abs(entry.viewKey - targetKey);
    if (minDiff === null || diff < minDiff) {
      minDiff = diff;
    }
  });

  const fallback = minDiff === null
    ? []
    : entries.filter(entry => typeof entry.viewKey === 'number' && Math.abs(Math.abs(entry.viewKey - targetKey) - minDiff) < 1e-9);

  return {
    entries: fallback,
    meta: {
      mode: 'fallback',
      referenceDate: dateKey_(referenceDate, tz),
      title: formatDisplayDate_(referenceDate, tz),
      labels: [],
      targetKey: targetKeyText,
      reason: 'Closest weekday/week match found.',
    },
  };
}

function getDefaultMatches_(entries, targetKey, targetKeyText) {
  const exact = entries.filter(entry => entry.viewKeyText === targetKeyText);
  if (exact.length) {
    return exact;
  }

  let minDiff = null;
  entries.forEach(entry => {
    if (typeof entry.viewKey !== 'number') {
      return;
    }
    const diff = Math.abs(entry.viewKey - targetKey);
    if (minDiff === null || diff < minDiff) {
      minDiff = diff;
    }
  });

  return minDiff === null
    ? []
    : entries.filter(entry => typeof entry.viewKey === 'number' && Math.abs(Math.abs(entry.viewKey - targetKey) - minDiff) < 1e-9);
}

function getSpecialDateDisplayMode_(config) {
  const mode = String(config && config.specialDateDisplayMode ? config.specialDateDisplayMode : '').trim();
  return mode === SPECIAL_DATE_DISPLAY_MODE_SPECIAL_ONLY
    ? SPECIAL_DATE_DISPLAY_MODE_SPECIAL_ONLY
    : SPECIAL_DATE_DISPLAY_MODE_SPECIAL_AND_DEFAULT;
}

function resolveJournalEntryDate_(entryDateValue, fallbackTimestamp) {
  const resolved = parseDateInput_(entryDateValue);
  if (resolved) {
    return startOfDay_(resolved);
  }

  const hasValue = entryDateValue !== '' && entryDateValue !== null && entryDateValue !== undefined;
  if (hasValue) {
    return null;
  }

  return fallbackTimestamp ? startOfDay_(fallbackTimestamp) : null;
}

function getSpecialDates_() {
  const sheet = getOrCreateSpecialDatesSheet_(true);
  ensureSpecialDatesHeader_(sheet);

  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    return [];
  }

  const rows = values.slice(1).filter(row => row.some(cell => cell !== '' && cell !== null));

  const seen = {};
  return rows
    .map(parseSpecialDateRow_)
    .filter(Boolean)
    .filter(item => {
      const key = specialDateKey_(item);
      if (seen[key]) {
        return false;
      }
      seen[key] = true;
      return true;
    });
}

function getLabelsForDate_(date, specialDates, tz) {
  const labels = [];
  const key = dateKey_(date, tz);
  const monthDay = monthDayKey_(date, tz);

  specialDates.forEach(item => {
    if (isSpecialDateActiveFor_(item, date, key, monthDay, tz)) {
      labels.push(item.label);
    }
  });

  return uniqueStrings_(labels);
}

function hasAnyLabel_(entryLabels, activeLabels) {
  if (!entryLabels || !entryLabels.length || !activeLabels || !activeLabels.length) {
    return false;
  }

  return entryLabels.some(label => activeLabels.indexOf(label) !== -1);
}

function getEasterSunday_(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;

  return new Date(year, month - 1, day);
}

function isNthWeekdayOfMonth_(date, nth, weekday, monthIndex) {
  return date.getMonth() === monthIndex &&
    date.getDay() === weekday &&
    date.getDate() >= 1 + ((nth - 1) * 7) &&
    date.getDate() <= (nth * 7);
}

function isSpecialDateActiveFor_(item, date, dateKey, monthDayKey, tz) {
  if (!item || item.enabled === false) {
    return false;
  }

  if (String(item.type || '').toLowerCase() === 'holiday') {
    return isHolidayRuleMatch_(item, date, tz);
  }

  if (item.repeatAnnually && item.monthDayKey === monthDayKey) {
    return true;
  }

  if (item.dateKey === dateKey) {
    return true;
  }

  return false;
}

function isHolidayRuleMatch_(item, date, tz) {
  const ruleType = String(item.ruleType || '').toLowerCase();
  const ruleValue = String(item.ruleValue || '').trim();
  if (ruleType === 'conference-weekend') {
    return isConferenceWeekendMatch_(date, ruleValue);
  }
  const target = getHolidayRuleDate_(date.getFullYear(), ruleType, ruleValue, tz);
  return target ? isSameDate_(date, target) : false;
}

function getHolidayRuleDate_(year, ruleType, ruleValue, tz) {
  const type = String(ruleType || '').toLowerCase();
  const value = String(ruleValue || '').trim();
  const timeZone = tz || Session.getScriptTimeZone();

  if (type === 'fixed-month-day') {
    const parts = value.split('/').map(part => Number(part.trim()));
    if (parts.length >= 2 && !parts.some(num => Number.isNaN(num))) {
      return new Date(year, parts[0] - 1, parts[1]);
    }
    const parsed = parseDateInput_(value);
    if (!parsed) {
      return null;
    }
    return new Date(year, parsed.getMonth(), parsed.getDate());
  }

  if (type === 'easter') {
    return getEasterSunday_(year);
  }

  if (type === 'nth-weekday') {
    const parts = value.split(',').map(part => Number(part.trim()));
    if (parts.length !== 3 || parts.some(num => Number.isNaN(num))) {
      return null;
    }
    return getNthWeekdayOfMonth_(year, parts[2], parts[1], parts[0]);
  }

  if (type === 'last-weekday') {
    const parts = value.split(',').map(part => Number(part.trim()));
    if (parts.length !== 2 || parts.some(num => Number.isNaN(num))) {
      return null;
    }
    return getLastWeekdayOfMonth_(year, parts[1], parts[0]);
  }

  if (type === 'relative') {
    const relative = decodeRelativeRuleValue_(value);
    if (!relative) {
      return null;
    }
    const base = getHolidayRuleDate_(year, relative.baseType, relative.baseValue, timeZone);
    if (!base) {
      return null;
    }
    const offsetDays = Number(relative.offsetDays || 0);
    if (Number.isNaN(offsetDays)) {
      return null;
    }
    const result = new Date(base);
    result.setDate(result.getDate() + offsetDays);
    return result;
  }

  return null;
}

function isLastWeekdayOfMonth_(date, weekday, monthIndex) {
  if (date.getMonth() !== monthIndex || date.getDay() !== weekday) {
    return false;
  }

  const nextWeek = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 7);
  return nextWeek.getMonth() !== monthIndex;
}

function getNthWeekdayOfMonth_(year, monthIndex, weekday, nth) {
  const date = new Date(year, monthIndex, 1);
  while (date.getDay() !== weekday) {
    date.setDate(date.getDate() + 1);
  }
  date.setDate(date.getDate() + (nth - 1) * 7);
  return date.getMonth() === monthIndex ? date : null;
}

function getLastWeekdayOfMonth_(year, monthIndex, weekday) {
  const date = new Date(year, monthIndex + 1, 0);
  while (date.getDay() !== weekday) {
    date.setDate(date.getDate() - 1);
  }
  return date;
}

function getViewKeyNumber_(date, tz) {
  const week = getWeekOfYear_(date, tz);
  const weekdayFraction = date.getDay() / 10;
  return week + weekdayFraction;
}

function getViewKeyText_(date, tz) {
  const week = pad2_(getWeekOfYear_(date, tz));
  return week + '.' + date.getDay();
}

function getWeekOfYear_(date, tz) {
  const local = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const jan1 = new Date(local.getFullYear(), 0, 1);
  const dayOfYear = Math.floor((local - jan1) / 86400000) + 1;
  return Math.ceil((dayOfYear + jan1.getDay()) / 7);
}

function isSameDate_(a, b) {
  return a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate();
}

function buildEntryId_(date, index) {
  if (!date) {
    return String(index);
  }

  return [
    date.getFullYear(),
    pad2_(date.getMonth() + 1),
    pad2_(date.getDate()),
    pad2_(date.getHours()),
    pad2_(date.getMinutes()),
    pad2_(date.getSeconds()),
    index,
  ].join('-');
}

function formatDisplayDate_(date, tz) {
  if (!date) {
    return '';
  }

  return Utilities.formatDate(date, tz, 'EEE, MMM d, yyyy');
}

function formatLongDisplayDate_(date, tz) {
  if (!date) {
    return '';
  }

  return Utilities.formatDate(date, tz, 'MMMM d, yyyy');
}

function dateKey_(date, tz) {
  return Utilities.formatDate(date, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function monthDayKey_(date, tz) {
  return Utilities.formatDate(date, tz || Session.getScriptTimeZone(), 'MM-dd');
}

function parseSpecialDateRow_(row) {
  if (!row || !row.length) {
    return null;
  }

  const first = String(row[0] ?? '').trim();
  const second = String(row[1] ?? '').trim();
  const third = row[2];
  const fourth = row[3];
  const fifth = row[4];
  const sixth = row[5];

  if (first.toLowerCase() === 'personal' || first.toLowerCase() === 'holiday') {
    if (!second) {
      return null;
    }
    const rawRuleValue = String(fifth ?? '').trim();
    const item = {
      type: first || 'Personal',
      label: second,
      repeatAnnually: toBoolean_(third),
      ruleType: String(fourth ?? '').trim(),
      ruleValue: rawRuleValue,
      enabled: row.length < 6 ? true : toBoolean_(sixth),
    };
    const parsed = parseDateInput_(rawRuleValue);
    if (parsed) {
      const tz = Session.getScriptTimeZone();
      if (String(item.ruleType || '').toLowerCase() === 'fixed-month-day') {
        item.ruleValue = Utilities.formatDate(parsed, tz, 'M/d');
      }
      item.dateKey = dateKey_(parsed, tz);
      item.monthDayKey = monthDayKey_(parsed, tz);
    }
    return item;
  }

  return null;
}

function parseDateInput_(value) {
  if (!value) {
    return null;
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return startOfDay_(value);
  }

  if (typeof value === 'number') {
    return spreadsheetSerialToDate_(value);
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }

    const match = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (match) {
      return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    }

    const slashOrDotMatch = trimmed.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/);
    if (slashOrDotMatch) {
      return new Date(Number(slashOrDotMatch[3]), Number(slashOrDotMatch[1]) - 1, Number(slashOrDotMatch[2]));
    }

    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) {
      return startOfDay_(parsed);
    }
  }

  return null;
}

function coerceDate_(value) {
  return parseDateInput_(value);
}

function spreadsheetSerialToDate_(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  return new Date(utcValue * 1000);
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatCityStateFromNominatim_(address) {
  return formatCityStateFromAddressObject_(address);
}

function formatCityStateFromGeocoderResponse_(response) {
  if (!response || !response.results || !response.results.length) {
    return '';
  }

  for (let i = 0; i < response.results.length; i++) {
    const value = formatCityStateFromGeocoderResult_(response.results[i]);
    if (value) {
      return value;
    }
  }

  return '';
}

function formatCityStateFromGeocoderResult_(result) {
  if (!result) {
    return '';
  }

  if (result.address_components) {
    const value = formatCityStateFromAddressComponents_(result.address_components);
    if (value) {
      return value;
    }
  }

  if (result.address) {
    const value = formatCityStateFromAddressObject_(result.address);
    if (value) {
      return value;
    }
  }

  if (result.formatted_address) {
    return formatCityStateFromFormattedAddress_(result.formatted_address);
  }

  return '';
}

function formatCityStateFromAddressComponents_(components) {
  if (!components || !components.length) {
    return '';
  }

  let city = '';
  let state = '';
  components.forEach(component => {
    const types = Array.isArray(component.types) ? component.types : [];
    if (!city && (types.indexOf('locality') !== -1 || types.indexOf('postal_town') !== -1 || types.indexOf('sublocality') !== -1)) {
      city = String(component.long_name || component.short_name || '').trim();
    }
    if (!state && types.indexOf('administrative_area_level_1') !== -1) {
      state = String(component.short_name || component.long_name || '').trim();
    }
  });

  return joinCityState_(city, state);
}

function formatCityStateFromAddressObject_(address) {
  if (!address) {
    return '';
  }

  const city = String(
    address.city ||
    address.town ||
    address.village ||
    address.municipality ||
    address.hamlet ||
    address.locality ||
    address.county ||
    ''
  ).trim();

  let state = String(address.state_code || address.state || '').trim();
  if (state && state.indexOf('-') !== -1) {
    state = state.split('-').pop().trim();
  }

  return joinCityState_(city, state);
}

function formatCityStateFromFormattedAddress_(value) {
  const text = String(value || '').trim();
  if (!text) {
    return '';
  }

  const parts = text.split(',').map(part => part.trim()).filter(Boolean);
  if (parts.length >= 2) {
    const city = parts[0];
    const state = String(parts[1].split(/\s+/)[0] || '').trim();
    return joinCityState_(city, state);
  }

  return text;
}

function joinCityState_(city, state) {
  if (city && state) {
    return city + ', ' + state;
  }
  return city || state || '';
}

function normalizeGpsCoordinate_(value) {
  const coords = parseGpsCoordinate_(value);
  if (!coords) {
    return '';
  }
  return coords.latitude.toFixed(6) + ', ' + coords.longitude.toFixed(6);
}

function savePhotoAttachment_(photoDataUrl) {
  const text = String(photoDataUrl || '').trim();
  if (!text) {
    return { url: '' };
  }

  const match = text.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) {
    return { url: '' };
  }

  const mimeType = String(match[1] || '').trim().toLowerCase();
  if (!/^image\/(png|jpeg|jpg|gif|webp|bmp|heic|heif)$/i.test(mimeType)) {
    return { url: '' };
  }

  const base64 = String(match[2] || '').trim();
  if (!base64) {
    return { url: '' };
  }

  const bytes = Utilities.base64Decode(base64);
  const extension = mimeTypeToExtension_(mimeType);
  const folder = getOrCreatePhotoFolder_();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const file = folder.createFile(Utilities.newBlob(bytes, mimeType, `Dust-${stamp}.${extension}`));
  const url = `https://drive.google.com/uc?export=view&id=${file.getId()}`;
  return {
    url: url,
    fileId: file.getId(),
    name: file.getName(),
  };
}

function getOrCreatePhotoFolder_() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function mimeTypeToExtension_(mimeType) {
  const type = String(mimeType || '').toLowerCase();
  if (type === 'image/jpeg' || type === 'image/jpg') return 'jpg';
  if (type === 'image/png') return 'png';
  if (type === 'image/gif') return 'gif';
  if (type === 'image/webp') return 'webp';
  if (type === 'image/bmp') return 'bmp';
  if (type === 'image/heic') return 'heic';
  if (type === 'image/heif') return 'heif';
  return 'img';
}

function buildPhotoImageFormula_(url) {
  const safeUrl = String(url || '').replace(/"/g, '""');
  if (!safeUrl) {
    return '';
  }
  return `=IMAGE("${safeUrl}")`;
}

function extractPhotoUrlFromCell_(value, formula) {
  const formulaText = String(formula || '').trim();
  if (formulaText) {
    const match = formulaText.match(/^=IMAGE\("([^"]+)"/i);
    if (match) {
      return match[1].replace(/""/g, '"');
    }
  }

  const text = String(value || '').trim();
  if (!text) {
    return '';
  }

  if (text.indexOf('http://') === 0 || text.indexOf('https://') === 0) {
    return text;
  }

  return '';
}

function parseGpsCoordinate_(value) {
  if (!value) {
    return null;
  }

  if (typeof value === 'object' && value.latitude != null && value.longitude != null) {
    const latObject = Number(value.latitude);
    const lonObject = Number(value.longitude);
    if (Number.isFinite(latObject) && Number.isFinite(lonObject)) {
      return { latitude: latObject, longitude: lonObject };
    }
  }

  const text = String(value || '').trim();
  if (!text) {
    return null;
  }

  const match = text.match(/^(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)$/);
  if (!match) {
    return null;
  }

  const latitude = Number(match[1]);
  const longitude = Number(match[2]);
  if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) {
    return null;
  }

  return { latitude: latitude, longitude: longitude };
}

function pad2_(value) {
  return String(value).padStart(2, '0');
}

function toBoolean_(value) {
  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'number') {
    return value !== 0;
  }

  if (typeof value === 'string') {
    return /^(true|yes|y|1)$/i.test(value.trim());
  }

  return false;
}

function uniqueStrings_(items) {
  const seen = {};
  return items.filter(item => {
    if (!item || seen[item]) {
      return false;
    }
    seen[item] = true;
    return true;
  });
}

function getSheetOrThrow_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error('Missing sheet: ' + name);
  }
  return sheet;
}

function getSheetIfExists_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getCachedAppData_(referenceDate, tz) {
  const cache = CacheService.getUserCache();
  const key = getAppDataCacheKey_(referenceDate, tz);
  const cached = cache.get(key);
  if (!cached) {
    return null;
  }

  try {
    return JSON.parse(cached);
  } catch (error) {
    return null;
  }
}

function setCachedAppData_(referenceDate, tz, data) {
  const cache = CacheService.getUserCache();
  const key = getAppDataCacheKey_(referenceDate, tz);
  try {
    cache.put(key, JSON.stringify(data), 300);
  } catch (error) {}
}

function getAppDataCacheKey_(referenceDate, tz) {
  return [
    'app-data',
    dateKey_(referenceDate, tz),
    String(tz || Session.getScriptTimeZone()),
    getAppDataCacheVersion_(),
  ].join('|');
}

function getAppDataCacheVersion_() {
  const props = PropertiesService.getScriptProperties();
  return String(props.getProperty('APP_DATA_CACHE_VERSION') || '0');
}

function invalidateAppDataCache_() {
  const props = PropertiesService.getScriptProperties();
  const current = Number(props.getProperty('APP_DATA_CACHE_VERSION') || '0');
  props.setProperty('APP_DATA_CACHE_VERSION', String(current + 1));
}

function getOrCreateSpecialDatesSheet_(seedDefaults) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SPECIAL_DATES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SPECIAL_DATES_SHEET_NAME);
    ensureSpecialDatesHeader_(sheet);
    if (seedDefaults) {
      seedDefaultHolidayRows_(sheet);
    }
    return sheet;
  }

  ensureSpecialDatesHeader_(sheet);
  return sheet;
}

function ensureSpecialDatesHeader_(sheet) {
  const header = SPECIAL_DATE_HEADER.slice();
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    return;
  }

  const row = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  const current = row.map(value => String(value || '').trim().toLowerCase());
  if (current[0] === 'type' && current[1] === 'label' && current[2] === 'repeatannually') {
    return;
  }

  sheet.getRange(1, 1, 1, header.length).setValues([header]);
}

function ensureJournalHeader_(sheet) {
  const header = JOURNAL_HEADER.slice();
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    return;
  }

  const row = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  const current = row.map(value => String(value || '').trim());
  const hasHeader = current[0] === 'Timestamp' || current[1] === 'Content' || current[2] === 'Date';

  if (!hasHeader) {
    return;
  }

  if (sheet.getLastColumn() < header.length || String(current[5] || '').trim() !== 'GPSCoordinate' || String(current[6] || '').trim() !== 'Photo') {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
}

function isHeaderRow_(row, expectedTerms) {
  const values = row.slice(0, 4).map(value => String(value || '').trim().toLowerCase());
  return expectedTerms.some(term => values.includes(term));
}

function seedDefaultHolidayRows_(sheet) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const values = sheet.getDataRange().getValues();
    const existing = {};

    values.slice(1).forEach(row => {
      const parsed = parseSpecialDateRow_(row);
      if (!parsed) {
        return;
      }

      const key = specialDateSeedKey_(parsed);
      existing[key] = true;
    });

    const rowsToAdd = DEFAULT_SPECIAL_DATE_ROWS.filter(item => {
      const key = specialDateSeedKey_({
        type: item.type,
        label: item.label,
        dateKey: '',
        ruleType: item.ruleType,
        ruleValue: item.ruleValue,
        repeatAnnually: item.repeatAnnually,
        enabled: true,
      });
      return !existing[key];
    }).map(item => ([
      item.type,
      item.label,
      item.repeatAnnually !== false,
      item.ruleType || '',
      item.ruleValue || '',
      true,
    ]));

    if (!rowsToAdd.length) {
      return;
    }

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAdd.length, SPECIAL_DATE_HEADER.length).setValues(rowsToAdd);
  } finally {
    lock.releaseLock();
  }
}

function specialDateKey_(item) {
  return [
    String(item && item.type ? item.type : '').toLowerCase(),
    String(item && item.label ? item.label : '').toLowerCase(),
    String(item && item.dateKey ? item.dateKey : '').toLowerCase(),
    String(item && item.ruleType ? item.ruleType : '').toLowerCase(),
    String(item && item.ruleValue ? item.ruleValue : '').toLowerCase(),
  ].join('|');
}

function specialDateSeedKey_(item) {
  const type = String(item && item.type ? item.type : '').toLowerCase();
  const ruleType = String(item && item.ruleType ? item.ruleType : '').toLowerCase();
  const ruleValue = String(item && item.ruleValue ? item.ruleValue : '').toLowerCase();

  if (type === 'holiday' || ruleType) {
    return ['holiday', ruleType, ruleValue].join('|');
  }

  return specialDateKey_(item);
}

function isConferenceWeekendMatch_(date, ruleValue) {
  const value = String(ruleValue || '').trim().toLowerCase();
  const monthIndex = conferenceMonthIndex_(value);
  if (monthIndex === null) {
    return false;
  }

  const firstSunday = getFirstSundayOfMonth_(date.getFullYear(), monthIndex);
  const saturday = new Date(firstSunday);
  saturday.setDate(firstSunday.getDate() - 1);

  return isSameDate_(date, saturday) || isSameDate_(date, firstSunday);
}

function conferenceMonthIndex_(ruleValue) {
  const value = String(ruleValue || '').trim().toLowerCase();
  if (!value) {
    return null;
  }

  if (value === 'april' || value === '4') {
    return 3;
  }

  if (value === 'october' || value === '10') {
    return 9;
  }

  const parsed = Number(value);
  if (!Number.isNaN(parsed) && parsed >= 1 && parsed <= 12) {
    return parsed - 1;
  }

  return null;
}

function decodeRelativeRuleValue_(value) {
  const parts = String(value || '').split('|');
  if (parts.length < 3) {
    return null;
  }

  return {
    baseType: String(parts[0] || '').trim().toLowerCase(),
    baseValue: String(parts[1] || '').trim(),
    offsetDays: String(parts[2] || '').trim(),
  };
}

function getFirstSundayOfMonth_(year, monthIndex) {
  const date = new Date(year, monthIndex, 1);
  while (date.getDay() !== 0) {
    date.setDate(date.getDate() + 1);
  }
  return date;
}
