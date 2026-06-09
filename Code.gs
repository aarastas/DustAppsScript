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
const JOURNAL_HEADER = ['Timestamp', 'Content', 'Location', 'Date', 'Modified', 'GPSCoordinate', 'Photo', 'Tag'];
const PHOTO_COLUMN_INDEX = 7;
const TAG_COLUMN_INDEX = 8;
const CODE_VERSION = '1.29'; // Version 1.29: Raised parsed-data cache TTL to 6h and memoized holiday rule dates per year.
const CODE_CHANGELOG = 'v1.29 | Code.gs | Raised parsed-data cache TTL to 6h and memoized holiday rule dates per year.';
const PARSED_JOURNAL_CACHE_PREFIX = 'parsed-journal';
const PARSED_SPECIAL_DATES_CACHE_PREFIX = 'parsed-special-dates';
const PARSED_DATA_CACHE_TTL_SECONDS = 21600; // 6h; writes invalidate via cache buster, so a long TTL is safe.
// In-memory memo for resolved holiday rule dates within a single execution.
// Keyed by `year|ruleType|ruleValue`. Cleared automatically when the script
// instance ends, so it never goes stale across edits.
const HOLIDAY_RULE_DATE_MEMO_ = {};
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
  { type: '
