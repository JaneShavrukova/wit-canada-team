// ============================================================
// WIT Canada — Utils
// Shared helper functions (no business logic)
// ============================================================


// ─────────────────────────────────────────────
// Sheets helpers
// ─────────────────────────────────────────────

function getMainSheet() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.MAIN);
}

function getSheet(name) {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(name);
}


// ─────────────────────────────────────────────
// Column mapping (SINGLE SOURCE)
// ─────────────────────────────────────────────

function normalizeHeader(value) {
  return value
    ?.toString()
    .replace(/\u00A0/g, ' ')   // fix non-breaking space
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function getColumnIndexMap(sheet) {
  const headers = sheet
    .getRange(CONFIG.SHEET.HEADER_ROW, 1, 1, sheet.getMaxColumns())
    .getValues()[0];

  const map = {};

  headers.forEach((header, index) => {
    const key = normalizeHeader(header);
    if (key) {
      map[key] = index + 1;
    }
  });

  return map;
}

// ─────────────────────────────────────────────
// Column access helper (SAFE)
// ─────────────────────────────────────────────

function getCol(colMap, key) {
  const normalizedKey = normalizeHeader(key);
  const col = colMap[normalizedKey];

  if (!col) {
    throw new Error(`Column not found: "${key}"`);
  }

  return col;
}


// ─────────────────────────────────────────────
// String helpers
// ─────────────────────────────────────────────

function normalize(str) {
  if (!str) return '';

  return str
    .toString()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function safeString(value) {
  return value ? value.toString().trim() : '';
}


// ─────────────────────────────────────────────
// Drive helpers
// ─────────────────────────────────────────────

function getFileNames(folderId) {
  try {
    const result = Drive.Files.list({
      q: `'${folderId}' in parents and trashed = false`,
      includeItemsFromAllDrives: true,
      supportsAllDrives: true,
      fields: 'files(name)',
      pageSize: 200
    });

    return (result.files || []).map(file => {
      const nameWithoutExt = file.name.replace(/\.[^/.]+$/, '');
      return normalize(nameWithoutExt);
    });

  } catch (e) {
    Logger.log(`❌ Could not access folder ${folderId}: ${e.message}`);
    return [];
  }
}

// ─────────────────────────────────────────────
// File matching (strict: last name required)
// ─────────────────────────────────────────────

/**
 * Returns true only if the normalized last name
 * is found somewhere in the file name.
 * First name is used as a secondary signal (logged, not required).
 */
function fileExistsForMember(fileNames, firstName, lastName) {
  if (!lastName) return false;

  const normLast  = normalize(lastName);
  const normFirst = normalize(firstName);

  return fileNames.some(name => {
    const hasLast  = name.includes(normLast);
    const hasFirst = normFirst ? name.includes(normFirst) : false;

    // Strong match: last name found
    if (hasLast) return true;

    // First name only → not enough, skip
    // (caller will add to unmatched list)
    return false;
  });
}


// ─────────────────────────────────────────────
// UI helpers
// ─────────────────────────────────────────────

function showAlert(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function runWithAlert(fn) {
  const ui = SpreadsheetApp.getUi();

  try {
    fn();
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

function buildMissingFieldsMessage({ firstName, lastName, personalEmail }) {
  const missing = [];

  if (!firstName)     missing.push('First Name');
  if (!lastName)      missing.push('Last Name');
  if (!personalEmail) missing.push('Personal Email');

  return `Cannot queue request. The following fields are missing:\n\n• ${missing.join('\n• ')}`;
}


// ─────────────────────────────────────────────
// Email helpers
// ─────────────────────────────────────────────

/**
 * Generates a suggested WIT email address from first and last name.
 * Strips accents, lowercases, removes non-alpha characters.
 *
 * @param {string} firstName
 * @param {string} lastName
 * @returns {string}
 */
function generateWITEmail(firstName, lastName) {
  const clean = (str) =>
    str
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/[^a-z]/g, '');

  return `${clean(firstName)}.${clean(lastName)}@women-in-tech.org`;
}