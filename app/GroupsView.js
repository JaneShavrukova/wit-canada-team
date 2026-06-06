// ============================================================
// WIT Canada — Groups View
// ============================================================

const GROUPS_VIEW_SHEET = "Report_Groups";
const EMAIL_ROW = 3;
const SOURCE_SHEETS = [CONFIG.SHEET.MAIN, "WIT_External"];

function buildGroupsView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.SHEET.MAIN);

  // Read group names (row 2) and emails (row 3)
  const allHeaders = mainSheet
    .getRange(CONFIG.SHEET.HEADER_ROW, 1, 1, mainSheet.getMaxColumns())
    .getValues()[0];

  const emailRow = mainSheet
    .getRange(EMAIL_ROW, 1, 1, mainSheet.getMaxColumns())
    .getValues()[0];

  // Build group definitions
  const groupDefs = [];
  allHeaders.forEach((header, i) => {
    const email = emailRow[i] ? emailRow[i].toString().trim() : "";
    if (email.includes("@")) {
      groupDefs.push({ name: header.toString().trim(), email, col: i + 1 });
    }
  });

  if (groupDefs.length === 0) {
    showAlert(
      "Error",
      "No group columns found. Check that row 3 contains group emails.",
    );
    return;
  }

  // Collect members { email: [ { name, role } ] }
  const groupMembers = {};
  groupDefs.forEach((g) => (groupMembers[g.email] = []));

  SOURCE_SHEETS.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const colMap = getColumnIndexMap(sheet);
    const firstCol = colMap[normalizeHeader(CONFIG.HEADERS.FIRST_NAME)];
    const lastCol = colMap[normalizeHeader(CONFIG.HEADERS.LAST_NAME)];
    const roleCol = colMap[normalizeHeader(CONFIG.HEADERS.ROLE)];

    if (!firstCol || !lastCol) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.SHEET.DATA_START_ROW) return;

    const data = sheet
      .getRange(
        CONFIG.SHEET.DATA_START_ROW,
        1,
        lastRow - CONFIG.SHEET.DATA_START_ROW + 1,
        sheet.getMaxColumns(),
      )
      .getValues();

    // Resolve group column positions for this sheet specifically.
    // WIT_Members and WIT_External have different column layouts,
    // so g.col (from mainSheet) cannot be used here.
    const groupColsForSheet = groupDefs.map((g) => ({
      email: g.email,
      col: colMap[normalizeHeader(g.name)] ?? null,
    }));

    data.forEach((row) => {
      const first = safeString(row[firstCol - 1]);
      const last = safeString(row[lastCol - 1]);
      if (!first && !last) return;

      const fullName = `${first} ${last}`.trim();
      const role = roleCol ? safeString(row[roleCol - 1]) : "";

      groupColsForSheet.forEach(({ email, col }) => {
        if (col && row[col - 1] === true) {
          groupMembers[email].push({
            name: fullName,
            role,
            isExternal: sheetName !== CONFIG.SHEET.MAIN,
          });
        }
      });
    });
  });

  // Delete and recreate sheet to avoid stale groupings.
  // Record position first so the tab order is preserved after recreation.
  const existing = ss.getSheetByName(GROUPS_VIEW_SHEET);
  const position = existing ? existing.getIndex() : null;
  if (existing) ss.deleteSheet(existing);
  const out = ss.insertSheet(GROUPS_VIEW_SHEET);
  if (position !== null) ss.moveActiveSheet(position);

  const values = [];
  const formats = [];
  const groupRanges = [];

  const timestamp = new Date().toLocaleString("en-CA", {
    timeZone: "America/Vancouver",
    dateStyle: "medium",
    timeStyle: "short",
  });

  values.push(["Groups Snapshot", "", ""]);
  formats.push("title");
  values.push([`Updated: ${timestamp}`, "", ""]);
  formats.push("timestamp");

  groupDefs.forEach((g) => {
    const members = [...groupMembers[g.email]].sort((a, b) =>
      a.name.localeCompare(b.name),
    );

    // Header row
    values.push([g.name, g.email, ""]);
    formats.push("header");

    const memberStartRow = values.length + 1;

    if (members.length === 0) {
      values.push(["", "— no members —", ""]);
      formats.push("empty");
    } else {
      members.forEach((m) => {
        values.push(["", m.name, m.role]);
        formats.push(m.isExternal ? "member-external" : "member");
      });
    }

    const memberEndRow = values.length;
    groupRanges.push({ start: memberStartRow, end: memberEndRow });

    // Gap row
    values.push(["", "", ""]);
    formats.push("gap");
  });

  // Write all values at once
  if (values.length > 0) {
    out.getRange(1, 1, values.length, 3).setValues(values);
  }

  // Apply formatting
  formats.forEach((type, i) => {
    const fullRow = out.getRange(i + 1, 1, 1, 3);
    const roleCell = out.getRange(i + 1, 3);

    if (type === "title") {
      fullRow
        .setBackground("#1a2fa3")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setFontSize(14);
    } else if (type === "timestamp") {
      fullRow.setBackground("#f8f9fa").setFontColor("#5f6368").setFontSize(10);
    } else if (type === "header") {
      fullRow
        .setBackground("#1a2fa3")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setFontSize(11);
    } else if (type === "member") {
      fullRow
        .setBackground("#ffffff")
        .setFontColor("#3c4043")
        .setFontWeight("normal")
        .setFontSize(11);
      roleCell.setFontColor("#80868b").setFontSize(10);
    } else if (type === "member-external") {
      fullRow
        .setBackground("#ffffff")
        .setFontColor("#80868b")
        .setFontWeight("normal")
        .setFontStyle("italic")
        .setFontSize(11);
      roleCell.setFontColor("#b0b4b8").setFontSize(10);
    } else if (type === "empty") {
      fullRow
        .setBackground("#ffffff")
        .setFontColor("#aaaaaa")
        .setFontStyle("italic")
        .setFontSize(10);
    }
  });

  // Add row groupings for collapse
  groupRanges.forEach(({ start, end }) => {
    if (end >= start) {
      out.getRange(start, 1, end - start + 1, 1).shiftRowGroupDepth(1);
    }
  });

  out.setFrozenRows(2);

  // Auto-resize columns to content width
  out.autoResizeColumns(1, 3);

  ss.setActiveSheet(out);
  showAlert(
    "✅ Done",
    `${GROUPS_VIEW_SHEET} updated — ${groupDefs.length} groups found.`,
  );
}
