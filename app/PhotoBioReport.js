// ============================================================
// WIT Canada — Photos & Bios Report
// ============================================================
// Region order and the PropertiesService key live in CONFIG
// (CONFIG.REGION_ORDER, CONFIG.PROPERTIES.SYNC_REPORT).


// ─────────────────────────────────────────────
// Save report to PropertiesService
// ─────────────────────────────────────────────

function saveSyncReport(byRegion) {
  const timestamp = reportTimestamp();

  PropertiesService
    .getScriptProperties()
    .setProperty(CONFIG.PROPERTIES.SYNC_REPORT, JSON.stringify({ timestamp, byRegion }));
}


// ─────────────────────────────────────────────
// Load report from PropertiesService
// ─────────────────────────────────────────────

function loadSyncReport() {
  const raw = PropertiesService
    .getScriptProperties()
    .getProperty(CONFIG.PROPERTIES.SYNC_REPORT);

  if (!raw) return null;
  return JSON.parse(raw);
}


// ─────────────────────────────────────────────
// Sidebar — quick view without re-running sync
// ─────────────────────────────────────────────

function showSyncReportSidebar() {
  const data = loadSyncReport();

  if (!data) {
    showAlert('No report', 'Run "Update photos & bios data" first.');
    return;
  }

  const { timestamp, byRegion } = data;

  const sectionsHtml = CONFIG.REGION_ORDER
    .filter(region => byRegion[region]?.length > 0)
    .map(region => {
      const members = byRegion[region].filter(m => !m.hasPhoto || !m.hasBio);
      if (members.length === 0) return '';

      const rows = members.map(m => {
        const photo = m.hasPhoto ? '✅' : '❌';
        const bio   = m.hasBio   ? '✅' : '❌';
        return `
          <tr>
            <td>${m.name}</td>
            <td style="color:#1a2fa3">${m.witEmail}</td>
            <td style="text-align:center">${photo}</td>
            <td style="text-align:center">${bio}</td>
          </tr>`;
      }).join('');

      return `
        <div class="region">
          <div class="region-label">${region}</div>
          <table>
            <thead>
              <tr>
                <th>Member</th>
                <th>WIT Email</th>
                <th>Photo</th>
                <th>Bio</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </div>`;
    }).join('');

  const allGood = !sectionsHtml.trim();

  const html = `
    <style>
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body {
        font-family: Google Sans, sans-serif;
        font-size: 12px;
        padding: 16px;
        background: #fff;
        color: #3c4043;
      }
      .meta {
        font-size: 11px;
        color: #80868b;
        margin-bottom: 14px;
        padding-bottom: 10px;
        border-bottom: 1px solid #f1f3f4;
      }
      .all-good {
        color: #34a853;
        font-weight: 600;
        font-size: 13px;
        margin-top: 24px;
        text-align: center;
      }
      .region { margin-bottom: 20px; }
      .region-label {
        font-weight: 700;
        font-size: 10px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        color: #5f6368;
        margin-bottom: 6px;
        padding-bottom: 4px;
        border-bottom: 1px solid #e8eaed;
      }
      table { width: 100%; border-collapse: collapse; }
      th {
        text-align: left;
        font-size: 10px;
        color: #80868b;
        font-weight: 600;
        padding: 3px 4px;
        text-transform: uppercase;
      }
      th:not(:first-child) { text-align: center; }
      td {
        padding: 5px 4px;
        border-bottom: 1px solid #f8f9fa;
        font-size: 12px;
      }
      tr:last-child td { border-bottom: none; }
    </style>

    <div class="meta">
      Last synced: <strong>${timestamp}</strong>
    </div>

    ${allGood
      ? '<div class="all-good">✅ All members have photos & bios</div>'
      : sectionsHtml
    }
  `;

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(html).setTitle('📋 Photos & Bios Status')
  );
}


// ─────────────────────────────────────────────
// Report sheet — full view with formatting
// ─────────────────────────────────────────────

function buildPhotoBioReportSheet() {
  const data   = loadSyncReport();
  const report = getOrCreateReportSheet(CONFIG.SHEET.REPORTS.PHOTO_BIO);

  if (!data) {
    report.getRange(1, 1).setValue('No sync data yet. Run "Update photos & bios data" first.');
    return;
  }

  const { timestamp, byRegion } = data;
  const COL_COUNT = 4;

  // Title + timestamp banner (rows 1–2); content begins at row 3.
  let currentRow = writeReportTitleBlock(report, 'Photos & Bios Status Report', COL_COUNT, timestamp);

  CONFIG.REGION_ORDER.forEach(region => {
    const members = (byRegion[region] || []).filter(m => !m.hasPhoto || !m.hasBio);
    if (members.length === 0) return;

    // Region header
    report.getRange(currentRow, 1, 1, COL_COUNT).merge()
      .setValue(region.toUpperCase())
      .setFontSize(10)
      .setFontWeight('bold')
      .setBackground(THEME.report.SECTION_BG)
      .setFontColor(THEME.report.SECTION_FG);
    currentRow++;

    // Column headers
    report.getRange(currentRow, 1, 1, COL_COUNT).setBackground(THEME.surfaceMuted).setFontWeight('bold');
    report.getRange(currentRow, 1).setValue('Member');
    report.getRange(currentRow, 2).setValue('WIT Email');
    report.getRange(currentRow, 3).setValue('Photo');
    report.getRange(currentRow, 4).setValue('Bio');
    currentRow++;

    // Member rows
    members.forEach(m => {
      report.getRange(currentRow, 1).setValue(m.name);
      report.getRange(currentRow, 2).setValue(m.witEmail);
      report.getRange(currentRow, 3).setValue(m.hasPhoto ? '✅' : '❌');
      report.getRange(currentRow, 4).setValue(m.hasBio   ? '✅' : '❌');
      currentRow++;
    });

    currentRow++; // gap between regions
  });

  // Column widths
  setReportColumnWidths(report, [200, 220, 80, 80]);

  report.setFrozenRows(2);
  ss.setActiveSheet(report);
}