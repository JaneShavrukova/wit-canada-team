// ============================================================
// WIT Canada — Contract Signed → Upload Prompt
// ============================================================
// onEdit (dispatched by EditRouter): when Contract Status transitions
// to 'signed', offer to open the signed-contracts Drive folder so the
// lead can upload the signed copy.
//   • "Upload now ↗" → opens the CONFIG.DRIVE.CONTRACTS_FOLDER_ID folder (new tab)
//   • "Not now"      → closes the modal
//
// Fires only on a transition INTO 'signed' (re-applying the same value
// or editing other columns does nothing).
// ============================================================

function processContractSignedOnEdit(e, ctx) {
  const { sheet, sheetName, row, col, colMap } = ctx || buildEditContext(e);
  if (sheetName !== CONFIG.SHEET.MAIN) return;

  // ── Guard: only the Contract Status column ───────────────
  const contractCol = colMap[normalizeHeader(CONFIG.HEADERS.CONTRACT_STATUS)];
  if (!contractCol) return;
  if (col !== contractCol) return;
  if (row < CONFIG.SHEET.DATA_START_ROW) return;

  // ── Guard: only on a transition INTO 'signed' ────────────
  const oldStatus = safeString(e.oldValue).toLowerCase();
  const newStatus = safeString(e.range.getValue()).toLowerCase();
  if (newStatus !== CONFIG.CONTRACT.SIGNED) return;
  if (oldStatus === newStatus) return;

  const get       = getRowValues(sheet, row, colMap);
  const firstName = safeString(get(CONFIG.HEADERS.FIRST_NAME));
  const lastName  = safeString(get(CONFIG.HEADERS.LAST_NAME));
  const fullName  = `${firstName} ${lastName}`.trim() || 'this member';

  _showLaunchCard({
    emoji:       '📄',
    title:       'Upload signed contract',
    description: `${fullName}'s contract is marked signed. Upload the signed copy to the contracts folder now?`,
    url:         `https://drive.google.com/drive/folders/${CONFIG.DRIVE.CONTRACTS_FOLDER_ID}`,
    buttonLabel: 'Upload now ↗',
    closeLabel:  'Not now',
    height:      250,
  });
}
