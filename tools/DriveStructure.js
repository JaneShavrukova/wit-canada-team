function exportDriveStructureToMd() {
  const SHARED_DRIVES = [
    { id: '0ABdfrNcsjynoUk9PVA', name: 'General' },
    { id: '0ADJyJDduRllQUk9PVA', name: 'Leadership' },
  ];
  const OUTPUT_FOLDER_ID = '10Yzu4SKwqWiepB5-LfvGN0eRaHSgektQ';
  const FOLDER_MIME = 'application/vnd.google-apps.folder';

  const lines = ['# WIT Canada — Drive Structure\n'];

  function listItems(folderId) {
    const params = {
      q: `'${folderId}' in parents and trashed = false`,
      includeItemsFromAllDrives: true,
      supportsAllDrives: true,
      corpora: 'allDrives',
      orderBy: 'folder,name',
      fields: 'nextPageToken, files(id, name, mimeType)',
      pageSize: 1000,
    };
    const items = [];
    do {
      const res = Drive.Files.list(params);
      (res.files || []).forEach(f => items.push(f));
      params.pageToken = res.nextPageToken;
    } while (params.pageToken);
    return items;
  }

  function walk(folderId, depth) {
    const indent = '  '.repeat(depth);
    const items = listItems(folderId);

    items.forEach(item => {
      if (item.mimeType.startsWith('image/')) return;
      if (item.mimeType === FOLDER_MIME) {
        lines.push(`${indent}- **${item.name}**`);
        walk(item.id, depth + 1);
      } else {
        lines.push(`${indent}- ${item.name}`);
      }
    });
  }

  SHARED_DRIVES.forEach(drive => {
    lines.push(`## ${drive.name}\n`);
    walk(drive.id, 0);
    lines.push('');
  });

  const content = lines.join('\n');
  const fileName = `WIT_Drive_Structure_${new Date().toISOString().slice(0, 10)}.md`;

  DriveApp.getFolderById(OUTPUT_FOLDER_ID)
    .createFile(fileName, content, MimeType.PLAIN_TEXT);

  SpreadsheetApp.getUi().alert(
    'Done',
    `File created: ${fileName}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
