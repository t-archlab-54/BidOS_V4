function authorizeDrive() {
  const file = DriveApp.getFiles().next();
  Logger.log(file.getName());
}

/********************************************************************
https://drive.google.com/file/d/FILE_ID/view
https://drive.google.com/document/d/FILE_ID/edit
https://docs.google.com/spreadsheets/d/FILE_ID/edit
https://docs.google.com/presentation/d/FILE_ID/edit
https://docs.google.com/forms/d/FILE_ID/edit
https://drive.google.com/open?id=FILE_ID
https://drive.google.com/uc?id=FILE_ID
https://drive.google.com/thumbnail?id=FILE_ID
https://drive.google.com/export?id=FILE_ID
https://drive.google.com/drive/folders/FOLDER_ID
https://.....?resourcekey=...
redirected/tracked/encoded URLs containing drive links
*********************************************************************/

function extractDriveId(url) {
  if (!url) return null;

  let text = String(url).trim();

  // Decode redirected / tracked links if encoded
  try {
    text = decodeURIComponent(text);
  } catch (e) { }

  const patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,            // /file/d/ID , /document/d/ID etc
    /\/folders\/([a-zA-Z0-9_-]+)/,      // folder URLs
    /[?&]id=([a-zA-Z0-9_-]+)/,          // ?id= or &id=
    /\/open\?id=([a-zA-Z0-9_-]+)/,      // open?id=
    /\/uc\?id=([a-zA-Z0-9_-]+)/,        // uc?id=
    /\/thumbnail\?id=([a-zA-Z0-9_-]+)/, // thumbnail?id=
    /\/export\?id=([a-zA-Z0-9_-]+)/     // export?id=
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return match[1];
  }

  return null;
}

function fillFileNamesFromSelection() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const range = sheet.getActiveRange();

  if (!range) return;

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  // Only read the header row instead of the whole sheet.
  const headers = normalizeHeaders_(sheet.getRange(1, 1, 1, lastCol).getValues()[0]);

  const urlCol = getHeaderIndex_(headers, CFG.HEADERS.PROFILE_TAB.RESUME_URL) + 1;
  const nameCol = getHeaderIndex_(headers, CFG.HEADERS.PROFILE_TAB.FILE_NAME) + 1;
  const createdAtCol = getHeaderIndex_(headers, CFG.HEADERS.PROFILE_TAB.CREATED_AT) + 1;

  if (urlCol === 0 || nameCol === 0 || createdAtCol === 0) {
    SpreadsheetApp.getUi().alert(
      "Missing headers: " + [CFG.HEADERS.PROFILE_TAB.RESUME_URL, CFG.HEADERS.PROFILE_TAB.FILE_NAME, CFG.HEADERS.PROFILE_TAB.CREATED_AT].join(" / ")
    );
    return;
  }

  const startRow = range.getRow();
  const numRows = range.getNumRows();

  // Batch read only the selected resume_url cells.
  const urls = sheet.getRange(startRow, urlCol, numRows, 1).getValues();

  // Batch read existing output values so skipped/blank/hidden rows stay unchanged.
  const fileNameRange = sheet.getRange(startRow, nameCol, numRows, 1);
  const createdAtRange = sheet.getRange(startRow, createdAtCol, numRows, 1);
  const fileNameOutput = fileNameRange.getValues();
  const createdAtOutput = createdAtRange.getValues();

  const timezone = Session.getScriptTimeZone();
  let changed = false;

  for (let i = 0; i < numRows; i++) {

    const row = startRow + i;

    // Preserve existing behavior: skip filtered rows entirely.
    if (sheet.isRowHiddenByFilter(row)) continue;

    const url = urls[i][0];

    if (!url) continue;

    const id = extractDriveId(url);

    if (!id) {
      fileNameOutput[i][0] = CFG.VALUES.FILE_LOOKUP.INVALID_URL;
      createdAtOutput[i][0] = "";
      changed = true;
      continue;
    }

    try {

      let item;

      // Try file first, then fallback folder. Same behavior as before.
      try {
        item = DriveApp.getFileById(id);
      } catch (err) {
        item = DriveApp.getFolderById(id);
      }

      fileNameOutput[i][0] = item.getName();
      createdAtOutput[i][0] = Utilities.formatDate(
        item.getDateCreated(),
        timezone,
        CFG.VALUES.FILE_LOOKUP.DATE_FORMAT
      );
      changed = true;

    } catch (e) {

      fileNameOutput[i][0] = CFG.VALUES.FILE_LOOKUP.NO_ACCESS;
      createdAtOutput[i][0] = CFG.VALUES.FILE_LOOKUP.NO_ACCESS;
      changed = true;
    }
  }

  // Two writes total instead of two writes per processed row.
  if (changed) {
    fileNameRange.setValues(fileNameOutput);
    createdAtRange.setValues(createdAtOutput);
  }
}