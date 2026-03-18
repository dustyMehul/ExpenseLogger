const SPREADSHEET_ID = 'PASTE_YOUR_SPREADSHEET_ID';
const SHEET_NAME = 'Expenses';
const PARENT_FOLDER_ID = 'PASTE_YOUR_PARENT_DRIVE_FOLDER_ID';

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Missing POST data');
    }

    const data = JSON.parse(e.postData.contents);

    const amount = data.amount || '';
    const vendor = data.vendor || '';
    const category = data.category || '';
    const expenseDate = data.date || '';
    const notes = data.notes || '';
    const fileBase64 = data.fileBase64 || data.imageBase64 || '';
    const mimeType = data.mimeType || 'application/octet-stream';
    const originalFileName = data.fileName || 'expense_file';

    if (!expenseDate) {
      throw new Error('Missing field: date');
    }

    if (!fileBase64) {
      throw new Error('Missing field: fileBase64');
    }

    const monthKey = expenseDate.substring(0, 7);
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    const monthFolder = getOrCreateMonthFolder_(parentFolder, monthKey);

    const safeVendor = vendor ? vendor.replace(/[^\w\-]+/g, '_') : 'expense';
    const extension = getExtensionFromMimeType_(mimeType, originalFileName);
    const fileName = `${expenseDate}_${safeVendor}_${Date.now()}.${extension}`;

    const cleanBase64 = fileBase64.replace(/^data:.+;base64,/, '');
    const bytes = Utilities.base64Decode(cleanBase64);
    const blob = Utilities.newBlob(bytes, mimeType, fileName);

    const file = monthFolder.createFile(blob);

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet not found: ${SHEET_NAME}`);
    }

    sheet.appendRow([
      new Date(),
      expenseDate,
      vendor,
      category,
      amount,
      notes,
      monthKey,
      file.getName(),
      file.getUrl()
    ]);

    return jsonOutput_({
      ok: true,
      message: 'Expense logged successfully',
      fileUrl: file.getUrl(),
      fileName: file.getName(),
      monthKey: monthKey,
      mimeType: mimeType
    });

  } catch (err) {
    return jsonOutput_({
      ok: false,
      error: String(err)
    });
  }
}

function getOrCreateMonthFolder_(parentFolder, monthKey) {
  const folders = parentFolder.getFoldersByName(monthKey);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(monthKey);
}

function getExtensionFromMimeType_(mimeType, originalFileName) {
  const map = {
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/png': 'png',
    'image/heic': 'heic',
    'application/pdf': 'pdf'
  };

  if (map[mimeType]) return map[mimeType];

  const parts = (originalFileName || '').split('.');
  return parts.length > 1 ? parts.pop().toLowerCase() : 'bin';
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function testWebhookLocally() {
  const sample = {
    amount: 499,
    vendor: 'Uber',
    category: 'Transport',
    date: '2026-03-18',
    notes: 'Airport ride',
    mimeType: 'image/jpeg',
    fileName: 'receipt.jpg',
    imageBase64: 'SGVsbG8='
  };

  const e = {
    postData: {
      contents: JSON.stringify(sample)
    }
  };

  const result = doPost(e);
  Logger.log(result.getContent());
}