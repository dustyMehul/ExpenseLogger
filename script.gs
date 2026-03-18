const SPREADSHEET_ID = 'PASTE_YOUR_SPREADSHEET_ID';
const SHEET_NAME = 'Expenses';
const PARENT_FOLDER_ID = 'PASTE_YOUR_SPREADSHEET_ID';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const amount = data.amount || '';
    const vendor = data.vendor || '';
    const category = data.category || '';
    const expenseDate = data.date || '';
    const notes = data.notes || '';
    const imageBase64 = data.imageBase64 || '';
    const mimeType = data.mimeType || 'image/jpeg';
    const originalFileName = data.fileName || 'receipt.jpg';

    if (!expenseDate) {
      throw new Error('Missing field: date');
    }

    if (!imageBase64) {
      throw new Error('Missing field: imageBase64');
    }

    const monthKey = expenseDate.substring(0, 7); // e.g. 2026-03

    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    const monthFolder = getOrCreateMonthFolder_(parentFolder, monthKey);

    const extension = getExtensionFromMimeType_(mimeType);
    const safeVendor = vendor ? vendor.replace(/[^\w\-]+/g, '_') : 'expense';
    const fileName = `${expenseDate}_${safeVendor}_${Date.now()}.${extension}`;

    const cleanBase64 = imageBase64.replace(/^data:.+;base64,/, '');
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
      monthKey: monthKey
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

function getExtensionFromMimeType_(mimeType) {
  const map = {
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/png': 'png',
    'application/pdf': 'pdf',
    'image/heic': 'heic'
  };
  return map[mimeType] || 'bin';
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