// ============================================================
// QuotePro — Google Apps Script Backend (Code.gs)
// ============================================================
// วิธีใช้:
// 1. ไปที่ https://script.google.com → สร้าง Project ใหม่
// 2. วางโค้ดนี้ใน Code.gs
// 3. ใส่ SPREADSHEET_ID ของคุณด้านล่าง
// 4. Deploy → New Deployment → Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy Web App URL ไปใส่ใน CONFIG.scriptUrl ใน QuotePro-GSheets.html
// ============================================================

const SPREADSHEET_ID = '1CFWnUvzrjRCavPBvirfOmOek5sxyvn1kzzhQrWaTWtE'; // ← Google Sheet ID
const DRIVE_FOLDER_ID = '1ZKxXkdQI2tBv7wyPgopbZfwD0PHd_Jdy';            // ← Google Drive Folder ID (สำหรับเก็บไฟล์ PO)

// โครงสร้าง Header ของแต่ละ Sheet
const SHEET_HEADERS = {
  Quotations:      ['ID','Title','QuoteDate','ClientName','ContactPerson','Phone','Email','PaymentMethod','Validity','ContractStart','ContractEnd','CustAddress','CustTaxID','CustTel','SaleName','SaleTel','SaleEmail','Status','PaidAmount','TotalAmount','SubTotal','DiscountPct','ApproverEmail','ApproverName','Note','Terms','DealStatus','CancelReason','CreatedBy','CreatedAt','UpdatedAt','PaymentDueDate'],
  QuoteItems:      ['ID','Title','QuoteID','Description','Quantity','Unit','UnitPrice','DiscountPct','LineTotal'],
  Catalog:         ['ID','Title','SKU','Price','Unit','Category','Description'],
  Customers:       ['ID','Title','ContactPerson','TaxID','Phone','Email','Address'],
  Settings:        ['ID','Title','Value'],
  Payments:        ['ID','Title','QuoteID','AmountPaid','PaymentDate','PaymentMethod','Reference','Note','SlipURL'],
  Approvals:       ['ID','Title','QuoteID','ApproveStatus','Comment','ApprovedDate','ApproverEmail'],
  CompanySettings: ['ID','Title','CompanyName','TaxID','Phone','Email','Address','PrimaryColor','FooterText','LogoURL','SignatureURL']
};

// ============================================================
// Fields ที่ต้องเก็บเป็น String เสมอ (เบอร์โทร, TaxID ฯลฯ)
const TEXT_FIELDS = ['Phone','CustTel','SaleTel','TaxID','CustTaxID','Reference','ContactPerson'];

// ============================================================
// UTILITIES
// ============================================================
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function generateId() {
  return Utilities.getUuid();
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = SHEET_HEADERS[name];
    if (headers) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getHeaders(sheet) {
  const last = sheet.getLastColumn();
  if (last === 0) return [];
  return sheet.getRange(1, 1, 1, last).getValues()[0];
}

function sheetToObjects(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      // Convert Date objects to ISO string
      if (row[i] instanceof Date) {
        obj[h] = row[i].toISOString();
      } else if (row[i] === '' || row[i] === null || row[i] === undefined) {
        obj[h] = null;
      } else if (TEXT_FIELDS.includes(h)) {
        // บังคับเป็น String เสมอ เพื่อรักษา 0 นำหน้าเบอร์โทร/TaxID
        obj[h] = String(row[i]);
      } else {
        obj[h] = row[i];
      }
    });
    // Map 'ID' column to lowercase 'id' for frontend compatibility
    if (obj.ID !== undefined) obj.id = String(obj.ID);
    return obj;
  }).filter(obj => obj.ID); // skip empty rows
}

function findRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(1, 1, lastRow, 1).getValues().flat().map(String);
  const idx = ids.indexOf(String(id));
  return idx === -1 ? -1 : idx + 1; // 1-based row number
}

function setTextColumnsFormat(sheet, headers) {
  const lastRow = Math.max(sheet.getLastRow() + 50, 200);
  headers.forEach(function(h, i) {
    if (TEXT_FIELDS.indexOf(h) !== -1) {
      sheet.getRange(2, i + 1, lastRow, 1).setNumberFormat('@');
    }
  });
}

function makeRow(headers, id, data) {
  return headers.map(h => {
    if (h === 'ID') return id;
    const val = data[h];
    if (val === undefined || val === null) return '';
    // บังคับ text fields เป็น String เพื่อกัน Google Sheets ตัด 0 นำหน้าออก
    if (TEXT_FIELDS.includes(h) && val !== '') return String(val);
    return val;
  });
}

// ============================================================
// RESPONSE HELPERS
// ============================================================
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(msg) {
  return jsonResponse({ error: msg });
}

// ============================================================
// GET — รับข้อมูลจาก Sheet
// ============================================================
function doGet(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || 'list';
    const sheetName = params.sheet || '';

    if (action === 'list') {
      if (!sheetName) return errorResponse('Missing sheet parameter');
      const ss = getSpreadsheet();
      const sheet = getOrCreateSheet(ss, sheetName);
      const rows = sheetToObjects(sheet);
      return jsonResponse({ data: rows });
    }

    if (action === 'ping') {
      return jsonResponse({ ok: true, timestamp: new Date().toISOString() });
    }

    return errorResponse('Unknown action: ' + action);

  } catch(err) {
    return errorResponse('doGet error: ' + err.toString());
  }
}

// ============================================================
// POST — เขียนข้อมูลลง Sheet
// ============================================================
function doPost(e) {
  try {
    const raw = e.postData ? e.postData.contents : '{}';
    const payload = JSON.parse(raw);
    const { action, sheet: sheetName, id, data } = payload;

    // uploadFile ไม่ต้องใช้ sheetName — จัดการก่อน
    if (action === 'uploadFile') {
      if (!data || !data.base64Data || !data.fileName) return errorResponse('Missing file data');
      const folderId = (data && data.folderId) || DRIVE_FOLDER_ID;
      const folder = DriveApp.getFolderById(folderId);
      const mimeType = (data && data.mimeType) || 'application/octet-stream';
      const decoded = Utilities.base64Decode(data.base64Data);
      const blob = Utilities.newBlob(decoded, mimeType, data.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const fileId = file.getId();
      return jsonResponse({
        success: true,
        fileId: fileId,
        fileUrl: 'https://drive.google.com/file/d/' + fileId + '/view',
        fileName: file.getName()
      });
    }

    if (!sheetName) return errorResponse('Missing sheet name');

    const ss = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, sheetName);
    const headers = getHeaders(sheet);

    switch (action) {

      // ---- ADD ROW ----
      case 'add': {
        const newId = generateId();
        const row = makeRow(headers, newId, data || {});
        // ตั้ง format @TEXT สำหรับ phone/tax columns ก่อน append
        setTextColumnsFormat(sheet, headers);
        sheet.appendRow(row);
        return jsonResponse({ success: true, id: newId });
      }

      // ---- UPDATE ROW ----
      case 'update': {
        if (!id) return errorResponse('Missing id for update');
        const rowIdx = findRowById(sheet, id);
        if (rowIdx === -1) return errorResponse('Row not found: ' + id);
        // อ่านค่าเดิมก่อน แล้ว merge กับค่าใหม่
        const currentVals = sheet.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
        const currentObj = {};
        headers.forEach((h, i) => { currentObj[h] = currentVals[i]; });
        // Merge: ใช้ค่าใหม่ถ้ามี, ไม่งั้นใช้ค่าเดิม
        const merged = { ...currentObj, ...(data || {}) };
        merged.ID = id; // ป้องกัน ID ถูกเขียนทับ
        const row = makeRow(headers, id, merged);
        // ตั้ง format @TEXT สำหรับ phone/tax columns ก่อน update
        setTextColumnsFormat(sheet, headers);
        sheet.getRange(rowIdx, 1, 1, headers.length).setValues([row]);
        return jsonResponse({ success: true });
      }

      // ---- DELETE ROW ----
      case 'delete': {
        if (!id) return errorResponse('Missing id for delete');
        const rowIdx = findRowById(sheet, id);
        if (rowIdx === -1) return errorResponse('Row not found: ' + id);
        sheet.deleteRow(rowIdx);
        return jsonResponse({ success: true });
      }

      // ---- BULK DELETE (เช่น ลบ QuoteItems ทั้งหมดของ Quote) ----
      case 'deleteWhere': {
        // data = { field: 'QuoteID', value: 'some-uuid' }
        if (!data || !data.field || data.value === undefined) return errorResponse('Missing deleteWhere params');
        const fieldIdx = headers.indexOf(data.field);
        if (fieldIdx === -1) return errorResponse('Field not found: ' + data.field);
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return jsonResponse({ success: true, deleted: 0 });
        const colVals = sheet.getRange(2, fieldIdx + 1, lastRow - 1, 1).getValues().flat();
        // ลบจากล่างขึ้นบน เพื่อไม่ให้ row index เลื่อน
        let deleted = 0;
        for (let i = colVals.length - 1; i >= 0; i--) {
          if (String(colVals[i]) === String(data.value)) {
            sheet.deleteRow(i + 2); // +2 เพราะ header row + 0-based index
            deleted++;
          }
        }
        return jsonResponse({ success: true, deleted });
      }

      // ---- UPLOAD FILE TO GOOGLE DRIVE ----
      case 'uploadFile': {
        if (!data || !data.base64Data || !data.fileName) return errorResponse('Missing file data');
        const folderId = data.folderId || DRIVE_FOLDER_ID;
        const folder = DriveApp.getFolderById(folderId);
        const mimeType = data.mimeType || 'application/octet-stream';
        const decoded = Utilities.base64Decode(data.base64Data);
        const blob = Utilities.newBlob(decoded, mimeType, data.fileName);
        const file = folder.createFile(blob);
        // ตั้งค่าให้ทุกคนที่มี link ดูได้
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileId = file.getId();
        return jsonResponse({
          success: true,
          fileId: fileId,
          fileUrl: 'https://drive.google.com/file/d/' + fileId + '/view',
          fileName: file.getName()
        });
      }

      // ---- SEND EMAIL ----
      case 'sendEmail': {
        if (!data || !data.to || !data.subject) return errorResponse('Missing email params');
        MailApp.sendEmail({
          to: data.to,
          subject: data.subject,
          body: data.body || '',
        });
        return jsonResponse({ success: true });
      }

      default:
        return errorResponse('Unknown action: ' + action);
    }

  } catch(err) {
    return errorResponse('doPost error: ' + err.toString());
  }
}

// ============================================================
// SETUP — สร้าง Sheets ทั้งหมดพร้อม Header
// ============================================================
function setupAllSheets() {
  const ss = getSpreadsheet();
  Object.keys(SHEET_HEADERS).forEach(name => {
    getOrCreateSheet(ss, name);
    Logger.log('Created/Verified sheet: ' + name);
  });
  Logger.log('Setup complete!');
}

// ============================================================
// DAILY REMINDERS — รันทุกเช้าผ่าน Time-based Trigger
// ============================================================

function checkAndSendReminders() {
  var ss = getSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var tomorrow = Utilities.formatDate(new Date(Date.now() + 86400000), tz, 'yyyy-MM-dd');

  var sheet = getOrCreateSheet(ss, 'Quotations');
  var rows = sheetToObjects(sheet);

  rows.forEach(function(q) {
    if (q.Status === 'Closed' || q.Status === 'Cancelled') return;

    // 1. แจ้งเตือนกำหนดวันชำระ (PaymentDueDate)
    if (q.PaymentDueDate) {
      var dueStr = String(q.PaymentDueDate).split('T')[0];
      if (dueStr === today || dueStr === tomorrow) {
        var toEmail = q.SaleEmail || '';
        if (toEmail && toEmail.indexOf('@') !== -1) {
          var label = dueStr === today ? 'วันนี้' : 'พรุ่งนี้';
          MailApp.sendEmail({
            to: toEmail,
            subject: '[แจ้งเตือน] ครบกำหนดชำระเงิน' + label + ' — ' + (q.Title||'') + ' / ' + (q.ClientName||''),
            body: 'ใบเสนอราคา : ' + (q.Title||'') + '\n'
              + 'ลูกค้า      : ' + (q.ClientName||'') + '\n'
              + 'มูลค่ารวม   : ' + Number(q.TotalAmount||0).toLocaleString() + ' บาท\n'
              + 'ยอดค้างชำระ : ' + Number((q.TotalAmount||0)-(q.PaidAmount||0)).toLocaleString() + ' บาท\n'
              + 'กำหนดชำระ  : ' + dueStr + ' (' + label + ')\n\n'
              + 'กรุณาติดตามการชำระเงินจากลูกค้า'
          });
          Logger.log('Sent payment due reminder to ' + toEmail + ' for ' + q.Title);
        }
      }
    }

    // 2. แจ้งเตือนติดตามจัดซื้อ (reminderDate ใน DealStatus JSON)
    if (q.DealStatus) {
      try {
        var ds = JSON.parse(q.DealStatus);
        var reminderDate = ds.reminderDate;
        if (reminderDate) {
          var remStr = String(reminderDate).split('T')[0];
          if (remStr === today || remStr === tomorrow) {
            var toEmail2 = q.SaleEmail || '';
            if (toEmail2 && toEmail2.indexOf('@') !== -1) {
              var label2 = remStr === today ? 'วันนี้' : 'พรุ่งนี้';
              MailApp.sendEmail({
                to: toEmail2,
                subject: '[แจ้งเตือน] ติดตามจัดซื้อ ' + label2 + ' — ' + (q.Title||'') + ' / ' + (q.ClientName||''),
                body: 'มีการตั้งแจ้งเตือนติดตามจัดซื้อสำหรับใบเสนอราคานี้\n\n'
                  + 'ใบเสนอราคา : ' + (q.Title||'') + '\n'
                  + 'ลูกค้า      : ' + (q.ClientName||'') + '\n'
                  + 'สถานะ       : ' + (q.Status||'') + '\n'
                  + 'วันที่ตั้งเตือน: ' + remStr + ' (' + label2 + ')\n\n'
                  + 'กรุณาติดตามสถานะการจัดซื้อ'
              });
              Logger.log('Sent procurement reminder to ' + toEmail2 + ' for ' + q.Title);
            }
          }
        }
      } catch(e) { /* DealStatus parse error - skip */ }
    }
  });

  Logger.log('checkAndSendReminders completed for ' + today);
}

// รันฟังก์ชันนี้ 1 ครั้งจาก Apps Script Editor เพื่อสร้าง Trigger อัตโนมัติ
function setupDailyReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkAndSendReminders') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('checkAndSendReminders')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  Logger.log('Daily reminder trigger created — runs every day at 8am');
}
