/**
 * Google Apps Script Backend สำหรับระบบกรอกและติดตามแผนงาน PM (PM Plan)
 * เชื่อมต่อกับ Google Spreadsheet: 1TJM3PnV12H__uoCv2VuwtfbceSRfmR3l4lq81xq8vJE
 *
 * วิธีติดตั้ง:
 * 1. เปิดไฟล์ Google Sheet -> เมนู "ส่วนขยาย" (Extensions) -> "Apps Script"
 * 2. ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดไฟล์นี้ลงไป -> กดบันทึก (Ctrl+S)
 * 3. กดปุ่ม "ทำให้ใช้งานได้" (Deploy) -> "การทำให้ใช้งานได้รายการใหม่" (New deployment)
 * 4. เลือกประเภทเป็น "เว็บแอป" (Web app):
 *    - ดำเนินการในฐานะ (Execute as): "ฉัน" (Me)
 *    - ผู้ที่มีสิทธิ์เข้าถึง (Who has access): "ทุกคน" (Anyone)
 * 5. กด "ทำให้ใช้งานได้" (Deploy) -> อนุญาตสิทธิ์ -> คัดลอก Web App URL (ลงท้ายด้วย /exec)
 * 6. นำ URL ไปวางในปุ่ม "ตั้งค่าเชื่อมต่อ API" บนหน้าเว็บ index.html
 */

const SPREADSHEET_ID = '1TJM3PnV12H__uoCv2VuwtfbceSRfmR3l4lq81xq8vJE';

const STATUS_OPTIONS = [
  'รอดำเนินการ',
  'ส่งหนังสือแจ้งแผนงานเรียบร้อย',
  'อยู่ระหว่างดำเนินการ',
  'พบปัญหา / รอดำเนินการแก้ไข',
  'ดำเนินการแล้วเสร็จ',
  'เอกสารจัดทำแล้วเสร็จ'
];

const SHEET_CONFIG = {
  LK: {
    key: 'LK',
    name: 'ลข',
    gid: 475478333,
    activeCols: 15,
    dateCols: [4, 6, 8, 10, 12, 13],
    statusCols: [5, 7, 9, 11, 14]
  },
  BSGW: {
    key: 'BSGW',
    name: 'BS/GW',
    gid: 1755639936,
    activeCols: 25,
    dateCols: [4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 23],
    statusCols: [5, 7, 9, 11, 13, 15, 17, 19, 21, 24]
  },
  BSSC: {
    key: 'BSSC',
    name: 'BSSC',
    gid: 0,
    activeCols: 24,
    dateCols: [4, 6, 8, 10, 12, 14, 16, 18, 20, 22],
    statusCols: [5, 7, 9, 11, 13, 15, 17, 19, 21, 23]
  }
};

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action;

    if (action === 'ping') {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const result = {
        ok: true,
        title: ss.getName(),
        spreadsheetId: SPREADSHEET_ID,
        updatedAt: new Date().toISOString()
      };
      if (params.callback) {
        const safeCb = String(params.callback).replace(/[^a-zA-Z0-9_$.]/g, '');
        return ContentService.createTextOutput(safeCb + '(' + JSON.stringify(result) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return jsonResponse_(result);
    }

    if (action === 'getAll') {
      const result = readAllSheets_();
      if (params.callback) {
        const safeCb = String(params.callback).replace(/[^a-zA-Z0-9_$.]/g, '');
        return ContentService.createTextOutput(safeCb + '(' + JSON.stringify(result) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return jsonResponse_(result);
    }

    // หากเปิดผ่านเบราว์เซอร์หรือมือถือโดยตรง จะโหลดหน้าเว็บให้อัตโนมัติ (หากสร้างไฟล์ HTML ชื่อ index ไว้ในโปรเจกต์)
    try {
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('ระบบกรอกและติดตามแผนงาน PM Plan')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=5.0');
    } catch (_) {
      return jsonResponse_(readAllSheets_());
    }
  } catch (err) {
    return jsonResponse_({ ok: false, error: err && err.message ? err.message : String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const payload = JSON.parse(rawBody);
    const action = payload.action;

    if (action === 'getAll') {
      return jsonResponse_(readAllSheets_());
    }
    if (action === 'updateCell') {
      return jsonResponse_(updateCell_(payload));
    }
    if (action === 'updateRow') {
      return jsonResponse_(updateRow_(payload));
    }
    if (action === 'bulkUpdate') {
      return jsonResponse_(bulkUpdate_(payload));
    }
    if (action === 'addRow') {
      return jsonResponse_(addRow_(payload));
    }
    if (action === 'deleteRow') {
      return jsonResponse_(deleteRow_(payload));
    }

    return jsonResponse_({ ok: false, error: 'ไม่รู้จักคำสั่ง action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err && err.message ? err.message : String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getTargetSheet_(ss, tabKey) {
  const cfg = SHEET_CONFIG[tabKey];
  if (!cfg) {
    throw new Error('ไม่พบการตั้งค่าแท็บ: ' + tabKey);
  }
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === cfg.gid) {
      return { sheet: sheets[i], cfg: cfg };
    }
  }
  const byName = ss.getSheetByName(cfg.name);
  if (byName) {
    return { sheet: byName, cfg: cfg };
  }
  throw new Error('ไม่พบชีต "' + cfg.name + '" (gid=' + cfg.gid + ') ในไฟล์ Google Sheet');
}

function readAllSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const out = {};

  Object.keys(SHEET_CONFIG).forEach(function(key) {
    const target = getTargetSheet_(ss, key);
    const sheet = target.sheet;
    const cfg = target.cfg;
    const lastRow = sheet.getLastRow();

    if (lastRow < 1) {
      out[key] = { key: cfg.key, name: sheet.getName(), gid: String(cfg.gid), headers: [], rows: [] };
      return;
    }

    const displayGrid = sheet.getRange(1, 1, lastRow, cfg.activeCols).getDisplayValues();
    const headers = displayGrid[0];
    const rows = [];

    for (let r = 1; r < displayGrid.length; r++) {
      const rowCells = displayGrid[r];
      const hasAnyValue = rowCells.some(function(v) { return String(v || '').trim() !== ''; });
      if (!hasAnyValue) continue;
      rows.push({
        sheetRow: r + 1,
        cells: rowCells
      });
    }

    out[key] = {
      key: cfg.key,
      name: sheet.getName(),
      gid: String(cfg.gid),
      headers: headers,
      rows: rows
    };
  });

  return {
    ok: true,
    updatedAt: new Date().toISOString(),
    sheets: out
  };
}

function toSheetValue_(val, colIndex, cfg) {
  const str = String(val === null || val === undefined ? '' : val).trim();
  if (str === '') return '';

  if (cfg.dateCols.indexOf(colIndex) !== -1) {
    const m = str.match(/^(\d{1,2})\/(\d{4})$/);
    if (m) {
      const month = parseInt(m[1], 10);
      const year = parseInt(m[2], 10);
      if (month >= 1 && month <= 12) {
        return new Date(year, month - 1, 1);
      }
    }
  }
  return str;
}

function applyRowFormattingAndValidation_(sheet, rowNumber, cfg) {
  cfg.dateCols.forEach(function(colIdx) {
    sheet.getRange(rowNumber, colIdx + 1).setNumberFormat('mm/yyyy');
  });

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();

  cfg.statusCols.forEach(function(colIdx) {
    sheet.getRange(rowNumber, colIdx + 1).setDataValidation(rule);
  });
}

function updateCell_(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const target = getTargetSheet_(ss, payload.tab);
  const sheet = target.sheet;
  const cfg = target.cfg;
  const sheetRow = parseInt(payload.sheetRow, 10);
  const colIndex = parseInt(payload.colIndex, 10);

  if (!sheetRow || sheetRow < 2) throw new Error('หมายเลขแถวไม่ถูกต้อง');
  if (isNaN(colIndex) || colIndex < 0 || colIndex >= cfg.activeCols) throw new Error('คอลัมน์ไม่ถูกต้อง');

  const cell = sheet.getRange(sheetRow, colIndex + 1);
  const parsedVal = toSheetValue_(payload.value, colIndex, cfg);
  cell.setValue(parsedVal);

  if (cfg.dateCols.indexOf(colIndex) !== -1 && parsedVal instanceof Date) {
    cell.setNumberFormat('mm/yyyy');
  }

  SpreadsheetApp.flush();
  return {
    ok: true,
    tab: payload.tab,
    sheetRow: sheetRow,
    colIndex: colIndex,
    displayValue: cell.getDisplayValue(),
    updatedAt: new Date().toISOString()
  };
}

function updateRow_(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const target = getTargetSheet_(ss, payload.tab);
  const sheet = target.sheet;
  const cfg = target.cfg;
  const sheetRow = parseInt(payload.sheetRow, 10);

  if (!sheetRow || sheetRow < 2) throw new Error('หมายเลขแถวไม่ถูกต้อง');
  if (!Array.isArray(payload.cells)) throw new Error('ข้อมูล cells ไม่ถูกต้อง');

  const rowValues = [];
  for (let c = 0; c < cfg.activeCols; c++) {
    rowValues.push(toSheetValue_(payload.cells[c], c, cfg));
  }

  const range = sheet.getRange(sheetRow, 1, 1, cfg.activeCols);
  range.setValues([rowValues]);

  cfg.dateCols.forEach(function(colIdx) {
    if (rowValues[colIdx] instanceof Date) {
      sheet.getRange(sheetRow, colIdx + 1).setNumberFormat('mm/yyyy');
    }
  });

  SpreadsheetApp.flush();
  return {
    ok: true,
    tab: payload.tab,
    sheetRow: sheetRow,
    cells: range.getDisplayValues()[0],
    updatedAt: new Date().toISOString()
  };
}

function bulkUpdate_(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const target = getTargetSheet_(ss, payload.tab);
  const sheet = target.sheet;
  const cfg = target.cfg;
  const updates = payload.updates;

  if (!Array.isArray(updates) || updates.length === 0) {
    throw new Error('ไม่มีรายการอัปเดต');
  }

  let count = 0;
  updates.forEach(function(item) {
    const sheetRow = parseInt(item.sheetRow, 10);
    if (!sheetRow || sheetRow < 2 || !item.changes) return;

    Object.keys(item.changes).forEach(function(colKey) {
      const colIndex = parseInt(colKey, 10);
      if (isNaN(colIndex) || colIndex < 0 || colIndex >= cfg.activeCols) return;
      const cell = sheet.getRange(sheetRow, colIndex + 1);
      const parsedVal = toSheetValue_(item.changes[colKey], colIndex, cfg);
      cell.setValue(parsedVal);
      if (cfg.dateCols.indexOf(colIndex) !== -1 && parsedVal instanceof Date) {
        cell.setNumberFormat('mm/yyyy');
      }
    });
    count++;
  });

  SpreadsheetApp.flush();
  return {
    ok: true,
    tab: payload.tab,
    updatedRows: count,
    updatedAt: new Date().toISOString()
  };
}

function addRow_(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const target = getTargetSheet_(ss, payload.tab);
  const sheet = target.sheet;
  const cfg = target.cfg;

  if (!Array.isArray(payload.cells)) throw new Error('ข้อมูล cells ไม่ถูกต้อง');

  const newRowNumber = Math.max(sheet.getLastRow() + 1, 2);
  const rowValues = [];
  for (let c = 0; c < cfg.activeCols; c++) {
    rowValues.push(toSheetValue_(payload.cells[c], c, cfg));
  }

  // คัดลอกรูปแบบตัวอักษร/เส้นตารางจากแถวก่อนหน้า (ถ้ามี)
  if (newRowNumber > 2) {
    const templateRange = sheet.getRange(newRowNumber - 1, 1, 1, cfg.activeCols);
    const destRange = sheet.getRange(newRowNumber, 1, 1, cfg.activeCols);
    templateRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }

  const destRange = sheet.getRange(newRowNumber, 1, 1, cfg.activeCols);
  destRange.setValues([rowValues]);
  applyRowFormattingAndValidation_(sheet, newRowNumber, cfg);

  SpreadsheetApp.flush();
  return {
    ok: true,
    tab: payload.tab,
    sheetRow: newRowNumber,
    cells: destRange.getDisplayValues()[0],
    updatedAt: new Date().toISOString()
  };
}

function deleteRow_(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const target = getTargetSheet_(ss, payload.tab);
  const sheet = target.sheet;
  const sheetRow = parseInt(payload.sheetRow, 10);

  if (!sheetRow || sheetRow < 2) throw new Error('ไม่สามารถลบแถวหัวตารางได้');

  const currentStation = String(sheet.getRange(sheetRow, 1).getDisplayValue() || '').trim();
  if (payload.expectedStation && currentStation !== String(payload.expectedStation).trim()) {
    throw new Error('ชื่อสถานีในแถวที่ ' + sheetRow + ' ไม่ตรงกับข้อมูลล่าสุด กรุณารีเฟรชหน้าเว็บก่อนลองใหม่');
  }

  sheet.deleteRow(sheetRow);
  SpreadsheetApp.flush();

  return {
    ok: true,
    tab: payload.tab,
    deletedRow: sheetRow,
    updatedAt: new Date().toISOString()
  };
}
