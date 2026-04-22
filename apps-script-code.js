// =================================================================
// 零用金記帳系統 · Google Apps Script 後端
// 版本:3.0(雙 Sheet 架構:零用金收支 + 今日便當)
// =================================================================

// Sheet 名稱設定(要和試算表的分頁名稱完全一致)
const SHEET_LEDGER = '零用金收支';  // 主表,若不存在則使用第一個工作表
const SHEET_LUNCH = '今日便當';     // 便當明細表

// ============= 取得工作表 =============
function getLedgerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_LEDGER);
  if (!sheet) sheet = ss.getSheets()[0]; // 沒有就用第一個
  return sheet;
}

function getLunchSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_LUNCH);
  if (!sheet) {
    // 自動建立「今日便當」分頁
    sheet = ss.insertSheet(SHEET_LUNCH);
    sheet.appendRow(['日期', '姓名', '便當口味', '少飯', '金額', '類型', '登記人']);
    sheet.getRange('A1:G1').setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 7, 110);
  }
  return sheet;
}

// ============= GET 請求(讀取,JSONP)=============
function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;
  let result;

  try {
    if (action === 'ping') {
      result = { status: 'ok', msg: '連線成功' };
    } else if (action === 'getAll') {
      result = { status: 'ok', records: readAllLedger() };
    } else if (action === 'getLunchByDate') {
      const date = e.parameter.date;
      result = { status: 'ok', records: readLunchByDate(date) };
    } else if (action === 'getByDate') {
      // 向下相容舊版
      const date = e.parameter.date;
      result = { status: 'ok', records: readLedgerByDate(date) };
    } else {
      result = { status: 'error', msg: '未知的動作: ' + action };
    }
  } catch (err) {
    result = { status: 'error', msg: err.toString() };
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============= POST 請求(寫入)=============
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'append') {
      appendToLedger(getLedgerSheet(), data.record);
      return jsonResponse({ status: 'ok' });
    }

    if (data.action === 'appendBatch') {
      const sheet = getLedgerSheet();
      (data.records || []).forEach(r => appendToLedger(sheet, r));
      return jsonResponse({ status: 'ok', count: (data.records || []).length });
    }

    // 新版:同時寫便當明細 + 零用金主表
    if (data.action === 'submitLunch') {
      // 1. 寫入「今日便當」明細
      const lunchSheet = getLunchSheet();
      (data.details || []).forEach(d => appendToLunch(lunchSheet, d));

      // 2. 寫入「零用金收支」主表
      const ledgerSheet = getLedgerSheet();
      (data.ledger || []).forEach(r => appendToLedger(ledgerSheet, r));

      return jsonResponse({
        status: 'ok',
        detailCount: (data.details || []).length,
        ledgerCount: (data.ledger || []).length
      });
    }

    return jsonResponse({ status: 'error', msg: '未知的動作' });

  } catch (err) {
    return jsonResponse({ status: 'error', msg: err.toString() });
  }
}

// ============= 寫入主表 =============
// 欄位:A=日期 | B=項目 | C=類型 | D=收入 | E=支出 | F=經手人 | G=備註 | H=餘額
function appendToLedger(sheet, r) {
  sheet.appendRow([
    r.date || '',
    r.item || '',
    r.type || '',
    r.income === '' ? '' : Number(r.income),
    r.expense === '' ? '' : Number(r.expense),
    r.handler || '',
    r.note || '',
    '' // 餘額由試算表公式計算
  ]);
}

// ============= 寫入便當明細 =============
// 欄位:A=日期 | B=姓名 | C=便當口味 | D=少飯 | E=金額 | F=類型 | G=登記人
function appendToLunch(sheet, d) {
  sheet.appendRow([
    d.date || '',
    d.person || '',
    d.bento || '',
    d.lessRice || '',
    Number(d.price || 0),
    d.type || '便當',
    d.handler || ''
  ]);
}

// ============= 讀取主表全部 =============
function readAllLedger() {
  const sheet = getLedgerSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  return values.map(row => ({
    date: formatDate(row[0]),
    item: row[1] || '',
    type: row[2] || '',
    income: row[3] === '' ? '' : Number(row[3]),
    expense: row[4] === '' ? '' : Number(row[4]),
    handler: row[5] || ''
  })).filter(r => r.item);
}

// ============= 讀取主表指定日期 =============
function readLedgerByDate(dateStr) {
  return readAllLedger().filter(r => r.date === dateStr);
}

// ============= 讀取便當明細指定日期 =============
function readLunchByDate(dateStr) {
  const sheet = getLunchSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  return values
    .map(row => ({
      date: formatDate(row[0]),
      person: row[1] || '',
      bento: row[2] || '',
      lessRice: row[3] || '',
      price: Number(row[4] || 0),
      type: row[5] || '',
      handler: row[6] || ''
    }))
    .filter(r => r.date === dateStr && r.person);
}

// ============= 日期格式化(強化版) =============
function formatDate(val) {
  if (!val && val !== 0) return '';

  // 如果是 Date 物件
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }

  const s = String(val).trim();
  if (!s) return '';

  // 已經是 YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // YYYY/M/D 或 YYYY-M-D
  const m1 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m1) {
    return m1[1] + '-' + String(m1[2]).padStart(2, '0') + '-' + String(m1[3]).padStart(2, '0');
  }

  // 其他格式(例如 Date.toString() 的結果)嘗試 parse
  try {
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      return y + '-' + m + '-' + dd;
    }
  } catch (e) {}

  return s;
}

// ============= 工具 =============
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
