// =================================================================
// 零用金記帳系統 · Google Apps Script 後端
// 版本:2.0(支援讀取、寫入、批次寫入、JSONP)
// =================================================================

// 試算表的第一個工作表就是記帳表(不需特別設定)
function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

// ============= GET 請求(讀取用,JSONP 模式)=============
function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;
  let result;

  try {
    if (action === 'ping') {
      result = { status: 'ok', msg: '連線成功' };
    } else if (action === 'getAll') {
      result = { status: 'ok', records: readAllRecords() };
    } else if (action === 'getByDate') {
      const date = e.parameter.date;
      result = { status: 'ok', records: readRecordsByDate(date) };
    } else {
      result = { status: 'error', msg: '未知的動作: ' + action };
    }
  } catch (err) {
    result = { status: 'error', msg: err.toString() };
  }

  // JSONP 回應
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============= POST 請求(寫入用)=============
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    if (data.action === 'append') {
      appendOne(sheet, data.record);
      return jsonResponse({ status: 'ok' });
    }

    if (data.action === 'appendBatch') {
      (data.records || []).forEach(r => appendOne(sheet, r));
      return jsonResponse({ status: 'ok', count: (data.records || []).length });
    }

    return jsonResponse({ status: 'error', msg: '未知的動作' });

  } catch (err) {
    return jsonResponse({ status: 'error', msg: err.toString() });
  }
}

// ============= 寫入一筆記錄 =============
// 欄位:A=日期 | B=項目 | C=類型 | D=收入金額 | E=支出金額 | F=經手人 | G=備註 | H=總餘額
function appendOne(sheet, r) {
  sheet.appendRow([
    r.date || '',
    r.item || '',
    r.type || '',
    r.income === '' ? '' : Number(r.income),
    r.expense === '' ? '' : Number(r.expense),
    r.handler || '',
    r.note || '',
    '' // H欄總餘額由試算表公式計算,留空
  ]);
}

// ============= 讀取所有記錄 =============
function readAllRecords() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // 讀 A 到 F 欄,共 6 欄(總餘額 H 欄不需要傳回)
  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  return values.map(row => ({
    date: formatDate(row[0]),
    item: row[1] || '',
    type: row[2] || '',
    income: row[3] === '' ? '' : Number(row[3]),
    expense: row[4] === '' ? '' : Number(row[4]),
    handler: row[5] || ''
  })).filter(r => r.item); // 過濾空行
}

// ============= 讀取指定日期的記錄 =============
function readRecordsByDate(dateStr) {
  const all = readAllRecords();
  return all.filter(r => r.date === dateStr);
}

// ============= 日期格式化 =============
function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  // 如果是文字格式,可能是 2026/4/22 或 2026-04-22
  const s = String(val).trim();
  const m1 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m1) {
    return `${m1[1]}-${String(m1[2]).padStart(2,'0')}-${String(m1[3]).padStart(2,'0')}`;
  }
  return s;
}

// ============= 工具 =============
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
