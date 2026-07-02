// Google Apps Script — 代課追蹤資料庫
// 貼到 Google Apps Script 編輯器後，先執行 setupSheet()，再部署成 Web App

const SHEET_NAME = '代課追蹤';
const HEADERS = [
  'id', '狀態', '結算月份', '代課類型', '代課教師',
  '日期', '星期', '節次或日代', '班級', '科目', '授課教師',
  '代課節數', '代課日數', '導師日數', '備註', '假別', '摘要',
  '建立時間', '更新時間'
];

// ── 初始化工作表 ──────────────────────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  // 欄位標題
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
       .setFontWeight('bold')
       .setBackground('#daeaf6');
  sheet.setFrozenRows(1);

  // 欄寬
  const widths = [140, 80, 100, 70, 100, 100, 60, 120, 60, 120, 100, 70, 70, 70, 160, 80, 200, 150, 150];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // 下拉選單：狀態（B欄）
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['已詢問', '已完成', '已取消'], true).build();
  sheet.getRange(2, headerIndex('狀態'), 1000, 1).setDataValidation(statusRule);

  // 下拉選單：代課類型（D欄）
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['節代', '日代'], true).build();
  sheet.getRange(2, headerIndex('代課類型'), 1000, 1).setDataValidation(typeRule);

  // 下拉選單：假別（P欄）
  const leaveRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['公', '事', '病', '喪', '身心調適假', '其他'], true).build();
  sheet.getRange(2, headerIndex('假別'), 1000, 1).setDataValidation(leaveRule);

  SpreadsheetApp.getUi().alert('✅ 工作表「' + SHEET_NAME + '」設定完成！');
}

function headerIndex(name) {
  return HEADERS.indexOf(name) + 1;
}

// ── 讀取資料 (GET) ────────────────────────────────────────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return jsonResponse({ rows: [] });

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return jsonResponse({ rows: [] });

    const headers = data[0];
    const rows = data.slice(1).map(row =>
      Object.fromEntries(headers.map((h, i) => [
        h,
        row[i] instanceof Date
          ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm')
          : row[i]
      ]))
    );

    return jsonResponse({ rows });
  } catch (err) {
    return jsonResponse({ rows: [], error: err.toString() });
  }
}

// ── 新增資料 (POST) ───────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('找不到工作表「' + SHEET_NAME + '」，請先執行 setupSheet()');

    const rows = Array.isArray(payload) ? payload : [payload];
    rows.forEach(row => {
      sheet.appendRow(HEADERS.map(h => row[h] !== undefined ? row[h] : ''));
    });

    return jsonResponse({ ok: true, count: rows.length });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
