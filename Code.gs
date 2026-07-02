// Google Apps Script - 代課追蹤資料庫
// 貼到 Apps Script 後先執行 setupSheet()，再部署成 Web App

var SHEET_NAME = '代課追蹤';
var HEADERS = [
  'id', '狀態', '結算月份', '代課類型', '代課教師',
  '日期', '星期', '節次或日代', '班級', '科目', '授課教師',
  '代課節數', '代課日數', '導師日數', '備註', '假別', '摘要',
  '建立時間', '更新時間'
];

function headerIndex(name) {
  return HEADERS.indexOf(name) + 1;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 初始化工作表
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  sheet.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight('bold')
    .setBackground('#daeaf6');
  sheet.setFrozenRows(1);

  var widths = [140, 80, 100, 70, 100, 100, 60, 120, 60, 120, 100, 70, 70, 70, 160, 80, 200, 150, 150];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['已詢問', '已完成', '已取消'], true)
    .build();
  sheet.getRange(2, headerIndex('狀態'), 1000, 1).setDataValidation(statusRule);

  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['節代', '日代'], true)
    .build();
  sheet.getRange(2, headerIndex('代課類型'), 1000, 1).setDataValidation(typeRule);

  var leaveRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['公', '事', '病', '喪', '身心調適假', '其他'], true)
    .build();
  sheet.getRange(2, headerIndex('假別'), 1000, 1).setDataValidation(leaveRule);

  SpreadsheetApp.getUi().alert('工作表「' + SHEET_NAME + '」設定完成！');
}

// 讀取資料
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({ rows: [] });
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return jsonResponse({ rows: [] });
    }

    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var val = row[j];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
        }
        obj[headers[j]] = val;
      }
      rows.push(obj);
    }

    return jsonResponse({ rows: rows });
  } catch (err) {
    return jsonResponse({ rows: [], error: err.toString() });
  }
}

// 新增資料
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('找不到工作表「' + SHEET_NAME + '」，請先執行 setupSheet()');
    }

    // 刪除資料
    if (payload.action === 'delete') {
      var targetId = payload.id;
      var data = sheet.getDataRange().getValues();
      var idCol = HEADERS.indexOf('id');
      for (var r = data.length - 1; r >= 1; r--) {
        if (String(data[r][idCol]) === String(targetId)) {
          sheet.deleteRow(r + 1);
          return jsonResponse({ ok: true, deleted: targetId });
        }
      }
      return jsonResponse({ ok: false, error: '找不到該筆資料' });
    }

    // 新增資料
    var rows = Array.isArray(payload) ? payload : [payload];
    for (var i = 0; i < rows.length; i++) {
      var rowData = rows[i];
      var values = [];
      for (var j = 0; j < HEADERS.length; j++) {
        var key = HEADERS[j];
        values.push(rowData[key] !== undefined ? rowData[key] : '');
      }
      sheet.appendRow(values);
    }

    return jsonResponse({ ok: true, count: rows.length });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}
