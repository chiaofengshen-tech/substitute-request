// Google Apps Script - 代課追蹤資料庫
// 貼到 Apps Script 後：
//   1. 執行 setupSheet() — 建立代課追蹤工作表
//   2. 執行 importTeachers() — 匯入代課教師資料庫（只需執行一次）
//   3. 部署成 Web App

var SHEET_NAME = '代課追蹤';
var TEACHER_SHEET_NAME = '代課教師';
var HEADERS = [
  'id', '狀態', '結算月份', '代課類型', '代課教師',
  '日期', '星期', '節次或日代', '班級', '科目', '授課教師',
  '代課節數', '代課日數', '導師日數', '備註', '假別', '摘要',
  '建立時間', '更新時間'
];
var TEACHER_HEADERS = ['代課教師', '顯示名稱', '薪俸', '學歷', '教師證', '備註', '資料狀態'];

function headerIndex(name) {
  return HEADERS.indexOf(name) + 1;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 初始化代課追蹤工作表
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
    .requireValueInList(['已詢問', '已確認', '已完成', '已取消'], true)
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

// 匯入代課教師資料庫（執行一次即可，之後直接在 Sheets 新增列）
function importTeachers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEACHER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TEACHER_SHEET_NAME);
  } else {
    sheet.clearContents();
  }

  sheet.getRange(1, 1, 1, TEACHER_HEADERS.length)
    .setValues([TEACHER_HEADERS])
    .setFontWeight('bold')
    .setBackground('#e8f5e9');
  sheet.setFrozenRows(1);

  var widths2 = [100, 200, 60, 300, 100, 200, 120];
  for (var i = 0; i < widths2.length; i++) {
    sheet.setColumnWidth(i + 1, widths2[i]);
  }

  var teachers = TEACHER_DATA;
  var values = [];
  for (var j = 0; j < teachers.length; j++) {
    var t = teachers[j];
    var row = [];
    for (var k = 0; k < TEACHER_HEADERS.length; k++) {
      row.push(t[TEACHER_HEADERS[k]] || '');
    }
    values.push(row);
  }
  if (values.length > 0) {
    sheet.getRange(2, 1, values.length, TEACHER_HEADERS.length).setValues(values);
  }

  SpreadsheetApp.getUi().alert('已匯入 ' + teachers.length + ' 筆代課教師資料！');
}

// 代課教師資料已匯入 Google Sheets，此陣列清空以避免個資存在公開 repo
// 如需重新匯入，請自行填入資料後執行 importTeachers()
var TEACHER_DATA = [
  // 請在此填入教師資料，格式：{"代課教師":"姓名","顯示名稱":"顯示名稱","薪俸":"金額","學歷":"學歷","教師證":"有/無","備註":"","資料狀態":"正常"},
];

// 讀取資料
function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action ? e.parameter.action : 'get';
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 讀取教師清單
    if (action === 'teachers') {
      var tsheet = ss.getSheetByName(TEACHER_SHEET_NAME);
      if (!tsheet) return jsonResponse({ teachers: [] });
      var tdata = tsheet.getDataRange().getValues();
      if (tdata.length <= 1) return jsonResponse({ teachers: [] });
      var theaders = tdata[0];
      var teachers = [];
      for (var ti = 1; ti < tdata.length; ti++) {
        var trow = tdata[ti];
        var tobj = {};
        for (var tj = 0; tj < theaders.length; tj++) {
          tobj[theaders[tj]] = trow[tj];
        }
        teachers.push(tobj);
      }
      return jsonResponse({ teachers: teachers });
    }

    // 讀取代課追蹤清單
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return jsonResponse({ rows: [] });

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return jsonResponse({ rows: [] });

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

// 新增 / 刪除代課資料
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('找不到工作表「' + SHEET_NAME + '」，請先執行 setupSheet()');
    }

    // 封存月份資料
    if (payload.action === 'archive') {
      var archiveMonth = payload.month;
      if (!archiveMonth) return jsonResponse({ ok: false, error: '未指定月份' });

      var data = sheet.getDataRange().getValues();
      var monthCol = HEADERS.indexOf('結算月份');
      var rowsToArchive = [];
      var rowIndexToDelete = [];

      for (var ai = 1; ai < data.length; ai++) {
        if (String(data[ai][monthCol]) === String(archiveMonth)) {
          rowsToArchive.push(data[ai]);
          rowIndexToDelete.push(ai + 1); // 1-indexed
        }
      }

      if (rowsToArchive.length === 0) {
        return jsonResponse({ ok: false, error: '找不到「' + archiveMonth + '」的資料' });
      }

      // 建立或取得封存分頁
      var archiveSheet = ss.getSheetByName(archiveMonth);
      if (!archiveSheet) {
        archiveSheet = ss.insertSheet(archiveMonth);
        archiveSheet.getRange(1, 1, 1, HEADERS.length)
          .setValues([HEADERS])
          .setFontWeight('bold')
          .setBackground('#ffe0b2');
        archiveSheet.setFrozenRows(1);
      }

      // 寫入封存分頁
      var lastRow = archiveSheet.getLastRow();
      archiveSheet.getRange(lastRow + 1, 1, rowsToArchive.length, HEADERS.length)
        .setValues(rowsToArchive);

      // 從主表刪除（從後往前刪）
      for (var di = rowIndexToDelete.length - 1; di >= 0; di--) {
        sheet.deleteRow(rowIndexToDelete[di]);
      }

      return jsonResponse({ ok: true, archived: rowsToArchive.length, month: archiveMonth });
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
