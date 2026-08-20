// Google Apps Script - 代課追蹤資料庫
// 貼到 Apps Script 後：
//   1. 執行 setupSheet() — 建立代課追蹤工作表
//   2. 執行 importTeachers() — 匯入代課教師資料庫（只需執行一次）
//   3. 部署成 Web App

var SHEET_NAME = '代課追蹤';
var TEACHER_SHEET_NAME = '代課教師';
var AVAIL_SHEET_NAME = '可代時段';
var AVAIL_HEADERS = ['代課教師', '星期一', '星期二', '星期三', '星期四', '星期五', '可日代', '備註'];
var WEEKDAY_NAMES = ['星期一', '星期二', '星期三', '星期四', '星期五'];
var HEADERS = [
  'id', '狀態', '結算月份', '代課類型', '代課教師',
  '日期', '星期', '節次或日代', '班級', '科目', '授課教師',
  '代課節數', '代課日數', '導師日數', '備註', '假別', '摘要',
  '建立時間', '更新時間'
];
var TEACHER_HEADERS = ['代課教師', '顯示名稱', '薪俸', '學歷', '教師證', '備註', '資料狀態', '特殊可代科目'];
var SPECIAL_SUBJECTS = ['英語', '舞蹈', '電腦', '體育'];

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

// 將 115 學年度一至六年級導師補入「代課教師」母名單。
// 可重複執行：已存在的姓名不會重複新增，也不會覆蓋既有資料。
function addHomeroomTeachers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEACHER_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('找不到「' + TEACHER_SHEET_NAME + '」工作表。');
    return;
  }

  var homeroomTeachers = [
    // 一年級（5 班）
    '王麗君', '李孟潔', '游瑟玫', '陳香君', '周玟慧',
    // 二年級（6 班）
    '陳富美', '許靖宜', '陳宣伶', '吳翠屏', '吳佩儒', '陳芬婷',
    // 三年級（6 班）
    '江筱帆', '吳曜崴', '葉依婷', '陳貞宜', '顧介鈞', '顏秀芳',
    // 四年級（6 班）
    '周素萍', '林湘穎', '林美智', '洪靖雯', '林綵紋', '詹騏璟',
    // 五年級（6 班）
    '蔡旻均', '彭傳家', '柯明月', '王翠津', '黃暐茹', '江敏滋',
    // 六年級（7 班）
    '郭秀雯', '吳宜儒', '林怡如', '王書漢', '余璨同', '林嘉威', '周雷世舫'
  ];

  var data = sheet.getDataRange().getValues();
  if (data.length === 0) {
    SpreadsheetApp.getUi().alert('「' + TEACHER_SHEET_NAME + '」工作表沒有標題列。');
    return;
  }

  var headers = data[0].map(String);
  var nameIdx = headers.indexOf('代課教師');
  var displayIdx = headers.indexOf('顯示名稱');
  var noteIdx = headers.indexOf('備註');
  var statusIdx = headers.indexOf('資料狀態');
  if (nameIdx < 0) {
    SpreadsheetApp.getUi().alert('「' + TEACHER_SHEET_NAME + '」工作表缺少「代課教師」欄位。');
    return;
  }

  function normalizeName(value) {
    return String(value || '').replace(/老師$/, '').trim();
  }

  var existing = {};
  for (var r = 1; r < data.length; r++) {
    var existingName = normalizeName(data[r][nameIdx]);
    if (existingName) existing[existingName] = true;
  }

  var rowsToAdd = [];
  for (var i = 0; i < homeroomTeachers.length; i++) {
    var name = homeroomTeachers[i];
    if (existing[normalizeName(name)]) continue;

    var row = [];
    for (var c = 0; c < headers.length; c++) row.push('');
    row[nameIdx] = name;
    if (displayIdx >= 0) row[displayIdx] = name;
    if (noteIdx >= 0) row[noteIdx] = '115學年度導師';
    if (statusIdx >= 0) row[statusIdx] = '正常';
    rowsToAdd.push(row);
    existing[normalizeName(name)] = true;
  }

  if (rowsToAdd.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, headers.length).setValues(rowsToAdd);
  }

  SpreadsheetApp.getUi().alert(
    '導師名單處理完成：新增 ' + rowsToAdd.length + '位，已存在 ' +
    (homeroomTeachers.length - rowsToAdd.length) + '位。'
  );
}

// 建立「可代時段」工作表（執行一次，之後直接在 Sheets 維護）
function setupAvailability() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AVAIL_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AVAIL_SHEET_NAME);
  } else {
    sheet.clearContents();
  }

  sheet.getRange(1, 1, 1, AVAIL_HEADERS.length)
    .setValues([AVAIL_HEADERS])
    .setFontWeight('bold')
    .setBackground('#e8eaf6');
  sheet.setFrozenRows(1);

  var widths3 = [100, 120, 120, 120, 120, 120, 60, 250];
  for (var i = 0; i < widths3.length; i++) {
    sheet.setColumnWidth(i + 1, widths3[i]);
  }

  // 初始資料（從 substitute-teacher-availability/代課老師可代時段_整理待確認.md 整理）
  var initialData = [
    ['淑珠老師', '', '1,2,3,4,5,6,7', '', '', '1,2,3,4,5,6,7', '是', '每週二、五整天'],
    ['吳洸堯',   '1,2,3,4', '1,2,3,4', '1,2,3,4', '1,2,3,4', '1,2,3,4', '否', '平日上午'],
    ['黃詺媞',   '', '1,2,3,4,5,6,7', '', '', '', '是', '週二'],
    ['王耕玄',   '1,2,3', '1,2,3', '1,2,3', '1,2,3', '1,2,3', '否', '週一至週五第1-3節'],
    ['李佳真',   '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '是', '皆可'],
    ['黃湘美',   '1,2,3,4', '1,2,3,4,5,6', '1,2,3,4', '1,2,3,4,5,6', '1,2,3,4,5,6', '否', '週一三12:00前；週二四五15:30前'],
    ['林佩欣',   '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '是', '皆可'],
    ['林貝蒂',   '5,6,7', '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '1,2,3,4', '1,2,3,4,5,6,7', '否', '週一12:00前不可；週四12:00後不可'],
    ['曾唯皓',   '1,2,3,4,5,6', '1,2,3,4,5,6', '1,2,3,4,5,6', '1,2,3,4,5,6', '1,2,3,4,5,6', '否', '週一至週五至第6節'],
    ['譚定群',   '1,2,3,4', '', '1,2,3,4', '1,2,3,4', '', '否', '週一、三、四上午至12:40'],
    ['吳麗娟',   '', '1,2,3,4,5,6,7', '', '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '是', '週二、四、五'],
    ['吳慧鈴',   '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '', '', '1,2,3,4,5,6,7', '是', '週一五優先；週二備用'],
    ['王怡雯',   '1,2,3,4,5,6,7', '1,2,3,4,5,6,7', '1,2,3,4', '1,2,3,4', '1,2,3,4', '是', '週一二整天；週三至五第1-4節'],
    ['許芳綺',   '', '', '', '', '', '否', '資料待確認'],
    ['楊惠雯',   '', '', '', '', '', '否', '資料待確認'],
    ['曾琬翎',   '', '', '', '', '', '否', '資料待確認'],
    ['李佩郁',   '', '', '', '', '', '否', '資料待確認']
  ];

  sheet.getRange(2, 1, initialData.length, AVAIL_HEADERS.length).setValues(initialData);
  SpreadsheetApp.getUi().alert('「可代時段」工作表設定完成！共 ' + initialData.length + ' 位老師。');
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

    // 篩選可代老師
    if (action === 'availability') {
      var weekday  = (e.parameter.weekday  || '').trim();   // '一','二',...
      var period   = (e.parameter.period   || '').trim();   // '1'~'7'
      var atype    = (e.parameter.type     || '節代').trim(); // '節代' or '日代'
      var adate    = (e.parameter.date     || '').trim();   // 'yyyy/MM/dd'
      var asubject = (e.parameter.subject  || '').trim();   // 科目

      var avSheet = ss.getSheetByName(AVAIL_SHEET_NAME);
      if (!avSheet) return jsonResponse({ teachers: [], error: '請先執行 setupAvailability() 建立可代時段工作表' });

      var avData = avSheet.getDataRange().getValues();
      var avH = avData[0].map(String);

      var nameIdx   = avH.indexOf('代課教師');
      var noteIdx   = avH.indexOf('備註');
      var wdIdx     = weekday ? avH.indexOf('星期' + weekday) : -1;

      // 「代課教師」是完整候選母名單；「可代時段」只記錄已知限制。
      var teacherSheet = ss.getSheetByName(TEACHER_SHEET_NAME);
      if (!teacherSheet) return jsonResponse({ teachers: [], error: '找不到「代課教師」工作表' });
      var teacherData = teacherSheet.getDataRange().getValues();
      if (teacherData.length <= 1) return jsonResponse({ teachers: [] });
      var teacherH = teacherData[0].map(String);
      var teacherNameIdx = teacherH.indexOf('代課教師');
      var teacherSpecialIdx = teacherH.indexOf('特殊可代科目');
      if (teacherNameIdx < 0) return jsonResponse({ teachers: [], error: '「代課教師」工作表缺少「代課教師」欄位' });

      function normalizeTeacherName(name) {
        return String(name || '').replace(/老師$/, '').trim();
      }

      var availabilityByName = {};
      for (var ai = 1; ai < avData.length; ai++) {
        var avName = nameIdx >= 0 ? String(avData[ai][nameIdx] || '').trim() : '';
        if (!avName) continue;
        availabilityByName[normalizeTeacherName(avName)] = avData[ai];
      }

      // x/× 是明確不可；空白或未建資料代表尚未回覆，仍可先詢問。
      var candidates = [];
      var candidateNamesSeen = {};
      for (var ci = 1; ci < teacherData.length; ci++) {
        var cname = String(teacherData[ci][teacherNameIdx] || '').trim();
        if (!cname) continue;
        var normalizedName = normalizeTeacherName(cname);
        if (!normalizedName || candidateNamesSeen[normalizedName]) continue;
        candidateNamesSeen[normalizedName] = true;

        var crow = availabilityByName[normalizedName] || null;
        var cnote = crow && noteIdx >= 0 ? String(crow[noteIdx] || '').trim() : '';
        var wdPeriods = crow && wdIdx >= 0 ? String(crow[wdIdx] || '').trim() : '';
        var isExplicitUnavailable = /^(x|×|✕|不可|不能)$/i.test(wdPeriods);
        if (isExplicitUnavailable) continue;

        var eligible = false;
        var availabilityKnown = wdPeriods !== '';
        var availablePeriods = wdPeriods.match(/[1-7]/g) || [];
        if (!availabilityKnown) {
          eligible = true;
          var unknownNote = '該日可代時段尚未回覆，請先詢問';
          cnote = cnote ? cnote + '；' + unknownNote : unknownNote;
        } else if (atype === '日代') {
          // 已回覆數字節次者，日代必須完整包含 1~7 節。
          eligible = ['1','2','3','4','5','6','7'].every(function(p) {
            return availablePeriods.indexOf(p) >= 0;
          });
        } else {
          eligible = period !== '' && availablePeriods.indexOf(period) >= 0;
        }

        if (eligible) candidates.push({
          name: cname,
          note: cnote,
          availabilityKnown: availabilityKnown
        });
      }

      // 先顯示已明確回覆可代的老師，再顯示時段尚未確認者。
      candidates.sort(function(a, b) {
        return (a.availabilityKnown === b.availabilityKnown) ? 0 : (a.availabilityKnown ? -1 : 1);
      });

      // 特殊科目篩選：支援關鍵字比對（英語彈性→英語、雙語電腦→電腦）
      var specialKeywords = [
        { subject: '英語', aliases: ['英語', '英文'] },
        { subject: '舞蹈', aliases: ['舞蹈'] },
        { subject: '電腦', aliases: ['電腦', '資訊'] },
        { subject: '體育', aliases: ['體育'] }
      ];
      var detectedSpecial = '';
      for (var ki = 0; ki < specialKeywords.length; ki++) {
        var keywordGroup = specialKeywords[ki];
        var matchedAlias = keywordGroup.aliases.some(function(alias) {
          return asubject.indexOf(alias) >= 0;
        });
        if (matchedAlias) { detectedSpecial = keywordGroup.subject; break; }
      }
      var isSpecial = detectedSpecial !== '';
      if (isSpecial) {
        var eligibleSpecial = {};
        if (teacherSpecialIdx >= 0) {
            for (var si = 1; si < teacherData.length; si++) {
              var sName    = String(teacherData[si][teacherNameIdx] || '').trim();
              var sSubject = String(teacherData[si][teacherSpecialIdx] || '').trim();
              var subjectMatches = false;
              if (detectedSpecial === '英語' && (sSubject === '英語' || sSubject === '英文')) subjectMatches = true;
              else if (detectedSpecial === '電腦' && (sSubject === '電腦' || sSubject === '資訊')) subjectMatches = true;
              else if (sSubject === detectedSpecial) subjectMatches = true;
              if (subjectMatches) eligibleSpecial[normalizeTeacherName(sName)] = true;
          }
        }
        candidates = candidates.filter(function(c) {
          return eligibleSpecial[normalizeTeacherName(c.name)];
        });
      }

      // 比對代課追蹤，找衝突
      var conflicts = {};
      var trackSheet2 = ss.getSheetByName(SHEET_NAME);
      if (trackSheet2 && adate) {
        var td = trackSheet2.getDataRange().getValues();
        var th = td[0].map(String);
        var tdDateIdx    = th.indexOf('日期');
        var tdStatusIdx  = th.indexOf('狀態');
        var tdTeacherIdx = th.indexOf('代課教師');
        var tdPeriodIdx  = th.indexOf('節次或日代');
        var tdTypeIdx    = th.indexOf('代課類型');

        for (var ti2 = 1; ti2 < td.length; ti2++) {
          var tr2 = td[ti2];
          var tStatus2 = String(tr2[tdStatusIdx] || '').trim();
          if (tStatus2 === '已取消') continue;

          var tRawDate = tr2[tdDateIdx];
          var tDateStr = '';
          if (tRawDate instanceof Date) {
            tDateStr = Utilities.formatDate(tRawDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
          } else {
            tDateStr = String(tRawDate || '').split(' ')[0].trim();
          }
          if (tDateStr !== adate) continue;

          var tTeacher2 = String(tr2[tdTeacherIdx] || '').replace(/老師$/, '').trim();
          var tType2    = String(tr2[tdTypeIdx]    || '').trim();
          var tPeriod2  = String(tr2[tdPeriodIdx]  || '').trim();

          // 衝突判斷
          var isConflict = false;
          if (tType2 === '日代' || atype === '日代') {
            isConflict = true;
          } else if (period) {
            var tPnums = tPeriod2.match(/\d/g) || [];
            isConflict = tPnums.indexOf(period) >= 0;
          } else {
            isConflict = true;
          }

          if (isConflict) {
            var prev = conflicts[tTeacher2];
            if (tStatus2 === '已確認' || tStatus2 === '已完成') {
              conflicts[tTeacher2] = '已確認';
            } else if (tStatus2 === '已詢問' && prev !== '已確認') {
              conflicts[tTeacher2] = '已詢問';
            }
          }
        }
      }

      // 組合結果
      var avResult = candidates.map(function(c) {
        var stripped = c.name.replace(/老師$/, '').trim();
        var cs = conflicts[stripped] || conflicts[c.name] || '可詢問';
        return { name: c.name, status: cs, note: c.note };
      });
      return jsonResponse({ teachers: avResult });
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

    // 更新單筆資料（例如：已詢問 → 已確認）
    if (payload.action === 'update') {
      var updateId = payload.id;
      var changes = payload.changes || {};
      if (!updateId) return jsonResponse({ ok: false, error: '未指定要更新的 id' });

      var updateData = sheet.getDataRange().getValues();
      var updateIdCol = HEADERS.indexOf('id');
      for (var ur = 1; ur < updateData.length; ur++) {
        if (String(updateData[ur][updateIdCol]) !== String(updateId)) continue;

        for (var uc = 0; uc < HEADERS.length; uc++) {
          var updateKey = HEADERS[uc];
          if (updateKey === 'id' || changes[updateKey] === undefined) continue;
          sheet.getRange(ur + 1, uc + 1).setValue(changes[updateKey]);
        }
        if (changes['更新時間'] === undefined) {
          sheet.getRange(ur + 1, HEADERS.indexOf('更新時間') + 1).setValue(new Date());
        }
        return jsonResponse({ ok: true, updated: updateId });
      }
      return jsonResponse({ ok: false, error: '找不到該筆資料' });
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
