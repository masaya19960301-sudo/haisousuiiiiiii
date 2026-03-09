/**
 * ============================================================
 * 家具配送センター 満車管理Webアプリ
 * ============================================================
 *
 * 【セットアップ手順】
 * 1. Googleスプレッドシートを新規作成
 * 2. 「拡張機能」→「Apps Script」でスクリプトエディタを開く
 * 3. このファイル（コード.gs）の内容をスクリプトエディタに貼り付け
 * 4. index.html を HTMLファイルとして追加
 * 5. スクリプトプロパティに SPREADSHEET_ID を設定
 *    （ファイル→プロジェクトの設定→スクリプトプロパティ）
 * 6. setup() 関数を実行（シート・ヘッダー・サンプルデータを自動生成）
 * 7. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」として公開
 *
 * ============================================================
 */

// ==================== 定数 ====================

var SHEET_NAMES = {
  MAIN: '満車管理',
  DAY_SETTINGS: '曜日設定',
  SPECIAL_DAYS: '特別日設定',
  CHANGE_LOG: '変更ログ'
};

var MAIN_HEADERS = [
  '配送日', '曜日', '台数', 'ステータス',
  '合計件数', '合計金額',
  '1台あたり件数', '1台あたり金額',
  '昨日までの合計件数', '昨日までの合計金額',
  '当日契約件数', '当日契約金額',
  '備考'
];

var DAY_SETTINGS_HEADERS = ['曜日', '通常台数', '有効フラグ'];

var SPECIAL_DAYS_HEADERS = ['日付', '台数', '種別', 'メモ'];

var CHANGE_LOG_HEADERS = ['変更日時', '変更者', '対象日付', '変更前', '変更後', '項目名'];

var DAY_NAMES = ['日', '月', '火', '水', '木', '金', '土'];

// ==================== スプレッドシート取得 ====================

function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#4a86c8')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ==================== セットアップ ====================

function setup() {
  var ss = getSpreadsheet();

  // シート作成
  getOrCreateSheet(ss, SHEET_NAMES.MAIN, MAIN_HEADERS);
  getOrCreateSheet(ss, SHEET_NAMES.DAY_SETTINGS, DAY_SETTINGS_HEADERS);
  getOrCreateSheet(ss, SHEET_NAMES.SPECIAL_DAYS, SPECIAL_DAYS_HEADERS);
  getOrCreateSheet(ss, SHEET_NAMES.CHANGE_LOG, CHANGE_LOG_HEADERS);

  // デフォルトの曜日設定を挿入
  var daySheet = ss.getSheetByName(SHEET_NAMES.DAY_SETTINGS);
  if (daySheet.getLastRow() <= 1) {
    var defaults = [
      ['月', 3, '配送あり'],
      ['火', 3, '配送あり'],
      ['水', 3, '配送あり'],
      ['木', 3, '配送あり'],
      ['金', 3, '配送あり'],
      ['土', 2, '配送あり'],
      ['日', 0, '配送なし']
    ];
    daySheet.getRange(2, 1, defaults.length, defaults[0].length).setValues(defaults);
  }

  // サンプルデータ（直近2週間分）を挿入
  var mainSheet = ss.getSheetByName(SHEET_NAMES.MAIN);
  if (mainSheet.getLastRow() <= 1) {
    insertSampleData(mainSheet);
  }

  // デフォルトのSheet1を削除（存在する場合）
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }
}

function insertSampleData(sheet) {
  var today = new Date();
  var rows = [];

  for (var i = -7; i <= 7; i++) {
    var d = new Date(today);
    d.setDate(d.getDate() + i);
    var dayIndex = d.getDay(); // 0=日, 1=月, ...
    var dayName = DAY_NAMES[dayIndex];

    // 日曜は休業
    if (dayIndex === 0) continue;

    var trucks = (dayIndex === 6) ? 2 : 3;
    var status = '受付中';
    var yesterdayCount = Math.floor(Math.random() * 20) + 5;
    var yesterdayAmount = Math.floor(Math.random() * 500000) + 100000;
    var todayCount = Math.floor(Math.random() * 10);
    var todayAmount = Math.floor(Math.random() * 300000);

    if (i < 0) {
      status = '締切済';
    } else if (i === 0) {
      status = '受付中';
    }

    var totalCount = yesterdayCount + todayCount;
    var totalAmount = yesterdayAmount + todayAmount;
    var perTruckCount = trucks > 0 ? Math.round(totalCount / trucks * 10) / 10 : 0;
    var perTruckAmount = trucks > 0 ? Math.round(totalAmount / trucks) : 0;

    rows.push([
      formatDate(d), dayName, trucks, status,
      totalCount, totalAmount,
      perTruckCount, perTruckAmount,
      yesterdayCount, yesterdayAmount,
      todayCount, todayAmount,
      ''
    ]);
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

// ==================== Web App ====================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('満車管理アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==================== データ取得 ====================

function getRows(startDate, endDate) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.MAIN);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MAIN_HEADERS.length).getValues();
  var result = [];

  var start = startDate ? new Date(startDate) : null;
  var end = endDate ? new Date(endDate) : null;
  if (start) start.setHours(0, 0, 0, 0);
  if (end) end.setHours(23, 59, 59, 999);

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateVal = row[0];
    if (!dateVal) continue;

    var rowDate;
    if (dateVal instanceof Date) {
      rowDate = dateVal;
    } else {
      rowDate = new Date(dateVal);
    }

    if (start && rowDate < start) continue;
    if (end && rowDate > end) continue;

    result.push({
      date: formatDate(rowDate),
      dayName: row[1],
      trucks: row[2],
      status: row[3],
      totalCount: row[4],
      totalAmount: row[5],
      perTruckCount: row[6],
      perTruckAmount: row[7],
      yesterdayCount: row[8],
      yesterdayAmount: row[9],
      todayCount: row[10],
      todayAmount: row[11],
      memo: row[12] || ''
    });
  }

  // 日付でソート
  result.sort(function(a, b) {
    return new Date(a.date) - new Date(b.date);
  });

  return result;
}

// ==================== セル保存 ====================

function saveCell(date, column, value) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.MAIN);
  if (!sheet) throw new Error('満車管理シートが見つかりません');

  var rowIndex = findRowByDate(sheet, date);
  if (rowIndex === -1) throw new Error('該当日のデータが見つかりません: ' + date);

  var colMap = {
    'trucks': 3,
    'status': 4,
    'yesterdayCount': 9,
    'yesterdayAmount': 10,
    'todayCount': 11,
    'todayAmount': 12,
    'memo': 13
  };

  var colIndex = colMap[column];
  if (!colIndex) throw new Error('無効な列名: ' + column);

  // 変更ログ記録
  var oldValue = sheet.getRange(rowIndex, colIndex).getValue();
  logChange(date, column, String(oldValue), String(value));

  // 値を設定
  var numericColumns = ['trucks', 'yesterdayCount', 'yesterdayAmount', 'todayCount', 'todayAmount'];
  if (numericColumns.indexOf(column) !== -1) {
    value = Number(value) || 0;
  }
  sheet.getRange(rowIndex, colIndex).setValue(value);

  // 自動計算列を更新
  recalculateRow(sheet, rowIndex);

  // 更新後の行データを返す
  return getRowData(sheet, rowIndex);
}

function recalculateRow(sheet, rowIndex) {
  var trucks = Number(sheet.getRange(rowIndex, 3).getValue()) || 0;
  var yesterdayCount = Number(sheet.getRange(rowIndex, 9).getValue()) || 0;
  var yesterdayAmount = Number(sheet.getRange(rowIndex, 10).getValue()) || 0;
  var todayCount = Number(sheet.getRange(rowIndex, 11).getValue()) || 0;
  var todayAmount = Number(sheet.getRange(rowIndex, 12).getValue()) || 0;

  var totalCount = yesterdayCount + todayCount;
  var totalAmount = yesterdayAmount + todayAmount;
  var perTruckCount = trucks > 0 ? Math.round(totalCount / trucks * 10) / 10 : 0;
  var perTruckAmount = trucks > 0 ? Math.round(totalAmount / trucks) : 0;

  sheet.getRange(rowIndex, 5).setValue(totalCount);
  sheet.getRange(rowIndex, 6).setValue(totalAmount);
  sheet.getRange(rowIndex, 7).setValue(perTruckCount);
  sheet.getRange(rowIndex, 8).setValue(perTruckAmount);
}

function getRowData(sheet, rowIndex) {
  var row = sheet.getRange(rowIndex, 1, 1, MAIN_HEADERS.length).getValues()[0];
  var dateVal = row[0] instanceof Date ? formatDate(row[0]) : String(row[0]);
  return {
    date: dateVal,
    dayName: row[1],
    trucks: row[2],
    status: row[3],
    totalCount: row[4],
    totalAmount: row[5],
    perTruckCount: row[6],
    perTruckAmount: row[7],
    yesterdayCount: row[8],
    yesterdayAmount: row[9],
    todayCount: row[10],
    todayAmount: row[11],
    memo: row[12] || ''
  };
}

// ==================== 行操作 ====================

function addRow(date) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.MAIN);
  if (!sheet) throw new Error('満車管理シートが見つかりません');

  // 重複チェック
  if (findRowByDate(sheet, date) !== -1) {
    throw new Error('指定日 ' + date + ' のデータは既に存在します');
  }

  var d = new Date(date);
  var dayIndex = d.getDay();
  var dayName = DAY_NAMES[dayIndex];

  // 台数を取得（特別日 > 曜日設定）
  var trucks = getTrucksForDate(date);

  var newRow = [
    date, dayName, trucks, '受付中',
    0, 0, 0, 0,
    0, 0, 0, 0,
    ''
  ];

  sheet.appendRow(newRow);

  return {
    date: date,
    dayName: dayName,
    trucks: trucks,
    status: '受付中',
    totalCount: 0,
    totalAmount: 0,
    perTruckCount: 0,
    perTruckAmount: 0,
    yesterdayCount: 0,
    yesterdayAmount: 0,
    todayCount: 0,
    todayAmount: 0,
    memo: ''
  };
}

function deleteRow(date) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.MAIN);
  if (!sheet) throw new Error('満車管理シートが見つかりません');

  var rowIndex = findRowByDate(sheet, date);
  if (rowIndex === -1) throw new Error('該当日のデータが見つかりません: ' + date);

  // 変更ログ
  logChange(date, '行削除', '存在', '削除');

  sheet.deleteRow(rowIndex);
  return { success: true };
}

function getTrucksForDate(date) {
  var d = new Date(date);
  var ss = getSpreadsheet();

  // 特別日設定を確認
  var specialSheet = ss.getSheetByName(SHEET_NAMES.SPECIAL_DAYS);
  if (specialSheet && specialSheet.getLastRow() > 1) {
    var specialData = specialSheet.getRange(2, 1, specialSheet.getLastRow() - 1, 4).getValues();
    for (var i = 0; i < specialData.length; i++) {
      var specialDate = specialData[i][0];
      if (specialDate instanceof Date) {
        specialDate = formatDate(specialDate);
      }
      if (specialDate === date) {
        var type = specialData[i][2];
        if (type === '休業日') return 0;
        return Number(specialData[i][1]) || 0;
      }
    }
  }

  // 曜日設定から取得
  var dayIndex = d.getDay();
  var dayName = DAY_NAMES[dayIndex];
  var daySheet = ss.getSheetByName(SHEET_NAMES.DAY_SETTINGS);
  if (daySheet && daySheet.getLastRow() > 1) {
    var dayData = daySheet.getRange(2, 1, daySheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < dayData.length; i++) {
      if (dayData[i][0] === dayName) {
        if (dayData[i][2] === '配送なし') return 0;
        return Number(dayData[i][1]) || 0;
      }
    }
  }

  return 3; // デフォルト
}

// ==================== 曜日設定 ====================

function getDaySettings() {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.DAY_SETTINGS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data.map(function(row) {
    return {
      dayName: row[0],
      trucks: row[1],
      enabled: row[2]
    };
  });
}

function saveDaySettings(settings) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.DAY_SETTINGS);
  if (!sheet) throw new Error('曜日設定シートが見つかりません');

  // 既存データをクリア
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
  }

  var rows = settings.map(function(s) {
    return [s.dayName, Number(s.trucks), s.enabled];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  return { success: true };
}

// ==================== 特別日設定 ====================

function getSpecialDays() {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.SPECIAL_DAYS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return data.map(function(row) {
    var dateVal = row[0] instanceof Date ? formatDate(row[0]) : String(row[0]);
    return {
      date: dateVal,
      trucks: row[1],
      type: row[2],
      memo: row[3] || ''
    };
  }).sort(function(a, b) {
    return new Date(a.date) - new Date(b.date);
  });
}

function saveSpecialDay(data) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.SPECIAL_DAYS);
  if (!sheet) throw new Error('特別日設定シートが見つかりません');

  // 既存の同一日付を検索
  var existingRow = -1;
  if (sheet.getLastRow() > 1) {
    var dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < dates.length; i++) {
      var d = dates[i][0] instanceof Date ? formatDate(dates[i][0]) : String(dates[i][0]);
      if (d === data.date) {
        existingRow = i + 2;
        break;
      }
    }
  }

  var rowData = [data.date, Number(data.trucks) || 0, data.type, data.memo || ''];

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, 4).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return { success: true };
}

function deleteSpecialDay(date) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.SPECIAL_DAYS);
  if (!sheet || sheet.getLastRow() <= 1) return { success: true };

  var dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = dates.length - 1; i >= 0; i--) {
    var d = dates[i][0] instanceof Date ? formatDate(dates[i][0]) : String(dates[i][0]);
    if (d === date) {
      sheet.deleteRow(i + 2);
      break;
    }
  }

  return { success: true };
}

// ==================== CSVエクスポート ====================

function exportCSV(startDate, endDate) {
  var rows = getRows(startDate, endDate);
  if (rows.length === 0) return '';

  var headers = ['配送日', '曜日', '台数', 'ステータス', '合計件数', '合計金額',
    '1台あたり件数', '1台あたり金額', '昨日までの合計件数', '昨日までの合計金額',
    '当日契約件数', '当日契約金額', '備考'];

  var csv = headers.join(',') + '\n';

  rows.forEach(function(r) {
    csv += [
      r.date, r.dayName, r.trucks, r.status,
      r.totalCount, r.totalAmount,
      r.perTruckCount, r.perTruckAmount,
      r.yesterdayCount, r.yesterdayAmount,
      r.todayCount, r.todayAmount,
      '"' + (r.memo || '').replace(/"/g, '""') + '"'
    ].join(',') + '\n';
  });

  return csv;
}

// ==================== 変更ログ ====================

function getChangeLog(limit) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.CHANGE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  var result = data.map(function(row) {
    return {
      timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(row[0]),
      user: row[1],
      targetDate: row[2],
      oldValue: row[3],
      newValue: row[4],
      field: row[5]
    };
  });

  // 新しい順にソート
  result.reverse();

  if (limit && limit > 0) {
    result = result.slice(0, limit);
  }

  return result;
}

function logChange(targetDate, field, oldValue, newValue) {
  try {
    var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.CHANGE_LOG);
    if (!sheet) return;

    var user = Session.getActiveUser().getEmail() || '不明';
    var now = new Date();
    var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([timestamp, user, targetDate, oldValue, newValue, field]);
  } catch (e) {
    // ログ記録失敗は無視
  }
}

// ==================== ユーティリティ ====================

function findRowByDate(sheet, date) {
  if (sheet.getLastRow() <= 1) return -1;

  var dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0];
    var formatted;
    if (d instanceof Date) {
      formatted = formatDate(d);
    } else {
      formatted = String(d);
    }
    if (formatted === date) {
      return i + 2; // 1-indexed + header row
    }
  }
  return -1;
}

function formatDate(d) {
  var year = d.getFullYear();
  var month = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  return year + '-' + month + '-' + day;
}
