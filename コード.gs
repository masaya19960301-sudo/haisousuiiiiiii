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
 * 【金額について】
 * 全ての金額は千円単位で入力・保存されます。
 * 例: 150 = 150千円 = 15万円
 *
 * ============================================================
 */

// ==================== 定数 ====================

var SHEET_NAMES = {
  MAIN: '満車管理',
  DAY_SETTINGS: '曜日設定',
  SPECIAL_DAYS: '特別日設定',
  HOLIDAYS: '祝日・販促期間',
  CHANGE_LOG: '変更ログ',
  CONTRACT_LOG: '契約履歴ログ',
  AREA_MASTER: 'エリアマスタ',
  AREA_EXCEPTION: 'エリア例外設定'
};

var MAIN_HEADERS = [
  '配送日', '曜日', '台数', 'ステータス',
  '合計件数', '合計金額(千円)',
  '1台あたり件数', '1台あたり金額(千円)',
  '昨日までの合計件数', '昨日までの合計金額(千円)',
  '当日契約件数', '当日契約金額(千円)',
  '備考', '満車確定日', '満車時件数'
];

var DAY_SETTINGS_HEADERS = ['曜日', '通常台数', '有効フラグ'];

var SPECIAL_DAYS_HEADERS = ['日付', '台数', '種別', 'メモ'];

var HOLIDAYS_HEADERS = ['日付', '終了日', '種別', '名称', 'メモ'];

var CHANGE_LOG_HEADERS = ['変更日時', '変更者', '対象日付', '変更前', '変更後', '項目名'];

var CONTRACT_LOG_HEADERS = [
  '契約日', '契約曜日', '契約日販促フラグ',
  '配送日', '配送曜日', '配送日販促フラグ',
  'リードタイム(日)', '件数', '金額',
  '契約時点の最短空き日', '最短空きまでの日数', '最短空きの曜日'
];

var AREA_MASTER_HEADERS = ['エリアID', 'エリア名', '通常配送曜日', '備考'];

var AREA_EXCEPTION_HEADERS = ['日付', 'エリアID', '例外種別', '理由'];

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

  getOrCreateSheet(ss, SHEET_NAMES.MAIN, MAIN_HEADERS);
  getOrCreateSheet(ss, SHEET_NAMES.DAY_SETTINGS, DAY_SETTINGS_HEADERS);
  getOrCreateSheet(ss, SHEET_NAMES.SPECIAL_DAYS, SPECIAL_DAYS_HEADERS);
  getOrCreateSheet(ss, SHEET_NAMES.HOLIDAYS, HOLIDAYS_HEADERS);
  getOrCreateSheet(ss, SHEET_NAMES.CHANGE_LOG, CHANGE_LOG_HEADERS);
  getOrCreateSheet(ss, SHEET_NAMES.CONTRACT_LOG, CONTRACT_LOG_HEADERS);
  getOrCreateSheet(ss, SHEET_NAMES.AREA_MASTER, AREA_MASTER_HEADERS);
  getOrCreateSheet(ss, SHEET_NAMES.AREA_EXCEPTION, AREA_EXCEPTION_HEADERS);

  // 既存満車管理シートに新列がなければ追加
  var mainSheet = ss.getSheetByName(SHEET_NAMES.MAIN);
  var headerRow = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  if (headerRow.indexOf('満車確定日') === -1) {
    var nextCol = headerRow.length + 1;
    mainSheet.getRange(1, nextCol).setValue('満車確定日').setFontWeight('bold').setBackground('#4a86c8').setFontColor('#ffffff');
    mainSheet.getRange(1, nextCol + 1).setValue('満車時件数').setFontWeight('bold').setBackground('#4a86c8').setFontColor('#ffffff');
  }

  // デフォルトの曜日設定
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

  // サンプル祝日データ
  var holidaySheet = ss.getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (holidaySheet.getLastRow() <= 1) {
    var year = new Date().getFullYear();
    var sampleHolidays = [
      [year + '-01-01', '', '祝日', '元日', ''],
      [year + '-01-13', '', '祝日', '成人の日', ''],
      [year + '-02-11', '', '祝日', '建国記念の日', ''],
      [year + '-02-23', '', '祝日', '天皇誕生日', ''],
      [year + '-03-20', '', '祝日', '春分の日', ''],
      [year + '-04-29', '', '祝日', '昭和の日', ''],
      [year + '-05-03', year + '-05-05', '祝日', '憲法記念日〜こどもの日', 'GW連休'],
      [year + '-07-21', '', '祝日', '海の日', ''],
      [year + '-08-11', '', '祝日', '山の日', ''],
      [year + '-09-15', '', '祝日', '敬老の日', ''],
      [year + '-09-23', '', '祝日', '秋分の日', ''],
      [year + '-10-13', '', '祝日', 'スポーツの日', ''],
      [year + '-11-03', '', '祝日', '文化の日', ''],
      [year + '-11-23', '', '祝日', '勤労感謝の日', ''],
      [year + '-03-01', year + '-03-31', '販促期間', '新生活応援フェア', '引越しシーズン'],
      [year + '-08-01', year + '-08-31', '販促期間', 'サマーセール', ''],
      [year + '-12-01', year + '-12-25', '販促期間', '年末セール', '']
    ];
    holidaySheet.getRange(2, 1, sampleHolidays.length, 5).setValues(sampleHolidays);
  }

  // サンプルデータ（直近の配送日30日分）
  if (mainSheet.getLastRow() <= 1) {
    insertSampleData(mainSheet, ss);
  }

  // デフォルトのSheet1を削除
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }
}

function insertSampleData(sheet, ss) {
  var today = new Date();
  var rows = [];
  var daySettings = getDaySettingsMap(ss);
  var specialDaysMap = getSpecialDaysMap(ss);
  var count = 0;
  var maxDays = 30;

  for (var i = -10; count < maxDays && i < 60; i++) {
    var d = new Date(today);
    d.setDate(d.getDate() + i);
    var dateStr = formatDate(d);
    var dayIndex = d.getDay();
    var dayName = DAY_NAMES[dayIndex];

    var trucks = resolveTrucks(dateStr, dayIndex, daySettings, specialDaysMap);
    if (trucks <= 0) continue;

    var status = '受付中';
    var yesterdayCount = Math.floor(Math.random() * 15) + 5;
    var yesterdayAmount = Math.floor(Math.random() * 500) + 100;
    var todayCount = Math.floor(Math.random() * 8);
    var todayAmount = Math.floor(Math.random() * 300);

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
      '', '', ''
    ]);
    count++;
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

// ==================== 自動30日生成 ====================

function getNext30DeliveryDays() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.MAIN);
  var daySettings = getDaySettingsMap(ss);
  var specialDaysMap = getSpecialDaysMap(ss);

  deduplicateMainSheet(sheet);

  var existingMap = {};
  var numCols = Math.max(sheet.getLastColumn(), MAIN_HEADERS.length);
  if (sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
    for (var i = 0; i < data.length; i++) {
      var dateVal = data[i][0];
      var key = dateVal instanceof Date ? formatDate(dateVal) : String(dateVal);
      existingMap[key] = {
        date: key,
        dayName: data[i][1],
        trucks: data[i][2],
        status: data[i][3],
        totalCount: data[i][4],
        totalAmount: data[i][5],
        perTruckCount: data[i][6],
        perTruckAmount: data[i][7],
        yesterdayCount: data[i][8],
        yesterdayAmount: data[i][9],
        todayCount: data[i][10],
        todayAmount: data[i][11],
        memo: data[i][12] || '',
        manshaDate: data[i][13] || '',
        manshaCount: data[i][14] || ''
      };
    }
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var result = [];
  var count = 0;

  for (var i = 0; count < 30 && i < 90; i++) {
    var d = new Date(today);
    d.setDate(d.getDate() + i);
    var dateStr = formatDate(d);
    var dayIndex = d.getDay();

    var trucks = resolveTrucks(dateStr, dayIndex, daySettings, specialDaysMap);
    if (trucks <= 0) continue;

    if (existingMap[dateStr]) {
      var row = existingMap[dateStr];
      row.defaultTrucks = trucks;
      result.push(row);
    } else {
      var dayName = DAY_NAMES[dayIndex];
      var newRow = [dateStr, dayName, trucks, '受付中', 0, 0, 0, 0, 0, 0, 0, 0, '', '', ''];
      sheet.appendRow(newRow);
      result.push({
        date: dateStr, dayName: dayName, trucks: trucks, status: '受付中',
        totalCount: 0, totalAmount: 0, perTruckCount: 0, perTruckAmount: 0,
        yesterdayCount: 0, yesterdayAmount: 0, todayCount: 0, todayAmount: 0,
        memo: '', manshaDate: '', manshaCount: '', defaultTrucks: trucks
      });
    }
    count++;
  }

  return result;
}

// ==================== データ取得 ====================

function getRows(startDate, endDate) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.MAIN);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  deduplicateMainSheet(sheet);

  var daySettings = getDaySettingsMap(ss);
  var specialDaysMap = getSpecialDaysMap(ss);
  var numCols = Math.max(sheet.getLastColumn(), MAIN_HEADERS.length);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  var result = [];

  var start = startDate ? new Date(startDate) : null;
  var end = endDate ? new Date(endDate) : null;
  if (start) start.setHours(0, 0, 0, 0);
  if (end) end.setHours(23, 59, 59, 999);

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateVal = row[0];
    if (!dateVal) continue;

    var rowDate = dateVal instanceof Date ? dateVal : new Date(dateVal);
    if (start && rowDate < start) continue;
    if (end && rowDate > end) continue;

    var dateStr = formatDate(rowDate);
    var dayIndex = rowDate.getDay();
    var defTrucks = resolveTrucks(dateStr, dayIndex, daySettings, specialDaysMap);

    result.push({
      date: dateStr,
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
      memo: row[12] || '',
      manshaDate: row[13] || '',
      manshaCount: row[14] || '',
      defaultTrucks: defTrucks
    });
  }

  result.sort(function(a, b) { return new Date(a.date) - new Date(b.date); });
  return result;
}

// ==================== カレンダー用データ取得 ====================

function getCalendarData(year, month) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.MAIN);
  var daySettings = getDaySettingsMap(ss);
  var specialDaysMap = getSpecialDaysMap(ss);

  // 月の初日と末日
  var firstDay = new Date(year, month - 1, 1);
  var lastDay = new Date(year, month, 0);
  var startStr = formatDate(firstDay);
  var endStr = formatDate(lastDay);

  var rows = getRows(startStr, endStr);
  var rowMap = {};
  rows.forEach(function(r) { rowMap[r.date] = r; });

  // 祝日・販促期間
  var holidays = getHolidays();
  var holidayDateSet = {};
  var promoDateSet = {};
  for (var h = 0; h < holidays.length; h++) {
    var hol = holidays[h];
    var hStart = new Date(hol.date);
    var hEnd = hol.endDate ? new Date(hol.endDate) : hStart;
    for (var dd = new Date(hStart); dd <= hEnd; dd.setDate(dd.getDate() + 1)) {
      var ds = formatDate(dd);
      if (hol.type === '祝日') holidayDateSet[ds] = hol.name;
      else if (hol.type === '販促期間') promoDateSet[ds] = hol.name;
    }
  }

  // 配送日を持つ日のセット
  var deliveryDays = {};
  for (var d = new Date(firstDay); d <= lastDay; d.setDate(d.getDate() + 1)) {
    var ds2 = formatDate(d);
    var di = d.getDay();
    var trucks = resolveTrucks(ds2, di, daySettings, specialDaysMap);
    if (trucks > 0) deliveryDays[ds2] = trucks;
  }

  // 全配送データから最短空きを計算（今日以降で受付中/注意の最も近い日）
  var todayStr = formatDate(new Date());
  var allRows = getRows(todayStr, null);
  var nearestOpen = '';
  for (var i = 0; i < allRows.length; i++) {
    if (allRows[i].date >= todayStr && (allRows[i].status === '受付中' || allRows[i].status === '注意')) {
      nearestOpen = allRows[i].date;
      break;
    }
  }

  var calDays = [];
  for (var d2 = new Date(firstDay); d2 <= lastDay; d2.setDate(d2.getDate() + 1)) {
    var ds3 = formatDate(d2);
    var rowData = rowMap[ds3] || null;
    var isDelivery = !!deliveryDays[ds3];
    calDays.push({
      date: ds3,
      dayName: DAY_NAMES[d2.getDay()],
      dayOfWeek: d2.getDay(),
      isDelivery: isDelivery,
      defaultTrucks: deliveryDays[ds3] || 0,
      status: rowData ? rowData.status : (isDelivery ? '受付中' : ''),
      trucks: rowData ? rowData.trucks : (deliveryDays[ds3] || 0),
      totalCount: rowData ? rowData.totalCount : 0,
      totalAmount: rowData ? rowData.totalAmount : 0,
      perTruckCount: rowData ? rowData.perTruckCount : 0,
      memo: rowData ? rowData.memo : '',
      isHoliday: !!holidayDateSet[ds3],
      holidayName: holidayDateSet[ds3] || '',
      isPromo: !!promoDateSet[ds3],
      promoName: promoDateSet[ds3] || '',
      manshaDate: rowData ? (rowData.manshaDate || '') : '',
      manshaCount: rowData ? (rowData.manshaCount || '') : ''
    });
  }

  return {
    year: year,
    month: month,
    days: calDays,
    nearestOpen: nearestOpen,
    today: todayStr
  };
}

// ==================== セル保存 ====================

function saveCell(date, column, value) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.MAIN);
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

  var oldValue = sheet.getRange(rowIndex, colIndex).getValue();
  logChange(date, column, String(oldValue), String(value));

  var numericColumns = ['trucks', 'yesterdayCount', 'yesterdayAmount', 'todayCount', 'todayAmount'];
  if (numericColumns.indexOf(column) !== -1) {
    value = Number(value) || 0;
  }
  sheet.getRange(rowIndex, colIndex).setValue(value);

  // ステータスが「満車」に変わった場合、満車確定日と満車時件数を記録
  if (column === 'status' && value === '満車' && oldValue !== '満車') {
    var todayStr = formatDate(new Date());
    var totalCount = Number(sheet.getRange(rowIndex, 5).getValue()) || 0;
    sheet.getRange(rowIndex, 14).setValue(todayStr);
    sheet.getRange(rowIndex, 15).setValue(totalCount);
  }

  recalculateRow(sheet, rowIndex);

  // 当日契約件数・金額が保存されたら契約履歴ログに記録
  if (column === 'todayCount' || column === 'todayAmount') {
    try {
      recordContractLog(ss, sheet, rowIndex, date);
    } catch(e) {}
  }

  return getRowData(sheet, rowIndex);
}

// ==================== 契約履歴ログ自動記録 ====================

function recordContractLog(ss, mainSheet, rowIndex, deliveryDate) {
  var todayCount = Number(mainSheet.getRange(rowIndex, 11).getValue()) || 0;
  var todayAmount = Number(mainSheet.getRange(rowIndex, 12).getValue()) || 0;

  // 件数・金額が両方0なら記録しない
  if (todayCount === 0 && todayAmount === 0) return;

  var today = new Date();
  var contractDateStr = formatDate(today);
  var contractDayName = DAY_NAMES[today.getDay()];

  var deliveryD = new Date(deliveryDate);
  var deliveryDayName = DAY_NAMES[deliveryD.getDay()];
  var leadTime = Math.round((deliveryD - today) / (1000 * 60 * 60 * 24));

  // 販促フラグ
  var holidays = getHolidays();
  var promoDateSet = {};
  for (var h = 0; h < holidays.length; h++) {
    var hol = holidays[h];
    if (hol.type !== '販促期間') continue;
    var hStart = new Date(hol.date);
    var hEnd = hol.endDate ? new Date(hol.endDate) : hStart;
    for (var dd = new Date(hStart); dd <= hEnd; dd.setDate(dd.getDate() + 1)) {
      promoDateSet[formatDate(dd)] = true;
    }
  }
  var contractPromo = promoDateSet[contractDateStr] ? 1 : 0;
  var deliveryPromo = promoDateSet[deliveryDate] ? 1 : 0;

  // 最短空き日を検索
  var todayStr = contractDateStr;
  var allRows = getRows(todayStr, null);
  var nearestOpen = '';
  var nearestDays = '';
  var nearestDayName = '';
  for (var i = 0; i < allRows.length; i++) {
    if (allRows[i].date >= todayStr && (allRows[i].status === '受付中' || allRows[i].status === '注意')) {
      nearestOpen = allRows[i].date;
      nearestDays = Math.round((new Date(allRows[i].date) - today) / (1000 * 60 * 60 * 24));
      nearestDayName = allRows[i].dayName;
      break;
    }
  }

  var logSheet = getOrCreateSheet(ss, SHEET_NAMES.CONTRACT_LOG, CONTRACT_LOG_HEADERS);

  // 同じ契約日×配送日が既存なら上書き
  var existingRow = -1;
  if (logSheet.getLastRow() > 1) {
    var logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues();
    for (var j = 0; j < logData.length; j++) {
      var ld = logData[j][0] instanceof Date ? formatDate(logData[j][0]) : String(logData[j][0]);
      var dd2 = logData[j][3] instanceof Date ? formatDate(logData[j][3]) : String(logData[j][3]);
      if (ld === contractDateStr && dd2 === deliveryDate) {
        existingRow = j + 2;
        break;
      }
    }
  }

  var rowData = [
    contractDateStr, contractDayName, contractPromo,
    deliveryDate, deliveryDayName, deliveryPromo,
    leadTime, todayCount, todayAmount,
    nearestOpen, nearestDays, nearestDayName
  ];

  if (existingRow > 0) {
    logSheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    logSheet.appendRow(rowData);
  }
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
  var numCols = Math.max(sheet.getLastColumn(), MAIN_HEADERS.length);
  var row = sheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];
  var dateVal = row[0] instanceof Date ? formatDate(row[0]) : String(row[0]);
  var rowDate = row[0] instanceof Date ? row[0] : new Date(row[0]);
  var ss = getSpreadsheet();
  var daySettings = getDaySettingsMap(ss);
  var specialDaysMap = getSpecialDaysMap(ss);
  var defTrucks = resolveTrucks(dateVal, rowDate.getDay(), daySettings, specialDaysMap);
  return {
    date: dateVal, dayName: row[1], trucks: row[2], status: row[3],
    totalCount: row[4], totalAmount: row[5], perTruckCount: row[6], perTruckAmount: row[7],
    yesterdayCount: row[8], yesterdayAmount: row[9], todayCount: row[10], todayAmount: row[11],
    memo: row[12] || '', manshaDate: row[13] || '', manshaCount: row[14] || '',
    defaultTrucks: defTrucks
  };
}

// ==================== 行操作 ====================

function addRow(date) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.MAIN);
  if (!sheet) throw new Error('満車管理シートが見つかりません');

  var existingRowIndex = findRowByDate(sheet, date);
  if (existingRowIndex !== -1) {
    return getRowData(sheet, existingRowIndex);
  }

  var d = new Date(date);
  var dayIndex = d.getDay();
  var dayName = DAY_NAMES[dayIndex];
  var trucks = getTrucksForDate(date);
  var newRow = [date, dayName, trucks, '受付中', 0, 0, 0, 0, 0, 0, 0, 0, '', '', ''];
  sheet.appendRow(newRow);

  return {
    date: date, dayName: dayName, trucks: trucks, status: '受付中',
    totalCount: 0, totalAmount: 0, perTruckCount: 0, perTruckAmount: 0,
    yesterdayCount: 0, yesterdayAmount: 0, todayCount: 0, todayAmount: 0,
    memo: '', manshaDate: '', manshaCount: ''
  };
}

function deleteRow(date) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.MAIN);
  if (!sheet) throw new Error('満車管理シートが見つかりません');
  var rowIndex = findRowByDate(sheet, date);
  if (rowIndex === -1) throw new Error('該当日のデータが見つかりません: ' + date);
  logChange(date, '行削除', '存在', '削除');
  sheet.deleteRow(rowIndex);
  return { success: true };
}

function getTrucksForDate(date) {
  var ss = getSpreadsheet();
  var d = new Date(date);
  var dayIndex = d.getDay();
  var daySettings = getDaySettingsMap(ss);
  var specialDaysMap = getSpecialDaysMap(ss);
  return resolveTrucks(date, dayIndex, daySettings, specialDaysMap);
}

// ==================== 内部ヘルパー: 台数解決 ====================

function getDaySettingsMap(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.DAY_SETTINGS);
  var map = {};
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      map[data[i][0]] = { trucks: Number(data[i][1]) || 0, enabled: data[i][2] };
    }
  }
  return map;
}

function getSpecialDaysMap(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.SPECIAL_DAYS);
  var map = {};
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    for (var i = 0; i < data.length; i++) {
      var dateVal = data[i][0] instanceof Date ? formatDate(data[i][0]) : String(data[i][0]);
      map[dateVal] = { trucks: Number(data[i][1]) || 0, type: data[i][2] };
    }
  }
  return map;
}

function resolveTrucks(dateStr, dayIndex, daySettings, specialDaysMap) {
  if (specialDaysMap[dateStr]) {
    if (specialDaysMap[dateStr].type === '休業日') return 0;
    return specialDaysMap[dateStr].trucks;
  }
  var dayName = DAY_NAMES[dayIndex];
  if (daySettings[dayName]) {
    if (daySettings[dayName].enabled === '配送なし') return 0;
    return daySettings[dayName].trucks;
  }
  return 3;
}

// ==================== 曜日設定 ====================

function getDaySettings() {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.DAY_SETTINGS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data.map(function(row) {
    return { dayName: row[0], trucks: row[1], enabled: row[2] };
  });
}

function saveDaySettings(settings) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.DAY_SETTINGS);
  if (!sheet) throw new Error('曜日設定シートが見つかりません');
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
  }
  var rows = settings.map(function(s) { return [s.dayName, Number(s.trucks), s.enabled]; });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  return { success: true };
}

// ==================== 目標設定 ====================

function getTargets() {
  var props = PropertiesService.getScriptProperties();
  return {
    targetCount: Number(props.getProperty('TARGET_PER_TRUCK_COUNT')) || 0,
    targetAmount: Number(props.getProperty('TARGET_PER_TRUCK_AMOUNT')) || 0
  };
}

function saveTargets(data) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('TARGET_PER_TRUCK_COUNT', String(Number(data.targetCount) || 0));
  props.setProperty('TARGET_PER_TRUCK_AMOUNT', String(Number(data.targetAmount) || 0));
  return { success: true };
}

// ==================== 特別日設定（連休対応） ====================

function getSpecialDays() {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.SPECIAL_DAYS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return data.map(function(row) {
    var dateVal = row[0] instanceof Date ? formatDate(row[0]) : String(row[0]);
    return { date: dateVal, trucks: row[1], type: row[2], memo: row[3] || '' };
  }).sort(function(a, b) { return new Date(a.date) - new Date(b.date); });
}

function saveSpecialDay(data) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.SPECIAL_DAYS);
  if (!sheet) throw new Error('特別日設定シートが見つかりません');

  var startDate = new Date(data.date);
  var endDate = data.endDate ? new Date(data.endDate) : startDate;

  for (var d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    var dateStr = formatDate(d);
    var existingRow = findSpecialDayRow(sheet, dateStr);
    var rowData = [dateStr, Number(data.trucks) || 0, data.type, data.memo || ''];

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, 4).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
  }

  return { success: true };
}

function deleteSpecialDay(date) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.SPECIAL_DAYS);
  if (!sheet || sheet.getLastRow() <= 1) return { success: true };
  var dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = dates.length - 1; i >= 0; i--) {
    var d = dates[i][0] instanceof Date ? formatDate(dates[i][0]) : String(dates[i][0]);
    if (d === date) { sheet.deleteRow(i + 2); break; }
  }
  return { success: true };
}

function findSpecialDayRow(sheet, dateStr) {
  if (sheet.getLastRow() <= 1) return -1;
  var dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0] instanceof Date ? formatDate(dates[i][0]) : String(dates[i][0]);
    if (d === dateStr) return i + 2;
  }
  return -1;
}

// ==================== 祝日・販促期間 ====================

function getHolidays() {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return data.map(function(row) {
    var startDate = row[0] instanceof Date ? formatDate(row[0]) : String(row[0]);
    var endDate = row[1] ? (row[1] instanceof Date ? formatDate(row[1]) : String(row[1])) : '';
    return { date: startDate, endDate: endDate, type: row[2], name: row[3], memo: row[4] || '' };
  }).sort(function(a, b) { return new Date(a.date) - new Date(b.date); });
}

function saveHoliday(data) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet) throw new Error('祝日・販促期間シートが見つかりません');
  sheet.appendRow([data.date, data.endDate || '', data.type, data.name, data.memo || '']);
  return { success: true };
}

function deleteHoliday(date, name) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet || sheet.getLastRow() <= 1) return { success: true };
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    var d = data[i][0] instanceof Date ? formatDate(data[i][0]) : String(data[i][0]);
    if (d === date && data[i][3] === name) { sheet.deleteRow(i + 2); break; }
  }
  return { success: true };
}

function importHolidays(rows) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet) throw new Error('祝日・販促期間シートが見つかりません');

  var imported = 0;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r.date || !r.type || !r.name) continue;
    sheet.appendRow([r.date, r.endDate || '', r.type, r.name, r.memo || '']);
    imported++;
  }
  return { success: true, count: imported };
}

// ==================== エリアマスタ ====================

function getAreas() {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.AREA_MASTER);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return data.map(function(row) {
    return { areaId: row[0], areaName: row[1], deliveryDays: row[2] || '', memo: row[3] || '' };
  });
}

function saveArea(data) {
  var ss = getSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.AREA_MASTER, AREA_MASTER_HEADERS);

  // 既存チェック
  if (sheet.getLastRow() > 1) {
    var existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]) === String(data.areaId)) {
        sheet.getRange(i + 2, 1, 1, 4).setValues([[data.areaId, data.areaName, data.deliveryDays || '', data.memo || '']]);
        return { success: true };
      }
    }
  }
  sheet.appendRow([data.areaId, data.areaName, data.deliveryDays || '', data.memo || '']);
  return { success: true };
}

function deleteArea(areaId) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.AREA_MASTER);
  if (!sheet || sheet.getLastRow() <= 1) return { success: true };
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]) === String(areaId)) { sheet.deleteRow(i + 2); break; }
  }
  return { success: true };
}

// ==================== エリア例外設定 ====================

function getAreaExceptions() {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.AREA_EXCEPTION);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  // 直近3ヶ月分のみ
  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - 3);
  return data.map(function(row) {
    var dateVal = row[0] instanceof Date ? formatDate(row[0]) : String(row[0]);
    return { date: dateVal, areaId: row[1], exceptionType: row[2], reason: row[3] || '' };
  }).filter(function(r) {
    return new Date(r.date) >= cutoff;
  }).sort(function(a, b) { return new Date(a.date) - new Date(b.date); });
}

function saveAreaException(data) {
  var ss = getSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.AREA_EXCEPTION, AREA_EXCEPTION_HEADERS);
  sheet.appendRow([data.date, data.areaId, data.exceptionType, data.reason || '']);
  return { success: true };
}

function deleteAreaException(date, areaId) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.AREA_EXCEPTION);
  if (!sheet || sheet.getLastRow() <= 1) return { success: true };
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    var d = data[i][0] instanceof Date ? formatDate(data[i][0]) : String(data[i][0]);
    if (d === date && String(data[i][1]) === String(areaId)) { sheet.deleteRow(i + 2); break; }
  }
  return { success: true };
}

// ==================== 契約履歴ログ取得 ====================

function getContractLogs(limit) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.CONTRACT_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, CONTRACT_LOG_HEADERS.length).getValues();
  var result = data.map(function(row) {
    return {
      contractDate: row[0] instanceof Date ? formatDate(row[0]) : String(row[0]),
      contractDay: row[1],
      contractPromo: row[2],
      deliveryDate: row[3] instanceof Date ? formatDate(row[3]) : String(row[3]),
      deliveryDay: row[4],
      deliveryPromo: row[5],
      leadTime: row[6],
      count: row[7],
      amount: row[8],
      nearestOpen: row[9] instanceof Date ? formatDate(row[9]) : String(row[9]),
      nearestDays: row[10],
      nearestDay: row[11]
    };
  });
  result.sort(function(a, b) { return new Date(b.contractDate) - new Date(a.contractDate); });
  if (limit && limit > 0) result = result.slice(0, limit);
  return result;
}

// ==================== 拡張分析データ取得 ====================

function getAnalysisData(startDate, endDate) {
  var rows = getRows(startDate, endDate);
  if (rows.length === 0) return { rows: [], holidays: [], contractLogs: [] };

  var holidays = getHolidays();
  var holidayDateSet = {};
  var promoDateSet = {};

  for (var h = 0; h < holidays.length; h++) {
    var hol = holidays[h];
    var start = new Date(hol.date);
    var end = hol.endDate ? new Date(hol.endDate) : start;
    for (var dd = new Date(start); dd <= end; dd.setDate(dd.getDate() + 1)) {
      var ds = formatDate(dd);
      if (hol.type === '祝日') {
        holidayDateSet[ds] = hol.name;
      } else if (hol.type === '販促期間') {
        promoDateSet[ds] = hol.name;
      }
    }
  }

  var enriched = rows.map(function(r) {
    var dayIndex = new Date(r.date).getDay();
    var isWeekend = (dayIndex === 0 || dayIndex === 6);
    var isHoliday = !!holidayDateSet[r.date];
    var isPromo = !!promoDateSet[r.date];
    var dayType = isHoliday ? '祝日' : isWeekend ? '土日' : '平日';

    return {
      date: r.date,
      dayName: r.dayName,
      dayType: dayType,
      isHoliday: isHoliday,
      isPromo: isPromo,
      holidayName: holidayDateSet[r.date] || '',
      promoName: promoDateSet[r.date] || '',
      trucks: Number(r.trucks) || 0,
      totalCount: Number(r.totalCount) || 0,
      totalAmount: Number(r.totalAmount) || 0,
      perTruckCount: Number(r.perTruckCount) || 0,
      perTruckAmount: Number(r.perTruckAmount) || 0,
      yesterdayCount: Number(r.yesterdayCount) || 0,
      yesterdayAmount: Number(r.yesterdayAmount) || 0,
      todayCount: Number(r.todayCount) || 0,
      todayAmount: Number(r.todayAmount) || 0,
      manshaDate: r.manshaDate || '',
      manshaCount: Number(r.manshaCount) || 0,
      status: r.status,
      dayOfWeek: dayIndex,
      weekendFlag: isWeekend ? 1 : 0,
      holidayFlag: isHoliday ? 1 : 0,
      promoFlag: isPromo ? 1 : 0
    };
  });

  var contractLogs = getContractLogs();

  return { rows: enriched, holidays: holidays, contractLogs: contractLogs };
}

// ==================== 予測用データ取得 ====================

function getPredictionData(deliveryDate) {
  var ss = getSpreadsheet();
  var d = new Date(deliveryDate);
  var dayIndex = d.getDay();
  var dayName = DAY_NAMES[dayIndex];
  var daySettings = getDaySettingsMap(ss);
  var specialDaysMap = getSpecialDaysMap(ss);
  var trucks = resolveTrucks(deliveryDate, dayIndex, daySettings, specialDaysMap);

  // 販促判定
  var holidays = getHolidays();
  var promoDateSet = {};
  for (var h = 0; h < holidays.length; h++) {
    var hol = holidays[h];
    if (hol.type !== '販促期間') continue;
    var hStart = new Date(hol.date);
    var hEnd = hol.endDate ? new Date(hol.endDate) : hStart;
    for (var dd = new Date(hStart); dd <= hEnd; dd.setDate(dd.getDate() + 1)) {
      promoDateSet[formatDate(dd)] = true;
    }
  }
  var isPromo = !!promoDateSet[deliveryDate];
  var isWeekend = (dayIndex === 0 || dayIndex === 6);

  // 過去の同条件データ
  var allRows = getRows(null, null);
  var similar = [];
  for (var i = 0; i < allRows.length; i++) {
    var r = allRows[i];
    if (r.date >= deliveryDate) continue;
    var rDate = new Date(r.date);
    var rIsWeekend = (rDate.getDay() === 0 || rDate.getDay() === 6);
    var rIsPromo = !!promoDateSet[r.date];
    // 同条件: 同曜日区分 & 台数が±1以内
    if (rIsWeekend === isWeekend && rIsPromo === isPromo &&
        Math.abs(Number(r.trucks) - trucks) <= 1 &&
        Number(r.totalCount) > 0) {
      similar.push(r);
    }
  }

  // 満車タイミング（同条件で満車確定日がある日のリードタイム）
  var manshaLeadTimes = [];
  for (var j = 0; j < similar.length; j++) {
    if (similar[j].manshaDate) {
      var mDate = new Date(similar[j].manshaDate);
      var dDate = new Date(similar[j].date);
      var lt = Math.round((dDate - mDate) / (1000 * 60 * 60 * 24));
      if (lt >= 0) manshaLeadTimes.push(lt);
    }
  }

  return {
    dayName: dayName,
    dayIndex: dayIndex,
    trucks: trucks,
    isPromo: isPromo,
    isWeekend: isWeekend,
    similarRows: similar.sort(function(a, b) { return new Date(b.date) - new Date(a.date); }).slice(0, 30),
    manshaLeadTimes: manshaLeadTimes
  };
}

// ==================== CSVエクスポート ====================

function exportCSV(startDate, endDate) {
  var rows = getRows(startDate, endDate);
  if (rows.length === 0) return '';

  var headers = ['配送日', '曜日', '台数', 'ステータス', '合計件数', '合計金額(千円)',
    '1台あたり件数', '1台あたり金額(千円)', '昨日までの合計件数', '昨日までの合計金額(千円)',
    '当日契約件数', '当日契約金額(千円)', '備考', '満車確定日', '満車時件数'];

  var csv = headers.join(',') + '\n';
  rows.forEach(function(r) {
    csv += [
      r.date, r.dayName, r.trucks, r.status,
      r.totalCount, r.totalAmount,
      r.perTruckCount, r.perTruckAmount,
      r.yesterdayCount, r.yesterdayAmount,
      r.todayCount, r.todayAmount,
      '"' + (r.memo || '').replace(/"/g, '""') + '"',
      r.manshaDate || '', r.manshaCount || ''
    ].join(',') + '\n';
  });
  return csv;
}

// ==================== 変更ログ ====================

function getChangeLog(limit) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.CHANGE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var result = data.map(function(row) {
    return {
      timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(row[0]),
      user: row[1], targetDate: row[2], oldValue: row[3], newValue: row[4], field: row[5]
    };
  });
  result.reverse();
  if (limit && limit > 0) result = result.slice(0, limit);
  return result;
}

function logChange(targetDate, field, oldValue, newValue) {
  try {
    var sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.CHANGE_LOG);
    if (!sheet) return;
    var user = Session.getActiveUser().getEmail() || '不明';
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([timestamp, user, targetDate, oldValue, newValue, field]);
  } catch (e) {}
}

// ==================== ユーティリティ ====================

function findRowByDate(sheet, date) {
  if (sheet.getLastRow() <= 1) return -1;
  var dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0];
    var formatted = d instanceof Date ? formatDate(d) : String(d);
    if (formatted === date) return i + 2;
  }
  return -1;
}

function deduplicateMainSheet(sheet) {
  if (sheet.getLastRow() <= 1) return;
  var numCols = Math.max(sheet.getLastColumn(), MAIN_HEADERS.length);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  var dateMap = {};

  for (var i = 0; i < data.length; i++) {
    var dateVal = data[i][0];
    var key = dateVal instanceof Date ? formatDate(dateVal) : String(dateVal);
    if (!key || key === 'undefined' || key === '') continue;

    if (!dateMap[key]) {
      dateMap[key] = { bestIdx: i, deleteIndices: [] };
    } else {
      var existingRow = data[dateMap[key].bestIdx];
      var currentRow = data[i];
      var existingScore = rowDataScore(existingRow);
      var currentScore = rowDataScore(currentRow);

      if (currentScore > existingScore) {
        dateMap[key].deleteIndices.push(dateMap[key].bestIdx);
        dateMap[key].bestIdx = i;
      } else {
        dateMap[key].deleteIndices.push(i);
      }
    }
  }

  var rowsToDelete = [];
  for (var k in dateMap) {
    for (var j = 0; j < dateMap[k].deleteIndices.length; j++) {
      rowsToDelete.push(dateMap[k].deleteIndices[j] + 2);
    }
  }

  if (rowsToDelete.length === 0) return;

  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var d = 0; d < rowsToDelete.length; d++) {
    sheet.deleteRow(rowsToDelete[d]);
  }
}

function rowDataScore(row) {
  return (Number(row[2]) || 0) + (Number(row[4]) || 0) + (Number(row[5]) || 0) +
         (Number(row[8]) || 0) + (Number(row[9]) || 0) + (Number(row[10]) || 0) + (Number(row[11]) || 0);
}

function formatDate(d) {
  var year = d.getFullYear();
  var month = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  return year + '-' + month + '-' + day;
}
