// ================================
// 設定ファイル
// ================================

// スプレッドシートIDを取得
function getSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

// 営業日リストスプレッドシートIDを取得
function getBusinessDaysSheetId() {
  return PropertiesService.getScriptProperties().getProperty('BUSINESS_DAYS_SHEET_ID');
}

// カレンダーIDを取得
function getCalendarId() {
  temp = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');
  console.log(temp)
  return PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');
}

// スプレッドシートオブジェクトを取得
function getSpreadsheet() {
  return SpreadsheetApp.openById(getSpreadsheetId());
}

// 各シートを取得
function getReservationsSheet() {
  return getSpreadsheet().getSheetByName('Reservations');
}

function getVisitDatesSheet() {
  return getSpreadsheet().getSheetByName('VisitDates');
}

function getConfigSheet() {
  return getSpreadsheet().getSheetByName('Config');
}

// Config情報を取得
function getConfigValue(key) {
  const sheet = getConfigSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}

// 職員通知メールアドレスを取得
function getStaffEmail() {
  return getConfigValue('職員通知メールアドレス');
}

// テスト関数
function testConfig() {
  Logger.log('スプレッドシートID: ' + getSpreadsheetId());
  Logger.log('営業日リストID: ' + getBusinessDaysSheetId());
  Logger.log('カレンダーID: ' + getCalendarId());
  Logger.log('職員メール: ' + getStaffEmail());
}

// フォームIDを取得
function getFormId() {
  return getConfigValue('フォームID');
}

// 共有フォルダIDを取得
function getSharedFolderId() {
  return getConfigValue('共有フォルダID');
}