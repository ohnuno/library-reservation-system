// ================================
// 学外者入館管理システム - 設定ファイル
// ================================

/**
 * 重要: このシステムは「学外者用訪問予約受付システム」と同じスプレッドシートを参照します
 * スプレッドシートIDは、予約受付システムと同じものを使用してください
 */

/**
 * スプレッドシートIDを取得
 * @return {string} スプレッドシートID
 */
function getSpreadsheetId() {
  // スクリプトプロパティから取得
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  
  if (!id) {
    throw new Error('SPREADSHEET_IDが設定されていません。\n' +
                    'スクリプトプロパティで設定してください。\n' +
                    '値: 予約受付システムと同じスプレッドシートID');
  }
  
  return id;
}

/**
 * カレンダーIDを取得
 * @return {string} カレンダーID
 */
function getCalendarId() {
  const id = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');
  
  if (!id) {
    Logger.log('警告: CALENDAR_IDが設定されていません');
  }
  
  return id;
}

/**
 * スプレッドシートオブジェクトを取得
 * @return {Spreadsheet} スプレッドシート
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(getSpreadsheetId());
}

/**
 * Configシートを取得
 * @return {Sheet} Configシート
 */
function getConfigSheet() {
  return getSpreadsheet().getSheetByName('Config');
}

/**
 * Reservationsシートを取得
 * @return {Sheet} Reservationsシート
 */
function getReservationsSheet() {
  return getSpreadsheet().getSheetByName('Reservations');
}

/**
 * VisitDatesシートを取得
 * @return {Sheet} VisitDatesシート
 */
function getVisitDatesSheet() {
  return getSpreadsheet().getSheetByName('VisitDates');
}

/**
 * Calendarシートを取得
 * @return {Sheet} Calendarシート
 */
function getCalendarSheet() {
  return getSpreadsheet().getSheetByName('Calendar');
}

/**
 * Config情報を取得
 * @param {string} key - 設定キー
 * @return {string} 設定値
 */
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

/**
 * Config情報を更新
 * @param {string} key - 設定キー
 * @param {string} value - 設定値
 */
function updateConfig(key, value) {
  const sheet = getConfigSheet();
  const data = sheet.getDataRange().getValues();
  
  // 既存の行を探す
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      Logger.log(`Config更新: ${key} = ${value}`);
      return;
    }
  }
  
  // 存在しない場合は新規追加
  sheet.appendRow([key, value]);
  Logger.log(`Config追加: ${key} = ${value}`);
}

/**
 * 職員通知メールアドレスを取得
 * @return {string} メールアドレス
 */
function getStaffEmail() {
  return getConfigValue('職員通知メールアドレス');
}

/**
 * 現在の年度を取得
 * @return {number} 年度
 */
function getCurrentYear() {
  const year = getConfigValue('CURRENT_YEAR');
  return year ? parseInt(year) : new Date().getFullYear();
}

/**
 * 現在の年度を取得（自動計算版）
 * 年度の定義: 4月1日〜翌年3月31日
 * @return {number} 年度
 */
function getCurrentYear() {
  // まずConfigシートの設定を確認
  const configYear = getConfigValue('CURRENT_YEAR');
  
  // 設定がある場合はそれを優先（手動設定）
  if (configYear) {
    return parseInt(configYear);
  }
  
  // 設定がない場合は自動計算
  return calculateCurrentYear();
}

/**
 * 現在の年度を自動計算
 * 年度の定義: 4月1日〜翌年3月31日
 * @return {number} 年度
 */
function calculateCurrentYear() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1; // 0-11 → 1-12
  
  // 4月1日より前の場合は前年度
  if (month < 4) {
    return year - 1;
  }
  
  // 4月1日以降は現在年が年度
  return year;
}

/**
 * 指定した日付が属する年度を計算
 * @param {Date|string} date - 日付
 * @return {number} 年度
 */
function getYearForDate(date) {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = d.getMonth() + 1;
  
  if (month < 4) {
    return year - 1;
  }
  
  return year;
}

/**
 * 年度の開始日を取得
 * @param {number} year - 年度
 * @return {Date} 年度開始日（4月1日）
 */
function getYearStartDate(year) {
  return new Date(year, 3, 1); // 月は0始まりなので3 = 4月
}

/**
 * 年度の終了日を取得
 * @param {number} year - 年度
 * @return {Date} 年度終了日（翌年3月31日）
 */
function getYearEndDate(year) {
  return new Date(year + 1, 2, 31); // 2 = 3月
}

// ================================
// テスト関数
// ================================

/**
 * 設定のテスト
 */
function testConfig() {
  Logger.log('=== 設定確認テスト ===');
  
  try {
    Logger.log('スプレッドシートID: ' + getSpreadsheetId());
    Logger.log('カレンダーID: ' + (getCalendarId() || '(未設定)'));
    Logger.log('職員メールアドレス: ' + (getStaffEmail() || '(未設定)'));
    Logger.log('現在の年度: ' + getCurrentYear());
    
    Logger.log('\n✓ 設定確認完了');
    
  } catch (error) {
    Logger.log('\n✗ 設定確認失敗');
    Logger.log('エラー: ' + error.message);
  }
}

/**
 * シート接続テスト
 */
function testSheetAccess() {
  Logger.log('=== シート接続テスト ===');
  
  try {
    const reservations = getReservationsSheet();
    Logger.log('Reservationsシート: ' + reservations.getName() + ' (' + reservations.getLastRow() + '行)');
    
    const visitDates = getVisitDatesSheet();
    Logger.log('VisitDatesシート: ' + visitDates.getName() + ' (' + visitDates.getLastRow() + '行)');
    
    const config = getConfigSheet();
    Logger.log('Configシート: ' + config.getName() + ' (' + config.getLastRow() + '行)');
    
    Logger.log('\n✓ シート接続成功');
    
  } catch (error) {
    Logger.log('\n✗ シート接続失敗');
    Logger.log('エラー: ' + error.message);
  }
}

/**
 * 年度計算のテスト
 */
function testYearCalculation() {
  Logger.log('=== 年度計算テスト ===');
  
  const now = new Date();
  Logger.log('現在日時: ' + formatDateTime(now));
  Logger.log('カレンダー年: ' + now.getFullYear());
  Logger.log('自動計算年度: ' + calculateCurrentYear());
  Logger.log('getCurrentYear(): ' + getCurrentYear());
  
  Logger.log('\n年度の範囲:');
  const currentYear = calculateCurrentYear();
  Logger.log(`${currentYear}年度: ${formatDate(getYearStartDate(currentYear))} 〜 ${formatDate(getYearEndDate(currentYear))}`);
  
  Logger.log('\n日付ごとの年度判定:');
  const testDates = [
    '2025/03/31',
    '2025/04/01',
    '2026/03/31',
    '2026/04/01'
  ];
  
  testDates.forEach(dateStr => {
    const year = getYearForDate(dateStr);
    Logger.log(`  ${dateStr} → ${year}年度`);
  });
  
  Logger.log('\n✓ テスト完了');
}