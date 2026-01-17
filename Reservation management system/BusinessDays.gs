// ================================
// 営業日管理
// ================================

/**
 * 営業日リストを取得
 * @return {Array} [[日付, 曜日, 開館日時, 備考], ...]
 */
function getBusinessDaysList() {
  const sheet = SpreadsheetApp.openById(getBusinessDaysSheetId())
                               .getSheetByName('Calendar');
  const data = sheet.getDataRange().getValues();
  
  // ヘッダー行を除いたデータを返す
  return data.slice(1);
}

/**
 * 指定した日付が営業日かどうかを判定
 * @param {string|Date} dateStr - 判定する日付 (例: "2025/11/15" または Date オブジェクト)
 * @return {boolean} 営業日ならtrue、休館日ならfalse
 */
function isBusinessDay(dateStr) {
  const targetDate = formatDate(dateStr);
  const businessDays = getBusinessDaysList();
  
  for (let i = 0; i < businessDays.length; i++) {
    const businessDate = formatDate(businessDays[i][0]); // A列: 日付
    const openingHours = String(businessDays[i][2]); // C列: 開館日時
    
    if (businessDate === targetDate) {
      // C列に数字が含まれていれば営業日
      return /\d/.test(openingHours);
    }
  }
  
  return false; // リストにない日付は非営業日
}

/**
 * 今日から指定日数分の営業日を取得
 * @param {number} days - 取得する日数 (デフォルト: 60日)
 * @return {Array} [[日付文字列, 開館時間], ...] の配列
 */
function getAvailableBusinessDays(days = 60) {
  const today = getToday();
  const endDate = new Date(today);
  endDate.setDate(endDate.getDate() + days);
  
  const businessDays = getBusinessDaysList();
  const availableDays = [];
  
  for (let i = 0; i < businessDays.length; i++) {
    const date = new Date(businessDays[i][0]); // A列: 日付
    const openingHours = String(businessDays[i][2]); // C列: 開館日時
    
    // 今日以降、指定日数以内、かつ営業日（数字を含む）の場合
    if (date >= today && date <= endDate && /\d/.test(openingHours)) {
      availableDays.push([
        formatDate(date),
        openingHours
      ]);
    }
  }
  
  return availableDays;
}

/**
 * 営業日リストをフォーム用の選択肢形式で取得
 * @return {Array} ["2025/11/15 (9:00-20:00)", ...] の配列
 */
function getBusinessDaysForForm() {
  try {
    const businessDays = getAvailableBusinessDays(365); // 1年分取得
    const tomorrow = new Date(getToday());
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = formatDate(tomorrow);
    
    Logger.log('tomorrow文字列: ' + tomorrowStr);
    
    // 日付文字列で比較
    const filteredDays = businessDays.filter(day => day[0] >= tomorrowStr);
    Logger.log('フィルタ後: ' + filteredDays.length + '件');
    
    return filteredDays.map(day => day[0]); // 日付文字列のみ返す
  } catch (error) {
    Logger.log('営業日取得エラー: ' + error.message);
    throw error;
  }
}

// ================================
// テスト関数
// ================================

/**
 * 営業日機能のテスト
 */
function testBusinessDays() {
  Logger.log('=== 営業日リスト取得テスト ===');
  const list = getBusinessDaysList();
  Logger.log(`取得件数: ${list.length}件`);
  Logger.log('最初の3件:');
  for (let i = 0; i < Math.min(3, list.length); i++) {
    Logger.log(`  ${formatDate(list[i][0])} | ${list[i][1]} | ${list[i][2]} | ${list[i][3]}`);
  }
  
  Logger.log('\n=== 営業日判定テスト ===');
  // 営業日リストの最初の日付でテスト
  if (list.length > 0) {
    const testDate = formatDate(list[0][0]);
    Logger.log(`${testDate}は営業日?: ${isBusinessDay(testDate)}`);
  }
  
  Logger.log('\n=== 利用可能な営業日取得テスト ===');
  const available = getAvailableBusinessDays(60);
  Logger.log(`今日から60日以内の営業日: ${available.length}件`);
  Logger.log('最初の5件:');
  for (let i = 0; i < Math.min(5, available.length); i++) {
    Logger.log(`  ${available[i][0]} (${available[i][1]})`);
  }
  
  Logger.log('\n=== フォーム用選択肢テスト ===');
  const formOptions = getBusinessDaysForForm();
  Logger.log(`フォーム選択肢: ${formOptions.length}件`);
  Logger.log('最初の3件:');
  for (let i = 0; i < Math.min(3, formOptions.length); i++) {
    Logger.log(`  ${formOptions[i]}`);
  }
}
