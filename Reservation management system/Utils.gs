// ================================
// 学外者入館管理システム - ユーティリティ関数
// ================================

/**
 * 日付を yyyy/MM/dd 形式にフォーマット
 * @param {Date|string} date - 日付オブジェクトまたは文字列
 * @return {string} フォーマットされた日付
 */
function formatDate(date) {
  if (!date) return '';
  
  const d = new Date(date);
  if (isNaN(d.getTime())) return String(date);
  
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  
  return `${year}/${month}/${day}`;
}

/**
 * 日時を yyyy/MM/dd HH:mm:ss 形式にフォーマット
 * @param {Date|string} date - 日付オブジェクトまたは文字列
 * @return {string} フォーマットされた日時
 */
function formatDateTime(date) {
  if (!date) return '';
  
  const d = new Date(date);
  if (isNaN(d.getTime())) return String(date);
  
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const hours = String(d.getHours()).padStart(2, '0');
  const minutes = String(d.getMinutes()).padStart(2, '0');
  const seconds = String(d.getSeconds()).padStart(2, '0');
  
  return `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
}

/**
 * 時刻を HH:mm 形式にフォーマット
 * @param {Date|string} date - 日付オブジェクトまたは文字列
 * @return {string} フォーマットされた時刻
 */
function formatTime(date) {
  if (!date) return '';
  
  const d = new Date(date);
  if (isNaN(d.getTime())) return String(date);
  
  const hours = String(d.getHours()).padStart(2, '0');
  const minutes = String(d.getMinutes()).padStart(2, '0');
  
  return `${hours}:${minutes}`;
}

/**
 * 今日の日付を取得（時刻を0:00:00にリセット）
 * @return {Date} 今日の日付
 */
function getToday() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return today;
}

/**
 * 現在日時を取得
 * @return {Date} 現在日時
 */
function getNow() {
  return new Date();
}

/**
 * 日付が今日かどうかをチェック
 * @param {Date|string} date - チェックする日付
 * @return {boolean} 今日ならtrue
 */
function isToday(date) {
  const targetDate = new Date(date);
  targetDate.setHours(0, 0, 0, 0);
  
  const today = getToday();
  
  return targetDate.getTime() === today.getTime();
}

/**
 * レコードIDを生成（年度を含む）
 * @return {string} レコードID（例: 2026-0218-001）
 */
function generateRecordId() {
  const now = getNow();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const random = String(Math.floor(Math.random() * 1000)).padStart(3, '0');
  
  return `${year}-${month}${day}-${random}`;
}

/**
 * HTMLエスケープ処理
 * @param {string} text - エスケープする文字列
 * @return {string} エスケープされた文字列
 */
function escapeHtml(text) {
  if (!text) return '';
  
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  
  return String(text).replace(/[&<>"']/g, m => map[m]);
}

/**
 * CSVエスケープ処理
 * @param {string} text - エスケープする文字列
 * @return {string} エスケープされた文字列
 */
function escapeCsv(text) {
  if (!text) return '';
  
  const str = String(text);
  
  // カンマ、改行、ダブルクォートが含まれる場合はダブルクォートで囲む
  if (str.includes(',') || str.includes('\n') || str.includes('"')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  
  return str;
}

/**
 * 配列をCSV行に変換
 * @param {Array} array - データ配列
 * @return {string} CSV形式の文字列
 */
function arrayToCsvRow(array) {
  return array.map(cell => escapeCsv(cell)).join(',');
}

/**
 * 2次元配列をCSVテキストに変換
 * @param {Array<Array>} data - 2次元配列
 * @return {string} CSV形式のテキスト
 */
function arrayToCsv(data) {
  return data.map(row => arrayToCsvRow(row)).join('\n');
}

/**
 * 日付文字列を比較（yyyy/MM/dd形式）
 * @param {string} date1 - 日付1
 * @param {string} date2 - 日付2
 * @return {number} date1 < date2: -1, date1 = date2: 0, date1 > date2: 1
 */
function compareDates(date1, date2) {
  const d1 = new Date(date1);
  const d2 = new Date(date2);
  
  if (d1 < d2) return -1;
  if (d1 > d2) return 1;
  return 0;
}

/**
 * 日付範囲のチェック
 * @param {Date|string} targetDate - チェック対象の日付
 * @param {Date|string} startDate - 開始日
 * @param {Date|string} endDate - 終了日
 * @return {boolean} 範囲内ならtrue
 */
function isDateInRange(targetDate, startDate, endDate) {
  const target = new Date(targetDate);
  const start = new Date(startDate);
  const end = new Date(endDate);
  
  target.setHours(0, 0, 0, 0);
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);
  
  return target >= start && target <= end;
}

/**
 * 空文字列・null・undefinedをチェック
 * @param {*} value - チェックする値
 * @return {boolean} 空ならtrue
 */
function isEmpty(value) {
  return value === null || value === undefined || value === '';
}

/**
 * 配列の重複を除去
 * @param {Array} array - 配列
 * @return {Array} 重複除去後の配列
 */
function uniqueArray(array) {
  return [...new Set(array)];
}

/**
 * オブジェクトの深いコピー
 * @param {Object} obj - コピー元オブジェクト
 * @return {Object} コピーされたオブジェクト
 */
function deepCopy(obj) {
  return JSON.parse(JSON.stringify(obj));
}

// ================================
// テスト関数
// ================================

/**
 * 日付フォーマットのテスト
 */
function testDateFormat() {
  Logger.log('=== 日付フォーマットテスト ===');
  
  const now = new Date();
  
  Logger.log('formatDate: ' + formatDate(now));
  Logger.log('formatDateTime: ' + formatDateTime(now));
  Logger.log('formatTime: ' + formatTime(now));
  Logger.log('getToday: ' + formatDate(getToday()));
  Logger.log('isToday(now): ' + isToday(now));
  Logger.log('isToday("2026/01/01"): ' + isToday('2026/01/01'));
  
  Logger.log('\n✓ 日付フォーマットテスト完了');
}

/**
 * ID生成のテスト
 */
function testIdGeneration() {
  Logger.log('=== ID生成テスト ===');
  
  for (let i = 0; i < 5; i++) {
    Logger.log('レコードID: ' + generateRecordId());
  }
  
  Logger.log('\n✓ ID生成テスト完了');
}

/**
 * エスケープ処理のテスト
 */
function testEscape() {
  Logger.log('=== エスケープ処理テスト ===');
  
  const htmlTest = '<script>alert("test")</script>';
  Logger.log('HTML: ' + escapeHtml(htmlTest));
  
  const csvTest = 'test, "quoted", new\nline';
  Logger.log('CSV: ' + escapeCsv(csvTest));
  
  const arrayTest = [
    ['名前', '値1', '値2'],
    ['山田太郎', 'test, comma', 'test\nnewline']
  ];
  Logger.log('CSV配列:\n' + arrayToCsv(arrayTest));
  
  Logger.log('\n✓ エスケープ処理テスト完了');
}

/**
 * ユーティリティ関数の総合テスト
 */
function testUtils() {
  Logger.log('=== ユーティリティ総合テスト ===\n');
  
  testDateFormat();
  Logger.log('');
  testIdGeneration();
  Logger.log('');
  testEscape();
  
  Logger.log('\n=== すべてのテスト完了 ===');
}