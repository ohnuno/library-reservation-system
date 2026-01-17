// ================================
// ユーティリティ関数
// ================================

// 予約IDを生成
function generateReservationId() {
  const date = new Date();
  const dateStr = Utilities.formatDate(date, 'JST', 'yyyyMMdd');
  const randomStr = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  return `RSV${dateStr}${randomStr}`;
}

// 予約トークンを生成 (32文字のランダム文字列)
function generateToken() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let token = '';
  for (let i = 0; i < 32; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

// 日付を yyyy/MM/dd 形式にフォーマット
function formatDate(date) {
  return Utilities.formatDate(new Date(date), 'JST', 'yyyy/MM/dd');
}

// 日付時刻を yyyy/MM/dd HH:mm:ss 形式にフォーマット
function formatDateTime(date) {
  return Utilities.formatDate(new Date(date), 'JST', 'yyyy/MM/dd HH:mm:ss');
}

// 今日の日付を取得 (時刻を0:00:00にリセット)
function getToday() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return today;
}

// 日付が今日かどうかをチェック
function isToday(dateStr) {
  const targetDate = new Date(dateStr);
  targetDate.setHours(0, 0, 0, 0);
  
  const today = getToday();
  
  return targetDate.getTime() === today.getTime();
}

// テスト関数
function testUtils() {
  Logger.log('予約ID: ' + generateReservationId());
  Logger.log('トークン: ' + generateToken());
  Logger.log('今日: ' + formatDate(getToday()));
  Logger.log('2025/11/15は今日?: ' + isToday('2025/11/15'));
}