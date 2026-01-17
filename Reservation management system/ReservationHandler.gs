// ================================
// 予約処理メインハンドラー
// ================================

/**
 * フォーム送信時のトリガー関数
 * @param {Object} e - フォーム送信イベントオブジェクト
 */
function onFormSubmit(e) {
  try {
    Logger.log('=== 予約処理開始 ===');
    
    // 1. フォーム回答を取得
    const formResponse = e.response;
    
    // 2. 回答データを抽出
    const reservationData = extractReservationData(formResponse);
    Logger.log('予約データ抽出完了: ' + JSON.stringify(reservationData));
    
    // 3. 予約ID・トークン生成
    reservationData.reservationId = generateReservationId();
    reservationData.token = generateToken();
    reservationData.status = '有効';
    reservationData.createdAt = new Date();
    reservationData.updatedAt = new Date();
    
    Logger.log('予約ID: ' + reservationData.reservationId);
    Logger.log('トークン: ' + reservationData.token);
    
    // 4. 営業日チェック
    const invalidDates = checkBusinessDays(reservationData.visitDates);
    if (invalidDates.length > 0) {
      Logger.log('警告: 以下の日付は営業日ではありません: ' + invalidDates.join(', '));
      // 営業日でない日付は除外
      reservationData.visitDates = reservationData.visitDates.filter(date => !invalidDates.includes(date));
    }
    
    if (reservationData.visitDates.length === 0) {
      throw new Error('有効な訪問日がありません');
    }
    
    Logger.log('有効な訪問日: ' + reservationData.visitDates.join(', '));
    
    // 5. Reservationsシートに保存
    const rowIndex = saveToReservationsSheet(reservationData);
    Logger.log('Reservationsシートに保存: 行' + rowIndex);
    
    // 6. VisitDatesシートに保存
    saveToVisitDatesSheet(reservationData);
    Logger.log('VisitDatesシートに保存完了');
    
    // 7. QRコード生成
    Logger.log('\nQRコード生成中...');
    const qrCodeResult = generateAndSaveQRCode(reservationData.reservationId, reservationData.token);
    reservationData.qrCodeEmailUrl = qrCodeResult.emailImageUrl;      // メール用URL(外部アクセス可能)
    reservationData.qrCodeBackupUrl = qrCodeResult.backupUrl;         // バックアップURL
    reservationData.qrCodeFileId = qrCodeResult.backupFileId;         // バックアップファイルID
    reservationData.qrContent = qrCodeResult.qrContent;               // QRコードの内容
    
    // バックアップURLをReservationsシートに保存
    if (qrCodeResult.backupUrl) {
      updateQRCodeInReservation(reservationData.reservationId, qrCodeResult.backupUrl);
    }
    Logger.log('QRコード生成完了');  

    // 8. カレンダーに登録
    Logger.log('\nカレンダーに登録中...');
    const eventIds = addReservationToCalendar(reservationData);
    reservationData.calendarEventIds = eventIds;
    
    // カレンダーイベントIDをReservationsシートに保存
    updateCalendarEventIdsInReservation(reservationData.reservationId, eventIds);
    
    // カレンダーイベントIDをVisitDatesシートに保存
    updateCalendarEventIdsInVisitDates(reservationData.reservationId, reservationData.visitDates, eventIds);
    
    Logger.log('カレンダー登録完了');
    
    // 9. ユーザーに予約完了メール送信
    Logger.log('\n予約完了メールを送信中...');
    sendConfirmationEmail(reservationData);
    Logger.log('予約完了メール送信完了');
    
    // 10. 職員に通知メール送信
    Logger.log('\n職員通知メールを送信中...');
    sendStaffNotification(reservationData);
    Logger.log('職員通知メール送信完了');
    
    Logger.log('=== 予約処理完了 ===');
    
  } catch (error) {
    Logger.log('エラー: ' + error.message);
    Logger.log(error.stack);
    
    // エラー通知をシステム管理者に送信
    sendErrorNotification(error);
  }
}

/**
 * フォーム回答から予約データを抽出
 * @param {FormResponse} formResponse - フォーム回答オブジェクト
 * @return {Object} 予約データ
 */
function extractReservationData(formResponse) {
  const data = {
    submittedAt: formResponse.getTimestamp(),
    email: getAnswerByItemId('email', formResponse)
  };
  
  // FormQuestionsシートから全項目を動的に取得
  const questions = getFormQuestions();
  
  questions.forEach(q => {
    if (q.itemId === 'email') {
      // 既に取得済みのためスキップ
      return;
    }
    
    if (q.itemId === 'visit_dates') {
      // 訪問日は特別処理（フォーマット変換）
      data.visitDatesRaw = getAnswerByItemId('visit_dates', formResponse);
      data.visitDates = parseVisitDates(data.visitDatesRaw);
    } else {
      // その他の項目は動的に取得
      const value = getAnswerByItemId(q.itemId, formResponse);
      data[q.itemId] = value || '';
    }
  });
  
  return data;
}

/**
 * 訪問希望日の文字列を解析
 * @param {Array} visitDatesRaw - フォームから取得した訪問希望日の配列
 * @return {Array} 日付文字列の配列 ["2025/11/15", "2025/11/20", ...]
 */
function parseVisitDates(visitDatesRaw) {
  if (!visitDatesRaw || !Array.isArray(visitDatesRaw)) {
    return [];
  }
  
  return visitDatesRaw.map(dateStr => {
    // "2025/11/15 (9:00-20:00)" から "2025/11/15" を抽出
    const match = dateStr.match(/^(\d{4}\/\d{2}\/\d{2})/);
    return match ? match[1] : dateStr;
  });
}

/**
 * 営業日チェック
 * @param {Array} dates - チェックする日付の配列
 * @return {Array} 営業日でない日付の配列
 */
function checkBusinessDays(dates) {
  const invalidDates = [];
  
  for (let date of dates) {
    if (!isBusinessDay(date)) {
      invalidDates.push(date);
    }
  }
  
  return invalidDates;
}

/**
 * Reservationsシートに保存
 * @param {Object} data - 予約データ
 * @return {number} 保存した行番号
 */
function saveToReservationsSheet(data) {
  const sheet = getReservationsSheet();
  
  // OPAC URLsを結合（最大5件）
  const opacUrls = [];
  for (let i = 1; i <= 5; i++) {
    const url = data[`opac_url_${i}`];
    if (url) opacUrls.push(url);
  }
  
  const row = [
    data.reservationId,                                      // A: 予約ID
    formatDateTime(data.submittedAt),                        // B: 申請日時
    data.name,                                               // C: 氏名
    data.email,                                              // D: メールアドレス
    data.phone,                                              // E: 電話番号
    data.address,                                            // F: 所属機関
    data.purpose,                                            // G: 訪問目的
    data.visitDates.map(d => formatDate(d)).join(','),      // H: 訪問日リスト
    data.status,                                             // I: ステータス
    '',                                                      // J: QRコードURL
    data.token,                                              // K: 予約トークン
    formatDateTime(data.createdAt),                          // L: 作成日時
    formatDateTime(data.updatedAt),                          // M: 更新日時
    '',                                                      // N: カレンダーイベントID
    data.companion_count || '',                              // O: 同行人数
    opacUrls.join(','),                                      // P: OPAC URL
    data.database_name || ''                                 // Q: データベース名称
  ];
  
  sheet.appendRow(row);
  
  return sheet.getLastRow();
}

/**
 * VisitDatesシートに保存
 * @param {Object} data - 予約データ
 */
function saveToVisitDatesSheet(data) {
  const sheet = getVisitDatesSheet();
  
  data.visitDates.forEach(date => {
    const row = [
      data.reservationId,     // A: 予約ID
      date,                   // B: 訪問日
      data.status,            // C: ステータス
      '',                     // D: カレンダーイベントID (後で設定)
      'FALSE'                 // E: QRコード有効フラグ (後で更新)
    ];
    
    sheet.appendRow(row);
  });
}

/**
 * 予約IDで予約を取得
 * @param {string} reservationId - 予約ID
 * @return {Object|null} 予約データ
 */
function getReservationById(reservationId) {
  try {
    const sheet = getReservationsSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reservationId) {
        let visitDatesStr = '';
        if (data[i][7] instanceof Date) {
          visitDatesStr = formatDate(data[i][7]);
        } else {
          visitDatesStr = String(data[i][7] || '');
        }
        
        return {
          reservationId: String(data[i][0]),
          submittedAt: data[i][1],
          name: String(data[i][2]),
          email: String(data[i][3]),
          phone: String(data[i][4]),
          address: String(data[i][5]),
          purpose: String(data[i][6]),
          visitDates: visitDatesStr,
          status: String(data[i][8]),
          qrCodeUrl: String(data[i][9]),
          token: String(data[i][10])
        };
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('予約取得エラー: ' + error.message);
    return null;
  }
}


/**
 * エラー通知をシステム管理者に送信
 * @param {Error} error - エラーオブジェクト
 */
function sendErrorNotification(error) {
  const adminEmail = getConfigValue('システム管理者メール');
  
  if (!adminEmail) {
    Logger.log('システム管理者メールが設定されていません');
    return;
  }
  
  const subject = '【エラー】訪問予約受付システム';
  const body = `予約処理中にエラーが発生しました。\n\n` +
               `エラーメッセージ: ${error.message}\n\n` +
               `スタックトレース:\n${error.stack}\n\n` +
               `発生日時: ${formatDateTime(new Date())}`;
  
  try {
    GmailApp.sendEmail(adminEmail, subject, body);
    Logger.log('エラー通知を送信しました: ' + adminEmail);
  } catch (e) {
    Logger.log('エラー通知の送信に失敗: ' + e.message);
  }
}

// ================================
// テスト関数
// ================================

/**
 * 訪問希望日解析のテスト
 */
function testParseVisitDates() {
  Logger.log('=== 訪問希望日解析テスト ===');
  
  const testData = [
    "2025/11/15 (9:00-20:00)",
    "2025/11/20 (9:00-20:00)",
    "2025/11/25 (13:00-20:00)"
  ];
  
  const result = parseVisitDates(testData);
  Logger.log('入力: ' + JSON.stringify(testData));
  Logger.log('出力: ' + JSON.stringify(result));
}

/**
 * 予約データ抽出のテスト（手動テスト用）
 * 注意: 実際のフォーム送信でテストしてください
 */
function testManualReservation() {
  Logger.log('=== 手動予約テストデータ作成 ===');
  
  // テストデータ
  const testData = {
    submittedAt: new Date(),
    email: 'test@example.com',
    name: 'テスト太郎',
    phone: '03-1234-5678',
    address: '東京都千代田区1-1-1',
    purpose: 'システムテスト',
    visitDatesRaw: [
      "2025/11/15 (9:00-20:00)",
      "2025/11/20 (9:00-20:00)"
    ]
  };
  
  testData.visitDates = parseVisitDates(testData.visitDatesRaw);
  testData.reservationId = generateReservationId();
  testData.token = generateToken();
  testData.status = '有効';
  testData.createdAt = new Date();
  testData.updatedAt = new Date();
  
  Logger.log('テストデータ: ' + JSON.stringify(testData, null, 2));
  
  // 営業日チェック
  const invalidDates = checkBusinessDays(testData.visitDates);
  Logger.log('無効な日付: ' + (invalidDates.length > 0 ? invalidDates.join(', ') : 'なし'));
  
  // 保存
  try {
    const rowIndex = saveToReservationsSheet(testData);
    Logger.log('Reservationsシートに保存: 行' + rowIndex);
    
    saveToVisitDatesSheet(testData);
    Logger.log('VisitDatesシートに保存完了');
    
    Logger.log('=== テスト完了 ===');
    Logger.log('スプレッドシートを確認してください');
  } catch (error) {
    Logger.log('エラー: ' + error.message);
  }
}
