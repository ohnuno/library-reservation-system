// ================================
// メール本文動的生成
// ================================

/**
 * フォーム設定に基づいて予約情報のHTMLを動的生成
 * @param {Object} reservationData - 予約データ
 * @return {string} HTML
 */
function generateReservationInfoHTML(reservationData) {
  const questions = getFormQuestions();
  let html = '';
  
  for (let q of questions) {
    // email, visit_datesは別途表示するのでスキップ
    if (q.itemId === 'email' || q.itemId === 'visit_dates') {
      continue;
    }
    
    // 予約データから値を取得
    let value = reservationData[q.itemId];
    
    if (value) {
      // 改行を<br>に変換
      if (typeof value === 'string') {
        value = value.replace(/\n/g, '<br>');
      }
      
      html += `
        <div class="info-row">
          <span class="label">${q.title}:</span> ${value}
        </div>`;
    }
  }
  
  return html;
}

/**
 * フォーム設定に基づいて予約情報のプレーンテキストを動的生成
 * @param {Object} reservationData - 予約データ
 * @return {string} プレーンテキスト
 */
function generateReservationInfoPlain(reservationData) {
  const questions = getFormQuestions();
  let text = '';
  
  for (let q of questions) {
    // email, visit_datesは別途表示するのでスキップ
    if (q.itemId === 'email' || q.itemId === 'visit_dates') {
      continue;
    }
    
    // 予約データから値を取得
    let value = reservationData[q.itemId];
    
    if (value) {
      text += `${q.title}: ${value}\n`;
    }
  }
  
  return text;
}

/**
 * 職員通知用の予約情報HTMLを動的生成
 * @param {Object} reservationData - 予約データ
 * @return {string} HTML
 */
function generateStaffReservationInfoHTML(reservationData) {
  const questions = getFormQuestions();
  let html = '';
  
  for (let q of questions) {
    // visit_datesは別途表示するのでスキップ
    if (q.itemId === 'visit_dates') {
      continue;
    }
    
    // 予約データから値を取得
    let value = reservationData[q.itemId];
    
    if (value) {
      // 改行を<br>に変換
      if (typeof value === 'string') {
        value = value.replace(/\n/g, '<br>');
      }
      
      html += `
        <div class="info-row">
          <span class="label">${q.title}:</span> ${value}
        </div>`;
    }
  }
  
  return html;
}

/**
 * 職員通知用の予約情報プレーンテキストを動的生成
 * @param {Object} reservationData - 予約データ
 * @return {string} プレーンテキスト
 */
function generateStaffReservationInfoPlain(reservationData) {
  const questions = getFormQuestions();
  let text = '';
  
  for (let q of questions) {
    // visit_datesは別途表示するのでスキップ
    if (q.itemId === 'visit_dates') {
      continue;
    }
    
    // 予約データから値を取得
    let value = reservationData[q.itemId];
    
    if (value) {
      text += `${q.title}: ${value}\n`;
    }
  }
  
  return text;
}
// ================================
// メール送信
// ================================

/**
 * 予約完了メールをユーザーに送信
 * @param {Object} reservationData - 予約データ
 */
function sendConfirmationEmail(reservationData) {
  try {
    const subject = '【予約完了】訪問予約を受け付けました';
    const htmlBody = createConfirmationEmailHTML(reservationData);
    const plainBody = createConfirmationEmailPlain(reservationData);
    
    GmailApp.sendEmail(
      reservationData.email,
      subject,
      plainBody,
      {
        htmlBody: htmlBody,
        name: '訪問予約受付システム'
      }
    );
    
    Logger.log('予約完了メールを送信しました: ' + reservationData.email);
    
  } catch (error) {
    Logger.log('予約完了メール送信エラー: ' + error.message);
    throw error;
  }
}

/**
 * 予約完了メールのHTML本文を作成
 * @param {Object} reservationData - 予約データ
 * @return {string} HTML本文
 */
function createConfirmationEmailHTML(reservationData) {
  // 訪問日リストを整形
  const visitDatesList = reservationData.visitDates
    .map(date => `  ・${date}`)
    .join('<br>');
  
  // WebアプリURL（変更・キャンセル用、再発行用）
  const webAppUrl = getConfigValue('WebアプリURL') || 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
  const cancelUrl = `${webAppUrl}?action=cancel&id=${reservationData.reservationId}`;
  const reissueUrl = `${webAppUrl}?action=reissue&id=${reservationData.reservationId}`;
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {
      font-family: "Hiragino Kaku Gothic ProN", "Hiragino Sans", Meiryo, sans-serif;
      line-height: 1.6;
      color: #333;
    }
    .container {
      max-width: 600px;
      margin: 0 auto;
      padding: 20px;
    }
    .header {
      background-color: #4285f4;
      color: white;
      padding: 20px;
      text-align: center;
      border-radius: 5px 5px 0 0;
    }
    .content {
      background-color: #f9f9f9;
      padding: 20px;
      border: 1px solid #ddd;
      border-radius: 0 0 5px 5px;
    }
    .section {
      margin-bottom: 20px;
      padding: 15px;
      background-color: white;
      border-left: 4px solid #4285f4;
    }
    .section-title {
      font-weight: bold;
      font-size: 16px;
      margin-bottom: 10px;
      color: #4285f4;
    }
    .info-row {
      margin: 8px 0;
    }
    .label {
      font-weight: bold;
      color: #666;
    }
    .qr-container {
      text-align: center;
      padding: 20px;
      background-color: white;
      margin: 20px 0;
    }
    .qr-image {
      max-width: 300px;
      height: auto;
    }
    .notice {
      background-color: #fff3cd;
      border: 1px solid #ffc107;
      padding: 15px;
      border-radius: 5px;
      margin: 20px 0;
    }
    .button {
      display: inline-block;
      padding: 10px 20px;
      background-color: #4285f4;
      color: white !important;
      text-decoration: none;
      border-radius: 5px;
      margin: 5px;
    }
    .footer {
      margin-top: 30px;
      padding-top: 20px;
      border-top: 1px solid #ddd;
      font-size: 12px;
      color: #666;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1 style="margin: 0;">訪問予約完了</h1>
    </div>
    
    <div class="content">
      <p>${reservationData.name} 様</p>
      <p>訪問予約を受け付けました。<br>以下の内容をご確認ください。</p>
      
      <div class="section">
        <div class="section-title">■ 予約内容</div>
        <div class="info-row">
          <span class="label">予約ID:</span> ${reservationData.reservationId}
        </div>
        ${generateReservationInfoHTML(reservationData)}
        <div class="info-row">
          <span class="label">訪問日:</span><br>
          ${visitDatesList}
        </div>
      </div>

      <div class="section">
        <div class="section-title">▪️ 訪問当日について</div>
        <p>訪問当日は、下記のQRコードを受付にてご提示ください。</p>
        
        <div class="qr-container">
          <img src="${reservationData.qrCodeEmailUrl}" alt="QRコード" class="qr-image">
          <p style="margin-top: 15px; font-size: 12px; color: #666; word-break: break-all;">
            <strong>確認用URL:</strong><br>
            <a href="${reservationData.qrContent}" style="color: #4285f4;">${reservationData.qrContent}</a>
          </p>
        </div>
        
        <div class="notice">
          <strong>⚠️ 重要</strong><br>
          QRコードは訪問日当日のみ有効です。<br>
          訪問日以外では使用できませんのでご注意ください。
        </div>
      </div>
      
      <div class="section">
        <div class="section-title">▪️ 予約の変更・キャンセル</div>
        <p>予約内容の変更やキャンセルをご希望の場合は、下記よりお手続きください。</p>
        <p style="text-align: center;">
          <a href="${getConfigValue('変更キャンセルフォームURL')}" class="button">予約の変更・キャンセル</a>
        </p>
      </div>
      
      <div class="section">
        <div class="section-title">▪️ QRコードの再発行</div>
        <p>QRコードを紛失された場合は、下記より再発行できます。</p>
        <p style="text-align: center;">
          <a href="${reissueUrl}" class="button">QRコード再発行</a>
        </p>
      </div>
      
      <div class="footer">
        <p>このメールは自動送信されています。</p>
        <p>ご不明な点がございましたら、下記までお問い合わせください。</p>
        <p>
          お問い合わせ先: ${getStaffEmail() || '(未設定)'}<br>
          訪問予約受付システム
        </p>
      </div>
    </div>
  </div>
</body>
</html>
`;
}

/**
 * 予約完了メールのプレーンテキスト本文を作成
 * @param {Object} reservationData - 予約データ
 * @return {string} プレーンテキスト本文
 */
function createConfirmationEmailPlain(reservationData) {
  const visitDatesList = reservationData.visitDates
    .map(date => `  ・${date}`)
    .join('\n');
  
  const webAppUrl = getConfigValue('WebアプリURL') || 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
  const cancelUrl = `${webAppUrl}?action=cancel&id=${reservationData.reservationId}`;
  const reissueUrl = `${webAppUrl}?action=reissue&id=${reservationData.reservationId}`;
  
  return `${reservationData.name} 様

訪問予約を受け付けました。
以下の内容をご確認ください。

━━━━━━━━━━━━━━━━━━━━
■ 予約内容
━━━━━━━━━━━━━━━━━━━━
予約ID: ${reservationData.reservationId}
${generateReservationInfoPlain(reservationData)}訪問日:
${visitDatesList}

━━━━━━━━━━━━━━━━━━━━
▪️ 訪問当日について
━━━━━━━━━━━━━━━━━━━━
訪問当日は、メールに添付のQRコードを受付にてご提示ください。

確認用URL:
${reservationData.qrContent}

⚠️ 重要
QRコードは訪問日当日のみ有効です。
訪問日以外では使用できませんのでご注意ください。

━━━━━━━━━━━━━━━━━━━━
▪️ 予約の変更・キャンセル
━━━━━━━━━━━━━━━━━━━━
予約内容の変更やキャンセルをご希望の場合は、下記URLよりお手続きください。
${getConfigValue('変更キャンセルフォームURL')}

━━━━━━━━━━━━━━━━━━━━
▪️ QRコードの再発行
━━━━━━━━━━━━━━━━━━━━
QRコードを紛失された場合は、下記URLより再発行できます。
${reissueUrl}

━━━━━━━━━━━━━━━━━━━━

このメールは自動送信されています。

ご不明な点がございましたら、下記までお問い合わせください。
お問い合わせ先: ${getStaffEmail() || '(未設定)'}
訪問予約受付システム
`;
}

/**
 * 職員向け通知メールを送信
 * @param {Object} reservationData - 予約データ
 */
function sendStaffNotification(reservationData) {
  try {
    const staffEmail = getStaffEmail();
    
    if (!staffEmail) {
      Logger.log('警告: 職員通知メールアドレスが設定されていません');
      return;
    }
    
    const subject = `【新規予約】${reservationData.name} 様 (${reservationData.visitDates[0]}他)`;
    const htmlBody = createStaffNotificationHTML(reservationData);
    const plainBody = createStaffNotificationPlain(reservationData);
    
    GmailApp.sendEmail(
      staffEmail,
      subject,
      plainBody,
      {
        htmlBody: htmlBody,
        name: '訪問予約受付システム'
      }
    );
    
    Logger.log('職員通知メールを送信しました: ' + staffEmail);
    
  } catch (error) {
    Logger.log('職員通知メール送信エラー: ' + error.message);
    // 職員通知メール送信失敗でも処理は続行
  }
}

/**
 * 職員向け通知メールのHTML本文を作成
 * @param {Object} reservationData - 予約データ
 * @return {string} HTML本文
 */
function createStaffNotificationHTML(reservationData) {
  const visitDatesList = reservationData.visitDates
    .map(date => `  ・${date}`)
    .join('<br>');
  
  const calendarUrl = `https://calendar.google.com/calendar/r?cid=${encodeURIComponent(getCalendarId())}`;
  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${getSpreadsheetId()}`;
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {
      font-family: "Hiragino Kaku Gothic ProN", "Hiragino Sans", Meiryo, sans-serif;
      line-height: 1.6;
      color: #333;
    }
    .container {
      max-width: 600px;
      margin: 0 auto;
      padding: 20px;
    }
    .header {
      background-color: #34a853;
      color: white;
      padding: 20px;
      text-align: center;
      border-radius: 5px 5px 0 0;
    }
    .content {
      background-color: #f9f9f9;
      padding: 20px;
      border: 1px solid #ddd;
      border-radius: 0 0 5px 5px;
    }
    .section {
      margin-bottom: 20px;
      padding: 15px;
      background-color: white;
      border-left: 4px solid #34a853;
    }
    .section-title {
      font-weight: bold;
      font-size: 16px;
      margin-bottom: 10px;
      color: #34a853;
    }
    .info-row {
      margin: 8px 0;
    }
    .label {
      font-weight: bold;
      color: #666;
      display: inline-block;
      width: 120px;
    }
    .button {
      display: inline-block;
      padding: 10px 20px;
      background-color: #34a853;
      color: white !important;
      text-decoration: none;
      border-radius: 5px;
      margin: 5px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1 style="margin: 0;">【新規予約通知】</h1>
    </div>
    
    <div class="content">
      <p>新しい訪問予約が登録されました。</p>
      
      <div class="section">
        <div class="section-title">■ 予約情報</div>
        <div class="info-row">
          <span class="label">予約ID:</span> ${reservationData.reservationId}
        </div>
        <div class="info-row">
          <span class="label">申請日時:</span> ${formatDateTime(reservationData.submittedAt)}
        </div>
        ${generateStaffReservationInfoHTML(reservationData)}
      </div>
      
      <div class="section">
        <div class="section-title">■ 訪問日</div>
        ${visitDatesList}
      </div>

      <div class="section">
        <div class="section-title">▪️ 訪問目的</div>
        <p>${reservationData.purpose.replace(/\n/g, '<br>')}</p>
      </div>
      
      <div class="section">
        <div class="section-title">▪️ クイックアクセス</div>
        <p style="text-align: center;">
          <a href="${calendarUrl}" class="button">カレンダーを開く</a>
          <a href="${spreadsheetUrl}" class="button">予約一覧を開く</a>
        </p>
      </div>
    </div>
  </div>
</body>
</html>
`;
}

/**
 * 職員向け通知メールのプレーンテキスト本文を作成
 * @param {Object} reservationData - 予約データ
 * @return {string} プレーンテキスト本文
 */
function createStaffNotificationPlain(reservationData) {
  const visitDatesList = reservationData.visitDates
    .map(date => `  ・${date}`)
    .join('\n');
  
  const calendarUrl = `https://calendar.google.com/calendar/r?cid=${encodeURIComponent(getCalendarId())}`;
  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${getSpreadsheetId()}`;
  
  return `【新規予約通知】

新しい訪問予約が登録されました。

━━━━━━━━━━━━━━━━━━━━
■ 予約情報
━━━━━━━━━━━━━━━━━━━━
予約ID: ${reservationData.reservationId}
申請日時: ${formatDateTime(reservationData.submittedAt)}
${generateStaffReservationInfoPlain(reservationData)}

━━━━━━━━━━━━━━━━━━━━
■ 訪問日
━━━━━━━━━━━━━━━━━━━━
${visitDatesList}

━━━━━━━━━━━━━━━━━━━━
▪️ 訪問目的
━━━━━━━━━━━━━━━━━━━━
${reservationData.purpose}

━━━━━━━━━━━━━━━━━━━━
▪️ クイックアクセス
━━━━━━━━━━━━━━━━━━━━
カレンダー: ${calendarUrl}
予約一覧: ${spreadsheetUrl}

---
訪問予約受付システム
`;
}

// ================================
// テスト関数
// ================================

/**
 * 予約完了メール送信のテスト
 */
function testSendConfirmationEmail() {
  Logger.log('=== 予約完了メール送信テスト ===');
  
  // テストデータ
  const testData = {
    reservationId: 'RSV20251113TEST',
    name: 'テスト太郎',
    email: Session.getActiveUser().getEmail(), // 自分のメールアドレス
    phone: '03-1234-5678',
    address: '東京都千代田区1-1-1',
    purpose: 'メール送信テスト\n複数行のテストです。',
    visitDates: ['2025/11/20', '2025/11/25'],
    status: '有効',
    submittedAt: new Date(),
    qrCodeEmailUrl: 'https://quickchart.io/qr?text=TEST&size=300'
  };
  
  try {
    sendConfirmationEmail(testData);
    
    Logger.log('\n✓ 予約完了メール送信成功');
    Logger.log('送信先: ' + testData.email);
    Logger.log('\n次の作業:');
    Logger.log('1. メールボックスを確認してメールが届いているか確認してください');
    Logger.log('2. HTMLメールとして正しく表示されているか確認してください');
    Logger.log('3. QRコード画像が表示されているか確認してください');
    
  } catch (error) {
    Logger.log('\n✗ 予約完了メール送信失敗');
    Logger.log('エラー: ' + error.message);
  }
}

/**
 * 職員通知メール送信のテスト
 */
function testSendStaffNotification() {
  Logger.log('=== 職員通知メール送信テスト ===');
  
  const staffEmail = getStaffEmail();
  
  if (!staffEmail) {
    Logger.log('エラー: 職員通知メールアドレスが設定されていません');
    Logger.log('Configシートの「職員通知メールアドレス」を設定してください');
    return;
  }
  
  // テストデータ
  const testData = {
    reservationId: 'RSV20251113TEST',
    name: 'テスト太郎',
    email: 'test@example.com',
    phone: '03-1234-5678',
    address: '東京都千代田区1-1-1',
    purpose: '職員通知メールテスト',
    visitDates: ['2025/11/20', '2025/11/25'],
    status: '有効',
    submittedAt: new Date()
  };
  
  try {
    sendStaffNotification(testData);
    
    Logger.log('\n✓ 職員通知メール送信成功');
    Logger.log('送信先: ' + staffEmail);
    Logger.log('\n次の作業:');
    Logger.log('1. 職員用メールボックスを確認してメールが届いているか確認してください');
    Logger.log('2. HTMLメールとして正しく表示されているか確認してください');
    
  } catch (error) {
    Logger.log('\n✗ 職員通知メール送信失敗');
    Logger.log('エラー: ' + error.message);
  }
}