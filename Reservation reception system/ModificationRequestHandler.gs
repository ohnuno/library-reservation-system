/**
 * 変更・キャンセルフォーム送信時の処理
 */
function onModificationFormSubmit(e) {
  try {
    Logger.log('=== 変更・キャンセル依頼処理開始 ===');
    
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();
    
    // フォーム回答を抽出
    const requestData = {
      submittedAt: formResponse.getTimestamp(),
      email: formResponse.getRespondentEmail(),
      reservationId: '',
      requestType: '',
      newVisitDates: [],
      cancelReason: ''
    };
    
    itemResponses.forEach(itemResponse => {
      const title = itemResponse.getItem().getTitle();
      const response = itemResponse.getResponse();
      
      if (title === '予約ID') {
        requestData.reservationId = response;
      } else if (title === '依頼内容') {
        requestData.requestType = response;
      } else if (title === '変更後の訪問希望日') {
        requestData.newVisitDates = parseVisitDates(response);
      } else if (title === 'キャンセル理由') {
        requestData.cancelReason = response;
      }
    });
    
    Logger.log('依頼データ: ' + JSON.stringify(requestData));
    
    // 本人確認
    const reservation = getReservationById(requestData.reservationId);
    
    if (!reservation) {
      sendModificationErrorEmail(requestData.email, '予約IDが見つかりません。');
      Logger.log('エラー: 予約が見つかりません');
      return;
    }
    
    if (reservation.email !== requestData.email) {
      sendModificationErrorEmail(requestData.email, '予約IDとメールアドレスが一致しません。');
      Logger.log('エラー: メールアドレス不一致');
      return;
    }
    
    // 職員への通知メール送信
    sendStaffModificationRequest(requestData, reservation);
    
    // 訪問者への受付確認メール送信
    sendModificationRequestConfirmation(requestData);
    
    Logger.log('=== 変更・キャンセル依頼処理完了 ===');
    
  } catch (error) {
    Logger.log('エラー: ' + error.message);
    Logger.log(error.stack);
  }
}

/**
 * エラーメール送信
 */
function sendModificationErrorEmail(email, message) {
  const subject = '【エラー】訪問予約 変更・キャンセル依頼';
  const body = `ご依頼いただいた内容を処理できませんでした。\n\nエラー内容:\n${message}\n\n予約完了メールに記載の予約IDとメールアドレスをご確認ください。`;
  
  GmailApp.sendEmail(email, subject, body);
  Logger.log('エラーメール送信: ' + email);
}

/**
 * 職員への通知メール送信
 */
function sendStaffModificationRequest(requestData, reservation) {
  const staffEmail = getStaffEmail();
  if (!staffEmail) return;
  
  const subject = `【変更・キャンセル依頼】${reservation.name} 様 (${requestData.reservationId})`;
  
  let body = `訪問予約の変更・キャンセル依頼がありました。\n\n`;
  body += `■ 予約情報\n`;
  body += `予約ID: ${requestData.reservationId}\n`;
  body += `氏名: ${reservation.name}\n`;
  body += `メール: ${reservation.email}\n`;
  body += `現在の訪問日: ${reservation.visitDates}\n`;
  body += `現在のステータス: ${reservation.status}\n\n`;
  body += `■ 依頼内容\n`;
  body += `依頼タイプ: ${requestData.requestType}\n`;
  
  if (requestData.requestType === '訪問日変更') {
    body += `変更希望日: ${requestData.newVisitDates.join(', ')}\n`;
  } else {
    body += `キャンセル理由: ${requestData.cancelReason || '(記載なし)'}\n`;
  }
  
  body += `\n依頼日時: ${formatDateTime(requestData.submittedAt)}\n`;
  body += `\n※職員用管理画面で処理を行ってください。`;
  
  GmailApp.sendEmail(staffEmail, subject, body);
  Logger.log('職員通知メール送信');
}

/**
 * 訪問者への受付確認メール送信
 */
function sendModificationRequestConfirmation(requestData) {
  const subject = '【受付完了】訪問予約 変更・キャンセル依頼';
  
  let body = `${requestData.reservationId} の変更・キャンセル依頼を受け付けました。\n\n`;
  body += `依頼内容: ${requestData.requestType}\n`;
  body += `受付日時: ${formatDateTime(requestData.submittedAt)}\n\n`;
  body += `職員が確認後、処理を行います。\n`;
  body += `完了後、改めてメールでお知らせします。\n\n`;
  body += `※このメールは自動送信されています。`;
  
  GmailApp.sendEmail(requestData.email, subject, body);
  Logger.log('受付確認メール送信');
}