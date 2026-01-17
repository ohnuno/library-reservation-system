/**
 * 職員用WebアプリのGETリクエスト処理
 */
function doGet(e) {
  return doGetStaff(e);
}

function doGetStaff(e) {
  const template = HtmlService.createTemplateFromFile('StaffInterface');
  return template.evaluate()
    .setTitle('訪問予約管理システム - 職員用')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 予約を検索
 * @param {string} searchType - 検索タイプ (id/name/email)
 * @param {string} searchValue - 検索値
 * @return {Array} 予約リスト
 */
function searchReservations(searchType, searchValue) {
  try {
    const sheet = getReservationsSheet();
    const data = sheet.getDataRange().getValues();
    const results = [];
    
    for (let i = 1; i < data.length; i++) {
      let match = false;
      
      if (searchType === 'id' && String(data[i][0]).includes(searchValue)) {
        match = true;
      } else if (searchType === 'name' && String(data[i][2]).includes(searchValue)) {
        match = true;
      } else if (searchType === 'email' && String(data[i][3]).includes(searchValue)) {
        match = true;
      }
      
      if (match) {
        results.push({
          reservationId: String(data[i][0]),
          submittedAt: formatDateTime(data[i][1]),
          name: String(data[i][2]),
          email: String(data[i][3]),
          visitDates: String(data[i][7]),
          status: String(data[i][8])
        });
      }
    }
    
    return results;
  } catch (error) {
    Logger.log('検索エラー: ' + error.message);
    throw error;
  }
}

/**
 * 予約詳細を取得
 * @param {string} reservationId - 予約ID
 * @return {Object} 予約詳細
 */
function getReservationDetail(reservationId) {
  try {
    const reservation = getReservationById(reservationId);
    if (!reservation) {
      return { error: true, message: '予約が見つかりません' };
    }
    
    // VisitDatesシートから訪問日詳細を取得
    const visitDatesSheet = getVisitDatesSheet();
    const visitData = visitDatesSheet.getDataRange().getValues();
    const visitDetails = [];
    
    for (let i = 1; i < visitData.length; i++) {
      if (visitData[i][0] === reservationId) {
        visitDetails.push({
          date: formatDate(visitData[i][1]),
          status: String(visitData[i][2]),
          canModify: visitData[i][2] === '有効'
        });
      }
    }
    
    return {
      error: false,
      reservation: {
        reservationId: reservation.reservationId,
        submittedAt: formatDateTime(reservation.submittedAt),
        name: reservation.name,
        email: reservation.email,
        phone: reservation.phone,
        address: reservation.address,
        purpose: reservation.purpose,
        status: reservation.status
      },
      visitDetails: visitDetails
    };
  } catch (error) {
    Logger.log('詳細取得エラー: ' + error.message);
    return { error: true, message: error.message };
  }
}

/**
 * 訪問日を変更
 * @param {string} reservationId - 予約ID
 * @param {Array} datesToRemove - 削除する日付
 * @param {Array} datesToAdd - 追加する日付
 * @return {Object} 結果
 */
function modifyVisitDates(reservationId, datesToRemove, datesToAdd) {
  try {
    Logger.log(`訪問日変更: ${reservationId}`);
    
    const reservation = getReservationById(reservationId);
    if (!reservation || reservation.status !== '有効') {
      return { error: true, message: '変更できない予約です' };
    }
    
    const visitDatesSheet = getVisitDatesSheet();
    const data = visitDatesSheet.getDataRange().getValues();
    
    // 削除: ステータスをキャンセルに
    datesToRemove.forEach(date => {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === reservationId && formatDate(data[i][1]) === date) {
          visitDatesSheet.getRange(i + 1, 3).setValue('キャンセル');
          
          // カレンダーイベント削除
          const eventId = data[i][3];
          if (eventId) deleteCalendarEvent(eventId);
        }
      }
    });
    
    // 追加: 新規行作成
    const newEventIds = [];
    datesToAdd.forEach(date => {
      visitDatesSheet.appendRow([
        reservationId,
        date,
        '有効',
        '',
        'TRUE'
      ]);
      
      // カレンダーイベント作成（変更履歴付き）
      const eventId = createModifiedCalendarEvent(reservationId, date, datesToRemove, datesToAdd);
      if (eventId) {
        const lastRow = visitDatesSheet.getLastRow();
        visitDatesSheet.getRange(lastRow, 4).setValue(eventId);
      }
    });
        
    // Reservationsシート更新
    updateReservationVisitDates(reservationId);
    
    // 確認メール送信
    sendVisitDateChangeConfirmation(reservation, datesToRemove, datesToAdd);
    
    return { error: false, message: '訪問日を変更しました' };
  } catch (error) {
    Logger.log('変更エラー: ' + error.message);
    return { error: true, message: error.message };
  }
}

/**
 * 変更履歴付きカレンダーイベント作成
 */
function createModifiedCalendarEvent(reservationId, dateStr, removed, added) {
  try {
    const reservation = getReservationById(reservationId);
    const calendar = getCalendar();
    const date = new Date(dateStr);
    
    let description = createEventDescription(reservation, dateStr);
    description += '\n\n【変更履歴】';
    description += '\n変更日時: ' + formatDateTime(new Date());
    if (removed.length > 0) {
      description += '\n削除した訪問日: ' + removed.join(', ');
    }
    if (added.length > 0) {
      description += '\n追加した訪問日: ' + added.join(', ');
    }
    
    const event = calendar.createAllDayEvent(
      `訪問予約: ${reservation.name}`,
      date,
      { description: description }
    );
    
    return event.getId();
  } catch (error) {
    Logger.log('イベント作成エラー: ' + error.message);
    return null;
  }
}

/**
 * 予約をキャンセル
 * @param {string} reservationId - 予約ID
 * @return {Object} 結果
 */
function cancelReservation(reservationId) {
  try {
    Logger.log(`予約キャンセル: ${reservationId}`);
    
    const reservation = getReservationById(reservationId);
    if (!reservation) {
      return { error: true, message: '予約が見つかりません' };
    }
    
    // Reservationsシート更新
    const reservationsSheet = getReservationsSheet();
    const data = reservationsSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reservationId) {
        reservationsSheet.getRange(i + 1, 9).setValue('キャンセル');
        reservationsSheet.getRange(i + 1, 13).setValue(formatDateTime(new Date()));
        break;
      }
    }
    
    // VisitDatesシート更新
    const visitDatesSheet = getVisitDatesSheet();
    const visitData = visitDatesSheet.getDataRange().getValues();
    
    for (let i = 1; i < visitData.length; i++) {
      if (visitData[i][0] === reservationId && visitData[i][2] === '有効') {
        visitDatesSheet.getRange(i + 1, 3).setValue('キャンセル');
        visitDatesSheet.getRange(i + 1, 5).setValue('FALSE');
        
        // カレンダーイベント削除
        const eventId = visitData[i][3];
        if (eventId) deleteCalendarEvent(eventId);
      }
    }
    
    // 確認メール送信
    sendCancellationConfirmationEmail(reservation);
    
    return { error: false, message: '予約をキャンセルしました' };
  } catch (error) {
    Logger.log('キャンセルエラー: ' + error.message);
    return { error: true, message: error.message };
  }
}

/**
 * Reservationsシートの訪問日リストを更新
 */
function updateReservationVisitDates(reservationId) {
  const visitDatesSheet = getVisitDatesSheet();
  const data = visitDatesSheet.getDataRange().getValues();
  
  const validDates = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId && data[i][2] === '有効') {
      validDates.push(formatDate(data[i][1]));
    }
  }
  
  const reservationsSheet = getReservationsSheet();
  const resData = reservationsSheet.getDataRange().getValues();
  
  for (let i = 1; i < resData.length; i++) {
    if (resData[i][0] === reservationId) {
      reservationsSheet.getRange(i + 1, 8).setValue(validDates.join(','));
      reservationsSheet.getRange(i + 1, 13).setValue(formatDateTime(new Date()));
      break;
    }
  }
}

/**
 * 単一カレンダーイベント作成
 */
function createSingleCalendarEvent(reservationId, dateStr) {
  try {
    const reservation = getReservationById(reservationId);
    const calendar = getCalendar();
    const date = new Date(dateStr);
    
    const event = calendar.createAllDayEvent(
      `訪問予約: ${reservation.name}`,
      date,
      { description: createEventDescription(reservation, dateStr) }
    );
    
    return event.getId();
  } catch (error) {
    Logger.log('イベント作成エラー: ' + error.message);
    return null;
  }
}

/**
 * 訪問日変更確認メール
 */
function sendVisitDateChangeConfirmation(reservation, removed, added) {
  const subject = '【変更完了】訪問予約の訪問日を変更しました';
  let body = `${reservation.name} 様\n\n`;
  body += `予約ID: ${reservation.reservationId}\n`;
  body += `訪問日の変更が完了しました。\n\n`;
  
  if (removed.length > 0) {
    body += `削除した訪問日:\n${removed.map(d => '  ・' + d).join('\n')}\n\n`;
  }
  if (added.length > 0) {
    body += `追加した訪問日:\n${added.map(d => '  ・' + d).join('\n')}\n\n`;
  }
  
  body += `変更日時: ${formatDateTime(new Date())}\n`;
  
  GmailApp.sendEmail(reservation.email, subject, body);
  Logger.log('変更確認メール送信');
}