// ================================
// Googleカレンダー管理
// ================================

/**
 * カレンダーオブジェクトを取得
 * @return {Calendar} カレンダーオブジェクト
 */
function getCalendar() {
  const calendarId = getCalendarId();
  
  if (!calendarId) {
    throw new Error('カレンダーIDが設定されていません');
  }
  
  return CalendarApp.getCalendarById(calendarId);
}

/**
 * 予約情報をカレンダーに登録
 * @param {Object} reservationData - 予約データ
 * @return {Array} イベントIDの配列
 */
function addReservationToCalendar(reservationData) {
  const calendar = getCalendar();
  const eventIds = [];
  
  try {
    // 各訪問日ごとにイベントを作成
    for (let dateStr of reservationData.visitDates) {
      const date = new Date(dateStr);
      
      // イベントのタイトル
      const title = `訪問予約: ${reservationData.name}`;
      
      // イベントの説明
      const description = createEventDescription(reservationData, dateStr);
      
      // 終日イベントとして作成
      const event = calendar.createAllDayEvent(
        title,
        date,
        {
          description: description
        }
      );
      
      const eventId = event.getId();
      eventIds.push(eventId);
      
      Logger.log(`カレンダーイベント作成: ${dateStr} (ID: ${eventId})`);
    }
    
    Logger.log(`合計 ${eventIds.length} 件のイベントを登録しました`);
    return eventIds;
    
  } catch (error) {
    Logger.log('カレンダー登録エラー: ' + error.message);
    throw error;
  }
}

/**
 * カレンダーイベントの説明文を作成
 * @param {Object} reservationData - 予約データ
 * @param {string} dateStr - 訪問日
 * @return {string} 説明文
 */
function createEventDescription(reservationData, dateStr) {
  return `【訪問予約情報】

予約ID: ${reservationData.reservationId}
訪問日: ${dateStr}

■ 訪問者情報
氏名: ${reservationData.name}
メールアドレス: ${reservationData.email}
電話番号: ${reservationData.phone}
住所: ${reservationData.address}

■ 訪問目的
${reservationData.purpose}

■ ステータス
${reservationData.status}

■ 申請日時
${formatDateTime(reservationData.submittedAt)}
`;
}

/**
 * ReservationsシートにカレンダーイベントIDを保存
 * @param {string} reservationId - 予約ID
 * @param {Array} eventIds - イベントIDの配列
 */
function updateCalendarEventIdsInReservation(reservationId, eventIds) {
  const sheet = getReservationsSheet();
  const data = sheet.getDataRange().getValues();
  
  // 予約IDで検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      // N列(14列目)にイベントIDをカンマ区切りで保存
      const eventIdsStr = eventIds.join(',');
      sheet.getRange(i + 1, 14).setValue(eventIdsStr);
      Logger.log('カレンダーイベントIDを更新: 行' + (i + 1));
      return true;
    }
  }
  
  Logger.log('警告: 予約ID ' + reservationId + ' が見つかりません');
  return false;
}

/**
 * VisitDatesシートにカレンダーイベントIDを保存
 * @param {string} reservationId - 予約ID
 * @param {Array} visitDates - 訪問日の配列
 * @param {Array} eventIds - イベントIDの配列
 */
function updateCalendarEventIdsInVisitDates(reservationId, visitDates, eventIds) {
  const sheet = getVisitDatesSheet();
  const data = sheet.getDataRange().getValues();
  
  // 訪問日とイベントIDのマッピングを作成
  const dateEventMap = {};
  for (let i = 0; i < visitDates.length; i++) {
    dateEventMap[visitDates[i]] = eventIds[i];
  }
  
  // VisitDatesシートを検索して更新
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      const dateStr = formatDate(data[i][1]);
      const eventId = dateEventMap[dateStr];
      
      if (eventId) {
        // D列(4列目)にイベントIDを保存
        sheet.getRange(i + 1, 4).setValue(eventId);
        Logger.log(`VisitDatesシート更新: 行${i + 1}, 日付${dateStr}, イベントID=${eventId}`);
      }
    }
  }
}

/**
 * カレンダーイベントを削除
 * @param {string} eventId - イベントID
 * @return {boolean} 成功したらtrue
 */
function deleteCalendarEvent(eventId) {
  if (!eventId) {
    Logger.log('イベントIDが指定されていません');
    return false;
  }
  
  try {
    const calendar = getCalendar();
    const event = calendar.getEventById(eventId);
    
    if (event) {
      event.deleteEvent();
      Logger.log('カレンダーイベントを削除しました: ' + eventId);
      return true;
    } else {
      Logger.log('イベントが見つかりません: ' + eventId);
      return false;
    }
  } catch (error) {
    Logger.log('カレンダーイベント削除エラー: ' + error.message);
    return false;
  }
}

/**
 * 予約IDに紐づくすべてのカレンダーイベントを削除
 * @param {string} reservationId - 予約ID
 * @return {number} 削除したイベント数
 */
function deleteReservationCalendarEvents(reservationId) {
  const sheet = getReservationsSheet();
  const data = sheet.getDataRange().getValues();
  
  // 予約IDで検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      const eventIdsStr = data[i][13]; // N列(14列目)
      
      if (!eventIdsStr) {
        Logger.log('カレンダーイベントIDが登録されていません');
        return 0;
      }
      
      const eventIds = eventIdsStr.split(',');
      let deletedCount = 0;
      
      for (let eventId of eventIds) {
        if (deleteCalendarEvent(eventId.trim())) {
          deletedCount++;
        }
      }
      
      Logger.log(`${deletedCount} 件のカレンダーイベントを削除しました`);
      return deletedCount;
    }
  }
  
  Logger.log('予約ID ' + reservationId + ' が見つかりません');
  return 0;
}

/**
 * カレンダーイベントを更新
 * @param {string} eventId - イベントID
 * @param {Object} updates - 更新内容 {title: ..., description: ...}
 * @return {boolean} 成功したらtrue
 */
function updateCalendarEvent(eventId, updates) {
  if (!eventId) {
    Logger.log('イベントIDが指定されていません');
    return false;
  }
  
  try {
    const calendar = getCalendar();
    const event = calendar.getEventById(eventId);
    
    if (!event) {
      Logger.log('イベントが見つかりません: ' + eventId);
      return false;
    }
    
    if (updates.title) {
      event.setTitle(updates.title);
    }
    
    if (updates.description) {
      event.setDescription(updates.description);
    }
    
    Logger.log('カレンダーイベントを更新しました: ' + eventId);
    return true;
    
  } catch (error) {
    Logger.log('カレンダーイベント更新エラー: ' + error.message);
    return false;
  }
}

// ================================
// テスト関数
// ================================

/**
 * カレンダー登録のテスト
 */
function testCalendarRegistration() {
  Logger.log('=== カレンダー登録テスト ===');
  
  // テストデータ
  const testData = {
    reservationId: 'RSV20251113TEST',
    name: 'テスト太郎',
    email: 'test@example.com',
    phone: '03-1234-5678',
    address: '東京都千代田区1-1-1',
    purpose: 'カレンダー登録テスト',
    visitDates: ['2025/11/20', '2025/11/25'],
    status: '有効',
    submittedAt: new Date()
  };
  
  try {
    const eventIds = addReservationToCalendar(testData);
    
    Logger.log('\n✓ カレンダー登録成功');
    Logger.log('登録されたイベントID:');
    eventIds.forEach((id, index) => {
      Logger.log(`  ${index + 1}. ${id}`);
    });
    
    Logger.log('\n次の作業:');
    Logger.log('1. Googleカレンダーを開いて、イベントが登録されているか確認してください');
    Logger.log('2. イベントの詳細を開いて、説明文が正しく表示されているか確認してください');
    Logger.log('3. 確認後、testDeleteCalendarEvents() を実行してテストイベントを削除してください');
    
    // テストイベントIDを保存（削除用）
    PropertiesService.getScriptProperties().setProperty('TEST_EVENT_IDS', eventIds.join(','));
    
  } catch (error) {
    Logger.log('\n✗ カレンダー登録失敗');
    Logger.log('エラー: ' + error.message);
  }
}

/**
 * テストイベントの削除
 */
function testDeleteCalendarEvents() {
  Logger.log('=== テストイベント削除 ===');
  
  const eventIdsStr = PropertiesService.getScriptProperties().getProperty('TEST_EVENT_IDS');
  
  if (!eventIdsStr) {
    Logger.log('削除するテストイベントがありません');
    return;
  }
  
  const eventIds = eventIdsStr.split(',');
  let deletedCount = 0;
  
  for (let eventId of eventIds) {
    if (deleteCalendarEvent(eventId.trim())) {
      deletedCount++;
    }
  }
  
  Logger.log(`${deletedCount} 件のテストイベントを削除しました`);
  
  // 保存したイベントIDをクリア
  PropertiesService.getScriptProperties().deleteProperty('TEST_EVENT_IDS');
}

/**
 * カレンダーアクセステスト
 */
function testCalendarAccess() {
  Logger.log('=== カレンダーアクセステスト ===');
  
  try {
    const calendar = getCalendar();
    Logger.log('カレンダー名: ' + calendar.getName());
    Logger.log('カレンダーID: ' + calendar.getId());
    Logger.log('タイムゾーン: ' + calendar.getTimeZone());
    
    // 今日から7日間のイベントを取得
    const today = new Date();
    const nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
    const events = calendar.getEvents(today, nextWeek);
    
    Logger.log(`\n今日から7日間のイベント数: ${events.length}件`);
    
    if (events.length > 0) {
      Logger.log('最初の3件:');
      for (let i = 0; i < Math.min(3, events.length); i++) {
        Logger.log(`  ${i + 1}. ${events[i].getTitle()} (${formatDate(events[i].getStartTime())})`);
      }
    }
    
    Logger.log('\n✓ カレンダーアクセス成功');
    
  } catch (error) {
    Logger.log('\n✗ カレンダーアクセス失敗');
    Logger.log('エラー: ' + error.message);
  }
}