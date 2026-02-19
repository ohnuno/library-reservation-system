// ================================
// 学外者入館管理システム - 職員用API
// ================================

/**
 * 入館中の訪問者を取得（退館処理用）
 * @return {Array} 入館中の訪問者リスト
 */
function getActiveVisitors() {
  const visitDatesSheet = getVisitDatesSheet();
  const reservationsSheet = getReservationsSheet();
  
  const visitData = visitDatesSheet.getDataRange().getValues();
  const reservationData = reservationsSheet.getDataRange().getValues();
  
  const today = formatDate(getToday());
  const visitors = [];
  
  // 予約データをマップに変換
  const reservationMap = {};
  for (let i = 1; i < reservationData.length; i++) {
    const reservationId = reservationData[i][0]; // A列: 予約ID
    reservationMap[reservationId] = {
      name: reservationData[i][2],        // C列: 氏名
      purpose: reservationData[i][6],     // G列: 訪問目的
      affiliation: reservationData[i][5]  // F列: 所属機関
    };
  }
  
  // VisitDatesシートから本日の入館中レコードを取得
  for (let i = 1; i < visitData.length; i++) {
    const visitDate = formatDate(visitData[i][1]); // B列: 訪問日
    const status = visitData[i][2];                // C列: ステータス
    const reservationId = visitData[i][0];         // A列: 予約ID
    
    // 本日かつステータスが「入館済」（受付済み）のレコード
    if (visitDate === today && status === '入館済') {
      const reservation = reservationMap[reservationId];
      
      if (reservation) {
        visitors.push({
          reservationId: reservationId,
          visitDate: visitDate,
          name: reservation.name,
          purpose: reservation.purpose,
          affiliation: reservation.affiliation,
          status: status
        });
      }
    }
  }
  
  // 予約ID順にソート
  visitors.sort((a, b) => b.reservationId.localeCompare(a.reservationId));
  
  return visitors;
}

/**
 * 退館を記録（現在時刻）
 * @param {string} reservationId - 予約ID
 * @return {Object} 結果
 */
function recordExitNow(reservationId) {
  return recordExit(reservationId, null);
}

/**
 * 退館を記録（時刻指定）- VisitDatesシートに記録する版
 * @param {string} reservationId - 予約ID
 * @param {string} exitTime - 退館時刻（HH:mm形式、nullの場合は現在時刻）
 * @return {Object} 結果
 */
function recordExit(reservationId, exitTime) {
  try {
    const visitDatesSheet = getVisitDatesSheet();
    const data = visitDatesSheet.getDataRange().getValues();
    const today = formatDate(getToday());
    
    let foundRow = -1;
    
    // 予約IDと訪問日で検索（本日の訪問レコードを探す）
    for (let i = 1; i < data.length; i++) {
      const vid = data[i][0];           // A列: 予約ID
      const vdate = formatDate(data[i][1]); // B列: 訪問日
      
      if (vid === reservationId && vdate === today) {
        foundRow = i + 1;
        break;
      }
    }
    
    if (foundRow === -1) {
      return {
        success: false,
        message: '本日の訪問記録が見つかりません'
      };
    }
    
    // 退館時刻を決定
    let timeToRecord;
    if (exitTime) {
      // 時刻指定の場合
      const todayDate = getToday();
      const [hours, minutes] = exitTime.split(':');
      todayDate.setHours(parseInt(hours), parseInt(minutes), 0, 0);
      timeToRecord = todayDate;
    } else {
      // 現在時刻
      timeToRecord = getNow();
    }
    
    // VisitDatesシートに退館時刻を記録
    visitDatesSheet.getRange(foundRow, 8).setValue(timeToRecord); // H列: 退館時刻
    
    // ステータスを「退館済」に更新
    visitDatesSheet.getRange(foundRow, 3).setValue('退館済'); // C列: ステータス
    
    // Reservationsシートのステータスも更新（すべての訪問日が退館済なら）
    updateReservationStatusIfAllExited(reservationId);
    
    Logger.log(`退館記録: ${reservationId} (${today}) at ${formatDateTime(timeToRecord)}`);
    
    return {
      success: true,
      message: '退館を記録しました',
      exitTime: formatDateTime(timeToRecord)
    };
    
  } catch (error) {
    Logger.log('退館記録エラー: ' + error.message);
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.message
    };
  }
}

/**
 * 入館記録を検索
 * @param {Object} criteria - 検索条件
 * @return {Array} 検索結果
 */
function searchVisitRecords(criteria) {
  const reservationsSheet = getReservationsSheet();
  const data = reservationsSheet.getDataRange().getValues();
  
  const records = [];
  
  for (let i = 1; i < data.length; i++) {
    const record = {
      reservationId: data[i][0],      // A列: 予約ID
      submittedAt: formatDateTime(data[i][1]), // B列: 申請日時
      name: data[i][2],               // C列: 氏名
      email: data[i][3],              // D列: メールアドレス
      phone: data[i][4],              // E列: 電話番号
      affiliation: data[i][5],        // F列: 所属機関
      purpose: data[i][6],            // G列: 訪問目的
      visitDates: data[i][7],         // H列: 訪問日リスト
      status: data[i][8],             // I列: ステータス
      createdAt: formatDateTime(data[i][11]), // L列: 作成日時
      updatedAt: formatDateTime(data[i][12]), // M列: 更新日時
      exitTime: data[i][17] ? formatDateTime(data[i][17]) : '' // 🔥 追加: R列（18列目）退館時刻
    };
    
    // 検索条件でフィルタ
    let match = true;
    
    if (criteria.startDate) {
      const visitDatesList = String(record.visitDates).split(',');
      const hasDateInRange = visitDatesList.some(vd => {
        return compareDates(vd.trim(), criteria.startDate) >= 0;
      });
      if (!hasDateInRange) match = false;
    }
    
    if (criteria.endDate) {
      const visitDatesList = String(record.visitDates).split(',');
      const hasDateInRange = visitDatesList.some(vd => {
        return compareDates(vd.trim(), criteria.endDate) <= 0;
      });
      if (!hasDateInRange) match = false;
    }
    
    if (criteria.name && !record.name.includes(criteria.name)) {
      match = false;
    }
    
    if (criteria.purpose && record.purpose !== criteria.purpose) {
      match = false;
    }
    
    if (criteria.status && record.status !== criteria.status) {
      match = false;
    }
    
    if (match) {
      records.push(record);
    }
  }
  
  // 申請日時の降順でソート
  records.sort((a, b) => b.submittedAt.localeCompare(a.submittedAt));
  
  return records;
}

/**
 * 予約詳細を取得
 * @param {string} reservationId - 予約ID
 * @return {Object} 予約詳細
 */
function getRecordDetail(reservationId) {
  const reservationsSheet = getReservationsSheet();
  const visitDatesSheet = getVisitDatesSheet();
  
  const data = reservationsSheet.getDataRange().getValues();
  
  // Reservationsシートから検索
  let record = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      record = {
        reservationId: data[i][0],
        submittedAt: formatDateTime(data[i][1]),
        name: data[i][2],
        email: data[i][3],
        phone: data[i][4],
        affiliation: data[i][5],
        purpose: data[i][6],
        visitDates: data[i][7],
        status: data[i][8],
        qrCodeUrl: data[i][9],
        createdAt: formatDateTime(data[i][11]),
        updatedAt: formatDateTime(data[i][12]),
        exitTime: data[i][17] ? formatDateTime(data[i][17]) : '' // 🔥 追加: R列（18列目）退館時刻
      };
      break;
    }
  }
  
  if (!record) {
    return null;
  }
  
  // VisitDatesシートから詳細を取得
  const visitData = visitDatesSheet.getDataRange().getValues();
  const visitDetails = [];
  
  for (let i = 1; i < visitData.length; i++) {
    if (visitData[i][0] === reservationId) {
      visitDetails.push({
        visitDate: formatDate(visitData[i][1]),
        status: visitData[i][2]
      });
    }
  }
  
  // 訪問日順にソート
  visitDetails.sort((a, b) => compareDates(a.visitDate, b.visitDate));
  
  record.visitDetails = visitDetails;
  
  return record;
}


/**
 * 入館を記録（QRコード受付時）
 * ※ 既存の予約受付システムのcompleteCheckIn関数に相当
 * @param {string} reservationId - 予約ID
 * @param {string} visitDate - 訪問日（yyyy/MM/dd）
 * @return {Object} 結果
 */
function recordEntry(reservationId, visitDate) {
  try {
    const visitDatesSheet = getVisitDatesSheet();
    const data = visitDatesSheet.getDataRange().getValues();
    
    let foundRow = -1;
    
    // 予約IDと訪問日で検索
    for (let i = 1; i < data.length; i++) {
      const vid = data[i][0];           // A列: 予約ID
      const vdate = formatDate(data[i][1]); // B列: 訪問日
      
      if (vid === reservationId && vdate === visitDate) {
        foundRow = i + 1;
        break;
      }
    }
    
    if (foundRow === -1) {
      return {
        success: false,
        message: '訪問記録が見つかりません'
      };
    }
    
    const entryTime = getNow();
    
    // 入館時刻を記録
    visitDatesSheet.getRange(foundRow, 6).setValue(entryTime); // F列: 入館時刻
    
    // ステータスを「入館済」に更新
    visitDatesSheet.getRange(foundRow, 3).setValue('入館済'); // C列: ステータス
    
    Logger.log(`入館記録: ${reservationId} (${visitDate}) at ${formatDateTime(entryTime)}`);
    
    return {
      success: true,
      message: '入館を記録しました',
      entryTime: formatDateTime(entryTime)
    };
    
  } catch (error) {
    Logger.log('入館記録エラー: ' + error.message);
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.message
    };
  }
}

/**
 * 再入館を記録
 * @param {string} reservationId - 予約ID
 * @param {string} visitDate - 訪問日（yyyy/MM/dd）
 * @return {Object} 結果
 */
function recordReEntry(reservationId, visitDate) {
  try {
    const visitDatesSheet = getVisitDatesSheet();
    const data = visitDatesSheet.getDataRange().getValues();
    
    let foundRow = -1;
    
    // 予約IDと訪問日で検索
    for (let i = 1; i < data.length; i++) {
      const vid = data[i][0];           // A列: 予約ID
      const vdate = formatDate(data[i][1]); // B列: 訪問日
      
      if (vid === reservationId && vdate === visitDate) {
        foundRow = i + 1;
        break;
      }
    }
    
    if (foundRow === -1) {
      return {
        success: false,
        message: '訪問記録が見つかりません'
      };
    }
    
    const reEntryTime = getNow();
    
    // 再入館時刻を記録
    visitDatesSheet.getRange(foundRow, 7).setValue(reEntryTime); // G列: 再入館時刻
    
    Logger.log(`再入館記録: ${reservationId} (${visitDate}) at ${formatDateTime(reEntryTime)}`);
    
    return {
      success: true,
      message: '再入館を記録しました',
      reEntryTime: formatDateTime(reEntryTime)
    };
    
  } catch (error) {
    Logger.log('再入館記録エラー: ' + error.message);
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.message
    };
  }
}

/**
 * 日付範囲指定統計
 * @param {string} startDate - 開始日（yyyy/MM/dd）
 * @param {string} endDate - 終了日（yyyy/MM/dd）
 * @return {Object} 統計データ
 */
function getStatistics(startDate, endDate) {
  const visitDatesSheet = getVisitDatesSheet();
  const reservationsSheet = getReservationsSheet();
  
  const visitData = visitDatesSheet.getDataRange().getValues();
  const resData = reservationsSheet.getDataRange().getValues();
  
  // 予約データをマップ化
  const reservationMap = {};
  for (let i = 1; i < resData.length; i++) {
    reservationMap[resData[i][0]] = {
      name: resData[i][2],
      purpose: resData[i][6]
    };
  }
  
  const start = new Date(startDate);
  const end = new Date(endDate);
  const today = getToday();
  
  let completedVisits = 0;
  let incompleteVisits = 0;
  let scheduledVisits = 0;
  let cancelledVisits = 0;
  
  const purposeCount = {};
  const dailyCount = {};
  const incompleteRecords = [];
  
  for (let i = 1; i < visitData.length; i++) {
    const visitDate = new Date(visitData[i][1]);
    
    if (visitDate >= start && visitDate <= end) {
      const reservationId = visitData[i][0];
      const status = visitData[i][2];
      const entryTime = visitData[i][5];
      const exitTime = visitData[i][7];
      
      const reservation = reservationMap[reservationId];
      
      // 訪問予定（未来日・有効）
      if (visitDate > today && status === '有効') {
        scheduledVisits++;
        continue;
      }
      
      // キャンセル・未訪問
      if (status === 'キャンセル' || (visitDate <= today && !entryTime && !exitTime)) {
        cancelledVisits++;
        continue;
      }
      
      // 要確認（片方のみ）
      if ((entryTime && !exitTime) || (!entryTime && exitTime)) {
        incompleteVisits++;
        incompleteRecords.push({
          reservationId: reservationId,
          visitDate: formatDate(visitDate),
          name: reservation ? reservation.name : '(不明)',
          entryTime: entryTime ? formatDateTime(entryTime) : '',
          exitTime: exitTime ? formatDateTime(exitTime) : '',
          status: status
        });
        continue;
      }
      
      // 完了訪問
      if (entryTime && exitTime) {
        completedVisits++;
        
        if (reservation) {
          const purpose = reservation.purpose || '不明';
          purposeCount[purpose] = (purposeCount[purpose] || 0) + 1;
        }
        
        const dateStr = formatDate(visitDate);
        dailyCount[dateStr] = (dailyCount[dateStr] || 0) + 1;
      }
    }
  }
  
  const purposeList = Object.keys(purposeCount)
    .map(key => ({ name: key, count: purposeCount[key] }))
    .sort((a, b) => b.count - a.count);
  
  const dailyList = Object.keys(dailyCount)
    .map(key => ({ date: key, count: dailyCount[key] }))
    .sort((a, b) => compareDates(a.date, b.date));
  
  return {
    totalVisits: completedVisits,
    incompleteVisits: incompleteVisits,
    scheduledVisits: scheduledVisits,
    cancelledVisits: cancelledVisits,
    purposeCount: purposeList,
    dailyCount: dailyList,
    incompleteRecords: incompleteRecords
  };
}

/**
 * すべての訪問日が退館済みかチェックし、Reservationsシートを更新
 * @param {string} reservationId - 予約ID
 */
function updateReservationStatusIfAllExited(reservationId) {
  const visitDatesSheet = getVisitDatesSheet();
  const data = visitDatesSheet.getDataRange().getValues();
  
  let allExited = true;
  
  // 同じ予約IDのすべての訪問日をチェック
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      const status = data[i][2]; // C列: ステータス
      if (status !== '退館済' && status !== 'キャンセル') {
        allExited = false;
        break;
      }
    }
  }
  
  // すべて退館済なら、Reservationsシートも更新
  if (allExited) {
    const reservationsSheet = getReservationsSheet();
    const resData = reservationsSheet.getDataRange().getValues();
    
    for (let i = 1; i < resData.length; i++) {
      if (resData[i][0] === reservationId) {
        // I列: ステータスを「退館済」に更新
        reservationsSheet.getRange(i + 1, 9).setValue('退館済');
        // M列: 更新日時を更新
        reservationsSheet.getRange(i + 1, 13).setValue(formatDateTime(getNow()));
        
        Logger.log(`Reservationsシート更新: ${reservationId} → 退館済`);
        break;
      }
    }
  }
}

/**
 * 年度別統計を取得（VisitDatesシート基準・入館退館記録あり限定）
 * @param {number} year - 年度
 * @return {Object} 統計データ
 */
function getStatisticsByYear(year) {
  const visitDatesSheet = getVisitDatesSheet();
  const reservationsSheet = getReservationsSheet();
  
  const visitData = visitDatesSheet.getDataRange().getValues();
  const resData = reservationsSheet.getDataRange().getValues();
  
  // 予約データをマップ化
  const reservationMap = {};
  for (let i = 1; i < resData.length; i++) {
    reservationMap[resData[i][0]] = {
      name: resData[i][2],        // C列: 氏名
      purpose: resData[i][6]      // G列: 訪問目的
    };
  }
  
  // 年度の範囲を計算
  const startDate = new Date(year, 3, 1);      // 4月1日
  const endDate = new Date(year + 1, 2, 31);   // 翌年3月31日
  const today = getToday();
  
  let completedVisits = 0;      // 完了訪問（入館・退館両方あり）
  let incompleteVisits = 0;     // 要確認（片方のみ）
  let scheduledVisits = 0;      // 訪問予定（未来日・有効）
  let cancelledVisits = 0;      // キャンセル・未訪問
  
  const purposeCount = {};
  const dailyCount = {};
  const incompleteRecords = [];  // 要確認レコード
  
  // VisitDatesシートから該当年度のデータを集計
  for (let i = 1; i < visitData.length; i++) {
    const visitDate = new Date(visitData[i][1]); // B列: 訪問日
    
    // 年度範囲内かチェック
    if (visitDate >= startDate && visitDate <= endDate) {
      const reservationId = visitData[i][0];     // A列: 予約ID
      const status = visitData[i][2];            // C列: ステータス
      const entryTime = visitData[i][5];         // F列: 入館時刻
      const exitTime = visitData[i][7];          // H列: 退館時刻
      
      const reservation = reservationMap[reservationId];
      
      // 1. 訪問予定の判定（未来日かつ有効）
      if (visitDate > today && status === '有効') {
        scheduledVisits++;
        continue;
      }
      
      // 2. キャンセル・未訪問の判定
      if (status === 'キャンセル' || (visitDate <= today && !entryTime && !exitTime)) {
        cancelledVisits++;
        continue;
      }
      
      // 3. 要確認の判定（片方のみ記録あり）
      if ((entryTime && !exitTime) || (!entryTime && exitTime)) {
        incompleteVisits++;
        incompleteRecords.push({
          reservationId: reservationId,
          visitDate: formatDate(visitDate),
          name: reservation ? reservation.name : '(不明)',
          entryTime: entryTime ? formatDateTime(entryTime) : '',
          exitTime: exitTime ? formatDateTime(exitTime) : '',
          status: status
        });
        continue;
      }
      
      // 4. 完了訪問の判定（入館・退館両方あり）
      if (entryTime && exitTime) {
        completedVisits++;
        
        if (reservation) {
          // 訪問目的別カウント
          const purpose = reservation.purpose || '不明';
          purposeCount[purpose] = (purposeCount[purpose] || 0) + 1;
        }
        
        // 日別カウント
        const dateStr = formatDate(visitDate);
        dailyCount[dateStr] = (dailyCount[dateStr] || 0) + 1;
      }
    }
  }
  
  // ランキング形式に変換
  const purposeList = Object.keys(purposeCount)
    .map(key => ({ name: key, count: purposeCount[key] }))
    .sort((a, b) => b.count - a.count);
  
  const dailyList = Object.keys(dailyCount)
    .map(key => ({ date: key, count: dailyCount[key] }))
    .sort((a, b) => compareDates(a.date, b.date));
  
  return {
    year: year,
    totalVisits: completedVisits,
    incompleteVisits: incompleteVisits,
    scheduledVisits: scheduledVisits,
    cancelledVisits: cancelledVisits,
    purposeCount: purposeList,
    dailyCount: dailyList,
    incompleteRecords: incompleteRecords
  };
}


/**
 * 予約をキャンセル
 * @param {string} reservationId - 予約ID
 * @return {Object} 結果
 */
function cancelRecord(reservationId) {
  try {
    const reservationsSheet = getReservationsSheet();
    const visitDatesSheet = getVisitDatesSheet();
    
    const data = reservationsSheet.getDataRange().getValues();
    let foundRow = -1;
    
    // 予約IDで検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reservationId) {
        foundRow = i + 1;
        break;
      }
    }
    
    if (foundRow === -1) {
      return {
        success: false,
        message: '予約が見つかりません'
      };
    }
    
    // Reservationsシートのステータスを「キャンセル」に更新
    reservationsSheet.getRange(foundRow, 9).setValue('キャンセル'); // I列: ステータス
    reservationsSheet.getRange(foundRow, 13).setValue(formatDateTime(getNow())); // M列: 更新日時
    
    // VisitDatesシートも更新
    const visitData = visitDatesSheet.getDataRange().getValues();
    for (let i = 1; i < visitData.length; i++) {
      if (visitData[i][0] === reservationId) {
        visitDatesSheet.getRange(i + 1, 3).setValue('キャンセル'); // C列: ステータス
      }
    }
    
    Logger.log(`予約キャンセル: ${reservationId}`);
    
    return {
      success: true,
      message: '予約をキャンセルしました'
    };
    
  } catch (error) {
    Logger.log('キャンセルエラー: ' + error.message);
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.message
    };
  }
}

// ================================
// テスト関数
// ================================

/**
 * 入館中の訪問者取得テスト
 */
function testGetActiveVisitors() {
  Logger.log('=== 入館中訪問者取得テスト ===');
  
  const visitors = getActiveVisitors();
  Logger.log(`入館中の訪問者: ${visitors.length}名`);
  
  visitors.forEach((v, i) => {
    Logger.log(`${i + 1}. ${v.name} (${v.reservationId})`);
  });
  
  Logger.log('\n✓ テスト完了');
}

/**
 * 検索機能テスト
 */
function testSearch() {
  Logger.log('=== 検索機能テスト ===');
  
  const criteria = {
    startDate: '2026/02/01',
    endDate: '2026/02/28',
    name: '',
    purpose: '',
    status: ''
  };
  
  const results = searchVisitRecords(criteria);
  Logger.log(`検索結果: ${results.length}件`);
  
  results.slice(0, 5).forEach((r, i) => {
    Logger.log(`${i + 1}. ${r.name} - ${r.purpose} (${r.status})`);
  });
  
  Logger.log('\n✓ テスト完了');
}

/**
 * 統計機能テスト
 */
function testStatistics() {
  Logger.log('=== 統計機能テスト ===');
  
  const stats = getStatistics('2026/02/01', '2026/02/28');
  
  Logger.log(`総訪問者数: ${stats.totalVisitors}名`);
  Logger.log(`\n訪問目的別:`);
  stats.purposeCount.forEach(p => {
    Logger.log(`  ${p.name}: ${p.count}名`);
  });
  
  Logger.log('\n✓ テスト完了');
}

/**
 * 退館記録のテスト
 */
function testRecordExit() {
  Logger.log('=== 退館記録テスト ===');
  
  // 実際の予約IDに置き換えてください
  const testReservationId = 'RSV20260218227';
  
  // 即時退館
  const result = recordExitNow(testReservationId);
  Logger.log('結果: ' + JSON.stringify(result));
  
  // スプレッドシートのR列に退館時刻が記録されているか確認
}

/**
 * 詳細取得のテスト
 */
function testGetDetail() {
  Logger.log('=== 詳細取得テスト ===');
  
  const testReservationId = 'RSV20251113935';
  const detail = getRecordDetail(testReservationId);
  
  Logger.log('退館時刻: ' + (detail.exitTime || '(未記録)'));
}