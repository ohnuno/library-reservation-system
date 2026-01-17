// ================================
// Webアプリケーション
// ================================

/**
 * WebアプリのGETリクエスト処理
 */
function doGet(e) {
  try {
    Logger.log('=== doGet called ===');
    Logger.log('Parameters: ' + JSON.stringify(e.parameter));
    
    // パラメータを取得
    const reservationId = e.parameter.id || '';
    const token = e.parameter.token || '';
    const action = e.parameter.action || '';
    
    Logger.log('reservationId: ' + reservationId);
    Logger.log('token: ' + token);
    Logger.log('action: ' + action);
    
    // HTMLテンプレートを作成
    const template = HtmlService.createTemplateFromFile('Index');
    
    // パラメータをテンプレートに渡す
    template.reservationId = reservationId;
    template.token = token;
    template.action = action;
    
    // HTMLを評価して返す
    const html = template.evaluate()
      .setTitle('訪問予約確認システム')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    return html;
    
  } catch (error) {
    Logger.log('doGet error: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    
    // エラー時は基本的なHTMLを返す
    return HtmlService.createHtmlOutput(
      '<h1>エラーが発生しました</h1><p>' + error.message + '</p>'
    );
  }
}

/**
 * 予約情報を確認
 * @param {string} reservationId - 予約ID
 * @param {string} token - 予約トークン
 * @return {Object} 予約情報またはエラー
 */
function checkReservation(reservationId, token) {
  try {
    Logger.log('=== checkReservation called ===');
    Logger.log(`予約ID: ${reservationId}`);
    Logger.log(`トークン: ${token}`);
    
    // 予約データを取得
    Logger.log('予約データ取得中...');
    const reservation = getReservationById(reservationId);
    Logger.log('予約データ取得結果: ' + (reservation ? 'あり' : 'なし'));
    
    if (!reservation) {
      Logger.log('エラー: 予約が見つかりません');
      return {
        error: true,
        message: '予約が見つかりません。予約IDを確認してください。'
      };
    }
    
    Logger.log('予約データ: ' + JSON.stringify(reservation));
    
    // トークン検証
    Logger.log('トークン検証中...');
    Logger.log('DBのトークン: ' + reservation.token);
    Logger.log('受信トークン: ' + token);
    
    if (reservation.token !== token) {
      Logger.log('エラー: トークン不一致');
      return {
        error: true,
        message: '無効なQRコードです。'
      };
    }
    
    Logger.log('トークン一致');
    
    // ステータス確認
    Logger.log('ステータス確認: ' + reservation.status);
    if (reservation.status !== '有効') {
      Logger.log('エラー: ステータスが有効ではない');
      return {
        error: true,
        message: 'この予約は' + reservation.status + 'です。'
      };
    }
    
    // 訪問日チェック
    Logger.log('訪問日チェック...');
    const today = formatDate(getToday());
    Logger.log('今日: ' + today);
    
    // visitDatesを文字列に変換
    let visitDatesStr = String(reservation.visitDates || '');
    Logger.log('訪問日リスト(文字列): ' + visitDatesStr);
    
    // カンマで分割
    const visitDates = visitDatesStr.split(',').map(d => d.trim());
    Logger.log('訪問日配列: ' + JSON.stringify(visitDates));
    
    if (!visitDates.includes(today)) {
      Logger.log('エラー: 本日は訪問予定日ではない');
      return {
        error: true,
        message: '本日は訪問予定日ではありません。\n訪問予定日: ' + visitDates.join(', ')
      };
    }
    
    Logger.log('訪問日チェック通過');
    
    // 予約情報を返す
    Logger.log('予約情報を返します');

    // Dateオブジェクトを文字列に変換
    const cleanReservation = {
      reservationId: reservation.reservationId,
      submittedAt: formatDateTime(reservation.submittedAt),  // Dateを文字列に
      name: reservation.name,
      email: reservation.email,
      phone: reservation.phone,
      address: reservation.address,
      purpose: reservation.purpose,
      visitDates: reservation.visitDates,
      status: reservation.status,
      qrCodeUrl: reservation.qrCodeUrl,
      token: reservation.token
    };

    const result = {
      error: false,
      reservationId: reservation.reservationId,
      visitDate: today,
      status: reservation.status,
      reservation: cleanReservation  // 予約データ全体を返す
    };
    
    Logger.log('返却データ: ' + JSON.stringify(result));
    return result;
    
  } catch (error) {
    Logger.log('=== checkReservation エラー ===');
    Logger.log('エラーメッセージ: ' + error.message);
    Logger.log('エラースタック: ' + error.stack);
    
    return {
      error: true,
      message: 'システムエラーが発生しました。職員にお問い合わせください。'
    };
  }
}

/**
 * 受付完了処理
 * @param {string} reservationId - 予約ID
 * @param {string} visitDate - 訪問日
 * @return {Object} 結果
 */
function completeCheckIn(reservationId, visitDate) {
  try {
    Logger.log('=== completeCheckIn called ===');
    Logger.log(`予約ID: ${reservationId}`);
    Logger.log(`訪問日: ${visitDate}`);
    
    // VisitDatesシートで該当する行のステータスを「完了」に更新
    const sheet = getVisitDatesSheet();
    const data = sheet.getDataRange().getValues();
    
    Logger.log(`VisitDatesシート行数: ${data.length}`);
    
    let foundRow = false;
    
    for (let i = 1; i < data.length; i++) {
      Logger.log(`行${i + 1}: 予約ID=${data[i][0]}, 訪問日=${formatDate(data[i][1])}`);
      
      if (data[i][0] === reservationId && formatDate(data[i][1]) === visitDate) {
        Logger.log(`✓ 一致する行を発見: 行${i + 1}`);
        
        // C列(3列目)のステータスを「完了」に更新
        sheet.getRange(i + 1, 3).setValue('完了');
        Logger.log(`C列(ステータス)を「完了」に更新`);
        
        // E列(5列目)のQRコード有効フラグをFALSEに更新
        sheet.getRange(i + 1, 5).setValue('FALSE');
        Logger.log(`E列(QRコード有効フラグ)を「FALSE」に更新`);
        
        foundRow = true;
        break;
      }
    }
    
    if (!foundRow) {
      Logger.log('警告: 該当する行が見つかりませんでした');
      Logger.log(`検索条件: 予約ID=${reservationId}, 訪問日=${visitDate}`);
    }
    
    // すべての訪問日が完了したか確認
    Logger.log('すべての訪問日が完了したか確認中...');
    let allCompleted = true;
    let totalVisitDates = 0;
    let completedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reservationId) {
        totalVisitDates++;
        const status = sheet.getRange(i + 1, 3).getValue(); // 更新後の値を取得
        Logger.log(`行${i + 1}: ステータス=${status}`);
        
        if (status === '完了' || status === 'キャンセル') {
          completedCount++;
        } else {
          allCompleted = false;
        }
      }
    }
    
    Logger.log(`訪問日の総数: ${totalVisitDates}, 完了数: ${completedCount}`);
    Logger.log(`すべて完了: ${allCompleted}`);
    
    // すべて完了していたら、Reservationsシートのステータスも「完了」に更新
    if (allCompleted && totalVisitDates > 0) {
      Logger.log('Reservationsシートを更新します');
      
      const reservationsSheet = getReservationsSheet();
      const reservationsData = reservationsSheet.getDataRange().getValues();
      
      for (let i = 1; i < reservationsData.length; i++) {
        if (reservationsData[i][0] === reservationId) {
          Logger.log(`Reservationsシート: 行${i + 1}を更新`);
          
          // I列(9列目)のステータスを「完了」に更新
          reservationsSheet.getRange(i + 1, 9).setValue('完了');
          Logger.log(`I列(ステータス)を「完了」に更新`);
          
          // M列(13列目)の更新日時を更新
          const now = new Date();
          reservationsSheet.getRange(i + 1, 13).setValue(formatDateTime(now));
          Logger.log(`M列(更新日時)を更新: ${formatDateTime(now)}`);
          
          break;
        }
      }
    }
    
    Logger.log('=== completeCheckIn 完了 ===');
    
    // 予約情報を取得(表示用)
    const reservation = getReservationById(reservationId);
    
    // 残りの訪問日数を計算
    const remainingCount = totalVisitDates - completedCount;
    
    // 処理日時
    const now = new Date();
    const completedAt = formatDateTime(now);
    
    return {
      success: true,
      message: '受付完了しました。',
      reservationId: reservationId,
      name: reservation ? reservation.name : '',
      visitDate: visitDate,
      completedAt: completedAt,
      totalVisitDates: totalVisitDates,
      completedCount: completedCount,
      remainingCount: remainingCount,
      allCompleted: allCompleted
    };
    
  } catch (error) {
    Logger.log('=== completeCheckIn エラー ===');
    Logger.log('エラーメッセージ: ' + error.message);
    Logger.log('エラースタック: ' + error.stack);
    
    return {
      error: true,
      message: 'エラーが発生しました: ' + error.message
    };
  }
}

/**
 * 予約IDから予約情報を取得
 * @param {string} reservationId - 予約ID
 * @return {Object|null} 予約情報
 */
function getReservationById(reservationId) {
  try {
    Logger.log('getReservationById: ' + reservationId);
    
    const sheet = getReservationsSheet();
    const data = sheet.getDataRange().getValues();
    
    Logger.log('データ行数: ' + data.length);
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reservationId) {
        Logger.log('予約発見: 行' + (i + 1));
        
        // 各カラムのデータ型をログ出力
        Logger.log('A列(予約ID): ' + data[i][0] + ' (型: ' + typeof data[i][0] + ')');
        Logger.log('H列(訪問日リスト): ' + data[i][7] + ' (型: ' + typeof data[i][7] + ')');
        Logger.log('K列(トークン): ' + data[i][10] + ' (型: ' + typeof data[i][10] + ')');
        
        // 訪問日リストを文字列に変換
        let visitDatesStr = '';
        if (data[i][7]) {
          // Dateオブジェクトの場合はフォーマット、文字列の場合はそのまま
          if (data[i][7] instanceof Date) {
            visitDatesStr = formatDate(data[i][7]);
          } else {
            visitDatesStr = String(data[i][7]);
          }
        }
        
        Logger.log('変換後の訪問日リスト: ' + visitDatesStr);
        
        return {
          reservationId: String(data[i][0]),
          submittedAt: data[i][1],
          name: String(data[i][2]),
          email: String(data[i][3]),
          phone: String(data[i][4]),
          address: String(data[i][5]),
          purpose: String(data[i][6]),
          visitDates: visitDatesStr,  // 文字列として返す
          status: String(data[i][8]),
          qrCodeUrl: String(data[i][9]),
          token: String(data[i][10])
        };
      }
    }
    
    Logger.log('予約が見つかりませんでした');
    return null;
    
  } catch (error) {
    Logger.log('getReservationByIdエラー: ' + error.message);
    Logger.log('スタック: ' + error.stack);
    throw error;
  }
}

// ================================
// テスト関数
// ================================

/**
 * WebアプリURLのテスト
 */
function testWebAppUrl() {
  Logger.log('=== WebアプリURL確認 ===');
  
  const webAppUrl = getConfigValue('WebアプリURL');
  
  if (webAppUrl) {
    Logger.log('WebアプリURL: ' + webAppUrl);
  } else {
    Logger.log('WebアプリURLが設定されていません。');
    Logger.log('デプロイ後、ConfigシートにWebアプリURLを設定してください。');
  }
}

/**
 * テスト用QRコードを生成
 */
function generateTestQRCode() {
  Logger.log('=== テスト用QRコード生成 ===');
  
  // Reservationsシートから最新の予約を取得
  const sheet = getReservationsSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    Logger.log('予約データがありません。先にフォームから予約を作成してください。');
    return;
  }
  
  // 最新の予約(最後の行)
  const lastRow = data[data.length - 1];
  const reservationId = lastRow[0];
  const token = lastRow[10];
  
  Logger.log('予約ID: ' + reservationId);
  Logger.log('トークン: ' + token);
  
  // WebアプリURL
  const webAppUrl = getConfigValue('WebアプリURL');
  const fullUrl = `${webAppUrl}?id=${reservationId}&token=${token}`;
  
  Logger.log('\nQRコード用URL:');
  Logger.log(fullUrl);
  
  // QuickChart APIでQRコード生成
  const qrCodeUrl = `https://quickchart.io/qr?text=${encodeURIComponent(fullUrl)}&size=300`;
  
  Logger.log('\nQRコード画像URL (ブラウザで開いてテスト):');
  Logger.log(qrCodeUrl);
  
  Logger.log('\n次の作業:');
  Logger.log('1. 上記「QRコード画像URL」をスマートフォンのブラウザで開く');
  Logger.log('2. 表示されたQRコードを別のデバイスで読み取る、または');
  Logger.log('3. 「QRコード用URL」を直接スマートフォンのブラウザで開く');
  Logger.log('4. 組織アカウントでログインする(未ログインの場合)');
  Logger.log('5. 予約情報が表示されることを確認');
}

/**
 * フォーム設定を取得(クライアント側で使用)
 * @return {Array} フォーム質問設定(email, visit_datesを除く)
 */
function getFormQuestionsForClient() {
  const questions = getFormQuestions();
  
  // email, visit_datesは除外
  return questions.filter(q => 
    q.itemId !== 'email' && q.itemId !== 'visit_dates'
  ).map(q => ({
    itemId: q.itemId,
    title: q.title
  }));
}