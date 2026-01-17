// ================================
// QRコード生成
// ================================

/**
 * QRコードを生成(メール埋め込み用とバックアップ用)
 * @param {string} reservationId - 予約ID
 * @param {string} token - 予約トークン
 * @return {Object} {emailImageUrl: メール用URL, backupUrl: バックアップファイルURL, backupFileId: ファイルID}
 */
function generateAndSaveQRCode(reservationId, token) {
  try {
    // WebアプリのURLを取得 (後で設定)
    let webAppUrl = getConfigValue('WebアプリURL');
    
    // WebアプリURLが未設定の場合は仮のURLを使用
    if (!webAppUrl) {
      webAppUrl = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
      Logger.log('警告: WebアプリURLが未設定です。仮のURLを使用します。');
    }
    
    // QRコードの内容 (WebアプリURL + パラメータ)
    const qrContent = `${webAppUrl}?id=${reservationId}&token=${token}`;
    
    Logger.log('QRコード内容: ' + qrContent);
    
    // メール埋め込み用: QuickChart APIのURL (外部からアクセス可能)
    const emailImageUrl = generateQRCodeImageUrl(qrContent);
    Logger.log('メール用QRコード画像URL: ' + emailImageUrl);
    
    // バックアップ用: Googleドライブに保存 (職員確認用)
    let backupUrl = null;
    let backupFileId = null;
    
    try {
      const response = UrlFetchApp.fetch(emailImageUrl);
      const blob = response.getBlob();
      blob.setName(`QR_${reservationId}.png`);
      
      const file = saveQRCodeToFolder(blob);
      backupUrl = file.getUrl();
      backupFileId = file.getId();
      
      Logger.log('バックアップ用QRコード画像を保存しました');
      Logger.log('バックアップURL: ' + backupUrl);
      
      // 共有設定を試みる (エラーが出ても処理を続行)
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        Logger.log('ファイルの共有設定を変更しました');
      } catch (shareError) {
        Logger.log('共有設定の変更をスキップ: ' + shareError.message);
        Logger.log('(フォルダの共有設定が継承されます)');
      }
      
    } catch (backupError) {
      Logger.log('警告: バックアップ保存に失敗しましたが、処理を続行します');
      Logger.log('エラー内容: ' + backupError.message);
    }
    
    return {
      emailImageUrl: emailImageUrl,        // メール埋め込み用(外部アクセス可能)
      backupUrl: backupUrl,                 // バックアップURL(Googleドライブ)
      backupFileId: backupFileId,           // バックアップファイルID
      qrContent: qrContent                  // QRコードに埋め込まれた内容
    };
    
  } catch (error) {
    Logger.log('QRコード生成エラー: ' + error.message);
    throw error;
  }
}

/**
 * QRコード画像のURLを生成 (QuickChart API使用)
 * @param {string} content - QRコードに埋め込む内容
 * @return {string} QRコード画像のURL
 */
function generateQRCodeImageUrl(content) {
  // QuickChart.io のQRコード生成API
  // サイズ: 300x300ピクセル
  const encodedContent = encodeURIComponent(content);
  return `https://quickchart.io/qr?text=${encodedContent}&size=300`;
}

/**
 * QRコード画像を適切なフォルダに保存
 * @param {Blob} blob - 画像データ
 * @return {File} 保存されたファイル
 */
function saveQRCodeToFolder(blob) {
  const folderId = getSharedFolderId();
  
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      
      // QRコード用のサブフォルダを作成または取得
      let qrFolder = null;
      const subFolders = folder.getFoldersByName('QRコード');
      
      if (subFolders.hasNext()) {
        qrFolder = subFolders.next();
      } else {
        qrFolder = folder.createFolder('QRコード');
        Logger.log('QRコード用フォルダを作成しました');
      }
      
      return qrFolder.createFile(blob);
      
    } catch (error) {
      Logger.log('共有フォルダへの保存に失敗: ' + error.message);
      Logger.log('マイドライブに保存します');
    }
  }
  
  // フォルダIDが未設定、またはエラーの場合はマイドライブに保存
  return DriveApp.createFile(blob);
}

/**
 * 予約データにQRコード情報を追加
 * @param {string} reservationId - 予約ID
 * @param {string} qrCodeUrl - QRコード画像のURL
 */
function updateQRCodeInReservation(reservationId, qrCodeUrl) {
  const sheet = getReservationsSheet();
  const data = sheet.getDataRange().getValues();
  
  // 予約IDで検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservationId) {
      // J列(10列目)にQRコードURLを保存
      sheet.getRange(i + 1, 10).setValue(qrCodeUrl);
      Logger.log('QRコードURLを更新: 行' + (i + 1));
      return true;
    }
  }
  
  Logger.log('警告: 予約ID ' + reservationId + ' が見つかりません');
  return false;
}

/**
 * QRコード画像を削除
 * @param {string} fileId - GoogleドライブのファイルID
 */
function deleteQRCodeFile(fileId) {
  if (!fileId) {
    Logger.log('ファイルIDが指定されていません');
    return false;
  }
  
  try {
    const file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
    Logger.log('QRコードファイルを削除しました: ' + fileId);
    return true;
  } catch (error) {
    Logger.log('QRコードファイルの削除に失敗: ' + error.message);
    return false;
  }
}

// ================================
// テスト関数
// ================================

/**
 * QRコード生成のテスト
 */
function testQRCodeGeneration() {
  Logger.log('=== QRコード生成テスト ===');
  
  const testReservationId = 'RSV20251113TEST';
  const testToken = generateToken();
  
  Logger.log('テスト予約ID: ' + testReservationId);
  Logger.log('テストトークン: ' + testToken);
  
  try {
    const result = generateAndSaveQRCode(testReservationId, testToken);
    
    Logger.log('\n✓ QRコード生成成功');
    Logger.log('\n【メール埋め込み用URL】(外部からアクセス可能)');
    Logger.log(result.emailImageUrl);
    Logger.log('\n【バックアップURL】(職員確認用・Googleドライブ)');
    Logger.log(result.backupUrl || '(保存されませんでした)');
    Logger.log('\n【QRコードの内容】');
    Logger.log(result.qrContent);
    Logger.log('\n次の作業:');
    Logger.log('1. 上記の「メール埋め込み用URL」をブラウザで開いてQRコードを確認してください');
    Logger.log('2. スマートフォンのQRコードリーダーでスキャンして内容を確認してください');
    Logger.log('3. バックアップURLがあれば、Googleドライブで確認してください');
    
  } catch (error) {
    Logger.log('\n✗ QRコード生成失敗');
    Logger.log('エラー: ' + error.message);
    Logger.log('スタックトレース: ' + error.stack);
  }
}

/**
 * QRコードURL生成のテスト
 */
function testQRCodeUrl() {
  Logger.log('=== QRコードURL生成テスト ===');
  
  const testContent = 'https://example.com/check?id=TEST123&token=ABC456';
  const url = generateQRCodeImageUrl(testContent);
  
  Logger.log('テスト内容: ' + testContent);
  Logger.log('生成されたURL: ' + url);
  Logger.log('\n次の作業:');
  Logger.log('上記URLをブラウザで開いて、QRコード画像が表示されることを確認してください');
}

/**
 * メール用HTML生成のテスト
 */
function testQRCodeEmailHTML() {
  Logger.log('=== メール用HTML生成テスト ===');
  
  const testReservationId = 'RSV20251113TEST';
  const testToken = generateToken();
  const result = generateAndSaveQRCode(testReservationId, testToken);
  
  const htmlSample = `
<html>
<body>
  <h2>訪問予約を受け付けました</h2>
  <p>訪問当日は、下記のQRコードを受付にてご提示ください。</p>
  <p style="text-align: center;">
    <img src="${result.emailImageUrl}" alt="QRコード" style="width: 300px; height: 300px;">
  </p>
</body>
</html>
`;
  
  Logger.log('生成されたHTML:');
  Logger.log(htmlSample);
  Logger.log('\nこのHTMLをメールで送信すると、QRコードが表示されます。');
}