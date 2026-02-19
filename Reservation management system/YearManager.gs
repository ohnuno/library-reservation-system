// ================================
// 学外者入館管理システム - 年度管理
// ================================

/**
 * 年度を切り替える（古いデータをアーカイブ）
 * @return {Object} 結果オブジェクト
 */
function switchToNewYear() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const currentYear = getCurrentYear();
    const newYear = currentYear + 1;
    
    Logger.log(`年度切替開始: ${currentYear}年度 → ${newYear}年度`);
    
    // Reservationsシートをアーカイブ
    const reservationsSheet = ss.getSheetByName('Reservations');
    if (!reservationsSheet) {
      throw new Error('Reservationsシートが見つかりません');
    }
    
    const archiveName = 'Reservations_' + currentYear;
    
    // 既に同名のアーカイブが存在する場合はエラー
    if (ss.getSheetByName(archiveName)) {
      throw new Error('アーカイブシート ' + archiveName + ' は既に存在します');
    }
    
    // シートをコピーしてアーカイブ
    const archivedSheet = reservationsSheet.copyTo(ss);
    archivedSheet.setName(archiveName);
    Logger.log(`Reservationsシートをアーカイブ: ${archiveName}`);
    
    // 元のシートをクリア（ヘッダー行は残す）
    const lastRow = reservationsSheet.getLastRow();
    if (lastRow > 1) {
      reservationsSheet.deleteRows(2, lastRow - 1);
      Logger.log(`Reservationsシートをクリア: ${lastRow - 1}行削除`);
    }
    
    // VisitDatesシートをアーカイブ
    const visitDatesSheet = ss.getSheetByName('VisitDates');
    if (visitDatesSheet) {
      const visitArchiveName = 'VisitDates_' + currentYear;
      
      if (!ss.getSheetByName(visitArchiveName)) {
        const visitArchivedSheet = visitDatesSheet.copyTo(ss);
        visitArchivedSheet.setName(visitArchiveName);
        Logger.log(`VisitDatesシートをアーカイブ: ${visitArchiveName}`);
        
        // 元のシートをクリア
        const visitLastRow = visitDatesSheet.getLastRow();
        if (visitLastRow > 1) {
          visitDatesSheet.deleteRows(2, visitLastRow - 1);
          Logger.log(`VisitDatesシートをクリア: ${visitLastRow - 1}行削除`);
        }
      }
    }
    
    // 年度を更新
    updateConfig('CURRENT_YEAR', newYear);
    Logger.log(`年度を更新: ${newYear}`);
    
    Logger.log('年度切替完了');
    
    return {
      success: true,
      oldYear: currentYear,
      newYear: newYear,
      archiveName: archiveName
    };
    
  } catch (error) {
    Logger.log('年度切替エラー: ' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 過去年度のデータを取得
 * @param {number} year - 年度
 * @return {Array} レコードの配列
 */
function getArchivedRecords(year) {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const archiveName = 'Reservations_' + year;
    const archivedSheet = ss.getSheetByName(archiveName);
    
    if (!archivedSheet) {
      Logger.log(`アーカイブシートが見つかりません: ${archiveName}`);
      return [];
    }
    
    const data = archivedSheet.getDataRange().getValues();
    const records = [];
    
    for (let i = 1; i < data.length; i++) {
      records.push({
        reservationId: data[i][0],
        submittedAt: formatDateTime(data[i][1]),
        name: data[i][2],
        email: data[i][3],
        phone: data[i][4],
        affiliation: data[i][5],
        purpose: data[i][6],
        visitDates: data[i][7],
        status: data[i][8],
        createdAt: formatDateTime(data[i][11]),
        updatedAt: formatDateTime(data[i][12]),
        exitTime: data[i][17] ? formatDateTime(data[i][17]) : ''
      });
    }
    
    Logger.log(`${year}年度のアーカイブデータ: ${records.length}件`);
    return records;
    
  } catch (error) {
    Logger.log('アーカイブデータ取得エラー: ' + error.message);
    return [];
  }
}

/**
 * 利用可能な年度一覧を取得
 * @return {Array} 年度の配列（現在年度 + アーカイブされた年度）
 */
function getAvailableYears() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const currentYear = getCurrentYear();
    const years = [currentYear];
    
    // すべてのシート名を取得
    const sheets = ss.getSheets();
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      // "Reservations_YYYY" の形式のシートを探す
      const match = name.match(/^Reservations_(\d{4})$/);
      if (match) {
        const year = parseInt(match[1]);
        if (!years.includes(year)) {
          years.push(year);
        }
      }
    });
    
    // 降順でソート
    years.sort((a, b) => b - a);
    
    Logger.log(`利用可能な年度: ${years.join(', ')}`);
    return years;
    
  } catch (error) {
    Logger.log('年度一覧取得エラー: ' + error.message);
    return [getCurrentYear()];
  }
}

/**
 * アーカイブシートの統計を取得
 * @param {number} year - 年度
 * @return {Object} 統計データ
 */
function getArchiveStatistics(year) {
  try {
    const records = getArchivedRecords(year);
    
    if (records.length === 0) {
      return {
        totalRecords: 0,
        statusCount: {},
        purposeCount: {}
      };
    }
    
    const statusCount = {};
    const purposeCount = {};
    
    records.forEach(record => {
      // ステータス別カウント
      const status = record.status || '不明';
      statusCount[status] = (statusCount[status] || 0) + 1;
      
      // 訪問目的別カウント
      const purpose = record.purpose || '不明';
      purposeCount[purpose] = (purposeCount[purpose] || 0) + 1;
    });
    
    return {
      year: year,
      totalRecords: records.length,
      statusCount: statusCount,
      purposeCount: purposeCount
    };
    
  } catch (error) {
    Logger.log('アーカイブ統計取得エラー: ' + error.message);
    return {
      totalRecords: 0,
      statusCount: {},
      purposeCount: {}
    };
  }
}

/**
 * 年度切替の検証（実際には実行しない）
 * @return {Object} 検証結果
 */
function validateYearSwitch() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const currentYear = getCurrentYear();
    const newYear = currentYear + 1;
    
    // チェック項目
    const checks = {
      reservationsSheetExists: false,
      visitDatesSheetExists: false,
      archiveNameAvailable: false,
      currentDataCount: 0
    };
    
    // Reservationsシートの存在確認
    const reservationsSheet = ss.getSheetByName('Reservations');
    checks.reservationsSheetExists = !!reservationsSheet;
    
    if (reservationsSheet) {
      checks.currentDataCount = reservationsSheet.getLastRow() - 1;
    }
    
    // VisitDatesシートの存在確認
    const visitDatesSheet = ss.getSheetByName('VisitDates');
    checks.visitDatesSheetExists = !!visitDatesSheet;
    
    // アーカイブ名の利用可能性確認
    const archiveName = 'Reservations_' + currentYear;
    checks.archiveNameAvailable = !ss.getSheetByName(archiveName);
    
    const canSwitch = checks.reservationsSheetExists && 
                      checks.visitDatesSheetExists && 
                      checks.archiveNameAvailable;
    
    return {
      success: true,
      canSwitch: canSwitch,
      currentYear: currentYear,
      newYear: newYear,
      checks: checks,
      message: canSwitch ? 
        '年度切替が可能です' : 
        '年度切替の条件が満たされていません'
    };
    
  } catch (error) {
    Logger.log('年度切替検証エラー: ' + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

// ================================
// テスト関数
// ================================

/**
 * 年度切替の検証テスト
 */
function testValidateYearSwitch() {
  Logger.log('=== 年度切替検証テスト ===');
  
  const result = validateYearSwitch();
  
  Logger.log('検証結果:');
  Logger.log('  切替可能: ' + result.canSwitch);
  Logger.log('  現在年度: ' + result.currentYear);
  Logger.log('  新年度: ' + result.newYear);
  Logger.log('  現在データ件数: ' + result.checks.currentDataCount);
  Logger.log('  メッセージ: ' + result.message);
  
  Logger.log('\n✓ テスト完了');
}

/**
 * 利用可能な年度一覧テスト
 */
function testGetAvailableYears() {
  Logger.log('=== 利用可能な年度一覧テスト ===');
  
  const years = getAvailableYears();
  
  Logger.log('利用可能な年度: ' + years.join(', '));
  
  Logger.log('\n✓ テスト完了');
}

/**
 * アーカイブ統計テスト
 */
function testGetArchiveStatistics() {
  Logger.log('=== アーカイブ統計テスト ===');
  
  const years = getAvailableYears();
  
  years.forEach(year => {
    const stats = getArchiveStatistics(year);
    Logger.log(`\n${year}年度:`);
    Logger.log('  総レコード数: ' + stats.totalRecords);
    Logger.log('  ステータス別: ' + JSON.stringify(stats.statusCount));
    Logger.log('  訪問目的別: ' + JSON.stringify(stats.purposeCount));
  });
  
  Logger.log('\n✓ テスト完了');
}

/**
 * 年度切替テスト（実際には実行しない - 検証のみ）
 */
function testSwitchToNewYear() {
  Logger.log('=== 年度切替テスト（検証のみ） ===');
  Logger.log('注意: このテストは実際には年度切替を実行しません');
  
  const validation = validateYearSwitch();
  
  if (validation.canSwitch) {
    Logger.log('\n年度切替が可能です。');
    Logger.log('実際に実行する場合は、switchToNewYear() を手動で実行してください。');
  } else {
    Logger.log('\n年度切替ができません: ' + validation.message);
  }
  
  Logger.log('\n✓ テスト完了');
}