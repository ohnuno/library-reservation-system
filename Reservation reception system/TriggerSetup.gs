// ================================
// トリガー設定管理
// ================================

/**
 * フォーム送信トリガーを設定
 * @param {string} formId - フォームID (省略時はConfigシートから取得)
 */
function setupFormSubmitTrigger(formId = null) {
  // フォームIDが指定されていない場合、Configシートから取得
  if (!formId) {
    formId = getConfigValue('フォームID');
  }
  
  if (!formId) {
    Logger.log('エラー: フォームIDが見つかりません');
    return false;
  }
  
  try {
    // 既存のonFormSubmitトリガーを削除
    removeFormSubmitTriggers();
    
    // 新しいトリガーを作成
    ScriptApp.newTrigger('onFormSubmit')
      .forForm(formId)
      .onFormSubmit()
      .create();
    
    Logger.log('フォーム送信トリガーを設定しました');
    Logger.log('フォームID: ' + formId);
    
    return true;
    
  } catch (error) {
    Logger.log('トリガー設定エラー: ' + error.message);
    return false;
  }
}

/**
 * 営業日更新の日次トリガーを設定
 */
function setupDailyBusinessDaysUpdateTrigger() {
  // 既存の同名トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'updateBusinessDaysInForm') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを作成（毎日午前0時）
  ScriptApp.newTrigger('updateBusinessDaysInForm')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();
  
  Logger.log('営業日更新トリガーを設定しました（毎日0時実行）');
}

/**
 * 変更・キャンセルフォーム送信トリガーを設定
 * @param {string} formId - フォームID
 */
function setupModificationFormTrigger(formId = null) {
  if (!formId) {
    formId = getConfigValue('変更キャンセルフォームID');
  }
  
  if (!formId) {
    Logger.log('エラー: 変更キャンセルフォームIDが見つかりません');
    return false;
  }
  
  try {
    // 既存トリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onModificationFormSubmit') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // 新トリガー作成
    ScriptApp.newTrigger('onModificationFormSubmit')
      .forForm(formId)
      .onFormSubmit()
      .create();
    
    Logger.log('変更・キャンセルフォーム送信トリガー設定完了');
    return true;
  } catch (error) {
    Logger.log('トリガー設定エラー: ' + error.message);
    return false;
  }
}

/**
 * 既存のフォーム送信トリガーを削除
 */
function removeFormSubmitTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let removedCount = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
    }
  });
  
  if (removedCount > 0) {
    Logger.log('既存のトリガーを削除しました: ' + removedCount + '件');
  }
  
  return removedCount;
}

/**
 * すべてのトリガーを表示
 */
function listAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  Logger.log('=== 設定されているトリガー一覧 ===');
  Logger.log('トリガー数: ' + triggers.length);
  
  if (triggers.length === 0) {
    Logger.log('トリガーは設定されていません');
    return;
  }
  
  triggers.forEach((trigger, index) => {
    Logger.log(`\n[${index + 1}]`);
    Logger.log('  関数: ' + trigger.getHandlerFunction());
    Logger.log('  イベント: ' + trigger.getEventType());
    
    // トリガーIDを取得
    try {
      Logger.log('  トリガーID: ' + trigger.getUniqueId());
    } catch (e) {
      // トリガーIDが取得できない場合もある
    }
  });
}

/**
 * すべてのトリガーを削除 (注意: 全削除)
 */
function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  Logger.log('すべてのトリガーを削除します...');
  
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  Logger.log('削除完了: ' + triggers.length + '件');
}

// ================================
// テスト関数
// ================================

/**
 * トリガー設定のテスト
 */
function testSetupTrigger() {
  Logger.log('=== トリガー設定テスト ===');
  
  const result = setupFormSubmitTrigger();
  
  if (result) {
    Logger.log('✓ トリガー設定成功');
    Logger.log('\n設定後のトリガー一覧:');
    listAllTriggers();
  } else {
    Logger.log('✗ トリガー設定失敗');
  }
}