// ================================
// Google Forms 管理 (スプレッドシート管理方式)
// ================================

/**
 * フォーム質問設定シートを取得
 */
function getFormQuestionsSheet() {
  return getSpreadsheet().getSheetByName('FormQuestions');
}

/**
 * フォーム質問設定を取得
 * @return {Array} 質問設定の配列
 */
function getFormQuestions() {
  // FormQuestionsシートを取得
  const sheet = getFormQuestionsSheet();
  const data = sheet.getDataRange().getValues();
  
  // 各行を質問オブジェクトに変換
  const questions = [];
  for (let i = 1; i < data.length; i++) {
    questions.push({
      order: data[i][0],        // A列: 順序
      itemId: data[i][1],       // B列: 項目ID
      type: data[i][2],         // C列: 質問タイプ
      title: data[i][3],        // D列: 質問タイトル
      helpText: data[i][4],     // E列: ヘルプテキスト
      required: data[i][5],     // F列: 必須
      choices: data[i][6] ? String(data[i][6]).split(',').map(c => c.trim()) : [],  // G列: 選択肢（カンマ区切り）
      section: data[i][7] || '' // H列: セクション名
    });
  }
  
  // 順序でソート
  questions.sort((a, b) => a.order - b.order);
  return questions;
}

/**
 * 予約フォームを作成
 * @return {string} 作成されたフォームのURL
 */
function createReservationForm() {
  // フォームを作成
  const form = FormApp.create('訪問予約申請フォーム');

  // 共有フォルダに移動
  const folderId = getSharedFolderId();
  if (folderId) {
    try {
      const formFile = DriveApp.getFileById(form.getId());
      const folder = DriveApp.getFolderById(folderId);
      const parents = formFile.getParents();
      while (parents.hasNext()) {
        parents.next().removeFile(formFile);
      }
      folder.addFile(formFile);
      Logger.log('フォームを共有フォルダに移動しました');
    } catch (error) {
      Logger.log('警告: フォームの移動に失敗: ' + error.message);
    }
  }
  
  // 質問設定を取得
  const questions = getFormQuestions();
  
  // セクションごとに質問をグループ化
  const sectionMap = {};
  questions.forEach(q => {
    const sectionName = q.section || '基本情報';
    if (!sectionMap[sectionName]) {
      sectionMap[sectionName] = [];
    }
    sectionMap[sectionName].push(q);
  });
  
  // セクションオブジェクトを保持（条件分岐用）
  const sections = {};
  let purposeItem = null;  // 訪問目的の質問項目（条件分岐の起点）
  
  // === 基本情報セクション（最初のページ） ===
  if (sectionMap['基本情報']) {
    // 基本情報の説明をフォーム説明に追加
    const sectionDesc = getSectionDescription('基本情報');
    if (sectionDesc) {
      form.setDescription(
        form.getDescription() + '\n\n' + sectionDesc
      );
    }
    
    // メールアドレスの自動収集を有効化
    form.setCollectEmail(true);
    
    // 基本情報の各質問を追加
    sectionMap['基本情報'].forEach(q => {
      if (q.itemId === 'email') return;
      
      const item = addFormItem(form, q);
      
      // 訪問目的の項目を保存（MultipleChoiceItemとして取得し直す）
      if (q.itemId === 'purpose' && item) {
        // フォームから直接取得し直す
        const items = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
        purposeItem = items[items.length - 1];  // 最後に追加されたラジオボタン項目
      }
    });
  }
  
  // === 見学セクション ===
  if (sectionMap['見学']) {
    // ページ区切りを追加してセクション作成
    const section = form.addPageBreakItem().setTitle('見学');
    const sectionDesc = getSectionDescription('見学');
    if (sectionDesc) section.setHelpText(sectionDesc);
    sections['見学'] = section;
    
    // 見学セクションの質問を追加
    sectionMap['見学'].forEach(q => addFormItem(form, q));
  }
  
  // === 資料閲覧セクション ===
  if (sectionMap['資料閲覧']) {
    const section = form.addPageBreakItem().setTitle('資料閲覧');
    const sectionDesc = getSectionDescription('資料閲覧');
    if (sectionDesc) section.setHelpText(sectionDesc);
    sections['資料閲覧'] = section;
    
    sectionMap['資料閲覧'].forEach(q => addFormItem(form, q));
  }
  
  // === データベース利用セクション ===
  if (sectionMap['データベース利用']) {
    const section = form.addPageBreakItem().setTitle('データベース利用');
    const sectionDesc = getSectionDescription('データベース利用');
    if (sectionDesc) section.setHelpText(sectionDesc);
    sections['データベース利用'] = section;
    
    sectionMap['データベース利用'].forEach(q => addFormItem(form, q));
  }

  // === 条件分岐の設定 ===
  if (purposeItem) {
    const items = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
    purposeItem = items[items.length - 1];
    
    const multipleChoice = purposeItem.asMultipleChoiceItem();
    const choices = [];
    
    if (sections['見学']) {
      choices.push(multipleChoice.createChoice('見学', sections['見学']));
    }
    if (sections['資料閲覧']) {
      choices.push(multipleChoice.createChoice('館内資料閲覧', sections['資料閲覧']));
    }
    if (sections['データベース利用']) {
      choices.push(multipleChoice.createChoice('データベース利用', sections['データベース利用']));
    }
    
    multipleChoice.setChoices(choices);
  }

  // === 各セクションから送信へ遷移 ===
  Object.values(sections).forEach(section => {
    section.setGoToPage(FormApp.PageNavigationType.SUBMIT);
  });
  
  // 確認メッセージを設定
  form.setConfirmationMessage(
    'ご予約ありがとうございます。\n' +
    '登録されたメールアドレスに予約内容とQRコードを送信しました。\n' +
    'メールが届かない場合は、迷惑メールフォルダをご確認ください。'
  );
  
  // フォームIDとURLをConfigシートに保存
  const formId = form.getId();
  const formUrl = form.getPublishedUrl();
  
  Logger.log('フォームが作成されました');
  Logger.log('フォームID: ' + formId);
  Logger.log('フォームURL: ' + formUrl);
  
  saveFormIdToConfig(formId, formUrl);
  
  // フォーム送信トリガーを設定
  Logger.log('\nフォーム送信トリガーを設定中...');
  const triggerResult = setupFormSubmitTrigger(formId);
  if (triggerResult) {
    Logger.log('✓ トリガー設定完了');
  }

  // 営業日更新の日次トリガーを設定
  Logger.log('\n営業日更新トリガーを設定中...');
  setupDailyBusinessDaysUpdateTrigger();
  
  return formUrl;
}

/**
 * フォームに質問項目を追加
 * @param {Form} form - フォームオブジェクト
 * @param {Object} q - 質問設定 {order, itemId, type, title, helpText, required, choices, section}
 * @return {Item} 追加された項目
 */
function addFormItem(form, q) {
  let item = null;
  
  if (q.type === 'text') {
    item = form.addTextItem().setTitle(q.title).setRequired(q.required);
    if (q.helpText) item.setHelpText(q.helpText);
    
  } else if (q.type === 'paragraph') {
    item = form.addParagraphTextItem().setTitle(q.title).setRequired(q.required);
    if (q.helpText) item.setHelpText(q.helpText);
    
  } else if (q.type === 'checkbox') {
    item = form.addCheckboxItem().setTitle(q.title).setRequired(q.required);
    if (q.helpText) item.setHelpText(q.helpText);
    
    if (q.itemId === 'visit_dates') {
      const businessDays = getBusinessDaysForForm();
      item.setChoiceValues(businessDays);
    } else if (q.choices && q.choices.length > 0) {
      item.setChoiceValues(q.choices);
    }
    
  } else if (q.type === 'radio') {
    item = form.addMultipleChoiceItem().setTitle(q.title).setRequired(q.required);
    if (q.helpText) item.setHelpText(q.helpText);

    if (q.itemId !== 'purpose' && q.choices && q.choices.length > 0) {
      item.setChoiceValues(q.choices);
    }
  }
  
  return item;
}

/**
 * フォームIDとURLをConfigシートに保存
 */
function saveFormIdToConfig(formId, formUrl) {
  const sheet = getConfigSheet();
  const data = sheet.getDataRange().getValues();
  
  let formIdRowIndex = -1;
  let formUrlRowIndex = -1;
  
  // 既存の行を探す
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'フォームID') {
      formIdRowIndex = i + 1;
    }
    if (data[i][0] === 'フォームURL') {
      formUrlRowIndex = i + 1;
    }
  }
  
  // フォームIDの保存
  if (formIdRowIndex === -1) {
    sheet.appendRow(['フォームID', formId]);
  } else {
    sheet.getRange(formIdRowIndex, 2).setValue(formId);
  }
  
  // フォームURLの保存
  if (formUrlRowIndex === -1) {
    sheet.appendRow(['フォームURL', formUrl]);
  } else {
    sheet.getRange(formUrlRowIndex, 2).setValue(formUrl);
  }
  
  Logger.log('ConfigシートにフォームID・URLを保存しました');
}

/**
 * 既存のフォームを完全に再構築
 * ※質問の順序や内容が変更された場合に使用
 */
function rebuildForm() {
  const formId = getConfigValue('フォームID');
  
  if (!formId) {
    Logger.log('エラー: フォームIDが見つかりません。先にcreateReservationForm()を実行してください。');
    return;
  }
  
  // 既存フォームを削除
  try {
    const oldForm = FormApp.openById(formId);
    const oldFormFile = DriveApp.getFileById(formId);
    oldFormFile.setTrashed(true);
    Logger.log('既存のフォームを削除しました');
  } catch (e) {
    Logger.log('既存フォームの削除をスキップ: ' + e.message);
  }
  
  // ★★★ 既存のトリガーを削除 ★★★
  removeFormSubmitTriggers();
  
  // 新しいフォームを作成
  Logger.log('新しいフォームを作成します...');
  return createReservationForm();
}

/**
 * 既存のフォームの訪問希望日の選択肢を更新
 */
function updateBusinessDaysInForm() {
  const formId = getConfigValue('フォームID');
  
  if (!formId) {
    Logger.log('エラー: フォームIDが見つかりません');
    return;
  }
  
  const form = FormApp.openById(formId);
  const items = form.getItems(FormApp.ItemType.CHECKBOX);
  
  // 「訪問希望日」のitemIdを持つ質問のタイトルを取得
  const questions = getFormQuestions();
  let visitDatesTitle = '訪問希望日'; // デフォルト
  
  for (let q of questions) {
    if (q.itemId === 'visit_dates') {
      visitDatesTitle = q.title;
      break;
    }
  }
  
  // 該当する質問を探して更新
  for (let i = 0; i < items.length; i++) {
    const checkboxItem = items[i].asCheckboxItem();
    if (checkboxItem.getTitle() === visitDatesTitle) {
      const businessDays = getBusinessDaysForForm();
      checkboxItem.setChoiceValues(businessDays);
      Logger.log('訪問希望日の選択肢を更新しました (' + businessDays.length + '件)');
      return;
    }
  }
  
  Logger.log('エラー: 「' + visitDatesTitle + '」の項目が見つかりません');
}

/**
 * フォームの基本情報を取得
 */
function getFormInfo() {
  const formId = getConfigValue('フォームID');
  
  if (!formId) {
    Logger.log('フォームがまだ作成されていません');
    return null;
  }
  
  const form = FormApp.openById(formId);
  
  return {
    id: formId,
    title: form.getTitle(),
    publishedUrl: form.getPublishedUrl(),
    editUrl: form.getEditUrl(),
    itemCount: form.getItems().length
  };
}

/**
 * フォーム回答から項目IDで値を取得
 * @param {string} itemId - 項目ID
 * @param {Object} formResponse - フォーム回答オブジェクト
 * @return {string|Array} 回答値
 */
function getAnswerByItemId(itemId, formResponse) {
  const questions = getFormQuestions();
  const itemResponses = formResponse.getItemResponses();
  
  // itemIdから質問タイトルを取得
  let targetTitle = null;
  for (let q of questions) {
    if (q.itemId === itemId) {
      targetTitle = q.title;
      break;
    }
  }
  
  if (!targetTitle) {
    return null;
  }
  
  // タイトルで回答を検索
  for (let itemResponse of itemResponses) {
    if (itemResponse.getItem().getTitle() === targetTitle) {
      return itemResponse.getResponse();
    }
  }
  
  // メールアドレスは特別扱い
  if (itemId === 'email') {
    return formResponse.getRespondentEmail();
  }
  
  return null;
}

// ================================
// テスト関数
// ================================

/**
 * FormQuestions設定の確認
 */
function testFormQuestions() {
  Logger.log('=== フォーム質問設定確認 ===');
  const questions = getFormQuestions();
  
  Logger.log('質問数: ' + questions.length);
  Logger.log('');
  
  questions.forEach(q => {
    Logger.log(`[${q.order}] ${q.itemId} (${q.type})`);
    Logger.log(`  タイトル: ${q.title}`);
    Logger.log(`  ヘルプ: ${q.helpText || '(なし)'}`);
    Logger.log(`  必須: ${q.required}`);
    Logger.log('');
  });
}

/**
 * フォーム作成のテスト
 */
function testCreateForm() {
  Logger.log('=== フォーム作成テスト ===');
  const url = createReservationForm();
  Logger.log('フォームURL: ' + url);
  Logger.log('\n次の作業:');
  Logger.log('1. 上記URLをブラウザで開いてフォームを確認してください');
  Logger.log('2. フォームの内容を確認したら、testFormInfo() を実行してください');
}

/**
 * フォーム情報確認のテスト
 */
function testFormInfo() {
  Logger.log('=== フォーム情報確認 ===');
  const info = getFormInfo();
  
  if (info) {
    Logger.log('フォームタイトル: ' + info.title);
    Logger.log('フォームID: ' + info.id);
    Logger.log('公開URL: ' + info.publishedUrl);
    Logger.log('編集URL: ' + info.editUrl);
    Logger.log('質問項目数: ' + info.itemCount);
  }
}

/**
 * 営業日更新のテスト
 */
function testUpdateBusinessDays() {
  Logger.log('=== 訪問希望日の選択肢更新テスト ===');
  updateBusinessDaysInForm();
  Logger.log('完了しました。フォームを開いて確認してください。');
}

/**
 * フォーム再構築のテスト
 */
function testRebuildForm() {
  Logger.log('=== フォーム再構築テスト ===');
  Logger.log('注意: 既存のフォームが削除され、新しいフォームが作成されます。');
  const url = rebuildForm();
  Logger.log('新しいフォームURL: ' + url);
}

/**
 * セクション説明文を取得
 * @param {string} sectionName - セクション名
 * @return {string} 説明文
 */
function getSectionDescription(sectionName) {
  const docId = getConfigValue(`説明文_${sectionName}`);
  if (!docId) return '';
  
  try {
    const doc = DocumentApp.openById(docId);
    return doc.getBody().getText().trim();
  } catch (error) {
    Logger.log(`説明文取得エラー(${sectionName}): ${error.message}`);
    return '';
  }
}