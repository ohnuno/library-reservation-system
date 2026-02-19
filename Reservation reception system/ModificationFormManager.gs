/**
 * 変更・キャンセル依頼フォームを作成
 * @return {string} フォームURL
 */
function createModificationRequestForm() {

  // 既存のフォームを削除
  const existingFormId = getConfigValue('変更キャンセルフォームID');
  if (existingFormId) {
    try {
      const existingForm = FormApp.openById(existingFormId);
      const formFile = DriveApp.getFileById(existingFormId);
      formFile.setTrashed(true);
      Logger.log('既存のフォームを削除しました: ' + existingFormId);
    } catch (error) {
      Logger.log('既存フォーム削除エラー（スキップ）: ' + error.message);
    }
  }

  const form = FormApp.create('訪問予約 変更・キャンセル依頼フォーム');
  
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
      Logger.log('フォームを共有フォルダに移動');
    } catch (error) {
      Logger.log('フォーム移動失敗: ' + error.message);
    }
  }
  
  form.setDescription(
    '訪問予約の変更またはキャンセルを依頼できます。\n' +
    '予約IDとメールアドレスで本人確認を行います。\n' +
    '職員が内容を確認後、処理を行います。'
  );
  
  // メールアドレス自動収集
  form.setCollectEmail(true);
  
  // 予約ID
  form.addTextItem()
    .setTitle('予約ID')
    .setHelpText('予約完了メールに記載されている予約ID（例: RSV20260101001）')
    .setRequired(true);
  
  // 依頼タイプ
  const typeItem = form.addMultipleChoiceItem()
    .setTitle('依頼内容')
    .setRequired(true);
  
  // 変更希望日セクション
  const changeSection = form.addPageBreakItem()
    .setTitle('訪問日変更');
  
  form.addCheckboxItem()
    .setTitle('変更後の訪問希望日')
    .setHelpText('新しい訪問希望日を選択してください（複数選択可）')
    .setChoiceValues(getBusinessDaysForForm())
    .setRequired(true);
  
  // キャンセル理由セクション
  const cancelSection = form.addPageBreakItem()
    .setTitle('予約キャンセル');
  
  form.addParagraphTextItem()
    .setTitle('キャンセル理由')
    .setHelpText('キャンセル理由を教えてください（任意）')
    .setRequired(false);
  
  // 条件分岐設定
  typeItem.setChoices([
    typeItem.createChoice('訪問日変更', changeSection),
    typeItem.createChoice('予約キャンセル', cancelSection)
  ]);
  
  // 各セクションから送信へ
  changeSection.setGoToPage(FormApp.PageNavigationType.SUBMIT);
  cancelSection.setGoToPage(FormApp.PageNavigationType.SUBMIT);
  
  form.setConfirmationMessage(
    'ご依頼ありがとうございます。\n' +
    '職員が確認後、対応いたします。\n' +
    '処理完了後、メールでお知らせします。'
  );
  
  // ConfigシートにフォームIDとURLを保存
  const formId = form.getId();
  const formUrl = form.getPublishedUrl();
  
  saveModificationFormIdToConfig(formId, formUrl);
  
  // トリガー設定
  setupModificationFormTrigger(formId);
  
  Logger.log('変更・キャンセルフォーム作成完了');
  Logger.log('URL: ' + formUrl);
  
  return formUrl;
}

/**
 * ConfigシートにフォームIDとURLを保存
 */
function saveModificationFormIdToConfig(formId, formUrl) {
  const sheet = getConfigSheet();
  const data = sheet.getDataRange().getValues();
  
  let idRow = -1, urlRow = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === '変更キャンセルフォームID') idRow = i + 1;
    if (data[i][0] === '変更キャンセルフォームURL') urlRow = i + 1;
  }
  
  if (idRow === -1) {
    sheet.appendRow(['変更キャンセルフォームID', formId]);
  } else {
    sheet.getRange(idRow, 2).setValue(formId);
  }
  
  if (urlRow === -1) {
    sheet.appendRow(['変更キャンセルフォームURL', formUrl]);
  } else {
    sheet.getRange(urlRow, 2).setValue(formUrl);
  }
  
  Logger.log('ConfigシートにフォームIDとURLを保存');
}