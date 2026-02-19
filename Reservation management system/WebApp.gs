/**
 * Webアプリのエントリーポイント
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('StaffUI')
    .setTitle('学外者入館管理システム - 職員用')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}