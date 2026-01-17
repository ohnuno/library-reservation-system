# 訪問予約管理システム

図書館向けの訪問予約管理システムです。Google Apps Script (GAS) とGoogleスプレッドシートを使用して実装されています。

## 機能概要

### 利用者向け機能
- **予約申請**: Googleフォームから訪問予約を申請
- **QRコード受付**: 訪問当日、QRコードをスキャンして受付
- **予約変更・キャンセル**: 専用フォームから変更・キャンセル依頼

### 職員向け機能
- **予約管理Webアプリ**: 予約の検索・詳細表示・変更・キャンセル処理
- **カレンダー連携**: Google Calendarに自動登録
- **自動メール送信**: 予約確認・変更通知・キャンセル通知

## システム構成

### 1. 予約受付システム（メインプロジェクト）
- 予約フォーム管理
- QRコード生成・受付処理
- 変更・キャンセル依頼フォーム

### 2. 職員用管理システム（別プロジェクト）
- 予約検索・詳細表示
- 訪問日変更・キャンセル処理

### 3. Googleスプレッドシート
以下のシートで構成:
- **Config**: システム設定
- **FormQuestions**: フォーム質問項目
- **Reservations**: 予約データ
- **VisitDates**: 訪問日詳細
- **Calendar**: 営業日カレンダー

## セットアップ

詳細は [INSTALL.md](INSTALL.md) を参照してください。

## 使用方法

- **利用者向け**: [USER_GUIDE.md](USER_GUIDE.md)
- **職員向け**: [STAFF_GUIDE.md](STAFF_GUIDE.md)

## 技術スタック

- Google Apps Script (JavaScript)
- Google Spreadsheet
- Google Forms
- Google Calendar
- Gmail API
- Google Drive API

## ライセンス

MIT License