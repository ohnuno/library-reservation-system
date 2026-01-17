# システム構成

訪問予約管理システムのファイル構成とアーキテクチャを説明します。

## プロジェクト構成
```
訪問予約管理システム/
├── 予約受付システム（メインプロジェクト）
│   ├── Config.gs              # 設定管理
│   ├── Utils.gs               # 共通ユーティリティ
│   ├── BusinessDays.gs        # 営業日管理
│   ├── FormManager.gs         # フォーム作成・管理
│   ├── ReservationHandler.gs  # 予約データ処理
│   ├── CalendarManager.gs     # カレンダー連携
│   ├── EmailSender.gs         # メール送信
│   ├── QRCodeGenerator.gs     # QRコード生成
│   ├── WebApp.gs              # QRコード受付Webアプリ
│   ├── Index.html             # QRコード受付UI
│   ├── TriggerSetup.gs        # トリガー設定
│   ├── ModificationFormManager.gs      # 変更・キャンセルフォーム管理
│   └── ModificationRequestHandler.gs   # 変更・キャンセル依頼処理
│
├── 職員用管理システム（別プロジェクト）
│   ├── Config.gs              # 設定管理（共通）
│   ├── Utils.gs               # 共通ユーティリティ（共通）
│   ├── BusinessDays.gs        # 営業日管理（共通）
│   ├── ReservationHandler.gs  # 予約データ処理（共通）
│   ├── CalendarManager.gs     # カレンダー連携（共通）
│   ├── EmailSender.gs         # メール送信（共通）
│   ├── StaffWebApp.gs         # 職員用Webアプリ
│   └── StaffInterface.html    # 職員用UI
│
└── Googleスプレッドシート
    ├── Config                 # システム設定
    ├── FormQuestions          # フォーム質問項目
    ├── Reservations           # 予約データ
    ├── VisitDates             # 訪問日詳細
    └── Calendar               # 営業日カレンダー
```

## ファイル詳細

### 予約受付システム

#### Config.gs
**機能**: システム設定の読み書き
- `getConfigValue(key)`: 設定値取得
- `getSpreadsheet()`: スプレッドシート取得
- `getConfigSheet()`: Configシート取得
- `getReservationsSheet()`: Reservationsシート取得
- `getVisitDatesSheet()`: VisitDatesシート取得
- `getFormQuestionsSheet()`: FormQuestionsシート取得

#### Utils.gs
**機能**: 日付フォーマット、ID生成など
- `formatDate(date)`: 日付を"yyyy/MM/dd"形式に
- `formatDateTime(date)`: 日時を"yyyy/MM/dd HH:mm:ss"形式に
- `getToday()`: 今日の日付を取得
- `generateReservationId()`: 予約ID生成
- `generateToken()`: セキュリティトークン生成

#### BusinessDays.gs
**機能**: 営業日カレンダー管理
- `getBusinessDaysList()`: 営業日リスト取得
- `getAvailableBusinessDays()`: 予約可能な営業日取得
- `updateBusinessDaysInForm()`: フォームの営業日選択肢を更新

#### FormManager.gs
**機能**: Googleフォームの作成・管理
- `createReservationForm()`: 予約フォーム作成
- `rebuildForm()`: フォーム再構築
- `getFormQuestions()`: フォーム質問設定取得
- `addFormItem(form, question)`: フォームに質問項目追加
- `getSectionDescription(sectionName)`: セクション説明取得

#### ReservationHandler.gs
**機能**: 予約データの処理
- `onFormSubmit(e)`: フォーム送信時の処理
- `extractReservationData(formResponse)`: フォーム回答からデータ抽出
- `saveToReservationsSheet(data)`: Reservationsシートに保存
- `saveToVisitDatesSheet(data)`: VisitDatesシートに保存
- `getReservationById(id)`: 予約IDで予約取得
- `getAnswerByItemId(itemId, formResponse)`: 質問IDで回答取得

#### CalendarManager.gs
**機能**: Googleカレンダー連携
- `getCalendar()`: カレンダー取得
- `createCalendarEvents(reservation)`: カレンダーイベント作成
- `deleteCalendarEvent(eventId)`: イベント削除
- `createEventDescription(reservation, date)`: イベント説明文作成

#### EmailSender.gs
**機能**: メール送信
- `sendConfirmationEmail(reservation)`: 予約確認メール送信
- `sendStaffNotification(reservation)`: 職員通知メール送信
- `createConfirmationEmailHTML(reservation)`: HTMLメール本文作成
- `createConfirmationEmailPlain(reservation)`: テキストメール本文作成
- `sendCancellationConfirmationEmail(reservation)`: キャンセル確認メール
- `sendVisitDateChangeConfirmation(reservation, removed, added)`: 変更確認メール

#### QRCodeGenerator.gs
**機能**: QRコード生成
- `generateQRCode(reservation)`: QRコード生成
- `generateQRCodeUrl(reservationId, token)`: QRコードURL生成
- `saveQRCodeToReservation(reservationId, qrCodeUrl)`: QRコードURLを保存

#### WebApp.gs
**機能**: QRコード受付Webアプリ
- `doGet(e)`: GET リクエスト処理
- `checkReservation(id, token)`: 予約確認
- `completeCheckIn(id, visitDate)`: 受付完了処理
- `getFormQuestionsForClient()`: クライアント用質問設定取得

#### Index.html
**機能**: QRコード受付UI
- 予約情報表示
- 受付完了ボタン
- エラー表示

#### TriggerSetup.gs
**機能**: トリガー設定
- `setupFormSubmitTrigger(formId)`: フォーム送信トリガー設定
- `setupModificationFormTrigger(formId)`: 変更フォーム送信トリガー設定
- `setupDailyBusinessDaysUpdateTrigger()`: 営業日更新トリガー設定

#### ModificationFormManager.gs
**機能**: 変更・キャンセル依頼フォーム管理
- `createModificationRequestForm()`: 変更・キャンセルフォーム作成
- `saveModificationFormIdToConfig(formId, formUrl)`: フォームID保存

#### ModificationRequestHandler.gs
**機能**: 変更・キャンセル依頼処理
- `onModificationFormSubmit(e)`: フォーム送信時処理
- `sendStaffModificationRequest(data, reservation)`: 職員通知
- `sendModificationRequestConfirmation(data)`: 受付確認メール
- `sendModificationErrorEmail(email, message)`: エラーメール

### 職員用管理システム

#### StaffWebApp.gs
**機能**: 職員用管理Webアプリ
- `doGetStaff(e)`: GET リクエスト処理
- `searchReservations(type, value)`: 予約検索
- `getReservationDetail(id)`: 予約詳細取得
- `modifyVisitDates(id, remove, add)`: 訪問日変更
- `cancelReservation(id)`: 予約キャンセル
- `getBusinessDaysForForm()`: 営業日リスト取得（受付期限なし）
- `createModifiedCalendarEvent(...)`: 変更履歴付きイベント作成
- `updateReservationVisitDates(id)`: 訪問日リスト更新

#### StaffInterface.html
**機能**: 職員用UI
- 予約検索フォーム
- 検索結果表示
- 予約詳細表示
- 訪問日変更フォーム
- キャンセルボタン

## データフロー

### 予約申請フロー
```
利用者 → Googleフォーム → onFormSubmit
  ↓
extractReservationData → saveToReservationsSheet
  ↓                       ↓
  └─→ saveToVisitDatesSheet
  ↓
generateQRCode → createCalendarEvents → sendConfirmationEmail
```

### QRコード受付フロー
```
職員 → QRコードスキャン → WebApp(doGet)
  ↓
checkReservation → 予約情報表示
  ↓
受付完了ボタン → completeCheckIn
  ↓
VisitDatesシート更新 → Reservationsシート更新
```

### 変更・キャンセル依頼フロー
```
利用者 → 変更・キャンセルフォーム → onModificationFormSubmit
  ↓
本人確認 → sendStaffModificationRequest
  ↓            ↓
  └─→ sendModificationRequestConfirmation
         ↓
職員 → 管理Webアプリ → modifyVisitDates / cancelReservation
  ↓
データ更新 → カレンダー更新 → 確認メール送信
```

## 技術仕様

### 予約ID形式
`RSV + yyyyMMdd + 3桁連番`
例: RSV20260117001

### トークン
32文字のランダム英数字

### QRコードURL形式
```
https://script.google.com/a/macros/組織ドメイン/s/デプロイID/exec?id=[予約ID]&token=[トークン]
```

### ステータス
- **有効**: 予約確定、訪問前
- **完了**: すべての訪問日の受付完了
- **キャンセル**: 予約キャンセル

### トリガー
1. **フォーム送信トリガー**: 予約フォーム送信時
2. **変更フォーム送信トリガー**: 変更・キャンセルフォーム送信時
3. **日次トリガー**: 毎日0時、営業日更新

## セキュリティ

- トークンによる予約の検証
- 組織内ユーザーのみWebアプリにアクセス可能
- スクリプトプロパティで機密情報を管理