# セットアップガイド

## 前提条件

- Googleアカウント（Google Workspace推奨）
- Google Drive, Spreadsheet, Forms, Calendar, Apps Scriptへのアクセス権限

## セットアップ手順

### 1. Googleスプレッドシートの準備

#### 1-1. 新規スプレッドシートを作成
1. Google Driveで新しいスプレッドシートを作成
2. 名前を「訪問予約管理」などに設定

#### 1-2. 必要なシートを作成
以下の5つのシートを作成してください:

**Config シート**
| A列(設定名) | B列(値) |
|---|---|
| SPREADSHEET_ID | (このスプレッドシートのID) |
| CALENDAR_ID | (Google CalendarのID) |
| STAFF_EMAIL | (職員の通知先メールアドレス) |
| SHARED_FOLDER_ID | (共有フォルダのID、任意) |
| QR_FOLDER_ID | (QRコード保存フォルダID、任意) |

**FormQuestions シート**
| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| 順序 | 項目ID | 質問タイプ | 質問タイトル | ヘルプテキスト | 必須 | 選択肢 | セクション |
| 1 | email | email | メールアドレス | 予約完了メールの送信先 | TRUE | | 基本情報 |
| 2 | name | text | 氏名 | | TRUE | | 基本情報 |
| 3 | phone | text | 電話番号 | | TRUE | | 基本情報 |
| 4 | address | text | 所属機関 | | TRUE | | 基本情報 |
| 5 | visit_dates | checkbox | 訪問希望日 | 複数選択可 | TRUE | | 基本情報 |
| 6 | purpose | radio | 訪問目的 | | TRUE | 見学,館内資料閲覧,データベース利用,その他 | 基本情報 |

**Reservations シート**
ヘッダー行:
| A | B | C | D | E | F | G | H | I | J | K | L | M | N | O | P | Q |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 予約ID | 申請日時 | 氏名 | メール | 電話 | 所属 | 訪問目的 | 訪問日リスト | ステータス | QRコードURL | トークン | 作成日時 | 更新日時 | カレンダーイベントID | 同行人数 | OPAC URL | データベース名称 |

**VisitDates シート**
ヘッダー行:
| A | B | C | D | E |
|---|---|---|---|---|
| 予約ID | 訪問日 | ステータス | カレンダーイベントID | 有効フラグ |

**Calendar シート**
| A | B | C | D |
|---|---|---|---|
| 日付 | 営業日フラグ | 開館時間 | 備考 |
| 2026/01/17 | TRUE | 9:00-20:00 | |
| 2026/01/18 | FALSE | | 休館日 |

### 2. Google Apps Scriptプロジェクト作成（予約受付システム）

#### 2-1. Apps Scriptを開く
1. スプレッドシートから「拡張機能」→「Apps Script」
2. プロジェクト名を「訪問予約システム」に変更

#### 2-2. ファイルを作成
以下のファイルをGitHubリポジトリからコピーして作成:
- Config.gs
- Utils.gs
- BusinessDays.gs
- FormManager.gs
- ReservationHandler.gs
- CalendarManager.gs
- EmailSender.gs
- QRCodeGenerator.gs
- WebApp.gs
- Index.html
- TriggerSetup.gs
- ModificationFormManager.gs
- ModificationRequestHandler.gs

#### 2-3. スクリプトプロパティ設定
「プロジェクトの設定」→「スクリプトプロパティ」で追加:
- `SPREADSHEET_ID`: スプレッドシートのID
- `CALENDAR_ID`: Google CalendarのID
- `BUSINESS_DAYS_SHEET_ID`: スプレッドシートのID（同じ）

#### 2-4. 初期セットアップ実行
1. 関数選択で `initialSetup` を選択
2. 実行して権限を承認
3. `rebuildForm()` を実行してフォーム作成

#### 2-5. Webアプリとしてデプロイ
1. 「デプロイ」→「新しいデプロイ」
2. 種類: ウェブアプリ
3. 次のユーザーとして実行: 自分
4. アクセスできるユーザー: 組織内のユーザー
5. デプロイURLをConfigシートの「WebアプリURL」に保存

### 3. Google Apps Scriptプロジェクト作成（職員用管理システム）

#### 3-1. 新規プロジェクト作成
1. スプレッドシートから「拡張機能」→「Apps Script」→新規プロジェクト
2. プロジェクト名を「訪問予約管理 - 職員用」に変更

#### 3-2. ファイルを作成
以下のファイルをコピー:
- Config.gs
- Utils.gs
- BusinessDays.gs
- ReservationHandler.gs
- CalendarManager.gs
- EmailSender.gs
- StaffWebApp.gs
- StaffInterface.html

#### 3-3. スクリプトプロパティ設定
予約受付システムと同じ内容を設定

#### 3-4. Webアプリとしてデプロイ
1. 「デプロイ」→「新しいデプロイ」
2. 種類: ウェブアプリ
3. アクセスできるユーザー: 組織内のユーザー
4. デプロイURLをConfigシートの「職員用WebアプリURL」に保存

### 4. Googleドキュメント作成（セクション説明文）

各セクションの説明文用に4つのGoogleドキュメントを作成:
1. 基本情報の説明
2. 見学の説明
3. 資料閲覧の説明
4. データベース利用の説明

各ドキュメントIDをConfigシートに保存:
- `説明文_基本情報`
- `説明文_見学`
- `説明文_資料閲覧`
- `説明文_データベース利用`

### 5. 動作確認

1. 予約フォームから予約を送信
2. メール受信を確認
3. QRコードで受付を確認
4. 職員用Webアプリで予約検索を確認

## トラブルシューティング

### フォームが作成されない
- スクリプトプロパティが正しく設定されているか確認
- 権限の承認を確認

### メールが送信されない
- GmailApp.sendEmailの権限を確認
- STAFF_EMAILが正しく設定されているか確認

### カレンダー連携が動作しない
- CALENDAR_IDが正しいか確認
- カレンダーの共有設定を確認