# Google Drive クローラー仕様書 (Google Apps Script版)

## 1. 概要

本システムは、Google Apps Script (GAS) を使用し、指定されたGoogle Driveのフォルダ（共有ドライブを含む）配下のすべてのフォルダとファイルを再帰的にクロールし、指定された情報をスプレッドシートに出力する。

GWSの実行時間制限（最大30分）を考慮し、処理を自動で中断・再開（レジューム）する機能を搭載する。

## 2. 目的

*   特定のドライブ配下にある全ファイルの棚卸し（一覧化）。
*   ファイルのメタデータ（更新日時、オーナー、ファイルタイプ等）の収集。
*   （将来的な拡張）特定のファイル（例：スプレッドシート）の内容を解析するバッチ処理の基盤とする。

## 3. システム構成

*   **実行環境:** Google Apps Script (GWS)
*   **コア技術:**
    
    *   Google Drive API v3 (ファイルのリスト取得、変更履歴の取得)
    *   Spreadsheet Service (設定の読み取り、結果の書き込み)
    *   Properties Service (レジューム用データ（処理キュー）の保存)
*   **インターフェース:** 実行管理用のGoogleスプレッドシート（カスタムメニュー）

## 4. 主な機能

1.  **設定機能:**
    
    *   起点となるフォルダID、結果出力シート名などをスプレッドシートの「設定」シートで管理する。
2.  **フルクロール機能:**
    
    *   起点フォルダから全階層をクロールし、ファイル情報と**全フォルダIDリスト**を取得・出力する。
3.  **更新確認機能 (差分取得):**
    
    *   フルクロールとは別に実行。Drive APIの `Changes` API を使用し、前回チェック以降に変更（追加・更新・削除）があったアイテムのみを取得する。
4.  **レジューム機能:**
    
    *   処理が実行時間制限（安全マージンを取り25分）に達した場合、未処理のフォルダキューを `PropertiesService` に保存し、自動的に処理を中断する。
    *   次回実行時に保存されたキューを読み込み、中断箇所から処理を再開する。
5.  **リセット機能:**
    
    *   保存されたレジュームデータ（キュー）をクリアし、次回実行時に最初からクロールをやり直す。（※各種結果シートも手動でクリアする必要がある）

## 5. 機能詳細

### 5.1. 設定シート (`設定`)

スプレッドシートに「設定」という名前のシートを作成し、以下の項目を設ける。

| セル  | 項目名          | 説明                                            |
| --- | ------------ | --------------------------------------------- |
| B1  | 起点フォルダID     | クロールを開始する最上位のフォルダID（またはURL）。                  |
| B2  | フルクロール結果シート名 | 全ファイル情報を出力するシート名。（例: FullCrawl_List）          |
| B3  | 更新確認結果シート名   | 差分更新情報を出力するシート名。（例: Update_Log）               |
| B4  | 実行ステータス      | スクリプトの実行状況（処理中、中断、完了）をリアルタイムで出力する。            |
| B5  | 全フォルダリストシート名 | クロール対象配下の全フォルダIDを出力・参照するシート名。（例: Folder_List） |
| B6  | エラーログシート名    | 実行中のエラー（権限不足等）を記録するシート名。（例: Error_Log）        |

### 5.2. フルクロール機能 (main\_FullCrawl)

キュー方式による再帰的クロール。

#### 5.2.1. 処理フロー

1.  **初期化:**
    
    *   `SCRIPT_START_TIME` に現在時刻をセット。
    *   `getConfiguration()` で「設定」シートから情報を読み込む。
    *   **レジュームデータ読み込み:**
        
        *   `PropertiesService` からレジュームデータ (`FOLDER_QUEUE`) を読み込む。
    *   **処理済みリスト構築 (容量制限対策):**
        
        *   `フルクロール結果シート名`(B2)のシートを取得。A列（ファイルID列）の既存データをすべて読み込み、処理済みファイルIDの `Set` (`processedFilesSet`) をメモリ上に構築する。
        *   `全フォルダリストシート名`(B5)のシートを取得。A列（フォルダID列）の既存データをすべて読み込み、処理済みフォルダIDの `Set` (`allFolderIdSet`) をメモリ上に構築する。
    *   **書き込み待機配列の初期化:**
        
        *   `filesToWrite = []` (ファイル情報用)
        *   `foldersToWrite = []` (フォルダ情報用)
2.  **キューの準備:**
    
    *   `FOLDER_QUEUE` が空（初回実行）の場合:
        
        *   起点フォルダID(B1)の情報を `Drive.Files.get` で取得する（`id`, `name`）。（失敗時は `5.5. エラーハンドリング` に基づき終了）
        *   取得したフォルダ情報（例: `startFolder`）を `FOLDER_QUEUE` に追加する (`{id: startFolder.id, name: startFolder.name}`)。
        *   **起点フォルダをリストに追加:**
            
            *   `allFolderIdSet.add(startFolder.id)` を実行。
            *   `foldersToWrite` 配列に起点フォルダのIDと名前を追加。
    *   `FOLDER_QUEUE` が存在する場合（再開時）、「処理再開」としてステータスセル(B4)を更新。
3.  **メインループ (while `FOLDER_QUEUE.length > 0`):**
    
    *   **時間チェック:** `checkTimeLimit()` (実行時間が25分を超えていないか) を確認。超えていればループを `break`。
    *   **キュー取得:** `FOLDER_QUEUE` から先頭のフォルダ (`currentFolder`) を取り出す (`shift()`)。
    *   **ステータス更新:** 「処理中: {フォルダ名} (残り: X件)」をB4セルに書き込む。
    *   **アイテム一覧取得 (Drive API v3):**
        
        *   `5.4.1. 指数バックオフ` を使用し、`Drive.Files.list` を呼び出す。
        *   `q` パラメータ: `'${currentFolder.id}' in parents and trashed=false`
        *   `fields`: `nextPageToken, files(id, name, mimeType, parents, createdTime, modifiedTime, owners)`
        *   `supportsAllDrives: true`, `includeItemsFromAllDrives: true` を指定。
        *   **エラー処理:** `5.5. エラーハンドリング` に基づき、フォルダ取得失敗時はログに残し、次のフォルダへ進む。
    *   **アイテム処理 (for `item` of `items`):**
        
        *   **時間チェック:** `checkTimeLimit()` を確認。時間切れなら `currentFolder` をキューに戻し (`unshift()`)、ループを `break`。
        *   **フォルダの場合 (`mimeType === FOLDER`):**
            
            *   `FOLDER_QUEUE` の末尾に追加 (`push({ id: item.id, name: item.name })`)。
            *   `allFolderIdSet.has(item.id)` で重複チェック。
            *   未処理の場合、`allFolderIdSet.add(item.id)` を実行し、`foldersToWrite` 配列に追加。
        *   **ファイルの場合:**
            
            *   `processedFilesSet.has(item.id)` で重複チェック。
            *   未処理の場合、取得したメタ情報（5.2.2 参照）を `filesToWrite` 配列に追加。
            *   `processedFilesSet.add(item.id)` を実行。
    *   **結果書き込み (フォルダごと):**
        
        *   `for item` ループが正常に完了した後（＝1フォルダの処理完了後）、`filesToWrite` と `foldersToWrite` にデータがあれば、対象シートに一括書き込み（`appendRows`）し、両配列を空にする。
4.  **終了処理:**
    
    *   **中断時:**
        
        *   `filesToWrite` と `foldersToWrite` に残データがあれば、対象シートに一括書き込みする（中断直前のデータを失わないため）。
        *   `FOLDER_QUEUE` のみ\*\*を `PropertiesService` に保存する。ステータスセル(B4)に「中断」メッセージをセット。
    *   **完了時:**
        
        *   ステータスセル(B4)に「完了」メッセージをセット。
        *   `PropertiesService` の `FOLDER_QUEUE` データを削除 (`deleteProperty`) する。
        *   **`5.3.1. 前提 (トークンの保存)` の処理を実行する。**

#### 5.2.2. 取得する情報 (フルクロール結果シート出力項目)

| 列   | 項目名            | 取得元 (Drive API v3 filesリソース)                                  |
| --- | -------------- | ------------------------------------------------------------- |
| A   | ファイルID         | id                                                            |
| B   | ファイル名          | name                                                          |
| C   | ファイルURL        | https://docs.google.com/spreadsheets/d/{id} (MIME Typeに応じて分岐) |
| D   | MIME Type      | mimeType                                                      |
| E   | 親フォルダID (全て)   | item.parents.join(', ')                                       |
| F   | 発見時の親フォルダ名     | クロール時の currentFolder.name                                     |
| G   | 作成日時           | createdTime                                                   |
| H   | 最終更新日時         | modifiedTime                                                  |
| I   | オーナー (メールアドレス) | owners[0].emailAddress                                        |
| J   | 取得日時           | (GAS実行時の new Date())                                          |

### 5.3. 更新確認機能 (main\_CheckUpdates)

フルクロールとは別に実行する。Drive API v3 の `Changes` API を利用する。

#### 5.3.1. 前提 (トークンの保存)

*   **`main_FullCrawl` が正常に完了した時点**、または**リセット機能の実行時**に、`Drive.Changes.getStartPageToken` を呼び出し、最新の `startPageToken` を取得。
*   取得した `startPageToken` を `PropertiesService` に `START_PAGE_TOKEN` キーで保存する。

#### 5.3.2. 処理フロー

1.  **初期化:**
    
    *   `PropertiesService` から `START_PAGE_TOKEN` を読み込む。トークンがなければエラーとし、先にフルクロールを実行するよう促す。
    *   「設定」シートから `更新確認結果シート名` (B3)、`全フォルダリストシート名` (B5)、`フルクロール結果シート名` (B2) を取得。
    *   `全フォルダリストシート名`(B5)のシートからA列の全フォルダIDを読み込み、`allFolderIdSet` (`Set`) をメモリ上に構築する。
    *   `フルクロール結果シート名`(B2)のシートからA列の全ファイルIDを読み込み、`processedFilesSet` (`Set`) をメモリ上に構築する。
2.  **変更履歴の取得 (Drive API v3):**
    
    *   `pageToken` に `PropertiesService` から読み込んだ `START_PAGE_TOKEN` をセットする。
    *   `5.4.1. 指数バックオフ` を使用し、`Drive.Changes.list` を呼び出す。
    *   `fields`: `nextPageToken, newStartPageToken, changes(fileId, time, removed, file(id, name, mimeType, parents, createdTime, modifiedTime, owners))`
    *   `supportsAllDrives: true`, `includeItemsFromAllDrives: true` を指定。
3.  **ループ処理 (while `pageToken`):**
    
    *   **中断フラグ:** `wasInterrupted = false` を初期化。
    *   **時間チェック:** `checkTimeLimit()` を確認。超えていれば `wasInterrupted = true` をセットし、ループを `break`。
    *   `Drive.Changes.list` を `pageToken` を更新しながら実行。
    *   **`Invalid Value` 自動回復:**
        
        *   `Invalid Value` エラーを `catch` した場合、`saveStartPageToken` を呼び出してトークンを再取得する。
        *   **APIキャッシュ問題:** 再取得した `newToken` が `pageToken` と同じ場合、`wasInterrupted = true` をセットし、`break` する。
        *   **回復成功:** `newToken` が異なる場合、`pageToken` を `newToken` に更新し、`response` をリトライ取得する。
    *   **変更アイテム処理 (for `change` of `changes`):**
        
        *   `change.removed === true` の場合:
            
            *   **ファイル削除の検知:**
                
                *   `processedFilesSet.has(change.fileId)` をチェックし、管理対象のファイルだった場合、「削除」としてログ出力(B3)。(ファイルID, 変更日時, "削除")
                *   `processedFilesSet.delete(change.fileId)` を実行。（`FullCrawl_List` シート(B2)からの削除は行わない）
            *   **フォルダ削除の検知:**
                
                *   `allFolderIdSet.has(change.fileId)` をチェックし、管理対象のフォルダだった場合、「削除」としてログ出力(B3)。(フォルダID, 変更日時, "削除 \[フォルダ\]")
                *   `allFolderIdSet.delete(change.fileId)` を実行。
                *   `Folder_List` シート (B5) からのリアルタイム行削除は行わない。
        *   `change.removed === false` の場合: `change.file` (Fileリソース) が存在。
            
            *   **管理対象かチェック:** `change.file.parents` 配列のいずれかのIDが、`allFolderIdSet` に含まれているかチェックする (`isTargetParent`)。
            *   `isTargetParent` が `true` の場合（＝管理対象フォルダ配下の変更）:
                
                *   **フォルダの場合 (`change.file.mimeType === FOLDER`):**
                    
                    *   `allFolderIdSet.has(change.file.id)` をチェック。
                    *   `true` (更新) の場合: 「更新」としてログ出力(B3)。
                    *   `false` (新規追加) の場合: 「追加」としてログ出力(B3)。
                    *   `allFolderIdSet.add(change.file.id)` を実行し、メモリ上のリストを更新。
                    *   `Folder_List` シート (B5) に新しいフォルダIDと名前を追記する。
                *   **ファイルの場合:**
                    
                    *   `processedFilesSet.has(change.file.id)` をチェック。
                    *   `true` (更新) の場合: 「更新」としてログ出力(B3)。
                    *   （`FullCrawl_List` (B2) の既存行の更新はログ記録のみ）
                    *   `false` (新規追加) の場合: 「追加」としてログ出力(B3)。
                    *   `processedFilesSet.add(change.file.id)` を実行し、メモリ上のリストを更新。
                    *   `FullCrawl_List` シート (B2) に、`change.file` の情報（5.2.2の項目）を新しい行として追記する。
    *   **トークン更新:** ループの最後に `newStartPageToken` を取得。
4.  **終了処理:**
    
    *   **中断時 (`pageToken` が `true` OR `wasInterrupted` が `true` の場合):**
        
        *   `START_PAGE_TOKEN` を**更新せず**に終了する。次回も同じ `startPageToken` から処理が再開される。
    *   **正常完了時 (`pageToken` が `null` AND `wasInterrupted` が `false` の場合):**
        
        *   `PropertiesService` の `START_PAGE_TOKEN` を、最後に取得した `newStartPageToken` で上書き保存する。

### 5.4. ユーティリティ関数（共通処理）

#### 5.4.1. 指数バックオフ (exponentialBackoff)

*   Drive APIの呼び出し（`Drive.Files.list` など）をラップする関数。
*   API呼び出しが失敗した場合、`try-catch` でエラーを捕捉する。
*   エラーがレートリミット（`403 rateLimitExceeded` や `5xx` サーバーエラー）に起因する場合、`Utilities.sleep()` を使用して待機（例: 1s, 2s, 4s...）し、最大5回程度リトライする。
*   最大回数リトライしても失敗した場合、エラーをスローし、呼び出し元の `5.5. エラーハンドリング` 処理に委ねる。

### 5.5. エラーハンドリング

*   **エラーログシート:**
    
    *   `設定` シート (B6) で指定された `エラーログシート名` にエラーを記録する。
    *   シートが存在しない場合は、ヘッダー（`日時`, `発生箇所`, `対象ID`, `エラーメッセージ`）を付与して作成する。
*   **処理:**
    
    *   `Drive.Files.list` や `Drive.Changes.list` が（指数バックオフ後も）失敗した場合（例: フォルダへのアクセス権限がない）、そのフォルダIDとエラーメッセージをエラーログシートに記録し、そのフォルダの処理はスキップして次のキューに進む。
    *   起点フォルダ(B1)の取得に失敗した場合、処理続行不可能なため、ステータスセル(B4)にエラーを書き込み、スクリプトを終了する。

## 6. GASプロジェクト設定

### 6.1. 必要なサービス

*   Google Drive API v3 (v2ではなくv3を明示的に有効化する)

### 6.2. `appsscript.json` (マニフェストファイル)

```
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {
    "enabledAdvancedServices": [{
      "userSymbol": "Drive",
      "serviceId": "drive",
      "version": "v3"
    }]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "[https://www.googleapis.com/auth/script.scriptapp](https://www.googleapis.com/auth/script.scriptapp)",
    "[https://www.googleapis.com/auth/script.storage](https://www.googleapis.com/auth/script.storage)",
    "[https://www.googleapis.com/auth/spreadsheets](https://www.googleapis.com/auth/spreadsheets)",
    "[https://www.googleapis.com/auth/drive](https://www.googleapis.com/auth/drive)"
  ]
}

```