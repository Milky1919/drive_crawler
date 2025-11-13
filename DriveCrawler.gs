/**
 * @OnlyCurrentDoc
 *
 * Google Drive クローラー (GAS版)
 *
 * 指定されたGoogle Driveフォルダ配下をクロールし、ファイル情報をスプレッドシートに出力します。
 * GWSの実行時間制限（30分）を考慮し、レジューム（中断・再開）機能を搭載しています。
 *
 * @BasedOn "Google Drive クローラー仕様書 (GAS版)"
 * @Version 3.0 (Changes API アプローチの根本的見直し (v2.6準拠) および 終了処理バグ(L382)の修正)
 */

//--- グローバル設定 ---
const MIME_TYPE_FOLDER = 'application/vnd.google-apps.folder';

// 実行時間制限（仕様書 4.4項: 安全マージンを取り25分）
const MAX_EXECUTION_TIME_MINUTES = 25;
const MAX_EXECUTION_TIME_MS = MAX_EXECUTION_TIME_MINUTES * 60 * 1000;
let SCRIPT_START_TIME = new Date(); // スクリプト開始時刻

// レジューム機能（スクリプトプロパティ）のキー
const RESUME_QUEUE_KEY = 'FOLDER_QUEUE'; // 未処理フォルダキュー
const START_PAGE_TOKEN_KEY = 'START_PAGE_TOKEN'; // Changes API トークン

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する
 * (仕様書 3. インターフェース)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('[Drive クローラー]')
    .addItem('1. フルクロール実行 (続きから)', 'main_FullCrawl')
    .addItem('2. 更新確認実行 (差分)', 'main_CheckUpdates')
    .addSeparator()
    .addItem('3. 実行状態をリセット', 'resetResumeData')
    .addToUi();
}

// ---------------------------------------------------
// メイン機能 (1) フルクロール
// ---------------------------------------------------

/**
 * メイン関数: フルクロールを実行する
 * (仕様書 5.2. フルクロール機能)
 */
function main_FullCrawl() {
  SCRIPT_START_TIME = new Date(); // 実行開始時刻をセット
  let config, filesToWrite = [], foldersToWrite = [];
  let folderQueue = [], processedFilesSet = new Set(), allFolderIdSet = new Set();
  const scriptProperties = PropertiesService.getScriptProperties();

  try {
    // 1. 初期化 (仕様書 5.2.1 1項)
    // getConfiguration() 内で起点フォルダの driveId もチェックする (v2.0 修正)
    config = getConfiguration();
    if (!config) return; // 設定取得失敗
    config.statusCell.setValue('処理開始: フルクロールを開始します...');
    SpreadsheetApp.flush();

    // 1-1. レジュームデータ読み込み
    try {
      const queueData = scriptProperties.getProperty(RESUME_QUEUE_KEY);
      folderQueue = queueData ? JSON.parse(queueData) : [];
    } catch (e) {
      logError(config, 'レジュームデータ読込失敗', null, e);
      folderQueue = []; // キューをリセットして続行
    }

    // 1-2. 処理済みリスト構築 (容量制限対策)
    config.statusCell.setValue('処理中: 既存ファイルリストを構築中...');
    SpreadsheetApp.flush();
    processedFilesSet = loadIdSetFromSheet(config.outputSheet, 0); // A列 (ファイルID)
    allFolderIdSet = loadIdSetFromSheet(config.folderListSheet, 0); // A列 (フォルダID)

    // 2. キューの準備 (仕様書 5.2.1 2項)
    if (folderQueue.length === 0) {
      // (v2.0 修正) config.startFolder は getConfiguration() で取得済み
      const startFolder = config.startFolderInfo;
      
      folderQueue.push({ id: startFolder.id, name: startFolder.name });
      if (!allFolderIdSet.has(startFolder.id)) {
        allFolderIdSet.add(startFolder.id);
        foldersToWrite.push([startFolder.id, startFolder.name]); // ヘッダーに合わせて調整
      }
    } else {
      config.statusCell.setValue(`処理再開: 未処理フォルダ ${folderQueue.length} 件...`);
      SpreadsheetApp.flush();
    }

    // 3. メインループ (仕様書 5.2.1 3項)
    while (folderQueue.length > 0) {
      // 3-1. 時間チェック
      if (checkTimeLimit()) {
        config.statusCell.setValue('中断 (時間制限): 処理を中断します。');
        SpreadsheetApp.flush();
        break; // whileループを抜ける
      }

      // 3-2. キュー取得
      const currentFolder = folderQueue.shift();
      const statusMsg = `処理中: ${currentFolder.name} (残りキュー: ${folderQueue.length})`;
      config.statusCell.setValue(statusMsg);
      Logger.log(statusMsg);
      SpreadsheetApp.flush();

      let pageToken = null;
      try {
        do {
          // 3-3. アイテム一覧取得 (Drive API v3)
          if (checkTimeLimit()) break; // ページネーション途中でも中断

          // (v2.0 修正) Files.list は driveId を必要としない (supportsAllDrives のみ)
          const response = exponentialBackoff(() => {
            return Drive.Files.list({
              q: `'${currentFolder.id}' in parents and trashed=false`,
              fields: 'nextPageToken, files(id, name, mimeType, parents, createdTime, modifiedTime, owners)',
              pageToken: pageToken,
              pageSize: 500, // 高速化のため最大数を要求
              supportsAllDrives: true,
              includeItemsFromAllDrives: true
            });
          });

          if (!response || !response.files) continue;

          // 3-4. アイテム処理
          for (const item of response.files) {
            if (checkTimeLimit()) break; // アイテム処理途中でも中断

            if (item.mimeType === MIME_TYPE_FOLDER) {
              // フォルダの場合
              folderQueue.push({ id: item.id, name: item.name });
              if (!allFolderIdSet.has(item.id)) {
                allFolderIdSet.add(item.id);
                foldersToWrite.push([item.id, item.name]); // [ID, Name]
              }
            } else {
              // ファイルの場合
              if (!processedFilesSet.has(item.id)) {
                const fileInfo = formatFileRow(item, currentFolder.name);
                filesToWrite.push(fileInfo);
                processedFilesSet.add(item.id);
              }
            }
          } // for (item)
          
          if (checkTimeLimit()) break; // 内部ループで中断した場合、pageTokenループも抜ける
          pageToken = response.nextPageToken;

        } while (pageToken);

      } catch (e) {
        // 5.5項: フォルダ取得失敗時の処理
        logError(config, 'フォルダ一覧取得失敗', currentFolder.id, e);
        // このフォルダはスキップして次のキューに進む
      }

      // 3-5. 結果書き込み (フォルダごと)
      flushWriteBuffers(config, filesToWrite, foldersToWrite);
      
      if (checkTimeLimit()) {
         // メインループの時間チェックがtrueになるよう再設定
         config.statusCell.setValue('中断 (時間制限): フォルダ処理後に中断します。');
         SpreadsheetApp.flush();
         break;
      }
    } // while (queue)

    // 4. 終了処理 (仕様書 5.2.1 4項)
    flushWriteBuffers(config, filesToWrite, foldersToWrite); // 残バファを書き込み

    if (folderQueue.length > 0) {
      // 中断時
      scriptProperties.setProperty(RESUME_QUEUE_KEY, JSON.stringify(folderQueue));
      const msg = `処理中断: 続きは次回実行されます。 (残りキュー: ${folderQueue.length})`;
      config.statusCell.setValue(msg);
      Logger.log(msg);
    } else {
      // 完了時
      scriptProperties.deleteProperty(RESUME_QUEUE_KEY);
      const msg = `処理完了: フルクロールが完了しました。(${new Date().toLocaleString('ja-JP')})`;
      config.statusCell.setValue(msg);
      Logger.log(msg);
      
      // 5.3.1項: Changes APIトークンを保存
      // (v2.0 修正) config (driveId) を渡す
      saveStartPageToken(config);
    }

  } catch (e) {
    logError(config, 'フルクロール全体エラー', null, e);
    config.statusCell.setValue(`[致命的エラー] ${e.message}`);
    // 中断時と同様にキューを保存
    if (folderQueue.length > 0) {
      scriptProperties.setProperty(RESUME_QUEUE_KEY, JSON.stringify(folderQueue));
    }
  }
}

/**
 * ファイル情報をシート出力用の配列にフォーマットする
 * (仕様書 5.2.2)
 */
function formatFileRow(item, currentFolderName) {
  const owner = (item.owners && item.owners.length > 0) ? item.owners[0].emailAddress : 'N/A';
  let fileUrl = `https://drive.google.com/file/d/${item.id}/view`;
  // MIME Typeに応じたURL分岐 (例)
  if (item.mimeType.includes('spreadsheet')) {
    fileUrl = `https://docs.google.com/spreadsheets/d/${item.id}/edit`;
  } else if (item.mimeType.includes('document')) {
    fileUrl = `https://docs.google.com/document/d/${item.id}/edit`;
  } else if (item.mimeType.includes('presentation')) {
    fileUrl = `https://docs.google.com/presentation/d/${item.id}/edit`;
  }

  return [
    item.id,                            // A: ファイルID
    item.name,                          // B: ファイル名
    fileUrl,                            // C: ファイルURL
    item.mimeType,                      // D: MIME Type
    item.parents ? item.parents.join(', ') : '', // E: 親フォルダID (全て)
    currentFolderName,                  // F: 発見時の親フォルダ名
    item.createdTime,                   // G: 作成日時
    item.modifiedTime,                  // H: 最終更新日時
    owner,                              // I: オーナー
    new Date()                          // J: 取得日時
  ];
}


// ---------------------------------------------------
// メイン機能 (2) 更新確認
// ---------------------------------------------------

/**
 * メイン関数: 更新確認を実行する
 * (仕様書 5.3. 更新確認機能)
 */
function main_CheckUpdates() {
  SCRIPT_START_TIME = new Date();
  let config;
  const scriptProperties = PropertiesService.getScriptProperties();

  try {
    // 1. 初期化 (仕様書 5.3.2 1項)
    config = getConfiguration();
    if (!config) return;
    config.statusCell.setValue('処理開始: 更新確認を開始します...');
    SpreadsheetApp.flush();

    let startPageToken = scriptProperties.getProperty(START_PAGE_TOKEN_KEY);
    if (!startPageToken) {
      config.statusCell.setValue('[エラー] トークンがありません。先にフルクロールを実行してください。');
      return;
    }

    config.statusCell.setValue('処理中: 既存リストをメモリに構築中...');
    SpreadsheetApp.flush();
    const allFolderIdSet = loadIdSetFromSheet(config.folderListSheet, 0); // A列 (フォルダID)
    const processedFilesSet = loadIdSetFromSheet(config.outputSheet, 0); // A列 (ファイルID)
    
    // ログ書き込み用バッファ
    let logsToWrite = [];
    let newFilesToWrite = [];
    let newFoldersToWrite = [];

    // 2. 変更履歴の取得 (仕様書 5.3.2 2項)
    let pageToken = null; // [v3.1 修正] 初回は null で開始
    let newStartPageToken = null;
    let response = null; 
    let wasInterrupted = false; // [v3.0 修正] 終了処理(L382)のバグ修正用フラグ

    config.statusCell.setValue('処理中: Drive APIから変更履歴を取得中...');
    SpreadsheetApp.flush();
    
    // (v2.0 修正) 共有ドライブの場合 driveId が必須
    const changesApiParams = {
      startPageToken: startPageToken,
      fields: 'nextPageToken, newStartPageToken, changes(fileId, time, removed, file(id, name, mimeType, parents, createdTime, modifiedTime, owners))',
      pageSize: 500,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    };
    if (config.targetDriveId) {
      changesApiParams.driveId = config.targetDriveId;
    }

    // 3. ループ処理 (仕様書 5.3.2 3項)
    do {
      if (checkTimeLimit()) {
        config.statusCell.setValue('中断 (時間制限): 処理を中断します。');
        SpreadsheetApp.flush();
        wasInterrupted = true; // [v3.0 修正]
        break; // whileループを抜ける
      }
      
      // pageTokenを更新
      changesApiParams.pageToken = pageToken;

      try {
        response = exponentialBackoff(() => {
          // (v2.0 修正) 共有ドライブ対応済みのパラメータを使用
          return Drive.Changes.list(changesApiParams);
        });
      } catch (e) {
        const errorMsg = e.message.toLowerCase();
        if (errorMsg.includes('invalid value')) {
          // [v3.1] Invalid Value 自動回復処理
          const recoveryResult = handleInvalidTokenError(config, startPageToken, e);
          if (recoveryResult.recovered) {
            startPageToken = recoveryResult.newToken; // トークンを更新
            changesApiParams.startPageToken = startPageToken; // APIパラメータも更新
            pageToken = null; // ループを最初からやり直す
            response = null; // responseをリセット
            continue; // 次のループへ
          } else {
            wasInterrupted = true; // 回復失敗
            break;
          }
        } else {
          // その他のAPIエラー
          logError(config, 'Changes.list API失敗', `Token: ${pageToken}`, e);
          config.statusCell.setValue(`[エラー] Changes.list API失敗。詳細はログ(B6)確認。`);
          wasInterrupted = true;
          break;
        }
      }

      if (!response) {
        // APIエラー（バックオフ失敗）
        const errorMsg = `Changes.list API失敗 (pageToken: ${pageToken}, DriveId: ${config.targetDriveId})`;
        logError(config, 'Changes.list API失敗', `Token: ${pageToken}`, new Error(errorMsg));
        config.statusCell.setValue(`[エラー] Changes.list API失敗。詳細はログ(B6)確認。`);
        wasInterrupted = true; // [v3.0 修正]
        break;
      }
      
      newStartPageToken = response.newStartPageToken; // 常に最新のトークンを保持

      if (response.changes && response.changes.length > 0) {
        // 変更アイテム処理
        for (const change of response.changes) {
          if (checkTimeLimit()) {
            wasInterrupted = true; // [v3.0 修正]
            break;
          }
          
          if (change.removed === true) {
            // 削除処理
            if (processedFilesSet.has(change.fileId)) {
              // ファイル削除
              logsToWrite.push([new Date(change.time), change.fileId, '削除', 'ファイル']);
              processedFilesSet.delete(change.fileId);
            } else if (allFolderIdSet.has(change.fileId)) {
              // フォルダ削除
              logsToWrite.push([new Date(change.time), change.fileId, '削除 [フォルダ]', 'フォルダ']);
              allFolderIdSet.delete(change.fileId);
              // シートからのリアルタイム行削除は行わない (5.3.2 3項)
            }
          } else if (change.file && change.file.parents) {
            // 追加/更新処理
            // 管理対象かチェック (親が管理対象Setに含まれるか)
            const isTargetParent = change.file.parents.some(parentId => allFolderIdSet.has(parentId));
            
            if (isTargetParent) {
              if (change.file.mimeType === MIME_TYPE_FOLDER) {
                // フォルダの場合
                if (allFolderIdSet.has(change.file.id)) {
                  // 更新
                  logsToWrite.push([new Date(change.time), change.file.id, '更新', 'フォルダ', change.file.name]);
                } else {
                  // 新規追加
                  logsToWrite.push([new Date(change.time), change.file.id, '追加', 'フォルダ', change.file.name]);
                  allFolderIdSet.add(change.file.id);
                  newFoldersToWrite.push([change.file.id, change.file.name]);
                }
              } else {
                // ファイルの場合
                if (processedFilesSet.has(change.file.id)) {
                  // 更新
                  logsToWrite.push([new Date(change.time), change.file.id, '更新', 'ファイル', change.file.name]);
                  // FullCrawl_List の更新は行わない (5.3.2 3項)
                } else {
                  // 新規追加
                  logsToWrite.push([new Date(change.time), change.file.id, '追加', 'ファイル', change.file.name]);
                  processedFilesSet.add(change.file.id);
                  // FullCrawl_List (B2) にも追記
                  const fileInfo = formatFileRow(change.file, '(Changes API)'); // 親名は不明
                  newFilesToWrite.push(fileInfo);
                }
              }
            }
          }
        } // for (change)
        
        // バッファ書き込み
        flushUpdateBuffers(config, logsToWrite, newFilesToWrite, newFoldersToWrite);

      } // if (response.changes)

      if (checkTimeLimit()) {
        wasInterrupted = true; // [v3.0 修正]
        break; // 内側で時間切れ
      }
      pageToken = response.nextPageToken;

    } while (pageToken);
    
    // 4. 終了処理 (仕様書 5.3.2 4項)
    flushUpdateBuffers(config, logsToWrite, newFilesToWrite, newFoldersToWrite);

    // [v3.0 修正] 終了判定ロジック (B4セル上書きバグ修正)
    // pageToken が null (while が完了) -> 正常完了
    // wasInterrupted が true (break で中断) -> 中断
    if (pageToken || wasInterrupted) { 
      // [v3.0 修正] 中断 (時間切れ or APIキャッシュ or API失敗)
      const msg = `処理中断: 更新確認を中断しました。次回同じトークンから再開します。`;
      Logger.log(msg);

      if (checkTimeLimit()) {
        // [v2.5 修正] L252 (時間切れ) で中断した場合
        logError(config, '中断（時間制限）', `startToken: ${startPageToken}`, new Error(msg));
      } else {
        // [v3.0 修正] L290 (APIキャッシュ問題) または L310 (API失敗) で中断した場合
        // B4セルとエラーログには既に原因が記載されている
      }

    } else {
      // 正常完了時 (pageTokenがnullになった)
      if (newStartPageToken) {
        scriptProperties.setProperty(START_PAGE_TOKEN_KEY, newStartPageToken);
        const msg = `処理完了: 更新確認が完了しました。(${new Date().toLocaleString('ja-JP')})`;
        config.statusCell.setValue(msg);
        Logger.log(msg);
      } else {
        // 変更が全くなかった場合、newStartPageTokenがnullのことがあるが、
        // startPageToken (ループ開始時のトークン) は有効なままなので、それを再利用する
        if (startPageToken) {
          scriptProperties.setProperty(START_PAGE_TOKEN_KEY, startPageToken);
           const msg = `処理完了: 変更はありませんでした。(${new Date().toLocaleString('ja-JP')})`;
           config.statusCell.setValue(msg);
           Logger.log(msg);
        } else {
          logError(config, 'トークン更新失敗', null, new Error('newStartPageTokenが取得できませんでした。'));
          config.statusCell.setValue('[エラー] 新しいトークンの取得に失敗しました。');
        }
      }
    }

  } catch (e) {
    logError(config, '更新確認全体エラー', null, e);
    config.statusCell.setValue(`[致命的エラー] ${e.message}`);
  }
}

/**
 * 更新確認の書き込みバッファをフラッシュする
 */
function flushUpdateBuffers(config, logsToWrite, newFilesToWrite, newFoldersToWrite) {
  try {
    if (logsToWrite.length > 0) {
      const sheet = config.updateLogSheet;
      sheet.getRange(sheet.getLastRow() + 1, 1, logsToWrite.length, logsToWrite[0].length).setValues(logsToWrite);
      logsToWrite.length = 0; // バッファクリア
    }
    if (newFilesToWrite.length > 0) {
      const sheet = config.outputSheet;
      sheet.getRange(sheet.getLastRow() + 1, 1, newFilesToWrite.length, newFilesToWrite[0].length).setValues(newFilesToWrite);
      newFilesToWrite.length = 0;
    }
    if (newFoldersToWrite.length > 0) {
      const sheet = config.folderListSheet;
      sheet.getRange(sheet.getLastRow() + 1, 1, newFoldersToWrite.length, newFoldersToWrite[0].length).setValues(newFoldersToWrite);
      newFoldersToWrite.length = 0;
    }
  } catch (e) {
    logError(config, '更新確認バッファ書込エラー', null, e);
  }
}


// ---------------------------------------------------
// メイン機能 (3) リセット
// ---------------------------------------------------

/**
 * レジュームデータ（キュー）をリセットし、Changes APIトークンを再取得する
 * (仕様書 4.5, 5.3.1)
 */
function resetResumeData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    '本当に実行状態（未処理キューとAPIトークン）をリセットしますか？\n' +
    '※シート(B2, B3, B5, B6)のデータは手動でクリアしてください。',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    let config;
    try {
      config = getConfiguration(); // B4セル等の取得のため
      if (!config) return; // 設定読み込み失敗
      
      // 1. キューを削除
      PropertiesService.getScriptProperties().deleteProperty(RESUME_QUEUE_KEY);
      Logger.log('未処理キュー (FOLDER_QUEUE) を削除しました。');

      // 2. 新しいトークンを取得・保存 (v2.0 修正)
      saveStartPageToken(config);
      
      const msg = 'リセット完了: 未処理キューを削除し、新しいAPIトークンを取得しました。';
      if (config && config.statusCell) {
        config.statusCell.setValue(msg);
      }
      ui.alert(msg);

    } catch (e) {
      // [v2.5 修正] saveStartPageToken がスローしたエラーは既にログ済み
      if (config && config.statusCell && !config.statusCell.getValue().includes('[エラー]')) {
        logError(config, 'リセット処理エラー', null, e);
        config.statusCell.setValue(`[エラー] リセットに失敗: ${e.message}`);
      }
      ui.alert(`[エラー] リセットに失敗しました: ${e.message}`);
    }
  }
}

/**
 * Changes API の開始トークンを取得し、PropertiesService に保存する
 * (仕様書 5.3.1)
 * (v2.3 修正) includeItemsFromAllDrives: true を追加
 * (v3.1 修正) 自動回復モード (isRecovery) を追加
 * @param {object} config 設定オブジェクト
 * @param {boolean} [isRecovery=false] 自動回復モードか
 * @returns {string|null} 成功時は新しいトークン、失敗時はnull
 */
function saveStartPageToken(config, isRecovery = false) {
  try {
    if (!config) {
       config = getConfiguration(); // configが渡されなかった場合 (リセット時など)
       if (!config) return null;
    }
    
    if (!isRecovery) {
      config.statusCell.setValue('処理中: 新しいAPIトークンを取得中...');
      SpreadsheetApp.flush();
    }
    
    // (v2.3 修正) 共有ドライブの場合 driveId と includeItemsFromAllDrives が必須
    const apiParams = {
      supportsAllDrives: true,
      includeItemsFromAllDrives: true // [v2.3 修正] 不足していたパラメータを追加
    };
    if (config.targetDriveId) {
      apiParams.driveId = config.targetDriveId;
    }
    
    const response = exponentialBackoff(() => {
      return Drive.Changes.getStartPageToken(apiParams);
    });
    
    if (response && response.startPageToken) {
      PropertiesService.getScriptProperties().setProperty(START_PAGE_TOKEN_KEY, response.startPageToken);
      Logger.log(`新しいAPIトークンを保存しました。 Token: ${response.startPageToken} (DriveId: ${config.targetDriveId || 'N/A'})`);
      if (!isRecovery) {
        config.statusCell.setValue('APIトークンを更新しました。');
        SpreadsheetApp.flush();
      }
      return response.startPageToken; // [v3.1] 成功時はトークンを返す
    } else {
      throw new Error('APIからトークンが返されませんでした。');
    }
  } catch (e) {
    // [v2.2 修正] エラーメッセージを具体的にする
    let errorMsg = e.message;
    const lowerError = e.message.toLowerCase();

    // 権限不足 (Forbidden) または Invalid Value の場合
    if (lowerError.includes('invalid value') || lowerError.includes('forbidden') || lowerError.includes('not found') || lowerError.includes('file not found')) {
        errorMsg = `APIトークンの取得に失敗。実行者の権限が対象ドライブの「コンテンツ管理者」以上であることを確認してください。(詳細: ${e.message})`;
    } else {
        errorMsg = `APIトークンの取得に失敗: ${e.message}`;
    }
    
    // [v2.5 修正] config が null でも Logger.log は実行する
    logError(config, 'APIトークン取得失敗', config ? config.targetDriveId : 'N/A', new Error(errorMsg));
    
    if (config && config.statusCell && !isRecovery) {
      config.statusCell.setValue(`[エラー] ${errorMsg}`);
      SpreadsheetApp.flush();
    }

    if (isRecovery) {
      // [v3.1] 自動回復モードではエラーをスローせず null を返す
      return null;
    } else {
      // 通常モードではエラーをスローして停止させる
      throw new Error(errorMsg);
    }
  }
}


/**
 * [v3.1 追加] Invalid Value エラーの自動回復処理
 * @param {object} config 設定オブジェクト
 * @param {string} oldToken エラーが発生した古いトークン
 * @param {Error} error 発生したエラーオブジェクト
 * @returns {{recovered: boolean, newToken: string|null}} 回復結果
 */
function handleInvalidTokenError(config, oldToken, error) {
  const errorMessage = (error && error.message) ? error.message : String(error);
  logError(config, 'Invalid Value (pageToken Mismatch)', `Token: ${oldToken} / DriveId: ${config.targetDriveId || 'N/A'}`, new Error(`自動回復処理を開始します。${errorMessage}`));

  try {
    // 新しいトークンを再取得
    const newToken = saveStartPageToken(config, true); // isRecovery=true を渡す

    if (newToken && newToken !== oldToken) {
      Logger.log(`Invalid Valueから回復成功。新しいトークン: ${newToken}`);
      config.statusCell.setValue('処理中: APIトークンを自動更新し、処理を継続します...');
      SpreadsheetApp.flush();
      return { recovered: true, newToken: newToken };
    } else {
      // APIは同じトークンを返した (キャッシュ問題)
      // [修正] 中断せず、API側のレプリケーション遅延を考慮して待機する
      const waitTimeMs = 5000; // 待機時間 (例: 5秒)
      const msg = `APIキャッシュ問題（待機）: Token: ${oldToken}。 APIキャッシュの同期を ${waitTimeMs / 1000}秒 待機して同じトークンでリトライします。`;

      logError(config, 'APIキャッシュ問題（待機）', `Token: ${oldToken}`, new Error(msg)); // ログレベルを変更
      config.statusCell.setValue(`処理中: APIキャッシュの同期を ${waitTimeMs / 1000}秒 待機中...`); // ステータスを変更
      SpreadsheetApp.flush();

      Utilities.sleep(waitTimeMs); // ★待機する

      // [修正] 回復成功として、同じトークン (newToken または oldToken) を返す
      return { recovered: true, newToken: newToken };
    }
  } catch (e) {
    // saveStartPageToken 内でエラーが発生した場合
    logError(config, '自動回復処理失敗', `Token: ${oldToken}`, e);
    config.statusCell.setValue(`[エラー] 自動回復処理に失敗しました: ${e.message}`);
    SpreadsheetApp.flush();
    return { recovered: false, newToken: null };
  }
}


/**
 * [v3.1 追加] Drive API の基本的な権限 (drive.readonly) が付与されているか事前チェックする
 * @returns {boolean} 権限があれば true, なければ false
 */
function checkApiPermissions() {
  try {
    // 最も軽量な readonly API をテスト呼び出し
    Drive.About.get({ fields: 'kind' });
    return true;
  } catch (e) {
    const errorMsg = e.message.toLowerCase();
    // 一般的な権限エラーのキーワード
    if (errorMsg.includes('authorization') || errorMsg.includes('forbidden') || errorMsg.includes('access denied')) {
       const msg = '[権限エラー] Drive APIへのアクセスが許可されていません。\n\n' +
        '【確認してください】\n' +
        '1. スクリプトが要求するGoogle Driveの権限を許可しましたか？\n' +
        '2. (Workspace管理者様) ドメイン全体のポリシーでAPIアクセスが制限されていませんか？';
      SpreadsheetApp.getUi().alert(msg);
      // getConfiguration内のログで詳細を記録するため、ここではUI表示のみ
    }
    return false;
  }
}


// ---------------------------------------------------
// ユーティリティ関数
// ---------------------------------------------------

/**
 * '設定' シートから実行時設定を取得する
 * (仕様書 5.1)
 * (v2.0 修正) 起点フォルダの driveId を取得する
 */
function getConfiguration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('設定');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('エラー: "設定" シートが見つかりません。');
    return null;
  }

  let config = {};
  try {
    // [v3.1 追加] API権限の事前チェック
    if (!checkApiPermissions()) {
      const error = new Error('Drive APIの権限がありません。スクリプトの実行を中止します。');
      try {
        const errorLogSheetName = sheet.getRange('B6').getValue().toString().trim();
        if (errorLogSheetName) {
            const errorLogSheet = getOrCreateSheet(ss, errorLogSheetName, ['日時', '発生箇所', '対象ID', 'エラーメッセージ']);
            errorLogSheet.appendRow([new Date(), 'API権限チェック (getConfiguration)', 'N/A', error.message]);
        }
      } catch (logError) {
        Logger.log(`[致命的エラー] API権限エラーのログ記録に失敗: ${logError.message}`);
      }
      return null;
    }

    config = {
      ss: ss,
      startFolderId: extractIdFromUrl(sheet.getRange('B1').getValue().toString().trim()),
      outputSheetName: sheet.getRange('B2').getValue().toString().trim(),
      updateLogSheetName: sheet.getRange('B3').getValue().toString().trim(),
      statusCell: sheet.getRange('B4'),
      folderListSheetName: sheet.getRange('B5').getValue().toString().trim(),
      errorLogSheetName: sheet.getRange('B6').getValue().toString().trim(),
      targetDriveId: null, // (v2.0 追加)
      startFolderInfo: null // (v2.0 追加)
    };

    if (!config.startFolderId || !config.outputSheetName || !config.updateLogSheetName || !config.folderListSheetName || !config.errorLogSheetName) {
      throw new Error('B1～B6のすべてのセルに入力が必要です。');
    }

    // [v2.5 修正] エラーログシートを早期に取得 (起点フォルダ取得失敗時もログするため)
    config.errorLogSheet = getOrCreateSheet(ss, config.errorLogSheetName, [
      '日時', '発生箇所', '対象ID', 'エラーメッセージ'
    ]);

    // (v2.0 追加) 起点フォルダの情報を取得し、driveId (共有ドライブか) を確認
    try {
      config.statusCell.setValue('処理中: 起点フォルダの情報を取得中...');
      SpreadsheetApp.flush();
      const startFolder = exponentialBackoff(() => {
        return Drive.Files.get(config.startFolderId, {
          fields: 'id, name, driveId', // driveId をリクエスト
          supportsAllDrives: true
        });
      });
      if (!startFolder) throw new Error('起点フォルダが見つかりません。');
      
      config.startFolderInfo = startFolder;
      if (startFolder.driveId) {
        config.targetDriveId = startFolder.driveId;
        Logger.log(`起点フォルダは共有ドライブです。DriveId: ${config.targetDriveId}`);
      } else {
        Logger.log('起点フォルダはマイドライブです。');
      }

    } catch (e) {
      logError(config, '起点フォルダ取得失敗', config.startFolderId, e);
      config.statusCell.setValue(`[致命的エラー] 起点フォルダ(B1)の取得に失敗: ${e.message}`);
      return null; // 5.5項: 続行不可能
    }


    // 各シートを取得または作成
    config.outputSheet = getOrCreateSheet(ss, config.outputSheetName, [
      'ファイルID', 'ファイル名', 'ファイルURL', 'MIME Type', '親フォルダID (全て)', '発見時の親フォルダ名', '作成日時', '最終更新日時', 'オーナー', '取得日時'
    ]);
    config.updateLogSheet = getOrCreateSheet(ss, config.updateLogSheetName, [
      '日時', 'ID', '変更タイプ', 'アイテム種別', '名前'
    ]);
    config.folderListSheet = getOrCreateSheet(ss, config.folderListSheetName, [
      'フォルダID', 'フォルダ名'
    ]);
    
    config.statusCell.clearContent();
    return config;

  } catch (e) {
    const errorMsg = (e.message || '不明なエラー');
    Logger.log(`設定エラー: ${errorMsg}`);
    // [v2.5 修正] config.errorLogSheet があればログを試みる
    logError(config, '設定エラー (getConfiguration)', null, e);
    
    if (config.statusCell) config.statusCell.setValue(`[設定エラー] ${errorMsg}`);
    SpreadsheetApp.getUi().alert(`設定エラー: ${errorMsg}`);
    return null;
  }
}

/**
 * DriveのURLからIDを抽出する
 * (v2.1 修正) 共有ドライブのルートURL (drive-list) に対応
 */
function extractIdFromUrl(input) {
  if (!input) return null;
  // 既にID形式 (英数字と-_のみ)
  if (/^[\w-]{15,}$/.test(input)) {
    return input;
  }
  // 共有ドライブのルート (drive/drive-list/...)
  const matchDriveList = input.match(/drive-list\/([-\w]{15,})/);
  if (matchDriveList) {
    return matchDriveList[1];
  }
  // フォルダ (folders/...) または (id=...)
  const matchFolder = input.match(/folders\/([-\w]{25,})|id=([-\w]{25,})/);
  if (matchFolder) {
    return matchFolder[1] || matchFolder[2];
  }
  // ファイル (d/...)
  const matchFile = input.match(/d\/([-\w]{25,})/);
  return matchFile ? matchFile[1] : null;
}

/**
 * ヘッダー付きでシートを取得または作成する
 */
function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * シートの指定列からIDを読み込みSetを構築する
 * @param {Sheet} sheet 対象シート
 * @param {number} colIndex 読み込む列 (0-indexed)
 */
function loadIdSetFromSheet(sheet, colIndex) {
  const set = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) { // ヘッダー行を除く
    const range = sheet.getRange(2, colIndex + 1, lastRow - 1, 1);
    const values = range.getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0]) {
        set.add(values[i][0].toString());
      }
    }
  }
  Logger.log(`シート "${sheet.getName()}" から ${set.size} 件のIDを読み込みました。`);
  return set;
}

/**
 * フルクロールの書き込みバッファをシートに書き込む
 */
function flushWriteBuffers(config, filesToWrite, foldersToWrite) {
  try {
    if (filesToWrite.length > 0) {
      const sheet = config.outputSheet;
      sheet.getRange(sheet.getLastRow() + 1, 1, filesToWrite.length, filesToWrite[0].length).setValues(filesToWrite);
      filesToWrite.length = 0; // バッファクリア
    }
    if (foldersToWrite.length > 0) {
      const sheet = config.folderListSheet;
      sheet.getRange(sheet.getLastRow() + 1, 1, foldersToWrite.length, foldersToWrite[0].length).setValues(foldersToWrite);
foldersToWrite.length = 0; // バッファクリア
    }
  } catch (e) {
    logError(config, 'バッファ書込エラー', null, e);
  }
}

/**
 * 実行時間が上限（25分）に近いかチェックする
 * (仕様書 4.4)
 */
function checkTimeLimit() {
  const elapsedTime = new Date().getTime() - SCRIPT_START_TIME.getTime();
  return elapsedTime >= MAX_EXECUTION_TIME_MS;
}

/**
 * エラーログシートにエラーを記録する
 * (仕様書 5.5)
 * (v2.5 修正) config が null でも Logger.log は実行する
 */
function logError(config, location, targetId, error) {
  const msg = (error && error.message) ? error.message : String(error);
  Logger.log(`[エラー発生] 箇所: ${location}, 対象: ${targetId}, MSG: ${msg}`);
  
  // [v2.5 修正] config と errorLogSheet が利用可能な場合のみシートに書き込む
  if (config && config.errorLogSheet) {
    try {
      config.errorLogSheet.appendRow([
        new Date(),
        location,
        targetId || 'N/A',
        msg
      ]);
      SpreadsheetApp.flush(); // エラーは即時書き込み
    } catch (e) {
      Logger.log(`[致命的エラー] エラーログの書き込みに失敗: ${e.message}`);
    }
  }
}

/**
 * API呼び出しを指数バックオフでラップする
 * (仕様書 5.4.1)
 * @param {function} apiCallFunction 実行するAPI呼び出し (例: () => Drive.Files.list({...}))
 */
function exponentialBackoff(apiCallFunction, maxRetries = 5) {
  let attempts = 0;
  while (attempts < maxRetries) {
    try {
      return apiCallFunction(); // API呼び出しを実行
    } catch (e) {
      const errorMsg = e.message.toLowerCase();
      // レートリミット (403) または サーバーエラー (5xx) か判定
      // (v2.0 修正) "invalid value" はリトライ対象外とする (400 Bad Request)
      if ((errorMsg.includes('ratelimitexceeded') || 
          errorMsg.includes('user rate limit exceeded') ||
          errorMsg.includes('backend error') ||
          errorMsg.includes('service invoked too many times') ||
          e.message.startsWith('Exception: Service error') || // 5xx系
          errorMsg.includes('internal error')) &&
          !errorMsg.includes('invalid value') // 400系はリトライしない
      ) {
        attempts++;
        if (attempts >= maxRetries) {
          Logger.log(`APIエラー: 最大リトライ回数 (${maxRetries}) に達しました。エラーをスローします。 MSG: ${e.message}`);
          throw e; // 最大リトライ回数を超えたらエラーを投げる
        }
        const waitTime = Math.pow(2, attempts) * 1000 + Math.random() * 1000; // 2^n秒 + ランダムなミリ秒
        Logger.log(`APIエラー (試行 ${attempts} / ${maxRetries}): ${waitTime}ms 待機後にリトライします。 MSG: ${e.message}`);
        Utilities.sleep(waitTime);
      } else {
        // レートリミト以外のエラー (例: 404 Not Found, 401 Unauthorized, 400 Invalid Value) は即時スロー
        Logger.log(`APIエラー (リトライ対象外): ${e.message}`);
        throw e;
      }
    }
  }
}