/**
 * SlideCreator.gs
 * 論文解析データからGoogleスライドを生成するGASスクリプト
 * 
 * 機能：
 * 1. テンプレートスライドの複製
 * 2. プレースホルダーテキストの置換
 * 3. フォルダへの保存と管理
 */

/**
 * 解析結果からスライドを生成
 * @param {Object} paperInfo 論文解析情報
 * @returns {string} 生成したスライドのURL
 */
function createSlide(paperInfo) {
  try {
    console.log('スライド生成を開始します...');
    
    // スライドテンプレートIDをスクリプトプロパティから取得
    const TEMPLATE_SLIDE_ID = PropertiesService.getScriptProperties().getProperty('PDF_SLIDE_TEMPLATE_ID');
    if (!TEMPLATE_SLIDE_ID) {
      throw new Error('PDF_SLIDE_TEMPLATE_IDがスクリプトプロパティに設定されていません。');
    }
    
    // スライド出力フォルダIDを取得
    const SLIDE_OUTPUT_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('SLIDE_OUTPUT_FOLDER_ID');
    
    // 新しいファイル名を設定（日時_タイトル形式）
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    
    // paperInfo.Titleの安全なアクセス
    const title = (paperInfo && paperInfo.Title) ? paperInfo.Title.substring(0, 50) : 'Untitled';
    const fileName = `${timestamp}_${title}`;
    
    console.log(`スライドファイル名: ${fileName}`);
    
    // テンプレートをコピー
    const templateFile = DriveApp.getFileById(TEMPLATE_SLIDE_ID);
    const targetFolder = SLIDE_OUTPUT_FOLDER_ID ? 
                        DriveApp.getFolderById(SLIDE_OUTPUT_FOLDER_ID) : 
                        DriveApp.getRootFolder();
    
    const newFile = templateFile.makeCopy(fileName, targetFolder);
    
    console.log(`スライドテンプレートをコピーしました: ${newFile.getName()}`);
    
    // スライドを開く
    const presentation = SlidesApp.openById(newFile.getId());
    
    // プレースホルダー置換（安全なアクセス）
    const replacements = {
      '{Title}': (paperInfo && paperInfo.Title) ? paperInfo.Title : 'タイトル情報なし',
      '{Japanese_Title}': (paperInfo && paperInfo.Japanese_Title) ? paperInfo.Japanese_Title : 'タイトル情報なし',
      '{Journal}': (paperInfo && paperInfo.Journal) ? paperInfo.Journal : '雑誌情報なし',
      '{Citation}': (paperInfo && paperInfo.Citation) ? paperInfo.Citation : '引用情報なし',
      '{Limitation}': (paperInfo && paperInfo.Limitation) ? paperInfo.Limitation : '記載なし',
      '{ClinicalImplication}': (paperInfo && paperInfo.ClinicalImplication) ? paperInfo.ClinicalImplication : '記載なし',
      '{ResearchImplication}': (paperInfo && paperInfo.ResearchImplication) ? paperInfo.ResearchImplication : '記載なし',
      '{PolicyImplication}': (paperInfo && paperInfo.PolicyImplication) ? paperInfo.PolicyImplication : '記載なし',
      '{Summary}': (paperInfo && paperInfo.Summary) ? paperInfo.Summary : '記載なし'
    };
    
    console.log('置換対象値:', replacements);
    
    console.log('プレースホルダー置換処理を開始します');
    
    // スライド内の全テキストを置換
    const slides = presentation.getSlides();
    for (const slide of slides) {
      for (const placeholder in replacements) {
        try {
          slide.replaceAllText(placeholder, replacements[placeholder]);
        } catch (replaceError) {
          console.warn(`プレースホルダー ${placeholder} の置換中にエラーが発生しました: ${replaceError}`);
          // 置換エラーは全体の処理を停止しない
        }
      }
    }
    
    // 変更を保存
    presentation.saveAndClose();
    
    console.log(`スライド生成が完了しました: ${newFile.getUrl()}`);
    return newFile.getUrl();
  } catch (error) {
    console.error(`スライド生成中にエラーが発生しました: ${error}`);
    throw error;
  }
}

/**
 * 特定のフォルダ内の全PDFファイルを処理する
 * 特定のフォルダからPDFを取得し、未処理のファイルをClaudeで解析してスライド作成
 */
function processAllPDFsInFolder() {
  try {
    console.log('フォルダ内のPDFファイル処理を開始します');
    
    // フォルダIDとスプレッドシートIDを取得
    const folderId = PropertiesService.getScriptProperties().getProperty('PDF_FOLDER_ID');
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('TRACKING_SPREADSHEET_ID');
    
    if (!folderId) {
      throw new Error('PDF_FOLDER_IDが設定されていません');
    }
    
    if (!spreadsheetId) {
      throw new Error('TRACKING_SPREADSHEET_IDが設定されていません');
    }
    
    const folder = DriveApp.getFolderById(folderId);
    
    // PDFファイルを取得
    const pdfFiles = folder.getFilesByType(MimeType.PDF);
    
    // 処理済みファイルIDを取得
    const processedIds = getProcessedFileIds();
    let processedCount = 0;
    let errorCount = 0;
    
    // PDFファイルを処理
    while (pdfFiles.hasNext()) {
      const file = pdfFiles.next();
      const fileId = file.getId();
      
      // 処理済みならスキップ
      if (processedIds.has(fileId)) {
        console.log(`ファイルはすでに処理済み: ${file.getName()}`);
        continue;
      }
      
      try {
        console.log(`処理開始: ${file.getName()}`);
        
        // 処理状態の更新
        updateFileStatus(fileId, file.getName(), file.getUrl(), '処理中');
        
        // PDFをClaudeで解析
        const paperInfo = processPDFFromDrive(fileId);
        
        // paperInfoの存在確認
        if (!paperInfo) {
          throw new Error('PDF解析結果が空です');
        }
        
        // デバッグログ
        console.log('解析結果:', paperInfo);
        
        // JSON情報を文字列化
        let jsonInfoStr = '';
        try {
          jsonInfoStr = JSON.stringify(paperInfo);
        } catch (jsonError) {
          console.warn(`JSON文字列化エラー: ${jsonError}`);
          jsonInfoStr = `{"error": "JSON変換エラー: ${jsonError.message}"}`;
        }
        
        // スライド生成
        const slideUrl = createSlide(paperInfo);
        
        // 処理状態の更新（JSON情報、完了フラグ、スライドURLを含む）
        updateFileStatus(
          fileId, 
          file.getName(), 
          file.getUrl(), 
          '処理済', 
          jsonInfoStr,  // JSON情報（E列）
          '',          // エラー情報（F列）
          slideUrl,    // スライドURL（I列）
          '完了'       // 完了フラグ（H列）
        );
        
        console.log(`処理完了: ${file.getName()}`);
        processedCount++;
        
        // 1回の実行で処理するファイル数を制限（GASの実行時間制限対策）
        if (processedCount >= 3) {
          console.log('処理ファイル数上限に達しました。次回の実行で残りを処理します。');
          break;
        }
      } catch (error) {
        console.error(`ファイル処理エラー: ${error}`);
        updateFileStatus(
          fileId, 
          file.getName(), 
          file.getUrl(), 
          'エラー', 
          '',          // JSON情報（E列）
          error.toString(), // エラー情報（F列）
          '',          // スライドURL（I列）
          'エラー'     // 完了フラグ（H列）
        );
        errorCount++;
      }
    }
    
    console.log(`処理完了: ${processedCount}ファイル処理、${errorCount}件のエラー`);
  } catch (error) {
    console.error(`全体処理エラー: ${error}`);
  }
}

/**
 * 処理済みファイルIDのセットを取得
 * @returns {Set} 処理済みファイルIDのセット
 */
function getProcessedFileIds() {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('TRACKING_SPREADSHEET_ID');
    if (!spreadsheetId) {
      console.warn('TRACKING_SPREADSHEET_IDが設定されていません');
      return new Set();
    }
    
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('PDF処理状態');
    
    if (!sheet) {
      return new Set();
    }
    
    const data = sheet.getDataRange().getValues();
    const processedIds = new Set();
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const status = data[i][3]; // 処理状態列
      if (status === '処理済') {
        processedIds.add(data[i][0]); // PDF ID列
      }
    }
    
    console.log(`${processedIds.size}件の処理済みファイルIDを取得しました`);
    return processedIds;
  } catch (error) {
    console.error(`処理済みファイル取得エラー: ${error}`);
    return new Set();
  }
}

/**
 * ファイルの処理状態を更新
 * @param {string} fileId ファイルID
 * @param {string} fileName ファイル名
 * @param {string} fileUrl ファイルURL
 * @param {string} status 処理状態
 * @param {string} jsonInfo JSON情報（オプション）
 * @param {string} errorInfo エラー情報（オプション）
 * @param {string} slideUrl スライドURL（オプション）
 * @param {string} completionFlag 完了フラグ（オプション）
 */
function updateFileStatus(fileId, fileName, fileUrl, status, jsonInfo = '', errorInfo = '', slideUrl = '', completionFlag = '') {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('TRACKING_SPREADSHEET_ID');
    if (!spreadsheetId) {
      console.warn('TRACKING_SPREADSHEET_IDが設定されていません');
      return;
    }
    
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = ss.getSheetByName('PDF処理状態');
    
    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet('PDF処理状態');
      sheet.appendRow(['PDF ID', 'ファイル名', 'ファイルURL', '処理状態', 'JSON情報', 'エラー情報', '処理日時', '完了フラグ', 'スライドURL']);
    }
    
    // ファイルIDで既存の行を検索
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === fileId) {
        rowIndex = i + 1; // 1-indexed
        break;
      }
    }
    
    const timestamp = new Date().toISOString();
    
    if (rowIndex > 0) {
      // 既存行の更新
      sheet.getRange(rowIndex, 4).setValue(status); // 処理状態
      if (jsonInfo) sheet.getRange(rowIndex, 5).setValue(jsonInfo); // JSON情報
      if (errorInfo) sheet.getRange(rowIndex, 6).setValue(errorInfo); // エラー情報
      sheet.getRange(rowIndex, 7).setValue(timestamp); // 処理日時
      if (completionFlag) sheet.getRange(rowIndex, 8).setValue(completionFlag); // 完了フラグ
      if (slideUrl) sheet.getRange(rowIndex, 9).setValue(slideUrl); // スライドURL
    } else {
      // 新規行の追加
      sheet.appendRow([fileId, fileName, fileUrl, status, jsonInfo, errorInfo, timestamp, completionFlag, slideUrl]);
    }
    
    console.log(`ファイル ${fileName} の状態を「${status}」に更新しました`);
  } catch (error) {
    console.error(`状態更新エラー: ${error}`);
  }
}

/**
 * 定期実行のためのトリガーをセットアップ
 */
function setupTrigger() {
  try {
    // 既存のトリガーをクリア
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'processAllPDFsInFolder') {
        ScriptApp.deleteTrigger(trigger);
      }
    }
    
    // 新しいトリガーを作成（毎日午前9時に実行）
    ScriptApp.newTrigger('processAllPDFsInFolder')
      .timeBased()
      .atHour(9)
      .everyDays(1)
      .create();
    
    console.log('毎日午前9時に実行するトリガーを設定しました');
    return '毎日午前9時に実行するトリガーを設定しました';
  } catch (error) {
    console.error(`トリガー設定エラー: ${error}`);
    return `トリガー設定エラー: ${error}`;
  }
}

/**
 * スプレッドシートに保存されているJSON情報からスライドを生成する
 * @param {string} rowId スプレッドシートの行ID（通常はPDF ID）
 * @returns {string} 生成したスライドのURL
 */
function createSlideFromSpreadsheetJSON(rowId) {
  try {
    // IDをチェック
    if (!rowId || rowId === 'undefined' || rowId === 'null') {
      throw new Error('有効なIDが提供されていません');
    }
    
    console.log(`ID '${rowId}' のJSON情報からスライド生成を開始します...`);
    
    // スプレッドシートIDをプロパティから取得
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('TRACKING_SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('TRACKING_SPREADSHEET_IDが設定されていません');
    }
    
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('PDF処理状態');
    
    if (!sheet) {
      throw new Error('PDF処理状態シートが見つかりません');
    }
    
    // データを取得
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let fileName = '';
    let fileUrl = '';
    let jsonStr = '';
    
    // ヘッダー行をスキップしてrowIdを検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === rowId) {
        rowIndex = i + 1; // 1-indexed
        fileName = data[i][1];
        fileUrl = data[i][2];
        jsonStr = data[i][4]; // E列: JSON情報
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error(`ID '${rowId}' の行がスプレッドシートに見つかりません`);
    }
    
    if (!jsonStr) {
      throw new Error(`ID '${rowId}' のJSON情報が空です`);
    }
    
    console.log(`JSON情報を取得しました: ${jsonStr.substring(0, 100)}...`);
    
    // JSON文字列をパース
    let paperInfo;
    try {
      paperInfo = JSON.parse(jsonStr);
    } catch (error) {
      throw new Error(`JSON解析エラー: ${error}`);
    }
    
    // スライド生成
    const slideUrl = createSlide(paperInfo);
    
    // 処理状態を更新
    sheet.getRange(rowIndex, 8).setValue('完了'); // H列: 完了フラグ
    sheet.getRange(rowIndex, 9).setValue(slideUrl); // I列: スライドURL
    sheet.getRange(rowIndex, 7).setValue(new Date().toISOString()); // G列: 処理日時
    
    console.log(`スライド作成完了: ${slideUrl}`);
    return slideUrl;
  } catch (error) {
    console.error(`スプレッドシートからのスライド生成中にエラーが発生しました: ${error}`);
    throw error;
  }
}

/**
 * JSON文字列から直接スライドを生成する
 * @param {string} jsonString JSON形式の論文情報文字列
 * @returns {string} 生成したスライドのURL
 */
function createSlideFromJSON(jsonString) {
  try {
    console.log('JSON文字列からスライド生成を開始します...');
    
    if (!jsonString) {
      throw new Error('JSON文字列が提供されていません');
    }
    
    // JSON文字列をパース
    let paperInfo;
    try {
      paperInfo = JSON.parse(jsonString);
    } catch (error) {
      throw new Error(`JSON解析エラー: ${error}`);
    }
    
    // スライド生成
    return createSlide(paperInfo);
  } catch (error) {
    console.error(`JSON文字列からのスライド生成中にエラーが発生しました: ${error}`);
    throw error;
  }
}

/**
 * スプレッドシートの未処理エントリからスライドを生成する
 * JSON情報が存在し、スライドURLが未設定のエントリを処理
 * PDF_FOLDER_ID不要
 * @returns {Object} 処理結果の要約
 */
function processPendingEntriesFromSpreadsheet() {
  try {
    console.log('スプレッドシートの未処理エントリからスライド生成を開始します');
    
    // スプレッドシートIDを取得
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('TRACKING_SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('TRACKING_SPREADSHEET_IDが設定されていません');
    }
    
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('PDF処理状態');
    
    if (!sheet) {
      throw new Error('PDF処理状態シートが見つかりません');
    }
    
    // データを取得
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0]; // ヘッダー行
    
    // 列インデックスを確認（0ベース）
    const idIdx = 0; // PDF ID列
    const fileNameIdx = 1; // ファイル名列
    const jsonIdx = 4; // JSON情報列
    const slideUrlIdx = 8; // スライドURL列
    
    let processedCount = 0;
    let errorCount = 0;
    const processedEntries = [];
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fileId = row[idIdx];
      const fileName = row[fileNameIdx];
      const jsonStr = row[jsonIdx];
      const slideUrl = row[slideUrlIdx];
      
      const completionFlag = row[7]; // H列: 完了フラグ
      
      // JSON情報があり、H列（完了フラグ）が空の行を処理
      if (jsonStr && !completionFlag) {
        try {
          // IDチェック
          if (!fileId) {
            console.warn(`行 ${i+1}: ファイルIDが未定義です。スキップします。`);
            continue;
          }
          
          console.log(`未処理エントリを処理: ${fileName || '名前なし'} (ID: ${fileId})`);
          
          // スライド生成
          const url = createSlideFromSpreadsheetJSON(fileId);
          
          processedEntries.push({
            id: fileId, 
            name: fileName,
            url: url
          });
          
          processedCount++;
          
          // 1回の実行で処理するエントリ数を制限（GASの実行時間制限対策）
          if (processedCount >= 3) {
            console.log('処理エントリ数上限に達しました。次回の実行で残りを処理します。');
            break;
          }
        } catch (error) {
          console.error(`エントリ処理エラー (${fileId}): ${error}`);
          
          // エラー情報を更新
          const rowIndex = i + 1; // 1-indexed
          sheet.getRange(rowIndex, 6).setValue(error.toString()); // F列: エラー情報
          sheet.getRange(rowIndex, 8).setValue('エラー'); // H列: 完了フラグ
          sheet.getRange(rowIndex, 7).setValue(new Date().toISOString()); // G列: 処理日時
          
          errorCount++;
        }
      }
    }
    
    // 処理結果の要約を返す
    const result = {
      processed: processedCount,
      errors: errorCount,
      entries: processedEntries
    };
    
    console.log(`処理完了: ${processedCount}エントリ処理、${errorCount}件のエラー`);
    return result;
  } catch (error) {
    console.error(`スプレッドシート処理エラー: ${error}`);
    throw error;
  }
}

function testSlideCreation() {
  try {
    // テスト用の論文情報
    const testPaperInfo = {
      Title: "Test Paper Title",
      Japanese_Title: "テスト論文タイトル",
      Journal: "Test Journal Name",
      Citation: "Test Author et al. (2025) Test Journal, 1(1), 1-10",
      Limitation: "これはテスト用のリミテーションです。",
      ClinicalImplication: "これはテスト用の臨床的意義です。",
      ResearchImplication: "これはテスト用の研究的意義です。",
      PolicyImplication: "これはテスト用の政策的意義です。",
      Summary: "これはテスト用の要約です。"
    };
    
    // スライド生成
    const slideUrl = createSlide(testPaperInfo);
    return `テストスライドを生成しました: ${slideUrl}`;
  } catch (error) {
    return `テストエラー: ${error}`;
  }
}
