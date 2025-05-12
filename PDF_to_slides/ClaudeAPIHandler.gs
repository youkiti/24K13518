/**
 * ClaudeAPIHandler.gs
 * Claude APIを使用してPDFファイルの解析を行うGASスクリプト
 * 
 * 機能：
 * 1. PDFファイルをBase64エンコード
 * 2. Claude APIを使用してPDF内容を解析
 * 3. function callingを使用して構造化データを取得
 */

/**
 * PDFデータをClaudeに送信して解析する
 * @param {Blob} pdfBlob PDFファイルのBlob
 * @returns {Object} 抽出された論文情報
 */
function analyzePDFWithClaude(pdfBlob) {
  try {
    const startMsg = 'Claude APIでPDF解析を開始します...';
    console.log(startMsg);
    Logger.log(startMsg);
    
    // Claude APIキーを取得
    const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!apiKey) {
      throw new Error('CLAUDE_API_KEYが設定されていません');
    }
    
    // PDFをBase64エンコード
    const base64PDF = Utilities.base64Encode(pdfBlob.getBytes());
    
    // function callingのためのツール定義
    const tools = [
      {
        "name": "extract_paper_info",
        "description": "論文PDFから重要な情報を抽出する",
        "input_schema": {
          "type": "object",
          "properties": {
            "Title": {
              "type": "string",
              "description": "論文のオリジナルタイトル"
            },
            "Japanese_Title": {
              "type": "string",
              "description": "論文タイトルの日本語訳"
            },
            "Journal": {
              "type": "string",
              "description": "論文が掲載された雑誌名（英語のまま）"
            },
            "Citation": {
              "type": "string",
              "description": "論文の引用情報（著者、ジャーナル、出版年など）"
            },
            "Limitation": {
              "type": "string",
              "description": "論文の制限事項・限界点"
            },
            "ClinicalImplication": {
              "type": "string",
              "description": "論文に書かれた臨床的意義"
            },
            "ResearchImplication": {
              "type": "string",
              "description": "論文に書かれた研究的意義"
            },
            "PolicyImplication": {
              "type": "string",
              "description": "論文に書かれた政策的意義"
            },
            "Summary": {
              "type": "string",
              "description": "論文の要約"
            }
          },
          "required": ["Title", "Japanese_Title", "Journal", "Citation", "Limitation", "ClinicalImplication", "ResearchImplication", "PolicyImplication", "Summary"]
        }
      }
    ];
    
    // APIリクエスト準備
    const payload = {
      "model": "claude-3-7-sonnet-20250219",
      "max_tokens": 4000,
      "tools": tools,
      "tool_choice": {"type": "auto"},
      "system": "あなたは学術論文の専門的なアシスタントです。PDFの論文を解析し、指定されたツールを使用して必要な情報を抽出してください。英語論文の場合は日本語タイトルを適切に翻訳してください。特に抄録（Abstract）と考察（Discussion）セクション、そしてハイライト（Highlights）セクションがあれば、そこから各種Implicationとリミテーションを日本語で抽出してください。情報が見つからない場合は「記載なし」と返してください。",
      "messages": [
        {
          "role": "user",
          "content": [
            {
              "type": "document",
              "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": base64PDF
              }
            },
            {
              "type": "text",
              "text": "この論文を解析し、必要な情報を抽出してください。特に以下の点に注意してください：\n\n1. 論文のタイトルと日本語訳\n2. 雑誌名（英語のままで抽出）\n3. 引用情報（著者、ジャーナル、年など）\n4. 抄録と考察部分を中心に、LimitationとClinical/Research/Policy Implicationを探してください\n5. 該当する情報が明示的に書かれていない場合は「記載なし」と返してください\n6. すべての項目を日本語で抽出してください（タイトル原文と雑誌名を除く）"
            }
          ]
        }
      ]
    };
    
    // APIリクエスト送信
    console.log('Claude APIへリクエスト送信中...');
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    // レスポンス処理
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      const errMsg = `API応答エラー (${responseCode}): ${response.getContentText()}`;
      Logger.log(errMsg);
      throw new Error(errMsg);
    }
    
    const responseJson = JSON.parse(response.getContentText());
    const responseMsg = 'Claude APIからの応答を受信しました';
    console.log(responseMsg);
    Logger.log(responseMsg);
    
    // ツール使用の結果を解析
    if (responseJson.content && responseJson.content.length) {
      for (const item of responseJson.content) {
        if (item.type === 'tool_use' && item.name === 'extract_paper_info') {
          const successMsg = '論文情報の抽出に成功しました';
          console.log(successMsg);
          Logger.log(successMsg);
          return item.input; // 抽出された論文情報を返す
        }
      }
    }
    
    // ツール使用の結果が見つからない場合
    const notFoundMsg = 'Claude APIからの応答に必要な情報が含まれていません';
    Logger.log(notFoundMsg);
    throw new Error(notFoundMsg);
  } catch (error) {
    const errorMsg = `PDF解析中にエラーが発生しました: ${error}`;
    console.error(errorMsg);
    Logger.log(errorMsg);
    throw error;
  }
}

/**
 * DriveからPDFファイルを取得してClaudeに解析させる
 * @param {string} fileId GoogleドライブのファイルID
 * @returns {Object} 抽出された論文情報
 */
function processPDFFromDrive(fileId) {
  try {
    const startMsg = `ファイルID ${fileId} のPDF処理を開始します`;
    console.log(startMsg);
    Logger.log(startMsg);
    
    // ファイルをBlobとして取得
    const file = DriveApp.getFileById(fileId);
    const pdfBlob = file.getBlob();
    
    // Claude APIで解析
    const paperInfo = analyzePDFWithClaude(pdfBlob);
    
    // ファイル名を含む情報を追加
    paperInfo.fileName = file.getName();
    
    console.log(`ファイル "${file.getName()}" の処理が完了しました`);
    return paperInfo;
  } catch (error) {
    console.error(`ファイル処理エラー: ${error}`);
    throw error;
  }
}

/**
 * PDF_FOLDER_IDから指定されたフォルダ内のPDFファイルを取得して処理
 * 単体実行用のエントリーポイント
 */
function processNextPDF() {
  try {
    console.log('次の未処理PDFの処理を開始します');
    
    // 必要なプロパティを取得
    const pdfFolderId = PropertiesService.getScriptProperties().getProperty('PDF_FOLDER_ID');
    const processedFolderId = PropertiesService.getScriptProperties().getProperty('PROCESSED_FOLDER_ID');
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('TRACKING_SPREADSHEET_ID');
    
    // プロパティチェック
    if (!pdfFolderId) {
      throw new Error('PDF_FOLDER_IDが設定されていません');
    }
    if (!processedFolderId) {
      throw new Error('PROCESSED_FOLDER_IDが設定されていません');
    }
    if (!spreadsheetId) {
      throw new Error('TRACKING_SPREADSHEET_IDが設定されていません');
    }
    
    // フォルダを取得
    const pdfFolder = DriveApp.getFolderById(pdfFolderId);
    const processedFolder = DriveApp.getFolderById(processedFolderId);
    
    // PDFファイルを取得
    const pdfFiles = pdfFolder.getFilesByType(MimeType.PDF);
    
    // トラッキングスプレッドシートを初期化
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = ss.getSheetByName('PDF処理状態');
    
    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet('PDF処理状態');
      sheet.appendRow(['PDF ID', 'ファイル名', 'ファイルURL', '処理状態', 'JSON情報', 'エラー情報', '処理日時']);
      console.log('「PDF処理状態」シートを新規作成しました');
      
      // PDFファイルがなければ終了
      if (!pdfFiles.hasNext()) {
        console.log('処理対象のPDFファイルがありません');
        return;
      }
    } else {
      // 既存シートの場合、列名が設定されているか確認
      const headerRow = sheet.getRange(1, 1, 1, 7).getValues()[0];
      
      // 列名が設定されていないか確認（1列目が空など）
      if (!headerRow[0]) {
        // 列名が未設定の場合は1行目に挿入
        sheet.insertRowBefore(1);
        sheet.getRange(1, 1, 1, 7).setValues([['PDF ID', 'ファイル名', 'ファイルURL', '処理状態', 'JSON情報', 'エラー情報', '処理日時']]);
        console.log('既存シートに列名を追加しました');
      }
    }
    
    // 処理済みファイルIDを取得
    const processedIds = getProcessedFileIds(sheet);
    
    // PDFファイルを処理
    while (pdfFiles.hasNext()) {
      const file = pdfFiles.next();
      const fileId = file.getId();
      const fileName = file.getName();
      
      // 処理済みならスキップ
      if (processedIds.has(fileId)) {
        console.log(`ファイルはすでに処理済み: ${fileName}`);
        continue;
      }
      
      try {
        console.log(`"${fileName}" (${fileId}) の処理を開始します`);
        
        // 処理状態を更新
        updateFileStatus(sheet, fileId, fileName, file.getUrl(), '処理中');
        
        // PDFを処理
        const paperInfo = processPDFFromDrive(fileId);
        
        // 処理済みフォルダにファイルを移動
        file.moveTo(processedFolder);
        console.log(`"${fileName}" を処理済みフォルダに移動しました`);
        
        // 処理状態とJSON情報を更新
        updateFileStatus(sheet, fileId, fileName, file.getUrl(), '処理済', JSON.stringify(paperInfo));
        
        console.log(`"${fileName}" の処理が完了しました`);
        return; // 1件だけ処理して終了
      } catch (error) {
        // エラータイプの識別と適切なメッセージ設定
        let errorStatus = 'エラー';
        const errorMessage = error.toString();
        
        // ファイルサイズエラーの検出
        if (errorMessage.includes('exceeds the maximum size') || 
            errorMessage.includes('file too large') || 
            errorMessage.includes('size limit exceeded')) {
          errorStatus = 'サイズ超過';
        } 
        // 著作権エラーの検出（Claude APIからの応答に基づく）
        else if (errorMessage.includes('copyright') || 
                errorMessage.includes('著作権') || 
                errorMessage.includes('intellectual property')) {
          errorStatus = '著作権制限';
        }
        
        // エラー情報を更新
        updateFileStatus(sheet, fileId, fileName, file.getUrl(), errorStatus, '', errorMessage);
        console.error(`"${fileName}" の処理中にエラーが発生しました: ${errorMessage}`);
        return;
      }
    }
    
    const message = '処理対象のPDFがありません';
    console.log(message);
    Logger.log(message);
    return message;
  } catch (error) {
    const errorMsg = `全体処理エラー: ${error}`;
    console.error(errorMsg);
    Logger.log(errorMsg);
    return errorMsg;
  }
}

/**
 * 処理済みファイルIDのセットを取得
 * @param {Sheet} sheet 処理状態を記録しているシート
 * @returns {Set} 処理済みファイルIDのセット
 */
function getProcessedFileIds(sheet) {
  try {
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
 * @param {Sheet} sheet 処理状態を記録しているシート
 * @param {string} fileId ファイルID
 * @param {string} fileName ファイル名
 * @param {string} fileUrl ファイルURL
 * @param {string} status 処理状態
 * @param {string} jsonInfo JSON情報（オプション）
 * @param {string} errorInfo エラー情報（オプション）
 */
function updateFileStatus(sheet, fileId, fileName, fileUrl, status, jsonInfo = '', errorInfo = '') {
  try {
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
      // 既存行の更新 - 同じセルに処理状態を上書き
      sheet.getRange(rowIndex, 4).setValue(status); // 処理状態（処理中→処理済などを上書き）
      if (jsonInfo) sheet.getRange(rowIndex, 5).setValue(jsonInfo); // JSON情報
      if (errorInfo) sheet.getRange(rowIndex, 6).setValue(errorInfo); // エラー情報
      sheet.getRange(rowIndex, 7).setValue(timestamp); // 処理日時
    } else {
      // 新規行の追加
      sheet.appendRow([fileId, fileName, fileUrl, status, jsonInfo, errorInfo, timestamp]);
    }
    
    console.log(`ファイル ${fileName} の状態を「${status}」に更新しました`);
  } catch (error) {
    console.error(`状態更新エラー: ${error}`);
  }
}

/**
 * 手動実行用のエントリーポイント
 */
function runManually() {
  Logger.log("手動処理を開始します");
  processNextPDF();
  Logger.log("手動処理が完了しました");
  return "処理完了"; // 戻り値を追加
}

/**
 * トリガーセットアップ用関数
 * GASウェブエディタから手動で実行することで、1時間ごとのトリガーを設定
 */
function setupHourlyTrigger() {
  try {
    // 既存のトリガーを削除（重複防止）
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'processNextPDF') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // 1時間ごとのトリガーを設定
    ScriptApp.newTrigger('processNextPDF')
      .timeBased()
      .everyHours(1)
      .create();
    
    const message = '1時間ごとのトリガーを設定しました';
    Logger.log(message);
    return message;
  } catch (error) {
    const errorMsg = `トリガー設定エラー: ${error}`;
    Logger.log(errorMsg);
    return errorMsg;
  }
}

/**
 * APIテスト用の関数
 */
function testClaudeAPI() {
  try {
    // テスト用のAPIキー確認
    const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!apiKey) {
      return 'CLAUDE_API_KEYが設定されていません';
    }
    
    // 簡単なAPIリクエスト
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
      },
      payload: JSON.stringify({
        'model': 'claude-3-7-sonnet-20250219',
        'max_tokens': 100,
        'messages': [
          {
            'role': 'user',
            'content': 'Hello, Claude. Can you hear me?'
          }
        ]
      }),
      muteHttpExceptions: true
    });
    
    const result = `Claude API接続テスト: ${response.getResponseCode()} ${response.getContentText().substring(0, 100)}...`;
    Logger.log(result); // 実行ログに出力
    return result;
  } catch (error) {
    const errorMsg = `APIテストエラー: ${error}`;
    Logger.log(errorMsg); // エラーログも実行ログに出力
    return errorMsg;
  }
}
