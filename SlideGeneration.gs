/**
 * Googleスライド生成と通知を行うモジュール
 */

// Slide URL列に値が入っているかを判定　あればその行はスキップ　なければ以下の処理

/**
 * OpenAI APIを使用してstructured outputを生成（JSONスキーマ形式）
 * @param {Object} article 記事情報
 * @returns {Object} パース済みの構造化データ
 * { Category, Time, Place, Person, "Summary of Article" }
 */
function generateStructuredOutput(article) {
  Logger.log(`generateStructuredOutput: start, article title = ${article.title}`);
  const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('openai_api');
  const OPENAI_MODEL = "gpt-4o-mini";

  // システムとユーザーのメッセージを定義
  const messages = [
    {
      role: "system",
      content: "You are an expert at structured data extraction. You will be given unstructured text from a research paper and should convert it into the given structure."
    },
    {
      role: "user",
      content: `以下の論文情報から、以下の構造化された情報を抽出してください：

タイトル: ${article.title}
要約: ${article.abstract}`
    }
  ];

  // JSONスキーマ形式での出力を指定（Category以外は日本語と英語の両方を返す）
  const payload = {
    model: OPENAI_MODEL,
    messages: messages,
    temperature: 0,
    response_format: {
      type: "json_schema",
      json_schema: {
        name: "research_paper_extraction",
        schema: {
          type: "object",
          properties: {
            "Thema": {
              type: "string",
              description: "タイトルと抄録から論文の内容を一言で表現してください。必要に応じて国名等の情報も含めてください。情報が得られない場合は『不明』と記入してください。"
            },
            "Category": {
              type: "string",
              enum: [
                "Case mangament",
                "Epidemiology",
                "Laboratory",
                "Public health",
                "ICT(infection control team)",
                "Vaccine"
              ],
              description: "論文が該当するカテゴリーを、上記の選択肢の中から1つ選んでください。情報が得られない場合は『不明』と記入してください。"
            },
            "Time": {
              type: "object",
              properties: {
                "ja": {
                  type: "string",
                  description: "日本語で、論文内で言及されている日時や調査期間（例：実験日、調査期間など）を記述してください。情報が得られない場合は『不明』と記入してください。"
                },
                "en": {
                  type: "string",
                  description: "In English, describe the date, time, or investigation period mentioned in the paper (e.g., the experimental date or study period). If the information is unavailable, answer 'unknown'."
                }
              },
              required: ["ja", "en"],
              additionalProperties: false
            },
            "Place": {
              type: "object",
              properties: {
                "ja": {
                  type: "string",
                  description: "日本語で、論文内で言及されている場所や地域を記述してください。情報が得られない場合は『不明』と記入してください。"
                },
                "en": {
                  type: "string",
                  description: "In English, describe the location or region mentioned in the paper. If the information is unavailable, answer 'unknown'."
                }
              },
              required: ["ja", "en"],
              additionalProperties: false
            },
            "Person": {
              type: "object",
              properties: {
                "ja": {
                  type: "string",
                  description: "日本語で、論文内で言及されている主要な人物または関係者を記述してください。情報が得られない場合は『不明』と記入してください。"
                },
                "en": {
                  type: "string",
                  description: "In English, describe the key person or relevant individuals mentioned in the paper. If the information is unavailable, answer 'unknown'."
                }
              },
              required: ["ja", "en"],
              additionalProperties: false
            },
            "Summary of Article": {
              type: "object",
              properties: {
                "ja": {
                  type: "string",
                  description: "日本語で、論文の要約を箇条書き形式で記述してください。専門用語を避け、誰にでも分かりやすい平易な表現で説明してください。情報が得られない場合は『不明』と記入してください。"
                },
                "en": {
                  type: "string",
                  description: "In English, provide a bullet-point summary of the article using plain language that avoids technical jargon. If the information is unavailable, answer 'unknown'."
                }
              },
              required: ["ja", "en"],
              additionalProperties: false
            }
          },
          required: ["Thema", "Category", "Time", "Place", "Person", "Summary of Article"],
          additionalProperties: false
        },
        strict: true
      }
    }
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  });

  const result = JSON.parse(response.getContentText());
  Logger.log('generateStructuredOutput: API response received');
  try {
    // result.choices[0].message.content にJSON文字列が入っているのでパースする
    const parsedResult = JSON.parse(result.choices[0].message.content);
    Logger.log('generateStructuredOutput: end, result parsed successfully');
    return parsedResult;
  } catch (error) {
    Logger.log(`generateStructuredOutput: error parsing result - ${error}`);
    throw new Error("Structured outputのパースに失敗しました: " + error);
  }
}

/**
 * Googleスライドを生成
 * @param {Object} structuredData 構造化データ
 * @param {Object} article 元の記事情報
 * @returns {string} スライドのURL
 */
function createGoogleSlide(structuredData, article) {
  Logger.log(`createGoogleSlide: start, article title = ${article.title}`);
  
  // 必要なプロパティを取得
  const folder_id = PropertiesService.getScriptProperties().getProperty('PDF_FOLDER_ID');
  const template_id = PropertiesService.getScriptProperties().getProperty('Slide_template_id');
  
  if (!folder_id || !template_id) {
    Logger.log('createGoogleSlide: 必要なプロパティが設定されていません');
    throw new Error('folder_idまたはSlide_template_idが設定されていません');
  }
  
  // ファイル名を「出版日_雑誌名_論文名」形式で作成（ファイル名用の日付：yyyy-MM-dd形式、表示用の日付：yyyy/MM/dd形式）
  const dateObj = new Date(article.pubDate);
  const fileNamePubDate = Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy-MM-dd');
  const displayPubDate = Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy/MM/dd');
  const fileName = `${fileNamePubDate}_${article.journal.replace(/[^\w\s]/g, '')}_${article.title.replace(/[^\w\s]/g, '')}`;
  
  try {
    // テンプレートファイルをコピー
    const templateFile = DriveApp.getFileById(template_id);
    const newFile = templateFile.makeCopy(fileName, DriveApp.getFolderById(folder_id));
    const presentation = SlidesApp.openById(newFile.getId());
    
    // プレースホルダーの置換用データを準備
    const replacements = {
      '{Thema}': structuredData['Thema'] || '不明',
      '{PublicationDate}': displayPubDate || '不明',
      '{Title}': article.title || '不明',
      '{Citation}': article.citation || '不明',
      '{Category}': structuredData['Category'] || '不明',
      '{Time}': (structuredData['Time'] && structuredData['Time'].ja) || '不明',
      '{Place}': (structuredData['Place'] && structuredData['Place'].ja) || '不明',
      '{Person}': (structuredData['Person'] && structuredData['Person'].ja) || '不明',
      '{Summary}': (structuredData['Summary of Article'] && structuredData['Summary of Article'].ja) || '不明'
    };
    
    // 最初のスライドのみを対象にプレースホルダーを置換
    const firstSlide = presentation.getSlides()[0];
    Object.entries(replacements).forEach(([placeholder, value]) => {
      firstSlide.replaceAllText(placeholder, value);
    });
    
    const slideUrl = presentation.getUrl();
    Logger.log(`createGoogleSlide: end, presentation created with URL ${slideUrl}`);
    return slideUrl;
    
  } catch (error) {
    Logger.log(`createGoogleSlide: エラーが発生しました: ${error.toString()}`);
    throw error;
  }
}

/**
 * 設定スプレッドシートから送信先メールアドレスを取得
 * @returns {string} カンマ区切りのメールアドレス
 */
function getRecipientEmails() {
  Logger.log('getRecipientEmails: start');
  const settingSpreadsheetId = PropertiesService.getScriptProperties().getProperty('setting_spreadsheet_id');
  const ss = SpreadsheetApp.openById(settingSpreadsheetId);
  const sheet = ss.getSheets()[0]; // 最初のシート
  const emails = sheet.getRange('D2').getValue();
  Logger.log(`getRecipientEmails: result: ${emails}`);
  return emails;
}

/**
 * 通知情報を整形
 * @param {Object} article 記事情報
 * @param {string} slideUrl スライドのURL
 * @returns {string} 整形された通知情報
 */
function formatNotificationInfo(article, slideUrl) {
  Logger.log(`formatNotificationInfo: start, article title = ${article.title}`);
  const notificationText = `
論文タイトル: ${article.title}
ジャーナル: ${article.journal}
出版日: ${article.pubDate}
スライドURL: ${slideUrl}
----------------------------------------`;
  Logger.log('formatNotificationInfo: end');
  return notificationText;
}

/**
 * まとめて通知を送信
 * @param {string[]} notifications 通知情報の配列
 * @param {string} recipientEmails カンマ区切りの送信先メールアドレス
 */
function sendSummaryNotification(notifications, recipientEmails) {
  Logger.log(`sendSummaryNotification: start, number of notifications = ${notifications.length}`);
  
  if (notifications.length === 0) {
    Logger.log('sendSummaryNotification: no notifications to send');
    return;
  }

  const subject = `新規スライド生成通知 (${new Date().toLocaleDateString()})`;
  const body = `本日生成された新規スライド一覧：

${notifications.join('\n')}`;

  GmailApp.sendEmail(recipientEmails, subject, body);
  Logger.log('sendSummaryNotification: email sent successfully');
}

/**
 * メイン実行関数
 */
function executeSlideGeneration() {
  Logger.log('executeSlideGeneration: start');
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('spreadsheet_id');
    if (!spreadsheetId) {
      Logger.log('executeSlideGeneration: spreadsheet_idが設定されていません');
      throw new Error('spreadsheet_idが設定されていません');
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('SearchResults');
    if (!sheet) {
      Logger.log('executeSlideGeneration: SearchResultsシートが見つかりません');
      throw new Error('SearchResultsシートが見つかりません');
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log(`executeSlideGeneration: 合計${lastRow}行を検出`);
    
    if (lastRow <= 1) {  // ヘッダーのみの場合
      Logger.log('executeSlideGeneration: 処理対象の記事がありません');
      return;
    }

    // Slide URL列の位置を特定または作成（大文字小文字無視）
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let slideUrlColumn = -1;
    for (let i = 0; i < headerRow.length; i++) {
      if ((headerRow[i] || "").toString().toLowerCase() === "slide url".toLowerCase()) {
        slideUrlColumn = i + 1;
        break;
      }
    }
    const startColumn = 10; // 構造化データの出力開始位置（J列）
    const outputHeaders = [
      "Thema", "Category", "Time", "Place", "Person", "Summary of Article"
    ];

    if (slideUrlColumn === -1) {  // Slide URL列が存在しない場合
      // 構造化データの右隣（startColumn + outputHeaders.length）にSlide URL列を追加
      slideUrlColumn = startColumn + outputHeaders.length;
      sheet.getRange(1, slideUrlColumn).setValue('Slide URL');
    }

    // 通知情報を格納する配列
    const notifications = [];

    // 2行目から最終行まで処理
    for (let currentRow = 2; currentRow <= lastRow; currentRow++) {
      try {
        // Slide URL列の値をチェック。値がある場合は、この行は既に処理済みなのでスキップ
        const slideUrlValue = sheet.getRange(currentRow, slideUrlColumn).getValue();
        if (slideUrlValue) {
          Logger.log(`${currentRow}行目: Slide URL列に値があるため、この行はスキップします。`);
          continue;
        }

        // 1～9列目のデータ取得（H列がCitation）
        const data = sheet.getRange(currentRow, 1, 1, 9).getValues()[0];
        const article = {
          pubmedId: data[1],    // PubMed ID (2列目)
          title: data[2],       // Title (3列目)
          journal: data[3],     // Journal (4列目)
          pubDate: data[4],     // Publication Date (5列目)
          authors: data[5],     // Authors (6列目)
          citation: data[7],    // Citation (H列, 8列目)
          abstract: data[8]     // Abstract (I列, 9列目)
        };

        // GPT-4 APIを用いてstructured outputの生成
        const structuredData = generateStructuredOutput(article);
        
        // structuredDataをスプレッドシートに追加列として保存
        // ヘッダーを常に更新する
        sheet.getRange(1, startColumn, 1, outputHeaders.length).setValues([outputHeaders]);
        
        // ヘルパー関数：値がオブジェクトの場合は "ja" と "en" の値を整形して返す
        function formatValue(value) {
          if (value && typeof value === 'object') {
            let parts = [];
            if (value.ja !== undefined) parts.push("ja: " + value.ja);
            if (value.en !== undefined) parts.push("en: " + value.en);
            return parts.join("; ");
          }
          return value;
        }
        
        // structuredDataの各値を、outputHeaders順に配列へ変換（オブジェクトの場合は整形して文字列に）
        const rowValues = outputHeaders.map(header => formatValue(structuredData[header] || ''));
        
        // 対象行の追加列（startColumn以降）にstructuredDataを書き込み
        sheet.getRange(currentRow, startColumn, 1, outputHeaders.length).setValues([rowValues]);
        
        // Googleスライドの作成
        const slideUrl = createGoogleSlide(structuredData, article);
        
        // 固定のSlide URL列にURLを保存
        sheet.getRange(currentRow, slideUrlColumn).setValue(slideUrl);
        
        // 通知情報を配列に追加
        notifications.push(formatNotificationInfo(article, slideUrl));
        
        Logger.log(`${currentRow}行目の処理が完了しました`);
      } catch (rowError) {
        Logger.log(`${currentRow}行目の処理中にエラーが発生しました: ${rowError.toString()}`);
        // 行個別のエラーは全体の処理を止めない
        continue;
      }
    }
    
    // 全ての処理が完了したら、まとめて通知を送信
    if (notifications.length > 0) {
      Logger.log(`executeSlideGeneration: ${notifications.length}件の通知を送信準備中`);
      const recipientEmails = getRecipientEmails();
      sendSummaryNotification(notifications, recipientEmails);
    } else {
      Logger.log('executeSlideGeneration: 送信する通知はありません');
    }
    
    Logger.log('executeSlideGeneration: 全ての未処理行の処理が完了しました');
  } catch (error) {
    Logger.log(`executeSlideGeneration: エラーが発生しました: ${error.toString()}`);
    throw error;
  }
}
