/**
 * PubMed検索と結果保存を行うモジュール（デバッグ用コード追加）
 */

/**
 * CSVファイルから設定を読み込む
 * @returns {Object} 設定オブジェクト
 */
function loadSearchSettings() {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SETTING_SPREADSHEET_ID');
    if (!spreadsheetId) {
      Logger.log("エラー: SETTING_SPREADSHEET_ID がスクリプト プロパティに設定されていません。");
      return null;
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('シート1');
    if (!sheet) {
      Logger.log("エラー: シート「シート1」が見つかりません。");
      return null;
    }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("エラー: 設定データが存在しません（2行目以降が空です）。");
      return null;
    }
    
    const settings = {
      searchKeywords: data[1][0],
      includeJournals: data[1][1] ? data[1][1].split(',') : [],
      excludeJournals: data[1][2] ? data[1][2].split(',') : [],
      email: data[1][3]
    };
    Logger.log("設定を読み込みました: %s", JSON.stringify(settings));
    return settings;
  } catch (e) {
    Logger.log("loadSearchSettings エラー: " + e);
    return null;
  }
}

/**
 * PubMedで検索し、eFetch APIで詳細情報（XML）を取得し、
 * eSummary API（XML形式）を利用して Publication date を抽出する関数
 * （eSearchでIDリストを取得 → eFetchでXML取得 → eSummaryで XML 解析して Publication date を取得 → 各項目を組み合わせ）
 * @param {string} query 検索クエリ
 * @returns {Array} 記事情報の配列（各記事はオブジェクトとして返す）
 */
function searchPubMed(query) {
  try {
    var PUBMED_API_BASE = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils/';
    
    // --- 1. eSearchでPubMed IDリストを取得 ---
    var searchUrl = PUBMED_API_BASE + 'esearch.fcgi?db=pubmed&term=' 
      + encodeURIComponent(query) + '&retmax=30&sort=edat&retmode=json';
    Logger.log('PubMed検索URL: ' + searchUrl);
    
    var searchResponse = UrlFetchApp.fetch(searchUrl);
    var searchResult = JSON.parse(searchResponse.getContentText());
    var ids = searchResult.esearchresult.idlist;
    Logger.log('取得したPubMed ID: ' + ids);
    
    if (ids.length === 0) {
      Logger.log('警告: 検索クエリ「' + query + '」でPubMed IDが見つかりませんでした。');
      return [];
    }
    
    // --- 2. eFetch APIで詳細情報（XML形式）を取得 ---
    var fetchUrl = PUBMED_API_BASE + 'efetch.fcgi?db=pubmed&id=' 
      + ids.join(',') + '&retmode=xml';
    Logger.log('eFetch URL: ' + fetchUrl);
    
    var fetchResponse = UrlFetchApp.fetch(fetchUrl);
    var xmlText = fetchResponse.getContentText();
    
    // XMLパース（efetchの結果）
    var document = XmlService.parse(xmlText);
    var root = document.getRootElement(); // <PubmedArticleSet>
    var pubmedArticles = root.getChildren('PubmedArticle');
    
    // --- 3. eSummary APIで Publication date を取得（XML形式） ---
var esummaryUrl = PUBMED_API_BASE + 'esummary.fcgi?db=pubmed&id=' 
  + ids.join(',') + '&retmode=xml';
Logger.log('eSummary URL: ' + esummaryUrl);

var esummaryResponse = UrlFetchApp.fetch(esummaryUrl);
var esumXml = XmlService.parse(esummaryResponse.getContentText());
var esumRoot = esumXml.getRootElement(); // <eSummaryResult>
var docSums = esumRoot.getChildren('DocSum');
var publicationDateMap = {};

for (var i = 0; i < docSums.length; i++) {
  var docSum = docSums[i];
  var uid = docSum.getChildText('Id');
  var items = docSum.getChildren('Item');
  var publicationDate = '';
  for (var j = 0; j < items.length; j++) {
    var item = items[j];
    var nameAttr = item.getAttribute('Name');
    if (nameAttr) {
      if (nameAttr.getValue() === 'History') {
        // History の中から medline を取得
        var historyItems = item.getChildren('Item');
        for (var k = 0; k < historyItems.length; k++) {
          var histItem = historyItems[k];
          var histNameAttr = histItem.getAttribute('Name');
          if (histNameAttr && histNameAttr.getValue() === 'medline') {
            // 日付文字列から時間部分を除去（YYYY/MM/DD HH:MM -> YYYY/MM/DD）
            publicationDate = histItem.getText().split(' ')[0];
            break;
          }
        }
        if (publicationDate !== '') {
          break;
        }
      } else if (nameAttr.getValue() === 'CreateDate') {
        // 万が一 Publication date がある場合はこちらを使用
        // 日付文字列から時間部分を除去（YYYY/MM/DD HH:MM -> YYYY/MM/DD）
        publicationDate = item.getText().split(' ')[0];
        break;
      }
    }
  }
  publicationDateMap[uid] = publicationDate;
}

    
    // --- 4. eFetch の XML 結果から他の項目を抽出 ---
    var articles = [];
    for (var i = 0; i < pubmedArticles.length; i++) {
      var pubmedArticle = pubmedArticles[i];
      var medlineCitation = pubmedArticle.getChild('MedlineCitation');
      if (!medlineCitation) continue;
      
      // PubMed ID
      var pmidElement = medlineCitation.getChild('PMID');
      var pubmedId = pmidElement ? pmidElement.getText() : '';
      
      var articleElement = medlineCitation.getChild('Article');
      
      // タイトル
      var title = articleElement ? articleElement.getChildText('ArticleTitle') : '';
      
      // ジャーナル情報と発行日
      var journalName = '';
      var pubDate = '';
      var journalElement = articleElement ? articleElement.getChild('Journal') : null;
      if (journalElement) {
        journalName = journalElement.getChildText('Title') || '';
        var journalIssue = journalElement.getChild('JournalIssue');
        if (journalIssue) {
          var pubDateElement = journalIssue.getChild('PubDate');
          if (pubDateElement) {
            var year = pubDateElement.getChildText('Year');
            var month = pubDateElement.getChildText('Month');
            var day = pubDateElement.getChildText('Day');
            if (year) {
              pubDate = year;
              if (month) pubDate += '/' + month;
              if (day) pubDate += '/' + day;
            } else {
              pubDate = pubDateElement.getChildText('MedlineDate') || '';
            }
          }
        }
      }
      
      // Publication dateは eSummary の XML から取得
      var publicationDate = publicationDateMap[pubmedId] || '';
      
      // 著者情報
      var authors = '';
      var authorListElement = articleElement.getChild('AuthorList');
      if (authorListElement) {
        var authorElements = authorListElement.getChildren('Author');
        var authorNames = [];
        for (var j = 0; j < authorElements.length; j++) {
          var author = authorElements[j];
          var lastName = author.getChildText('LastName');
          var foreName = author.getChildText('ForeName');
          if (lastName && foreName) {
            authorNames.push(foreName + ' ' + lastName);
          } else {
            var collectiveName = author.getChildText('CollectiveName');
            if (collectiveName) {
              authorNames.push(collectiveName);
            }
          }
        }
        authors = authorNames.join(', ');
      }
      
      // DOIの取得
      var doi = '';
      var articleIdList = pubmedArticle.getChild('PubmedData')?.getChild('ArticleIdList');
      if (articleIdList) {
        var articleIds = articleIdList.getChildren('ArticleId');
        for (var j = 0; j < articleIds.length; j++) {
          var idType = articleIds[j].getAttribute('IdType')?.getValue();
          if (idType === 'doi') {
            doi = articleIds[j].getText();
            break;
          }
        }
      }

      // Article Identifierの取得
      var articleIdentifier = '';
      if (articleIdList) {
        var articleIds = articleIdList.getChildren('ArticleId');
        for (var j = 0; j < articleIds.length; j++) {
          var idType = articleIds[j].getAttribute('IdType')?.getValue();
          if (idType === 'pii') {
            articleIdentifier = articleIds[j].getText();
            break;
          }
        }
      }

      // Abstract（要旨）
      var abstractText = '';
      var abstractElement = articleElement.getChild('Abstract');
      if (abstractElement) {
        var abstractTexts = abstractElement.getChildren('AbstractText');
        var abstracts = [];
        for (var j = 0; j < abstractTexts.length; j++) {
          abstracts.push(abstractTexts[j].getText());
        }
        abstractText = abstracts.join('\n');
      }

      // Citation用の日付フォーマット
      function formatCitationDate(dateStr) {
        if (!dateStr) return '';
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const parts = dateStr.split('/');
        if (parts.length === 3) {
          const month = months[parseInt(parts[1], 10) - 1];
          return `${parts[0]} ${month} ${parts[2]}`;
        }
        return dateStr;
      }

      // Citation生成
      var citation = '';
      if (authors) {
        const authorParts = authors.split(', ');
        const firstAuthor = authorParts[0].split(' ').pop(); // 最初の著者の姓を取得
        const authorText = authorParts.length > 1 ? `${firstAuthor} et al` : firstAuthor;
        const journalAbbrev = journalName.replace(/\./g, '').split(' ').map(word => word[0]).join('');
        const formattedDate = formatCitationDate(publicationDate);
        citation = `${authorText}. ${journalAbbrev}. ${formattedDate}`;
        if (articleIdentifier) {
          citation += `:${articleIdentifier}`;
        }
        if (doi) {
          citation += `. doi: ${doi}`;
        }
        citation += '. Epub ahead of print';
      }
      
      articles.push({
        pubmedId: pubmedId,
        title: title,
        journal: journalName,
        publicationDate: publicationDate, // eSummaryから取得した Publication date（例："1999/06/09"）
        authors: authors,
        doi: doi,
        citation: citation,
        abstract: abstractText
      });
    }
    
    Logger.log('取得した記事数: ' + articles.length);
    return articles;
    
  } catch (e) {
    Logger.log('searchPubMed エラー: ' + e);
    return [];
  }
}

/**
 * "yyyy/MM/dd" 形式の日付文字列をDateオブジェクトに変換するヘルパー関数
 * @param {string} dateStr 日付文字列（例："1999/06/09"）
 * @returns {Date}
 */
function parseDateString(dateStr) {
  var parts = dateStr.split('/');
  if (parts.length === 3) {
    return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  }
  return new Date(dateStr);
}

/**
 * 除外するジャーナルフィルタを適用（必要に応じてフィルタ内容を変更してください）
 * @param {Array} articles 記事リスト
 * @param {Array} excludeJournals 除外するジャーナル
 * @returns {Array} フィルタ適用後の記事リスト
 */
function applyJournalFilter(articles, excludeJournals) {
  try {
    if (!excludeJournals || excludeJournals.length === 0) {
      Logger.log("除外するジャーナルが指定されていないため、フィルタ処理をスキップします。");
      return articles;
    }
    
    const filteredArticles = articles.filter(article => {
      const journal = article.journal.toLowerCase();
      return !excludeJournals.some(exclude => journal.includes(exclude.toLowerCase()));
    });
    Logger.log("フィルタ適用後の記事数: %s（元記事数: %s）", filteredArticles.length, articles.length);
    return filteredArticles;
  } catch (e) {
    Logger.log("applyJournalFilter エラー: " + e);
    return articles;
  }
}

/**
 * 検索結果を指定したスプレッドシートに保存
 * シートの最初の列にはシリアル番号を追加し、Publication dateは文字列またはDateオブジェクトとして保存します。
 * ヘッダーの「Medline Date」を「Publication date」に変更しています。
 * @param {Array} articles 保存する記事リスト
 */
function saveToSpreadsheet(articles) {
  try {
    Logger.log("保存対象の記事数: %s", articles.length);
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SAVE_SPREADSHEET_ID');
    if (spreadsheetId === 'SAVE_SPREADSHEET_ID') {
      Logger.log("警告: 保存先のスプレッドシートIDが正しく設定されていません。");
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = 'SearchResults';
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("シート「%s」が存在しないため、新規作成します。", sheetName);
      sheet = ss.insertSheet(sheetName);
    }
    
    // ヘッダーがなければ追加（Serial列とPublication date列を含む）
    if (sheet.getLastRow() === 0) {
      Logger.log("ヘッダーが存在しないため、ヘッダー行を追加します。");
      sheet.appendRow(['Serial', 'PubMed ID', 'Title', 'Journal', 'Publication date', 'Authors', 'URL', 'Citation', 'Abstract']);
    }
    
    // 既存のPubMed IDを取得して重複チェック（PubMed IDは第2列に保存されているため）
    const lastRow = sheet.getLastRow();
    const existingIds = lastRow > 1 
      ? sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().map(String)
      : [];
    Logger.log("既存のPubMed ID: %s", existingIds);
    
    // 現在のシリアル番号を計算（ヘッダー行があるため、シリアル番号は (現在の行数-1)）
    let currentSerial = lastRow > 1 ? lastRow - 1 : 0;
    let addedCount = 0;
    articles.forEach(article => {
      if (!existingIds.includes(article.pubmedId)) {
        currentSerial++;
        sheet.appendRow([
          currentSerial,
          article.pubmedId,
          article.title,
          article.journal,
          article.publicationDate, // eSummaryから取得した Publication date を保存
          article.authors,
          article.doi ? 'https://doi.org/' + article.doi : '',
          article.citation,
          article.abstract
        ]);
        Logger.log("追加: PubMed ID %s (Serial: %s)", article.pubmedId, currentSerial);
        addedCount++;
      } else {
        Logger.log("重複のためスキップ: PubMed ID %s", article.pubmedId);
      }
    });
    Logger.log("新規追加記事数: %s", addedCount);
  } catch (e) {
    Logger.log("saveToSpreadsheet エラー: " + e);
  }
}

/**
 * メイン関数：設定の読み込み、PubMed検索、フィルタ、保存を実行
 */
function main() {
  try {
    Logger.log("----- スクリプト開始 -----");
    
    const settings = loadSearchSettings();
    if (!settings) {
      Logger.log("有効な設定が取得できなかったため、処理を終了します。");
      return;
    }
    
    const query = settings.searchKeywords;
    const articles = searchPubMed(query);
    if (articles.length === 0) {
      Logger.log("検索結果が得られなかったため、処理を終了します。");
      return;
    }
    
    const filteredArticles = applyJournalFilter(articles, settings.excludeJournals);
    if (filteredArticles.length === 0) {
      Logger.log("フィルタ後、記事が存在しなかったため、処理を終了します。");
      return;
    }
    
    saveToSpreadsheet(filteredArticles);
    Logger.log("----- スクリプト終了 -----");
  } catch (e) {
    Logger.log("main エラー: " + e);
  }
}
