/**
 * 厳密漢字判定ツール (KanjiVG) + 縦書きドリル — サーバー側
 *
 * 【単漢字・KanjiVG】
 * - KANJI_SHEET_ID … ストローク用スプレッドシート ID
 *
 * 【縦書きドリル教材】次のいずれかでブック一覧を解決します（優先順）。
 * 1. KANJI_DRILL_BOOK_IDS … カンマ区切りのスプレッドシート ID（教材フォルダ内のブックを明示登録）
 * 2. KANJI_DRILL_BOOK_ID … 単一ブック ID（従来 DRILL_BOOK_ID もフォールバックで受理）
 * 3. KANJI_DRILL_FOLDER_ID … Drive 上の「教材」フォルダ ID。フォルダ内の Google スプレッドシートを一覧
 *
 * シート1行目はヘッダ行。想定列例:
 * セット, 漢字, 訓読みA_読み, 訓A_例文1, 訓A_例文2, … 音読みD_読み, 音D_例文1, 音D_例文2
 */

/**
 * ウェブアプリとしてのエントリポイント
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('漢字書き順・美文字ドリル PRO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * ドリル用ブック一覧（スプレッドシート）
 */
function getDrillBooksList() {
  const prop = PropertiesService.getScriptProperties();
  const idsCsv =
    prop.getProperty('KANJI_DRILL_BOOK_IDS') ||
    prop.getProperty('DRILL_BOOK_IDS');
  const folderId =
    prop.getProperty('KANJI_DRILL_FOLDER_ID') ||
    prop.getProperty('KANJI_DRILL_MATERIAL_FOLDER_ID');
  const singleId =
    prop.getProperty('KANJI_DRILL_BOOK_ID') ||
    prop.getProperty('DRILL_BOOK_ID');

  const books = [];

  if (idsCsv && String(idsCsv).trim()) {
    String(idsCsv).split(',').map(function (s) {
      return String(s).trim();
    }).filter(Boolean).forEach(function (id) {
      try {
        var ss = SpreadsheetApp.openById(id);
        books.push({ id: id, name: ss.getName() });
      } catch (e) {
        console.warn('skip book id ' + id + ': ' + e.message);
      }
    });
    return { success: true, books: books };
  }

  if (singleId && String(singleId).trim()) {
    try {
      var ss0 = SpreadsheetApp.openById(String(singleId).trim());
      books.push({ id: String(singleId).trim(), name: ss0.getName() });
      return { success: true, books: books };
    } catch (e) {
      return { error: '❌ KANJI_DRILL_BOOK_ID でブックを開けません。\n' + e.message, books: [] };
    }
  }

  if (folderId && String(folderId).trim()) {
    try {
      var folder = DriveApp.getFolderById(String(folderId).trim());
      var it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
      while (it.hasNext()) {
        var f = it.next();
        books.push({ id: f.getId(), name: f.getName() });
      }
      return { success: true, books: books };
    } catch (e) {
      return { error: '❌ 教材フォルダを読めません。KANJI_DRILL_FOLDER_ID とアクセス権を確認してください。\n' + e.message, books: [] };
    }
  }

  return {
    error:
      "❌ ドリル教材が未設定です。スクリプトプロパティに次のいずれかを設定してください:\n" +
      "・KANJI_DRILL_BOOK_IDS（カンマ区切り ID）\n" +
      '・KANJI_DRILL_BOOK_ID（単一）\n' +
      '・KANJI_DRILL_FOLDER_ID（「教材」フォルダ）',
    books: []
  };
}

/**
 * 指定ブック内の全シートをドリル用レコードに変換して返す
 * @param {string} bookId スプレッドシート ID
 */
function getDrillData(bookId) {
  if (!bookId) {
    var prop2 = PropertiesService.getScriptProperties();
    bookId =
      prop2.getProperty('KANJI_DRILL_BOOK_ID') ||
      prop2.getProperty('DRILL_BOOK_ID');
  }
  if (!bookId) {
    return { error: "❌ getDrillData(bookId) にブック ID がありません。" };
  }

  try {
    var ss = SpreadsheetApp.openById(String(bookId));
    var sheets = ss.getSheets();
    var drillData = {};

    sheets.forEach(function (sheet) {
      var name = sheet.getName();
      var values = sheet.getDataRange().getValues();
      if (values.length < 2) return;

      var headers = values[0].map(function (h) {
        return String(h == null ? '' : h).trim();
      });
      var records = [];
      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var empty = true;
        for (var c = 0; c < row.length; c++) {
          if (row[c] != null && String(row[c]).trim() !== '') {
            empty = false;
            break;
          }
        }
        if (empty) continue;

        var record = {};
        for (var j = 0; j < headers.length; j++) {
          if (!headers[j]) continue;
          var cell = row[j];
          record[headers[j]] = cell;
        }
        records.push(record);
      }
      drillData[name] = records;
    });

    return { success: true, data: drillData, bookId: bookId, bookName: ss.getName() };
  } catch (e) {
    return {
      error:
        '❌ ドリル用ブックの読み込みに失敗しました。ID と共有権を確認してください。\n' +
        e.message,
    };
  }
}

/**
 * アプリ起動時の初期設定情報を取得（KanjiVG 単漢字用ブック）
 */
function getAppInitData() {
  const prop = PropertiesService.getScriptProperties();
  const sheetId = prop.getProperty('KANJI_SHEET_ID');

  if (!sheetId) {
    throw new Error("❌ スクリプトプロパティ 'KANJI_SHEET_ID' が未設定です。");
  }

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheets = ss.getSheets().map(s => s.getName());

    return {
      bookName: ss.getName(),
      sheets: sheets
    };
  } catch (e) {
    throw new Error('❌ スプレッドシートにアクセスできません。IDと共有権限を確認してください。\n' + e.message);
  }
}

/**
 * 特定のシートから漢字データ（SVGパス）を抽出して整形する
 * @param {string} sheetName 読み込むシート名
 */
function getKanjiDataFromSheet(sheetName) {
  const prop = PropertiesService.getScriptProperties();
  const sheetId = prop.getProperty('KANJI_SHEET_ID');

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {};

    const values = sheet.getDataRange().getValues();
    const kanjiMap = {};

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const kanji = row[0];

      if (kanji && String(kanji).trim().length === 1) {
        const paths = [];

        for (let j = 2; j < row.length; j++) {
          const cellValue = row[j];
          if (!cellValue) continue;

          const strVal = String(cellValue).trim();
          if (strVal === '') continue;

          if (strVal.includes('|')) {
            strVal.split('|').forEach(p => {
              const cleaned = p.trim();
              if (cleaned.startsWith('M') || cleaned.startsWith('m')) {
                paths.push(cleaned);
              }
            });
          } else if (strVal.startsWith('M') || strVal.startsWith('m')) {
            paths.push(strVal);
          }
        }

        if (paths.length > 0) {
          kanjiMap[kanji] = paths;
        }
      }
    }

    return kanjiMap;

  } catch (e) {
    console.error('Data Fetch Error: ' + e.message);
    return {};
  }
}
