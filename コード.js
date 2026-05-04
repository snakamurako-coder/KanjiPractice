/**
 * 厳密漢字判定ツール (KanjiVG) + 縦書きドリル — サーバー側
 *
 * 【初回起動】プロパティ未設定時、Drive に教材フォルダとサンプルブックを自動作成し、
 * 見出し行とサンプルデータを書き込み、スクリプトプロパティを登録します。
 *
 * 【単漢字・KanjiVG】
 * - KANJI_SHEET_ID … ストローク用スプレッドシート ID（未設定時は自動作成）
 *
 * 【縦書きドリル教材】
 * - KANJI_DRILL_FOLDER_ID … 教材フォルダ（未設定時は自動作成）
 * - KANJI_DRILL_BOOK_IDS … カンマ区切り ID（未設定時はサンプルブックを自動作成）
 * - 従来: KANJI_DRILL_BOOK_ID / DRILL_BOOK_ID / KANJI_DRILL_MATERIAL_FOLDER_ID も利用可
 *
 * シート1行目はヘッダ行。列:
 * セット, 漢字, 訓読みA_読み, 訓A_例文1, … 音D_例文2
 */

var SAMPLE_KANJI_PATH_ICHI =
  'M11,54.25c3.19,0.62,6.25,0.75,9.73,0.5c20.64-1.5,50.39-5.12,68.58-5.24c3.6-0.02,5.77,0.24,7.57,0.49';
var SAMPLE_KANJI_PATH_MIGI =
  'M53.5,21.5c0.62,1.12,0.69,2.23,0.25,4C49.62,42,39.5,61,25.25,74.25 | M13,42.15c1.9,0.56,5.9,0.52,7.79,0.34c23.41-2.24,49.76-5.74,67.67-6.3c3.24-0.1,6.45,0.31,9.17,0.81 | M41.75,66.5c0.75,0.75,1.35,1.93,1.54,2.95c0.94,5,2.38,16.66,3.07,22.76c0.24,2.15,0.39,2.8,0.39,3.54 | M43.25,68c5.25-0.5,29.75-3.25,37-3.75c1.75-0.12,3.24,1.52,3,2.75c-1,5.12-3.38,18-4.5,23.25 | M47,93.25c5.79-0.2,19.51-1.58,28.25-2.23c2.21-0.17,4.18-0.27,5.75-0.27';
var SAMPLE_KANJI_PATH_AME =
  'M25.75,22.37c1.87,0.4,4.47,0.62,6.32,0.4c11.68-1.39,28.28-3.77,41.25-4.64c2.49-0.17,4.37-0.12,7.18,0.28 | M15.5,41.25c1.25,1.5,1.66,3.26,1.89,5.19c1.24,10.69,2.19,26.61,2.66,36.31c0.13,2.7,0.2,5,0.2,6 | M18.25,44.25c1.42-0.09,62.76-5.33,69.5-6c2.5-0.25,4.61,1,4.5,3.75c-0.5,12.75-1.77,28.11-6,44.75c-1.88,7.38-5.38,1.88-8.5-1.25 | M52.25,23.5C53.31,24.56,54,26.25,54,28c0,0.82-0.25,37.8-0.43,53c-0.04,3.43-0.07,5.74-0.07,6.25 | M31,53.5c4.21,1.24,8.95,3.94,11.25,6 | M30.5,68.75c3.8,1.26,9.68,5.89,11.75,8 | M66.88,48.88c4.98,1.99,10.63,5.97,12.62,7.62 | M67.25,66.5c2.75,1,9,5.5,11,7.75';

var DRILL_HEADER_ROW = [
  'セット',
  '漢字',
  '訓読みA_読み',
  '訓A_例文1',
  '訓A_例文2',
  '訓読みB_読み',
  '訓B_例文1',
  '訓B_例文2',
  '訓読みC_読み',
  '訓C_例文1',
  '訓C_例文2',
  '訓読みD_読み',
  '訓D_例文1',
  '訓D_例文2',
  '音読みA_読み',
  '音A_例文1',
  '音A_例文2',
  '音読みB_読み',
  '音B_例文1',
  '音B_例文2',
  '音読みC_読み',
  '音C_例文1',
  '音C_例文2',
  '音読みD_読み',
  '音D_例文1',
  '音D_例文2',
];

var DRILL_SAMPLE_ROWS = [
  [
    1,
    '一',
    'ひと',
    '一つある。',
    'もう一つだ。',
    'いっ',
    'けしゴムを一こかした。',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    'いち',
    '一ばんだよ。',
    '一月はふゆだ。',
    'いつ',
    'きん一にまぜる。',
    'とう一する。',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
  ],
  [
    1,
    '右',
    'みぎ',
    '右をむく。',
    '右手を見る。',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    'う',
    'うせつする。',
    'さゆうを見る。',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
  ],
  [
    1,
    '雨',
    'あめ',
    '雨がふる。',
    '大雨がふる。',
    'あま',
    '雨水がでる。',
    '雨ぐもをみる。',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    'う',
    '雨てんちゅうしだ。',
    'ごう雨になる。',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
    '×',
  ],
];

function _propNonEmpty(props, key) {
  var v = props.getProperty(key);
  return v != null && String(v).trim() !== '';
}

function moveSpreadsheetToFolder_(spreadsheetId, folder) {
  DriveApp.getFileById(spreadsheetId).moveTo(folder);
}

function seedDrillSampleSpreadsheet_(ss) {
  var sh = ss.getSheets()[0];
  sh.setName('第1課サンプル');
  var rows = [DRILL_HEADER_ROW].concat(DRILL_SAMPLE_ROWS);
  sh.getRange(1, 1, rows.length, DRILL_HEADER_ROW.length).setValues(rows);
  sh.setFrozenRows(1);
}

function seedKanjiVGSampleSpreadsheet_(ss) {
  var sh = ss.getSheets()[0];
  sh.setName('サンプル');
  sh.getRange(1, 1, 1, 3).setValues([['漢字', 'Unicode', 'ストローク（KanjiVG）']]);
  var body = [
    ['一', '4.00E+00', SAMPLE_KANJI_PATH_ICHI],
    ['右', '53f3', SAMPLE_KANJI_PATH_MIGI],
    ['雨', '9.60E+09', SAMPLE_KANJI_PATH_AME],
  ];
  // getRange(行, 列, 行数, 列数) — データ行数は body.length のみ（1+ は誤りで次元不一致になる）
  sh.getRange(2, 1, body.length, 3).setValues(body);
  sh.setFrozenRows(1);
}

/**
 * 初回のみ: 教材フォルダ + ドリル用・KanjiVG 用のサンプルスプレッドシートを作成しプロパティを埋める
 */
function ensureKanjiPracticeBootstrap_() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(45000);
  } catch (e) {
    throw new Error(
      '初回教材セットアップのロック取得に失敗しました（同時アクセスが集中している可能性があります）。数十秒後にもう一度開いてください。' +
      (e && e.message ? ' 詳細: ' + e.message : '')
    );
  }
  try {
    var props = PropertiesService.getScriptProperties();

    var needDrillBook =
      !_propNonEmpty(props, 'KANJI_DRILL_BOOK_IDS') &&
      !_propNonEmpty(props, 'KANJI_DRILL_BOOK_ID') &&
      !_propNonEmpty(props, 'DRILL_BOOK_ID');
    var needKanjiBook = !_propNonEmpty(props, 'KANJI_SHEET_ID');

    if (!needDrillBook && !needKanjiBook) {
      return;
    }

    var folderId = props.getProperty('KANJI_DRILL_FOLDER_ID') || props.getProperty('KANJI_DRILL_MATERIAL_FOLDER_ID');
    var folder = null;
    if (folderId) {
      try {
        folder = DriveApp.getFolderById(String(folderId).trim());
      } catch (e) {
        folderId = null;
      }
    }
    if (!folder) {
      folder = DriveApp.createFolder('KanjiPractice教材');
      props.setProperty('KANJI_DRILL_FOLDER_ID', folder.getId());
    }

    if (needDrillBook) {
      var ssDrill = SpreadsheetApp.create('縦書きドリル・サンプル教材');
      moveSpreadsheetToFolder_(ssDrill.getId(), folder);
      seedDrillSampleSpreadsheet_(ssDrill);
      props.setProperty('KANJI_DRILL_BOOK_IDS', ssDrill.getId());
    }

    if (needKanjiBook) {
      var ssKanji = SpreadsheetApp.create('KanjiVG・ストロークサンプル');
      moveSpreadsheetToFolder_(ssKanji.getId(), folder);
      seedKanjiVGSampleSpreadsheet_(ssKanji);
      props.setProperty('KANJI_SHEET_ID', ssKanji.getId());
    }

    props.setProperty('KANJI_PRACTICE_BOOTSTRAPPED_AT', new Date().toISOString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * ウェブアプリとしてのエントリポイント
 */
function doGet() {
  ensureKanjiPracticeBootstrap_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('漢字書き順・美文字ドリル PRO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * ドリル用ブック一覧（スプレッドシート）
 */
function getDrillBooksList() {
  ensureKanjiPracticeBootstrap_();
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
    String(idsCsv)
      .split(',')
      .map(function (s) {
        return String(s).trim();
      })
      .filter(Boolean)
      .forEach(function (id) {
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
      return {
        error:
          '❌ 教材フォルダを読めません。KANJI_DRILL_FOLDER_ID とアクセス権を確認してください。\n' +
          e.message,
        books: [],
      };
    }
  }

  return {
    error:
      "❌ ドリル教材が未設定です。スクリプトプロパティに次のいずれかを設定してください:\n" +
      "・KANJI_DRILL_BOOK_IDS（カンマ区切り ID）\n" +
      '・KANJI_DRILL_BOOK_ID（単一）\n' +
      '・KANJI_DRILL_FOLDER_ID（「教材」フォルダ）',
    books: [],
  };
}

/**
 * 指定ブック内の全シートをドリル用レコードに変換して返す
 * @param {string} bookId スプレッドシート ID
 */
function getDrillData(bookId) {
  ensureKanjiPracticeBootstrap_();
  if (!bookId) {
    var prop2 = PropertiesService.getScriptProperties();
    bookId =
      prop2.getProperty('KANJI_DRILL_BOOK_ID') ||
      prop2.getProperty('DRILL_BOOK_ID');
  }
  if (!bookId) {
    return { error: '❌ getDrillData(bookId) にブック ID がありません。' };
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
  ensureKanjiPracticeBootstrap_();
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
      sheets: sheets,
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
  ensureKanjiPracticeBootstrap_();
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
