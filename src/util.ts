/**
 * プロパティ情報を表示
 */
function putProp() {
  // プロパティ一覧文字列の生成
  // const uuid = SETTING?.uuid || '';
  const props = { ...SCRIPT_PROP.getProperties() };
  const prop_str = Object.keys(props).sort()
    .map((key) =>
      `${key} : ${props[key].slice(0, 200)}`
    );

  const caches = { ...SCRIPT_CACHE.getAll(SETTING.cache_keys.map((key) => key)) };
  const cache_str = Object.keys(caches).sort()
    .map((key) =>
      `${key} : ${caches[key].slice(0, 200)}`
    );

  // ポップアップに表示
  const html = HtmlService
    .createHtmlOutput(`<pre>${[...prop_str, ...cache_str].join('\n')}</pre>`)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'プロパティ情報');
}

/** 設定値・トリガの初期化 */
function beginInit() {
  // 確認ダイアログを表示
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('初期化', '設定を全て初期化します', ui.ButtonSet.OK_CANCEL);

  if (resp !== ui.Button.OK)
    return;

  // 設定初期化
  SETTING.init();
  // トリガー削除
  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t));

  return '初期化が完了しました';
}

/**
 * ログをセットする
 * @param {string[]} text 文字列
 * @param {boolean} force キャッシュの書き込みフラグ
 */
function putLog(texts: string[], force = true) {
  console.log(texts);
  const ss = SpreadsheetApp.getActive();
  const s = ss.getSheetByName('log');
  if (s) {
    const ts = Utilities.formatDate(new Date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    // 列数・文字数をカット
    const trimmed = texts.slice(0, 20).map((text) => text.slice(0, 2000));
    // ログ用キャッシュに格納
    const { logs } = SETTING;
    logs.push([ts, Session.getActiveUser().getEmail() || 'ー', ...trimmed]);
    SETTING.logs = logs;

    // 一定数を超えたら書き込み
    if (force || logs.length > 10)
      flushLog();

    // 最大値チェック
    const MAX_LOG_NUM = 30000;
    if (s.getLastRow() > MAX_LOG_NUM) {
      const range = {
        sheetId: s.getSheetId(),
        dimension: 'ROWS',
        startIndex: 1,
        endIndex: MAX_LOG_NUM / 2
      };
      Sheets.Spreadsheets?.batchUpdate(
        { requests: [{ deleteDimension: { range } }] },
        ss.getId()
      );
    }
  }
}

/** キャッシュされたログを書き出す */
function flushLog() {
  const { logs } = SETTING;
  if (logs.length) {
    const ss = SpreadsheetApp.getActive();
    Sheets.Spreadsheets?.Values?.append(
      { values: logs },
      ss.getId(),
      'log!A:Z',
      { valueInputOption: 'USER_ENTERED' }
    );
    SETTING.logs = [];
  }
}

/** 列番号をアルファベットに変換 */
function numeric2Colname(num: number) {
  /** アルファベット総数 */
  const RADIX = 26;
  /** Aの文字コード */
  const A = 'A'.charCodeAt(0);

  let n = num;
  let s = '';
  while (n >= 1) {
    n--;
    s = String.fromCharCode(A + (n % RADIX)) + s;
    n = Math.floor(n / RADIX);
  }
  return s;
}

/** アルファベットを列番号に変換 */
function colname2number(column_name: string) {
  const base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const column_number = column_name.toUpperCase().split('').reduce((acc, c) => {
    acc = acc * 26 + base.indexOf(c) + 1;
    return acc;
  }, 0);
  return column_number;
}

/**
 * １次元配列を、指定した要素数で分割した２次元配列に変換
 * @param[in] {Array<T>} arr １次元配列
 * @param[in] {number} chunk 分割する要素数
 */
function bunch<T>(arr: Array<T>, chunk: number) {
  return [...Array(Math.ceil(arr.length / chunk))].map((_, i) =>
    arr.slice(i * chunk, (i + 1) * chunk)
  );
}

/**
 * １次元配列から重複を取り除いた新しい配列を作成する
 * @param[in] {Array<T>} arr １次元配列
 */
function unique<T>(arr: Array<T>) {
  return Array.from(new Set(arr));
}

/**
 * 日付のシリアル値をUnixTimeに変換
 */
function dateSerialToDate(serialDate: number | string) {
  // スプレッドシートの日付シリアル値は1900年1月1日を基準とする
  const dateOrigin = new Date(1899, 11, 30); // JavaScriptのDateは0ベースの月を使用するため、12月は11となります
  const msecPerDay = 24 * 60 * 60 * 1000; // 1日あたりのミリ秒
  if (typeof serialDate === 'string')
    return new Date(dateOrigin.getTime() + parseFloat(serialDate) * msecPerDay);
  return new Date(dateOrigin.getTime() + serialDate * msecPerDay);
}

/**
 * 日付のシリアル値をUnixTimeに変換
 */
function dateToDateSerial(date: Date) {
  // スプレッドシートの日付シリアル値は1900年1月1日を基準とする
  const dateOrigin = new Date(1899, 11, 30); // JavaScriptのDateは0ベースの月を使用するため、12月は11となります
  const msecPerDay = 24 * 60 * 60 * 1000; // 1日あたりのミリ秒
  return (date.getTime() - dateOrigin.getTime()) / msecPerDay;
}

/** オブジェクトからクエリパラメータ用の文字列を生成 */
function generateQueryString(query: { [key: string]: string }) {
  const query_param = Object.keys(query)
    .map((k) => encodeURIComponent(k) + '=' + encodeURIComponent(query[k]))
    .join('&amp;');
  return query_param;
}

/** シートAPIのチートシート */
function sheetsApiCheatSheet() {
  const ss = SpreadsheetApp.getActive();
  const s = ss.getSheetByName('log');

  // 行の追加
  const append = () => {
    Sheets.Spreadsheets?.Values?.append(
      { values: [] },
      ss.getId(),
      'log!A:A',
      { valueInputOption: 'USER_ENTERED' }
    );
  };

  // 値の取得
  const batchGet = () => {
    const resp = Sheets.Spreadsheets?.Values?.batchGet(
      ss.getId(),
      {
        ranges: ['main!A:A'],
        // valueRenderOption: 'FORMULA',  数式を取得する場合
        // dateTimeRenderOption: 'FORMATTED_STRING',  日付をシリアル値で取らないために指定
      }
    );

    // 見出し行をキーとした配列に変換
    const df = resp?.valueRanges ? (resp.valueRanges[0].values || [[]]) : [[]];
    const [head] = df;
    const records = namingCellValues(df);
  };

  // 値のクリア
  const batchClear = () => {
    Sheets.Spreadsheets?.Values?.batchClear(
      {
        ranges: ['logs!A1:B2']
      },
      ss.getId()
    );
  };

  // 値の更新
  const batchUpdate = () => {
    Sheets.Spreadsheets?.Values?.batchUpdate(
      {
        valueInputOption: 'USER_ENTERED',
        data: [
          { range: 'logs!A1:B2', values: [[1, 2], [3, 4]], },
          { range: 'logs!A1:B2', values: [[1, 2], [3, 4]], },
        ]
      },
      ss.getId()
    );
  };

  // 行の削除
  const deleteRows = () => {
    const range = {
      sheetId: s?.getSheetId(),
      dimension: 'ROWS',
      startIndex: 1,
      endIndex: 30000
    };
    Sheets.Spreadsheets?.batchUpdate(
      { requests: [{ deleteDimension: { range } }] },
      ss.getId()
    );
  };
}

/**
 * 二次元のテーブルデータを見出しをキーとしたオブジェクト配列に変換する
 */
function namingCellValues(df: CellValue[][]) {
  const [head, ...values] = df;
  const records = (values as CellValue[][]).map((r) =>
    (head as CellValue[]).reduce((acc, value, i) => {
      acc[value.toString()] = r[i]?.toString() || '';
      return acc;
    }, {} as { [key: string]: string })
  );
  return records;
}

function showSideBarFromHTML(title: string, file: string) {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService
    .createTemplateFromFile(file)
    .evaluate()
    .setHeight(700)
    .setTitle(title);
  ui.showSidebar(html);
}

function hasTrigger(func: string) {
  return ScriptApp.getProjectTriggers().some((t) => t.getHandlerFunction() === func);
}

function deleteTrigger(func: string) {
  putLog(['DeleteTrigger', func]);
  return ScriptApp.getProjectTriggers().map((t) => (t.getHandlerFunction() === func) && (ScriptApp.deleteTrigger(t)));
}

function formatDate(date: GoogleAppsScript.Base.Date, format: string) {
  const str = Utilities.formatDate(date, 'JST', format)
    .replace(/(Sun|Sunday)/, '日')
    .replace(/(Mon|Monday)/, '月')
    .replace(/(Tue|Tuesday)/, '火')
    .replace(/(Wed|Wednesday)/, '水')
    .replace(/(Thu|Thursday)/, '木')
    .replace(/(Fri|Friday)/, '金')
    .replace(/(Sat|Saturday)/, '土');
  return str;
}

/** 現在選択中の行番号を取得 */
function selectedRows() {
  const rows = SpreadsheetApp.getActiveSheet()
    ?.getActiveRangeList()
    ?.getRanges()
    .map((range) => {
      const row = { start: range.getRow(), end: range.getLastRow(), };
      return [...Array(row.end - row.start + 1)].map((_, i) => (row.start + i));
    }).flat() || [];
  return rows;
}

function createAfterTrigger(func: string, after_msec: number) {
  putLog(['NewTrigger', func]);
  return ScriptApp.newTrigger(func).timeBased().after(after_msec).create();
}

function include(file: string) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

/** 画像URLからファイルをDriveに保存 */
function downloadImageToDrive(url: string) {
  putLog(['DownloadImageToDrive', url]);

  // 格納先フォルダの取得(スプシのあるフォルダ)
  let dirs = DriveApp
    .getFileById(SpreadsheetApp.getActive().getId())
    .getParents();

  if (!dirs.hasNext())
    throw new Error('フォルダにアクセス出来ません');

  // 無ければ作成
  const parent_dir = dirs.next();
  dirs = parent_dir.searchFolders('title contains "img"');
  const dir = dirs.hasNext() ? dirs.next() : parent_dir.createFolder('img');
  dir.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // ファイルを作成
  const img = UrlFetchApp.fetch(url).getBlob();
  if (!img.getContentType()?.includes('image'))
    throw new Error('画像ファイルをダウンロード出来ませんでした');

  const file = dir.createFile(img);
  return file;
  // ダウンロード画像へのアクセス
  // https://drive.google.com/uc?export=download&id={file.id}
}

/** 改行付きCSVのParser */
function csv_parse(str: string) {
  const RE = new RegExp('"(?:[^"]|"")*"|"(?:[^"]|"")*$|[^,\\r\\n]+|\\r?\\n|\\r|,+', 'g');
  const csv = [['']];
  let m;

  str = str.replace(/[\r?\n|\r]+$/, ''); //最後の改行を除去
  while ((m = RE.exec(str)) && m) {
    const elem = m[0];
    const row = csv[csv.length - 1];
    const c = elem.charAt(0);

    // ,から始まる列 => 列を追加
    if (c === ',')
      [...Array(elem.length)].map((_) => row.push(''));
    // 改行から始まる列 => 行を追加
    else if (c === '\n' || c === '\r\n' || c === '\r')
      csv.push(['']);
    // quoteで始まる列 => quote除外した文字列をセット
    else if (c === '"')
      row[row.length - 1] = elem.replace(/^"|"$/g, '').replace(/""/g, '"');
    // その他 => 文字列をセット
    else
      row[row.length - 1] = elem;
  }

  return csv;
}

/** メール送信 */
function sendMail(recipient: string, name: string, subject: string, body: string) {
  // Quotaのために、下書きを作成して送信
  const draft = GmailApp.createDraft(
    recipient,
    subject,
    body, {
    name,
    // from: システムオーナーのメールアドレスとなるので注意
    replyTo: 'noreply@gmail.com',
  });
  GmailApp.getDraft(draft.getId()).send();

  putLog(['SendEMail', JSON.stringify({
    recipient,
    subject,
    body,
    name
  }, null, 2), `Quota: ${MailApp.getRemainingDailyQuota()}`]);
}

/** XML形式の文字列をオブジェクトに変換する */
function xmlToObj(xml: string) {
  const doc = XmlService.parse(xml);
  const elementToObject = (elem: GoogleAppsScript.XML_Service.Element) => {
    const result: any = {};

    // Attributesを取得
    elem.getAttributes().forEach((attr) =>
      result[attr.getName()] = attr.getValue()
    );

    // Child Elementを取得
    elem.getChildren().forEach((child) => {
      // 再帰的にもう一度この関数を実行して判定
      const value = elementToObject(child);

      // XMLをJSONに変換する
      const key = child.getName();
      result[key] = result[key] ?
        [...result[key], value] :
        [value];
    });

    // タグ内のテキストデータを取得
    elem.getText() && (result['Text'] = elem.getText());
    return result;
  };
  const root = doc.getRootElement();
  const obj = {
    [root.getName()]: elementToObject(root),
  };
  return obj;
}

