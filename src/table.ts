/** 指定した範囲のテーブルを読み書きするための抽象クラス */
class Table<THeader extends iHeader, TRow extends iRow, THash extends iHash> {
  /** 見出し行番号 */
  head_row: number;
  /** 最終行番号 */
  tail_row: number;
  /** 先頭列 */
  head_col: string;
  /** 最終列 */
  tail_col: string;
  /** シート名 */
  sheet: string;
  /** テーブル範囲 */
  range: string;
  /** Hashの主キー */
  primary_key: string;
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet;

  /** 見出し */
  head: THeader[] = [];
  /** データオブジェクト */
  records: TRow[] = [];
  /** データ実体 */
  hashes: THash[] = [];

  /** batchGetのレンダリングオプション */
  valueRenderOption: 'FORMATTED_VALUE' | 'UNFORMATTED_VALUE' | 'FORMULA';

  /** 具象Rowオブジェクトの生成処理 */
  rowFactory: (head: THeader[], hash: THash, row?: number) => TRow;

  /** Type Guard */
  static isTRow<TRow extends iRow>(value: any): value is TRow {
    return (value.row !== undefined && typeof value.row === 'number');
  }

  /**
   * Tableクラスコンストラクタ
   * @param range テーブル範囲文字列
   * @param primary_key 主キーのカラム名文字列
   * @param rowFactory 具象Rowオブジェクト生成用処理
   */
  constructor(range: string, primary_key: string, rowFactory: (head: THeader[], hash: THash, row?: number) => TRow
    = (head, hash, row) => {
      const primary_key = this.primary_key;
      const row_object: iRow = {
        head,
        hash,
        row: row || NaN,
        /** Hashを見出しの順番に従った配列に変換する */
        toValues() {
          return this.head.map((col) => hash[col] === 'NaN' ? undefined : hash[col]);
        },
        /** 指定したHashが一致しているか比較 */
        isEqual(other: THash) {
          return this.hash[primary_key] === other[primary_key];
        }
      }
      return row_object as TRow;
    }) {
    const [sheet, a1note] = range.split('!');
    this.sheet = sheet;
    const head_row = parseInt(a1note.split(':')[0].replace(/[^\d]+/, ''));
    this.head_row = isNaN(head_row) ? 1 : head_row;
    this.tail_row = parseInt(a1note.split(':')[1].replace(/[^\d]+/, '')); // NaNの場合は''で置換すること
    this.head_col = a1note.split(':')[0].replace(/(\d+)/, '');
    this.tail_col = a1note.split(':')[1].replace(/(\d+)/, '');
    this.range = range;
    this.primary_key = primary_key;
    this.rowFactory = rowFactory;
    this.ss = SpreadsheetApp.getActive();
    this.valueRenderOption = 'FORMATTED_VALUE';
    this.getExistRecords();
  }

  /**
   * レンダリングオプションの変更
   * FORMATTED_VALUE: セルの表示される形式で値を取得します。数値や日付は文字列として返されますが、数式は評価された値として返されます。
   * UNFORMATTED_VALUE: セルのフォーマットを無視して、値を取得します。数値や日付も数値型として返されます。
   * FORMULA: セルの数式を取得します。数式が存在する場合はそのまま返されます。
   */
  setValueRenderOption(option: 'FORMATTED_VALUE' | 'UNFORMATTED_VALUE' | 'FORMULA') {
    this.valueRenderOption = option;
    this.getExistRecords();
  }

  /** 二次元のテーブルデータを見出しをキーとしたオブジェクト配列に変換する */
  toRecords(df: CellValue[][]): THash[] {
    const [head, ...values_arr] = df;
    return values_arr.map((values) =>
      head.reduce((acc, col, i) => {
        (acc as iHash)[col.toString()] = values[i]?.toString() || '';
        return acc;
      }, {} as THash)
    );
  }

  /** 見出し名を列名に変換する */
  toCol(key: THeader) {
    return numeric2Colname(this.head.indexOf(key) + 1);
  }

  /** データが存在する最終行を探す */
  lastRow() {
    const [last_record] = this.records.slice(-1);
    return last_record?.row || this.head_row;
  }

  /**
   * 条件にマッチする先頭のRecordを返す
   * @param target 検索条件(Hash or Row or RowNumber)
   */
  findRecord(target: THash | TRow | number) {
    return this.records.find((r) =>
      Table.isTRow(target) ? r.isEqual(target.hash) : (typeof target === 'number') ?
        r.row === target : r.isEqual(target)
    );
  }

  /** 指定した範囲のデータをRow形式で取得 */
  getExistRecords() {
    // テーブル範囲の２次元配列を取得
    const resp = Sheets.Spreadsheets?.Values?.batchGet(this.ss.getId(), {
      ranges: [this.range],
      valueRenderOption: this.valueRenderOption,
    });
    const df = resp?.valueRanges?.[0].values || [[]];

    // Row形式の配列を生成
    const [head] = df as THeader[][];
    const records = this.toRecords(df).map((hash, i) =>
      this.rowFactory(head, hash, this.head_row + 1 + i)
    );
    this.head = head;
    this.records = records;
    this.hashes = records.map((r) => r.hash as THash);
    return { head, records };
  }

  /** 指定したデータでテーブルを再作成する */
  resetTable(records: THash[] | TRow[]) {
    const values = records
      .map((r) => Table.isTRow(r) ? r.hash : r)
      .map((hash) => this.rowFactory(this.head, hash as THash).toValues())

    // シートをクリア
    Sheets.Spreadsheets?.Values?.batchClear(
      { ranges: [`${this.sheet}!${this.head_col}${this.head_row + 1}:${this.tail_col}${this.tail_row || ''}`] },
      this.ss.getId()
    );
    if (values.length > 0) {
      // シートを上書き
      Sheets.Spreadsheets?.Values?.append(
        { values },
        this.ss.getId(),
        `${this.sheet}!${this.head_col}${this.head_row + 1}`,
        { valueInputOption: 'USER_ENTERED' }
      );
      // プロパティ更新
      this.getExistRecords();
    }
  }

  /** 指定したデータでシートを上書きする(一致するデータが無ければ末尾に追加) */
  updateRecords(records: THash[] | TRow[]) {
    // 上書き範囲の配列を生成
    let next_row = this.lastRow() + 1;
    const data = records.map((r) => Table.isTRow(r) ? (r.hash as THash) : r)
      .reduce((acc, hash) => {
        const record = this.findRecord(hash);
        if (record) {
          acc.push({
            range: `${this.sheet}!${this.head_col}${record.row}`,
            values: [this.rowFactory(this.head, hash, record.row).toValues()],
          });
        } else {
          // 見つからない場合は末尾に追記
          acc.push({
            range: `${this.sheet}!${this.head_col}${next_row}`,
            values: [this.rowFactory(this.head, hash, next_row).toValues()],
          });
          next_row += 1;
        }
        return acc;
      }, [] as GoogleAppsScript.Sheets.Schema.ValueRange[])

    if (data.length > 0) {
      // シートを上書き
      Sheets.Spreadsheets?.Values?.batchUpdate({
        valueInputOption: 'USER_ENTERED',
        data,
      }, this.ss.getId());
      // プロパティ更新
      this.getExistRecords();
    }
  }

  /** recordsの内容をシートに反映 */
  save() {
    this.updateRecords(this.records);
  }

  /** 指定したデータをテーブルの末尾に追記する */
  appendRecords(records: THash[] | TRow[]) {
    // 追記用二次元配列を生成
    const values = records
      .map((r) => Table.isTRow(r) ? r.hash : r)
      .map((hash) => this.rowFactory(this.head, hash as THash).toValues())

    // シートの末尾に追記
    if (values.length > 0) {
      Sheets.Spreadsheets?.Values?.append(
        { values },
        this.ss.getId(),
        `${this.sheet}!${this.head_col}${this.lastRow() + 1}`,
        { valueInputOption: 'USER_ENTERED' }
      );
      // プロパティ更新
      this.getExistRecords();
    }
  }
}

function tableSample() {
  // 具体的なシートを指定してTableクラスを生成
  // - Rowオブジェクトの生成処理を上書きする場合、引数にコールバック関数をセットすること
  const tbl = new Table('main!A1:ZZ', 'ID');
  tbl.getExistRecords();
  console.log(tbl.head);
  console.log(tbl.records);
  tbl.appendRecords(tbl.records.slice(0, 2));
  tbl.appendRecords([{ ID: '1004', foo: 'appended' }]);
  tbl.updateRecords([{ ID: '1004', foo: 'updated' }]);
  tbl.resetTable([{ ID: '101', foo: 'test' }, { ID: '102', foo: 'test2' }])
}

