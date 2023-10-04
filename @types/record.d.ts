// ジェネリクスのインタフェース
/** テーブルの見出し定義 */
type iHeader = string;
/** 見出しをキーとしたオブジェクト定義 */
type iHash = { [key in iHeader]: string };
/** Row : テーブルアクセス用のインタフェース */
type iRow = {
  head: iHeader[],
  hash: iHash,
  row: number,
  toValues: () => (string | undefined)[],
  isEqual(other: iHash): boolean,
}

declare namespace TableDef {
  namespace Main {
    // ジェネリクスの実体.実際のテーブルに合わせて見出し名を定義する
    /** テーブルの見出し定義 */
    type Header =
      'ID' |
      'foo';
    /** 見出しをキーとしたオブジェクト定義 */
    type Hash = { [key in Header]: string };
    /** Row : テーブルアクセス用のインタフェース */
    type Row = {
      head: Header[],
      hash: Hash,
      row: number,
      toValues: () => (string | undefined)[],
      isEqual(other: Hash): boolean,
    }
  }
}

