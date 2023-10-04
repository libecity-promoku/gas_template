/**
 * gas_template v0.1
 */

/** スクリプト実行用メニューの追加 */
function onOpen() {
  const menu = [
    { name: '⚙️ 各種設定', functionName: 'settingMenu' },
    null,
    { name: '🗒 プロパティ情報', functionName: 'putProp' },
    { name: '🗑 全初期化', functionName: 'beginInit' }
  ];
  SpreadsheetApp.getActiveSpreadsheet().addMenu('メニュー', menu);
}

