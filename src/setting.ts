/** スクリプトプロパティ */
const SCRIPT_PROP = PropertiesService.getScriptProperties();
const SCRIPT_CACHE = CacheService.getScriptCache();

/** 設定値 */
const SETTING = {

  cache_keys: ['logs'],

  /** 設定の初期化 */
  init() {
    Object.keys(SCRIPT_PROP.getProperties())
      .map((key) => SCRIPT_PROP.deleteProperty(key));
    SCRIPT_CACHE.removeAll(this.cache_keys.map((key) => key));
  },

  /** set property wrapper */
  write(key: string, val: string) {
    SCRIPT_PROP.setProperty(key, val);
  },
  /** get property wrapper */
  read(key: string) {
    return SCRIPT_PROP.getProperty(key);
  },

  get api_key() {
    return this.read('api_key') || '';
  },
  set api_key(key: string) {
    this.write('api_key', key);
  },

  // 永続的に保持しない値はCacheで保持
  get logs() {
    const logs = SCRIPT_CACHE.get('logs');
    return logs ? JSON.parse(logs) : [];
  },
  set logs(logs: string[][]) {
    SCRIPT_CACHE.put('logs', JSON.stringify(logs), 60 * 3600 * 1000);
  },
};

/** 設定画面を表示する */
function settingMenu() {
  showSideBarFromHTML('⚙️ 各種設定', 'html/settingroot');
}

/** 設定画面ボタンクリック時処理 */
function onSetting(selects: any) {
  const { api_key } = selects;
  putLog(['OnSetting', JSON.stringify(selects, null, 2)]);

  // 設定値の更新
  SETTING.api_key = api_key;

  return '設定が完了しました';
}

