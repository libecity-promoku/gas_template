# GAS開発用テンプレ

GASをローカル環境(clasp + TS + ESLint)で開発したい人向けのテンプレートです

## 使い方

- [clasp](https://github.com/google/clasp)をglobalにインストール

```bash
npm install -g @google/clasp
```

- claspでログインしていない場合は`clasp login`でGoogleアカウントにログイン

- `.clasp.json`のスクリプトIDをデプロイしたいGASプロジェクトIDで置き換え

> [プロジェクト管理](https://developers.google.com/apps-script/api/samples/manage?hl=ja)

- 必要なライブラリをインストール

```bash
npm install
```

## スクリプト

```json
"scripts": {
  "build": "./node_modules/.bin/tsc",
  "lint": "./node_modules/.bin/eslint src/*.ts",
  "push": "npm run build && clasp push",
  "watch": "./node_modules/.bin/tsc-watch --onSuccess \"clasp push\"",
  "deploy": "npm run push && ./deploy.sh"
},
```

- `npm run build` : `./src`以下の`.ts`ファイルをコンパイル
- `npm run lint` : `./src`以下の`.ts`ファイルを静的解析
- `npm run push` : `./dist`以下のコードをGAS環境にアップロード
- `npm run watch` : `./src`以下のコードを監視し、変更があれば`push`
- `npm run deploy` : デプロイ済みのWebアプリがある場合、現在のコードでデプロイプロジェクトを更新


## FAQ

- `tsc`で以下のエラーが出る

```
node_modules/@types/google-apps-script/google-apps-script.base.d.ts(590,13): error TS2403: Subsequent variable declarations must have the same type.  Variable 'MimeType' must be of type '{ new (): MimeType; prototype: MimeType; }', but here has type 'MimeType'.
```

✅ MimeTypeという名前が重複している(と思われる)ので、当該行をコメントアウトしてください

