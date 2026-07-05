# セットアップ手順

## 1. ファイルの配置
すでにNode.js + Reactのプロジェクトがある場合、`src/` フォルダに以下をコピーしてください。

- `main.jsx`
- `App.jsx`
- `App.css`
- `index.css`
- `storagePolyfill.js`
- （これまで作った）`gdm-review.jsx` ← 同じ `src/` フォルダに置く

プロジェクトのルートには以下も置いてください。

- `tailwind.config.js`
- `postcss.config.js`

## 2. 必要なパッケージのインストール
```
npm install lucide-react papaparse xlsx
npm install -D tailwindcss postcss autoprefixer
```

## 3. 起動
```
npm run dev
```
(Vite以外のツールを使っている場合は、そのビルドコマンドを実行してください)

## 重要な注意点

### window.storage について
`gdm-review.jsx` は Claude.ai のアーティファクト専用機能 `window.storage`
（クラウドの共有キーバリューストア）を使って、レッスン内容や復習記録を保存しています。

このAPIは claude.ai の外には存在しないため、`storagePolyfill.js` で
**ブラウザのlocalStorageを使った簡易的な代用品**を用意しました。
`main.jsx` で一番最初にインポートしているので、特別な作業なしでアプリは動きます。

ただし、これは **そのブラウザだけに保存される** ものです。つまり:

- 先生が入力したレッスン内容は、先生のブラウザにしか保存されません
- 生徒が別の端末で開いても、先生が作ったレッスンは見えません
- 「クラス全員で内容を共有する」という元々の狙いは、localStorageだけでは実現できません

クラス全員で本当に内容を共有したい場合は、Firebase RealtimeDatabase /
Firestore / Supabase などの無料枠があるバックエンドサービスを使い、
`storagePolyfill.js` の中身をそちらへの読み書きに置き換える必要があります。
（希望があれば、そのための実装もお手伝いできます）

### スプレッドシート連携について
アプリ内の「スプレッドシートから読み込み」機能は、Googleスプレッドシートの
「ウェブに公開 → CSV」で得られるURLを `fetch()` で直接取得しています。
これは claude.ai の外でも同様に動作するはずです（CORSがブロックされない前提）。
