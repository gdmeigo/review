// gdm-review.jsx は Claude.ai のアーティファクト専用API `window.storage`
// (クラウドの永続キーバリューストア) を使っています。
// このAPIは claude.ai の外では存在しないため、そのまま動かすと
// 保存・読み込みが常に失敗し、内容は何も保存されません。
//
// ここでは最低限動かせるように、ブラウザのlocalStorageで
// 同じインターフェースを再現しています。
//
// ⚠️ 重要な制限:
// localStorageは「このブラウザだけ」に保存されます。
// 元のアプリが想定していた「先生と生徒でクラウド上のレッスン内容を共有する」
// という機能は、これだけでは実現できません（各自のブラウザに別々に保存されます）。
// 本当にクラス全員で共有したい場合は、Firebase / Supabase などの
// バックエンドを別途用意して、この window.storage の実装を
// そちらに向けて書き換える必要があります。

if (typeof window !== "undefined" && !window.storage) {
  const STORE_KEY = "__gdm_review_kv_store__";

  const readAll = () => {
    try {
      return JSON.parse(localStorage.getItem(STORE_KEY) || "{}");
    } catch {
      return {};
    }
  };
  const writeAll = (data) => {
    localStorage.setItem(STORE_KEY, JSON.stringify(data));
  };
  const scopedKey = (key, shared) => (shared ? "shared:" : "personal:") + key;

  window.storage = {
    async get(key, shared = false) {
      const data = readAll();
      const k = scopedKey(key, shared);
      if (!(k in data)) return null;
      return { key, value: data[k], shared: !!shared };
    },
    async set(key, value, shared = false) {
      const data = readAll();
      data[scopedKey(key, shared)] = value;
      writeAll(data);
      return { key, value, shared: !!shared };
    },
    async delete(key, shared = false) {
      const data = readAll();
      const k = scopedKey(key, shared);
      const existed = k in data;
      delete data[k];
      writeAll(data);
      return existed ? { key, deleted: true, shared: !!shared } : null;
    },
    async list(prefix = "", shared = false) {
      const data = readAll();
      const scopePrefix = scopedKey(prefix, shared);
      const scopeTag = shared ? "shared:" : "personal:";
      const keys = Object.keys(data)
        .filter((k) => k.startsWith(scopePrefix))
        .map((k) => k.slice(scopeTag.length));
      return { keys, prefix, shared: !!shared };
    },
  };
}
