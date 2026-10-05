// LINE リッチメニューを Messaging API で作成し、全ユーザーのデフォルトに設定する。
// 使い方: cp .env.example .env → 値を記入 → node --env-file=.env deploy.mjs
import { readFile } from "node:fs/promises";

const here = new URL(".", import.meta.url);
const token = process.env.LINE_CHANNEL_ACCESS_TOKEN;
if (!token) throw new Error("LINE_CHANNEL_ACCESS_TOKEN が未設定です（.env を確認してください）");

// richmenu.json の {{URL_xxx}} を環境変数で置き換える
const template = await readFile(new URL("richmenu.json", here), "utf8");
const body = template.replace(/\{\{(\w+)\}\}/g, (_, key) => {
  const value = process.env[key];
  if (!value) throw new Error(`${key} が未設定です（.env を確認してください）`);
  return value;
});

async function call(url, init) {
  const res = await fetch(url, {
    ...init,
    headers: { Authorization: `Bearer ${token}`, ...init.headers },
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`${init.method} ${url} -> ${res.status}: ${text}`);
  return text ? JSON.parse(text) : {};
}

// 1. メニュー定義を検証・作成
await call("https://api.line.me/v2/bot/richmenu/validate", {
  method: "POST", headers: { "Content-Type": "application/json" }, body,
});
const { richMenuId } = await call("https://api.line.me/v2/bot/richmenu", {
  method: "POST", headers: { "Content-Type": "application/json" }, body,
});
console.log("作成:", richMenuId);

// 2. 画像をアップロード
const image = await readFile(new URL("richmenu.jpg", here));
await call(`https://api-data.line.me/v2/bot/richmenu/${richMenuId}/content`, {
  method: "POST", headers: { "Content-Type": "image/jpeg" }, body: image,
});
console.log("画像アップロード完了");

// 3. デフォルトメニューに設定
await call(`https://api.line.me/v2/bot/user/all/richmenu/${richMenuId}`, { method: "POST" });
console.log("デフォルトメニューに設定しました");
