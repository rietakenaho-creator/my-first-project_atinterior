# 無料相談 予約システム（Google Apps Script）

予約ページから申し込みが入ると、次の4つが自動で行われます。

1. Googleカレンダーに予定を登録（空き枠だけを表示するので二重予約なし）
2. スプレッドシート「無料相談 予約一覧」に記録
3. **あなたに通知メールを即送信**
4. 相談者に確認メール（Zoom URL付き）を送信

## ファイル

| ファイル | 内容 |
|---|---|
| `Code.gs` | 本体。冒頭の `CONFIG` に設定がまとまっています |
| `Index.html` | 予約ページの画面 |
| `appsscript.json` | GASの設定（タイムゾーン、公開範囲） |

## セットアップ（Claude Codeから）

このフォルダをPCに置き、Claude Codeで開いて「READMEの手順でセットアップして」と頼めば進められます。手順は次のとおりです。

1. **Apps Script APIを有効化**：ブラウザで https://script.google.com/home/usersettings を開き、「Google Apps Script API」をオンにする
2. **clasp をインストールしてログイン**
   ```bash
   npm install -g @google/clasp
   clasp login
   ```
3. **プロジェクトを作成してアップロード**
   ```bash
   clasp create --type standalone --title "無料相談予約" --rootDir .
   clasp push
   ```
   （`clasp create` が `appsscript.json` を上書きしようとしたら、このフォルダのものを残してください）
4. **Zoom URL を設定**：`Code.gs` の `ZOOM_URL`（必要なら `ZOOM_NOTE`）を書き換えて、もう一度 `clasp push`
5. **権限の許可と動作確認**：`clasp open-script` でエディタを開き、関数 `setup` を選んで「実行」。Googleの許可画面が出るので許可する。実行ログに通知先・予約一覧シートのURL・空き枠が表示されればOK
6. **公開**
   ```bash
   clasp deploy --description "v1"
   ```
   表示された Deployment ID を使った `https://script.google.com/macros/s/<ID>/exec` が予約ページのURLです。

> claspのバージョンによってコマンド名が少し違います（例：`clasp open-script` は古い版では `clasp open`）。エラーが出たら `clasp --help` で確認してください。

### 更新するとき

コードを修正したら `clasp push` → `clasp deploy -i <Deployment ID> --description "v2"` で、**同じURLのまま**更新されます（`-i` を付けないとURLが変わります）。

## よく変える設定（`Code.gs` の CONFIG）

- `TOPICS`：相談内容の選択肢（最後の「その他」を選ぶと自由記載が必須になります）
- `NOTIFY_EMAIL`：通知先（空欄ならGoogleアカウントのアドレス）
- `START_HOUR` / `END_HOUR` / `WEEKDAYS`：受付時間・曜日
- `SLOT_MINUTES` / `BUFFER_MINUTES`：1枠の長さ、前後の空き時間
- `DAYS_AHEAD` / `MIN_NOTICE_HOURS`：何日先まで、何時間後から受け付けるか

## 補足

- 祝日を自動で除外するには、Googleカレンダーに「日本の祝日」カレンダーを追加しておいてください（`setup` 実行時に未追加なら注意が出ます）
- 終日予定は「空き」として扱います。終日予定の日も埋めたい場合は `IGNORE_ALL_DAY_EVENTS: false`
- 無料のGoogleアカウントはメール送信数に1日あたりの上限があります（1件の予約で2通送信）。無料相談の受付なら通常は問題ありません
- 無料版Zoomは1回40分までです（30分枠なら収まります）
- キャンセル・日程変更は、確認メールへの返信で受ける運用にしています
