# アットインテリア 無料相談 予約システム（Google Apps Script）

## 予約から事前準備までの流れ

1. 相談者が予約ページで**日時を選んで予約**（お名前・メール・電話番号）
2. 予約完了と同時に、**あなたへ通知メール**。相談者には確認メール（Zoom URL・アンケートURL付き）
3. 予約完了画面からそのまま**事前アンケート**へ（「あとで回答する」も可）
4. アンケートが回答されると、**あなたへ回答内容つきの通知メール**。Googleカレンダーの予定にも回答が書き込まれる
5. 相談日の前日になっても未回答の方には、**アンケートのお願いメールを自動送信**（1回だけ）

予約と回答はすべてスプレッドシート「アットインテリア 無料相談 予約一覧」に1行ずつ記録されます。

## 事前アンケートの項目

| 質問 | 形式 |
|---|---|
| お住まいの種類 | 1つ選ぶ |
| ご家族の人数 | 1つ選ぶ |
| 今のお悩み・ご検討内容（引越し、リノベーション各種など） | 複数選択 |
| ご希望のご予算 | 1つ選ぶ |
| アットインテリアをどこで知りましたか？（Instagram、YouTube、広告など） | 複数選択 |
| アットインテリアのデザインで好きなところ（北欧的、モダンなど） | 複数選択・任意 |
| お悩みの詳細・ご希望など | 自由記載・任意 |

質問や選択肢は `Code.gs` の `CONFIG.SURVEY` を書き換えるだけで、画面・メール・スプレッドシートに反映されます。

> スプレッドシートの見出しは最初の予約時に作られます。運用開始後に質問の数や順番を変える場合は、シートの見出しも合わせて直してください。

## ファイル

| ファイル | 内容 |
|---|---|
| `Code.gs` | 本体。冒頭の `CONFIG` に設定がまとまっています |
| `Index.html` | 予約ページ・アンケートの画面 |
| `appsscript.json` | GASの設定（タイムゾーン、公開範囲） |

## セットアップ（Claude Codeから）

このフォルダをPCに置き、Claude Codeで開いて「READMEの手順でセットアップして」と頼めば進められます。

1. **Apps Script APIを有効化**：ブラウザで https://script.google.com/home/usersettings を開き、「Google Apps Script API」をオンにする
2. **clasp をインストールしてログイン**
   ```bash
   npm install -g @google/clasp
   clasp login
   ```
3. **プロジェクトを作成してアップロード**
   ```bash
   clasp create --type standalone --title "アットインテリア無料相談予約" --rootDir .
   clasp push
   ```
   （`clasp create` が `appsscript.json` を上書きしようとしたら、このフォルダのものを残してください）
4. **Zoom URL を設定**：`Code.gs` の `ZOOM_URL`（必要なら `ZOOM_NOTE`）を書き換えて `clasp push`
5. **権限の許可と動作確認**：`clasp open-script` でエディタを開き、関数 `setup` を選んで「実行」。Googleの許可画面が出るので許可する。
   これで予約一覧シートが作られ、アンケートのお願いメールの毎日チェック（9時台）も設定されます
6. **公開**
   ```bash
   clasp deploy --description "v1"
   ```
   表示された Deployment ID を使った `https://script.google.com/macros/s/<ID>/exec` が予約ページのURLです
7. **予約ページのURLを登録**：そのURLを `Code.gs` の `WEB_APP_URL` に入れて、`clasp push` → `clasp deploy -i <Deployment ID> --description "v1.1"`
   （メールに載せるアンケートURLに使います）
8. **テスト予約**：自分のメールで1件予約し、通知メール・確認メール・アンケート・カレンダーを確認。確認後、カレンダーの予定とシートの行を削除

> claspのバージョンによってコマンド名が少し違います（例：`clasp open-script` は古い版では `clasp open`）。エラーが出たら `clasp --help` で確認してください。

### 更新するとき

コードを修正したら `clasp push` → `clasp deploy -i <Deployment ID> --description "v2"` で、**同じURLのまま**更新されます（`-i` を付けないとURLが変わり、送信済みメールのアンケートURLが使えなくなります）。

## よく変える設定（`Code.gs` の CONFIG）

- `SURVEY`：事前アンケートの質問・選択肢
- `NOTIFY_EMAIL`：通知先（空欄ならGoogleアカウントのアドレス）
- `START_HOUR` / `END_HOUR` / `WEEKDAYS`：受付時間・曜日
- `SLOT_MINUTES` / `BUFFER_MINUTES`：1枠の長さ、前後の空き時間
- `DAYS_AHEAD` / `MIN_NOTICE_HOURS`：何日先まで、何時間後から受け付けるか
- `SURVEY_REMINDER`：未回答の方へのお願いメール（`false` で止める）

## 補足

- 祝日を自動で除外するには、Googleカレンダーに「日本の祝日」カレンダーを追加しておいてください
- 終日予定は「空き」として扱います。終日予定の日も埋めたい場合は `IGNORE_ALL_DAY_EVENTS: false`
- 無料のGoogleアカウントはメール送信数に1日あたりの上限があります（1件の予約で最大4通）。無料相談の受付なら通常は問題ありません
- 無料版Zoomは1回40分までです（30分枠なら収まります）
- キャンセル・日程変更は、確認メールへの返信で受ける運用にしています
