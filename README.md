# Mail Assist（Outlook アドイン）

定型文テンプレートの挿入と、宛名（「山田様」など）の自動入力ができる Outlook アドインです。
ビルドは不要で、このリポジトリをそのまま GitHub Pages で公開すれば使えます。

## できること

| 画面 | ボタン | 動作 |
|---|---|---|
| メール作成 | **テンプレート** | テンプレートを選んでカーソル位置に挿入（件名も設定可） |
| メール作成 | **宛名を挿入** | 宛先（To）の名前から「山田様」などを本文の先頭に入れる |
| メール閲覧 | **テンプレートで返信** | テンプレートを入れた返信／全員に返信を作る |

テンプレートはメールボックスに保存されるので、別の PC や Outlook on the web でも同じものが使えます（上限は約 30KB）。

## テンプレートで使える変数

| 書き方 | 入る内容（例） |
|---|---|
| `{宛名}` | 山田様／山田様、佐藤様／関係者各位 |
| `{自分}` | 自分の表示名 |
| `{件名}` | 今の件名 |
| `{今日}` `{明日}` `{明後日}` | 10月15日(木) |
| `{今日+7}` `{今日-1}` | 今日から N 日後・前 |
| `{3営業日後}` | 土日を除いて 3 日後 |
| `{来週月曜}` 〜 `{来週日曜}` | 来週のその曜日 |
| `{?資料名}` | 挿入前に入力欄を表示 |
| `{?補足:よろしく}` | 入力欄（空なら「よろしく」） |

宛名の作り方（姓のみ／フルネーム、敬称、まとめる人数）は ⚙ 設定で変えられます。

## 公開手順（初回のみ）

1. GitHub で `addin_outlook` という **Public** リポジトリを作り、このフォルダを push する
2. リポジトリの **Settings → Pages** で Source を「**GitHub Actions**」にする（以後は push するたびに `.github/workflows/pages.yml` がテストと公開を行う）
3. 数分後に `https://masaki5711-program.github.io/addin_outlook/taskpane.html` が開けることを確認する

ユーザー名やリポジトリ名が違う場合は、`manifest.xml` の
`https://masaki5711-program.github.io/addin_outlook` をすべて置き換えてください。

## Outlook への追加

1. https://aka.ms/olksideload を開く（または Outlook の「アプリ」→「アドインを取得」）
2. **個人用アドイン → カスタム アドインを追加 → ファイルから追加** で `manifest.xml` を選ぶ
3. メール作成画面のリボン（または「…」「アプリ」）に **Mail Assist** が出る

コードを更新して push すれば、Outlook 側は再読み込みだけで反映されます。
`manifest.xml` を変えたとき（ボタン追加など）は `<Version>` を上げて追加し直してください。

## 開発

```sh
npm test           # 変数展開・宛名作成のテスト
npm run validate   # マニフェストの検証（Microsoft のサービスを使用）
```

| パス | 内容 |
|---|---|
| `manifest.xml` | アドインの定義（ボタン・URL） |
| `taskpane.html` / `js/taskpane.js` | テンプレートの作業ウィンドウ |
| `commands.html` / `js/commands.js` | 「宛名を挿入」ボタンの処理 |
| `js/lib/render.js` | 変数の展開 |
| `js/lib/salutation.js` | 宛名の作成 |
| `js/lib/store.js` | 保存・書き出し・読み込み |
| `js/lib/defaults.js` | 初期テンプレート |
| `tools/make-icons.ps1` | アイコン PNG の生成 |
