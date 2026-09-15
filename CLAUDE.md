# 開発メモ

名城大学のWebClass / Google Classroom から課題を取得し、Google Tasks に登録する
Google Apps Script のプロジェクト。個人用で、配布はしていない。

## 反映方法

Apps Script へは clasp で直接反映する。**コピペはしない。**

```
clasp push -f
```

`gas/` 以下がそのまま Apps Script のファイルになる（`.clasp.json` の `rootDir`）。
`.clasp.json` はスクリプトIDを含むため gitignore 済み。認証情報は `~/.clasprc.json`。

コードを変更したら必ず `clasp push -f` まで行うこと。
リポジトリと Apps Script が食い違うと、どちらが正か分からなくなる。

Apps Script 側を確認したいときは、`gas/` を上書きしないよう別ディレクトリへ `clasp pull` して差分を見る。

## ファイル構成

`gas/` 配下のみが本番。

| ファイル | 役割 |
|---|---|
| `Main.js` | エントリポイント。メニュー、トリガー設定、`dailySystemRun` |
| `AppLogic.js` | 取得・Tasks同期・シート整理の中核 |
| `WebClassClient.js` | WebClassのSSO/SAMLログインとHTTP。Cookieとリダイレクトを手で管理 |
| `Parser.js` | WebClassのHTMLを正規表現で解析 |
| `Check.js` | 取りこぼしチェック（メニューから手動実行）。自動実行からは呼ばれない |
| `Utils.js` | ログ、設定、`Health`、`SheetUtils`、共通ヘルパー |
| `Config.js` | 定数。`COL` / `FLAG` / `TERMINAL_FLAGS` もここ |
| `appsscript.json` | マニフェスト。OAuthスコープを明示的に固定している |

## 踏みやすい落とし穴

### 実行の重複を防ぐ

`dailySystemRun` と `checkMissingTasks` は `LockService` のスクリプトロックを取る。
どちらも同じシートを読み書きするため、重なると Tasks ID を失ったり二重登録が起きる。

自動実行は待たずにスキップする（トリガーは1日複数回あるので次の機会に処理される）。
手動のチェックは30秒待ってから諦める。

### OAuthスコープは自動推論されない

`appsscript.json` に `oauthScopes` を書いているため、スコープの自動推論は無効。
新しいGoogleサービスを使うコードを足したら、スコープも手で追加する。
足し忘れると実行時に権限エラーになる。

スコープを変更したら、https://myaccount.google.com/permissions で既存の認可を
取り消してから再実行する。取り消さないと古いスコープのまま動く。

`courses.courseWork.list` に必要なのは `classroom.coursework.me.readonly`（学生向け）。
自動推論だと教師向けの `classroom.coursework.students.readonly` が選ばれ、
「The caller does not have permission」になる。

### ログの表記規則

| 記法 | 用途 |
|---|---|
| `--- 名前 (情報) ---` | 実行全体の開始・終了のみ |
| `[WebClass]` `[Classroom]` `[Tasks]` `[設定]` `[チェック]` | 段ごとの要約 |
| `  ・内容` | 明細（コース別の件数、個別の課題など） |
| `  → 内容` | 直前の処理の結果 |
| `✅` | 正常終了の確認 |
| `⚠️` | 注意（処理は続く） |
| `🚨` | 異常 |

件数が0のものは行に出さない。要約側で「10コース中4コースに項目あり」と示す。
1実行あたり20行程度に収めること。

### ログの先頭に `=` `+` `@` を使わない

スプレッドシートが数式として解釈し `#ERROR!` になる。
`log()` 側で `'` を付ける防御を入れてあるが、見出しには `---` を使う。

### Classroom の dueDate / dueTime は UTC

`fromClassroomDue()` を使うこと。ローカル時刻として組み立てると9時間ずれる。

### 列番号とフラグはベタ書きしない

`COL.LINK` / `FLAG.EXPIRED` / `TERMINAL_FLAGS` を使う。
`row[5]` や `'EXPIRED'` を直接書くと、タイポしてもエラーにならず静かに壊れる。

### `PERMISSION_DENIED` はまずアカウントを疑う

ブラウザの複数ログインが原因のことが多い。コードやスコープではない。

### WebClassへのリクエストは増やさない

大学のサーバーなので、リトライ（`retryOnTransient`）は Google API にだけ使う。
コース巡回の `Utilities.sleep(500)` も維持する。

## 設計上の判断

### 状態はスプレッドシートが持つ

課題の一覧と Tasks ID、処理済みフラグをシートに保存している。
`processTasksSync` が長い（約150行）のは、この設計の後始末が大半を占めるため。

Tasks リスト自体を状態の正とすれば300行以上減らせるが、
「利用者が手で削除したタスク」を覚える手段が別途必要になる。
現状で動いているため保留している。

### GitHub Actions への移行は検討して中止した

技術的には可能で、WebClassのSAMLログインがNodeで動くことまで検証した。
中止したのは、動機だった「気持ち悪さ」の正体が構成ではなく
「課題が取れていないのに黙って動き続けていたこと」だったため。
個人GCPプロジェクトのOAuth設定（大学ドメインから見て第三者アプリ扱いになる）の
手間に対して、得られるものが釣り合わない。

### 異常はメールではなく Tasks に出す

普段からウィジェットで見ている場所に出すのが確実なため。
`Health` がタスクを1件だけ置き、復旧すると自動で削除する。

## 未対応の既知の問題

### 締切が延長された課題が二度と登録されない

1. 締切が過ぎると行に `EXPIRED` が付く（ここまでは正しい）
2. 先生が締切を延ばすとシートの締切欄は更新されるが、`EXPIRED` はリンク基準で引き継がれる
3. 登録判定が `TERMINAL_FLAGS`（`EXPIRED` を含む）で弾く
4. `_cleanup` は締切が未来の行を消さないため、状態が固定される

同じ経路の `SKIPPED_NODATE` は `_cleanup` が毎回行を消すため自動回復する。詰むのは `EXPIRED` のみ。

修正案: 登録判定から `EXPIRED` を外し、恒久的な除外理由ではなく毎回再評価される状態として扱う。
`COMPLETED` / `DELETED` は利用者の意思なので維持する。

## 動作確認

Apps Script のコードは手元では実行できないため、ロジック単体を切り出して
Node で検証してから `clasp push` している（GASのAPIはスタブに差し替える）。
検証用のスクリプトはリポジトリに残していない。
