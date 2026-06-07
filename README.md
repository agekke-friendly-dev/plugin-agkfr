# plugin-agkfr

Claude Code 用プラグイン。**セッションログ管理 + Excel VBA + 操作監査ログ** を 1 つに集約した小規模プラグイン (`af` = agekke-friendly)。

Marketplace 名: `agkfr-tools` / Plugin 名: `af` / コマンド prefix: `/af:`

---

## Features

| 種別 | 名前 | 概要 |
|---|---|---|
| Command | `/af:log.read` | セッション開始時の引継ぎ読込（前回ログ + memo + ガイド）|
| Command | `/af:log.write` | セッション終了時のログ作成（タスク・問題・引継ぎを構造化）|
| Skill | `vba-manager` | Excel VBA の import/export/lint/inspect（**Windows + Excel + pywin32 必須**）|
| Hook | `auto_commit_hook` | プロンプト送信毎の自動 `git commit`（コミット忘れ防止）|
| Hook | `bash_logger` | Bash 実行コマンドの監査ログ |
| Hook | `file_op_logger` | Read/Write/Edit ファイル操作の監査ログ |
| Hook | `web_tool_logger` | WebSearch/WebFetch の監査ログ |

> Claude Code Plugin は、コマンド・スキル・フックをまとめて配布できる公式の拡張機構。詳細は [公式 docs](https://code.claude.com/docs/en/plugins-reference) 参照。

---

## Requirements

- **Claude Code v2.x 以上**（Plugin システム対応版）
- **Python 3.x**（hook の実行に使用、`python` コマンドが PATH に通っていること）
- **Windows + Excel + `pywin32`**（`vba-manager` skill のみ。他の機能は OS 非依存）
- （任意）**`pip install formulas`** — `vba-manager` の数式 lint 機能を使う場合のみ

---

## Installation

Claude Code 内で以下を実行:

```
/plugin marketplace add agekke-friendly-dev/plugin-agkfr
/plugin install af@agkfr-tools
/reload-plugins
```

その後 **Claude Code を一度再起動**（`/exit` → 起動し直し）。
hook は session 開始時の固定読込のため、再起動しないと有効化されない。

> ⚠️ **marketplace 名と repo 名は別物**:
> - Marketplace 名: **`agkfr-tools`**（`marketplace.json` の `name` フィールドで決定）
> - Repo 名: **`plugin-agkfr`**（GitHub 上の repo 名）
> - `install` コマンドは **`af@agkfr-tools`** を使う

---

## 動作確認

1. `/help` を実行 → 一覧に `/af:log.read` と `/af:log.write` が出る
2. 任意の Bash を実行（例: `date`）→ `~/.claude/personal-plugin-data/logs/bash_commands.txt` に行が追記される

両方確認できれば install 成功。

---

## Usage

### コマンド

- **`/af:log.read [vba|py|pyqt]`** — セッション開始時に前回までの作業内容・引継ぎ事項を一気に読み込む。プロジェクト内の最新セッションログ・`__memory.md`・`PROGRESS.md` を自動収集して並列 Read。引数で言語別ガイド（`vba` / `py` / `pyqt`）も同時 Read。**前日の続きから作業を再開したい朝イチに便利**。

- **`/af:log.write [git|lite|detail|update]`** — セッション終了時に作業内容を Markdown ログ化（`for_claude/session_logs/` 配下）。引数で挙動切替:
  - `git` — 全リポジトリ状態チェック + 同期
  - `lite` — 軽量モード（ログ作成のみ、同期/push スキップ）
  - `detail` — 詳細モード（試行錯誤・判断理由まで記録）
  - `update` — 直近ログに追記
  - 引数なし — 通常モード

### Skill

- **`vba-manager`** — Excel VBA モジュールの import/export/lint/inspect を Python で自動化する skill。`.xlsm` ⇄ `.bas/.cls`（UTF-8 / CP932 自動変換）の往復、重複宣言・構文エラー検出、`.Formula`/`.Formula2` の数式バリデーションに対応。**Excel マクロを Git で版管理したい開発者向け**。
  > ⚠️ 動作要件: **Windows のみ**（`pywin32` + `pythoncom` 必須）、**Excel インストール必須**、実行前に対象 Excel を完全に閉じておくこと。数式 lint 機能を使うには `pip install formulas` が必要。

### Hooks（自動実行）

- **`auto_commit_hook`**（`UserPromptSubmit`） — プロンプト送信毎に作業中ディレクトリで自動 `git commit`。メッセージ形式: `M: style.css | A: memo.txt | D: old.py (+3)`（状態接頭辞 + ファイル名、最大 3 件 + 残数）。「コミット忘れて作業内容を見失った」を防ぐ保険として機能。
- **`bash_logger`**（`PreToolUse: Bash`） — Claude が実行した Bash コマンドを 1 行 1 件で記録。
- **`file_op_logger`**（`PreToolUse: Read|Write|Edit`） — Claude が触れたファイルパスを記録。
- **`web_tool_logger`**（`PostToolUse: WebFetch|WebSearch`） — Claude の検索クエリ・取得 URL を記録。

### ログ出力先

すべての logger hook は以下に書き込む:

```
~/.claude/personal-plugin-data/logs/
├── bash_commands.txt       Bash 実行履歴
├── file_operations.txt     ファイル操作履歴
└── web_tool_usage.txt      Web ツール使用履歴
```

- 環境変数 **`CLAUDE_PERSONAL_DATA`** で出力先を上書き可
- 各ファイル **10,000 行ローテーション**（超過分は古い行から自動削除）
- UTF-8 / タブ区切り、PC 個別保存（`personal-plugin-data/` は `.gitignore` に追加推奨）

---

## Update（更新）

新しい commit を反映するとき:

```
/plugin marketplace update agkfr-tools
/plugin update af@agkfr-tools
/reload-plugins
```

→ Claude Code 再起動で完全反映。

> `/plugin update` の出力が `(no content)` でも正常（silent success の公式仕様）。
> 実際に反映されたかは `~/.claude/plugins/installed_plugins.json` の `gitCommitSha` が最新 commit と一致するかで確認できる。

---

## Uninstall

```
/plugin uninstall af@agkfr-tools
/plugin marketplace remove agkfr-tools
```

その後 `~/.claude/settings.json` に `agkfr-tools` の記述が残っていれば手動削除。

---

## Troubleshooting

| 症状 | 対処 |
|---|---|
| install 後 hook が動かない | `/exit` → Claude Code 再起動（settings.json は session 開始時固定読込）|
| `/plugin install` が「already installed」になる | `/plugin marketplace remove agkfr-tools` → 再 add → 再 install |
| `/plugin update` が `(no content)` で反映確認できない | `installed_plugins.json` の `gitCommitSha` を確認、最新 commit と一致なら成功 |
| `vba-manager` が動かない | Windows + Excel + `pip install pywin32` を確認、対象 Excel を完全に閉じてから実行 |
| private repo にしたい | 既知バグ多数のため非推奨（[#17201](https://github.com/anthropics/claude-code/issues/17201)）|

---

## Links

- [公式: Plugin marketplaces](https://code.claude.com/docs/en/plugin-marketplaces)
- [公式: Plugins reference](https://code.claude.com/docs/en/plugins-reference)

---

## License

社内・個人利用向け。© HideNaga3
