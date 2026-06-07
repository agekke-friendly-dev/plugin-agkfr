# plugin-agkfr

Claude Code 用プラグイン。**セッションログ管理 + Excel VBA + 操作監査ログ** を 1 つに集約した小規模プラグイン (`af` = agekke-friendly)。

Marketplace 名: `agkfr-tools` / Plugin 名: `af` / コマンド prefix: `/af:`

> ⚠️ このリポジトリは **private** です。同僚配布は **org collaborator 招待 + HTTPS 認証セットアップ** が前提です。下記「Prerequisites（同僚 PC の初回セットアップ）」を必ず先に完了させてください。

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

- **Claude Code v2.1.141 以上**（2026-05-13 リリース、`CLAUDE_CODE_PLUGIN_PREFER_HTTPS` 環境変数サポート版以降）
  - 確認: `claude --version`
  - 更新: `npm i -g @anthropic-ai/claude-code`
- **Python 3.x**（hook の実行に使用、`python` コマンドが PATH に通っていること）
- **GitHub CLI (`gh`)**（private repo 認証用、Windows: `winget install --id GitHub.cli`）
- **Windows + Excel + `pywin32`**（`vba-manager` skill のみ。他の機能は OS 非依存）
- （任意）**`pip install formulas`** — `vba-manager` の数式 lint 機能を使う場合のみ

---

## Prerequisites（同僚 PC の初回セットアップ）

private repo からの install には、初回 1 回のみ以下のセットアップが必要です。

### Step 0: org collaborator 招待を受領（事前）

`agekke-friendly-dev` org への招待メールを accept してください（HideNaga3 から個別送信）。
受領状態は以下で確認可能:

```powershell
gh api /user/memberships/orgs/agekke-friendly-dev
```

`state: "active"` なら OK。

### Step 1: Claude Code を v2.1.141 以上に更新

```powershell
claude --version
# 古ければ:
npm i -g @anthropic-ai/claude-code
```

### Step 2: GitHub CLI をインストール

```powershell
winget install --id GitHub.cli
```

### Step 3: GitHub 認証

```powershell
gh auth login
```

対話プロンプト:
- `GitHub.com` を選択
- **`HTTPS`** を選択（**SSH は選ばない**、Windows での認証バグ回避）
- `Yes (Authenticate Git with your GitHub credentials)`
- `Login with a web browser` → ブラウザで 8 桁コード入力 → 完了

### Step 4: git credential helper を gh に紐付け

```powershell
gh auth setup-git
```

これで HTTPS clone 時に `gh` の認証情報が自動で使われます。

### Step 5: SSH 強制バグの保険

```powershell
git config --global url."https://github.com/".insteadOf git@github.com:
```

Claude Code が誤って SSH を使うバグへの保険（[#27771](https://github.com/anthropics/claude-code/issues/27771), [#52234](https://github.com/anthropics/claude-code/issues/52234)）。

### Step 6: 環境変数 `CLAUDE_CODE_PLUGIN_PREFER_HTTPS` を永続設定

```powershell
[Environment]::SetEnvironmentVariable("CLAUDE_CODE_PLUGIN_PREFER_HTTPS", "1", "User")
```

v2.1.141 (2026-05-13) で追加された **HTTPS 強制クローンフラグ**。設定後は PowerShell とエクスプローラーを開き直してください（環境変数の再読込）。

### （任意）Step 7: auto-update 用 PAT を設定

Claude Code 起動時の background auto-update 用に GitHub Personal Access Token（`repo` scope）を設定:

```powershell
[Environment]::SetEnvironmentVariable("GITHUB_TOKEN", "ghp_xxxxxxxxxxxx", "User")
```

PAT は [Settings → Developer settings → Tokens (classic)](https://github.com/settings/tokens) で発行。
これがなくても手動 `/plugin marketplace update` は動作します。

---

## Installation

Prerequisites 完了後、**Claude Code を再起動してから** 以下を実行:

```
/plugin marketplace add agekke-friendly-dev/plugin-agkfr
/plugin install af@agkfr-tools
/reload-plugins
```

その後 もう一度 **Claude Code を再起動**（`/exit` → 起動し直し）。
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
| `/plugin marketplace add` で `authentication failed` | `gh auth status` で認証確認、`gh auth setup-git` 再実行、`$env:CLAUDE_CODE_PLUGIN_PREFER_HTTPS` が `1` か確認 |
| Claude Code が SSH 経由で clone してエラー | Step 5 の `insteadOf` 設定を再確認: `git config --global --get url."https://github.com/".insteadOf` が `git@github.com:` を返すこと |
| `gh auth login` 後も org repo が見えない | collaborator 招待状を accept したか確認: `gh api /user/memberships/orgs/agekke-friendly-dev` で `state: "active"` か |

---

## Fallback: 認証セットアップが動かない場合

Prerequisites の Step 1〜6 を完了しても `/plugin marketplace add` が失敗する場合、**手動 clone + directory source** で回避できます。

### F1. PAT を発行

[GitHub Settings → Developer settings → Personal access tokens (classic)](https://github.com/settings/tokens) で `repo` scope の PAT を発行。

### F2. 手動 clone

```powershell
git clone https://<USER>:<PAT>@github.com/agekke-friendly-dev/plugin-agkfr.git $HOME\agkfr-mirror
```

`<USER>` は GitHub username、`<PAT>` は発行した PAT。

### F3. settings.json に directory source として登録

`~/.claude/settings.json` を編集:

```json
{
  "extraKnownMarketplaces": {
    "agkfr-tools": {
      "source": {
        "source": "directory",
        "path": "C:\\Users\\<USER>\\agkfr-mirror"
      }
    }
  }
}
```

`<USER>` は Windows のユーザー名（`$env:USERNAME` で確認）。絶対パス必須。

### F4. Claude Code 再起動 → install

```
/plugin install af@agkfr-tools
/reload-plugins
```

### F5. 更新時

```powershell
cd $HOME\agkfr-mirror
git pull
```

→ Claude Code 内で `/plugin marketplace update agkfr-tools` + `/plugin update af@agkfr-tools` + `/reload-plugins`。

> ⚠️ Fallback の制約: auto-update は無効、更新は毎回 `git pull` を手動実行が必要。

---

## Links

- [公式: Plugin marketplaces](https://code.claude.com/docs/en/plugin-marketplaces) — Private repositories セクション (L508〜)
- [公式: Plugins reference](https://code.claude.com/docs/en/plugins-reference)

---

## License

社内・個人利用向け。© HideNaga3
