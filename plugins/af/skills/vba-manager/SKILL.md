---
name: vba-manager
description: VBAインポート/エクスポート/リンター。VBA import、VBA export、モジュール取込、VBAインポート、VBAエクスポート、bas、cls、lint、リンター、inspect、インスペクト、モジュール一覧、数式チェック、Formula検証、formulas、数式バリデーションなどのキーワードで起動。
allowed-tools: Read, Bash, Write
---

# VBA Manager

Excel VBAモジュールのインポート/エクスポート/リンターチェックを行うスキル。

## 機能

1. **エクスポート**: Excel VBAプロジェクト → .bas/.cls ファイル（UTF-8）
2. **インポート**: .bas/.cls ファイル → Excel VBAプロジェクト（CP932変換）
3. **リンター**: 重複宣言、構文エラー、数式バリデーション等のチェック
   - 数式チェック: `.Formula`/`.Formula2` 代入文字列を `formulas` ライブラリ(PyPI)でAST解析し括弧/クォート不整合を検出
   - `pip install formulas` 必須（未インストール時は数式チェックのみスキップ）

## ワークブックの判断

会話の文脈からメインのワークブックを判断する:

| キーワード | ワークブック |
|-----------|-------------|
| 本番、メイン、マクロ実行 | `macro/マクロ実行ブック.xlsm` |
| dev、開発 | `macro_sub/dev.xlsm` |
| test、テスト | `macro_test/test1_fx.xlsm` |
| ファイル名の一部 | 該当するファイル |

## スクリプトの場所

プロジェクト固定（`macro/` 配下）の VBA モジュール管理に使用:

```
${CLAUDE_PLUGIN_ROOT}/skills/vba-manager/vba_tools.py
```

引数体系: `--prod` / `--test` でワークブック切替。

## 使用方法

### 【重要】--prod / --test 必須

export / import では **必ず `--prod` または `--test` を指定**すること。

| フラグ | 対象 | フォルダ |
|--------|------|----------|
| `--prod` | 本番用 | `macro/vba_main_modules/` |
| `--test` | テスト用 | `macro/vba_test_modules/` |

### エクスポート

```bash
# 本番用
python ${CLAUDE_PLUGIN_ROOT}/skills/vba-manager/vba_tools.py export --prod

# テスト用
python ${CLAUDE_PLUGIN_ROOT}/skills/vba-manager/vba_tools.py export --test
```

### インポート

**【重要】複数モジュールは必ず1コマンドで一括指定すること。**
`&&` で個別に実行すると前回のExcelプロセスが残留しエラーになる。

```bash
# 複数モジュール（1回のExcelセッションで処理）← 必ずこの形式
python ${CLAUDE_PLUGIN_ROOT}/skills/vba-manager/vba_tools.py import DEF_16 A03_Check A00_Main --prod

# 単一モジュール
python ${CLAUDE_PLUGIN_ROOT}/skills/vba-manager/vba_tools.py import A01_Main --prod

# テスト用
python ${CLAUDE_PLUGIN_ROOT}/skills/vba-manager/vba_tools.py import A01_Main --test
```

### ワークブック指定ファイル

| ファイル | 用途 | フラグ |
|----------|------|--------|
| `*.main` | 本番用ワークブック | `--prod` |
| `*.test` | テスト用ワークブック | `--test` |

### プロシージャ一覧（inspect）

**Excel不要。** `.bas/.cls` ファイルを正規表現パースし、Sub/Function/Property一覧を表示。

```bash
# 全モジュール
python ${CLAUDE_PLUGIN_ROOT}/skills/vba-manager/vba_tools.py inspect --prod

# 特定モジュールのみ
python ${CLAUDE_PLUGIN_ROOT}/skills/vba-manager/vba_tools.py inspect --prod A01_Main

# Public変数/定数も表示
python ${CLAUDE_PLUGIN_ROOT}/skills/vba-manager/vba_tools.py inspect --prod --vars
```

### リンター

```bash
python ${CLAUDE_PLUGIN_ROOT}/skills/vba-manager/vba_tools.py lint macro/vba_main_modules/A01_Main.bas
```

## ディレクトリ構成

```
project/
├── macro/
│   ├── マクロ実行ブック.xlsm           # 本番（--prod）
│   ├── マクロ実行ブック.main           # 本番指定ファイル
│   ├── マクロ実行ブック_Claudeテスト用.xlsm  # テスト（デフォルト）
│   └── マクロ実行ブック_Claudeテスト用.test  # テスト指定ファイル
├── macro/vba_main_modules/             # 本番用 UTF-8（--prod時）
├── macro/vba_main_modules_origin/      # 本番用 CP932（--prod時）
├── macro/vba_test_modules/             # テスト用 UTF-8（--test時）
├── macro/vba_test_modules_origin/      # テスト用 CP932（--test時）
└── ...
```

**復元が楽**: 本番とテストでvba_modulesが分離されているため、テストで壊しても本番のvba_modulesは無傷。

## エンコーディング

- **macro/vba_main_modules/**, **macro/vba_test_modules/**: UTF-8（Claudeで直接読み書き可能）
- **macro/vba_main_modules_origin/**, **macro/vba_test_modules_origin/**: CP932（Excelエクスポートそのまま）
- **インポート時**: UTF-8 → CP932 に自動変換

## リンターチェック項目

- 重複変数宣言
- 重複Sub/Function定義
- 重複定数定義
- バックスラッシュエスケープエラー（VBAでは `""` を使用）

## 注意事項

1. **Excelを閉じてから実行**: インポート/エクスポート前にExcelを完全に終了
2. **git commit推奨**: VBA編集前にバックアップ
3. **UTF-8で編集**: macro/vba_main_modules/*.bas はUTF-8で管理
4. **Excel起動中エラー時は一時停止**: `Excelプロセスが実行中です` エラーが出たら、即座に作業を一時停止してユーザーに報告する。勝手に再試行しない
5. **インポート完了報告必須**: インポート成功後は必ず「インポートしました」とユーザーに報告すること

## Python依存関係

- `pywin32` (win32com.client)
- `pythoncom`

## 推奨モジュール構成

新規VBAプロジェクトでは以下の命名規則・構成を推奨：

| モジュール | 役割 |
|-----------|------|
| `A00_Main.bas` | ドライバー関数（MAIN_..., DEBUG_...） |
| `A01_Data.bas` | パブリック変数、定数、Enum |
| `A02_XXX.bas` | 機能モジュール（CSV処理など） |
| `A50_Log.bas` または `A99_Debug.bas` | ログ出力機能 |
| `A51_CommonUtil.bas` | 汎用ユーティリティ |

### MAIN_... / DEBUG_... パターン

```vba
' A00_Main.bas
Public Sub MAIN_CheckCSV()
    g_RunMode = MODE_PRODUCTION  ' MsgBox表示、ログシート出力
    Call RunCheckCSV
End Sub

Public Sub DEBUG_CheckCSV()
    g_RunMode = MODE_DEBUG       ' MsgBox非表示、debug_log.txt出力
    DebugLogClear
    Call RunCheckCSV
End Sub
```

### ログ出力の動作

| モード | MsgBox | ログシート | debug_log.txt |
|--------|--------|-----------|---------------|
| `MODE_PRODUCTION` | 表示 | 出力 | × |
| `MODE_DEBUG` | 非表示 | 出力 | 出力 |

### A01_Data.bas の例

```vba
' 実行モード
Public Enum RunMode
    MODE_PRODUCTION = 0
    MODE_DEBUG = 1
End Enum

Public g_RunMode As RunMode

' シート名
Public Const SHEET_LOG As String = "動作確認用"

' メッセージ出力（モードに応じて切替）
Public Sub ShowMsg(msg As String)
    Select Case g_RunMode
        Case MODE_PRODUCTION
            MsgBox msg, vbInformation
        Case MODE_DEBUG
            DebugLog "[MSG] " & msg
    End Select
End Sub
```

## トラブルシューティング

### Excelが開いているエラー

```
Excelプロセスが実行中です
```

→ タスクマネージャーでEXCEL.EXEを終了

### ファイルが見つからない

```
[ERROR] ワークブックが見つかりません
```

→ プロジェクトルートで実行しているか確認

### 文字化け

→ macro/vba_main_modules/*.bas がUTF-8で保存されているか確認
