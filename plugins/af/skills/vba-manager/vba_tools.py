# -*- coding: utf-8 -*-
"""
VBA Manager - 統合スクリプト
VBAモジュールのインポート/エクスポート/リンターを1ファイルに統合

使用方法:
  モジュール一覧:     python vba_tools.py list [--prod|--test]
  エクスポート:       python vba_tools.py export [--prod|--test]
  インポート:         python vba_tools.py import <module1> [module2 ...] [--prod|--test]
  マクロ実行:         python vba_tools.py run <macro_name> [--prod|--test]
  リンター:           python vba_tools.py lint [vba_file_path]
  プロシージャ一覧:   python vba_tools.py inspect [module_name] [--prod|--test] [--vars]
  パラメータ追加:     python vba_tools.py add-param <func_name> <new_param> <vba_file>

インポート例:
  python vba_tools.py import A01_Main --prod          # 単一モジュール
  python vba_tools.py import A01_Main A02_Data --prod # 複数モジュール（1回のExcelセッションで処理）

マクロ実行例:
  python vba_tools.py run DEBUG_CheckCSV --prod       # デバッグ用マクロ実行
  python vba_tools.py run A00_Main.MAIN_CheckCSV --test  # モジュール名付きで指定

ワークブック指定:
  デフォルト  → .testファイルで指定されたテスト用ワークブック（安全）
  --prod      → .mainファイルで指定された本番用ワークブック

ワークブック指定ファイル:
  .test  テスト用（デフォルト・安全）
  .main  本番用（--prod必須）

add-param例:
  python vba_tools.py add-param DebugLog "A02_Data.G_IsDebugMode" vba_modules/A01_Main.bas
  → 全てのDebugLog(...)呼び出しに第2引数を追加

2026-06-07: プラグイン化
"""
import re
import sys
import tempfile
import subprocess
from pathlib import Path
from collections import defaultdict
from typing import List, Optional, Set, Tuple

# ============================================================================
# パス検出
# ============================================================================

def get_project_root() -> Path:
    """プロジェクトルートを検出（カレントディレクトリを使用）"""
    return Path.cwd()


def find_workbook(hint: str = None, prod_mode: bool = False) -> Optional[Path]:
    """
    ワークブックを検出

    Args:
        hint: ヒント文字列（ファイル名の一部、パス等）
        prod_mode: Trueの場合、.mainファイル（本番）を使用。Falseは.test（デフォルト・安全）

    Returns:
        見つかったワークブックのパス、見つからない場合はNone

    Note:
        .main と .test ファイルは常に両方存在する前提。
        フォールバック検索は行わない（誤インポート防止）。
    """
    root = get_project_root()

    # ヒントがフルパスの場合
    if hint:
        hint_path = Path(hint)
        if hint_path.exists() and hint_path.suffix in ('.xlsm', '.xlsx', '.xlsb'):
            return hint_path

    # --prod モードの場合: .mainファイルを検索（本番）
    if prod_mode:
        main_files = [f for f in root.glob("**/*.main") if f.with_suffix('.xlsm').exists()]
        if len(main_files) > 1:
            print("[ERROR] .mainファイルが複数見つかりました。1つにしてください:")
            for f in main_files:
                print(f"  - {f.relative_to(root)}")
            return None
        if main_files:
            xlsm_path = main_files[0].with_suffix('.xlsm')
            print(f"[本番] .main指定 -> {xlsm_path.name}")
            return xlsm_path
        print("[お知らせ] .mainファイルが見つかりません")
        print("  本番用ワークブックと同名の .main ファイルを作成してください")
        print("  例: macro/マクロ実行ブック.main")
        return None

    # デフォルト: .testファイルを検索（テスト用・安全）
    test_files = [f for f in root.glob("**/*.test") if f.with_suffix('.xlsm').exists()]
    if len(test_files) > 1:
        print("[ERROR] .testファイルが複数見つかりました。1つにしてください:")
        for f in test_files:
            print(f"  - {f.relative_to(root)}")
        return None
    if test_files:
        xlsm_path = test_files[0].with_suffix('.xlsm')
        print(f"[テスト] .test指定 -> {xlsm_path.name}")
        return xlsm_path

    print("[お知らせ] .testファイルが見つかりません")
    print("  テスト用ワークブックと同名の .test ファイルを作成してください")
    print("  例: macro/マクロ実行ブック_テスト用.test")
    print("  本番用は --prod を指定")
    return None


def get_vba_modules_dir(workbook_path: Path, prod_mode: bool = False) -> Path:
    """ワークブックに対応するVBAモジュールディレクトリを取得

    Args:
        workbook_path: ワークブックのパス
        prod_mode: Trueで本番用(vba_modules)、Falseでテスト用(vba_modules_test)
    """
    root = get_project_root()
    wb_parent = workbook_path.parent.name

    # パターンマッチング
    if wb_parent == "macro_sub":
        return root / "macro_sub" / "dev_macro"
    elif wb_parent == "macro_test":
        return root / "macro_test" / "dev_macro"
    elif wb_parent == "macro":
        if prod_mode:
            return root / "macro" / "vba_main_modules"
        else:
            return root / "macro" / "vba_test_modules"
    else:
        # ワークブックと同じディレクトリにvba_main_modulesを作成
        return workbook_path.parent / "vba_main_modules"


def get_vba_modules_origin_dir(workbook_path: Path, prod_mode: bool = False) -> Path:
    """ワークブックに対応するVBAモジュールオリジナルディレクトリを取得（CP932エクスポート用）

    Args:
        workbook_path: ワークブックのパス
        prod_mode: Trueで本番用(vba_main_modules_origin)、Falseでテスト用(vba_test_modules_origin)
    """
    root = get_project_root()
    wb_parent = workbook_path.parent.name

    if wb_parent == "macro_sub":
        return root / "macro_sub" / "dev_macro_origin"
    elif wb_parent == "macro_test":
        return root / "macro_test" / "dev_macro_origin"
    elif wb_parent == "macro":
        if prod_mode:
            return root / "macro" / "vba_main_modules_origin"
        else:
            return root / "macro" / "vba_test_modules_origin"
    else:
        return workbook_path.parent / "vba_main_modules_origin"


# ============================================================================
# Excel環境チェック
# ============================================================================

def is_excel_running() -> bool:
    """Excelプロセスが実行中かチェック"""
    try:
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/NH'],
            capture_output=True,
            encoding='cp932',
            errors='ignore',
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
        )
        return 'EXCEL.EXE' in result.stdout.upper()
    except Exception:
        return False


def is_excel_file_open(file_path: Path) -> bool:
    """Excelファイルが開かれているかチェック"""
    temp_file = file_path.parent / f"~${file_path.name}"
    if temp_file.exists():
        return True

    try:
        with open(file_path, 'r+b') as f:
            pass
        return False
    except (PermissionError, IOError):
        return True


def check_excel_environment(file_path: Path) -> Tuple[bool, str]:
    """Excel環境をチェック"""
    messages = []

    if is_excel_running():
        messages.append("Excelプロセスが実行中です")

    if is_excel_file_open(file_path):
        messages.append(f"ファイル '{file_path.name}' が開かれています")

    if messages:
        return False, "\n".join(messages)
    return True, ""


# ============================================================================
# VBAファイル読み書き
# ============================================================================

def read_vba_file(vba_file_path: Path, encoding: str = 'utf-8') -> str:
    """VBAファイルを読み込む"""
    with open(vba_file_path, 'r', encoding=encoding, errors='replace') as f:
        return f.read()


def write_vba_file(vba_file_path: Path, content: str, encoding: str = 'utf-8') -> None:
    """VBAファイルに書き込む"""
    with open(vba_file_path, 'w', encoding=encoding) as f:
        f.write(content)


# ============================================================================
# VBAリンター
# ============================================================================

class VBALinter:
    """VBA構文チェッカー"""

    def __init__(self, vba_code: str, file_path: Path = None,
                 global_public_vars: Set[str] = None,
                 global_public_consts: Set[str] = None):
        self.lines = vba_code.split('\n')
        self.issues = []
        self.file_path = file_path
        self.is_class_module = file_path and str(file_path).endswith('.cls') if file_path else False
        self.module_level_vars = set()
        self._collect_module_level_vars()
        self.global_public_vars = global_public_vars or set()
        self.module_level_consts = set()
        self._collect_module_level_consts()
        self.global_public_consts = global_public_consts or set()

    def _collect_module_level_vars(self):
        """モジュールレベルの変数を収集"""
        in_procedure = False
        for line in self.lines:
            if line.strip().startswith("'"):
                continue
            if re.match(r'^\s*(Private|Public)?\s*(Sub|Function)\s+(\w+)', line, re.IGNORECASE):
                in_procedure = True
                continue
            if re.match(r'^\s*End\s+(Sub|Function)', line, re.IGNORECASE):
                in_procedure = False
                continue
            if not in_procedure:
                var_decl_match = re.match(r'^\s*(Public|Private|Dim)\s+(.+)', line, re.IGNORECASE)
                if var_decl_match:
                    vars_list = self._extract_variables_from_declaration(var_decl_match.group(2))
                    for var_name in vars_list:
                        self.module_level_vars.add(var_name.lower())

    def _collect_module_level_consts(self):
        """モジュールレベルの定数を収集"""
        in_procedure = False
        for line in self.lines:
            if line.strip().startswith("'"):
                continue
            if re.match(r'^\s*(Private|Public)?\s*(Sub|Function)\s+(\w+)', line, re.IGNORECASE):
                in_procedure = True
                continue
            if re.match(r'^\s*End\s+(Sub|Function)', line, re.IGNORECASE):
                in_procedure = False
                continue
            if not in_procedure:
                const_match = re.match(r'^\s*(Public|Private)?\s*Const\s+(\w+)', line, re.IGNORECASE)
                if const_match:
                    self.module_level_consts.add(const_match.group(2).lower())

    def _extract_variables_from_declaration(self, declaration_part: str) -> List[str]:
        """変数宣言部分から変数名を抽出"""
        variables = []
        parts = declaration_part.split(',')
        for part in parts:
            part = part.strip()
            match = re.match(r'(\w+)\s*(\(\s*\))?\s+As\s+', part, re.IGNORECASE)
            if match:
                variables.append(match.group(1))
        return variables

    # 継続行の最大数（VBA制限）
    MAX_CONTINUATION_LINES = 24

    def check_all(self) -> List[dict]:
        """全チェック実行"""
        self.check_duplicate_declarations()
        self.check_duplicate_procedures()
        self.check_duplicate_constants()
        self.check_backslash_in_strings()
        self.check_continuation_lines()
        self.check_formula_strings()
        return self.issues

    def check_formula_strings(self):
        """VBAコード内の .Formula/.Formula2 代入文字列を構文チェック"""
        try:
            from formulas import Parser
        except ImportError:
            return  # formulas未インストール時はスキップ

        parser = Parser()
        # 継続行を結合して論理行を構築
        logical_lines = []
        current_line = ""
        current_start = 0
        for i, line in enumerate(self.lines, start=1):
            stripped = line.rstrip()
            if current_line == "":
                current_start = i
            has_continuation = stripped.endswith(' _') or stripped.endswith('\t_')
            if has_continuation:
                current_line += stripped[:-1].strip() + " "
            else:
                current_line += stripped.strip()
                logical_lines.append((current_start, current_line))
                current_line = ""

        # .Formula / .Formula2 代入パターンを検索
        formula_pattern = re.compile(
            r'\.(Formula2?)\s*=\s*(.+)$',
            re.IGNORECASE
        )
        for line_num, logical_line in logical_lines:
            if logical_line.strip().startswith("'"):
                continue
            match = formula_pattern.search(logical_line)
            if not match:
                continue
            rhs = match.group(2).strip()
            # 行末コメント除去（文字列外の ' 以降を削除）
            rhs = self._strip_vba_comment(rhs)
            # VBA文字列連結を解決: "..." & var & "..." → 連結結果
            # 全ての文字列リテラルを抽出して連結、変数部分はダミーセル参照A1に
            excel_formula = self._resolve_vba_formula_concat(rhs)
            if excel_formula is None:
                continue
            # 数式として不完全な場合スキップ
            if not excel_formula.startswith('='):
                continue
            try:
                parser.ast(excel_formula)
            except Exception as e:
                self.issues.append({
                    'type': '数式構文エラー',
                    'line': line_num,
                    'message': f"Formula文字列の構文エラー: {type(e).__name__}",
                    'code': logical_line.strip()[:80] + ('...' if len(logical_line.strip()) > 80 else '')
                })

    @staticmethod
    def _strip_vba_comment(rhs: str) -> str:
        """文字列リテラル外の ' コメントを除去"""
        in_string = False
        i = 0
        while i < len(rhs):
            if rhs[i] == '"':
                if in_string:
                    if i + 1 < len(rhs) and rhs[i + 1] == '"':
                        i += 2
                        continue
                    in_string = False
                else:
                    in_string = True
            elif rhs[i] == "'" and not in_string:
                return rhs[:i].rstrip()
            i += 1
        return rhs

    @staticmethod
    def _resolve_vba_formula_concat(rhs: str) -> Optional[str]:
        """VBA右辺の文字列連結式を解決してExcel数式文字列を生成する
        例: "=IF(" & addr & "="""",""""," & addr & ")" → =IF(A1="","",A1)
        """
        # トークナイザー: 文字列リテラルと変数を順に抽出
        # VBA文字列リテラル: "..." （内部の "" はエスケープされた "）
        result = ""
        i = 0
        rhs = rhs.strip()
        while i < len(rhs):
            # スペース・& をスキップ
            if rhs[i] in (' ', '\t'):
                i += 1
                continue
            if rhs[i] == '&':
                i += 1
                continue
            # 文字列リテラル
            if rhs[i] == '"':
                j = i + 1
                while j < len(rhs):
                    if rhs[j] == '"':
                        if j + 1 < len(rhs) and rhs[j + 1] == '"':
                            j += 2  # "" エスケープをスキップ
                        else:
                            break  # 文字列終端
                    else:
                        j += 1
                content = rhs[i + 1:j]
                result += content.replace('""', '"')
                i = j + 1
            else:
                # 変数名/式 → 次の & またはスペースまで読む
                j = i
                while j < len(rhs) and rhs[j] not in ('&', ' ', '\t'):
                    j += 1
                # シート名コンテキスト: 直前が ' なら Sheet1 に置換
                if result.endswith("'"):
                    result += "Sheet1"
                else:
                    result += "A1"
                i = j
        return result if result else None

    def check_continuation_lines(self):
        """継続記号数をチェック（VBA制限: 24個まで）"""
        continuation_count = 0  # 継続記号（_）の数
        start_line = 0

        for i, line in enumerate(self.lines, start=1):
            stripped = line.rstrip()
            has_continuation = stripped.endswith(' _') or stripped.endswith('\t_')

            if continuation_count == 0:
                # 新しい論理行の開始
                if has_continuation:
                    start_line = i
                    continuation_count = 1
            else:
                # 継続中
                if has_continuation:
                    continuation_count += 1
                else:
                    # 論理行の終了 - チェック
                    if continuation_count > self.MAX_CONTINUATION_LINES:
                        self.issues.append({
                            'type': '継続行超過',
                            'line': start_line,
                            'message': f"継続記号が{continuation_count}個（制限: {self.MAX_CONTINUATION_LINES}）- インポート時にエラーになります",
                            'code': self.lines[start_line - 1].strip()[:60] + '...'
                        })
                    continuation_count = 0

    def check_duplicate_declarations(self):
        """重複変数宣言をチェック"""
        scope_stack = [{'name': 'Module', 'start': 0, 'vars': {}}]
        in_conditional_compile = False
        for i, line in enumerate(self.lines, start=1):
            stripped = line.strip()
            if stripped.startswith("'"):
                continue
            # コンパイラ分岐（#If/#ElseIf/#Else/#End If）内の重複は無視
            if re.match(r'^#If\s+', stripped, re.IGNORECASE):
                in_conditional_compile = True
                continue
            if re.match(r'^#End\s+If', stripped, re.IGNORECASE):
                in_conditional_compile = False
                continue
            if re.match(r'^#(Else|ElseIf)\b', stripped, re.IGNORECASE):
                continue
            if re.match(r'^\s*(Private|Public)?\s*(Sub|Function)\s+(\w+)', line, re.IGNORECASE):
                match = re.match(r'^\s*(Private|Public)?\s*(Sub|Function)\s+(\w+)', line, re.IGNORECASE)
                scope_stack.append({'name': match.group(3), 'start': i, 'vars': {}})
            elif re.match(r'^\s*End\s+(Sub|Function)', line, re.IGNORECASE):
                if len(scope_stack) > 1:
                    scope_stack.pop()
            var_decl_match = re.match(r'^\s*(Public|Private|Dim|Static)\s+(.+)', line, re.IGNORECASE)
            if var_decl_match and not in_conditional_compile:
                var_names = self._extract_variables_from_declaration(var_decl_match.group(2))
                current_scope = scope_stack[-1]
                for var_name in var_names:
                    var_name_lower = var_name.lower()
                    if var_name_lower in current_scope['vars']:
                        self.issues.append({
                            'type': '重複変数宣言',
                            'line': i,
                            'message': f"変数 '{var_name}' が重複宣言（最初: 行{current_scope['vars'][var_name_lower]}）",
                            'code': line.strip()
                        })
                    else:
                        current_scope['vars'][var_name_lower] = i

    def check_duplicate_procedures(self):
        """重複Sub/Function定義をチェック"""
        procedures = {}
        for i, line in enumerate(self.lines, start=1):
            match = re.match(r'^\s*(Private|Public)?\s*(Sub|Function)\s+(\w+)', line, re.IGNORECASE)
            if match:
                proc_name = match.group(3).lower()
                if proc_name in procedures:
                    self.issues.append({
                        'type': '重複プロシージャ',
                        'line': i,
                        'message': f"{match.group(2)} '{match.group(3)}' が重複（最初: 行{procedures[proc_name]}）",
                        'code': line.strip()
                    })
                else:
                    procedures[proc_name] = i

    def check_duplicate_constants(self):
        """重複定数定義をチェック"""
        constants = {}
        for i, line in enumerate(self.lines, start=1):
            if line.strip().startswith("'"):
                continue
            match = re.match(r'^\s*(Private|Public)?\s*Const\s+(\w+)', line, re.IGNORECASE)
            if match:
                const_name = match.group(2).lower()
                if const_name in constants:
                    self.issues.append({
                        'type': '重複定数定義',
                        'line': i,
                        'message': f"定数 '{match.group(2)}' が重複（最初: 行{constants[const_name]}）",
                        'code': line.strip()
                    })
                else:
                    constants[const_name] = i

    def check_backslash_in_strings(self):
        """バックスラッシュエスケープをチェック"""
        for i, line in enumerate(self.lines, start=1):
            if line.strip().startswith("'"):
                continue
            if re.search(r'=\s*\\"', line):
                if not re.search(r'&\s*"[^"]*\\', line):
                    self.issues.append({
                        'type': 'バックスラッシュエラー',
                        'line': i,
                        'message': r'\" は無効。VBAでは "" を使用',
                        'code': line.strip()
                    })

    def report(self) -> bool:
        """レポート出力。問題なければTrue"""
        if not self.issues:
            print("[OK] 問題なし")
            return True

        by_type = defaultdict(list)
        for issue in self.issues:
            by_type[issue['type']].append(issue)

        print(f"[WARNING] {len(self.issues)}個の問題")
        for issue_type, items in sorted(by_type.items()):
            print(f"\n【{issue_type}】")
            for issue in items:
                print(f"  行{issue['line']}: {issue['message']}")
        return False


def lint_vba_file(vba_file_path: Path,
                  global_public_vars: Set[str] = None,
                  global_public_consts: Set[str] = None) -> bool:
    """VBAファイルをリントチェック"""
    print(f"[リンター] {vba_file_path.name}")

    # UTF-8で読み込み
    try:
        code = read_vba_file(vba_file_path, 'utf-8')
    except UnicodeDecodeError:
        # CP932で再試行
        code = read_vba_file(vba_file_path, 'cp932')

    linter = VBALinter(code, vba_file_path, global_public_vars, global_public_consts)
    linter.check_all()
    return linter.report()


def collect_public_symbols(vba_dir: Path) -> Tuple[Set[str], Set[str]]:
    """ディレクトリ内の全VBAファイルからPublic変数・定数を収集"""
    public_vars = set()
    public_consts = set()

    if not vba_dir.exists():
        return public_vars, public_consts

    for bas_file in vba_dir.glob("*.bas"):
        try:
            code = read_vba_file(bas_file)
            for line in code.split('\n'):
                # Public変数
                var_match = re.match(r'^\s*Public\s+(\w+)\s+As', line, re.IGNORECASE)
                if var_match:
                    public_vars.add(var_match.group(1).lower())
                # Public定数
                const_match = re.match(r'^\s*Public\s+Const\s+(\w+)', line, re.IGNORECASE)
                if const_match:
                    public_consts.add(const_match.group(1).lower())
        except:
            pass

    return public_vars, public_consts


# ============================================================================
# プロシージャ一覧（inspect）
# ============================================================================

# Sub/Function宣言パターン
_RE_PROC = re.compile(
    r'^\s*(Private|Public|Friend)?\s*(Sub|Function)\s+(\w+)\s*(\(.*)?',
    re.IGNORECASE,
)

# Property Get/Let/Set宣言パターン
_RE_PROP = re.compile(
    r'^\s*(Private|Public|Friend)?\s*Property\s+(Get|Let|Set)\s+(\w+)\s*(\(.*)?',
    re.IGNORECASE,
)

# Public変数/定数パターン（--vars用）
_RE_PUB_VAR = re.compile(
    r'^\s*Public\s+(Const\s+)?(\w+)\s+As\s+',
    re.IGNORECASE,
)


def _parse_vba_procedures(code: str, include_vars: bool = False) -> list:
    """VBAコードからプロシージャ一覧をパースする

    Returns:
        list of dict: [{scope, kind, name, line}, ...]
    """
    results = []
    for i, line in enumerate(code.split('\n'), start=1):
        # コメント行をスキップ
        stripped = line.strip()
        if stripped.startswith("'"):
            continue

        # Sub/Function
        m = _RE_PROC.match(line)
        if m:
            scope = (m.group(1) or 'Public').capitalize()
            kind = m.group(2).capitalize()  # Sub / Function
            name = m.group(3)
            results.append({'scope': scope, 'kind': kind, 'name': name, 'line': i})
            continue

        # Property Get/Let/Set
        m = _RE_PROP.match(line)
        if m:
            scope = (m.group(1) or 'Public').capitalize()
            kind = f"Property {m.group(2).capitalize()}"
            name = m.group(3)
            results.append({'scope': scope, 'kind': kind, 'name': name, 'line': i})
            continue

        # Public変数/定数（--vars時のみ）
        if include_vars:
            m = _RE_PUB_VAR.match(line)
            if m:
                is_const = bool(m.group(1))
                name = m.group(2)
                kind = 'Const' if is_const else 'Variable'
                results.append({'scope': 'Public', 'kind': kind, 'name': name, 'line': i})

    return results


def inspect_vba_modules(vba_modules_dir: Path,
                        module_filter: list = None,
                        include_vars: bool = False) -> bool:
    """VBAモジュール内のプロシージャ一覧を表示（Excel不要）

    Args:
        vba_modules_dir: VBAモジュールディレクトリ
        module_filter: 表示するモジュール名リスト（Noneで全モジュール）
        include_vars: Public変数/定数も表示するか

    Returns:
        成功した場合True
    """
    if not vba_modules_dir.exists():
        print(f"[ERROR] ディレクトリが見つかりません: {vba_modules_dir}")
        return False

    # 対象ファイル収集
    vba_files = []
    for ext in ['*.bas', '*.cls', '*.frm']:
        vba_files.extend(vba_modules_dir.glob(ext))
    vba_files.sort(key=lambda f: f.name)

    # モジュール名フィルタ
    if module_filter:
        filter_lower = [m.lower() for m in module_filter]
        vba_files = [f for f in vba_files if f.stem.lower() in filter_lower]
        if not vba_files:
            print(f"[ERROR] 指定されたモジュールが見つかりません: {', '.join(module_filter)}")
            return False

    total_modules = 0
    total_procs = 0

    for vba_file in vba_files:
        try:
            code = read_vba_file(vba_file, 'utf-8')
        except UnicodeDecodeError:
            code = read_vba_file(vba_file, 'cp932')

        line_count = len(code.split('\n'))
        procs = _parse_vba_procedures(code, include_vars)

        if not procs:
            continue

        total_modules += 1
        total_procs += len(procs)

        print(f"## {vba_file.name} ({line_count}行)")
        print()
        print("| スコープ | 種別 | プロシージャ名 | 行 |")
        print("|----------|------|---------------|-----|")
        for p in procs:
            print(f"| {p['scope']} | {p['kind']} | {p['name']} | {p['line']} |")
        print()

    print("---")
    print(f"合計: {total_modules}モジュール, {total_procs}個のプロシージャ")
    return True


# ============================================================================
# モジュール整合性チェック
# ============================================================================

def check_module_consistency(vb_project, vba_modules_dir: Path) -> Tuple[bool, List[str], List[str]]:
    """
    ワークブック内モジュールとvba_modules/フォルダの整合性をチェック

    Args:
        vb_project: VBProject オブジェクト
        vba_modules_dir: VBAモジュールディレクトリ

    Returns:
        (整合性OK, ワークブックにのみ存在するモジュール, フォルダにのみ存在するファイル)
    """
    # ワークブック内のモジュール一覧（標準/クラス/フォームのみ）
    wb_modules = set()
    for component in vb_project.VBComponents:
        if component.Type in [1, 2, 3]:  # StdModule, ClassModule, MSForm
            wb_modules.add(component.Name.lower())

    # vba_modules/ 内のファイル一覧
    folder_modules = set()
    for ext in ['*.bas', '*.cls', '*.frm']:
        for f in vba_modules_dir.glob(ext):
            folder_modules.add(f.stem.lower())

    # 差分チェック
    only_in_wb = [m for m in wb_modules if m not in folder_modules]
    only_in_folder = [m for m in folder_modules if m not in wb_modules]

    # 元のケースで表示するために再取得
    only_in_wb_display = []
    for component in vb_project.VBComponents:
        if component.Name.lower() in [m.lower() for m in only_in_wb]:
            only_in_wb_display.append(component.Name)

    only_in_folder_display = []
    for ext in ['*.bas', '*.cls', '*.frm']:
        for f in vba_modules_dir.glob(ext):
            if f.stem.lower() in [m.lower() for m in only_in_folder]:
                only_in_folder_display.append(f.stem)

    is_consistent = len(only_in_wb) == 0 and len(only_in_folder) == 0
    return is_consistent, only_in_wb_display, only_in_folder_display


def report_module_consistency(vb_project, vba_modules_dir: Path) -> bool:
    """
    モジュール整合性チェック結果を表示

    Returns:
        整合性OKならTrue
    """
    is_ok, only_in_wb, only_in_folder = check_module_consistency(vb_project, vba_modules_dir)

    print()
    print("=" * 60)
    print("モジュール整合性チェック")
    print("=" * 60)

    if is_ok:
        print("[OK] ワークブックとフォルダの内容が一致しています")
        return True

    if only_in_wb:
        print()
        print("[警告] ワークブックにのみ存在（フォルダにない）:")
        for m in sorted(only_in_wb):
            print(f"  - {m}")
        print("  → export で取り出すか、不要なら削除を検討")

    if only_in_folder:
        print()
        print("[警告] フォルダにのみ存在（ワークブックにない）:")
        for m in sorted(only_in_folder):
            print(f"  - {m}")
        print("  → import で追加するか、不要なら削除を検討")

    return False


# ============================================================================
# インポート/エクスポート
# ============================================================================

def import_vba_modules(module_names: List[str],
                       workbook_path: Path,
                       skip_linter: bool = False,
                       prod_mode: bool = False) -> bool:
    """
    複数のVBAモジュールを1回のExcelセッションでインポート

    Args:
        module_names: モジュール名のリスト
        workbook_path: ワークブックのパス
        skip_linter: リンターをスキップ
        prod_mode: Trueで本番用(vba_modules)、Falseでテスト用(vba_modules_test)

    Returns:
        全て成功した場合True
    """
    import pythoncom
    import win32com.client

    vba_modules_dir = get_vba_modules_dir(workbook_path, prod_mode)

    # 環境チェック（1回だけ）
    env_ok, env_msg = check_excel_environment(workbook_path)
    if not env_ok:
        print(f"[ERROR] {env_msg}")
        print("Excelを終了してから再実行してください")
        return False

    # 各モジュールのファイルパスを解決し、リンターチェック
    module_files = []
    public_vars, public_consts = collect_public_symbols(vba_modules_dir)

    for module_name in module_names:
        vba_file_path = vba_modules_dir / f"{module_name}.bas"
        if not vba_file_path.exists():
            vba_file_path = vba_modules_dir / f"{module_name}.cls"
        if not vba_file_path.exists():
            vba_file_path = vba_modules_dir / f"{module_name}.frm"

        if not vba_file_path.exists():
            print(f"[ERROR] ファイルが見つかりません: {module_name}")
            return False

        # リンターチェック（.frmもチェック可能）
        if not skip_linter:
            if not lint_vba_file(vba_file_path, public_vars, public_consts):
                print(f"[ERROR] リンターエラー: {module_name}。インポートを中止")
                return False

        module_files.append((module_name, vba_file_path))

    # 一時ファイルを事前に作成
    temp_files = []
    temp_dir = None  # フォーム用一時ディレクトリ
    vba_origin_dir = get_vba_modules_origin_dir(workbook_path, prod_mode)

    try:
        import shutil

        for module_name, vba_file_path in module_files:
            content = read_vba_file(vba_file_path, 'utf-8')

            # .frmの場合は一時ディレクトリに元の名前でコピー（.frxも一緒に）
            if vba_file_path.suffix == '.frm':
                if temp_dir is None:
                    temp_dir = Path(tempfile.mkdtemp())

                # .frmをCP932変換してコピー
                temp_frm = temp_dir / vba_file_path.name
                write_vba_file(temp_frm, content, 'cp932')

                # origin版の.frxをコピー
                frx_origin = vba_origin_dir / vba_file_path.with_suffix('.frx').name
                if frx_origin.exists():
                    temp_frx = temp_dir / frx_origin.name
                    shutil.copy(frx_origin, temp_frx)
                else:
                    print(f"[ERROR] origin版.frxが見つかりません: {frx_origin}")
                    print("  先にexport --prodを実行してください")
                    return False

                temp_files.append((module_name, temp_frm))
                print(f"[フォーム] {module_name} (UTF-8 + origin .frx)")
            else:
                # .bas/.clsは従来通り
                temp_fd, temp_path = tempfile.mkstemp(suffix=vba_file_path.suffix)
                temp_file = Path(temp_path)
                write_vba_file(temp_file, content, 'cp932')
                temp_files.append((module_name, temp_file))

        print("[変換] UTF-8 → CP932")

        # COMでインポート（1回のセッションで全モジュール処理）
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            wb = excel.Workbooks.Open(str(workbook_path.absolute()))
            import time
            time.sleep(3)  # COMサーバー準備待ち（RPC_E_CALL_REJECTED対策）
            vb_project = wb.VBProject

            for module_name, temp_file in temp_files:
                # 既存モジュール削除
                for component in vb_project.VBComponents:
                    if component.Name == module_name:
                        vb_project.VBComponents.Remove(component)
                        print(f"[削除] 既存の {module_name}")
                        break

                # インポート
                vb_project.VBComponents.Import(str(temp_file.absolute()))
                print(f"[インポート] {module_name}")

            wb.Save()
            print(f"[保存] {workbook_path.name}")

            # モジュール整合性チェック
            report_module_consistency(vb_project, vba_modules_dir)

            return True

        except Exception as e:
            print(f"[ERROR] {e}")
            return False
        finally:
            try:
                if 'wb' in locals():
                    wb.Close(SaveChanges=False)
            except:
                pass
            pythoncom.CoUninitialize()

    finally:
        # 一時ファイルを削除（temp_dir外のファイルのみ）
        for _, temp_file in temp_files:
            # temp_dir内のファイルはrmtreeで削除するのでスキップ
            if temp_dir and temp_file.parent == temp_dir:
                continue
            if temp_file.exists():
                try:
                    temp_file.unlink()
                except:
                    pass
        # フォーム用一時ディレクトリを削除
        if temp_dir and temp_dir.exists():
            try:
                shutil.rmtree(temp_dir)
            except:
                pass


def import_vba_module(module_name: str,
                      workbook_path: Path,
                      vba_file_path: Path = None,
                      skip_linter: bool = False,
                      prod_mode: bool = False) -> bool:
    """
    VBAモジュールをインポート（単一モジュール用、後方互換性のため残す）

    Args:
        module_name: モジュール名
        workbook_path: ワークブックのパス
        vba_file_path: VBAファイルのパス（省略時は自動検出）
        skip_linter: リンターをスキップ
        prod_mode: Trueで本番用(vba_modules)、Falseでテスト用(vba_modules_test)

    Returns:
        成功した場合True
    """
    return import_vba_modules([module_name], workbook_path, skip_linter, prod_mode)


def export_vba_modules(workbook_path: Path, prod_mode: bool = False) -> bool:
    """
    VBAモジュールをエクスポート

    Args:
        workbook_path: ワークブックのパス
        prod_mode: Trueで本番用(vba_modules)、Falseでテスト用(vba_modules_test)

    Returns:
        成功した場合True
    """
    import pythoncom
    import win32com.client

    vba_modules_dir = get_vba_modules_dir(workbook_path, prod_mode)
    vba_origin_dir = get_vba_modules_origin_dir(workbook_path, prod_mode)

    vba_modules_dir.mkdir(parents=True, exist_ok=True)
    vba_origin_dir.mkdir(parents=True, exist_ok=True)

    # 環境チェック
    env_ok, env_msg = check_excel_environment(workbook_path)
    if not env_ok:
        print(f"[ERROR] {env_msg}")
        return False

    excel = None
    wb = None

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(workbook_path.absolute()))
        import time
        time.sleep(3)  # COMサーバー準備待ち（RPC_E_CALL_REJECTED対策）
        vb_project = wb.VBProject

        print(f"[ワークブック] {workbook_path.name}")
        print()

        # 古いファイルを削除
        for pattern in ['*.bas', '*.cls', '*.frm']:
            for old_file in vba_origin_dir.glob(pattern):
                old_file.unlink()
            for old_file in vba_modules_dir.glob(pattern):
                old_file.unlink()
        print("[クリーンアップ] 完了")

        exported = []
        for component in vb_project.VBComponents:
            # Type: 1=StdModule, 2=ClassModule, 3=MSForm
            if component.Type in [1, 2, 3]:
                module_name = component.Name
                ext = {1: ".bas", 2: ".cls", 3: ".frm"}.get(component.Type, ".bas")

                # CP932でエクスポート
                origin_file = vba_origin_dir / f"{module_name}{ext}"
                component.Export(str(origin_file))

                # UTF-8に変換して保存
                utf8_file = vba_modules_dir / f"{module_name}{ext}"
                content = read_vba_file(origin_file, 'cp932')
                write_vba_file(utf8_file, content, 'utf-8')

                print(f"[エクスポート] {module_name}")
                exported.append(utf8_file)

        print()
        print(f"[完了] {len(exported)}個のモジュール")

        # リンターチェック
        print()
        print("=" * 60)
        print("リンターチェック")
        print("=" * 60)

        public_vars, public_consts = collect_public_symbols(vba_modules_dir)
        all_ok = True
        for vba_file in exported:
            if not lint_vba_file(vba_file, public_vars, public_consts):
                all_ok = False
            print()

        return all_ok

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except:
            pass
        pythoncom.CoUninitialize()


def list_vba_modules(workbook_path: Path) -> bool:
    """
    ワークブック内のVBAモジュール一覧を表示

    Args:
        workbook_path: ワークブックのパス

    Returns:
        成功した場合True
    """
    import pythoncom
    import win32com.client

    excel = None
    wb = None

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(workbook_path.absolute()))
        import time
        time.sleep(3)  # COMサーバー準備待ち（RPC_E_CALL_REJECTED対策）
        vb_project = wb.VBProject

        print(f"[ワークブック] {workbook_path.name}")
        print()

        # Type: 1=StdModule, 2=ClassModule, 3=MSForm, 100=Document
        type_names = {1: "標準モジュール", 2: "クラスモジュール", 3: "フォーム", 100: "ドキュメント"}

        modules = []
        for component in vb_project.VBComponents:
            modules.append({
                'name': component.Name,
                'type': component.Type,
                'type_name': type_names.get(component.Type, f"不明({component.Type})")
            })

        # タイプ別にソート
        modules.sort(key=lambda x: (x['type'], x['name']))

        # 表示
        current_type = None
        for mod in modules:
            if mod['type'] != current_type:
                current_type = mod['type']
                print(f"【{mod['type_name']}】")
            print(f"  {mod['name']}")

        print()
        print(f"[合計] {len(modules)}個のコンポーネント")

        return True

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except:
            pass
        pythoncom.CoUninitialize()


# ============================================================================
# 関数呼び出しにパラメータ追加
# ============================================================================

def add_param_to_function_calls(code: str, func_name: str, new_param: str) -> Tuple[str, int]:
    """
    関数呼び出しに新しいパラメータを追加

    文字列リテラル内の括弧を正しく処理し、マルチライン呼び出しにも対応

    Args:
        code: VBAコード
        func_name: 関数名
        new_param: 追加するパラメータ

    Returns:
        (修正後のコード, 修正した呼び出し数)
    """
    lines = code.split('\n')
    result_lines = []
    modified_count = 0

    # マルチライン呼び出しを追跡
    in_multiline_call = False
    multiline_buffer = []
    multiline_start_idx = 0

    i = 0
    while i < len(lines):
        line = lines[i]

        # マルチライン継続中
        if in_multiline_call:
            multiline_buffer.append(line)

            # 継続行かチェック（行末の _ を除去して判定）
            stripped = line.rstrip()
            if stripped.endswith(' _') or stripped.endswith('\t_'):
                i += 1
                continue

            # マルチライン終了 - 結合して処理
            combined = '\n'.join(multiline_buffer)
            modified, was_modified = _add_param_to_single_call(combined, func_name, new_param)

            if was_modified:
                modified_count += 1
                # 修正後を分割して追加
                result_lines.extend(modified.split('\n'))
            else:
                result_lines.extend(multiline_buffer)

            in_multiline_call = False
            multiline_buffer = []
            i += 1
            continue

        # 関数呼び出しを検出
        if func_name in line:
            stripped = line.rstrip()

            # マルチライン開始をチェック
            if stripped.endswith(' _') or stripped.endswith('\t_'):
                in_multiline_call = True
                multiline_buffer = [line]
                multiline_start_idx = i
                i += 1
                continue

            # 単一行呼び出し
            modified, was_modified = _add_param_to_single_call(line, func_name, new_param)
            if was_modified:
                modified_count += 1
            result_lines.append(modified)
        else:
            result_lines.append(line)

        i += 1

    return '\n'.join(result_lines), modified_count


def _add_param_to_single_call(code: str, func_name: str, new_param: str) -> Tuple[str, bool]:
    """
    単一の関数呼び出し（単一行またはマルチライン結合後）にパラメータを追加

    文字列リテラル内の括弧を正しく処理
    """
    # 関数呼び出しパターンを検索
    pattern = re.escape(func_name) + r'\s*\('
    match = re.search(pattern, code)

    if not match:
        return code, False

    start_pos = match.end() - 1  # '(' の位置

    # 対応する閉じ括弧を見つける（文字列リテラルを考慮）
    close_pos = _find_matching_paren(code, start_pos)

    if close_pos == -1:
        return code, False

    # パラメータを挿入
    before = code[:close_pos]
    after = code[close_pos:]

    # 既にパラメータがあるかチェック（空の括弧でない場合はカンマを追加）
    inner_content = code[start_pos + 1:close_pos].strip()
    if inner_content:
        modified = before + ', ' + new_param + after
    else:
        modified = before + new_param + after

    return modified, True


def _find_matching_paren(code: str, open_pos: int) -> int:
    """
    開き括弧に対応する閉じ括弧の位置を見つける
    文字列リテラル内の括弧は無視

    Args:
        code: コード文字列
        open_pos: 開き括弧の位置

    Returns:
        閉じ括弧の位置、見つからない場合は-1
    """
    depth = 0
    in_string = False
    i = open_pos

    while i < len(code):
        char = code[i]

        # 文字列リテラルの処理
        if char == '"':
            if in_string:
                # エスケープされた引用符 "" をチェック
                if i + 1 < len(code) and code[i + 1] == '"':
                    i += 2
                    continue
                in_string = False
            else:
                in_string = True
            i += 1
            continue

        # 文字列内なら括弧を無視
        if in_string:
            i += 1
            continue

        # 括弧のカウント
        if char == '(':
            depth += 1
        elif char == ')':
            depth -= 1
            if depth == 0:
                return i

        i += 1

    return -1


def cmd_add_param(args: List[str]) -> bool:
    """add-paramコマンドの実行"""
    if len(args) < 3:
        print("使用方法: python vba_tools.py add-param <func_name> <new_param> <vba_file>")
        print("例: python vba_tools.py add-param DebugLog \"A02_Data.G_IsDebugMode\" vba_modules/A01_Main.bas")
        return False

    func_name = args[0]
    new_param = args[1]
    vba_file = Path(args[2])

    if not vba_file.exists():
        print(f"[ERROR] ファイルが見つかりません: {vba_file}")
        return False

    print(f"[add-param] {vba_file.name}")
    print(f"  関数: {func_name}")
    print(f"  追加パラメータ: {new_param}")

    # UTF-8で読み込み
    try:
        code = read_vba_file(vba_file, 'utf-8')
    except UnicodeDecodeError:
        code = read_vba_file(vba_file, 'cp932')

    # パラメータ追加
    modified_code, count = add_param_to_function_calls(code, func_name, new_param)

    if count == 0:
        print(f"[INFO] 修正対象なし（{func_name}の呼び出しが見つかりません）")
        return True

    # 書き込み
    write_vba_file(vba_file, modified_code, 'utf-8')
    print(f"[OK] {count}箇所を修正")

    return True


# ============================================================================
# マクロ実行
# ============================================================================

def run_vba_macro(macro_name: str,
                  workbook_path: Path,
                  close_after: bool = True) -> bool:
    """
    VBAマクロを実行

    Args:
        macro_name: 実行するマクロ名（例: DEBUG_CheckCSV, MAIN_CheckCSV）
        workbook_path: ワークブックのパス
        close_after: 実行後にワークブックを閉じるか（デフォルト: True）

    Returns:
        成功した場合True
    """
    import pythoncom
    import win32com.client

    excel = None
    wb = None

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        print(f"[ワークブック] {workbook_path.name}")
        wb = excel.Workbooks.Open(str(workbook_path.absolute()))

        # マクロ実行
        print(f"[実行] {macro_name}")
        excel.Run(macro_name)
        print(f"[完了] {macro_name}")

        # 保存
        wb.Save()
        print(f"[保存] {workbook_path.name}")

        return True

    except Exception as e:
        error_msg = str(e)
        print(f"[ERROR] {error_msg}")

        # よくあるエラーのヒント
        if "マクロ" in error_msg or "見つかりません" in error_msg:
            print()
            print("ヒント: マクロ名を確認してください")
            print("  - モジュール名.マクロ名 の形式が必要な場合があります")
            print("  例: A00_Main.DEBUG_CheckCSV")
        elif "実行中" in error_msg:
            print()
            print("ヒント: 別のExcelプロセスが実行中です")
            print("  タスクマネージャーでEXCEL.EXEを終了してください")

        import traceback
        traceback.print_exc()
        return False

    finally:
        try:
            if wb and close_after:
                wb.Close(SaveChanges=False)
        except:
            pass
        pythoncom.CoUninitialize()


# ============================================================================
# CLI
# ============================================================================

def main():
    """コマンドラインインターフェース"""
    if len(sys.argv) < 2:
        print(__doc__)
        return

    # --prod / --test フラグをパース（必須）
    has_prod = '--prod' in sys.argv
    has_test = '--test' in sys.argv
    args = [a for a in sys.argv if a not in ('--prod', '--test')]

    command = args[1].lower() if len(args) > 1 else ""

    # --vars フラグをパース（inspect用）
    has_vars = '--vars' in sys.argv
    args = [a for a in args if a != '--vars']

    # list / export / import / run / inspect は --prod または --test が必須
    if command in ("list", "export", "import", "run", "inspect"):
        if not has_prod and not has_test:
            print("[ERROR] --prod または --test を指定してください")
            print()
            print("  --prod  → 本番用 (macro/vba_main_modules/)")
            print("  --test  → テスト用 (macro/vba_test_modules/)")
            print()
            print("例:")
            print(f"  python vba_tools.py {command} --prod")
            print(f"  python vba_tools.py {command} --test")
            return
        if has_prod and has_test:
            print("[ERROR] --prod と --test は同時に指定できません")
            return

    prod_mode = has_prod

    if command == "list":
        # モジュール一覧
        workbook = find_workbook(None, prod_mode=prod_mode)
        if not workbook:
            print("[ERROR] ワークブックが見つかりません")
            return
        list_vba_modules(workbook)

    elif command == "inspect":
        # プロシージャ一覧（Excel不要）
        workbook = find_workbook(None, prod_mode=prod_mode)
        if not workbook:
            print("[ERROR] ワークブックが見つかりません")
            return
        vba_dir = get_vba_modules_dir(workbook, prod_mode=prod_mode)
        # モジュール名フィルタ（--で始まるものは除外）
        module_filter = [a for a in args[2:] if not a.startswith('--')]
        inspect_vba_modules(vba_dir, module_filter or None, include_vars=has_vars)

    elif command == "export":
        # エクスポート
        hint = args[2] if len(args) > 2 else None
        workbook = find_workbook(hint, prod_mode=prod_mode)
        if not workbook:
            print("[ERROR] ワークブックが見つかりません")
            return
        export_vba_modules(workbook, prod_mode=prod_mode)

    elif command == "import":
        # インポート（複数モジュール対応）
        if len(args) < 3:
            print("使用方法: python vba_tools.py import <module_name> [module_name2 ...] [--prod]")
            return
        # モジュール名を収集（--で始まるものは除外）
        module_names = [a for a in args[2:] if not a.startswith('--')]
        if not module_names:
            print("[ERROR] モジュール名を指定してください")
            return
        workbook = find_workbook(None, prod_mode=prod_mode)
        if not workbook:
            print("[ERROR] ワークブックが見つかりません")
            return
        import_vba_modules(module_names, workbook, prod_mode=prod_mode)

    elif command == "run":
        # マクロ実行
        if len(args) < 3:
            print("使用方法: python vba_tools.py run <macro_name> [--prod|--test]")
            print()
            print("例:")
            print("  python vba_tools.py run DEBUG_CheckCSV --prod")
            print("  python vba_tools.py run A00_Main.MAIN_CheckCSV --test")
            return
        macro_name = args[2]
        workbook = find_workbook(None, prod_mode=prod_mode)
        if not workbook:
            print("[ERROR] ワークブックが見つかりません")
            return
        # 環境チェック
        env_ok, env_msg = check_excel_environment(workbook)
        if not env_ok:
            print(f"[ERROR] {env_msg}")
            print("Excelを終了してから再実行してください")
            return
        run_vba_macro(macro_name, workbook)

    elif command == "lint":
        # リンター（複数ファイル対応）
        if len(sys.argv) < 3:
            print("使用方法: python vba_tools.py lint <vba_file_path> [vba_file_path2 ...]")
            return
        for vba_path in sys.argv[2:]:
            vba_file = Path(vba_path)
            if not vba_file.exists():
                print(f"[ERROR] ファイルが見つかりません: {vba_file}")
                continue
            lint_vba_file(vba_file)

    elif command == "add-param":
        # 関数呼び出しにパラメータ追加
        cmd_add_param(sys.argv[2:])

    else:
        print(f"不明なコマンド: {command}")
        print(__doc__)


if __name__ == "__main__":
    main()
