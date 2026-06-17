#!/usr/bin/env python3
# DUPLICATE NOTICE: This file is a duplicate of plugins/claude-session-tools/scripts/log_read.py.
# Kept here so c plugin commands can reference scripts/ via ${CLAUDE_PLUGIN_ROOT}.
# When modifying, update both files.

"""
log_read.py - セッション開始時の情報収集を一括実行

Usage:
  python log_read.py [vba|py|pyqt]

Output:
  - 読み込み対象ファイル一覧
"""
import argparse
import os
import re
import sys
from pathlib import Path

# Windows日本語出力対策
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')


def read_latest_entries(file_path, pattern, max_entries, exclude_completed=False):
    """ファイルから最新N件のエントリを抽出"""
    if not os.path.exists(file_path):
        return []

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    entries = []
    current_entry = []
    current_header = None

    for line in content.split('\n'):
        match = re.match(pattern, line)
        if match:
            if current_header and current_entry:
                entries.append((current_header, '\n'.join(current_entry).strip()))
            current_header = line
            current_entry = []
        elif current_header:
            current_entry.append(line)

    if current_header and current_entry:
        entries.append((current_header, '\n'.join(current_entry).strip()))

    # 完了マーク除外
    if exclude_completed:
        entries = [(h, c) for h, c in entries if not h.strip().startswith('## ✅')]

    # 最新N件（末尾がN件）
    return entries[-max_entries:] if entries else []


def find_latest_logs(log_dir, max_files=3):
    """最新のセッションログを検索（[main]付きを最大3件）"""
    if not os.path.exists(log_dir):
        return []

    md_files = sorted(Path(log_dir).glob('*.md'), key=lambda p: p.stat().st_mtime, reverse=True)
    if not md_files:
        return []

    result = []

    # [main]付きログを最大3件収集
    for f in md_files[:20]:  # 最大20件まで遡る
        if '[main]' in f.name:
            result.append(str(f))
            if len(result) >= max_files:
                break

    # [main]が見つからない場合、最新1件を返す
    if not result and md_files:
        result.append(str(md_files[0]))

    return result


def get_guide_paths(args):
    """言語ガイドのパスを取得 (plugin 内 guides/ 優先、フォールバックで SHARED)"""
    plugin_root = os.environ.get("CLAUDE_PLUGIN_ROOT")
    if plugin_root:
        guides_dir = os.path.join(plugin_root, "guides")
    else:
        # フォールバック: スクリプト位置から相対 (../guides)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        guides_dir = os.path.join(script_dir, "..", "guides")
        if not os.path.isdir(guides_dir):
            # 最終フォールバック: 旧 SHARED 配下
            guides_dir = os.path.join(os.path.expanduser('~'), '.claude', 'SHARED', 'guides')
    paths = []

    if 'vba' in args:
        p = os.path.join(guides_dir, 'vba.md')
        if os.path.exists(p):
            paths.append(p)
    if 'py' in args or 'python' in args or 'pyqt' in args:
        p = os.path.join(guides_dir, 'python.md')
        if os.path.exists(p):
            paths.append(p)
    if 'pyqt' in args:
        p = os.path.join(guides_dir, 'pyqt.md')
        if os.path.exists(p):
            paths.append(p)

    return paths


def display_width(s):
    """文字列の表示幅を計算（全角=2, 半角=1）"""
    width = 0
    for c in s:
        if ord(c) > 127:
            width += 2
        else:
            width += 1
    return width


def pad_to_width(s, width):
    """指定幅になるようにスペースでパディング"""
    current = display_width(s)
    return s + ' ' * (width - current)


def main():
    parser = argparse.ArgumentParser(description='セッション開始時の情報収集')
    parser.add_argument('guides', nargs='*', help='読み込むガイド (vba, py, pyqt)')
    args = parser.parse_args()

    cwd = os.getcwd()
    _warnings = []

    # === 読み込み対象ファイル ===
    print("\n---\n\n## 読み込み対象ファイル\n")
    files_to_read = []

    # セッションログ
    log_dir = os.path.join(cwd, 'for_claude', 'session_logs')
    logs = find_latest_logs(log_dir)
    for log in logs:
        files_to_read.append(log)

    # __memory.md
    memory_path = os.path.join(cwd, '__memory.md')
    if os.path.exists(memory_path):
        files_to_read.append(memory_path)

    # PROGRESS.md（docs/plans/）
    progress_path = os.path.join(cwd, 'docs', 'plans', 'PROGRESS.md')
    if os.path.exists(progress_path):
        files_to_read.append(progress_path)

    # 言語ガイド
    guides = get_guide_paths(args.guides)
    files_to_read.extend(guides)

    for f in files_to_read:
        print(f"- {f}")

    # === 警告セクション ===
    if _warnings:
        print("\n---\n\n## 警告\n")
        for w in _warnings:
            print(f"- {w}")


if __name__ == '__main__':
    main()
