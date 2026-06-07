#!/usr/bin/env python3
"""
Common utilities for Claude Code hooks.

Usage:
    from hook_utils import read_stdin_json, get_personal_plugin_data

    input_data = read_stdin_json()
    if input_data is None:
        sys.exit(0)

    log_path = get_personal_plugin_data("logs") / "bash_commands.txt"
"""

import json
import os
import sys
from pathlib import Path


def get_personal_plugin_data(subdir: str = "") -> Path:
    """Return `${CLAUDE_PERSONAL_DATA}` (+ optional subdir) or fallback `~/.claude/personal-plugin-data/`.

    plugin uninstall でも残る個人 PC データ用のパスを返す。
    既存ディレクトリも自動作成 (mkdir parents=True, exist_ok=True)。
    """
    env = os.environ.get("CLAUDE_PERSONAL_DATA")
    if env:
        path = Path(env)
    else:
        path = Path.home() / ".claude" / "personal-plugin-data"
    if subdir:
        path = path / subdir
    path.mkdir(parents=True, exist_ok=True)
    return path


def read_stdin_json():
    """
    Read JSON from stdin with proper UTF-8 handling.

    Claude Code sends UTF-8 encoded JSON, but Python's sys.stdin.encoding
    may be cp932 on Windows, causing Japanese paths to be corrupted.

    Returns:
        dict: Parsed JSON data, or None if stdin is empty or invalid.
    """
    stdin_bytes = sys.stdin.buffer.read()
    if not stdin_bytes:
        return None

    try:
        stdin_raw = stdin_bytes.decode('utf-8')
    except UnicodeDecodeError:
        # Fallback to cp932 if UTF-8 fails
        stdin_raw = stdin_bytes.decode('cp932', errors='replace')

    try:
        return json.loads(stdin_raw)
    except json.JSONDecodeError:
        return None
