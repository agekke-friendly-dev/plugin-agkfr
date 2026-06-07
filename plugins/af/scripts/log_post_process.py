#!/usr/bin/env python3
# DUPLICATE NOTICE: This file is a duplicate of plugins/claude-session-tools/scripts/log_post_process.py.
# Kept here so c plugin commands can reference scripts/ via ${CLAUDE_PLUGIN_ROOT}.
# When modifying, update both files.

"""
log.write 後処理スクリプト

処理内容:
  - requirements.txt 更新（.venv存在時のみ）

settings.json のバックアップは ConfigChange hook (config_change_backup.py) が
変更時に自動実行するため、ここでは行わない。

Usage:
    python log_post_process.py
"""
import os
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def update_requirements():
    """cwd直下に.venvがあればrequirements.txtを更新"""
    cwd = Path.cwd()
    venv_pip = cwd / ".venv" / "Scripts" / "pip.exe"

    if not venv_pip.exists():
        print("[SKIP] .venv not found")
        return

    import subprocess
    result = subprocess.run(
        [str(venv_pip), "freeze"],
        capture_output=True, text=True
    )

    if result.returncode != 0:
        print(f"[WARN] pip freeze failed: {result.stderr.strip()}")
        return

    req_path = cwd / "requirements.txt"
    req_path.write_text(result.stdout, encoding="utf-8")
    print(f"[OK] requirements.txt updated ({len(result.stdout.splitlines())} packages)")


def main():
    update_requirements()


if __name__ == "__main__":
    main()
