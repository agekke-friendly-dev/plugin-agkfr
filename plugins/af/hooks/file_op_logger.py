#!/usr/bin/env python3
"""PreToolUse(Read|Write|Edit) logger: 1 行 1 件で file_operations.txt に追記 (最大 10000 行)"""
import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from hook_utils import read_stdin_json, get_personal_plugin_data

MAX_LINES = 10000


def main():
    data = read_stdin_json()
    if not data:
        sys.exit(0)

    tool_name = data.get("tool_name", "")
    tool_input = data.get("tool_input", {})
    file_path = tool_input.get("file_path", "")
    cwd = data.get("cwd", "")

    log_path = get_personal_plugin_data("logs") / "file_operations.txt"

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{now}]\t{tool_name}\t{cwd}\t{file_path}\n"

    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line)

    with open(log_path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    if len(lines) > MAX_LINES:
        with open(log_path, "w", encoding="utf-8") as f:
            f.writelines(lines[-MAX_LINES:])

    sys.exit(0)


if __name__ == "__main__":
    main()
