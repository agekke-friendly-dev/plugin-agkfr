#!/usr/bin/env python3
"""PreToolUse(Bash) logger: 1 行 1 件で bash_commands.txt に追記 (最大 10000 行)"""
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

    tool_input = data.get("tool_input", {})
    command = tool_input.get("command", "")
    cwd = data.get("cwd", "")

    command = command.replace("\r", "").replace("\n", "\\n")

    log_path = get_personal_plugin_data("logs") / "bash_commands.txt"

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{now}]\t{cwd}\t{command}\n"

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
