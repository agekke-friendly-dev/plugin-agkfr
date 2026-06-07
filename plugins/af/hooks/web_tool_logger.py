#!/usr/bin/env python3
"""PostToolUse(WebFetch|WebSearch) logger: 1 行 1 件で web_tool_usage.txt に追記 (最大 10000 行)"""
import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from hook_utils import read_stdin_json, get_personal_plugin_data

MAX_LINES = 10000


def main():
    data = read_stdin_json()
    if data is None:
        sys.exit(0)

    tool_name = data.get("tool_name", "")
    if tool_name not in ("WebFetch", "WebSearch"):
        sys.exit(0)

    tool_input = data.get("tool_input", {})
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if tool_name == "WebSearch":
        query = tool_input.get("query", "")
        line = f"{now} | WebSearch | {query}\n"
    else:
        url = tool_input.get("url", "")
        prompt = tool_input.get("prompt", "")[:50]
        line = f"{now} | WebFetch | {url} | {prompt}...\n"

    log_path = get_personal_plugin_data("logs") / "web_tool_usage.txt"

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
