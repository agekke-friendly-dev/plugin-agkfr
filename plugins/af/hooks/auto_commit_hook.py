#!/usr/bin/env python3
"""
Auto git commit hook on user message submit
Triggers on UserPromptSubmit event

Commit message format: M: style.css | A: memo.txt | ...
- Status prefix per file (M/A/D/R/??)
- Filename only (no directory path)
- Max 3 files shown, rest as (+N)
- Priority: modified/new first, then deleted
"""
import sys
import subprocess
import os
from datetime import datetime

# Add hooks directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from hook_utils import read_stdin_json


# Map git status codes to display labels
STATUS_MAP = {
    "M": "M",   # Modified
    "A": "A",   # Added (staged)
    "??": "A",  # Untracked (new file) -> show as A
    "D": "D",   # Deleted
    "R": "R",   # Renamed
    "MM": "M",
    "AM": "M",
}


def get_changed_files(cwd):
    """Get changed files with status. Returns list of (status_label, filename).

    Priority: modified/new first, then deleted.
    """
    result = subprocess.run(
        ["git", "-c", "core.quotepath=false", "status", "--porcelain"],
        cwd=cwd, capture_output=True, text=True, check=False
    )
    if not result.stdout.strip():
        return []

    priority = []  # modified, new, renamed
    deleted = []   # deleted

    for line in result.stdout.strip().split("\n"):
        if len(line) < 4:
            continue
        status = line[:2].strip()
        filepath = line[3:].strip().strip('"')
        # Renamed files: "R  old -> new"
        if " -> " in filepath:
            filepath = filepath.split(" -> ")[-1]
        filename = os.path.basename(filepath) or os.path.basename(filepath.rstrip("/"))
        label = STATUS_MAP.get(status, status[0] if status else "?")

        if status == "D":
            deleted.append((label, filename))
        else:
            priority.append((label, filename))

    return priority + deleted


def build_commit_message(files, max_len=100):
    """Build commit message.

    Format: M: style.css | A: memo.txt | D: old.py (+2)
    Adds files until 100 char limit, rest as (+N).
    """
    if not files:
        return None

    parts = []
    for i, (label, name) in enumerate(files):
        candidate = f"{label}: {name}"
        # Check if adding this part would exceed max_len
        test_msg = " | ".join(parts + [candidate])
        remaining = len(files) - i - 1
        if remaining > 0:
            test_msg += f" (+{remaining})"
        if len(test_msg) > max_len and parts:
            # Over limit, stop here
            remaining = len(files) - i
            msg = "auto: " + " | ".join(parts) + f" (+{remaining})"
            return msg
        parts.append(candidate)

    return "auto: " + " | ".join(parts)


def main():
    # Display current datetime
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{now}]")

    # Get project directory from environment variable
    project_dir = os.environ.get("CLAUDE_PROJECT_DIR", os.getcwd())

    input_data = read_stdin_json()
    if input_data is None:
        sys.exit(0)

    cwd = input_data.get("cwd", project_dir)

    # Get changed files BEFORE git add
    files = get_changed_files(cwd)
    commit_msg = build_commit_message(files)

    if not commit_msg:
        sys.exit(0)  # No changes to commit

    # git add -A && git commit
    try:
        subprocess.run(
            ["git", "add", "-A"],
            cwd=cwd,
            capture_output=True,
            check=False
        )

        subprocess.run(
            ["git", "commit", "-m", commit_msg],
            cwd=cwd,
            capture_output=True,
            text=True,
            check=False
        )
    except Exception:
        pass

    sys.exit(0)

if __name__ == "__main__":
    main()
