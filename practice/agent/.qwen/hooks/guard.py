#!/usr/bin/env python3
# PreToolUse-hook: запрещает сетевые команды из shell.
import json
import re
import sys

data = json.load(sys.stdin)
tool = data.get("tool_name", "")
args = data.get("tool_input", {})


def decide(decision, reason):
    print(json.dumps({"hookSpecificOutput": {
        "hookEventName": "PreToolUse",
        "permissionDecision": decision,       # allow / deny
        "permissionDecisionReason": reason,
    }}, ensure_ascii=True))
    sys.exit(0)


if tool == "run_shell_command" and re.search(r"\b(curl|wget|nc|ncat)\b", args.get("command", "")):
    decide("deny", "Сетевые команды запрещены.")

decide("allow", "ok")