#!/usr/bin/env bash
# PostToolUse guard: reject em dashes (U+2014) and spaced en dashes in edited files.
# Exit 2 feeds the message back to Claude so the lines get fixed immediately.
# Fails open (exit 0) on anything unexpected so it can never wedge a session.
set -uo pipefail

command -v jq >/dev/null 2>&1 || exit 0

payload=$(cat) || exit 0
file=$(printf '%s' "$payload" | jq -r '.tool_input.file_path // empty' 2>/dev/null) || exit 0
[ -n "${file:-}" ] && [ -f "$file" ] || exit 0

# These two document the characters they ban.
case "$file" in
  */.claude/skills/humanize/*) exit 0 ;;
  */.claude/hooks/check-em-dash.sh) exit 0 ;;
esac

# Literal byte match: locale-independent, and never matches a plain hyphen.
hits=$(grep -n -e '—' -e ' – ' -- "$file" 2>/dev/null | head -20) || exit 0
[ -n "$hits" ] || exit 0

{
  echo "BLOCKED by no-em-dash rule: $file"
  echo "$hits"
  echo
  echo "Fix these now. Parenthetical -> commas or parentheses. Before an"
  echo "explanation -> colon. Joining two clauses -> semicolon or two sentences."
  echo "Prose range -> 'to'. Hyphens in compounds and identifiers are fine."
} >&2
exit 2
