#!/bin/bash

# Script to implement features from to-support.json using Claude
# Usage: ./implement-features.sh [--dry-run]

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
JSON_FILE="$SCRIPT_DIR/to-support.json"

DRY_RUN=false
if [[ "$1" == "--dry-run" ]]; then
  DRY_RUN=true
  echo "=== DRY RUN MODE ==="
fi

# Check dependencies
if ! command -v jq &> /dev/null; then
  echo "Error: jq is required. Install with: brew install jq"
  exit 1
fi

if ! command -v claude &> /dev/null; then
  echo "Error: claude CLI is required."
  exit 1
fi

# Get the first unimplemented feature
get_next_feature() {
  jq -r '
    map(select(.implemented == false))
    | sort_by(.priority)
    | first
    | if . == null then empty else . end
  ' "$JSON_FILE"
}

# Mark a feature as implemented
mark_implemented() {
  local feature_id="$1"
  local tmp_file=$(mktemp)

  jq --arg id "$feature_id" '
    map(if .id == $id then .implemented = true else . end)
  ' "$JSON_FILE" > "$tmp_file" && mv "$tmp_file" "$JSON_FILE"

  echo "Marked '$feature_id' as implemented in to-support.json"
}

# Count remaining features
count_remaining() {
  jq '[.[] | select(.implemented == false)] | length' "$JSON_FILE"
}

echo "=== MSG-2-EML Feature Implementation Script ==="
echo ""

remaining=$(count_remaining)
echo "Features remaining to implement: $remaining"
echo ""

while true; do
  feature_json=$(get_next_feature)

  if [[ -z "$feature_json" ]]; then
    echo "All features have been implemented!"
    break
  fi

  feature_id=$(echo "$feature_json" | jq -r '.id')
  feature_desc=$(echo "$feature_json" | jq -r '.description')
  feature_priority=$(echo "$feature_json" | jq -r '.priority')

  echo "----------------------------------------"
  echo "Feature: $feature_id (priority: $feature_priority)"
  echo "Description: $feature_desc"
  echo "----------------------------------------"
  echo ""

  prompt="Implement this feature in the MSG to EML converter project:

Feature ID: $feature_id
Description: $feature_desc

Instructions:
1. Read the existing code in src/server/msg-to-eml.ts to understand the current implementation
2. Implement the feature described above
3. Add or update tests in src/server/msg-to-eml.test.ts to cover the new functionality
4. Run 'npm test' to verify all tests pass
5. If tests pass, commit with message: 'feat($feature_id): <brief description>'
6. If you need to install new dependencies, use 'npm install <package> --save'

Do not deploy. Just implement, test, and commit."

  if [[ "$DRY_RUN" == true ]]; then
    echo "[DRY RUN] Would execute:"
    echo "claude --dangerously-skip-permissions -p \"$prompt\""
    echo ""
    echo "Press Enter to simulate completion, or Ctrl+C to exit"
    read
  else
    echo "Calling Claude to implement feature..."
    echo ""

    # Run claude with permissions for this project
    if claude --dangerously-skip-permissions -p "$prompt"; then
      echo ""
      echo "Claude completed work on '$feature_id'"
    else
      echo ""
      echo "Warning: Claude exited with non-zero status for '$feature_id'"
      echo "Check the implementation manually before continuing."
      echo "Press Enter to mark as implemented anyway, or Ctrl+C to exit"
      read
    fi
  fi

  # Mark as implemented
  mark_implemented "$feature_id"

  remaining=$(count_remaining)
  echo ""
  echo "Features remaining: $remaining"
  echo ""

  if [[ $remaining -eq 0 ]]; then
    break
  fi

  # Small delay between features
  sleep 2
done

echo ""
echo "=== Implementation Complete ==="
echo "Don't forget to:"
echo "  1. Review all changes: git log --oneline -10"
echo "  2. Run full test suite: npm test"
echo "  3. Deploy when ready: railway up --service msg-2-eml"
