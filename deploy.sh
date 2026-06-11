#!/usr/bin/env bash
#
# Full deploy chain:  commit → clasp push → clasp deploy → git push
#
# Usage:
#   ./deploy.sh                             # full chain, auto-generated message
#   ./deploy.sh "commit message"            # full chain, your message
#   ./deploy.sh --no-deploy                 # auto message, skip web-app bump
#   ./deploy.sh "commit message" --no-deploy  # your message, skip web-app bump
#
# If you don't pass a message, one is built from the changed file list.
# See DEPLOY.md for the manual steps and the why behind each one.

set -euo pipefail

# Parse args: a lone "--no-deploy" may appear with or without a message.
MSG=""
DEPLOY=1
for arg in "$@"; do
  if [[ "$arg" == "--no-deploy" ]]; then
    DEPLOY=0
  else
    MSG="$arg"
  fi
done

# Auto-generate a commit message from the changed files when none is given.
if [[ -z "$MSG" ]]; then
  FILES=$(git status --porcelain | sed 's/^...//' | xargs -n1 basename | sort -u | paste -sd ', ' -)
  if [[ -z "$FILES" ]]; then
    echo "✗ Nothing to commit — working tree is clean."
    exit 1
  fi
  MSG="chore: update ${FILES}"
  echo "▸ auto message: \"$MSG\""
fi

DEPLOYMENT_ID="AKfycbxh3-EQ7VUbbf1WMn9q9aSBBCnSimhRST5QwEjs6VDXij07JwJQMP0Md99DqpqrFNmU"

# 1. commit (docs belong in the SAME commit as code — see CLAUDE.md)
echo "▸ git commit"
git add -A
git commit -m "$MSG"

# 2. push code to Apps Script (updates editor / automations immediately)
echo "▸ clasp push"
clasp push -f

# 3. cut a new web-app version (live /exec URL keeps the old one until you do)
if [[ "$DEPLOY" == "1" ]]; then
  echo "▸ clasp deploy"
  clasp deploy --deploymentId "$DEPLOYMENT_ID" --description "$MSG"
else
  echo "▸ clasp deploy  (skipped: --no-deploy)"
fi

# 4. push to GitHub
echo "▸ git push"
git push origin main

echo "✓ Done."
