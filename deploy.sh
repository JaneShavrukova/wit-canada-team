#!/usr/bin/env bash
#
# Full deploy chain:  commit → clasp push → clasp deploy → git push
#
# Usage:
#   ./deploy.sh "commit message"            # full chain (code + web app)
#   ./deploy.sh "commit message" --no-deploy  # skip the web-app version bump
#                                             # (use when no served .html / doGet changed)
#
# See DEPLOY.md for the manual steps and the why behind each one.

set -euo pipefail

MSG="${1:-}"
if [[ -z "$MSG" ]]; then
  echo "✗ Commit message required."
  echo '  Usage: ./deploy.sh "commit message" [--no-deploy]'
  exit 1
fi

DEPLOY=1
[[ "${2:-}" == "--no-deploy" ]] && DEPLOY=0

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
