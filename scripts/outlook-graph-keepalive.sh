#!/usr/bin/env bash
set -euo pipefail

TOKEN_FILE="${OUTLOOK_GRAPH_TOKEN_FILE:-$HOME/.openclaw/workspace/state/outlook_graph_token.json}"
CLIENT_ID="${OUTLOOK_GRAPH_CLIENT_ID:-4a8603ea-1f64-4963-85bc-27c5758fd6f3}"
TENANT="${OUTLOOK_GRAPH_TENANT:-consumers}"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TOKEN_HELPER="$SCRIPT_DIR/outlook-graph-token.py"

TOKEN="$($TOKEN_HELPER "$TOKEN_FILE" "$CLIENT_ID" "$TENANT")"

python3 - "$TOKEN" <<'PY'
import json, sys, urllib.request

token = sys.argv[1]
req = urllib.request.Request(
    'https://graph.microsoft.com/v1.0/me/mailFolders/inbox?$select=id,displayName,unreadItemCount',
    headers={'Authorization': f'Bearer {token}', 'Accept': 'application/json'}
)
with urllib.request.urlopen(req, timeout=30) as r:
    data = json.loads(r.read().decode())
print('keepalive_ok', data.get('displayName'), 'unread=', data.get('unreadItemCount'))
PY
