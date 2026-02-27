#!/usr/bin/env bash
set -euo pipefail

TOKEN_FILE="${OUTLOOK_GRAPH_TOKEN_FILE:-$HOME/.openclaw/workspace/state/outlook_graph_token.json}"
CLIENT_ID="${OUTLOOK_GRAPH_CLIENT_ID:-4a8603ea-1f64-4963-85bc-27c5758fd6f3}"
TENANT="${OUTLOOK_GRAPH_TENANT:-consumers}"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TOKEN_HELPER="$SCRIPT_DIR/outlook-graph-token.py"

python3 - "$TOKEN_FILE" "$CLIENT_ID" "$TENANT" "$TOKEN_HELPER" <<'PY'
import json, sys, subprocess, urllib.request, urllib.error

token_file, client_id, tenant, helper = sys.argv[1:]

def get_token(force=False):
    cmd = [helper, token_file, client_id, tenant]
    if force:
        cmd.append('--force')
    return subprocess.check_output(cmd, text=True).strip()

def fetch(token):
    req = urllib.request.Request(
        'https://graph.microsoft.com/v1.0/me/mailFolders?$top=200&$select=id,displayName,totalItemCount,unreadItemCount',
        headers={'Authorization': f'Bearer {token}', 'Accept': 'application/json'}
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())

try:
    data = fetch(get_token(False))
except urllib.error.HTTPError as e:
    if e.code == 401:
        data = fetch(get_token(True))
    else:
        print('folders_error', e.code, e.read().decode()[:400])
        raise SystemExit(2)

for f in data.get('value', []):
    print(f"{f.get('displayName')} | {f.get('id')} | total={f.get('totalItemCount')} unread={f.get('unreadItemCount')}")
PY
