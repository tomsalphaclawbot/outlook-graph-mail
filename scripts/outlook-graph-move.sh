#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 2 ]]; then
  echo "usage: $0 <message_id> <destination_folder_id_or_wellknown_name>"
  exit 1
fi

MSG_ID="$1"
DEST="$2"  # e.g. inbox|archive|junkemail|deleteditems or folder id
TOKEN_FILE="${OUTLOOK_GRAPH_TOKEN_FILE:-$HOME/.openclaw/workspace/state/outlook_graph_token.json}"
CLIENT_ID="${OUTLOOK_GRAPH_CLIENT_ID:-4a8603ea-1f64-4963-85bc-27c5758fd6f3}"
TENANT="${OUTLOOK_GRAPH_TENANT:-consumers}"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TOKEN_HELPER="$SCRIPT_DIR/outlook-graph-token.py"

python3 - "$MSG_ID" "$DEST" "$TOKEN_FILE" "$CLIENT_ID" "$TENANT" "$TOKEN_HELPER" <<'PY'
import json, sys, subprocess, urllib.parse, urllib.request, urllib.error

msg_id, dest, token_file, client_id, tenant, helper = sys.argv[1:]

def get_token(force=False):
    cmd = [helper, token_file, client_id, tenant]
    if force:
        cmd.append('--force')
    return subprocess.check_output(cmd, text=True).strip()

def move(token):
    url = f"https://graph.microsoft.com/v1.0/me/messages/{urllib.parse.quote(msg_id, safe='')}/move"
    payload = {'destinationId': dest}
    req = urllib.request.Request(url, data=json.dumps(payload).encode(), headers={
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }, method='POST')
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())

try:
    body = move(get_token(False))
except urllib.error.HTTPError as e:
    if e.code == 401:
        body = move(get_token(True))
    else:
        print('move_error', e.code, e.read().decode()[:500])
        raise SystemExit(2)

print('move_ok', body.get('id', ''))
PY
