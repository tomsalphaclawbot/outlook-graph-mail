#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 1 ]]; then
  echo "usage: $0 <message_id>"
  exit 1
fi

MSG_ID="$1"
TOKEN_FILE="${OUTLOOK_GRAPH_TOKEN_FILE:-$HOME/.openclaw/workspace/state/outlook_graph_token.json}"
CLIENT_ID="${OUTLOOK_GRAPH_CLIENT_ID:-4a8603ea-1f64-4963-85bc-27c5758fd6f3}"
TENANT="${OUTLOOK_GRAPH_TENANT:-consumers}"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TOKEN_HELPER="$SCRIPT_DIR/outlook-graph-token.py"

python3 - "$MSG_ID" "$TOKEN_FILE" "$CLIENT_ID" "$TENANT" "$TOKEN_HELPER" <<'PY'
import json, sys, subprocess, urllib.parse, urllib.request, urllib.error

msg_id, token_file, client_id, tenant, helper = sys.argv[1:]

def get_token(force=False):
    cmd = [helper, token_file, client_id, tenant]
    if force:
        cmd.append('--force')
    return subprocess.check_output(cmd, text=True).strip()

def archive(token):
    encoded = urllib.parse.quote(msg_id, safe='')
    url = f'https://graph.microsoft.com/v1.0/me/messages/{encoded}/move'
    payload = {'destinationId': 'archive'}
    req = urllib.request.Request(url, data=json.dumps(payload).encode(), headers={
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }, method='POST')
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())

try:
    body = archive(get_token(False))
except urllib.error.HTTPError as e:
    if e.code == 401:
        body = archive(get_token(True))
    else:
        print('archive_error', e.code, e.read().decode()[:500])
        raise SystemExit(3)

print('archive_ok', body.get('id', ''))
PY
