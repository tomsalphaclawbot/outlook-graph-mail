#!/usr/bin/env bash
set -euo pipefail

FOLDER="${1:-inbox}"   # inbox|archive|junkemail|sentitems|deleteditems or folder id
TOP="${2:-20}"
SUBJECT_CONTAINS="${3:-}"
TOKEN_FILE="${OUTLOOK_GRAPH_TOKEN_FILE:-$HOME/.openclaw/workspace/state/outlook_graph_token.json}"
CLIENT_ID="${OUTLOOK_GRAPH_CLIENT_ID:-4a8603ea-1f64-4963-85bc-27c5758fd6f3}"
TENANT="${OUTLOOK_GRAPH_TENANT:-consumers}"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TOKEN_HELPER="$SCRIPT_DIR/outlook-graph-token.py"

python3 - "$FOLDER" "$TOP" "$SUBJECT_CONTAINS" "$TOKEN_FILE" "$CLIENT_ID" "$TENANT" "$TOKEN_HELPER" <<'PY'
import json, sys, subprocess, urllib.parse, urllib.request, urllib.error

folder, top_s, subj, token_file, client_id, tenant, helper = sys.argv[1:]
top = int(top_s)

def get_token(force=False):
    cmd = [helper, token_file, client_id, tenant]
    if force:
        cmd.append('--force')
    return subprocess.check_output(cmd, text=True).strip()

def fetch(token):
    folder_enc = urllib.parse.quote(folder, safe='')
    url = f'https://graph.microsoft.com/v1.0/me/mailFolders/{folder_enc}/messages?$top={top}&$select=id,receivedDateTime,subject,from,parentFolderId'
    req = urllib.request.Request(url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/json'})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())

try:
    data = fetch(get_token(False))
except urllib.error.HTTPError as e:
    if e.code == 401:
        data = fetch(get_token(True))
    else:
        print('read_folder_error', e.code, e.read().decode()[:500])
        raise SystemExit(2)

items = data.get('value', [])
if subj:
    s = subj.lower()
    items = [m for m in items if s in (m.get('subject') or '').lower()]
for m in items:
    sender = (((m.get('from') or {}).get('emailAddress') or {}).get('address') or '')
    print(f"{m.get('id','')} | {m.get('receivedDateTime','')} | {sender} | {m.get('subject','')}")
print(f"count={len(items)}")
PY
