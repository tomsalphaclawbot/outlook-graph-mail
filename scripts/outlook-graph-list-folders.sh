#!/usr/bin/env bash
set -euo pipefail

TOKEN_FILE="${OUTLOOK_GRAPH_TOKEN_FILE:-$HOME/.openclaw/workspace/state/outlook_graph_token.json}"

python3 - "$TOKEN_FILE" <<'PY'
import json, sys, urllib.request, urllib.error, pathlib
path=pathlib.Path(sys.argv[1])
if not path.exists():
    print('missing_token_file'); raise SystemExit(1)
tok=json.loads(path.read_text())
req=urllib.request.Request('https://graph.microsoft.com/v1.0/me/mailFolders?$top=200&$select=id,displayName,totalItemCount,unreadItemCount', headers={'Authorization':f"Bearer {tok['access_token']}", 'Accept':'application/json'})
try:
    with urllib.request.urlopen(req, timeout=30) as r: data=json.loads(r.read().decode())
except urllib.error.HTTPError as e:
    print('folders_error', e.code, e.read().decode()[:400]); raise SystemExit(2)
for f in data.get('value',[]):
    print(f"{f.get('displayName')} | {f.get('id')} | total={f.get('totalItemCount')} unread={f.get('unreadItemCount')}")
PY
