#!/usr/bin/env bash
set -euo pipefail

FOLDER="${1:-inbox}"   # inbox|archive|junkemail|sentitems|deleteditems or folder id
TOP="${2:-20}"
SUBJECT_CONTAINS="${3:-}"
TOKEN_FILE="${OUTLOOK_GRAPH_TOKEN_FILE:-$HOME/.openclaw/workspace/state/outlook_graph_token.json}"

python3 - "$FOLDER" "$TOP" "$SUBJECT_CONTAINS" "$TOKEN_FILE" <<'PY'
import json, sys, urllib.parse, urllib.request, urllib.error, pathlib
folder, top_s, subj, token_file = sys.argv[1:]
top=int(top_s)
path=pathlib.Path(token_file)
if not path.exists():
    print('missing_token_file'); raise SystemExit(1)
tok=json.loads(path.read_text())
folder_enc=urllib.parse.quote(folder, safe='')
url=f'https://graph.microsoft.com/v1.0/me/mailFolders/{folder_enc}/messages?$top={top}&$select=id,receivedDateTime,subject,from,parentFolderId'
req=urllib.request.Request(url, headers={'Authorization':f"Bearer {tok['access_token']}", 'Accept':'application/json'})
try:
    with urllib.request.urlopen(req, timeout=30) as r: data=json.loads(r.read().decode())
except urllib.error.HTTPError as e:
    print('read_folder_error', e.code, e.read().decode()[:500]); raise SystemExit(2)
items=data.get('value',[])
if subj:
    s=subj.lower(); items=[m for m in items if s in (m.get('subject') or '').lower()]
for m in items:
    sender=(((m.get('from') or {}).get('emailAddress') or {}).get('address') or '')
    print(f"{m.get('id','')} | {m.get('receivedDateTime','')} | {sender} | {m.get('subject','')}")
print(f"count={len(items)}")
PY
