#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 2 ]]; then
  echo "usage: $0 <message_id> <destination_folder_id_or_wellknown_name>"
  exit 1
fi

MSG_ID="$1"
DEST="$2"  # e.g. inbox|archive|junkemail|deleteditems or folder id
TOKEN_FILE="${OUTLOOK_GRAPH_TOKEN_FILE:-$HOME/.openclaw/workspace/state/outlook_graph_token.json}"

python3 - "$MSG_ID" "$DEST" "$TOKEN_FILE" <<'PY'
import json, sys, urllib.parse, urllib.request, urllib.error, pathlib
msg_id,dest,token_file=sys.argv[1:]
path=pathlib.Path(token_file)
if not path.exists():
    print('missing_token_file'); raise SystemExit(1)
tok=json.loads(path.read_text())
url=f"https://graph.microsoft.com/v1.0/me/messages/{urllib.parse.quote(msg_id, safe='')}/move"
payload={'destinationId':dest}
req=urllib.request.Request(url, data=json.dumps(payload).encode(), headers={'Authorization':f"Bearer {tok['access_token']}", 'Content-Type':'application/json','Accept':'application/json'}, method='POST')
try:
    with urllib.request.urlopen(req, timeout=30) as r:
        body=json.loads(r.read().decode())
        print('move_ok', body.get('id',''))
except urllib.error.HTTPError as e:
    print('move_error', e.code, e.read().decode()[:500]); raise SystemExit(2)
PY
