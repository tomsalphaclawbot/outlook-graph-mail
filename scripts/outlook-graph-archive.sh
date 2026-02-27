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

python3 - "$MSG_ID" "$TOKEN_FILE" "$CLIENT_ID" "$TENANT" <<'PY'
import json, sys, urllib.parse, urllib.request, urllib.error, pathlib
msg_id, token_file, client_id, tenant = sys.argv[1:]
path = pathlib.Path(token_file)
if not path.exists():
    print('missing_token_file'); raise SystemExit(1)

tok = json.loads(path.read_text())

def post_form(url, data):
    body = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, data=body, headers={'Content-Type':'application/x-www-form-urlencoded'})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.getcode(), json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        try: return e.code, json.loads(e.read().decode())
        except Exception: return e.code, {'error':'http_error'}

def ensure_token(t):
    if 'access_token' in t: return t
    c, rt = post_form(f'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token', {
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': t['refresh_token'],
        'scope': 'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send offline_access openid profile'
    })
    if c >= 400 or 'access_token' not in rt:
        print('refresh_failed', rt); raise SystemExit(2)
    path.write_text(json.dumps(rt, indent=2)); path.chmod(0o600)
    return rt

tok = ensure_token(tok)
encoded = urllib.parse.quote(msg_id, safe='')
url = f'https://graph.microsoft.com/v1.0/me/messages/{encoded}/move'
payload = {'destinationId': 'archive'}
req = urllib.request.Request(url, data=json.dumps(payload).encode(), headers={
    'Authorization': f"Bearer {tok['access_token']}",
    'Content-Type': 'application/json',
    'Accept': 'application/json'
}, method='POST')
try:
    with urllib.request.urlopen(req, timeout=30) as r:
        body = json.loads(r.read().decode())
        print('archive_ok', body.get('id',''))
except urllib.error.HTTPError as e:
    print('archive_error', e.code, e.read().decode()[:500])
    raise SystemExit(3)
PY
