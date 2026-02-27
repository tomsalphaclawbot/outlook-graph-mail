#!/usr/bin/env bash
set -euo pipefail

TOKEN_FILE="${OUTLOOK_GRAPH_TOKEN_FILE:-$HOME/.openclaw/workspace/state/outlook_graph_token.json}"
CLIENT_ID="${OUTLOOK_GRAPH_CLIENT_ID:-4a8603ea-1f64-4963-85bc-27c5758fd6f3}"
TENANT="${OUTLOOK_GRAPH_TENANT:-consumers}"
TOP="${1:-10}"
SUBJECT_CONTAINS="${2:-}"

python3 - "$TOKEN_FILE" "$CLIENT_ID" "$TENANT" "$TOP" "$SUBJECT_CONTAINS" <<'PY'
import json, sys, time, urllib.parse, urllib.request, urllib.error, pathlib

token_file, client_id, tenant, top_s, subject_contains = sys.argv[1:]
top = int(top_s)
path = pathlib.Path(token_file)
if not path.exists():
    print('missing_token_file')
    raise SystemExit(1)

tok = json.loads(path.read_text())

def post(url, data):
    body = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, data=body, headers={'Content-Type':'application/x-www-form-urlencoded'})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.getcode(), json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode())
        except Exception:
            return e.code, {'error': 'http_error', 'error_description': str(e)}

def ensure_access_token(t):
    if 'access_token' in t and t.get('expires_in'):
        return t
    if 'refresh_token' not in t:
        return t
    c, rt = post(f'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token', {
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': t['refresh_token'],
        'scope': 'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send offline_access openid profile'
    })
    if c >= 400 or 'access_token' not in rt:
        print('refresh_failed', rt)
        raise SystemExit(2)
    path.write_text(json.dumps(rt, indent=2)); path.chmod(0o600)
    return rt

tok = ensure_access_token(tok)
access = tok['access_token']
url = f'https://graph.microsoft.com/v1.0/me/messages?$top={top}&$select=id,parentFolderId,receivedDateTime,subject,from,webLink'
req = urllib.request.Request(url, headers={
    'Authorization': f'Bearer {access}',
    'Accept': 'application/json'
})
try:
    with urllib.request.urlopen(req, timeout=30) as r:
        data = json.loads(r.read().decode())
except urllib.error.HTTPError as e:
    print('graph_error', e.code, e.read().decode()[:500])
    raise SystemExit(3)

items = data.get('value', [])
if subject_contains:
    s = subject_contains.lower()
    items = [m for m in items if s in (m.get('subject') or '').lower()]

for m in items:
    sender = (((m.get('from') or {}).get('emailAddress') or {}).get('address') or '')
    print(f"{m.get('id','')} | {m.get('receivedDateTime','')} | {sender} | {m.get('subject','')}")
print(f"count={len(items)}")
PY
