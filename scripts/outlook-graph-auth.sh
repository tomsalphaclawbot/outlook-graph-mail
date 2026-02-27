#!/usr/bin/env bash
set -euo pipefail

CLIENT_ID="${OUTLOOK_GRAPH_CLIENT_ID:-4a8603ea-1f64-4963-85bc-27c5758fd6f3}"
TENANT="${OUTLOOK_GRAPH_TENANT:-consumers}"
TOKEN_FILE="${OUTLOOK_GRAPH_TOKEN_FILE:-$HOME/.openclaw/workspace/state/outlook_graph_token.json}"
SCOPES="${OUTLOOK_GRAPH_SCOPES:-https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send offline_access openid profile}"

python3 - "$CLIENT_ID" "$TENANT" "$TOKEN_FILE" "$SCOPES" <<'PY'
import json, sys, time, urllib.parse, urllib.request, urllib.error, pathlib
client_id, tenant, token_file, scopes = sys.argv[1:]

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

code, dc = post(f'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode', {
    'client_id': client_id,
    'scope': scopes,
})
if code >= 400:
    print('device_code_error:', dc)
    raise SystemExit(1)

print('Open:', dc.get('verification_uri') or dc.get('verification_uri_complete'))
print('Code:', dc.get('user_code'))
print('Waiting for approval...')

interval = int(dc.get('interval', 5))
expires = int(dc.get('expires_in', 900))
start = time.time()

tok = None
while time.time() - start < expires:
    code, tr = post(f'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token', {
        'grant_type': 'urn:ietf:params:oauth:grant-type:device_code',
        'client_id': client_id,
        'device_code': dc['device_code'],
    })
    if code < 400 and 'access_token' in tr:
        tok = tr
        break
    if tr.get('error') in ('authorization_pending', 'slow_down'):
        time.sleep(interval + (2 if tr.get('error') == 'slow_down' else 0))
        continue
    print('token_error:', tr)
    raise SystemExit(2)

if not tok:
    print('token_timeout')
    raise SystemExit(3)

path = pathlib.Path(token_file)
path.parent.mkdir(parents=True, exist_ok=True)
path.write_text(json.dumps(tok, indent=2))
path.chmod(0o600)
print(f'saved_token: {path}')
print('auth_ok')
PY
