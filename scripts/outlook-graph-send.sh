#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 3 ]]; then
  echo "usage: $0 <to_email> <subject> <body_text>"
  exit 1
fi

TO="$1"
SUBJECT="$2"
BODY="$3"
TOKEN_FILE="${OUTLOOK_GRAPH_TOKEN_FILE:-$HOME/.openclaw/workspace/state/outlook_graph_token.json}"
CLIENT_ID="${OUTLOOK_GRAPH_CLIENT_ID:-4a8603ea-1f64-4963-85bc-27c5758fd6f3}"
TENANT="${OUTLOOK_GRAPH_TENANT:-consumers}"

python3 - "$TO" "$SUBJECT" "$BODY" "$TOKEN_FILE" "$CLIENT_ID" "$TENANT" <<'PY'
import json, sys, urllib.parse, urllib.request, urllib.error, pathlib

to_email, subject, body_text, token_file, client_id, tenant = sys.argv[1:]
path = pathlib.Path(token_file)
if not path.exists():
    print('missing_token_file')
    raise SystemExit(1)

tok = json.loads(path.read_text())

def post_form(url, data):
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

if 'access_token' not in tok and 'refresh_token' in tok:
    c, rt = post_form(f'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token', {
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': tok['refresh_token'],
        'scope': 'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send offline_access openid profile'
    })
    if c >= 400 or 'access_token' not in rt:
        print('refresh_failed', rt)
        raise SystemExit(2)
    tok = rt
    path.write_text(json.dumps(tok, indent=2)); path.chmod(0o600)

payload = {
  'message': {
    'subject': subject,
    'body': {'contentType': 'Text', 'content': body_text},
    'toRecipients': [{'emailAddress': {'address': to_email}}]
  },
  'saveToSentItems': True
}

req = urllib.request.Request(
    'https://graph.microsoft.com/v1.0/me/sendMail',
    data=json.dumps(payload).encode(),
    headers={
      'Authorization': f"Bearer {tok['access_token']}",
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    method='POST'
)
try:
    with urllib.request.urlopen(req, timeout=30) as r:
        print('send_status', r.status)
except urllib.error.HTTPError as e:
    print('send_error', e.code, e.read().decode()[:500])
    raise SystemExit(3)
print('send_ok')
PY
