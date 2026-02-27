#!/usr/bin/env python3
import json, sys, time, urllib.parse, urllib.request, urllib.error, pathlib

if len(sys.argv) < 4:
    print("usage: outlook-graph-token.py <token_file> <client_id> <tenant> [--force]", file=sys.stderr)
    raise SystemExit(1)

token_file, client_id, tenant = sys.argv[1:4]
force = "--force" in sys.argv[4:]
path = pathlib.Path(token_file)
if not path.exists():
    print('missing_token_file')
    raise SystemExit(1)

tok = json.loads(path.read_text())
SCOPES = 'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send offline_access openid profile'


def post_form(url, data):
    body = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, data=body, headers={'Content-Type': 'application/x-www-form-urlencoded'})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.getcode(), json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode())
        except Exception:
            return e.code, {'error': 'http_error', 'error_description': str(e)}


def needs_refresh(t):
    if force:
        return True
    if 'refresh_token' not in t:
        return False
    if 'access_token' not in t:
        return True
    exp = int(t.get('expires_in') or 0)
    obt = int(t.get('obtained_at') or 0)
    if exp <= 0 or obt <= 0:
        return True
    return int(time.time()) >= (obt + exp - 300)


if needs_refresh(tok):
    if 'refresh_token' not in tok:
        print('refresh_token_missing')
        raise SystemExit(2)
    c, rt = post_form(f'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token', {
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': tok['refresh_token'],
        'scope': SCOPES,
    })
    if c >= 400 or 'access_token' not in rt:
        print('refresh_failed', rt)
        raise SystemExit(3)
    rt['obtained_at'] = int(time.time())
    tok = rt
    path.write_text(json.dumps(tok, indent=2))
    path.chmod(0o600)

print(tok['access_token'])
