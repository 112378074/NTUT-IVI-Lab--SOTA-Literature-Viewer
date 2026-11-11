#!/usr/bin/env python3
"""One-time Gmail authorisation — run this on YOUR OWN computer, not in the cloud.

It opens a browser, you sign in to the Gmail account that should send the
harvest emails, and it prints a refresh token.  Store that token (plus the
client ID and secret) as environment variables of the scheduled
environment; scripts/send_mail.py uses them from then on.

Only the gmail.send scope is requested: the token can send mail as you,
but cannot read, list or delete anything in the mailbox.

    python scripts/gmail_auth.py --client-id XXX.apps.googleusercontent.com --client-secret YYY

Needs an OAuth client of type "Desktop app" in Google Cloud Console with
the Gmail API enabled (steps in scripts/README.md).
"""
import argparse, http.server, json, secrets, sys, urllib.parse, urllib.request, webbrowser

AUTH_URL = 'https://accounts.google.com/o/oauth2/v2/auth'
TOKEN_URL = 'https://oauth2.googleapis.com/token'
SCOPE = 'https://www.googleapis.com/auth/gmail.send'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--client-id', required=True)
    ap.add_argument('--client-secret', required=True)
    ap.add_argument('--port', type=int, default=8765)
    a = ap.parse_args()

    redirect = f'http://127.0.0.1:{a.port}/'
    state = secrets.token_urlsafe(16)
    got = {}

    class H(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            got.update({k: v[0] for k, v in q.items()})
            ok = 'code' in q and q.get('state', [''])[0] == state
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(('<h2>授權完成，可以關閉這個分頁回到終端機。</h2>' if ok
                              else '<h2>授權失敗，請回到終端機查看訊息。</h2>').encode())

        def log_message(self, *args):
            pass

    srv = http.server.HTTPServer(('127.0.0.1', a.port), H)
    srv.timeout = 300

    url = AUTH_URL + '?' + urllib.parse.urlencode({
        'client_id': a.client_id, 'redirect_uri': redirect, 'response_type': 'code',
        'scope': SCOPE, 'access_type': 'offline', 'prompt': 'consent', 'state': state})
    print('Opening the browser. If it does not open, paste this URL:\n\n' + url + '\n')
    webbrowser.open(url)
    # Ignore stray requests (e.g. /favicon.ico) until the redirect carrying code/error lands.
    while not ('code' in got or 'error' in got):
        srv.handle_request()

    if got.get('state') != state or 'code' not in got:
        sys.exit(f'Authorisation failed: {got.get("error", "state mismatch or no code")}')

    body = urllib.parse.urlencode({
        'code': got['code'], 'client_id': a.client_id, 'client_secret': a.client_secret,
        'redirect_uri': redirect, 'grant_type': 'authorization_code'}).encode()
    with urllib.request.urlopen(urllib.request.Request(TOKEN_URL, data=body), timeout=30) as r:
        tok = json.loads(r.read())
    if 'refresh_token' not in tok:
        sys.exit('No refresh token returned. Revoke the app at https://myaccount.google.com/permissions and run again.')

    print('Done. Set these as environment variables of the scheduled environment:\n')
    print(f'GMAIL_CLIENT_ID={a.client_id}')
    print(f'GMAIL_CLIENT_SECRET={a.client_secret}')
    print(f'GMAIL_REFRESH_TOKEN={tok["refresh_token"]}')
    print('NOTIFY_TO=azaz31855@gmail.com,fctien@ntut.edu.tw\n')
    print('Keep the refresh token private: it can send mail as you.')


if __name__ == '__main__':
    main()
