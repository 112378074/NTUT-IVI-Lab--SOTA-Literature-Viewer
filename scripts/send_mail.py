#!/usr/bin/env python3
"""Send the harvest summary through the Gmail API.

Why the Gmail API rather than SMTP: the scheduled harvest runs in a remote
container that blocks raw SMTP (ports 587/465) and most mail-API hosts, but
gmail.googleapis.com and oauth2.googleapis.com are reachable over HTTPS.
Mail goes out from the authorising Gmail account itself, so no sender
domain has to be verified.

Config is read from the process environment, with scripts/.env overlaid on
top (the file wins).  The remote runner never has scripts/.env — it is
gitignored and every run is a fresh clone — so set these as environment
variables there:

    GMAIL_CLIENT_ID       OAuth client ID (Desktop app type)
    GMAIL_CLIENT_SECRET   OAuth client secret
    GMAIL_REFRESH_TOKEN   from scripts/gmail_auth.py, run once on your own machine
    NOTIFY_TO             comma-separated recipients

Usage:
    python scripts/send_mail.py --check                    # config + token only, sends nothing
    python scripts/send_mail.py --test                     # one test message to NOTIFY_TO
    python scripts/send_mail.py --subject S --body-file F [--html-file H] [--to a,b]

Exit status: 0 sent (or --check passed), 2 not configured, 1 failed.
"""
import argparse, base64, json, os, re, ssl, sys, urllib.error, urllib.parse, urllib.request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, make_msgid
from pathlib import Path

ENV_FILE = Path(__file__).resolve().parent / '.env'
KEYS = ('GMAIL_CLIENT_ID', 'GMAIL_CLIENT_SECRET', 'GMAIL_REFRESH_TOKEN', 'NOTIFY_TO')
TOKEN_URL = 'https://oauth2.googleapis.com/token'
SEND_URL = 'https://gmail.googleapis.com/gmail/v1/users/me/messages/send'


def _ssl_ctx():
    # Behind the agent proxy the CA bundle is exported via SSL_CERT_FILE.
    for var in ('SSL_CERT_FILE', 'REQUESTS_CA_BUNDLE'):
        p = os.environ.get(var)
        if p and os.path.exists(p):
            return ssl.create_default_context(cafile=p)
    return ssl.create_default_context()


SSL_CTX = _ssl_ctx()


def load_config():
    cfg = {k: os.environ[k] for k in KEYS if os.environ.get(k)}
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text(encoding='utf-8').splitlines():
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                k, v = line.split('=', 1)
                if k.strip() in KEYS:
                    cfg[k.strip()] = v.strip().strip('"').strip("'")
    return cfg


def recipients(raw):
    return [a for a in re.split(r'[,;\s]+', raw or '') if a]


def _post(url, data, headers):
    req = urllib.request.Request(url, data=data, headers=headers, method='POST')
    try:
        with urllib.request.urlopen(req, timeout=30, context=SSL_CTX) as r:
            return r.getcode(), r.read().decode('utf-8', 'replace')
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode('utf-8', 'replace')


def access_token(cfg):
    body = urllib.parse.urlencode({
        'client_id': cfg['GMAIL_CLIENT_ID'],
        'client_secret': cfg['GMAIL_CLIENT_SECRET'],
        'refresh_token': cfg['GMAIL_REFRESH_TOKEN'],
        'grant_type': 'refresh_token',
    }).encode()
    status, text = _post(TOKEN_URL, body, {'Content-Type': 'application/x-www-form-urlencoded'})
    if status != 200:
        hint = ''
        if 'invalid_grant' in text:
            hint = (' — refresh token expired or revoked. Re-run scripts/gmail_auth.py. '
                    'If this recurs about weekly, the OAuth app is still in "Testing": '
                    'publish it to "In production" so tokens stop expiring after 7 days.')
        raise RuntimeError(f'token exchange failed: HTTP {status}: {text[:300]}{hint}')
    return json.loads(text)['access_token']


def build_raw(to, subject, text, html=None):
    msg = MIMEMultipart('alternative')
    msg['To'] = ', '.join(to)
    msg['Subject'] = subject
    msg['Date'] = formatdate(localtime=True)
    msg['Message-ID'] = make_msgid()
    msg.attach(MIMEText(text, 'plain', 'utf-8'))
    if html:
        msg.attach(MIMEText(html, 'html', 'utf-8'))
    return base64.urlsafe_b64encode(msg.as_bytes()).decode()


def send(cfg, to, subject, text, html=None):
    token = access_token(cfg)
    payload = json.dumps({'raw': build_raw(to, subject, text, html)}).encode()
    status, resp = _post(SEND_URL, payload, {'Authorization': f'Bearer {token}',
                                              'Content-Type': 'application/json'})
    if status != 200:
        raise RuntimeError(f'send failed: HTTP {status}: {resp[:300]}')
    return json.loads(resp).get('id', '')


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('--subject')
    ap.add_argument('--body-file')
    ap.add_argument('--html-file')
    ap.add_argument('--to', help='override NOTIFY_TO for this send')
    ap.add_argument('--check', action='store_true', help='verify config and token; send nothing')
    ap.add_argument('--test', action='store_true', help='send one test message')
    a = ap.parse_args(argv)

    cfg = load_config()
    to = recipients(a.to or cfg.get('NOTIFY_TO', ''))
    missing = [k for k in KEYS if k != 'NOTIFY_TO' and not cfg.get(k)]
    if not to:
        missing.append('NOTIFY_TO')
    src = 'environment + scripts/.env' if ENV_FILE.exists() else 'environment (scripts/.env absent)'
    if missing:
        print(f'email skipped: not configured — missing {", ".join(missing)} in {src}')
        return 2

    try:
        if a.check:
            access_token(cfg)
            print(f'email ready: token OK, would send to {", ".join(to)}')
            return 0
        if a.test:
            subject, text, html = '[文獻收錄] 寄信測試', '這是一封測試信，代表每週收錄的通知信可以正常寄出。', None
        else:
            if not (a.subject and a.body_file):
                ap.error('--subject and --body-file are required unless --check or --test')
            subject = a.subject
            text = Path(a.body_file).read_text(encoding='utf-8')
            html = Path(a.html_file).read_text(encoding='utf-8') if a.html_file else None
        mid = send(cfg, to, subject, text, html)
        print(f'email sent to {", ".join(to)} (Gmail id {mid})')
        return 0
    except Exception as e:
        print(f'email failed: {type(e).__name__}: {e}')
        return 1


if __name__ == '__main__':
    sys.exit(main())
