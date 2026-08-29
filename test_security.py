#!/usr/bin/env python3
"""Manual security test suite — simulates common attacks against the local converter."""
import json, urllib.request, urllib.error, socket, urllib.parse

BASE = 'http://localhost:8787'
results = []

def check(name, cond, detail=''):
    status = 'PASS' if cond else 'FAIL'
    results.append((status, name, detail))
    print(f'[{status}] {name}' + (f'  -- {detail}' if detail else ''))

def get(path, raw=False):
    try:
        req = urllib.request.Request(BASE + path)
        with urllib.request.urlopen(req) as r:
            body = r.read()
            return r.status, body
    except urllib.error.HTTPError as e:
        return e.code, e.read()

def post_multipart(path, fields, files, content_length=None):
    boundary = '----testboundary123'
    body = b''
    for k, v in fields.items():
        body += f'--{boundary}\r\nContent-Disposition: form-data; name="{k}"\r\n\r\n{v}\r\n'.encode()
    for k, (fname, data) in files.items():
        body += f'--{boundary}\r\nContent-Disposition: form-data; name="{k}"; filename="{fname}"\r\nContent-Type: application/octet-stream\r\n\r\n'.encode()
        body += data + b'\r\n'
    body += f'--{boundary}--\r\n'.encode()
    cl = content_length if content_length is not None else len(body)
    req = urllib.request.Request(BASE + path, data=body, method='POST')
    req.add_header('Content-Type', f'multipart/form-data; boundary={boundary}')
    req.add_header('Content-Length', str(cl))
    try:
        with urllib.request.urlopen(req) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()

print('=== ATTACK 1: Path traversal on /dl ===')
for payload in ['../../etc/passwd', '..%2F..%2Fetc%2Fpasswd', '%2e%2e%2f%2e%2e%2fetc/passwd',
                '/etc/passwd', '....//....//etc/passwd', 'DESA MABUUN.xlsx/../../../etc/passwd']:
    code, body = get('/dl?name=' + urllib.parse.quote(payload))
    leaked = b'root:' in body
    check(f'traversal blocked: {payload[:40]}', code == 404 and not leaked, f'HTTP {code}')

print()
print('=== ATTACK 2: Reading app source / sensitive files ===')
for path in ['/server.py', '/app.py', '/../server.py', '/.env', '/output/../server.py']:
    code, body = get(path)
    check(f'no source leak: {path}', code == 404, f'HTTP {code}')

print()
print('=== ATTACK 3: Non-PDF uploads ===')
code, body = post_multipart('/convert', {'acuan': '2026-08-29'},
                              {'pdf': ('evil.pdf', b'this is not a pdf')})
check('plain text rejected', code == 400 and b'bukan PDF' in body, f'HTTP {code}')
# PDF header but garbage body
code, body = post_multipart('/convert', {'acuan': '2026-08-29'},
                              {'pdf': ('fake.pdf', b'%PDF-1.4 garbage garbage')})
check('header-only fake PDF handled', code in (400, 200), f'HTTP {code}')

print()
print('=== ATTACK 4: Missing/empty fields ===')
code, body = post_multipart('/convert', {}, {})
check('no file at all -> clean error', code == 400, f'HTTP {code}')
code, body = post_multipart('/convert', {'acuan': '2026-08-29'}, {'pdf': ('a.pdf', b'')})
check('empty pdf field -> clean error', code == 400, f'HTTP {code}')

print()
print('=== ATTACK 5: Bad date injection ===')
code, body = post_multipart('/convert', {'acuan': '2026-13-45'},
                              {'pdf': ('a.pdf', b'%PDF-notreal')})
check('invalid date rejected', code == 400, f'HTTP {code}')
code, body = post_multipart('/convert', {'acuan': "'; DROP TABLE x;--"},
                              {'pdf': ('a.pdf', b'%PDF-notreal')})
check('sql-ish date rejected', code == 400, f'HTTP {code}')

print()
print('=== ATTACK 6: Content-Length lies ===')
code, body = post_multipart('/convert', {'acuan': '2026-08-29'},
                              {'pdf': ('a.pdf', b'%PDF-realish')}, content_length=999999999)
check('huge claimed length handled', code in (400, 413), f'HTTP {code}')
code, body = post_multipart('/convert', {'acuan': '2026-08-29'},
                              {'pdf': ('a.pdf', b'%PDF-realish')}, content_length=-5)
check('negative length rejected', code in (400, 413), f'HTTP {code}')

print()
print('=== ATTACK 7: Oversized upload ===')
req = urllib.request.Request(BASE + '/convert', data=b'x', method='POST')
req.add_header('Content-Length', str(300 * 1024 * 1024))
req.add_header('Content-Type', 'multipart/form-data; boundary=--x')
try:
    with urllib.request.urlopen(req, timeout=5) as r:
        code = r.status
except urllib.error.HTTPError as e:
    code = e.code
except Exception as e:
    code = str(e)
check('oversize upload rejected 413', code == 413, f'HTTP {code}')

print()
print('=== ATTACK 8: Header injection via filename ===')
code, body = post_multipart('/convert', {'acuan': '2026-08-29'},
                              {'pdf': ('a.pdf\r\nX-Injected: 1', b'%PDF-x')})
check('CRLF filename handled', code in (400, 200), f'HTTP {code}')

print()
print('=== ATTACK 9: XSS in filename reflected? ===')
code, body = get('/dl?name=%3Cscript%3Ealert(1)%3C/script%3E')
check('xss filename not reflected/404', code == 404, f'HTTP {code}')

print()
print('=== ATTACK 10: server binds loopback only ===')
import subprocess
out = subprocess.run(['lsof', '-nP', '-iTCP:8787', '-sTCP:LISTEN'], capture_output=True, text=True).stdout
loopback_only = '127.0.0.1:8787' in out and '0.0.0.0:8787' not in out and '*:8787' not in out
check('bound to 127.0.0.1 only (not network-exposed)', loopback_only, out.strip().splitlines()[-1] if out.strip() else 'no listener')

print()
fails = [r for r in results if r[0] == 'FAIL']
print(f'TOTAL: {len(results)} checks, {len(fails)} failed')
for f in fails:
    print('  FAIL:', f[1], f[2])