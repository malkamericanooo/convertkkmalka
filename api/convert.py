from http.server import BaseHTTPRequestHandler
import json, datetime, urllib.parse, tempfile, os
import app as conv

class handler(BaseHTTPRequestHandler):
    def _send(self, code, ctype, body, filename=None):
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        if filename:
            quoted = urllib.parse.quote(filename)
            self.send_header('Content-Disposition', f"attachment; filename*=UTF-8''{quoted}")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path in ('/', '/index.html'):
            with open(os.path.join(os.path.dirname(__file__), '..', 'index.html'), 'rb') as f:
                html = f.read()
            self._send(200, 'text/html; charset=utf-8', html)
        elif self.path.startswith('/dl?'):
            try:
                qs = urllib.parse.urlparse(self.path).query
                name = os.path.basename(urllib.parse.parse_qs(qs).get('name', [''])[0])
                path = os.path.join('/tmp', 'output', name)
                if name and os.path.isfile(path):
                    with open(path, 'rb') as f:
                        data = f.read()
                    self._send(200, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', data, filename=name)
                else:
                    self._send(404, 'text/plain', b'file not found')
            except Exception:
                self._send(404, 'text/plain', b'file not found')
        else:
            self._send(404, 'text/plain', b'not found')

    def do_POST(self):
        if self.path != '/api/convert':
            self._send(404, 'text/plain', b'not found'); return
        ctype = self.headers.get('Content-Type', '')
        try:
            length = int(self.headers.get('Content-Length', 0))
        except ValueError:
            self._send(400, 'application/json', json.dumps({'ok': False, 'error': 'Content-Length tidak valid.'}).encode()); return
        if length <= 0 or length > 200 * 1024 * 1024:
            self._send(413, 'application/json', json.dumps({'ok': False, 'error': 'File terlalu besar (maksimal 200 MB).'}).encode()); return
        try:
            body = self.rfile.read(length)
            fields, files = self._parse_multipart(body, ctype)
            acuan_str = fields.get('acuan', '')
            data = files.get('pdf')
            if not data:
                raise ValueError('Tidak ada file PDF yang dipilih.')
            if not data.startswith(b'%PDF'):
                raise ValueError('File bukan PDF yang valid.')
            acuan = datetime.date.fromisoformat(acuan_str) if acuan_str else datetime.date.today()
            outdir = '/tmp/output'
            os.makedirs(outdir, exist_ok=True)
            res = conv.convert_bytes(data, acuan, outdir)
            warn = []
            qa = res['qa']
            if qa['tanpa_tanggal']:
                warn.append(f"{len(qa['tanpa_tanggal'])} baris tanpa tanggal lahir valid")
            if qa['umur_anomali']:
                warn.append(f"{len(qa['umur_anomali'])} umur anomali (<0 atau >120)")
            self._send(200, 'application/json', json.dumps({
                'ok': True, 'converted': res['converted'], 'summary': res['summary'],
                'families': res['families'], 'penduduk': qa['total_penduduk'],
                'dewasa': qa['total_dewasa'], 'ringkasan': res['ringkasan_keluarga'],
                'warn': ' · '.join(warn),
            }).encode())
        except Exception as e:
            self._send(400, 'application/json', json.dumps({'ok': False, 'error': str(e)}).encode())

    def _parse_multipart(self, body, content_type):
        boundary = None
        for part in content_type.split(';'):
            part = part.strip()
            if part.startswith('boundary='):
                boundary = part[9:].strip('"')
        if not boundary:
            return {}, {}
        boundary = ('--' + boundary).encode()
        fields, files = {}, {}
        for chunk in body.split(boundary):
            chunk = chunk.strip(b'\r\n-')
            if not chunk or chunk == b'--':
                continue
            if b'\r\n\r\n' not in chunk:
                continue
            head, payload = chunk.split(b'\r\n\r\n', 1)
            head = head.decode('utf-8', 'replace')
            name = filename = None
            for line in head.split('\r\n'):
                if 'name="' in line:
                    name = line.split('name="')[1].split('"')[0]
                if 'filename="' in line:
                    filename = line.split('filename="')[1].split('"')[0]
            if not name:
                continue
            if filename:
                files[name] = payload.rstrip(b'\r\n')
            else:
                fields[name] = payload.decode('utf-8', 'replace')
        return fields, files