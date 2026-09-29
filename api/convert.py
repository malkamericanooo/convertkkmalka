from http.server import BaseHTTPRequestHandler
import json, datetime, urllib.parse, os, io, base64
import app as converter


class handler(BaseHTTPRequestHandler):
    """Vercel serverless handler.
    POST /api/convert  -> JSON with base64-encoded xlsx files
    GET  /            -> serves index.html
    """

    def _send(self, code, ctype, body, filename=None):
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        if filename:
            quoted = urllib.parse.quote(filename)
            self.send_header('Content-Disposition', f"attachment; filename*=UTF-8''{quoted}")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, code, obj):
        self._send(code, 'application/json', json.dumps(obj).encode('utf-8'))

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
            if b'\r\n\r\n' in chunk:
                head, payload = chunk.split(b'\r\n\r\n', 1)
            else:  # field kosong (mis. semua centang dimatikan): tetap dicatat sebagai ''
                head, payload = chunk, b''
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

    def do_GET(self):
        halaman = {'/': 'index.html', '/index.html': 'index.html', '/kebijakan': 'kebijakan.html'}
        if self.path in halaman:
            html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), halaman[self.path])
            with open(html_path, 'rb') as f:
                html = f.read()
            self._send(200, 'text/html; charset=utf-8', html)
        else:
            self._send(404, 'text/plain', b'not found')

    def do_POST(self):
        if self.path != '/api/convert':
            self._send(404, 'text/plain', b'not found'); return

        ctype = self.headers.get('Content-Type', '')
        try:
            length = int(self.headers.get('Content-Length', 0))
        except ValueError:
            self._send_json(400, {'ok': False, 'error': 'Content-Length tidak valid.'}); return
        if length <= 0 or length > 200 * 1024 * 1024:
            self._send_json(413, {'ok': False, 'error': 'File terlalu besar (maksimal 200 MB).'}); return

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
            kategori = converter.parse_kategori(fields.get('kategori'))
            pilihan = converter.parse_pilihan_file(fields.get('file'))
            format2 = converter.parse_format2(fields.get('format2'))

            # Parse PDF + build only the chosen xlsx files, all in memory (tidak ada yang disimpan)
            res = converter.konversi(io.BytesIO(data), acuan, kategori, pilihan, format2)
            qa = res['qa']

            # Warnings
            warn = []
            if qa['tanpa_tanggal']:
                warn.append(f"{len(qa['tanpa_tanggal'])} baris tanpa tanggal lahir valid")
            if qa['umur_anomali']:
                warn.append(f"{len(qa['umur_anomali'])} umur anomali (<0 atau >120)")

            self._send_json(200, {
                'ok': True,
                'families': res['families'],
                'penduduk': qa['total_penduduk'],
                'dewasa': qa['total_dewasa'],
                'anak': qa['total_anak'],
                'ringkasan': res['ringkasan_keluarga'],
                'warn': ' · '.join(warn),
                # base64 so the page can offer each file as a download blob
                'files': [dict(f, data=base64.b64encode(f['data']).decode('utf-8'))
                          for f in res['files']],
                'kategori_tot': res['kategori_tot'],
                'format2': res['format2'],
            })
        except Exception as e:
            self._send_json(400, {'ok': False, 'error': str(e)})
