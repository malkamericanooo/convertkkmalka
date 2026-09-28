#!/usr/bin/env python3
"""Web app lokal: PDF Rekapitulasi -> file Excel pilihan (REKAP, form POPM, kategori umur).
Jalankan:  python3 server.py   lalu buka  http://localhost:8787
"""
import os, datetime, json
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
import urllib.parse
import app as conv


def parse_multipart(body, content_type):
    """Pengganti cgi.FieldStorage: ambil field file 'pdf' + field 'acuan'."""
    m = {}
    for part in content_type.split(';'):
        part = part.strip()
        if part.startswith('boundary='):
            m['boundary'] = part[9:].strip('"')
    if 'boundary' not in m:
        return {}, {}
    boundary = ('--' + m['boundary']).encode()
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

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, 'output')
PORT = int(os.environ.get('PORT', 8787))
MAX_UPLOAD = 200 * 1024 * 1024  # 200 MB

PAGE = '''<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Konverter Rekapitulasi KK</title>
<style>
:root{--bg:#0f172a;--card:#1e293b;--accent:#10b981;--accent2:#34d399;
      --txt:#e2e8f0;--mut:#94a3b8;--err:#f87171;}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
     background:var(--bg);color:var(--txt);min-height:100vh;
     display:flex;align-items:center;justify-content:center;padding:24px}
.card{background:var(--card);border-radius:16px;padding:32px;max-width:520px;
      width:100%;box-shadow:0 20px 60px rgba(0,0,0,.4)}
h1{font-size:1.3rem;margin-bottom:4px}
.sub{color:var(--mut);font-size:.9rem;margin-bottom:24px}
label{display:block;font-size:.85rem;color:var(--mut);margin:16px 0 6px}
input[type=date]{width:100%;padding:10px;border-radius:8px;border:1px solid #334155;
     background:#0f172a;color:var(--txt);font-size:1rem}
.drop{border:2px dashed #334155;border-radius:12px;padding:28px;text-align:center;
     cursor:pointer;transition:border-color .2s, background .2s;margin-top:6px}
.drop:hover,.drop.over{border-color:var(--accent);background:rgba(16,185,129,.06)}
.drop.has{border-color:var(--accent2);background:rgba(16,185,129,.08)}
.drop .fname{font-weight:600;color:var(--accent2);margin-top:8px;word-break:break-all}
input[type=file]{display:none}
button{width:100%;margin-top:24px;padding:14px;border:none;border-radius:10px;
     background:var(--accent);color:#04150e;font-size:1rem;font-weight:700;
     cursor:pointer;transition:background .2s}
button:hover{background:var(--accent2)}
button:disabled{background:#334155;color:var(--mut);cursor:not-allowed}
.status{margin-top:20px;font-size:.9rem;color:var(--mut);min-height:1.4em}
.status.err{color:var(--err)}
.result{margin-top:16px;display:none}
.result a{display:block;padding:12px 14px;background:#0f172a;border:1px solid #334155;
     border-radius:10px;color:var(--accent2);text-decoration:none;margin-top:8px;
     font-size:.9rem}
.result a:hover{border-color:var(--accent)}
.kat{display:flex;gap:10px;align-items:center;margin:6px 0;color:var(--txt);font-size:.9rem;cursor:pointer}
.kat small{color:var(--mut)}
.qa{margin-top:14px;font-size:.85rem;color:var(--mut);background:#0f172a;
     border-radius:10px;padding:12px 14px;line-height:1.7}
.hint{margin-top:18px;font-size:.78rem;color:var(--mut);line-height:1.6}
</style>
</head>
<body>
<div class="card">
  <h1>📋 Konverter Rekapitulasi Data Keluarga</h1>
  <div class="sub">PDF ekspor pemerintah → Excel, pilih file yang dibutuhkan saja</div>

  <form id="f" method="post" action="/convert" enctype="multipart/form-data">
    <label>1. Pilih tanggal acuan (untuk hitung umur)</label>
    <input type="date" name="acuan" id="acuan">

    <label>2. File yang dibuat</label>
    <div id="out-list">
      <label class="kat"><input type="checkbox" id="out-rekap" value="rekap" checked> REKAP lengkap <small>salinan tabel PDF (file paling besar)</small></label>
      <label class="kat"><input type="checkbox" id="out-popm" value="popm" checked> Form POPM: ringkasan KK <small>anggota ≥ 18 tahun</small></label>
      <label class="kat"><input type="checkbox" id="out-terpisah" value="terpisah" checked> Per kategori umur: file terpisah</label>
      <label class="kat"><input type="checkbox" id="out-gabungan" value="gabungan" checked> Per kategori umur: file gabungan</label>
    </div>

    <label>3. Kategori umur (untuk file per kategori, bisa lebih dari satu)</label>
    <div id="kat-list">
      <label class="kat"><input type="checkbox" id="kat-balita" value="balita"> Bayi Balita <small>0–59 bulan</small></label>
      <label class="kat"><input type="checkbox" id="kat-prasekolah" value="prasekolah"> Anak Pra Sekolah <small>5–6 tahun</small></label>
      <label class="kat"><input type="checkbox" id="kat-sekolah" value="sekolah"> Anak Usia Sekolah &amp; Remaja <small>7 tahun – 17 tahun 11 bulan 29 hari</small></label>
      <label class="kat"><input type="checkbox" id="kat-dewasa" value="dewasa"> Dewasa / Produktif <small>18–59 tahun</small></label>
      <label class="kat"><input type="checkbox" id="kat-lansia" value="lansia"> Lansia <small>60 tahun ke atas</small></label>
    </div>

    <label>4. Taruh / pilih file PDF</label>
    <div class="drop" id="drop">
      <div>Klik atau seret file PDF ke sini</div>
      <div class="fname" id="fname"></div>
    </div>
    <input type="file" name="pdf" id="pdf" accept="application/pdf">

    <button type="submit" id="go" disabled>Konversi</button>
  </form>

  <div class="status" id="status"></div>
  <div class="result" id="result">
    <div id="lfiles"></div>
    <div class="qa" id="qa"></div>
  </div>
  <div class="hint">
    Baris <b style="color:#86efac">hijau</b> = Kepala Keluarga, baris putih di
    bawahnya = anggota keluarga. Umur dihitung dari TANGGAL LAHIR (teks dd-mm-yyyy)
    pada tanggal acuan; anggota berumur ≥ 18 tahun dihitung ke kolom
    "Jumlah Anggota Keluarga". File kategori umur berisi jumlah KK di paling
    atas, lalu per KK jumlah anggota tiap kategori yang dipilih.
  </div>
</div>
<script>
const drop=document.getElementById('drop'),pdf=document.getElementById('pdf'),
      fname=document.getElementById('fname'),go=document.getElementById('go'),
      acuan=document.getElementById('acuan'),status=document.getElementById('status'),
      form=document.getElementById('f');
acuan.value=new Date().toISOString().slice(0,10);
drop.onclick=()=>pdf.click();
drop.ondragover=e=>{e.preventDefault();drop.classList.add('over')};
drop.ondragleave=()=>drop.classList.remove('over');
drop.ondrop=e=>{e.preventDefault();drop.classList.remove('over');
  if(e.dataTransfer.files[0]){pdf.files=e.dataTransfer.files;show()}};
pdf.onchange=show;
function show(){
  const f=pdf.files[0];
  if(!f)return;
  if(!f.name.toLowerCase().endsWith('.pdf')){status.textContent='Harus file PDF';status.className='status err';return}
  fname.textContent=f.name;drop.classList.add('has');go.disabled=false;
  status.textContent='';status.className='status';
}
form.onsubmit=async e=>{
  e.preventDefault();
  if(!pdf.files[0])return;
  const pilih=l=>[...document.querySelectorAll(l+' input:checked')].map(c=>c.value);
  const file=pilih('#out-list'),kategori=pilih('#kat-list');
  if(!file.includes('rekap')&&!file.includes('popm')&&!(kategori.length&&(file.includes('terpisah')||file.includes('gabungan')))){
    status.textContent='Belum ada file yang dipilih. Centang minimal satu file, atau pilih kategori umur untuk file per kategori.';
    status.className='status err';return}
  go.disabled=true;status.textContent='Mengonversi… (bisa butuh beberapa detik)';
  document.getElementById('result').style.display='none';
  const fd=new FormData(form);
  fd.set('file',file.join(','));fd.set('kategori',kategori.join(','));
  try{
    const r=await fetch('/convert',{method:'POST',body:fd});
    const j=await r.json();
    if(!j.ok){status.textContent=j.error;status.className='status err';go.disabled=false;return}
    status.textContent='Selesai ✅';status.className='status';
    const res=document.getElementById('result');res.style.display='block';
    const lfiles=document.getElementById('lfiles');lfiles.replaceChildren();
    for(const f of j.files){const a=document.createElement('a');
      a.href='/dl?name='+encodeURIComponent(f.name);a.textContent='⬇️ '+f.label;lfiles.appendChild(a)}
    document.getElementById('qa').innerHTML=
      'Keluarga: <b>'+j.families+'</b> · Anggota: <b>'+j.penduduk+'</b> · '+
      'Dewasa ≥18: <b>'+j.dewasa+'</b> · Anak &lt;18: <b>'+j.anak+'</b>'+
      j.kategori_tot.map(k=>' · '+k.label+': <b>'+k.jumlah+'</b>').join('')+
      (j.ringkasan? ' · RINGKASAN PDF: '+j.ringkasan+' KK':'')+
      (j.warn? '<br><span style="color:#f87171">'+j.warn+'</span>':'');
    go.disabled=false;
  }catch(err){status.textContent='Gagal: '+err;status.className='status err';go.disabled=false}
};
</script>
</body>
</html>'''

class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def _send(self, code, ctype, body, filename=None):
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        if filename:
            quoted = urllib.parse.quote(filename)
            self.send_header('Content-Disposition',
                             f"attachment; filename*=UTF-8''{quoted}")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        if u.path in ('/', '/index.html'):
            self._send(200, 'text/html; charset=utf-8', PAGE.encode())
        elif u.path == '/dl':
            try:
                name = os.path.basename(urllib.parse.parse_qs(u.query).get('name', [''])[0])
                path = os.path.join(OUT, name)
                if name and os.path.isfile(path):
                    with open(path, 'rb') as f:
                        data = f.read()
                    self._send(200, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               data, filename=name)
                else:
                    self._send(404, 'text/plain', b'file not found')
            except Exception:
                self._send(404, 'text/plain', b'file not found')
        else:
            self._send(404, 'text/plain', b'not found')

    def do_POST(self):
        if urllib.parse.urlparse(self.path).path != '/convert':
            self._send(404, 'text/plain', b'not found'); return
        ctype = self.headers.get('Content-Type', '')
        try:
            length = int(self.headers.get('Content-Length', 0))
        except ValueError:
            self._send(400, 'application/json',
                       json.dumps({'ok': False, 'error': 'Content-Length tidak valid.'}).encode())
            return
        if length <= 0 or length > MAX_UPLOAD:
            self._send(413, 'application/json',
                       json.dumps({'ok': False,
                                   'error': 'File terlalu besar (maksimal 200 MB).'}).encode())
            return
        try:
            body = self.rfile.read(length)
            fields, files = parse_multipart(body, ctype)
            acuan_str = fields.get('acuan', '')
            data = files.get('pdf')
            if not data:
                raise ValueError('Tidak ada file PDF yang dipilih.')
            if not data.startswith(b'%PDF'):
                raise ValueError('File bukan PDF yang valid.')
            acuan = datetime.date.fromisoformat(acuan_str) if acuan_str else datetime.date.today()
            kategori = conv.parse_kategori(fields.get('kategori'))
            pilihan = conv.parse_pilihan_file(fields.get('file'))
            res = conv.convert_bytes(data, acuan, OUT, kategori, pilihan)
            warn = []
            qa = res['qa']
            if qa['tanpa_tanggal']:
                warn.append(f"{len(qa['tanpa_tanggal'])} baris tanpa tanggal lahir valid")
            if qa['umur_anomali']:
                warn.append(f"{len(qa['umur_anomali'])} umur anomali (<0 atau >120)")
            self._send(200, 'application/json', json.dumps({
                'ok': True, 'files': res['files'], 'kategori_tot': res['kategori_tot'],
                'families': res['families'], 'penduduk': qa['total_penduduk'],
                'dewasa': qa['total_dewasa'], 'anak': qa['total_anak'],
                'ringkasan': res['ringkasan_keluarga'],
                'warn': ' · '.join(warn),
            }).encode())
        except Exception as e:
            self._send(400, 'application/json', json.dumps({'ok': False, 'error': str(e)}).encode())

if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)
    print(f'Buka http://localhost:{PORT} di browser')
    ThreadingHTTPServer(('127.0.0.1', PORT), H).serve_forever()
