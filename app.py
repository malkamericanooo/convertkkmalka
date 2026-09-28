#!/usr/bin/env python3
"""
Konverter PDF Rekapitulasi Data Keluarga (BKKBN) -> Excel
1) Replika hasil konverter berbayar: header berulang per halaman + data + RINGKASAN
2) Ringkasan per KK format "DATA RUMAH TANGGA/KK" (format DESA MABURAI)
3) Semua anggota: jumlah KK di atas, per KK anggota >= 18, < 18 dan total
"""
import re, datetime, os, sys, json
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

COL_XS = [40,110,290,421,600,674,754,794,844,1011,1061,1111,1189,1264,1339,
          1403,1465,1527,1581,1634,1688,1741,1795,1848,1910,1954]
VLINE = [{"object_type":"line","orientation":"v","top":0,"bottom":1000,
          "x0":x,"x1":x,"width":0,"height":1000} for x in COL_XS]
TBL_SETTINGS = {"vertical_strategy":"explicit","horizontal_strategy":"lines",
                "explicit_vertical_lines":VLINE,"snap_tolerance":2,
                "min_words_vertical":1,"min_words_horizontal":1}

GREEN_FILL = PatternFill('solid', fgColor='FFCCFFCC')
THIN = Side(style='thin'); DOUBLE = Side(style='double')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_TITLE = Font(name='Arial Bold', size=18)
F_HDR   = Font(name='Arial Bold', size=10)
F_LBL   = Font(name='Arial', size=12)
F_DATA  = Font(name='Arial', size=10)

HEADER_LABELS = {
    ('E',2): ('KECAMATAN',F_LBL), ('F',2): (':',F_LBL),
    ('A',3): ('RT',F_LBL), ('B',3): (':',F_LBL),
    ('E',4): ('KABUPATEN / KOTA',F_LBL), ('F',4): (':',F_LBL),
    ('A',5): ('DUSUN / RW',F_LBL), ('B',5): (':',F_LBL),
    ('E',6): ('PROVINSI',F_LBL), ('F',6): (':',F_LBL),
    ('A',7): ('DESA / KELURAHAN',F_LBL), ('B',7): (':',F_LBL),
    ('E',8): ('BULAN DAN TAHUN LAPOR',F_LBL), ('F',8): (':',F_LBL),
}
TABLE_HDR = {
    ('N',10): 'KESERTAAN BER-KB PUS', ('U',10): 'SASARAN DAN KESERTAAN DALAM POKTAN',
    ('A',11): 'NO.', ('E',11): 'HUBUNGAN', ('F',11): 'TANGAL LAHIR',
    ('H',11): 'JUMLAH', ('J',11): 'STATUS', ('K',11): 'STATUS',
    ('B',12): 'KKI', ('C',12): 'NIK', ('D',12): 'NAMA', ('G',12): 'USIA',
    ('I',12): 'KESERTAAN JKN', ('M',12): 'PESERTA KB', ('P',12): 'BUKAN PESERTA KB',
    ('S',12): 'BKB', ('U',12): 'BKR', ('W',12): 'BKL', ('Y',12): 'MUTASI',
    ('A',13): 'KELUARGA', ('E',13): 'DENGAN KK', ('H',13): 'ANAK',
    ('J',13): 'PUS', ('K',13): 'HAMIL', ('X',14): 'UPPKA',
    ('L',15): 'METODE', ('M',15): 'JALUR', ('N',15): 'LAMA BER-KB',
    ('O',15): 'Ingin Anak', ('P',15): 'Ingin Anak', ('Q',15): 'Tidak Ingin',
    ('R',15): 'Keluarga', ('T',15): 'Keluarga', ('V',15): 'Keluarga', ('X',15): '(KESERTAAN)',
    ('L',16): 'KONTRASEPSI', ('M',16): 'PELAYANAN', ('N',16): '(bulan)',
    ('O',16): 'Segera', ('P',16): 'Kemudian', ('Q',16): 'Anak Lagi',
    ('R',16): 'Sasaran', ('S',16): 'Kesertaan', ('T',16): 'Sasaran',
    ('U',16): 'Kesertaan', ('V',16): 'Sasaran', ('W',16): 'Kesertaan',
    ('O',17): '(IAS)', ('P',17): '(IAT)', ('Q',17): '(TIAL)',
}
COL_WIDTHS = {'A':27,'B':38,'C':42,'D':54,'E':26,'F':42,'H':23,'I':28,'K':8,
              'L':11,'M':10,'N':21,'P':15,'R':10,'S':9,'U':30,'W':9,'Z':8}

def clean(v):
    if v is None: return None
    s = str(v).replace('\n',' ').strip()
    return s if s else None

def parse_header(text):
    h = {}
    def grab(pat, key):
        m = re.search(pat, text)
        if m: h[key] = m.group(1).strip()
    grab(r'KECAMATAN\s*:\s*(.+)', 'kecamatan')
    m = re.search(r'^RT\s*:\s*(.+)$', text, re.MULTILINE)
    if m: h['rt'] = m.group(1).strip()
    grab(r'KABUPATEN / KOTA\s*:\s*(.+)', 'kabupaten')
    grab(r'DUSUN / RW\s*:\s*(.+)', 'dusun')
    grab(r'PROVINSI\s*:\s*(.+)', 'provinsi')
    grab(r"DESA / KELURAHAN\s*:\s*(.+)", 'desa')
    m = re.search(r'BULAN DAN TAHUN LAPOR\w*\s*:\s*(\S+)\s*-\s*(\d{4})', text)
    if m: h['bulan'], h['tahun'] = m.group(1), int(m.group(2))
    return h

def page_data(page):
    """Satu halaman -> (rows[25], greens: set index data-row yang KK)."""
    text = page.extract_text() or ''
    tables = page.find_tables(TBL_SETTINGS)
    if not tables: return [], set(), text
    t = tables[0]
    cells = t.extract()
    start = None
    for i, row in enumerate(cells):
        vals = [clean(c) for c in row]
        filled = [v for v in vals if v]
        if len(filled) >= 20 and all(filled[j] == str(j+1) for j in range(min(5, len(filled)))):
            start = i + 1; break
    if start is None: return [], set(), text
    data = cells[start:]
    # rect hijau = baris KK
    grects = []
    for r in page.rects:
        c = r.get('non_stroking_color')
        if c and isinstance(c,(tuple,list)) and len(c)==3 and 0.75<=c[0]<=0.85 and c[1]>=0.95 and 0.75<=c[2]<=0.85:
            grects.append((r['top'], r['bottom']))
    greens = set()
    for i in range(len(data)):
        b = t.rows[start + i].bbox
        yc = (b[1] + b[3]) / 2
        if any(g0-2 <= yc <= g1+2 for g0,g1 in grects):
            greens.add(i)
    return data, greens, text

def parse_ringkasan(page):
    out = {'ringkasan': {}, 'pelayanan': [], 'poktan': []}
    txt = page.extract_text() or ''
    pats = {
        'jumlah_keluarga': r'1\. Jumlah Keluarga\s*:\s*(\d+)',
        'jumlah_pus': r'2\. Jumlah PUS\s*:\s*(\d+)',
        'pus_bukan_kb': r'4\. Jumlah PUS bukan Peserta KB\s*:\s*(\d+)',
        'ias': r'1\) Ingin Anak Segera\s*:\s*(\d+)',
        'iat': r'2\) Ingin Anak Ditunda\s*:\s*(\d+)',
        'tial': r'3\) Tidak Ingin Anak Lagi\s*:\s*(\d+)',
        'hamil': r'5\. Jumlah Wanita Hamil\s*:\s*(\d+)',
        'unmet': r'6\. Jumlah Unmet Need\s*:\s*(\d+)',
        'sasaran_kel': r'7\. Keluarga Menjadi Sasaran Kelompok Kegiatan\s*:\s*([\d ]+)',
        'anggota_kel': r'8\. Keluarga Menjadi Anggota Kelompok Kegiatan\s*:\s*([\d ]+)',
    }
    for k, p in pats.items():
        m = re.search(p, txt)
        out['ringkasan'][k] = [int(x) for x in m.group(1).split()] if m else []
    for tb in page.extract_tables():
        if tb and tb[0] and str(tb[0][0] or '').strip() == 'METODE':
            out['pelayanan'] = [[clean(x) for x in row] for row in tb]
        elif tb and tb[0] and str(tb[0][0] or '').strip() == 'BKB':
            out['poktan'] = [[clean(x) for x in row] for row in tb]
    return out

def parse_pdf(pdf_path):
    doc = {'info': {}, 'pages': [], 'ringkasan': None, 'qa': []}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            is_summary = ('JUMLAH PELAYANAN' in text and 'MUTASI :' in text
                          and 'KESERTAAN BER-KB' not in text)
            if is_summary:
                doc['ringkasan'] = parse_ringkasan(page)
                continue
            rows, greens, text = page_data(page)
            rows = [[clean(v) for v in row[:25]] for row in rows]
            # buang baris kosong SEKALIGUS geser index greens agar tetap cocok
            kept, ng = [], set()
            for idx, r in enumerate(rows):
                if any(r):
                    if idx in greens: ng.add(len(kept))
                    kept.append(r)
            doc['pages'].append({'rows': kept, 'greens': ng, 'header': parse_header(text)})
        for pg in doc['pages']:
            if pg['header'].get('desa'):
                doc['info'] = pg['header']; break
    return doc

def kki_out(kki_full):
    """Aturan file berbayar: kolom B = KKI dasar (14 digit) + spasi + nomor
    keluarga 9 digit. PDF kadang menyambung token dengan spasi ganda."""
    k = ' '.join((kki_full or '').split())
    toks = k.split()
    if len(toks) >= 2:
        return toks[0] + ' ' + ''.join(toks[1:])
    if len(k) >= 24:
        return k[:14] + ' ' + k[14:]
    return k

def write_header_block(ws, row0, info):
    c = ws.cell(row=row0, column=4, value='REKAPITULASI DATA KELUARGA'); c.font = F_TITLE
    ws.row_dimensions[row0].height = 23
    for (col, r), (v, f) in HEADER_LABELS.items():
        ws[f'{col}{row0 + r - 1}'] = v; ws[f'{col}{row0 + r - 1}'].font = f
    vals = {'G2': info.get('kecamatan',''), 'C3': info.get('rt',''),
            'G4': info.get('kabupaten',''), 'C5': info.get('dusun',''),
            'H6': info.get('provinsi',''), 'C7': info.get('desa',''),
            'G8': info.get('bulan',''), 'H8': '-', 'I8': info.get('tahun','')}
    for coord, v in vals.items():
        cell = ws[f'{coord[0]}{int(coord[1:]) + row0 - 1}']
        cell.value = v; cell.font = F_LBL
    for (col, r), v in TABLE_HDR.items():
        cell = ws[f'{col}{row0 + r - 1}']; cell.value = v; cell.font = F_HDR
    for j in range(1, 26):
        cell = ws.cell(row=row0+17, column=j, value=j)
        cell.font = F_DATA; cell.border = BOX

def cast_row(row):
    out = list(row) + [None]*(25 - len(row))
    for idx in (6, 7):
        if out[idx] is not None and re.fullmatch(r'\d+', str(out[idx])):
            out[idx] = int(out[idx])
    return out[:25]

def build_converted(doc, out):
    """Build converted xlsx. out can be a path string or file-like object (BytesIO)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    for c, w in COL_WIDTHS.items(): ws.column_dimensions[c].width = w
    r0 = 1
    families = []   # untuk summary
    cur = None
    for pg in doc['pages']:
        write_header_block(ws, r0, pg['header'] if pg['header'].get('desa') else doc['info'])
        data0 = r0 + 18
        kk_full = None
        kk_no = None
        for i, row in enumerate(pg['rows']):
            rr = data0 + i
            vals = cast_row(row)
            is_green = i in pg['greens']
            hub = str(vals[4] or '').strip().upper()
            if is_green or hub == 'KK':
                if cur: families.append(cur)
                kk_full = vals[1] or kk_full
                kk_no = vals[0] or kk_no
                cur = {'nama': vals[3], 'lahir': vals[5], 'kk_no': vals[0],
                       'members': [(vals[4], vals[5])]}
            elif cur is not None:
                cur['members'].append((vals[4], vals[5]))
            # tulis sel
            for j, v in enumerate(vals, start=1):
                cell = ws.cell(row=rr, column=j)
                cell.value = v; cell.font = F_DATA; cell.border = BOX
                if is_green: cell.fill = GREEN_FILL
            # KKI kolom A/B sesuai aturan konverter berbayar
            if is_green:
                ws.cell(row=rr, column=1).value = kk_no
                ws.cell(row=rr, column=2).value = kki_out(kk_full)
            elif cur is not None:
                ws.cell(row=rr, column=1).value = None
                ws.cell(row=rr, column=2).value = kki_out(kk_full)
        r0 = data0 + len(pg['rows'])
    if cur: families.append(cur)
    # RINGKASAN
    if doc['ringkasan']:
        rk = doc['ringkasan']['ringkasan']
        ws.cell(row=r0, column=1, value='MUTASI :').font = F_DATA
        ws.cell(row=r0, column=2, value='1. MENIKAH').font = F_DATA
        ws.cell(row=r0, column=3, value='2. BERCERAI  3. MENINGGAL DUNIA').font = F_DATA
        ws.cell(row=r0, column=6, value='4. PINDAH  5. ANGGOTA KELUARGA BARU').font = F_DATA
        r0 += 1
        ws.cell(row=r0, column=1, value='RINGKASAN :').font = F_DATA
        r0 += 1
        def lbl(rr, cc, label):
            ws.cell(row=rr, column=cc, value=label).font = F_DATA
        def kv(rr, cc, label, key, idx=0):
            lbl(rr, cc, label); ws.cell(row=rr, column=cc+4, value=':').font = F_DATA
            v = rk.get(key, [])
            if idx < len(v): ws.cell(row=rr, column=cc+5, value=v[idx]).font = F_DATA
        kv(r0, 1, '1. Jumlah Keluarga', 'jumlah_keluarga'); kv(r0, 7, '4. Jumlah PUS bukan Peserta KB', 'pus_bukan_kb'); r0 += 1
        kv(r0, 1, '2. Jumlah PUS', 'jumlah_pus'); kv(r0, 7, '1) Ingin Anak Segera', 'ias'); r0 += 1
        lbl(r0, 2, '3. Jumlah Peserta KB Aktif'); kv(r0, 7, '2) Ingin Anak Ditunda', 'iat'); r0 += 1
        kv(r0, 7, '3) Tidak Ingin Anak Lagi', 'tial'); r0 += 1
        kv(r0, 7, '5. Jumlah Wanita Hamil', 'hamil'); r0 += 1
        for prow in doc['ringkasan']['pelayanan']:
            for j, v in enumerate(prow[:5]):
                ws.cell(row=r0, column=j+1, value=v).font = F_DATA
            r0 += 1
        kv(r0, 7, '6. Jumlah Unmet Need', 'unmet'); r0 += 1
        kv(r0, 7, '7. Keluarga Menjadi Sasaran Kelompok Kegiatan', 'sasaran_kel'); r0 += 1
        kv(r0, 7, '8. Keluarga Menjadi Anggota Kelompok Kegiatan', 'anggota_kel'); r0 += 1
        for prow in doc['ringkasan']['poktan']:
            for j, v in enumerate(prow[:4]):
                ws.cell(row=r0, column=9+j, value=v).font = F_DATA
            r0 += 1
    wb.save(out)
    return families

# ---------- tanggal / umur ----------
def parse_tanggal(s):
    if s is None: return None
    s = str(s).strip().replace('/', '-')
    p = s.split('-')
    if len(p) != 3: return None
    try: d, m, y = int(p[0]), int(p[1]), int(p[2])
    except ValueError: return None
    if y < 100: y += 1900 if y >= 30 else 2000
    try: return datetime.date(y, m, d)
    except ValueError: return None

def umur_pada(lahir, acuan):
    return acuan.year - lahir.year - ((acuan.month, acuan.day) < (lahir.month, lahir.day))

def hitung_anggota(fam, acuan, threshold=18):
    """Satu keluarga -> jumlah anggota >= threshold (dewasa), < threshold (anak),
    total anggota, plus baris yang tanggal lahirnya tak terbaca / umurnya aneh.
    Anggota tanpa umur valid masuk total tapi tidak ke dewasa maupun anak."""
    h = {'dewasa': 0, 'anak': 0, 'anggota': 0, 'tanpa_tanggal': [], 'umur_anomali': []}
    for hub, lahir in fam['members']:
        h['anggota'] += 1
        b = parse_tanggal(lahir)
        if b is None:
            if lahir not in (None, ''):
                h['tanpa_tanggal'].append([fam['nama'], lahir])
            continue
        u = umur_pada(b, acuan)
        if u < 0 or u > 120:
            h['umur_anomali'].append([fam['nama'], lahir, u]); continue
        if u >= threshold: h['dewasa'] += 1
        else: h['anak'] += 1
    return h

def ttl(x):
    return ' '.join(w[0].upper() + w[1:].lower() if w else w for w in (x or '').split())

def rt_pendek(info):
    rt = (info.get('rt') or '').strip()
    return re.sub(r'RT\s*0+', 'RT ', rt) if rt else ''

def build_summary(families, doc, acuan, out, threshold=18):
    """Build summary xlsx. out can be a path string or file-like object (BytesIO)."""
    info = doc['info']
    desa = ttl(info.get('desa'))
    kec  = ttl(info.get('kecamatan'))
    kab  = ttl(info.get('kabupaten'))
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sheet1'
    for c, w in {'A':12,'B':42,'C':24,'D':14,'E':20}.items():
        ws.column_dimensions[c].width = w
    ws['A1'] = 'DATA RUMAH TANGGA/KK DI KELURAHAN/DESA TERPILIH'
    ws['A2'] = 'KEGIATAN SURVEY DAMPAK POPM FILARIASIS (BIS)'
    ws['A3'] = f"BBLKM BANJARBARU TAHUN {acuan.year}"
    for coord in ('A1','A2','A3'): ws[coord].font = F_HDR
    head = [(5,'Nama Desa', f': {desa}'), (6,'Nama Kecamatan', f': {kec}'),
            (7,'Jumlah RT', ': 1  RT'), (8,'Jumlah KK', ':'),
            (9,'Jumlah Penduduk', ':'), (10,'Puskesmas', f': {desa}'),
            (11,'Kabupaten', f': {kab}')]
    for rr, lbl, v in head:
        ws[f'A{rr}'] = lbl; ws[f'C{rr}'] = v
        ws[f'A{rr}'].font = F_DATA; ws[f'C{rr}'].font = F_DATA
    headers = ['No. Urut', 'Nama KK', f'Jumlah Anggota Keluarga ≥ {threshold} Tahun',
               'Alamat (RT)', 'Keterangan']
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(row=13, column=j, value=h)
        cell.font = F_HDR; cell.border = BOX; cell.fill = GREEN_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    rt_short = rt_pendek(info)
    qa = {'tanpa_tanggal': [], 'umur_anomali': [], 'total_dewasa': 0, 'total_anak': 0,
          'total_penduduk': 0}
    n = 0
    for fam in families:
        n += 1
        h = hitung_anggota(fam, acuan, threshold)
        qa['tanpa_tanggal'] += h['tanpa_tanggal']; qa['umur_anomali'] += h['umur_anomali']
        qa['total_dewasa'] += h['dewasa']; qa['total_anak'] += h['anak']
        qa['total_penduduk'] += h['anggota']
        for j, v in enumerate([n, fam['nama'], h['dewasa'], rt_short, ''], start=1):
            cell = ws.cell(row=13+n, column=j, value=v)
            cell.font = F_DATA; cell.border = BOX
    ws['C8'] = f': {n}  KK'
    ws['C9'] = f': {qa["total_penduduk"]}'
    wb.save(out)
    return qa

def build_semua_anggota(families, doc, acuan, out, threshold=18):
    """File ke-3: jumlah KK + total anggota di paling atas, lalu per KK jumlah
    anggota >= threshold, < threshold dan total. out bisa path atau BytesIO."""
    info = doc['info']
    rt = rt_pendek(info)
    rows, tot = [], {'dewasa': 0, 'anak': 0, 'anggota': 0}
    for fam in families:
        h = hitung_anggota(fam, acuan, threshold)
        for k in tot: tot[k] += h[k]
        lain = h['anggota'] - h['dewasa'] - h['anak']
        ket = f'{lain} anggota tgl lahir kosong/tidak valid' if lain else ''
        rows.append([fam['nama'], h['dewasa'], h['anak'], h['anggota'], rt, ket])

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sheet1'
    for c, w in {'A':10,'B':42,'C':18,'D':18,'E':16,'F':12,'G':36}.items():
        ws.column_dimensions[c].width = w
    ws['A1'] = 'REKAP KEPALA KELUARGA DAN ANGGOTA KELUARGA'
    lokasi = [f'Desa {ttl(info["desa"])}' if info.get('desa') else '', rt,
              f'Kec. {ttl(info["kecamatan"])}' if info.get('kecamatan') else '',
              f'Kab. {ttl(info["kabupaten"])}' if info.get('kabupaten') else '']
    ws['A2'] = ', '.join(x for x in lokasi if x)
    ws['A3'] = f'Umur dihitung per tanggal {acuan:%d-%m-%Y}'
    ws['A1'].font = F_HDR; ws['A2'].font = F_DATA; ws['A3'].font = F_DATA

    rekap = [('Jumlah KK (Kepala Keluarga)', len(families)),
             (f'Jumlah Anggota Keluarga ≥ {threshold} Tahun', tot['dewasa']),
             (f'Jumlah Anggota Keluarga < {threshold} Tahun', tot['anak']),
             ('Jumlah Penduduk (semua anggota)', tot['anggota'])]
    lain = tot['anggota'] - tot['dewasa'] - tot['anak']
    if lain: rekap.append(('Tgl lahir kosong/tidak valid', lain))
    r = 5
    for lbl, v in rekap:
        ws.cell(row=r, column=1, value=lbl).font = F_HDR
        cell = ws.cell(row=r, column=3, value=v)
        cell.font = F_HDR; cell.alignment = Alignment(horizontal='left')
        r += 1

    hdr = r + 1
    headers = ['No. Urut', 'Nama KK', f'Jumlah Anggota Keluarga ≥ {threshold} Tahun',
               f'Jumlah Anggota Keluarga < {threshold} Tahun', 'Total Anggota Keluarga',
               'Alamat (RT)', 'Keterangan']
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(row=hdr, column=j, value=h)
        cell.font = F_HDR; cell.border = BOX; cell.fill = GREEN_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[hdr].height = 42
    for n, row in enumerate(rows, start=1):
        for j, v in enumerate([n] + row, start=1):
            cell = ws.cell(row=hdr + n, column=j, value=v)
            cell.font = F_DATA; cell.border = BOX
    ws.freeze_panes = f'A{hdr + 1}'
    wb.save(out)
    return tot

def convert(pdf_path, acuan, outdir):
    doc = parse_pdf(pdf_path)
    desa = (doc['info'].get('desa') or 'DESA').replace("'", '').replace(' ', '_')
    rt = (doc['info'].get('rt') or '').replace(' ', '')
    os.makedirs(outdir, exist_ok=True)
    conv = os.path.join(outdir, f'REKAP {desa} {rt} converted.xlsx')
    summ = os.path.join(outdir, f'DESA {desa}.xlsx')
    semua = os.path.join(outdir, f'DESA {desa} SEMUA ANGGOTA.xlsx')
    fams = build_converted(doc, conv)
    qa = build_summary(fams, doc, acuan, summ)
    build_semua_anggota(fams, doc, acuan, semua)
    ring = (doc['ringkasan'] or {}).get('ringkasan', {}).get('jumlah_keluarga', [None])[0]
    return {'converted': os.path.basename(conv), 'summary': os.path.basename(summ),
            'semua': os.path.basename(semua),
            'families': len(fams), 'ringkasan_keluarga': ring, 'qa': qa}

def convert_bytes(pdf_bytes, acuan, outdir):
    """Konversi dari bytes (untuk web upload)."""
    import tempfile
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tf:
        tf.write(pdf_bytes)
        tmp = tf.name
    try:
        return convert(tmp, acuan, outdir)
    finally:
        os.unlink(tmp)

if __name__ == '__main__':
    pdf_path = sys.argv[1]
    acuan = datetime.date.fromisoformat(sys.argv[2]) if len(sys.argv) > 2 else datetime.date.today()
    outdir = sys.argv[3] if len(sys.argv) > 3 else os.path.dirname(pdf_path) or '.'
    res = convert(pdf_path, acuan, outdir)
    res['converted'] = os.path.join(outdir, res['converted'])
    res['summary'] = os.path.join(outdir, res['summary'])
    print(json.dumps(res, indent=2, default=str))
