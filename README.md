# Konverter Rekapitulasi Data Keluarga

PDF ekspor pemerintah (Rekapitulasi Data Keluarga / BKKBN) → **2 file Excel**, tanpa bayar konverter online.

## Cara pakai

### Lokal (Python)
```bash
cd ~/mabuun-konverter
python3 server.py
# buka http://localhost:8787
```

### Vercel (Deploy)
1. Push ke GitHub
2. Import di Vercel → auto-detect Python
3. Done

## Struktur

| File | Fungsi |
|------|--------|
| `app.py` | Inti konverter (parse PDF → 2 xlsx) |
| `api/convert.py` | Vercel serverless handler (POST /api/convert, GET /api/dl) |
| `api/index.html` | UI (static) |
| `vercel.json` | Routing Vercel |
| `requirements.txt` | Dependencies |
| `test_e2e.py` | Playwright E2E test |
| `test_security.py` | 23 security checks |
| `Jalankan Konverter.command` | Double-click launcher (lokal) |
| `output/` | Hasil konversi |

## Logika ringkasan KK

- Baris **hijau** di PDF = Kepala Keluarga (mulai keluarga baru)
- Baris putih di bawahnya = anggota keluarganya, sampai hijau berikutnya
- TANGGAL LAHIR itu **teks** `dd-mm-yyyy` → diurai eksplisit hari-dulu (tidak ikut setelan regional komputer)
- Umur = pada **tanggal acuan** (bukan TODAY yang berubah-ubah)
- Kolom "Jumlah Anggota Keluarga ≥ 18 Tahun" = anggota (termasuk KK) yang berumur **≥ 18**
- Alamat (RT) diambil otomatis dari header PDF (mis. `RT 010`)

## Validasi (contoh PDF RT 010, MABU'UN)

| Cek | Hasil |
|---|---|
| Jumlah keluarga | **119** = angka resmi RINGKASAN di PDF ✅ |
| Baris hijau = baris KK | 119/119 cocok ✅ |
| Total anggota | 382 |
| Dewasa ≥ 18 | 290 |
| Baris tanpa tanggal lahir | 0 |
| Umur anomali (<0 / >120) | 0 |

## Keamanan

- Path traversal: ✅ blocked (23/23 attacks passed)
- Non-PDF upload: ✅ rejected
- Oversize: ✅ 413
- Loopback-only: ✅ 127.0.0.1