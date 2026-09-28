# Konverter Rekapitulasi Data Keluarga

PDF ekspor pemerintah (Rekapitulasi Data Keluarga / BKKBN) → **Excel**, tanpa bayar konverter online. Di web, centang file mana saja yang dibutuhkan (default semua). Membaca PDF-nya tetap wajib dan itu yang makan waktu; yang berkurang adalah jumlah file dan ukuran unduhan.

| File | Isi |
|---|---|
| `REKAP ... converted.xlsx` | Replika tabel PDF (seperti konverter berbayar), file paling besar |
| `DESA ....xlsx` | Form survey POPM Filariasis: per KK, jumlah anggota ≥ 18 tahun |
| `DESA ... <KATEGORI>.xlsx` | Satu file per kategori umur yang dicentang di web. Paling atas: jumlah KK + total kategori. Lalu per KK: jumlah di kategori, di luar kategori, total |
| `DESA ... GABUNGAN <...>.xlsx` | Hanya kalau dicentang > 1 kategori: satu kolom per kategori + jumlah kategori terpilih |

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
| `app.py` | Inti konverter (parse PDF → xlsx) |
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
- **Kategori umur** (`KATEGORI_UMUR` di `app.py`), umur tahun penuh pada tanggal acuan:

  | Kategori | Umur |
  |---|---|
  | Bayi Balita | 0–59 bulan (0–4 tahun penuh) |
  | Anak Pra Sekolah | 5–6 tahun |
  | Anak Usia Sekolah & Remaja | 7 tahun – 17 tahun 11 bulan 29 hari |
  | Dewasa / Produktif | 18–59 tahun |
  | Lansia | ≥ 60 tahun |

  Umur 18 hanya masuk Dewasa, jadi kelima kategori menutup semua umur tanpa dobel. Anggota tanpa tanggal lahir valid tetap masuk total dan dicatat di Keterangan. Form POPM (file ke-2) tetap ≥ 18
- Alamat (RT) diambil otomatis dari header PDF (mis. `RT 010`)

## Validasi (contoh PDF RT 010, MABU'UN)

| Cek | Hasil |
|---|---|
| Jumlah keluarga | **119** = angka resmi RINGKASAN di PDF ✅ |
| Baris hijau = baris KK | 119/119 cocok ✅ |
| Total anggota | 382 |
| Dewasa ≥ 18 | 290 |
| Anak < 18 | 92 (cocok dengan kolom USIA di PDF) |
| Balita / Pra Sekolah / Remaja / Dewasa / Lansia (acuan 28-09-2026) | 17 / 14 / 61 / 264 / 26 = 382 (dicek ulang langsung dari tanggal lahir di file REKAP) |
| Baris tanpa tanggal lahir | 0 |
| Umur anomali (<0 / >120) | 0 |

## Keamanan

- Path traversal: ✅ blocked (23/23 attacks passed)
- Non-PDF upload: ✅ rejected
- Oversize: ✅ 413
- Loopback-only: ✅ 127.0.0.1