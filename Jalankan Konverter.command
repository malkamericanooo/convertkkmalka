#!/bin/bash
cd "$(dirname "$0")"
echo "=========================================="
echo " Konverter Rekapitulasi Data Keluarga"
echo "=========================================="
echo ""
echo "Membuka browser di http://localhost:8787 ..."
echo "(Tutup jendela ini / Ctrl+C untuk stop server)"
echo ""
sleep 1 && open http://localhost:8787 &
python3 server.py
