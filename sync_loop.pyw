"""
sync_loop.pyw — jalankan sync_to_excel.py setiap 2 menit.
Ekstensi .pyw → tidak muncul console window di Windows.
Tambahkan shortcut file ini ke Startup folder agar berjalan otomatis saat login.
"""
import subprocess
import time
import sys
import os

INTERVAL = 30  # detik
script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sync_to_excel.py")
python = sys.executable

while True:
    try:
        subprocess.run([python, script], timeout=25)
    except Exception:
        pass
    time.sleep(INTERVAL)
