"""
sync_loop.pyw — jalankan sync_to_excel.py setiap 30 detik.
Ekstensi .pyw → tidak muncul console window di Windows.
Tambahkan shortcut file ini ke Startup folder agar berjalan otomatis saat login.

Setiap loop: sync transaksi dari bot → Excel
Setiap 10 loop (~5 menit): export data.json → git push ke GitHub Pages
"""
import subprocess
import time
import sys
import os

INTERVAL    = 30   # detik antar sync
PUSH_EVERY  = 10   # git push setiap N loop (~5 menit)

script_dir    = os.path.dirname(os.path.abspath(__file__))
python        = sys.executable
sync_script   = os.path.join(script_dir, "sync_to_excel.py")
export_script = os.path.join(script_dir, "export_json.py")

loop_count = 0

while True:
    # 1. Sync transaksi dari bot → Excel
    try:
        subprocess.run([python, sync_script], timeout=25)
    except Exception:
        pass

    # 2. Export Excel → data.json (setiap loop, cepat ~0.03s)
    try:
        subprocess.run([python, export_script], timeout=15)
    except Exception:
        pass

    # 3. Git push ke GitHub setiap PUSH_EVERY loop
    loop_count += 1
    if loop_count % PUSH_EVERY == 0:
        try:
            subprocess.run(
                ["git", "-C", script_dir, "add", "data.json"],
                timeout=10
            )
            subprocess.run(
                ["git", "-C", script_dir, "commit", "-m",
                 f"auto: update dashboard data [{time.strftime('%Y-%m-%d %H:%M')}]"],
                timeout=15
            )
            subprocess.run(
                ["git", "-C", script_dir, "push", "origin", "master"],
                timeout=30
            )
        except Exception:
            pass

    time.sleep(INTERVAL)
