"""
sync_loop.pyw — sync setiap 30 detik + push ke GitHub jika ada data baru.
Ekstensi .pyw = tidak muncul console window di Windows.
"""
import subprocess
import time
import sys
import os

INTERVAL = 30  # detik antar sync

script_dir    = os.path.dirname(os.path.abspath(__file__))
python        = sys.executable
sync_script   = os.path.join(script_dir, "sync_to_excel.py")
export_script = os.path.join(script_dir, "export_json.py")

# Wajib untuk proses tanpa console (.pyw) agar subprocess tidak crash 0xc0000142
NO_WINDOW = subprocess.CREATE_NO_WINDOW


def run(cmd, timeout=30):
    try:
        subprocess.run(cmd, timeout=timeout, capture_output=True,
                       creationflags=NO_WINDOW)
    except Exception:
        pass


def data_changed():
    """Return True jika data.json punya perubahan belum di-commit."""
    try:
        r = subprocess.run(
            ["git", "-C", script_dir, "diff", "--quiet", "data.json"],
            capture_output=True, creationflags=NO_WINDOW, timeout=10
        )
        return r.returncode != 0  # 1 = ada perubahan
    except Exception:
        return False


def push_if_changed():
    if not data_changed():
        return
    run(["git", "-C", script_dir, "add", "data.json"])
    run(["git", "-C", script_dir, "commit", "-m",
         f"auto: update data [{time.strftime('%H:%M')}]"])
    run(["git", "-C", script_dir, "push", "origin", "master"], timeout=45)


while True:
    # 1. Sync transaksi bot → Excel
    run([python, sync_script], timeout=25)

    # 2. Export Excel → data.json
    run([python, export_script], timeout=15)

    # 3. Push ke GitHub hanya jika ada perubahan
    push_if_changed()

    time.sleep(INTERVAL)
