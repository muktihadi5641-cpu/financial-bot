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


def run(cmd, timeout=30):
    try:
        subprocess.run(cmd, timeout=timeout, capture_output=True)
    except Exception:
        pass


def data_changed():
    """Return True jika data.json punya perubahan belum di-commit."""
    r = subprocess.run(
        ["git", "-C", script_dir, "diff", "--quiet", "data.json"],
        capture_output=True
    )
    return r.returncode != 0  # 1 = ada perubahan, 0 = tidak ada


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

    # 2. Export Excel → data.json (cepat ~0.03s)
    run([python, export_script], timeout=15)

    # 3. Push ke GitHub hanya jika data.json berubah
    push_if_changed()

    time.sleep(INTERVAL)
