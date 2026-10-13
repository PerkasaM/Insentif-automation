"""
run_pipeline.py
Penghubung: transform -> simpan "Raw data by finance ..." -> (opsional) insentif/performance per channel.

Dipanggil dari run_insentif.bat:
    python run_pipeline.py --channel none           (transform saja)
    python run_pipeline.py --channel modern_trade   (transform + achievement Modern Trade)

Menambah channel baru:
  1. buat script baru dengan fungsi run(df, label)   (contoh: incentive_direct.py)
  2. daftarkan di CHANNEL_MODULES di bawah
  3. tambahkan pilihannya di run_insentif.bat
"""

from __future__ import annotations

import argparse
import importlib
import sys
import time

# channel -> nama module yang punya fungsi run(df, label)
CHANNEL_MODULES = {
    "modern_trade": "incentive_modern_trade",
    # "direct":   "incentive_direct",
    # "project":  "incentive_project",
    # "indirect": "incentive_indirect",
}

CHANNEL_LABEL = {
    "modern_trade": "Modern Trade",
    "direct": "Direct",
    "project": "Project",
    "indirect": "Indirect",
}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--channel", default="none",
                    choices=["none", *CHANNEL_LABEL.keys()],
                    help="channel untuk insentif/performance; 'none' = transform saja")
    channel = ap.parse_args().channel

    # Cek di awal, supaya user tidak menunggu transform untuk channel yang belum ada
    if channel != "none" and channel not in CHANNEL_MODULES:
        print(f"[INFO] Script channel {CHANNEL_LABEL[channel]} belum tersedia.")
        return 1

    t0 = time.time()
    try:
        import transform_data
        df, raw_path, label = transform_data.run()
        print(f"\nTransform selesai ({time.time() - t0:.0f} detik).")

        if channel == "none":
            print("Tidak ada pembuatan insentif/performance yang diminta - selesai.")
            return 0

        print(f"\nLanjut ke channel {CHANNEL_LABEL[channel]}...")
        module = importlib.import_module(CHANNEL_MODULES[channel])
        module.run(df, label)
        print(f"\nSelesai ({time.time() - t0:.0f} detik).")
        return 0
    except Exception as e:
        print(f"\n[GAGAL] {type(e).__name__}: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
