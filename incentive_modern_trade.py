"""
incentive_modern_trade.py
Channel MODERN TRADE - Achievement per SDO per CYCLE (pivot).

Dipanggil oleh run_pipeline.py setelah transform selesai:
    run(df, label)  -> df = hasil transform, label = rentang tanggal (mis. "1 April - 30 September")

Atau dijalankan sendiri memakai file "Raw data by finance ..." terbaru di folder output:
    python incentive_modern_trade.py

Channel lain (direct, project, indirect) dibuat sebagai script terpisah dengan pola yang sama
(fungsi run(df, label)) lalu didaftarkan di CHANNEL_MODULES pada run_pipeline.py.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from transform_data import OUT_DIR, load_latest_raw

# =====================================================================
# CONFIG - MODERN TRADE
# =====================================================================

FILTER_CHANNEL = "MODERN TRADE"
FILTER_REMARKS = ["SALES", "RETUR"]      # CLAIM G di-exclude dari achievement
ACH_VALUE_COL = "NET VALUE2"             # nilai achievement (NET VALUE x PPN)

# Daftar SDO yang dihitung (semua SDO di Target KA)
SDO_INSENTIF_LIST = [
    "MARTA TEMMY NATASIA",
    "IMANUEL SIREGAR",
    "BAGUS YULIARTO",
    "ANDY",
    "AHMAD DAMYATI",
    "JEKY TIRTA",
    "MOHAMMAD HENDRA FARIZAL",
]

OUTPUT_PREFIX = "Achievement Modern Trade"


# =====================================================================
# PIVOT
# =====================================================================

def _parse_cycle(c):
    """'C0426' -> (26, 4) = (tahun, bulan/urutan cycle). Format lain -> None."""
    m = re.match(r"^C(\d{2})(\d{2})$", str(c).strip())
    return (int(m.group(2)), int(m.group(1))) if m else None


def _quarter(c):
    p = _parse_cycle(c)
    return (p[0], (p[1] - 1) // 3 + 1) if p else None


def _filter(df: pd.DataFrame):
    """Terapkan filter bertahap, kembalikan (data_final, catatan_cek)."""
    cek = {}
    ch = df["CHANNEL"].astype(str).str.upper().str.strip()
    d = df[ch == FILTER_CHANNEL]
    cek["Baris channel " + FILTER_CHANNEL] = (len(d), d[ACH_VALUE_COL].sum())
    if d.empty:
        ada = ", ".join(sorted(ch.dropna().unique())[:15])
        raise ValueError(f"Tidak ada baris dengan CHANNEL = '{FILTER_CHANNEL}'. Channel yang ada: {ada}")

    d = d[d["REMARKS"].isin(FILTER_REMARKS)]
    cek["Setelah filter REMARKS " + "/".join(FILTER_REMARKS)] = (len(d), d[ACH_VALUE_COL].sum())

    # SDO channel ini yang TIDAK ada di daftar (kemungkinan beda penulisan nama)
    luar = d[~d["SDO"].isin(SDO_INSENTIF_LIST)]
    di_luar = (luar.groupby(luar["SDO"].fillna("(KOSONG)"))[ACH_VALUE_COL]
               .agg(["size", "sum"]).sort_values("sum", ascending=False))
    d = d[d["SDO"].isin(SDO_INSENTIF_LIST)]
    cek["Setelah filter SDO di daftar insentif"] = (len(d), d[ACH_VALUE_COL].sum())

    tanpa_cycle = d[d["CYCLE"].isna()]
    cek["Baris tanpa CYCLE (tidak masuk pivot)"] = (len(tanpa_cycle), tanpa_cycle[ACH_VALUE_COL].sum())
    d = d[d["CYCLE"].notna()]
    cek["Masuk pivot"] = (len(d), d[ACH_VALUE_COL].sum())
    return d, cek, di_luar


def build_pivot(d: pd.DataFrame) -> pd.DataFrame:
    """Pivot SDO x CYCLE + subtotal per kuartal + TOTAL."""
    pv = d.pivot_table(index="SDO", columns="CYCLE", values=ACH_VALUE_COL,
                       aggfunc="sum", fill_value=0)
    pv = pv.reindex(SDO_INSENTIF_LIST, fill_value=0)

    cycles = sorted(pv.columns, key=lambda c: (_parse_cycle(c) or (99, 99), str(c)))
    by_q: dict = {}
    for c in cycles:
        by_q.setdefault(_quarter(c), []).append(c)

    out = pd.DataFrame(index=pv.index)
    for q, cs in by_q.items():
        for c in cs:
            out[c] = pv[c]
        if q is not None:
            out[f"TOTAL Q{q[1]} 20{q[0]}"] = pv[cs].sum(axis=1)
    out["TOTAL"] = pv.sum(axis=1)
    out.loc["TOTAL"] = out.sum()
    out.index.name = "SDO"
    return out


# =====================================================================
# SIMPAN
# =====================================================================

def _style_sheet(ws, n_header_cols: int, number_cols_from: int = 2):
    fill = PatternFill("solid", fgColor="1F4E78")
    for c in range(1, n_header_cols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "B2"


def save(pivot: pd.DataFrame, cek: dict, di_luar: pd.DataFrame, label: str) -> Path:
    path = OUT_DIR / f"{OUTPUT_PREFIX} {label}.xlsx"
    cek_df = pd.DataFrame(
        [(k, v[0], v[1]) for k, v in cek.items()],
        columns=["Tahap", "Jumlah baris", f"Total {ACH_VALUE_COL}"],
    )
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            pivot.to_excel(w, sheet_name="Achievement")
            cek_df.to_excel(w, sheet_name="Cek", index=False)
            n = len(cek_df)
            label_cell = w.sheets["Cek"].cell(
                row=n + 3, column=1, value="SDO di channel ini yang TIDAK masuk daftar insentif:")
            label_cell.font = Font(bold=True)
            if di_luar.empty:
                w.sheets["Cek"].cell(row=n + 4, column=1, value="(tidak ada)")
            else:
                di_luar.rename(columns={"size": "Jumlah baris", "sum": f"Total {ACH_VALUE_COL}"}) \
                       .to_excel(w, sheet_name="Cek", startrow=n + 3, index_label="SDO")

            ws = w.sheets["Achievement"]
            _style_sheet(ws, pivot.shape[1] + 1)
            ws.column_dimensions["A"].width = 30
            for i in range(2, pivot.shape[1] + 2):
                ws.column_dimensions[get_column_letter(i)].width = 18
                for r in range(2, pivot.shape[0] + 2):
                    ws.cell(row=r, column=i).number_format = "#,##0"
            for i in range(1, pivot.shape[1] + 2):
                ws.cell(row=pivot.shape[0] + 1, column=i).font = Font(bold=True)

            wc = w.sheets["Cek"]
            wc.column_dimensions["A"].width = 52
            wc.column_dimensions["B"].width = 16
            wc.column_dimensions["C"].width = 22
            for row in wc.iter_rows(min_row=2, min_col=2, max_col=3):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0"
    except PermissionError:
        raise PermissionError(f"Tidak bisa menulis {path.name} - tutup dulu file itu di Excel lalu jalankan ulang.")
    return path


# =====================================================================
# RUN
# =====================================================================

def run(df: pd.DataFrame, label: str) -> Path:
    print("[Modern Trade] Menghitung achievement per SDO per cycle...")
    d, cek, di_luar = _filter(df)
    for k, (n, v) in cek.items():
        print(f"   {k:<45s} {n:>9,} baris   {v:>20,.0f}")
    if not di_luar.empty:
        print(f"   PERHATIAN: {len(di_luar)} nama SDO di channel ini tidak ada di SDO_INSENTIF_LIST "
              f"(lihat sheet 'Cek')")
    if d.empty:
        raise ValueError("Tidak ada data yang lolos filter - pivot tidak dibuat.")

    pivot = build_pivot(d)
    path = save(pivot, cek, di_luar, label)
    print(f"   Tersimpan: {path}")
    return path


if __name__ == "__main__":
    try:
        data, lbl = load_latest_raw()
        run(data, lbl)
    except Exception as e:
        print(f"\n[GAGAL] {e}")
        sys.exit(1)
