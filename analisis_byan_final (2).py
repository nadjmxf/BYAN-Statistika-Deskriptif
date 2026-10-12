# -*- coding: utf-8 -*-
"""
STATISTIKA DESKRIPTIF & EKSPLORASI VISUAL - PT Bayan Resources Tbk (BYAN)
=========================================================================
Cakupan : pemeriksaan data, statistika deskriptif, perbandingan deskriptif
          antarperiode, eksplorasi visual. TIDAK ada uji hipotesis, p-value,
          fitting distribusi, regresi/korelasi, bootstrap, atau event study formal.
Input   : 3TIC_01_DataHistoris_EventRumorBYAN.xlsx
          - BYAN_Analysis : Date, Open, High, Low, Close, Volume, Return, %DVolume,
                            Event, Event_ID, Periode_Event, Rumor_Event
          - Event_Rumor   : Event_ID, Tanggal, Isi_Topik_Rumor, Status_Klarifikasi,
                            'Rumor (1=Rumor, 0=Non-Rumor)', Nomor_Surat, Baris_Data_BYAN
Output  : outputs/ (CSV + 1 Excel)   figures/ (PNG 300 dpi)

Satuan  : Return & |Return| disimpan sebagai PROPORSI (0,01 = 1%) pada tabel;
          grafik menampilkannya dalam %. Volume dalam LOT (1 lot = 100 lembar).
          log(Volume) = logaritma natural, hanya alat bantu visualisasi.
"""
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
from matplotlib.ticker import FuncFormatter
from scipy import stats

# ==========================================================================
# 00. KONFIGURASI
# ==========================================================================
INPUT_FILE = Path("3TIC_01_DataHistoris_EventRumorBYAN.xlsx")
OUT_DIR, FIG_DIR = Path("outputs"), Path("figures")
OUT_DIR.mkdir(exist_ok=True)
FIG_DIR.mkdir(exist_ok=True)

SHEET_DATA, SHEET_EVENT = "BYAN_Analysis", "Event_Rumor"
COL_RUMOR = "Rumor (1=Rumor, 0=Non-Rumor)"
PERIODS = ["Normal", "Pre", "Event", "Post"]
BIAS_CORRECTED = True      # True = koreksi sampel (sama dengan SKEW() dan KURT() Excel)
IQR_K = 1.5                # aturan pagar 1,5 x IQR (hanya untuk MENGIDENTIFIKASI)
WINDOW = 5                 # lebar tampilan event-centered: -5 .. +5 hari bursa
FIG_DPI, HIST_BINS = 300, 50
C_RUMOR, C_NON = "#C00000", "0.35"   # merah hanya untuk membedakan rumor


def fmt_id(x, pos=None):
    """Pemisah ribuan gaya Indonesia (titik)."""
    return f"{x:,.0f}".replace(",", ".")


def save_fig(fig, name):
    fig.savefig(FIG_DIR / name, dpi=FIG_DPI, bbox_inches="tight")
    plt.close(fig)
    print("  tersimpan:", FIG_DIR / name)


def save_csv(t, name, index=False):
    t.to_csv(OUT_DIR / name, index=index, encoding="utf-8-sig")  # presisi penuh
    print("  tersimpan:", OUT_DIR / name)


# ==========================================================================
# 01. MUAT & SIAPKAN DATA
# ==========================================================================
def load_and_prepare(path):
    xl = pd.ExcelFile(path)
    for s in (SHEET_DATA, SHEET_EVENT):
        if s not in xl.sheet_names:
            raise KeyError(f"Sheet '{s}' tidak ada. Tersedia: {xl.sheet_names}")
    df, ev = xl.parse(SHEET_DATA), xl.parse(SHEET_EVENT)
    need = ["Date", "Open", "High", "Low", "Close", "Volume", "Return", "Event_ID",
            "Periode_Event", "Rumor_Event"]
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise KeyError(f"Kolom {miss} tidak ada di {SHEET_DATA}. Tersedia: {df.columns.tolist()}")

    n_cols_raw = df.shape[1]                                  # jumlah kolom asli file (sebelum kolom turunan)
    df["Date"] = pd.to_datetime(df["Date"])
    df = df.sort_values("Date").reset_index(drop=True)
    df.attrs["n_cols_raw"] = n_cols_raw
    df.attrs["return_recalc"] = df["Close"].pct_change()          # (Ct - Ct-1)/Ct-1
    df["absRet"] = df["Return"].abs()
    df["logVol"] = np.log(df["Volume"].where(df["Volume"] > 0))  # nol -> NaN, tidak dimanipulasi
    df["Periode_Event"] = pd.Categorical(df["Periode_Event"], categories=PERIODS, ordered=True)
    df["Rumor_Num"] = pd.to_numeric(df["Rumor_Event"], errors="coerce")   # Normal ('-') -> NaN
    df["Rumor_Label"] = df["Rumor_Num"].map({1: "Rumor", 0: "Non-Rumor"})

    ev = ev.rename(columns={COL_RUMOR: "Rumor"})
    ev["Tanggal"] = pd.to_datetime(ev["Tanggal"])
    ev["Rumor_Label"] = ev["Rumor"].map({1: "Rumor", 0: "Non-Rumor"})
    return df, ev


# ==========================================================================
# 02. PEMERIKSAAN DATA  -> Tabel 1
# ==========================================================================
def data_quality_check(df, ev):
    r = []

    def add(item, hasil, status=""):
        r.append({"Pemeriksaan": item, "Hasil": hasil, "Status": status})

    ok = lambda b: "OK" if b else "PERIKSA"
    add("Jumlah baris data", len(df))
    add("Jumlah kolom data (kolom asli file)", df.attrs.get("n_cols_raw", df.shape[1]))
    add("Periode data", f"{df['Date'].min():%d-%m-%Y} s.d. {df['Date'].max():%d-%m-%Y}")
    add("Tanggal tidak valid", int(df["Date"].isna().sum()), ok(df["Date"].isna().sum() == 0))
    add("Tanggal duplikat", int(df["Date"].duplicated().sum()), ok(not df["Date"].duplicated().any()))
    for c in ["Open", "High", "Low", "Close", "Volume"]:
        add(f"Missing {c}", int(df[c].isna().sum()), ok(df[c].isna().sum() == 0))
    n_ret_na = int(df["Return"].isna().sum())
    add("Return hilang (hari pertama tidak punya hari sebelumnya)", n_ret_na, ok(n_ret_na == 1))
    add("Return valid", int(df["Return"].notna().sum()))
    diff = (df["Return"] - df.attrs["return_recalc"]).abs().max()
    add("Selisih maks Return bawaan vs hitung ulang dari Close", f"{diff:.2e}", ok(diff < 1e-9))
    add("Volume <= 0", int((df["Volume"] <= 0).sum()), ok((df["Volume"] <= 0).sum() == 0))
    add("OHLC tidak wajar: High < Low", int((df["High"] < df["Low"]).sum()),
        ok((df["High"] < df["Low"]).sum() == 0))
    out = ~df["Close"].between(df["Low"], df["High"])
    add("OHLC tidak wajar: Close di luar rentang Low-High", int(out.sum()), ok(out.sum() == 0))
    add("Return = 0 (harga penutupan tidak berubah)", int((df["Return"] == 0).sum()))
    add("Return > 0 / Return < 0", f"{int((df['Return'] > 0).sum())} / {int((df['Return'] < 0).sum())}")
    cnt = df["Periode_Event"].value_counts().reindex(PERIODS)
    for p in PERIODS:
        add(f"Jumlah observasi periode {p}", int(cnt[p]))
    add("Jumlah observasi Normal+Pre+Event+Post = n data", int(cnt.sum()), ok(cnt.sum() == len(df)))
    add("Jumlah Event_ID unik (sheet Event_Rumor)", int(ev["Event_ID"].nunique()),
        ok(ev["Event_ID"].is_unique))
    add("Event_ID antar-sheet cocok",
        sorted(set(df["Event_ID"].astype(str)) - {"-"}) == sorted(ev["Event_ID"]),
        ok(set(df["Event_ID"].astype(str)) - {"-"} == set(ev["Event_ID"])))
    add("Jumlah event Rumor / Non-rumor",
        f"{int((ev['Rumor'] == 1).sum())} / {int((ev['Rumor'] == 0).sum())}")
    return pd.DataFrame(r)


# ==========================================================================
# 03. STATISTIK DESKRIPTIF  -> Tabel 2 dan 3
# ==========================================================================
def descriptive_stats(x):
    x = pd.Series(x).dropna()
    n = len(x)
    q1, q3 = x.quantile(0.25), x.quantile(0.75)
    return {
        "n": n, "mean": x.mean(), "median": x.median(), "sd": x.std(ddof=1),
        "min": x.min(), "Q1": q1, "Q3": q3, "IQR": q3 - q1, "max": x.max(),
        "skewness": stats.skew(x, bias=not BIAS_CORRECTED) if n > 2 else np.nan,
        "excess_kurtosis": stats.kurtosis(x, fisher=True, bias=not BIAS_CORRECTED) if n > 3 else np.nan,
    }


VARS = ["Return", "absRet", "Volume"]          # variabel yang dibandingkan antarperiode
OVERALL_VARS = ["Close"] + VARS                # tabel keseluruhan juga memuat harga Close


def table_overall(df):
    t = pd.DataFrame({v: descriptive_stats(df[v]) for v in OVERALL_VARS}).T
    t["n"] = t["n"].astype(int)
    t.index.name = "Variabel"
    return t.reset_index()


def table_by_period(df):
    rows = []
    for v in VARS:
        for p in PERIODS:
            d = descriptive_stats(df.loc[df["Periode_Event"] == p, v])
            rows.append({"Variabel": v, "Periode": p, **d})
    t = pd.DataFrame(rows)
    t["n"] = t["n"].astype(int)
    return t


def table_rumor(df):
    """Hanya hari berlabel (Pre/Event/Post); Normal tidak berlabel. Deskriptif saja."""
    sub = df[df["Rumor_Label"].notna()]
    g = sub.groupby("Rumor_Label")
    t = pd.DataFrame({
        "n_hari": g.size(), "n_event": g["Event_ID"].nunique(),
        "mean_Return": g["Return"].mean(), "sd_Return": g["Return"].std(ddof=1),
        "median_absRet": g["absRet"].median(), "mean_absRet": g["absRet"].mean(),
        "median_Volume": g["Volume"].median()}).reindex(["Rumor", "Non-Rumor"])
    return t.reset_index().rename(columns={"Rumor_Label": "Kelompok"})


# ==========================================================================
# 04. RINGKASAN EVENT  -> Tabel 4
# ==========================================================================
def event_summary(df, ev):
    med_normal = df.loc[df["Periode_Event"] == "Normal", "Volume"].median()
    day = df[df["Periode_Event"] == "Event"].set_index("Event_ID")
    rows = []
    for _, r in ev.iterrows():
        d, win = day.loc[r["Event_ID"]], df[df["Event_ID"] == r["Event_ID"]]
        rows.append({
            "Event_ID": r["Event_ID"], "Tanggal": r["Tanggal"].date(),
            "Rumor_Label": r["Rumor_Label"], "Status_Klarifikasi": r.get("Status_Klarifikasi"),
            "Return_hari_event": d["Return"], "absRet_hari_event": d["absRet"],
            "Volume_hari_event": d["Volume"], "Median_Volume_Normal": med_normal,
            "Rasio_Volume_vs_Median_Normal": d["Volume"] / med_normal,
            "n_pre": int((win["Periode_Event"] == "Pre").sum()),
            "n_post": int((win["Periode_Event"] == "Post").sum())})
    return pd.DataFrame(rows)


# ==========================================================================
# 05. OBSERVASI EKSTREM (aturan 1,5 x IQR; tidak ada yang dihapus)  -> Tabel 5
# ==========================================================================
def extreme_observations(df):
    rows = []
    for v in ["Return", "Volume"]:
        x = df[v].dropna()
        q1, q3 = x.quantile(0.25), x.quantile(0.75)
        lo, hi = q1 - IQR_K * (q3 - q1), q3 + IQR_K * (q3 - q1)
        sel = df[(df[v] < lo) | (df[v] > hi)]
        for _, r in sel.iterrows():
            in_win = r["Periode_Event"] != "Normal"
            rows.append({
                "Variabel": v, "Tanggal": r["Date"].date(), "Nilai": r[v],
                "Batas_bawah": lo, "Batas_atas": hi,
                "Arah": "atas" if r[v] > hi else "bawah",
                "Periode_Event": str(r["Periode_Event"]), "Event_ID": r["Event_ID"],
                "Keterangan": ("dalam event window" if in_win
                               else "pergerakan ekstrem di luar event yang teridentifikasi")})
    t = pd.DataFrame(rows)
    t["_mag"] = t["Nilai"].abs()
    return t.sort_values(["Variabel", "_mag"], ascending=[True, False]).drop(columns="_mag").reset_index(drop=True)



# ==========================================================================
# 05b. ANGKA PENDUKUNG PEMBAHASAN  -> Tabel 6, 7, 8
#      Setiap angka yang dikutip di Bab IV tetapi bukan isi Tabel 1-5 keluar di sini.
# ==========================================================================
def table_supporting(df, t3, t5):
    """Tabel 6: angka pendukung (rasio, pembagi, hitungan bersyarat, statistik log(Volume), lebar kelas)."""
    rows = []

    def add(kategori, ukuran, nilai, dipakai):
        rows.append({"Kategori": kategori, "Ukuran": ukuran, "Nilai": nilai, "Dipakai_di": dipakai})

    d = df.dropna(subset=["Return"])                       # 649 hari ber-Return
    big = d["absRet"] > 0.05
    p90 = d["Volume"].quantile(0.90)
    hi_vol = d["Volume"] > p90

    # --- hari dengan |Return| > 5% dan hubungannya dengan volume (Gambar 4.6)
    add("|Return| > 5%", "Jumlah hari (total)", int(big.sum()), "4.4 Gambar 4.6")
    for p in PERIODS:
        add("|Return| > 5%", f"Jumlah hari, periode {p}", int((big & (d["Periode_Event"] == p)).sum()), "4.4 Gambar 4.6")
    add("Persentil ke-90 Volume", "Persentil ke-90 (lot, n = 649 hari ber-Return)", p90, "4.4 Gambar 4.6")
    add("Persentil ke-90 Volume", "Jumlah hari Volume > persentil ke-90", int(hi_vol.sum()), "4.4 Gambar 4.6")
    add("Persentil ke-90 Volume", "Hari |Return| > 5% dengan Volume > persentil ke-90", int((big & hi_vol).sum()), "4.4 Gambar 4.6")
    add("Persentil ke-90 Volume", "Proporsi hari |Return| > 5% yang Volume-nya > persentil ke-90",
        (big & hi_vol).sum() / big.sum(), "4.4 Gambar 4.6")
    nrm = d["Periode_Event"] == "Normal"
    m_big, m_small = d.loc[nrm & big, "Volume"].median(), d.loc[nrm & ~big, "Volume"].median()
    add("Normal: Volume menurut besar |Return|", "Median Volume, hari Normal |Return| > 5% (lot)", m_big, "4.4 Gambar 4.6")
    add("Normal: Volume menurut besar |Return|", "Median Volume, hari Normal |Return| <= 5% (lot)", m_small, "4.4 Gambar 4.6")
    add("Normal: Volume menurut besar |Return|", "Rasio kedua median (kali)", m_big / m_small, "4.4 Gambar 4.6")

    # --- log(Volume) per periode dan keseluruhan (Gambar 4.3 dan 4.5)
    for p in PERIODS:
        lv = df.loc[df["Periode_Event"] == p, "logVol"].dropna()
        add("log(Volume) per periode", f"{p}: median", lv.median(), "4.4 Gambar 4.5")
        add("log(Volume) per periode", f"{p}: Q1", lv.quantile(0.25), "4.4 Gambar 4.5")
        add("log(Volume) per periode", f"{p}: Q3", lv.quantile(0.75), "4.4 Gambar 4.5")
    lv = df["logVol"].dropna()
    add("log(Volume) keseluruhan", "Minimum", lv.min(), "4.4 Gambar 4.3")
    add("log(Volume) keseluruhan", "Median", lv.median(), "4.4 Gambar 4.3")
    add("log(Volume) keseluruhan", "Maksimum", lv.max(), "4.4 Gambar 4.3")
    vmax_n = df.loc[df["Periode_Event"] == "Normal", "Volume"].max()
    add("Volume Normal tertinggi", "Volume (lot)", vmax_n, "4.4 Gambar 4.5")
    add("Volume Normal tertinggi", "log(Volume)", np.log(vmax_n), "4.4 Gambar 4.5")

    # --- lebar kelas histogram (Gambar 4.2 dan 4.3)
    r = d["Return"] * 100
    add("Lebar kelas histogram", "Return (poin persen)", (r.max() - r.min()) / HIST_BINS, "4.4 Gambar 4.2")
    vmin, vmax = df["Volume"].min(), df["Volume"].max()
    w = (vmax - vmin) / HIST_BINS
    add("Lebar kelas histogram", "Volume (lot)", w, "4.4 Gambar 4.3")
    add("Lebar kelas histogram", "Batas atas kelas pertama Volume (lot)", vmin + w, "4.4 Gambar 4.3")

    # --- rasio dari Tabel 2 dan 3
    v, rt = df["Volume"].dropna(), d["Return"]
    add("Rasio Volume keseluruhan", "Mean / median", v.mean() / v.median(), "4.2")
    add("Rasio Volume keseluruhan", "Maksimum / median", v.max() / v.median(), "4.2")
    add("Rasio Volume keseluruhan", "Maksimum / Q3", v.max() / v.quantile(0.75), "4.2")
    add("Rasio Volume keseluruhan", "SD / IQR", v.std(ddof=1) / (v.quantile(0.75) - v.quantile(0.25)), "4.2")
    add("Rasio Return keseluruhan", "SD / IQR", rt.std(ddof=1) / (rt.quantile(0.75) - rt.quantile(0.25)), "4.2")
    g = t3.set_index(["Variabel", "Periode"])
    for var, col, lab in (("absRet", "median", "Median |Return|"), ("absRet", "mean", "Mean |Return|"),
                          ("Return", "IQR", "IQR Return"), ("Return", "sd", "SD Return"),
                          ("Volume", "median", "Median Volume")):
        base = g.loc[(var, "Normal"), col]
        for p in PERIODS[1:]:
            add("Rasio periode terhadap Normal (kali)", f"{lab}: {p} / Normal", g.loc[(var, p), col] / base, "4.3")

    # --- komposisi rumor vs non-rumor
    sub = df[df["Rumor_Label"].notna()]
    for k in ["Rumor", "Non-Rumor"]:
        c = sub[sub["Rumor_Label"] == k]["Periode_Event"].value_counts()
        n = int(c.sum())
        for p in ["Pre", "Event", "Post"]:
            add("Komposisi hari berlabel", f"{k}: hari {p}", int(c.get(p, 0)), "4.3 Tabel 4.5")
        add("Komposisi hari berlabel", f"{k}: proporsi hari Pre", c.get("Pre", 0) / n, "4.3 Tabel 4.5")

    # --- observasi ekstrem (turunan Tabel 5)
    t = t5.copy()
    t["Tanggal"] = pd.to_datetime(t["Tanggal"])
    outside = t["Keterangan"].str.startswith("pergerakan ekstrem di luar")
    for var in ["Return", "Volume"]:
        x = t[t["Variabel"] == var]
        o = outside[x.index]
        add("Observasi ekstrem", f"{var}: total ditandai", len(x), "4.6")
        add("Observasi ekstrem", f"{var}: ditandai di luar event window (Normal)", int(o.sum()), "4.6")
        add("Observasi ekstrem", f"{var}: proporsi di luar event window", o.mean(), "4.6")
        add("Observasi ekstrem", f"{var}: proporsi dalam event window", 1 - o.mean(), "4.6")
    ret_n = t[(t["Variabel"] == "Return") & outside]
    vol_n = t[(t["Variabel"] == "Volume") & outside]
    add("Observasi ekstrem Normal", "Return ditandai pada Jan-Sep 2026", int((ret_n["Tanggal"].dt.year == 2026).sum()), "4.6")
    add("Observasi ekstrem Normal", "Volume ditandai pada Jul-Agu 2024",
        int(((vol_n["Tanggal"].dt.year == 2024) & vol_n["Tanggal"].dt.month.isin([7, 8])).sum()), "4.6")
    add("Observasi ekstrem Normal", "Volume ditandai pada Jun, Agu, Sep 2026",
        int(((vol_n["Tanggal"].dt.year == 2026) & vol_n["Tanggal"].dt.month.isin([6, 8, 9])).sum()), "4.6")
    return pd.DataFrame(rows)


def table_event_window(df, ev):
    """Tabel 7: data harian -5..+5 hari bursa per event (sumber Gambar 4.7 dan pembahasan 4.5)."""
    med_normal = df.loc[df["Periode_Event"] == "Normal", "Volume"].median()
    rows = []
    for _, r in ev.iterrows():
        i0 = int(df.index[df["Date"] == r["Tanggal"]][0])
        for i in range(max(0, i0 - WINDOW), min(len(df) - 1, i0 + WINDOW) + 1):
            x = df.loc[i]
            rows.append({"Event_ID": r["Event_ID"], "Hari_relatif": i - i0, "Tanggal": x["Date"].date(),
                         "Periode_Event": str(x["Periode_Event"]), "Event_ID_baris": x["Event_ID"],
                         "Return": x["Return"], "absRet": x["absRet"], "Volume": x["Volume"],
                         "Rasio_Volume_vs_Median_Normal": x["Volume"] / med_normal})
    return pd.DataFrame(rows)


def table_extreme_by_month(t5):
    """Tabel 8: jumlah observasi ekstrem per bulan, dipisah dalam/luar event window."""
    t = t5.copy()
    t["Bulan"] = pd.to_datetime(t["Tanggal"]).dt.to_period("M").astype(str)
    t["Lokasi"] = np.where(t["Keterangan"].str.startswith("pergerakan ekstrem di luar"),
                           "luar event window", "dalam event window")
    return (t.groupby(["Variabel", "Bulan", "Lokasi"]).size().rename("Jumlah_hari").reset_index())

# ==========================================================================
# 06. FIGURE
# ==========================================================================
def fig1_time_series(df, ev):
    fig, (a1, a2) = plt.subplots(2, 1, figsize=(12, 7), sharex=True,
                                 gridspec_kw={"height_ratios": [3, 2]})
    a1.plot(df["Date"], df["Close"], color="black", lw=1.0)
    a2.plot(df["Date"], df["Volume"], color="black", lw=0.8)
    top = df["Close"].max()
    a1.set_ylim(df["Close"].min() * 0.93, top * 1.16)
    levels = [1.12, 1.07, 1.02]                      # label bertingkat agar event berdekatan terbaca
    for i, r in ev.reset_index(drop=True).iterrows():
        rumor = r["Rumor"] == 1
        col = C_RUMOR if rumor else C_NON
        for ax in (a1, a2):
            ax.axvline(r["Tanggal"], color=col, ls="-" if rumor else "--", lw=1.0, alpha=0.85)
        a1.text(r["Tanggal"], top * levels[i % 3], r["Event_ID"], color=col, fontsize=8,
                ha="center", va="center", fontweight="bold",
                bbox=dict(facecolor="white", edgecolor="none", pad=0.8))
    a1.set_ylabel("Harga penutupan (Rp)")
    a2.set_ylabel("Volume (lot)")
    a2.set_xlabel("Tanggal")
    for ax in (a1, a2):
        ax.yaxis.set_major_formatter(FuncFormatter(fmt_id))
    a2.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    a2.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    a1.legend(handles=[Line2D([0], [0], color="black", lw=1, label="Harga penutupan (Close)"),
                       Line2D([0], [0], color=C_RUMOR, ls="-", label="Tanggal event: rumor"),
                       Line2D([0], [0], color=C_NON, ls="--", label="Tanggal event: non-rumor")],
              loc="lower left", frameon=False)
    a1.set_title("Harga Penutupan dan Volume Harian BYAN dengan Penanda Event")
    fig.text(0.01, -0.01, "Sumber: BYAN_Analysis dan Event_Rumor. Garis vertikal = tanggal event; "
             "tidak menunjukkan hubungan sebab-akibat.", fontsize=8, ha="left")
    fig.tight_layout()
    save_fig(fig, "fig_4_1_harga_volume.png")


def fig2_hist_return(df):
    x = df["Return"].dropna() * 100
    fig, ax = plt.subplots(figsize=(9, 5.5))
    ax.hist(x, bins=HIST_BINS, color="0.65", edgecolor="0.25", lw=0.5)
    ax.axvline(0, color="black", ls="--", lw=1.0, label="Return = 0")
    ax.set_xlabel("Return harian (%)")
    ax.set_ylabel("Frekuensi (jumlah hari)")
    ax.set_title(f"Histogram Return Harian BYAN (n = {len(x)})")
    ax.legend(frameon=False)
    fig.text(0.01, -0.02, "Sumber: BYAN_Analysis.", fontsize=8, ha="left")
    fig.tight_layout()
    save_fig(fig, "fig_4_2_hist_return.png")


def fig3_volume_distribution(df):
    v, lv = df["Volume"].dropna(), df["logVol"].dropna()
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(12, 5))
    for ax, x, lab, ttl in ((a1, v, "Volume (lot)", f"Volume, skala asli (n = {len(v)})"),
                            (a2, lv, "log(Volume)  [Volume dalam lot]", f"log(Volume) (n = {len(lv)})")):
        ax.hist(x, bins=HIST_BINS, color="0.65", edgecolor="0.25", lw=0.5)
        ax.axvline(x.median(), color="black", ls="--", lw=1.0, label="Median")
        ax.set_xlabel(lab)
        ax.set_ylabel("Frekuensi (jumlah hari)")
        ax.set_title(ttl)
        ax.legend(frameon=False)
    a1.xaxis.set_major_formatter(FuncFormatter(fmt_id))
    fig.suptitle("Distribusi Volume Harian BYAN: Skala Asli dan Skala Logaritma Natural")
    fig.text(0.01, -0.02, "Sumber: BYAN_Analysis. Transformasi log hanya mengubah skala agar bentuk sebaran "
             "terbaca; observasi ekstrem tetap ada.", fontsize=8, ha="left")
    fig.tight_layout()
    save_fig(fig, "fig_4_3_distribusi_volume.png")


def _boxplot(df, col, scale, ylabel, title, fname):
    data, labels = [], []
    for p in PERIODS:
        x = df.loc[df["Periode_Event"] == p, col].dropna().to_numpy() * scale
        data.append(x)
        labels.append(f"{p}\n(n = {len(x)})")
    fig, ax = plt.subplots(figsize=(8, 6))
    bp = ax.boxplot(data, patch_artist=True, widths=0.55, medianprops=dict(color="black", lw=1.6),
                    flierprops=dict(marker="o", markersize=3.5, markerfacecolor="none", markeredgecolor="0.3"))
    for b in bp["boxes"]:
        b.set(facecolor="0.85", edgecolor="0.2")
    ax.set_xticks(range(1, 5))
    ax.set_xticklabels(labels)
    ax.set_xlabel("Periode")
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    fig.text(0.01, -0.02, "Sumber: BYAN_Analysis. Kotak: Q1-Q3; garis tengah: median; whisker: 1,5 x IQR; "
             "lingkaran: di luar whisker (tidak dihapus).", fontsize=8, ha="left")
    fig.tight_layout()
    save_fig(fig, fname)


def fig4_absret_period(df):
    _boxplot(df, "absRet", 100, "|Return| harian (%)",
             "Boxplot |Return| Harian BYAN per Periode", "fig_4_4_return_periode.png")


def fig5_logvol_period(df):
    _boxplot(df, "logVol", 1, "log(Volume)  [Volume dalam lot]",
             "Boxplot log(Volume) Harian BYAN per Periode", "fig_4_5_volume_periode.png")


def fig6_scatter(df):
    d = df.dropna(subset=["absRet", "logVol"])
    style = {"Normal": dict(marker="o", s=14, c="0.75", edgecolor="none", alpha=0.7, zorder=1),
             "Pre": dict(marker="^", s=38, c="0.45", edgecolor="black", lw=0.5, zorder=2),
             "Event": dict(marker="*", s=130, c="black", edgecolor="black", lw=0.5, zorder=3),
             "Post": dict(marker="s", s=34, c="white", edgecolor="black", lw=0.8, zorder=2)}
    fig, ax = plt.subplots(figsize=(9, 6))
    for p in PERIODS:
        s = d[d["Periode_Event"] == p]
        ax.scatter(s["logVol"], s["absRet"] * 100, label=f"{p} (n = {len(s)})", **style[p])
    ax.set_xlabel("log(Volume)  [Volume dalam lot]")
    ax.set_ylabel("|Return| harian (%)")
    ax.set_title(f"|Return| terhadap log(Volume) Harian BYAN (n = {len(d)})")
    ax.legend(title="Periode", frameon=False)
    fig.text(0.01, -0.02, "Sumber: BYAN_Analysis. Eksplorasi visual saja; tidak ada korelasi/regresi yang dihitung.",
             fontsize=8, ha="left")
    fig.tight_layout()
    save_fig(fig, "fig_4_6_return_volume.png")


def fig7_event_centered(df, ev, summary):
    """Pola tiap event pada -5..+5 hari bursa. Batang = |Return| (%), garis = volume / median Normal."""
    med_normal = df.loc[df["Periode_Event"] == "Normal", "Volume"].median()
    fig, axes = plt.subplots(2, 4, figsize=(15, 7.5), sharey=True)
    for ax, (_, r) in zip(axes.ravel(), ev.reset_index(drop=True).iterrows()):
        eid = r["Event_ID"]
        i0 = int(df.index[df["Date"] == r["Tanggal"]][0])
        idx = [i for i in range(i0 - WINDOW, i0 + WINDOW + 1) if 0 <= i < len(df)]
        w = df.loc[idx]
        rel = np.array(idx) - i0
        own = (w["Event_ID"] == eid).to_numpy()          # hari yang dilabeli sebagai jendela event ini
        event_col = C_RUMOR if r["Rumor"] == 1 else "black"
        for x, a, is_own, is_day in zip(rel, w["absRet"] * 100, own, x_is_day(rel)):
            ax.bar(x, a, width=0.7, color=event_col if is_day else "0.7", edgecolor="0.2", lw=0.4,
                   hatch=None if is_own else "///", alpha=1 if is_own else 0.55)
        ax.axvline(0, color="0.5", ls=":", lw=0.8)
        ax.set_xticks(range(-WINDOW, WINDOW + 1, 1))
        ax.tick_params(axis="x", labelsize=7)
        ax.set_title(f"{eid} ({r['Tanggal']:%d/%m/%Y}, {r['Rumor_Label'].replace('Non-Rumor', 'non-rumor').replace('Rumor', 'rumor')})",
                     fontsize=9)
        ax2 = ax.twinx()
        ax2.plot(rel, w["Volume"] / med_normal, color="black", marker="o", ms=3, lw=1.0)
        ax2.set_yscale("log")
        ax2.set_ylim(0.2, 400)
        if ax in axes[:, -1]:
            ax2.set_ylabel("Volume / median Volume Normal (kali, skala log)", fontsize=8)
        else:
            ax2.set_yticklabels([])
        ax2.tick_params(labelsize=7)
    for ax in axes[:, 0]:
        ax.set_ylabel("|Return| (%)")
    for ax in axes[-1, :]:
        ax.set_xlabel("Hari bursa relatif terhadap hari event")
    fig.legend(handles=[Patch(facecolor="0.7", edgecolor="0.2", label="|Return| hari dalam jendela event ini"),
                        Patch(facecolor="0.7", edgecolor="0.2", hatch="///", alpha=0.55,
                              label="|Return| hari di luar jendela event ini (milik event lain / Normal)"),
                        Patch(facecolor="black", label="|Return| hari event (merah = rumor)"),
                        Line2D([0], [0], color="black", marker="o", ms=3, label="Volume / median Volume Normal")],
               loc="lower center", ncol=2, frameon=False, fontsize=8, bbox_to_anchor=(0.5, -0.04))
    fig.suptitle("Pola |Return| dan Volume pada -5 s.d. +5 Hari Bursa di Sekitar Setiap Event")
    fig.text(0.01, -0.07, f"Sumber: BYAN_Analysis dan Event_Rumor. Median Volume Normal = {fmt_id(med_normal)} lot. "
             "Setiap event ditampilkan terpisah; tidak dirata-ratakan.", fontsize=8, ha="left")
    fig.tight_layout()
    save_fig(fig, "fig_4_7_event_centered.png")


def x_is_day(rel):
    return [x == 0 for x in rel]


# ==========================================================================
# 07. VALIDASI RINGKAS
# ==========================================================================
def validate(df, t_overall, t_period):
    ok = True
    for v in VARS:
        x = df[v].dropna().to_numpy(float)
        g = t_overall.set_index("Variabel").loc[v]
        ok &= (len(x) == g["n"]) and np.isclose(x.mean(), g["mean"]) and np.isclose(np.median(x), g["median"])
    ok &= t_period.query("Variabel == 'Volume'")["n"].sum() == len(df)
    ok &= t_period.query("Variabel == 'Return'")["n"].sum() == df["Return"].notna().sum()
    print("\nVALIDASI: tabel konsisten dengan data ->", "OK" if ok else "GAGAL, periksa sebelum dipakai")
    return ok


# ==========================================================================
# 08. MAIN
# ==========================================================================
def main():
    pd.set_option("display.width", 220)
    pd.set_option("display.max_columns", 50)
    pd.set_option("display.float_format", lambda v: f"{v:,.6f}")

    df, ev = load_and_prepare(INPUT_FILE)

    print("\n[Tabel 1] Pemeriksaan & kualitas data")
    t1 = data_quality_check(df, ev)
    print(t1.to_string(index=False))
    save_csv(t1, "tabel_1_kualitas_data.csv")

    print("\n[Tabel 2] Statistik deskriptif keseluruhan (Return & |Return| = proporsi)")
    t2 = table_overall(df)
    print(t2.to_string(index=False))
    save_csv(t2, "tabel_2_deskriptif_keseluruhan.csv")

    print("\n[Tabel 3] Statistik deskriptif per periode")
    t3 = table_by_period(df)
    print(t3.to_string(index=False))
    save_csv(t3, "tabel_3_deskriptif_periode.csv")
    t3b = table_rumor(df)
    print("\n[Tabel 3b] Rumor vs Non-rumor (deskriptif; hanya hari berlabel)")
    print(t3b.to_string(index=False))
    save_csv(t3b, "tabel_3b_rumor_nonrumor.csv")

    print("\n[Tabel 4] Ringkasan per event")
    t4 = event_summary(df, ev)
    print(t4.to_string(index=False))
    save_csv(t4, "tabel_4_ringkasan_event.csv")

    print("\n[Tabel 5] Observasi ekstrem (aturan 1,5 x IQR)")
    t5 = extreme_observations(df)
    print(t5.groupby(["Variabel", "Keterangan"]).size().to_string())
    print("\n10 observasi Return paling ekstrem:")
    print(t5[t5["Variabel"] == "Return"].head(10).to_string(index=False))
    save_csv(t5, "tabel_5_observasi_ekstrem.csv")

    print("\n[Tabel 6] Angka pendukung pembahasan")
    t6 = table_supporting(df, t3, t5)
    print(t6.to_string(index=False))
    save_csv(t6, "tabel_6_angka_pendukung.csv")

    print("\n[Tabel 7] Data harian -5..+5 hari bursa per event")
    t7 = table_event_window(df, ev)
    print(t7.to_string(index=False))
    save_csv(t7, "tabel_7_jendela_event_harian.csv")

    print("\n[Tabel 8] Observasi ekstrem per bulan")
    t8 = table_extreme_by_month(t5)
    print(t8.to_string(index=False))
    save_csv(t8, "tabel_8_ekstrem_per_bulan.csv")

    print("\n[Figure]")
    fig1_time_series(df, ev)
    fig2_hist_return(df)
    fig3_volume_distribution(df)
    fig4_absret_period(df)
    fig5_logvol_period(df)
    fig6_scatter(df)
    fig7_event_centered(df, ev, t4)

    validate(df, t2, t3)

    with pd.ExcelWriter(OUT_DIR / "hasil_statistik_deskriptif.xlsx", engine="openpyxl") as w:
        for name, t in (("1_Kualitas_Data", t1), ("2_Keseluruhan", t2), ("3_Periode", t3),
                        ("3b_Rumor", t3b), ("4_Event", t4), ("5_Ekstrem", t5),
                        ("6_Pendukung", t6), ("7_Jendela_Event", t7), ("8_Ekstrem_Bulan", t8)):
            t.to_excel(w, sheet_name=name, index=False)
    print("\n  tersimpan:", OUT_DIR / "hasil_statistik_deskriptif.xlsx")


if __name__ == "__main__":
    main()
