# Statistika Deskriptif Saham BYAN

Analisis deskriptif dan eksplorasi visual data historis saham **PT Bayan Resources Tbk (BYAN)**. Skrip memeriksa kualitas data, merangkum harga, return, dan volume, membandingkan periode terkait event, serta menghasilkan tabel dan grafik untuk mendukung eksplorasi data.

> **Batasan analisis:** proyek ini tidak melakukan uji hipotesis, menghitung *p-value*, memasang distribusi, regresi, korelasi, bootstrap, ataupun *event study* formal. Grafik di sekitar tanggal event hanya menunjukkan pola data, bukan bukti hubungan sebab-akibat.

## Isi repositori

```text
.
├── analisis_byan_final (2).py
├── 3TIC_01_DataHistoris_EventRumorBYAN.xlsx
├── outputs/
│   ├── hasil_statistik_deskriptif.xlsx
│   └── tabel_*.csv
└── figures/
    └── fig_4_*.png
```

Skrip dan workbook Excel merupakan sumber analisis. Direktori `outputs/` dan `figures/` berisi hasil yang dibuat atau diperbarui saat skrip dijalankan.

## Persyaratan

- Python 3
- Paket Python: `numpy`, `pandas`, `matplotlib`, `scipy`, dan `openpyxl`

Pasang paket yang diperlukan melalui terminal:

```bash
python -m pip install numpy pandas matplotlib scipy openpyxl
```

## Menjalankan analisis

Jalankan perintah dari direktori proyek agar skrip dapat menemukan workbook masukan dengan path relatif:

```bash
python "analisis_byan_final (2).py"
```

Skrip membaca workbook, mencetak ringkasan ke terminal, membuat direktori `outputs/` dan `figures/` bila belum tersedia, lalu menyimpan tabel dan grafik. Jika workbook atau sheet wajib tidak ditemukan, skrip akan menampilkan kesalahan.

## Data masukan

File masukan: [`3TIC_01_DataHistoris_EventRumorBYAN.xlsx`](3TIC_01_DataHistoris_EventRumorBYAN.xlsx)

Workbook harus menyediakan dua sheet:

### `BYAN_Analysis`

Kolom yang dibaca oleh skrip:

| Kolom | Keterangan |
| --- | --- |
| `Date` | Tanggal observasi perdagangan |
| `Open`, `High`, `Low`, `Close` | Harga saham |
| `Volume` | Volume perdagangan dalam lot |
| `Return` | Return harian |
| `Event_ID` | Identitas event; nilai `-` digunakan untuk hari tanpa event |
| `Periode_Event` | Label periode: `Normal`, `Pre`, `Event`, atau `Post` |
| `Rumor_Event` | Label rumor (`1`), non-rumor (`0`), atau tanpa label pada hari Normal |

Skrip mengurutkan data berdasarkan tanggal, menghitung `absRet` sebagai nilai absolut dari `Return`, dan menghitung `logVol` sebagai logaritma natural dari volume positif. Nilai volume nol tidak diubah dan tidak disertakan dalam perhitungan logaritma.

### `Event_Rumor`

Sheet ini memuat daftar event dan informasi klasifikasinya. Kolom yang diperlukan adalah `Event_ID`, `Tanggal`, dan `Rumor (1=Rumor, 0=Non-Rumor)`. Kolom seperti `Status_Klarifikasi` digunakan dalam ringkasan jika tersedia.

Skrip juga mencocokkan daftar `Event_ID` pada kedua sheet sebagai bagian dari pemeriksaan kualitas data.

## Cakupan analisis

- **Kualitas data:** tanggal duplikat/tidak valid, nilai hilang, konsistensi return terhadap perubahan `Close`, kewajaran nilai OHLC, volume, dan kecocokan ID event.
- **Statistik keseluruhan dan per periode:** jumlah observasi, mean, median, simpangan baku sampel, minimum, kuartil, IQR, maksimum, *skewness*, dan *excess kurtosis*.
- **Perbandingan rumor dan non-rumor:** ringkasan deskriptif untuk hari yang memiliki label saja.
- **Ringkasan event:** return dan volume pada hari event serta perbandingan volume terhadap median periode Normal.
- **Observasi ekstrem:** identifikasi menggunakan pagar `1,5 × IQR`. Observasi tetap dipertahankan; tidak ada penghapusan data ekstrem.
- **Eksplorasi visual:** seri waktu, histogram, boxplot, sebaran |Return| terhadap log(volume), dan grafik tiap event dalam jendela hari bursa −5 sampai +5.

Statistik per periode disajikan untuk `Return`, `absRet` (|Return|), dan `Volume`. Return pada file tabel disimpan sebagai **proporsi** (contoh: `0,01` berarti `1%`), sementara grafik return menampilkannya dalam persen. Volume dinyatakan dalam lot (`1 lot = 100 lembar`). Transformasi `log(Volume)` hanya digunakan untuk membantu membaca sebaran visual.

## Hasil yang dihasilkan

### Tabel CSV dan workbook Excel

Setiap tabel CSV berada di `outputs/`; workbook Excel menggabungkan tabel-tabel tersebut dalam sheet terpisah.

| Hasil | Isi |
| --- | --- |
| [`tabel_1_kualitas_data.csv`](outputs/tabel_1_kualitas_data.csv) | Pemeriksaan dan kualitas data |
| [`tabel_2_deskriptif_keseluruhan.csv`](outputs/tabel_2_deskriptif_keseluruhan.csv) | Statistik deskriptif keseluruhan, termasuk `Close` |
| [`tabel_3_deskriptif_periode.csv`](outputs/tabel_3_deskriptif_periode.csv) | Statistik `Return`, |Return|, dan volume menurut periode |
| [`tabel_3b_rumor_nonrumor.csv`](outputs/tabel_3b_rumor_nonrumor.csv) | Ringkasan hari berlabel rumor dan non-rumor |
| [`tabel_4_ringkasan_event.csv`](outputs/tabel_4_ringkasan_event.csv) | Ringkasan hari event dan volume relatif terhadap periode Normal |
| [`tabel_5_observasi_ekstrem.csv`](outputs/tabel_5_observasi_ekstrem.csv) | Observasi Return dan volume di luar pagar 1,5 × IQR |
| [`tabel_6_angka_pendukung.csv`](outputs/tabel_6_angka_pendukung.csv) | Angka pendukung pembahasan, rasio, dan hitungan tambahan |
| [`tabel_7_jendela_event_harian.csv`](outputs/tabel_7_jendela_event_harian.csv) | Data harian pada jendela −5 hingga +5 hari bursa tiap event |
| [`tabel_8_ekstrem_per_bulan.csv`](outputs/tabel_8_ekstrem_per_bulan.csv) | Jumlah observasi ekstrem per bulan dan lokasi relatif terhadap event |
| [`hasil_statistik_deskriptif.xlsx`](outputs/hasil_statistik_deskriptif.xlsx) | Gabungan Tabel 1–8 dalam sheet Excel |

### Grafik PNG

Grafik disimpan pada resolusi 300 dpi di `figures/`.

1. [`fig_4_1_harga_volume.png`](figures/fig_4_1_harga_volume.png) — harga penutupan dan volume dengan penanda event.
2. [`fig_4_2_hist_return.png`](figures/fig_4_2_hist_return.png) — histogram return harian.
3. [`fig_4_3_distribusi_volume.png`](figures/fig_4_3_distribusi_volume.png) — distribusi volume pada skala asli dan logaritma.
4. [`fig_4_4_return_periode.png`](figures/fig_4_4_return_periode.png) — boxplot |Return| menurut periode.
5. [`fig_4_5_volume_periode.png`](figures/fig_4_5_volume_periode.png) — boxplot log(volume) menurut periode.
6. [`fig_4_6_return_volume.png`](figures/fig_4_6_return_volume.png) — sebaran |Return| terhadap log(volume), dibedakan menurut periode.
7. [`fig_4_7_event_centered.png`](figures/fig_4_7_event_centered.png) — pola |Return| dan volume pada jendela hari bursa setiap event.

## Catatan interpretasi

- Hari pertama dalam seri tidak memiliki return harian jika tidak ada harga penutupan hari sebelumnya; return tersebut dapat bernilai kosong secara wajar.
- Ukuran sampel berbeda antarperiode. Statistik periode Event perlu dibaca dengan memperhatikan jumlah observasinya.
- Penanda ekstrem berdasarkan IQR adalah alat identifikasi deskriptif, bukan aturan untuk menghapus atau menyatakan data tidak valid.
- Pada grafik event-centered, tiap event ditampilkan sendiri-sendiri dan tidak dirata-ratakan. Di sekitar event yang berdekatan, hari dalam jendela satu event dapat termasuk periode event lain.
- File hasil merefleksikan isi workbook yang dipakai pada saat skrip terakhir dijalankan.

## Snapshot data yang disertakan

Pemeriksaan pada hasil yang saat ini tersedia menunjukkan 650 baris data dari 2 Januari 2024 sampai 1 Oktober 2026, dengan 8 `Event_ID` (2 rumor dan 6 non-rumor). Informasi ini menggambarkan workbook dan keluaran yang saat ini disertakan; hasil dapat berubah bila data masukan diperbarui.
