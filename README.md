# 🤖 BOT AR Telegram — Chatbot Piutang Real-Time

> **Query data piutang AR kapanpun, dari mana saja — cukup ketik nama atau kode pelanggan di Telegram, bot mengirim laporan bergambar dalam hitungan detik**

Bot Telegram interaktif yang membaca data AR dari `ARVIEWER.xlsm`, memuat seluruh data ke RAM saat startup, dan merespons permintaan anggota tim secara real-time. Pengguna mengetikkan nama pelanggan atau kode pelanggan — bot memandu melalui tiga langkah filter interaktif (Produk → Jatuh Tempo → FRAUD) dan mengirimkan tabel laporan piutang sebagai gambar PNG lengkap dengan status pembayaran per faktur (hanya di aktifkan oleh admin dan untuk produk tertentu (default: off).

---

## 📋 Daftar Isi

- [Gambaran Umum](#-gambaran-umum)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
  - [Langkah Setup (Sekali)](#langkah-setup-sekali)
  - [Menjalankan Bot](#menjalankan-bot)
  - [Menggunakan Bot di Telegram](#menggunakan-bot-di-telegram)
- [Alur Kerja Pipeline](#-alur-kerja-pipeline)
- [Detail Tiap Skrip](#-detail-tiap-skrip)
- [Konfigurasi `config.conf`](#-konfigurasi-configconf)
- [Sistem Pencarian & Resolusi Nama](#-sistem-pencarian--resolusi-nama)
- [Format Output: Gambar Laporan](#-format-output-gambar-laporan)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting & Keamanan](#-catatan-penting--keamanan)

---

## 🗂️ Gambaran Umum

Bot ini menjawab pertanyaan yang paling sering diajukan tim: *"Berapa piutang pelanggan X?"* — tanpa harus membuka Excel, tanpa harus memiliki akses ke ARVIEWER, dan dari perangkat apapun yang memiliki Telegram.

Data diambil dari ekosistem yang sama dengan proyek-proyek lain dalam seri ini: `ARVIEWER.xlsm` sebagai sumber AR, dan (`Hasil_Latihan_temp.xlsx` dari proyek `AR-Orderan-MachineLearning` & `Automasi-AR-Orderan`) sebagai basis resolusi nama serta kode pelanggan.

| Aspek | Keterangan |
|---|---|
| **Antarmuka** | Chat Telegram (teks masuk, gambar keluar) |
| **Akses** | Dilindungi `secret_key` — harus autentikasi sebelum bisa query |
| **Sumber data** | `ARVIEWER.xlsm` (Source, Nama Pelanggan SS, Pembayaran SS) |
| **Resolusi nama** | 6-lapis: branch rules → fb_dict → ml_dict → fuzzy ML → fuzzy FB → fallback |
| **Output** | Gambar PNG tabel matplotlib dikirim langsung ke chat |
| **Refresh data** | Background thread otomatis setiap N menit |

---

## ✨ Fitur Utama

- **Autentikasi per sesi** — Setiap pengguna harus memasukkan `secret_key` sebelum bisa query. Autentikasi bertahan selama bot berjalan.
- **Filter interaktif 3 langkah** — Setelah pencarian awal, pengguna memilih via tombol inline Telegram: (1) filter produk IRC/ZN/SEMUA, (2) filter jatuh tempo saja vs semua faktur, (3) sertakan data FRAUD atau tidak.
- **Multi-jenis query** — Mendukung tiga tipe input dalam satu antarmuka: kode pelanggan, nama pelanggan/kontak, dan kata kunci grup.
- **Multi-kode via `&`** — Satu query seperti `YY-2223 & YY-2224` menampilkan piutang dari dua kode pelanggan sekaligus.
- **Branch rules** — Mapping khusus untuk nama yang memiliki cabang berbeda: query yang mengandung kombinasi kata kunci tertentu diarahkan ke nama kanonik yang tepat (misalnya `sumo godong` vs `sumo purwodadi`).
- **Status pembayaran per faktur** — Setiap baris faktur diberi warna berdasarkan data dari sheet `Pembayaran SS`: biru (LUNAS), oranye (DICICIL), magenta (LEBIH BAYAR), hitam (belum ada data).
- **Gambar tabel berkualitas tinggi** — Tabel dirender via matplotlib dengan header biru gelap, pemisah bulan antar-kelompok faktur, format IDR, dan footer ringkasan total.
- **RAM preload + background refresh** — Seluruh data dimuat ke memori global saat startup; thread daemon memperbarui data secara periodik tanpa menghentikan bot.
- **Pengiriman fallback ke dokumen** — Jika gambar terlalu besar untuk dikirim sebagai foto, bot otomatis mengirim sebagai dokumen PNG.
- **Toleransi DPI otomatis** — Ukuran gambar dibatasi maks 9.000 piksel per sisi; DPI dikurangi otomatis jika data terlalu banyak baris.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl pyTelegramBotAPI matplotlib numpy rapidfuzz
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `pandas` | Skrip 1, 2, 3 | Baca Excel, transformasi, filter |
| `openpyxl` | Skrip 1, 2, 3 | Baca `.xlsm` dan `.xlsx` |
| `pyTelegramBotAPI` (`telebot`) | Skrip 3 | Klien Telegram Bot API |
| `matplotlib` | Skrip 3 | Render tabel sebagai gambar PNG |
| `numpy` | Skrip 3 | Operasi array (digunakan matplotlib) |
| `rapidfuzz` | Skrip 3 | Fuzzy matching nama pelanggan |
| `configparser`, `re`, `os`, `io`, `threading`, `time`, `datetime`, `shutil` | Semua | Standard library |

### Akun & Token yang diperlukan
- **Bot Telegram** — Buat bot baru via [@BotFather](https://t.me/BotFather) di Telegram → dapatkan `bot_token`.
- **`ARVIEWER.xlsm`** — File dashboard AR (dari proyek ARVIEWER). Harus diperbarui secara berkala oleh pipeline ARVIEWER sebelum bot dijalankan.
- **`Hasil_Latihan_temp.xlsx`** *(opsional)* — Hasil pelatihan ML dari proyek `AR-Orderan-MachineLearning`. Jika tidak ada, resolusi nama tetap bekerja via `FBackCust_temp.xlsx` dan fallback.

---

## 📁 Struktur Folder & File

```
📦 BOT-AR-TELEGRAM/
│
├── 📄 Jalankan BOT.py               ← Titik masuk utama. Jalankan ini
│
└── 📁 Dapur/                        ← Folder pipeline (jangan diubah strukturnya)
    ├── 📄 __init__.py
    ├── 📄 1_CopyData.py             ← Ekstrak 4 sumber dari ARVIEWER.xlsm + salin ML
    ├── 📄 2_AdjDateFormat.py        ← Normalisasi tanggal ke format Indonesia
    ├── 📄 3_ARBotTelegram.py        ← Bot Telegram (berjalan terus-menerus)
    └── 📄 config.conf               ← Semua konfigurasi bot dan path data
```

**File sementara yang dihasilkan di `Dapur/` saat runtime:**

| File | Sumber | Isi |
|---|---|---|
| `ARClean_temp.xlsx` | ARVIEWER sheet `arvi_ar_sheet` | Data AR per faktur (sumber utama query) |
| `FBackCust_temp.xlsx` | ARVIEWER sheet `arvi_name_out` | Master nama pelanggan (lookup nama) |
| `PaySS_temp.xlsx` | ARVIEWER sheet `arvi_pay_sales` | Data pembayaran per faktur |
| `Hasil_Latihan_temp.xlsx` | File ML (path `ml_trainning`) | Hasil pelatihan model resolusi nama |

---

## 🚀 Cara Penggunaan

### Langkah Setup (Sekali)

#### 1. Buat Bot Telegram

1. Buka [@BotFather](https://t.me/BotFather) di Telegram.
2. Kirim `/newbot` → ikuti instruksi (nama bot, username bot).
3. Salin **token** yang diberikan BotFather.

#### 2. Isi `config.conf`

Buka `Dapur/config.conf` dan isi nilai-nilai kritis:

```ini
[AUTH]
secret_key = KataRahasia123    ; Sandi yang harus dimasukkan pengguna sebelum bisa query

[BOT]
bot_token = 123456789:ABCdef...  ; Token dari BotFather

[DIR]
arvi = C:\path\ke\ARVIEWER.xlsm   ; Path absolut ke ARVIEWER.xlsm
arvi_ar_sheet = Source             ; Nama sheet AR di ARVIEWER
arvi_name_out = Nama Pelanggan SS  ; Nama sheet master nama
arvi_pay_sales = Pembayaran SS     ; Nama sheet pembayaran
ml_trainning = C:\path\ke\Hasil_Latihan_temp.xlsx  ; Path hasil training ML
```

Lihat panduan lengkap di [Konfigurasi `config.conf`](#-konfigurasi-configconf).

#### 3. Pastikan ARVIEWER.xlsm sudah diperbarui

Bot membaca data dari `ARVIEWER.xlsm` yang sudah ada. Jalankan pipeline ARVIEWER terlebih dahulu untuk memastikan data AR terkini sudah ada di file ini.

### Menjalankan Bot

```bash
python "Jalankan BOT.py"
```

Bot akan:
1. Memvalidasi folder dan file syarat
2. Menjalankan `1_CopyData.py` (ekstraksi data)
3. Menjalankan `2_AdjDateFormat.py` (normalisasi tanggal)
4. Memulai `3_ARBotTelegram.py` (bot aktif)

```
--> Memeriksa struktur folder dan file...
--> Semua file syarat ditemukan. Memulai proses...

--> Memulai eksekusi: 1_CopyData.py
--> Sheet 'Source' berhasil diekstrak ke 'ARClean_temp.xlsx'!
--> Sheet 'Nama Pelanggan SS' berhasil diekstrak ke 'FBackCust_temp.xlsx'!
--> Sheet 'Pembayaran SS' berhasil diekstrak ke 'PaySS_temp.xlsx'!
--> File ML berhasil disalin ke 'Hasil_Latihan_temp.xlsx'!
--> Selesai: 1_CopyData.py

--> Memulai eksekusi: 2_AdjDateFormat.py
--> Format tanggal berhasil diubah.
--> Selesai: 2_AdjDateFormat.py

--> ==================================================
--> Memulai eksekusi 3_ARBotTelegram.py
--> Bot Telegram aktif dan berjalan secara terus-menerus.
--> Tekan Ctrl+C kapan saja untuk menghentikan Bot.
--> ==================================================

--> Memulai Data RAM Preloader...
--> [09:00:00] RAM Data Caching Sukses Diperbarui!
--> Bot Telegram Piutang AR Aktif & Berjalan...
```

Tekan **`Ctrl+C`** untuk menghentikan bot.

### Menggunakan Bot di Telegram

#### Langkah 1 — Autentikasi

Buka bot di Telegram, kirim `/start`. Bot akan meminta sandi:
```
Akses Terkunci.
Masukkan Sandi/Token Internal untuk mengaktifkan bot:
```
Kirim nilai `secret_key` dari `config.conf`. Jika benar:
```
Akses Diterima! Silakan cari data piutang pelanggan.
```

#### Langkah 2 — Kirim Query

Ketikkan pencarian. Bot mendukung tiga format:

| Format | Contoh | Keterangan |
|---|---|---|
| Kode pelanggan | `YY-2223` | Satu kode, distandarisasi otomatis |
| Multi-kode | `YY-2223 & MGL-1045` | Gabungan beberapa kode dengan `&` |
| Nama pelanggan/kontak | `Wakid Kendal` | Fuzzy matching terhadap data AR |
| Kata kunci grup | `manis` | Cocokkan semua pelanggan dengan kata kunci grup |

#### Langkah 3 — Filter Bertahap (3 Langkah via Tombol)

Bot menampilkan jumlah faktur ditemukan lalu memandu 3 langkah filter:

```
Ditemukan 12 faktur piutang.

Langkah 1/3: Pilih Filter Produk:
[ IRC ]  [ ZN ]  [ SEMUA ]
```

```
Langkah 2/3: Pilih Filter Jatuh Tempo (JT):
[ HANYA JT ]  [ SEMUA DATA ]
```

```
Langkah 3/3: Sertakan data Sales FRAUD?
[ TANPA FRAUD ]  [ SERTAKAN FRAUD ]
```

#### Langkah 4 — Terima Gambar Laporan

Bot mengirim gambar tabel piutang dengan caption:
```
Laporan Piutang (IRC)
• Status JT: Hanya JT (>0 hari)
• Sales FRAUD: Dibuang (Tanpa Fraud)
```

---

## 🔄 Alur Kerja Pipeline

```
[Jalankan BOT.py]
   │
   ├─── Validasi folder Dapur/ + 4 file syarat
   │
   ├─── [1] 1_CopyData.py (sekali saat startup)
   │       Baca ARVIEWER.xlsm:
   │         sheet arvi_ar_sheet → ARClean_temp.xlsx
   │         sheet arvi_name_out → FBackCust_temp.xlsx
   │         sheet arvi_pay_sales → PaySS_temp.xlsx
   │       Salin Hasil_Latihan_temp.xlsx → Dapur/
   │
   ├─── [2] 2_AdjDateFormat.py (sekali saat startup)
   │       Baca ARClean_temp.xlsx
   │       Ubah Tgl Faktur & Jatuh Tempo → format Indonesia (misal: "15 Jan 2025")
   │       Simpan kembali ke ARClean_temp.xlsx
   │
   └─── [3] 3_ARBotTelegram.py (berjalan terus-menerus)
           │
           ├─── STARTUP: Preload semua data ke G_DATA (RAM global)
           │       df_ar (dari ARClean_temp)
           │       pay_summary (dari PaySS_temp)
           │       ml_dict + ml_list (dari Hasil_Latihan_temp)
           │       fb_dict + fb_list (dari FBackCust_temp)
           │       group_keywords + branch_rules (dari config.conf)
           │
           ├─── THREAD DAEMON: background_data_refresher()
           │       Setiap N menit → muat ulang semua file ke G_DATA
           │       (data terbaru tersedia tanpa restart bot)
           │
           └─── BOT LOOP: bot.infinity_polling()
                   │
                   ├─── /start atau /help
                   │       Jika belum auth → minta secret_key
                   │       Jika sudah auth → tampilkan instruksi query
                   │
                   ├─── Pesan teks biasa (belum auth)
                   │       Cek apakah == secret_key
                   │       Ya → autentikasi + konfirmasi
                   │       Tidak → "Sandi Salah"
                   │
                   └─── Pesan teks biasa (sudah auth)
                           │
                           ├─ Standarisasi kode (multi-kode via &)
                           ├─ Cari di df_ar berdasarkan Clean_Kode
                           │
                           ├─ Jika tidak ditemukan via kode:
                           │   resolve_target_name_fast() → nama kanonik
                           │   Cek group_keywords
                           │   Exact match Nama Pelanggan / Nama Kontak
                           │   Fuzzy match RapidFuzz token_sort_ratio (≥80%)
                           │
                           ├─ Simpan hasil ke user_sessions[user_id]
                           ├─ Tampilkan tombol filter Produk [IRC|ZN|SEMUA]
                           │
                           ├─ [callback] prod_ → simpan produk, tampilkan filter JT
                           │
                           ├─ [callback] jt_ → simpan filter_jt, tampilkan filter FRAUD
                           │
                           └─ [callback] fraud_ → terapkan semua filter →
                                   generate_ar_image() → kirim foto/dokumen
```

---

## 🔍 Detail Tiap Skrip

### `1_CopyData.py`

Mengekstrak empat sumber data dari `ARVIEWER.xlsm` ke folder `Dapur/`:

| Output | Metode baca | Keterangan |
|---|---|---|
| `ARClean_temp.xlsx` | Auto-detect header (keyword scan) | Data AR lengkap per faktur |
| `FBackCust_temp.xlsx` | `skiprows=1` | Master nama pelanggan, kolom KETERANGAN & NAMA |
| `PaySS_temp.xlsx` | `sheet_name=arvi_pay_sales` | Data pembayaran, header di baris ke-2 (`header=1`) |
| `Hasil_Latihan_temp.xlsx` | `shutil.copyfile` | Salinan hasil training ML |

---

### `2_AdjDateFormat.py`

Membaca `ARClean_temp.xlsx` dan mengonversi dua kolom tanggal ke teks Indonesia:

```
datetime(2025, 1, 15)  →  "15 Jan 2025"
datetime(2025, 12, 31) →  "31 Des 2025"
```

Nama bulan yang digunakan: Jan, Feb, Mar, Apr, Mei, Jun, Jul, Agu, Sep, Okt, Nov, Des.

---

### `3_ARBotTelegram.py`

Inti sistem. Semua data dimuat ke `G_DATA` — dictionary global yang diakses oleh semua handler dalam thread yang berbeda menggunakan `data_lock` dan `session_lock` untuk keamanan thread.

**Status autentikasi** disimpan di `authenticated_users` (set in-memory). Jika bot di-restart, semua pengguna harus autentikasi ulang.

**Sesi pengguna** disimpan di `user_sessions[user_id]` berisi `data` (DataFrame hasil pencarian), `pay_summary`, `prod`, dan `filter_jt` — digunakan antar-callback untuk mempertahankan konteks filter.

---

## ⚙️ Konfigurasi `config.conf`

### `[AUTH]` dan `[BOT]` — Wajib diisi

```ini
[AUTH]
secret_key = KataRahasia123Anda

[BOT]
bot_token = 123456789:ABCdefGHIjklMNO...
```

> ⚠️ Kedua nilai ini **wajib diisi**. Bot tidak akan berjalan jika salah satu kosong — akan menampilkan error `CRITICAL ERROR` saat startup.

---

### `[DIR]` — Path file sumber

```ini
[DIR]
arvi = E:\ADM IRC AND ZN\ARVIEWER.xlsm
arvi_ar_sheet = Source
arvi_name_out = Nama Pelanggan SS
arvi_pay_sales = Pembayaran SS
ml_trainning = E:\ADM IRC AND ZN\AR Pusat Machine Learning\ML\Hasil_Latihan_temp.xlsx
```

| Key | Keterangan |
|---|---|
| `arvi` | Path absolut ke `ARVIEWER.xlsm` |
| `arvi_ar_sheet` | Sheet AR per faktur (sumber data utama) |
| `arvi_name_out` | Sheet master nama pelanggan (kolom `KETERANGAN` dan `NAMA`) |
| `arvi_pay_sales` | Sheet rekap pembayaran (header di baris ke-2) |
| `ml_trainning` | Path ke `Hasil_Latihan_temp.xlsx` dari proyek ML |

---

### `[DISPLAY]` — Tampilan gambar

```ini
[DISPLAY]
show_pay_status = No    ; Ya → tampilkan kolom "Cek Pelunasan SS Sales" di gambar
                        ; No → sembunyikan kolom tersebut (gambar lebih ringkas)
```

---

### `[MAP]` — Awalan kode pelanggan

```ini
[MAP]
depo = SL|YY|MKS|MGL|PW|PWT|PLU|SG|SMG|TGL|PA|KDI
```

Daftar prefix kode pelanggan dipisah `|`. Digunakan oleh fungsi `standardize_code()` untuk menormalisasi format kode: `MGL1234` → `MGL-1234`, `YY 2223` → `YY-2223`.

---

### `[GROUP_KEYWORDS]` — Kata kunci grup pelanggan

```ini
[GROUP_KEYWORDS]
keyword1 = manis
keyword2 = depo
```

Setiap nilai adalah satu kata kunci. Jika query pengguna mengandung kata kunci ini (setelah resolusi nama), semua baris AR yang kolom `Nama Pelanggan` atau `Nama Kontak`-nya mengandung kata kunci ini akan ditampilkan bersama.

**Contoh penggunaan:** Jika ada beberapa pelanggan yang semuanya adalah cabang dari satu distributor "Manis", tambahkan `manis` sebagai group keyword. Query `manis` akan menampilkan semua faktur dari semua cabang distributor tersebut.

---

### `[BRANCH_RULES]` — Aturan resolusi nama cabang

```ini
[BRANCH_RULES]
sumo|godong = CV. SUMO SUKSES SENTOSA GODONG
sumo|purwodadi = CV. SUMO SUKSES SENTOSA
```

Format key: `kata_kunci_1|kata_kunci_2`. Jika query mengandung **keduanya**, nama kanonik di sisi kanan digunakan.

**Contoh:** Query `"sumo godong"` cocok dengan aturan pertama → nama `"CV. SUMO SUKSES SENTOSA GODONG"` digunakan untuk lookup AR. Query `"sumo purwodadi"` cocok dengan aturan kedua.

Jika key tidak mengandung `|`, pencocokan dilakukan sebagai satu kata kunci tunggal (cukup kata kunci ada di query).

---

## 🧠 Sistem Pencarian & Resolusi Nama

### Tipe Query yang Didukung

#### 1. Kode Pelanggan (prioritas pertama)
Query dibersihkan dan distandarisasi, lalu dicari di kolom `Clean_Kode`:

```
"YY2223"         → standardize → "YY-2223"  → cari di AR
"YY-2223 & MGL-1045"  → split & → ["YY-2223", "MGL-1045"]  → cari keduanya
```

Jika ada hasil via kode, proses langsung lanjut ke tampilkan tombol filter.

#### 2. Nama / Kontak Pelanggan (fallback jika kode tidak ditemukan)

Jika pencarian via kode kosong, `resolve_target_name_fast()` dijalankan:

```
Input: nama mentah dari query pengguna
  │
  ├─ 1. Cache hit → return dari cache
  ├─ 2. Branch rules (multi-keyword) → jika query cocok
  ├─ 3. Exact fb_dict (FBackCust, kolom KETERANGAN → NAMA)
  ├─ 4. Exact ml_dict (Hasil_Latihan)
  ├─ 5. Fuzzy ML via WRatio (threshold 75%)
  └─ 6. Fuzzy FB via WRatio (threshold 80%)
       Fallback: gunakan query apa adanya
```

Hasil nama kanonik kemudian digunakan untuk:
1. Cek group_keywords → tampilkan seluruh grup jika cocok
2. Exact match kolom `Nama Pelanggan` atau `Nama Kontak`
3. Fuzzy match `token_sort_ratio` terhadap semua `Nama Pelanggan` (threshold 80%)

### Status Pembayaran (`pay_summary`)

Dibaca dari `PaySS_temp.xlsx` (sheet `Pembayaran SS`). Kolom `nota/invoice/faktur` digunakan sebagai kunci, kolom `bayar/nominal` sebagai nilai:

| Kondisi | Status yang ditampilkan | Warna teks |
|---|---|---|
| `total_bayar == nilai_faktur` | `LUNAS (Total Nx bayar)` | Biru |
| `total_bayar < nilai_faktur` | `DICICIL: Baru Bayar X (Kurang: Y)` | Oranye |
| `total_bayar > nilai_faktur` | `LEBIH BAYAR: Total X` | Magenta |
| Tidak ada data | `Belum Ada Data` | Hitam |

---

## 📤 Format Output: Gambar Laporan

Bot mengirim gambar PNG tabel laporan dengan elemen-elemen berikut:

### Header Gambar
```
PT PRIMA TUNGGAL MANDIRI
```
Nama perusahaan ditampilkan di bagian atas gambar (hardcoded — ubah di `generate_ar_image()` jika perlu).

### Tabel Data

Kolom yang ditampilkan bergantung pada flag `show_pay_status`:

**`show_pay_status = Ya` (10 kolom):**

| Kolom | Lebar relatif | Keterangan |
|---|---|---|
| No. Faktur | 8% | Nomor faktur |
| Tgl Faktur | 7% | Tanggal faktur (format "15 Jan 2025") |
| Jatuh Tempo | 7% | Tanggal jatuh tempo |
| Nilai Faktur | 8% | Nilai faktur (format IDR dengan titik) |
| Sisa Piutang | 8% | Sisa piutang |
| Umur JT | 6% | Umur jatuh tempo |
| Nama Pelanggan | 16% | Dipotong max 30 karakter |
| Nama Penjual | 9% | Dipotong max 16 karakter |
| Nama Kontak | 18% | Dipotong max 35 karakter |
| Cek Pelunasan SS Sales | 13% | Berwarna sesuai status pembayaran |

**`show_pay_status = No` (9 kolom, tanpa kolom pelunasan).**

### Pemisah Bulan
Baris kosong disisipkan otomatis di antara kelompok faktur yang berbeda bulan/tahunnya.

### Styling Tabel
- **Header:** Background biru gelap `#002060`, teks putih tebal
- **Baris LUNAS:** Teks biru
- **Baris DICICIL:** Teks oranye
- **Baris LEBIH BAYAR:** Teks magenta
- **Belum ada data:** Teks hitam biasa

### Footer Gambar
```
Total Nilai Faktur: Rp X  |  Total Sisa Piutang: Rp Y  |  Total Jatuh Tempo (JT): Rp Z
```

### Spesifikasi Teknis Gambar
- **Lebar:** 18 inci
- **Tinggi:** Dinamis (0.32 inci per baris + 1.0 inci)
- **DPI:** 220 (dikurangi otomatis jika pixel > 9.000 per sisi)
- **Format:** PNG via `matplotlib` backend `Agg` (tanpa GUI)

### Caption
```
Laporan Piutang (IRC)
• Status JT: Hanya JT (>0 hari)
• Sales FRAUD: Dibuang (Tanpa Fraud)
```

---

## 🛠️ Troubleshooting

### ❌ `CRITICAL ERROR: 'secret_key' di seksi [AUTH] wajib diisi!`
Isi nilai `secret_key` di `Dapur/config.conf` seksi `[AUTH]`. Nilai tidak boleh kosong.

### ❌ `CRITICAL ERROR: 'bot_token' di seksi [BOT] wajib diisi!`
Isi token bot dari BotFather di `config.conf` seksi `[BOT]`.

### ❌ `[ERROR] File '1_CopyData.py' tidak ditemukan di dalam folder 'Dapur'`
Pastikan seluruh isi folder `Dapur/` ada (`1_CopyData.py`, `2_AdjDateFormat.py`, `3_ARBotTelegram.py`, `config.conf`).

### ❌ `File ARVIEWER tidak ditemukan pada path: ...`
Path di `[DIR] arvi` tidak valid. Pastikan path mengarah ke `ARVIEWER.xlsm` yang benar. Gunakan path absolut untuk menghindari ambiguitas.

### ❌ Bot tidak merespons setelah `/start`
Kemungkinan: (1) `bot_token` salah; (2) bot yang sama sudah aktif di tempat lain (satu bot hanya bisa dijalankan satu instance). Pastikan tidak ada proses bot yang berjalan sebelumnya.

### ❌ "Sandi Salah!" meski kelihatan sudah benar
`secret_key` bersifat exact match — perhatikan spasi di awal/akhir. Coba ketik ulang tanpa menyalin dari dokumen yang mungkin menambahkan spasi tersembunyi.

### ❌ "Data piutang tidak ditemukan" padahal pelanggan ada di AR
Tiga kemungkinan: (1) kode pelanggan tidak terstandarisasi — coba format `PREFIX-ANGKA` seperti `MGL-1234`; (2) nama tidak dikenali oleh sistem resolusi — periksa apakah ada di `FBackCust_temp.xlsx` atau `Hasil_Latihan_temp.xlsx`; (3) data AR belum diperbarui — jalankan ulang pipeline ARVIEWER dulu.

### ❌ Gambar tidak terkirim / error saat kirim foto
Bot secara otomatis akan mencoba kirim sebagai dokumen jika foto gagal. Jika keduanya gagal, kemungkinan masalah koneksi atau gambar korup. Cek log terminal untuk pesan error.

### ❌ `[ERROR CACHE REFRESH]: ...` di terminal
Background thread gagal memuat ulang data. Cek apakah file `ARClean_temp.xlsx` masih ada dan tidak terkunci oleh proses lain. Bot tetap berjalan dengan data lama yang sudah ada di RAM.

### ❌ Bot lambat merespons saat baru dimulai
Data sedang dimuat ke RAM. Jika bot merespons "Data sedang dimuat ke memori RAM, silakan coba beberapa detik lagi", tunggu hingga log terminal menampilkan `RAM Data Caching Sukses Diperbarui!`.

---

## 📌 Catatan Penting & Keamanan

- **`secret_key` dan `bot_token` bersifat rahasia** — Jangan pernah commit `config.conf` ke repositori publik. Tambahkan ke `.gitignore`.
- **Autentikasi tidak permanen** — Jika bot di-restart, semua pengguna harus memasukkan `secret_key` ulang. `authenticated_users` disimpan di memori, bukan di disk.
- **Satu instance per bot** — Telegram API tidak mengizinkan dua proses menggunakan token yang sama secara bersamaan. Pastikan tidak ada instance bot yang sudah berjalan sebelum menjalankan ulang.
- **Nama perusahaan hardcoded** — `"PT PRIMA TUNGGAL MANDIRI"` di header gambar dituliskan langsung di `generate_ar_image()`. Ubah di baris `plt.figtext(...)` jika diperlukan.
- **Bot membaca data statis** — `1_CopyData.py` dan `2_AdjDateFormat.py` hanya berjalan **sekali** saat startup. Pembaruan data AR setelahnya dilakukan oleh background thread yang membaca ulang dari file `*_temp.xlsx` yang sudah ada di `Dapur/`. Untuk memuat data ARVIEWER terbaru, restart bot atau jalankan ulang secara manual.
- **Bergantung pada ARVIEWER** — Bot ini adalah viewer, bukan updater. Pipeline ARVIEWER harus berjalan terpisah untuk menjaga data tetap segar.
- **Sesi pengguna tidak dibersihkan** — `user_sessions` dapat tumbuh seiring banyaknya pengguna. Untuk penggunaan intensif jangka panjang, pertimbangkan membersihkan sesi lama secara periodik.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*REIGHTEEN)*
