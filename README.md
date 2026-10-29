# Pipeline Scraping & Dataset ABSA Restoran Halal Bali

Repositori ini berisi pipeline otomatisasi untuk mengumpulkan ulasan Google Maps dari **147 restoran halal bersertifikat resmi (BPJPH / LPPOM MUI)** di Provinsi Bali. Pipeline ini menangani proses pencarian tempat otomatis (*auto-search & place profile*), ekstraksi nomor kontak (*telepon*) & tautan situs resmi (*website*), opsi layanan & fasilitas makan presisi, ulasan konsumen, tanggapan pemilik restoran (*balasan pemilik*), pembersihan teks (NLP), deteksi 8 aspek ABSA (*Aspect-Based Sentiment Analysis*), serta pembuatan buku kerja Excel multi-sheet komprehensif (`Dataset ABSA Restoran Halal Bali.xlsx`).

---

## 📁 Lingkup Berkas Proyek

Struktur repositori berfokus pada berkas-berkas utama berikut:

```text
scraping_bali_wulan/
├── data/
│   ├── restoran_halal_bali.csv         # Master data 147 restoran halal Bali (10 kolom murni)
│   ├── raw_ulasan_checkpoint.csv       # File checkpoint hasil scraping ulasan real-time
│   ├── dataset_absa_restoran_halal_bali.csv  # Dataset ulasan bersih hasil pemrosesan Tahap 3
│   ├── Dataset ABSA Restoran Halal Bali.xlsx # Buku kerja Excel 13 sheet lengkap siap pakai
│   └── browser_profile/                # Profil sesi Google Chrome persisten (bebas bot-detection)
├── 02_scrape_ulasan_restoran.py       # [TAHAP 2] Skrip scraping Google Maps restoran
├── 03_proses_dataset_absa.py          # [TAHAP 3] Skrip NLP, pemetaan ABSA & export Excel 13-sheet
├── login_google.py                    # Helper login akun Google ke profil Chrome persisten
├── run_pipeline.py                    # Menu CLI interaktif untuk menjalankan seluruh tahapan
├── requirements.txt                   # Daftar dependensi pustaka Python
└── README.md                          # Dokumentasi panduan lengkap proyek
```

### Rincian Fungsi Berkas

1. **[`02_scrape_ulasan_restoran.py`](02_scrape_ulasan_restoran.py)**:
   - Skrip utama untuk scraping ulasan Google Maps.
   - **Auto-Search**: Mencari restoran di Google Maps jika tautan belum tersedia.
   - **Ekstraksi Kontak**: Mengambil nomor telepon dan URL website resmi restoran.
   - **Ekstraksi Fasilitas Presisi**: Mengambil opsi layanan (`Dine-in`, `Takeaway`, `Delivery`) dan fasilitas aktual dari panel resmi tanpa teks sampah antarmuka.
   - **Multi-Action Infinite Scroll**: Mengombinasikan `scroll_into_view`, *simulated mouse wheel*, dan *container DOM scroll* dengan mekanisme *anti-stuck (nudge)* untuk mengambil ulasan tuntas hingga ratusan ulasan per restoran.
   - **Ekstraksi Balasan Pemilik**: Mendeteksi teks tanggapan pihak pengelola restoran.
   - **Auto-Resume**: Melanjutkan otomatis dari restoran terakhir di checkpoint jika proses sempat dihentikan.

2. **[`03_proses_dataset_absa.py`](03_proses_dataset_absa.py)**:
   - Skrip pemrosesan ulasan dan ekspor dataset akhir.
   - Pembersihan teks NLP (pembersihan karakter khusus, emoji, dan spasi ganda).
   - Deteksi bahasa ulasan menggunakan pustaka `langdetect`.
   - Konversi tanggal ulasan relatif ("3 bulan lalu") menjadi format tanggal ISO 8601.
   - Segmentasi kalimat ulasan dan pelabelan otomatis 8 aspek ABSA kuliner halal.
   - Ekspor buku kerja Excel 13 sheet (`Dataset ABSA Restoran Halal Bali.xlsx`) dan file CSV (`dataset_absa_restoran_halal_bali.csv`).

3. **[`login_google.py`](login_google.py)**:
   - Membuka peramban Google Chrome asli menggunakan folder profil `data/browser_profile/`.
   - Digunakan jika Anda ingin masuk (*login*) ke akun Google secara manual agar riwayat dan sesi login tersimpan permanen dan terhindar dari tantangan CAPTCHA / bot detection.

4. **[`run_pipeline.py`](run_pipeline.py)**:
   - Antarmuka menu interaktif berbasis terminal untuk mempermudah eksekusi tahapan tanpa perlu menghafal perintah baris terminal.

5. **[`requirements.txt`](requirements.txt)**:
   - Berisi seluruh dependensi paket Python: `playwright`, `pandas`, `openpyxl`, `langdetect`, `tqdm`, dan `regex`.

---

## 🛠️ Instalasi & Persiapan

### 1. Instalasi Dependensi Python
Buka terminal / PowerShell di direktori proyek ini, lalu jalankan:
```powershell
pip install -r requirements.txt
```

### 2. Instalasi Browser Playwright
Pastikan peramban Chromium terinstal untuk Playwright:
```powershell
playwright install chromium
```

### 3. (Opsional) Login Akun Google
Jika ingin mengaktifkan akun Google pada profil persisten:
```powershell
python login_google.py
```
*Jendela Chrome asli akan terbuka. Silakan login ke akun Google Anda, lalu tekan `ENTER` pada terminal untuk menyimpan sesi.*

---

## 🚀 Panduan Menjalankan

### Cara 1: Menggunakan Menu Interaktif (Direkomendasikan)
Jalankan file menu utama:
```powershell
python run_pipeline.py
```
Pilihan menu:
* `1` : Menjalankan **Tahap 2 (Scraping Ulasan Restoran)**
* `2` : Menjalankan **Tahap 3 (Pemrosesan ABSA & Export Excel)**
* `3` : Menjalankan alur berurutan (**Tahap 2 langsung ke Tahap 3**)
* `4` : Buka / Cek Sesi Login Akun Google (`login_google.py`)
* `0` : Keluar

---

### Cara 2: Menjalankan Langsung via Perintah Terminal

#### 1. Menjalankan Scraping Ulasan (Tahap 2)
Jalankan skrip `02_scrape_ulasan_restoran.py`:
```powershell
# Jalankan secara default (Chrome visual interaktif):
python 02_scrape_ulasan_restoran.py
```

#### Parameter Baris Perintah (CLI Arguments):
| Parameter | Alias | Deskripsi | Contoh Perintah |
|---|---|---|---|
| `--max` | - | Membatasi kuota ulasan maksimal per restoran | `python 02_scrape_ulasan_restoran.py --max 300` |
| `--start` | - | Memulai scraping dari urutan restoran tertentu (1-indexed) | `python 02_scrape_ulasan_restoran.py --start 10` |
| `--nama` | `--resto`, `--hotel` | Filter dan proses hanya restoran dengan nama tertentu | `python 02_scrape_ulasan_restoran.py --nama "Wong Solo"` |
| `--headless` | - | Menjalankan browser di latar belakang tanpa jendela visual | `python 02_scrape_ulasan_restoran.py --headless` |
| `--retry-incomplete` | - | Memproses kembali restoran yang kuotanya belum terpenuhi | `python 02_scrape_ulasan_restoran.py --retry-incomplete` |

##### Contoh Kasus Penggunaan CLI:
* **Ambil maksimal 300 ulasan per restoran untuk seluruh restoran:**
  ```powershell
  python 02_scrape_ulasan_restoran.py --max 300
  ```
* **Mulai langsung dari restoran nomor urut 25:**
  ```powershell
  python 02_scrape_ulasan_restoran.py --start 25 --max 200
  ```
* **Scraping restoran spesifik berdasarkan nama:**
  ```powershell
  python 02_scrape_ulasan_restoran.py --nama "Goemerot" --max 300
  ```

> **💡 Fitur Auto-Resume:**
> Jika proses scraping dihentikan di tengah jalan (`Ctrl+C`), Anda dapat langsung menjalankan kembali perintah yang sama. Skrip akan **secara otomatis melanjutkan dari restoran berikutnya** tanpa mengulang atau menghapus ulasan yang sudah tersimpan di `data/raw_ulasan_checkpoint.csv`.

---

#### 2. Menjalankan Pemrosesan ABSA & Export Excel (Tahap 3)
Setelah proses scraping selesai atau terdapat data ulasan pada file checkpoint, jalankan:
```powershell
python 03_proses_dataset_absa.py
```
Skrip ini akan memproses seluruh teks ulasan, mendeteksi bahasa, memetakan aspek ABSA, dan menghasilkan:
* `data/dataset_absa_restoran_halal_bali.csv`
* `data/Dataset ABSA Restoran Halal Bali.xlsx`

---

## 📊 Rincian Sheet Buku Kerja Excel (`Dataset ABSA Restoran Halal Bali.xlsx`)

Buku kerja Excel hasil Tahap 3 memiliki **13 lembar kerja (sheets)** yang terstruktur:

| No | Nama Sheet | Deskripsi Isi |
|:--:|---|---|
| 1 | `ulasan` | Dataset ulasan bersih 35 kolom lengkap (teks ulasan, tanggal ISO, rating, nama resto, alamat, telepon, website, fasilitas, balasan pemilik, jumlah like, dll.). |
| 2 | `kalimat_aspek` | Pasangan segmentasi kalimat ulasan yang dipetakan ke 8 aspek ABSA kuliner halal. |
| 3 | `daftar_restoran` | Profil lengkap 147 restoran halal Bali (rating rata-rata, estimasi ulasan, nomor telepon, website, fasilitas, nomor sertifikat, alamat, koordinat). |
| 4 | `olah_fasilitas` | Matriks biner indikator fasilitas restoran (*Dine-in, Takeaway, Delivery, Mushola, Halal Certified, Outdoor, Wi-Fi, dll.*) berstatus "Ya" / "Tidak". |
| 5 | `stat_fasilitas` | Analisis peringkat fasilitas terpopuler di restoran halal Bali beserta rata-rata ratingnya. |
| 6 | `stat_ringkas` | Ringkasan metrik utama dataset (total ulasan, total restoran, persentase ulasan berbalasan pemilik, rata-rata rating, dll.). |
| 7 | `stat_status_halal` | Distribusi kehalalan dan nomor sertifikat resmi. |
| 8 | `stat_aspek` | Rekapitulasi frekuensi kemunculan 8 aspek ABSA dan rata-rata rating per aspek. |
| 9 | `stat_bahasa` | Distribusi bahasa ulasan hasil deteksi (*Indonesian, English, dll.*). |
| 10 | `stat_rating` | Sebaran distribusi bintang ulasan (1 hingga 5). |
| 11 | `stat_tahun` | Tren sebaran volume ulasan per tahun rilis. |
| 12 | `stat_restoran` | Rekapitulasi performa ulasan dan rating rata-rata untuk masing-masing 147 restoran. |
| 13 | `kamus_aspek` | Dokumentasi kamus kata kunci (*regular expressions*) untuk 8 aspek ABSA kuliner halal. |

---

## 🎯 Kamus 8 Aspek ABSA Restoran Halal Bali

1. **`makanan_halal`**: Menilai kehalalan bahan makanan, sertifikasi MUI/BPJPH, kelezatan rasa, bumbu khas, higienitas penyajian, dan variasi menu.
2. **`fasilitas_ibadah`**: Keberadaan mushola, masjid terdekat, tempat wudhu, sajadah, mukena, dan penunjuk arah kiblat.
3. **`kesesuaian_syariah`**: Lingkungan bebas alkohol, suasana ramah keluarga (*family-friendly*), etika berpakaian, dan konsep syariah.
4. **`kebersihan`**: Kebersihan area makan, meja, piring/alat makan, toilet, wastafel, dan sanitasi restoran.
5. **`pelayanan`**: Keramahan staf, kesopanan pelayan/kasir, kecepatan penyajian makanan, kejelasan informasi menu, dan penanganan antrean.
6. **`suasana_fasilitas`**: Kenyamanan ruangan, pendingin udara (AC), jaringan Wi-Fi, area outdoor/indoor, tata musik, dan ketersediaan lahan parkir.
7. **`lokasi_akses`**: Kemudahan menemukan lokasi resto, akses dari jalan raya utama, kedekatan dengan tempat wisata/bandara, serta kondisi lalu lintas.
8. **`harga_value`**: Keterjangkauan harga menu, kesesuaian harga dengan porsi/rasa (*worth it*), paket hemat, dan transparansi tagihan.
