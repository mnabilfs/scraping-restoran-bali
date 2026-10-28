# Panduan Scraping & Pembentukan Dataset ABSA Restoran Halal Bali

Pipeline ini dirancang khusus untuk mengumpulkan data ulasan Google Maps dari restoran halal bersertifikat resmi (BPJPH / LPPOM MUI) di wilayah **Provinsi Bali** (mencakup Kota Denpasar, Kabupaten Badung, Gianyar, Tabanan, Buleleng, Karangasem, Jembrana, Klungkung, dan Bangli). Pipeline ini mengekstrak ulasan konsumen, tanggapan resmi pengelola restoran (*balasan pemilik*), opsi layanan dan fasilitas makan, melakukan pembersihan teks (NLP), mendeteksi aspek berbasis aspek (ABSA) halal & kualitas kuliner, serta menyusun buku kerja Excel multi-sheet komprehensif (`Dataset ABSA Restoran Halal Bali.xlsx`).

---

## 🌟 Fitur Unggulan Pipeline Restoran Halal Bali

1. **Dataset Master Restoran Halal Bersertifikat (`data/restoran_halal_bali.csv`)**:
   - Berisi 147 data restoran halal terdaftar resmi di Bali dengan atribut lengkap: *Nama Perusahaan, Nama Restaurant, No Sertifikat, Tanggal Terdaftar, Kecamatan, Kota, Alamat, Provinsi, Koordinat, dan Sumber Halal*.
2. **Auto Search & Cache Metadata Cerdas (Tahap 2)**:
   - Jika link Google Maps belum ada di file CSV, sistem secara otomatis mencari restoran di Google Maps menggunakan nama dan alamatnya.
   - Hasil pencarian (URL tempat, `place_id`, `data_id`, rating rata-rata, jumlah total ulasan, kategori, dan koordinat GPS) akan langsung disimpan dan diperbarui kembali ke file CSV master.
3. **Scraping Ulasan Mendalam (Full Infinite Scroll)**:
   - Mengambil ulasan konsumen secara mendalam (default: 100 ulasan per restoran, atau seluruh ulasan).
   - Otomatis mengeklik tombol `Lainnya` / `Lihat lainnya` agar teks ulasan panjang dan balasan pemilik tidak terpotong.
4. **Ekstraksi Fasilitas & Layanan Makan (`fasilitas`)**:
   - Mengambil opsi layanan utama restoran (*Dine-in / Makan di tempat, Takeaway / Bawa pulang, Delivery / Pesan antar, Wi-Fi, Tempat duduk luar, Parkir, Ramah anak, Reservasi, dll.*).
   - Menyediakan sheet khusus `olah_fasilitas` (matriks biner fasilitas per restoran) dan `stat_fasilitas` (analisis statistik sebaran fasilitas restoran di Bali).
5. **Ekstraksi Balasan / Tanggapan Pemilik Restoran (`balasan_pemilik`)**:
   - Mendeteksi dan mengekstrak respon resmi dari manajemen resto (`Tanggapan dari pemilik`) beserta waktu tanggapannya.
6. **Autentikasi Profil Browser Persisten (`data/browser_profile/`)**:
   - Profil Google Chrome tersimpan secara persisten, mencegah deteksi bot dan verifikasi captcha berulang.
7. **Pembersihan NLP & Deteksi 8 Aspek Kuliner Halal (Tahap 3)**:
   - Menghasilkan token kalimat, deteksi bahasa (menggunakan `langdetect`), segmentasi kalimat ulasan, serta pemetaan 8 aspek ABSA:
     1. `makanan_halal` (kehalalan menu, bahan, rasa, kelezatan, porsi, sertifikasi)
     2. `fasilitas_ibadah` (mushola, tempat wudhu, sajadah, arah kiblat, adzan)
     3. `kesesuaian_syariah` (bebas alkohol, citra syariah, ramah keluarga, mahram)
     4. `kebersihan` (higienitas makanan, meja, piring, toilet, lingkungan)
     5. `pelayanan` (keramahan staf, kesopanan pelayan, kecepatan saji, antrean)
     6. `suasana_fasilitas` (kenyamanan, AC, wifi, estetika tempat, musik, parkir)
     7. `lokasi_akses` (kemudahan dijangkau, dekat pantai/objek wisata, parkiran)
     8. `harga_value` (keterjangkauan harga, paket promo, porsi sebanding harga)
8. **Buku Kerja Excel 13 Sheet Lengkap (`Dataset ABSA Restoran Halal Bali.xlsx`)**:
   - Dirancang siap pakai untuk analisis data, visualisasi BI, dan pemodelan Machine Learning / Deep Learning (BERT / IndoBERT).

---

## 📁 Struktur Berkas Proyek

```text
scraping_bali_wulan/
├── data/
│   ├── restoran_halal_bali.csv         # Master data restoran halal Bali (147 entri)
│   ├── raw_ulasan_checkpoint.csv       # File checkpoint hasil scraping ulasan real-time
│   ├── dataset_absa_restoran_halal_bali.csv  # Dataset ulasan akhir bersih hasil Tahap 3
│   ├── Dataset ABSA Restoran Halal Bali.xlsx # Workbook Excel 13 sheet lengkap
│   └── browser_profile/                # Profil sesi Google Chrome persisten
├── 02_scrape_ulasan_restoran.py       # [TAHAP 2] Skrip scraping Google Maps restoran
├── 03_proses_dataset_absa.py            # [TAHAP 3] Skrip pembersihan & ekspor Excel ABSA
├── run_pipeline.py                      # Menu interaktif untuk menjalankan tahapan
├── login_google.py                      # Helper untuk login akun Google jika diperlukan
└── README.md                            # Panduan dokumentasi proyek
```

---

## 🚀 Panduan Menjalankan

### Cara 1: Menggunakan Menu Interaktif (Direkomendasikan)
Buka terminal / PowerShell di folder proyek, lalu jalankan:
```powershell
python run_pipeline.py
```
Pilih opsi menu:
- Ketik `1` untuk menjalankan **Tahap 2 (Scraping Ulasan Restoran)**
- Ketik `2` untuk menjalankan **Tahap 3 (Pemrosesan ABSA & Export Excel)**
- Ketik `3` untuk menjalankan secara berurutan (**Tahap 2 langsung ke Tahap 3**)

---

### Cara 2: Menjalankan Skrip Langsung dari Terminal

#### 1. Menjalankan Tahap 2 (Scraping Ulasan)
Anda dapat menjalankan file utama `02_scrape_ulasan_restoran.py`:
```powershell
# Jalankan secara default (mode jendela Chrome terlihat & interaktif):
python 02_scrape_ulasan_restoran.py
```

### ⚙️ Fleksibilitas Argumen CLI (Parameter Perintah)

Skrip Tahap 2 dilengkapi dengan berbagai opsi baris perintah (*CLI arguments*) yang sangat fleksibel:

| Parameter | Alias | Deskripsi | Contoh Perintah |
|---|---|---|---|
| `--max` | - | Membatasi jumlah kuota ulasan maksimal per restoran | `python 02_scrape_ulasan_restoran.py --max 50` |
| `--start` | - | Memulai proses langsung dari nomor urut restoran tertentu (1-indexed) | `python 02_scrape_ulasan_restoran.py --start 15` |
| `--nama` | `--resto`, `--hotel` | Filter dan proses hanya restoran yang cocok dengan nama tertentu | `python 02_scrape_ulasan_restoran.py --nama "Pagi Sore"` |
| `--headless` | - | Menjalankan browser di latar belakang tanpa memunculkan jendela GUI | `python 02_scrape_ulasan_restoran.py --headless` |
| `--retry-incomplete` | - | Menyapu dan melanjutkan kembali restoran yang ulasannya belum mencapai kuota target | `python 02_scrape_ulasan_restoran.py --retry-incomplete` |

#### 💡 Contoh Penggunaan Argumen CLI:

1. **Membatasi kuota ulasan menjadi 50 ulasan per restoran:**
   ```powershell
   python 02_scrape_ulasan_restoran.py --max 50
   ```

2. **Memulai scraping langsung dari restoran urutan ke-20:**
   ```powershell
   python 02_scrape_ulasan_restoran.py --start 20
   ```

3. **Scraping restoran tertentu berdasarkan nama (menggunakan `--nama`, `--resto`, atau `--hotel`):**
   ```powershell
   # Menggunakan flag --nama
   python 02_scrape_ulasan_restoran.py --nama "Bebek Tepi Sawah"

   # Menggunakan flag --resto
   python 02_scrape_ulasan_restoran.py --resto "Kebuli Jordan"

   # Menggunakan flag --hotel
   python 02_scrape_ulasan_restoran.py --hotel "Wong Solo"
   ```

4. **Kombinasi parameter (misalnya mulai dari restoran ke-10, kuota maksimal 30 ulasan per restoran):**
   ```powershell
   python 02_scrape_ulasan_restoran.py --start 10 --max 30
   ```

5. **Scraping spesifik satu restoran dengan kuota 100 ulasan:**
   ```powershell
   python 02_scrape_ulasan_restoran.py --nama "Restoran Pagi Sore" --max 100
   ```

> **💡 Catatan Auto-Resume:**
> Skrip Tahap 2 dilengkapi fitur auto-resume. Jika Anda menghentikan proses (Ctrl+C) dan menjalankannya kembali, skrip akan **secara otomatis melanjutkan dari restoran berikutnya**, tanpa menghapus atau mengulang ulasan yang sudah tersimpan di `data/raw_ulasan_checkpoint.csv`.

#### 2. Menjalankan Tahap 3 (Pemrosesan & Pembuatan Dataset Excel)
Setelah scraping selesai (atau jika sudah memiliki ulasan di checkpoint), jalankan:
```powershell
python 03_proses_dataset_absa.py
```
Skrip ini akan memproses seluruh data teks ulasan, mendeteksi bahasa, memetakan aspek, dan membuat file:
- `data/dataset_absa_restoran_halal_bali.csv`
- `data/Dataset ABSA Restoran Halal Bali.xlsx`

---

## 📊 Rincian Sheet pada File Excel Akhir

Buku kerja Excel `Dataset ABSA Restoran Halal Bali.xlsx` berisi 13 lembar kerja:

1. **`ulasan`**: Seluruh baris ulasan bersih (35 kolom lengkap: teks asli, terjemahan, rating, tanggal ISO, nama restoran, alamat, telepon, website resmi, status halal, opsi layanan, fasilitas, balasan pemilik, jumlah like, reviewer total ulasan, dll.).
2. **`kalimat_aspek`**: Pasangan kalimat ulasan yang telah disegmentasi dan dipetakan ke 8 aspek ABSA.
3. **`daftar_restoran`**: Master profil 147 restoran halal Bali beserta rating, total ulasan, nomor telepon, tautan website resmi, fasilitas, nomor sertifikat halal, dan koordinat.
4. **`olah_fasilitas`**: Matriks indikator fasilitas kuliner (Dine-in, Takeaway, Delivery, Mushola, Halal Certified, Outdoor seating, Wi-Fi, dll.) berstatus "Ya" / "Tidak".
5. **`stat_fasilitas`**: Rekapitulasi ranking fasilitas terpopuler di restoran halal Bali beserta rata-rata ratingnya.
6. **`stat_ringkas`**: Ringkasan umum metrik dataset.
7. **`stat_status_halal`**: Distribusi status kehalalan tempat.
8. **`stat_aspek`**: Rekapitulasi frekuensi aspek ABSA dan rata-rata ratingnya.
9. **`stat_bahasa`**: Sebaran bahasa ulasan (Indonesia, Inggris, dll.).
10. **`stat_rating`**: Distribusi bintang ulasan (1 hingga 5).
11. **`stat_tahun`**: Distribusi tahun ulasan ulasan.
12. **`stat_restoran`**: Rekapitulasi ulasan dan performa rating per restoran.
13. **`kamus_aspek`**: Dokumentasi kata kunci regex dan 8 aspek ABSA kuliner halal.
