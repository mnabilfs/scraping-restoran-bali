"""
TAHAP 3: Pemrosesan Dataset ABSA Restoran Halal Bali & Export ke Excel Lengkap
Input:
  - data/restoran_halal_bali.csv (Master 147 Restoran Halal)
  - data/raw_ulasan_checkpoint.csv
Output:
  - data/dataset_absa_restoran_halal_bali.csv
  - data/Dataset ABSA Restoran Halal Bali.xlsx (13 Sheet Lengkap: ulasan, kalimat_aspek, daftar_restoran, olah_fasilitas, stat_fasilitas, kamus_aspek, dan sheet statistik)
"""

import os
import sys
import re
from datetime import datetime, timedelta
from pathlib import Path
import warnings
warnings.filterwarnings('ignore', category=FutureWarning)
warnings.filterwarnings('ignore', message='.*select_dtypes.*')
import pandas as pd
from langdetect import detect

# Pastikan output konsol Windows mendukung UTF-8
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"

INPUT_RESTO_FILE = DATA_DIR / "restoran_halal_bali.csv"
REVIEW_FILE = DATA_DIR / "raw_ulasan_checkpoint.csv"

OUTPUT_CSV = DATA_DIR / "dataset_absa_restoran_halal_bali.csv"
OUTPUT_EXCEL = DATA_DIR / "Dataset ABSA Restoran Halal Bali.xlsx"

# Kamus 8 Aspek ABSA Khusus Restoran Halal di Bali
KAMUS_ASPEK = {
    "makanan_halal": {
        "regex": r"halal|haram|non[- ]?halal|pork|babi|bacon|ham\b|lard|sarapan|breakfast|makan|makanan|food|menu|restoran|restaurant|dapur|kitchen|catering|buffet|prasmanan|sertifikat|certified|certification|MUI|BPJPH|rasa|lezat|enak|gurih|sedap|pedas|asin|manis|sambal|bumbu|daging|ayam|bebek|ikan|seafood|sate|nasi|mie|goreng|bakar|kuah|sop|porsi",
        "inti_halal": True
    },
    "fasilitas_ibadah": {
        "regex": r"musala|mushola|musholla|masjid|mosque|sajadah|sarung|mukena|kiblat|qibla|qiblah|kibla|sholat|shalat|salat|prayer|pray|adzan|azan|adhan|quran|al[- ]?qur|wudhu|wudu|ablution|tempat wudhu",
        "inti_halal": True
    },
    "kesesuaian_syariah": {
        "regex": r"alkohol|alcohol|miras|bir\b|beer|wine|bar\b|nightclub|club|party|pesta|privasi|privacy|privat|hijab|berhijab|islami|islamic|syariah|sharia|muslim|moslem|family friendly|ramah keluarga|mahram|non[- ]?muslim|halal certified|jaminan halal",
        "inti_halal": True
    },
    "kebersihan": {
        "regex": r"bersih|kotor|clean|dirty|bau|smell|odor|debu|dust|jorok|higienis|hygien|rapi|tidy|meja|lantai|piring|sendok|garpu|gelas|kamar mandi|bathroom|toilet|wastafel|tempat cuci|kutu|kecoa|cockroach|semut|ant\b|lalat|fly",
        "inti_halal": False
    },
    "pelayanan": {
        "regex": r"layanan|pelayanan|service|staf|staff|karyawan|pegawai|pelayan|waiter|waitress|kasir|cashier|ramah|friendly|helpful|sopan|polite|responsif|responsive|lambat|slow|cepat|fast|antre|antrean|queue|waiting list|manajemen|management|hospitality|keramahan",
        "inti_halal": False
    },
    "suasana_fasilitas": {
        "regex": r"suasana|atmosphere|ambience|tempat|ruangan|meja|kursi|ac\b|air conditioner|pendingin|wifi|wi-fi|internet|musik|music|colokan|stopkontak|toilet|wastafel|parkir|parking|parkiran|luas|sempit|spacious|small|cozy|nyaman|estetik|aesthetic|outdoor|indoor|balkon|view|pemandangan|kolam",
        "inti_halal": False
    },
    "lokasi_akses": {
        "regex": r"lokasi|location|strategis|strategic|dekat|near|close|jauh|far|akses|access|pantai|beach|bandara|airport|walking distance|jalan kaki|macet|traffic|sekitar|area|lingkungan|neighborhood|pusat|central|pinggir jalan|mudah dicari",
        "inti_halal": False
    },
    "harga_value": {
        "regex": r"harga|price|priced|murah|cheap|mahal|expensive|worth|value|terjangkau|affordable|biaya|tarif|rate\b|promo|diskon|discount|sebanding|overpriced|budget|porsi|paket|hemat|bill|tagihan",
        "inti_halal": False
    }
}

# Regex khusus untuk kolom flag aspek halal per ulasan
REGEX_FLAGS_HALAL = {
    "flag_makanan_halal": r"\b(halal|haram|non[- ]?halal|pork|babi|bacon|ham|sarapan halal|makanan halal|daging halal)\b",
    "flag_fasilitas_ibadah": r"\b(musala|mushola|musholla|masjid|sajadah|sarung|mukena|kiblat|sholat|shalat|salat|wudhu|adzan)\b",
    "flag_citra_syariah": r"\b(syariah|sharia|islami|islamic|muslim|moslem|ramah keluarga|family friendly|mahram)\b",
    "flag_bebas_alkohol": r"\b(alkohol|miras|bir|wine|bar|cocktail)\b"
}

def parse_relative_date(date_str: str) -> tuple[datetime, str]:
    now = datetime.now()
    if not isinstance(date_str, str) or not date_str.strip():
        iso_str = now.strftime("%Y-%m-%dT%H:%M:%SZ")
        return now, iso_str

    clean_str = re.sub(r"^(Diedit|Ditulis)\s+", "", date_str, flags=re.IGNORECASE).strip()
    match = re.search(r"(\d+)\s+(hari|minggu|bulan|tahun|jam|menit)", clean_str, re.IGNORECASE)
    
    if match:
        val = int(match.group(1))
        unit = match.group(2).lower()
        if "hari" in unit:
            dt = now - timedelta(days=val)
        elif "minggu" in unit:
            dt = now - timedelta(weeks=val)
        elif "bulan" in unit:
            dt = now - timedelta(days=val * 30)
        elif "tahun" in unit:
            dt = now - timedelta(days=val * 365)
        elif "jam" in unit:
            dt = now - timedelta(hours=val)
        elif "menit" in unit:
            dt = now - timedelta(minutes=val)
        else:
            dt = now
    else:
        dt = now

    iso_str = dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    return dt, iso_str

def clean_review_text(text: str) -> str:
    if not isinstance(text, str):
        return ""
    t = re.sub(r"\s+", " ", text)
    t = re.sub(r"[\r\n\t]+", " ", t)
    return t.strip()

def safe_detect_lang(text: str) -> str:
    if not isinstance(text, str) or len(text.strip()) < 10:
        return "id"
    try:
        return detect(text)
    except Exception:
        return "id"

def clean_csv_cell(val) -> str:
    if val is None or pd.isna(val):
        return ""
    txt = str(val)
    txt = re.sub(r"[\r\n]+", " ", txt)
    txt = txt.replace("\t", " ")
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt

def sanitize_for_excel(df_in: pd.DataFrame) -> pd.DataFrame:
    df_out = df_in.copy()
    for col in df_out.select_dtypes(include=["object", "str"]).columns:
        df_out[col] = df_out[col].apply(
            lambda x: re.sub(r"^[=@]+\s*", "", str(x)) if isinstance(x, str) and len(x) > 0 and x[0] in ["=", "@"] else x
        )
    return df_out

def proses_pipeline():
    print("=" * 65)
    print("🚀 MEMULAI TAHAP 3: PREPROCESSING & EXPORT DATASET ABSA RESTORAN BALI")
    print("=" * 65)

    target_resto_file = INPUT_RESTO_FILE
    if not target_resto_file.exists():
        print(f"❌ File master restoran tidak ditemukan: {target_resto_file}")
        return

    if not REVIEW_FILE.exists():
        print(f"❌ File checkpoint ulasan belum ada: {REVIEW_FILE}")
        print("   Silakan jalankan '02_scrape_ulasan_restoran.py' terlebih dahulu untuk scraping ulasan!")
        return

    df_resto = pd.read_csv(target_resto_file, low_memory=False, dtype=str).fillna("")
    df_reviews = pd.read_csv(REVIEW_FILE, low_memory=False, dtype=str).fillna("")

    # Pastikan kolom-kolom standar ada di df_resto agar tidak KeyError
    for col in ["data_id", "place_id", "url", "rating_rata2", "jumlah_ulasan_total", "kategori", "fasilitas", "telepon", "website", "lat", "lon", "Koordinat"]:
        if col not in df_resto.columns:
            df_resto[col] = ""

    print(f"📋 Memproses {len(df_resto)} restoran dan {len(df_reviews)} ulasan...")

    # Buat kamus lookup restoran berdasarkan data_id dan nama_restaurant
    resto_map_by_id = {}
    resto_map_by_name = {}

    for _, r in df_resto.iterrows():
        did = str(r.get("data_id", "")).strip()
        nama = str(r.get("Nama Restaurant", "")).strip()
        r_dict = r.to_dict()
        if did:
            resto_map_by_id[did] = r_dict
        if nama:
            resto_map_by_name[nama.lower()] = r_dict

    # 1. Pemetaan Fasilitas, Kontak (Telepon & Website), data_id, dan Rating dari Ulasan ke Restoran
    fac_map_id, fac_map_name = {}, {}
    tel_map_id, tel_map_name = {}, {}
    web_map_id, web_map_name = {}, {}
    did_map_name = {}
    rat_map_name = {}

    for _, rf in df_reviews.iterrows():
        did_r = str(rf.get("data_id", "")).strip()
        nm_r = str(rf.get("nama_restoran") or rf.get("nama_penginapan") or "").strip().lower()
        fc_r = str(rf.get("fasilitas", "")).strip()
        tel_r = str(rf.get("telepon", "")).strip()
        web_r = str(rf.get("website", "")).strip()
        rat_r = str(rf.get("rating_restoran", "")).strip()

        if did_r and nm_r and nm_r not in did_map_name:
            did_map_name[nm_r] = did_r
        if fc_r:
            if did_r and did_r not in fac_map_id: fac_map_id[did_r] = fc_r
            if nm_r and nm_r not in fac_map_name: fac_map_name[nm_r] = fc_r
        if tel_r:
            if did_r and did_r not in tel_map_id: tel_map_id[did_r] = tel_r
            if nm_r and nm_r not in tel_map_name: tel_map_name[nm_r] = tel_r
        if web_r:
            if did_r and did_r not in web_map_id: web_map_id[did_r] = web_r
            if nm_r and nm_r not in web_map_name: web_map_name[nm_r] = web_r
        if rat_r:
            if nm_r and nm_r not in rat_map_name: rat_map_name[nm_r] = rat_r

    df_resto["data_id"] = df_resto.apply(lambda r: str(r.get("data_id", "")).strip() or did_map_name.get(str(r.get("Nama Restaurant", "")).strip().lower(), ""), axis=1)
    df_resto["fasilitas"] = df_resto.apply(lambda r: str(r.get("fasilitas", "")).strip() or fac_map_id.get(str(r.get("data_id", "")).strip()) or fac_map_name.get(str(r.get("Nama Restaurant", "")).strip().lower(), ""), axis=1)
    df_resto["telepon"] = df_resto.apply(lambda r: str(r.get("telepon", "")).strip() or tel_map_id.get(str(r.get("data_id", "")).strip()) or tel_map_name.get(str(r.get("Nama Restaurant", "")).strip().lower(), ""), axis=1)
    df_resto["website"] = df_resto.apply(lambda r: str(r.get("website", "")).strip() or web_map_id.get(str(r.get("data_id", "")).strip()) or web_map_name.get(str(r.get("Nama Restaurant", "")).strip().lower(), ""), axis=1)
    df_resto["rating_rata2"] = df_resto.apply(lambda r: str(r.get("rating_rata2", "")).strip() or rat_map_name.get(str(r.get("Nama Restaurant", "")).strip().lower(), ""), axis=1)

    df_reviews["fasilitas"] = df_reviews.apply(lambda r: str(r.get("fasilitas", "")).strip() or fac_map_id.get(str(r.get("data_id", "")).strip()) or fac_map_name.get(str(r.get("nama_restoran", "")).strip().lower()) or "", axis=1)

    # 2. Proses Setiap Ulasan
    processed_reviews = []
    kalimat_aspek_list = []

    for _, row in df_reviews.iterrows():
        did = str(row.get("data_id", "")).strip()
        nama_restoran = str(row.get("nama_restoran") or row.get("nama_penginapan") or "").strip()
        
        # Ambil profil restoran
        resto_info = resto_map_by_id.get(did, resto_map_by_name.get(nama_restoran.lower(), {}))
        
        place_id = str(resto_info.get("place_id", ""))
        # Sinkronkan persis dengan kolom 'Nama Restaurant' di file master
        if resto_info.get("Nama Restaurant"):
            nama_restoran = str(resto_info.get("Nama Restaurant")).strip()

        alamat = str(resto_info.get("Alamat", row.get("alamat", "")))
        kota = str(resto_info.get("Kota", row.get("kabupaten", "Bali")))
        kecamatan = str(resto_info.get("Kecamatan", ""))
        provinsi = str(resto_info.get("Provinsi", "Bali"))
        no_sertifikat = str(resto_info.get("No Sertifikat", row.get("kata_kunci", "")))
        nama_perusahaan = str(resto_info.get("Nama Perusahaan", ""))
        sumber_halal = str(resto_info.get("Sumber Halal", "BPJPH"))
        kategori = str(resto_info.get("kategori", "Restoran Halal"))
        url = str(resto_info.get("url", ""))

        lat = resto_info.get("lat")
        lon = resto_info.get("lon")
        if (not lat or not lon) and resto_info.get("Koordinat"):
            try:
                coords = str(resto_info["Koordinat"]).split(",")
                lat = float(coords[0].strip())
                lon = float(coords[1].strip())
            except Exception:
                pass

        rating_restoran = resto_info.get("rating_rata2", row.get("rating_restoran") or row.get("rating_penginapan") or "")
        total_ulasan_restoran = resto_info.get("jumlah_ulasan_total", "")
        telepon = str(resto_info.get("telepon") or row.get("telepon") or "").strip()
        website = str(resto_info.get("website") or row.get("website") or "").strip()

        teks_asli = str(row.get("teks_ulasan", "")).strip()
        teks_bersih = clean_review_text(teks_asli)
        bahasa_code = safe_detect_lang(teks_bersih)

        tgl_raw = str(row.get("tanggal", ""))
        dt_obj, iso_str = parse_relative_date(tgl_raw)

        # Seluruh restoran di dataset ini memiliki Sertifikat Halal resmi
        status_halal = "halal"
        bukti_halal = f"No Sertifikat: {no_sertifikat} ({sumber_halal})" if no_sertifikat else "Sertifikasi Halal Resmi"

        review_id = str(row.get("review_id", ""))
        reviewer_total = row.get("reviewer_total_ulasan", 1)
        rating_ulasan = row.get("rating_bintang", "5.0")
        likes = row.get("jumlah_like", 0)

        rev_item = {
            "place_id": place_id,
            "data_id": did,
            "nama_restoran": nama_restoran,
            "nama_perusahaan": nama_perusahaan,
            "no_sertifikat": no_sertifikat,
            "kategori": kategori,
            "alamat": alamat,
            "kecamatan": kecamatan,
            "kabupaten": kota,
            "provinsi": provinsi,
            "sumber_halal": sumber_halal,
            "status_halal": status_halal,
            "bukti_halal": bukti_halal,
            "rating_restoran": rating_restoran,
            "total_ulasan_restoran": total_ulasan_restoran,
            "latitude": lat,
            "longitude": lon,
            "url": url,
            "telepon": telepon,
            "website": website,
            "fasilitas": str(row.get("fasilitas", "")),
            "review_id": review_id,
            "reviewer_total_ulasan": reviewer_total,
            "rating_ulasan": rating_ulasan,
            "rating_bintang": rating_ulasan,
            "tanggal": tgl_raw,
            "iso_date": iso_str,
            "tanggal_iso": dt_obj,
            "ulasan": teks_asli,
            "teks_ulasan": teks_asli,
            "teks_bersih": teks_bersih,
            "teks_terjemahan_google": str(row.get("teks_terjemahan_google", "")),
            "likes": likes,
            "ada_balasan_pemilik": row.get("ada_balasan_pemilik", False),
            "balasan_pemilik": str(row.get("balasan_pemilik", "")),
            "bahasa": bahasa_code,
            "sumber": "Google Maps"
        }
        processed_reviews.append(rev_item)

        # 3. Ekstraksi pasangan Kalimat-Aspek untuk sheet kalimat_aspek
        kalimat_split = [k.strip() for k in re.split(r"[.!?\n]+", teks_bersih) if len(k.strip()) > 3]
        for k_idx, kal in enumerate(kalimat_split):
            for aspek_nama, aspek_cfg in KAMUS_ASPEK.items():
                if re.search(aspek_cfg["regex"], kal, re.IGNORECASE):
                    kalimat_aspek_list.append({
                        "review_id": review_id,
                        "data_id": did,
                        "place_id": place_id,
                        "nama_restoran": nama_restoran,
                        "status_halal": status_halal,
                        "kalimat_ke": k_idx,
                        "kalimat": kal,
                        "aspek": aspek_nama,
                        "aspek_inti_halal": aspek_cfg["inti_halal"],
                        "rating_ulasan": rating_ulasan,
                        "bahasa": bahasa_code,
                        "tanggal_iso": dt_obj
                    })

    df_full = pd.DataFrame(processed_reviews)
    df_kalimat = pd.DataFrame(kalimat_aspek_list)

    # 4. Tambahkan Flag Aspek Halal & Kolom Anotasi untuk file CSV Tahap 3
    df_absa_csv = df_full.copy()
    df_absa_csv["tahun"] = df_absa_csv["tanggal_iso"].apply(lambda d: d.year if pd.notnull(d) else "")
    df_absa_csv["jumlah_kata"] = df_absa_csv["teks_bersih"].apply(lambda t: len(str(t).split()))

    for flag_col, pattern in REGEX_FLAGS_HALAL.items():
        df_absa_csv[flag_col] = df_absa_csv["teks_bersih"].apply(lambda t: bool(re.search(pattern, str(t), re.IGNORECASE)))

    df_absa_csv["flag_ada_aspek_halal"] = (
        df_absa_csv["flag_makanan_halal"] |
        df_absa_csv["flag_fasilitas_ibadah"] |
        df_absa_csv["flag_citra_syariah"] |
        df_absa_csv["flag_bebas_alkohol"]
    )

    label_cols = [
        "label_makanan_halal", "label_fasilitas_ibadah", "label_kesesuaian_syariah",
        "label_kebersihan", "label_pelayanan", "label_suasana_fasilitas",
        "label_lokasi", "label_harga", "anotator_1", "anotator_2"
    ]
    for lc in label_cols:
        df_absa_csv[lc] = ""

    for col in df_absa_csv.columns:
        if df_absa_csv[col].dtype == object:
            df_absa_csv[col] = df_absa_csv[col].apply(clean_csv_cell)
    df_absa_csv.to_csv(OUTPUT_CSV, index=False, encoding="utf-8")
    print(f"✅ File CSV akhir tersimpan: {OUTPUT_CSV}")

    # 5. Persiapan Seluruh 13 Sheet Excel
    # Sheet 1: ulasan
    kolom_ulasan_lengkap = [
        "place_id", "data_id", "nama_restoran", "nama_perusahaan", "no_sertifikat", "kategori",
        "alamat", "kecamatan", "kabupaten", "provinsi", "sumber_halal", "status_halal", "bukti_halal",
        "rating_restoran", "total_ulasan_restoran", "latitude", "longitude", "url", "telepon", "website", "fasilitas",
        "review_id", "reviewer_total_ulasan", "rating_ulasan", "rating_bintang", "tanggal", "iso_date",
        "ulasan", "teks_bersih", "teks_terjemahan_google", "likes", "ada_balasan_pemilik",
        "balasan_pemilik", "bahasa", "sumber"
    ]
    for c in kolom_ulasan_lengkap:
        if c not in df_full.columns:
            df_full[c] = ""
    df_ulasan_sheet = df_full[kolom_ulasan_lengkap].copy()

    # Sheet 2: kalimat_aspek (df_kalimat)

    # Sheet 3: daftar_restoran (Profil 147 Restoran Halal)
    kolom_resto_sheet = [
        "data_id", "place_id", "Nama Restaurant", "Nama Perusahaan", "No Sertifikat",
        "Tanggal Terdaftar", "Kecamatan", "Kota", "Alamat", "Provinsi", "Sumber Halal",
        "kategori", "rating_rata2", "jumlah_ulasan_total", "telepon", "website", "fasilitas", "status_halal", "bukti_halal", "url"
    ]
    df_resto["status_halal"] = "halal"
    df_resto["bukti_halal"] = df_resto["No Sertifikat"].apply(lambda x: f"Sertifikat Halal {x}" if pd.notnull(x) and str(x).strip() else "Sertifikasi Halal")
    for c in kolom_resto_sheet:
        if c not in df_resto.columns:
            df_resto[c] = ""
    df_resto_sheet = df_resto[kolom_resto_sheet].copy()

    # Sheet 4: olah_fasilitas (Matrix Indikator Fasilitas & Layanan Restoran)
    TOP_RESTO_FACILITIES = [
        "Dine-in", "Makan di tempat", "Takeaway", "Bawa pulang", "Delivery", "Pesan antar",
        "Makanan halal", "Wi-Fi", "Wi-Fi gratis", "Tempat duduk outdoor", "Area outdoor",
        "AC", "Toilet", "Tempat ibadah", "Mushola", "Parkir", "Parkir gratis",
        "Ramah anak-anak", "Ramah keluarga", "Pembayaran kartu kredit", "Pembayaran seluler"
    ]

    rows_olah = []
    for _, r in df_resto_sheet.iterrows():
        f_str = str(r.get("fasilitas", "")).strip()
        f_list = [x.strip() for x in f_str.split(",") if x.strip()]
        item = {
            "data_id": r.get("data_id"),
            "nama_restaurant": r.get("Nama Restaurant"),
            "nama_perusahaan": r.get("Nama Perusahaan"),
            "no_sertifikat": r.get("No Sertifikat"),
            "kota": r.get("Kota"),
            "kecamatan": r.get("Kecamatan"),
            "kategori": r.get("kategori"),
            "rating_rata2": r.get("rating_rata2"),
            "jumlah_ulasan_total": r.get("jumlah_ulasan_total"),
            "status_halal": "halal",
            "total_fasilitas": len(f_list),
            "daftar_fasilitas": f_str
        }
        for tf in TOP_RESTO_FACILITIES:
            col_name = tf.lower().replace(" ", "_").replace("-", "_")
            item[col_name] = "Ya" if any(tf.lower() in x.lower() for x in f_list) else "Tidak"
        rows_olah.append(item)

    df_olah_fasilitas = pd.DataFrame(rows_olah)

    # Sheet 5: stat_fasilitas (Rekap ranking fasilitas di Restoran Bali)
    all_facs = set()
    for f in df_resto_sheet["fasilitas"]:
        for item in str(f).split(","):
            if item.strip():
                all_facs.add(item.strip())

    stat_rows = []
    total_r = len(df_resto_sheet)
    df_resto_sheet["rating_num"] = pd.to_numeric(df_resto_sheet["rating_rata2"], errors="coerce")

    for fac in sorted(all_facs):
        mask = df_resto_sheet["fasilitas"].astype(str).str.contains(fac, regex=False)
        sub = df_resto_sheet[mask]
        n = len(sub)
        avg_r = sub["rating_num"].mean()
        stat_rows.append({
            "fasilitas_layanan": fac,
            "jumlah_restoran": n,
            "persentase_restoran": f"{n/total_r*100:.1f}%" if total_r > 0 else "0.0%",
            "rata_rata_rating": round(avg_r, 2) if pd.notnull(avg_r) else 0.0
        })

    if stat_rows:
        df_stat_fasilitas = pd.DataFrame(stat_rows).sort_values(by="jumlah_restoran", ascending=False).reset_index(drop=True)
    else:
        df_stat_fasilitas = pd.DataFrame(columns=["fasilitas_layanan", "jumlah_restoran", "persentase_restoran", "rata_rata_rating"])

    # Sheet 6: kamus_aspek
    kamus_rows = [{"aspek": k, "regex": v["regex"], "inti_halal": v["inti_halal"]} for k, v in KAMUS_ASPEK.items()]
    df_kamus_sheet = pd.DataFrame(kamus_rows)

    # Sheet 7-13: Sheet Statistik
    df_stat_ringkas = pd.DataFrame([
        {"Keterangan": "Jumlah restoran dalam dataset", "Nilai": len(df_resto_sheet)},
        {"Keterangan": "Jumlah ulasan restoran (setelah pembersihan)", "Nilai": len(df_ulasan_sheet)},
        {"Keterangan": "Jumlah pasangan kalimat-aspek", "Nilai": len(df_kalimat)},
        {"Keterangan": "Jumlah jenis fasilitas & layanan terdata", "Nilai": len(df_stat_fasilitas)},
        {"Keterangan": "Rentang tanggal ulasan (awal)", "Nilai": str(df_full['tanggal_iso'].min())},
        {"Keterangan": "Rentang tanggal ulasan (akhir)", "Nilai": str(df_full['tanggal_iso'].max())}
    ])

    df_stat_status_halal = df_resto_sheet["status_halal"].value_counts().reset_index()
    df_stat_status_halal.columns = ["status_halal", "jumlah_restoran"]

    if not df_kalimat.empty:
        df_kalimat_stat = df_kalimat.copy()
        df_kalimat_stat["rating_ulasan"] = pd.to_numeric(df_kalimat_stat["rating_ulasan"], errors="coerce")
        df_stat_aspek = df_kalimat_stat.groupby("aspek").agg(
            jumlah_kalimat=("kalimat", "count"),
            jumlah_ulasan=("review_id", "nunique"),
            rata_rata_rating=("rating_ulasan", "mean")
        ).round(2).reset_index()
    else:
        df_stat_aspek = pd.DataFrame(columns=["aspek", "jumlah_kalimat", "jumlah_ulasan", "rata_rata_rating"])

    df_stat_bahasa = df_ulasan_sheet["bahasa"].value_counts().reset_index()
    df_stat_bahasa.columns = ["bahasa", "jumlah_ulasan"]

    df_stat_rating = df_ulasan_sheet["rating_ulasan"].value_counts().reset_index()
    df_stat_rating.columns = ["rating", "jumlah_ulasan"]

    df_stat_tahun = df_absa_csv["tahun"].value_counts().reset_index()
    df_stat_tahun.columns = ["tahun", "jumlah_ulasan"]

    df_ulasan_num = df_ulasan_sheet.copy()
    df_ulasan_num["rating_ulasan"] = pd.to_numeric(df_ulasan_num["rating_ulasan"], errors="coerce")
    df_stat_restoran = df_ulasan_num.groupby(["data_id", "nama_restoran"]).agg(
        jumlah_ulasan=("review_id", "count"),
        rata_rata_rating=("rating_ulasan", "mean")
    ).round(2).reset_index()

    # Ekspor ke Excel Multi-Sheet
    print(f"⏳ Mengekspor ke Excel: '{OUTPUT_EXCEL}'...")
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        sanitize_for_excel(df_ulasan_sheet).to_excel(writer, sheet_name="ulasan", index=False)
        sanitize_for_excel(df_kalimat).to_excel(writer, sheet_name="kalimat_aspek", index=False)
        sanitize_for_excel(df_resto_sheet).to_excel(writer, sheet_name="daftar_restoran", index=False)
        sanitize_for_excel(df_olah_fasilitas).to_excel(writer, sheet_name="olah_fasilitas", index=False)
        sanitize_for_excel(df_stat_fasilitas).to_excel(writer, sheet_name="stat_fasilitas", index=False)
        df_stat_ringkas.to_excel(writer, sheet_name="stat_ringkas", index=False)
        df_stat_status_halal.to_excel(writer, sheet_name="stat_status_halal", index=False)
        df_stat_aspek.to_excel(writer, sheet_name="stat_aspek", index=False)
        df_stat_bahasa.to_excel(writer, sheet_name="stat_bahasa", index=False)
        df_stat_rating.to_excel(writer, sheet_name="stat_rating", index=False)
        df_stat_tahun.to_excel(writer, sheet_name="stat_tahun", index=False)
        df_stat_restoran.to_excel(writer, sheet_name="stat_restoran", index=False)
        df_kamus_sheet.to_excel(writer, sheet_name="kamus_aspek", index=False)

    print("\n" + "=" * 65)
    print("🎉 SEMUA PROSES BERHASIL SELESAI 100%!")
    print(f"📁 Excel Lengkap: {OUTPUT_EXCEL}")
    print("=" * 65)

if __name__ == "__main__":
    proses_pipeline()
