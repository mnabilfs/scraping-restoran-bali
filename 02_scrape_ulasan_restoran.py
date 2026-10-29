"""
TAHAP 2: Scraping Ulasan Restoran Halal Bali dari Google Maps
Input: data/restoran_halal_bali.csv (Master 147 Restoran Halal - Bersih & Tidak Dimutasi)
Output: data/raw_ulasan_checkpoint.csv

Fitur Unggulan:
  1. Auto Search & Cache URL:
     Jika URL Google Maps belum ada di dataset master, skrip mencari restoran secara otomatis di Google Maps,
     membuka profil restoran, dan mengambil seluruh metadata yang diperlukan.
  2. Preservasi Master CSV Murni:
     Dataset master 'data/restoran_halal_bali.csv' (147 baris, 10 kolom) TIDAK dimutasi atau diubah.
     Semua kolom hasil scraping (telepon, website, fasilitas, data_id, dll.) disimpan khusus di file checkpoint.
  3. Ekstraksi Fasilitas & Layanan Presisi:
     Mengambil opsi layanan (Dine-in, Takeaway, Delivery) dan fasilitas aktual dari container 'div.E0DTEd' /
     'button.XJ8h0e' (pane.attributes.expand) tanpa teks sampah UI.
  4. Ekstraksi Balasan Pemilik:
     Mengambil teks tanggapan resmi dari pihak restoran beserta status responnya.
  5. Multi-Action Infinite Scroll Tangguh:
     Mengombinasikan scroll_into_view, mouse wheel, dan DOM scroll dengan mekanisme anti-stuck (nudge)
     sehingga dapat mengambil ulasan secara tuntas sesuai target --max.
  6. Auto-Resume Cerdas:
     Melanjutkan langsung dari restoran terakhir di file checkpoint ulasan tanpa mengulang dari awal.
"""

import sys
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

import os
import re
import time
import hashlib
import csv
import argparse
import urllib.parse
from pathlib import Path
import pandas as pd
from tqdm import tqdm
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

INPUT_RESTO_FILE = DATA_DIR / "restoran_halal_bali.csv"
OUTPUT_REVIEW_FILE = DATA_DIR / "raw_ulasan_checkpoint.csv"
USER_DATA_DIR = DATA_DIR / "browser_profile"
LOGIN_FLAG_FILE = USER_DATA_DIR / "login_completed.flag"

# Kolom standar dataset ulasan restoran
CSV_COLUMNS = [
    "review_id", "data_id", "nama_restoran", "kabupaten", "kelompok_sampel",
    "kata_kunci", "fasilitas", "teks_ulasan", "teks_terjemahan_google", "rating_restoran", "rating_bintang",
    "tanggal", "jumlah_like", "ada_balasan_pemilik", "balasan_pemilik", "sumber", "reviewer_total_ulasan",
    "telepon", "website"
]

def hash_review_id(text: str, author: str, date_str: str) -> str:
    seed = f"{author}_{date_str}_{text}".encode("utf-8")
    return hashlib.sha256(seed).hexdigest()[:24]

def clean_csv_cell(val) -> str:
    if val is None or pd.isna(val):
        return ""
    txt = str(val)
    if txt.strip().lower() in ["nan", "none", "null"]:
        return ""
    txt = re.sub(r"[\r\n]+", " ", txt)
    txt = txt.replace("\t", " ")
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt

def parse_url_coords_and_id(url: str):
    lat, lon, data_id, place_id = None, None, None, None
    if not url:
        return lat, lon, data_id, place_id

    coord_match = re.search(r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)", url) or re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", url)
    if coord_match:
        lat = float(coord_match.group(1))
        lon = float(coord_match.group(2))

    pid_match = re.search(r"!19s(ChIJ[0-9a-zA-Z_-]+)", url)
    if pid_match:
        place_id = pid_match.group(1)

    data_id_match = re.search(r"!1s(0x[0-9a-fA-F]+:0x[0-9a-fA-F]+)", url) or re.search(r"(0x[0-9a-fA-F]+:0x[0-9a-fA-F]+)", url)
    if data_id_match:
        data_id = data_id_match.group(1)

    return lat, lon, data_id, place_id

# Ekstraksi nomor telepon dan tautan website resmi restoran
JS_EXTRACT_CONTACT_INFO = r"""
() => {
    let telepon = '';
    const phoneEl = document.querySelector('button[data-item-id*="phone"], button[aria-label*="Telepon:" i], button[aria-label*="Phone:" i], a[href^="tel:"], button[data-tooltip*="telepon" i], button[data-tooltip*="phone" i]');
    if (phoneEl) {
        const aria = phoneEl.getAttribute('aria-label') || '';
        telepon = aria.replace(/^(Telepon|Phone):\s*/i, '').trim() || phoneEl.innerText.trim();
        telepon = telepon.replace(/^tel:\+?/i, '').trim();
    }

    let website = '';
    const webEl = document.querySelector('a[data-item-id="authority"], a[aria-label*="Situs Web" i], a[aria-label*="Website" i], a[data-tooltip*="situs web" i], a[data-tooltip*="website" i]');
    if (webEl) {
        let href = webEl.getAttribute('href') || '';
        const m = href.match(/[?&]q=([^&]+)/);
        if (m) {
            try { href = decodeURIComponent(m[1]); } catch(e){}
        }
        website = href;
    }

    return { telepon, website };
}
"""

# Ekstraksi fasilitas & layanan presisi:
# Hanya mengambil dari elemen spesifik (E0DTEd, LTs0Rc, dan tombol expand pane.attributes.expand)
JS_CHECK_AND_EXPAND_FACILITIES = r"""
() => {
    const facilities = new Set();

    function cleanItem(raw) {
        if (!raw || typeof raw !== 'string') return '';
        const lines = raw.split(/[\r\n]+/).map(s => s.trim()).filter(Boolean);
        for (let l of lines) {
            const c = l.replace(/^[\s·•✔️✔\-\uE000-\uF8FF]+\s*/gu, '').trim();
            if (c.length >= 2 && c.length <= 50 && !c.includes('http') && !c.includes('.com') && !c.includes('.id')) {
                const lower = c.toLowerCase();
                const junk = ['lihat foto', 'foto', 'restoran', 'ulasan', 'menu', 'rute', 'simpan', 'bagikan', 'tentang', 'histori maps', 'klaim', 'edit'];
                if (!junk.some(j => lower === j || lower.startsWith(j + ':'))) {
                    return c;
                }
            }
        }
        return '';
    }

    // 1. Ambil opsi layanan inline jika ada (div.E0DTEd atau div.LTs0Rc)
    document.querySelectorAll('div.E0DTEd div.LTs0Rc, div.LTs0Rc[role="group"]').forEach(el => {
        const aria = el.getAttribute('aria-label') || '';
        const txt = el.innerText || '';
        const c1 = cleanItem(aria);
        if (c1) facilities.add(c1);
        else {
            const c2 = cleanItem(txt);
            if (c2) facilities.add(c2);
        }
    });

    // 2. Cek apakah ada tombol ekspansi atribut fasilitas (button.XJ8h0e / pane.attributes.expand)
    const expandBtn = document.querySelector('button[jsaction*="pane.attributes.expand"], button.XJ8h0e');
    let wasExpanded = false;
    if (expandBtn) {
        try {
            expandBtn.click();
            wasExpanded = true;
        } catch(e) {}
    }

    return {
        wasExpanded,
        inlineItems: Array.from(facilities)
    };
}
"""

JS_COLLECT_EXPANDED_FACILITIES_AND_CLOSE = r"""
() => {
    const facilities = new Set();

    function cleanItem(raw) {
        if (!raw || typeof raw !== 'string') return '';
        const lines = raw.split(/[\r\n]+/).map(s => s.trim()).filter(Boolean);
        for (let l of lines) {
            const c = l.replace(/^[\s·•✔️✔\-\uE000-\uF8FF]+\s*/gu, '').trim();
            if (c.length >= 2 && c.length <= 50 && !c.includes('http') && !c.includes('.com') && !c.includes('.id')) {
                const lower = c.toLowerCase();
                const junk = ['lihat foto', 'foto', 'restoran', 'ulasan', 'menu', 'rute', 'simpan', 'bagikan', 'tentang', 'histori maps', 'klaim', 'edit'];
                if (!junk.some(j => lower === j || lower.startsWith(j + ':'))) {
                    return c;
                }
            }
        }
        return '';
    }

    // Ambil item dari panel atribut fasilitas yang terbuka
    document.querySelectorAll('ul.ZQ6we li, div.iP2tbe, div.LTs0Rc[role="group"]').forEach(el => {
        const aria = el.getAttribute('aria-label') || '';
        const txt = el.innerText || '';
        const c1 = cleanItem(aria);
        if (c1) facilities.add(c1);
        else {
            const c2 = cleanItem(txt);
            if (c2) facilities.add(c2);
        }
    });

    // Tutup panel atribut dengan tombol kembali
    const backBtn = document.querySelector('button[aria-label*="Kembali"], button[aria-label*="Back"], button[aria-label*="Tutup"], button[aria-label*="Close"], button[jsaction*="back"]');
    if (backBtn) {
        try { backBtn.click(); } catch(e) {}
    }

    return Array.from(facilities);
}
"""

JS_EXPAND_ALL_REVIEWS = r"""
() => {
    // Klik tombol 'Lainnya' pada ulasan dan balasan pemilik
    document.querySelectorAll("button.w8nwRe, button[aria-label*='Lainnya'], button[aria-label*='more']").forEach(b => {
        try { b.click(); } catch(e) {}
    });
}
"""

JS_EXTRACT_REVIEWS = r"""
() => {
    const cards = Array.from(document.querySelectorAll('div.jftiEf'));
    return cards.map(c => {
        const raw_id = c.getAttribute('data-review-id') || '';
        const authorEl = c.querySelector('.d4r55');
        const author = authorEl ? authorEl.innerText.trim() : 'Anonim';
        const infoEl = c.querySelector('.RfnDt');
        const info = infoEl ? infoEl.innerText.trim() : '';
        const starEl = c.querySelector('.kvMYJc');
        const star = starEl ? (starEl.getAttribute('aria-label') || starEl.innerText) : '5';
        const dateEl = c.querySelector('.rsqaWe');
        const date = dateEl ? dateEl.innerText.trim() : '';

        // Teks ulasan
        const textEl = c.querySelector('.wiI7Cc, .MyEned');
        const text = textEl ? textEl.innerText.trim() : '';

        // Deteksi & Ekstraksi Balasan Pemilik
        let has_reply = false;
        let reply_text = '';
        const replyEl = c.querySelector('div.CDe7pd, div.wiI7Cc.kK2LRe, div.kK2LRe, div[aria-label*="Tanggapan"]');
        if (replyEl) {
            has_reply = true;
            const replyContent = replyEl.querySelector('.wiI7Cc') || replyEl;
            reply_text = replyContent.innerText.trim();
        } else {
            const allSpans = Array.from(c.querySelectorAll('div, span'));
            for (let el of allSpans) {
                const t = (el.innerText || '').trim();
                if (/^Tanggapan dari pemilik/i.test(t)) {
                    has_reply = true;
                    const nextEl = el.parentElement ? el.parentElement.querySelector('.wiI7Cc') : null;
                    reply_text = nextEl ? nextEl.innerText.trim() : t.replace(/^Tanggapan dari pemilik\s*/i, '').trim();
                    break;
                }
            }
        }

        // Ekstraksi Jumlah Like (Jempol)
        let likes = 0;
        const likeBtn = c.querySelector('button[aria-label*="Suka"], button[aria-label*="Like"], button[aria-label*="like"]');
        if (likeBtn) {
            const likeAria = likeBtn.getAttribute('aria-label') || '';
            const match = likeAria.match(/(\d+)/);
            if (match) {
                likes = parseInt(match[1]);
            } else {
                const btnTxt = likeBtn.innerText.trim();
                const numMatch = btnTxt.match(/(\d+)/);
                if (numMatch) likes = parseInt(numMatch[1]);
            }
        }

        return {
            raw_id,
            author,
            info,
            star,
            date,
            text,
            has_reply,
            reply_text,
            likes
        };
    });
}
"""

def load_clean_resto_master():
    if not INPUT_RESTO_FILE.exists():
        print(f"❌ File master restoran tidak ditemukan di: {INPUT_RESTO_FILE}")
        sys.exit(1)
    df = pd.read_csv(INPUT_RESTO_FILE).fillna("")
    return df

def switch_to_reviews_tab(page) -> bool:
    # 1. Coba klik tab Ulasan / Reviews
    for sel in [
        "button[role='tab']:has-text('Ulasan')",
        "button[role='tab']:has-text('Review')",
        "button:has-text('Ulasan')",
        "button:has-text('Reviews')"
    ]:
        try:
            loc = page.locator(sel).first
            if loc.count() > 0:
                loc.click(timeout=2500)
                return True
        except Exception:
            pass

    # 2. Coba klik angka/link ulasan di header (div.F7nice atau tombol ulasan lainnya)
    try:
        cnt_btn = page.locator('div.F7nice button, div.F7nice span:has-text("("), button[aria-label*="ulasan" i]').first
        if cnt_btn.count() > 0:
            cnt_btn.click(timeout=2500)
            return True
    except Exception:
        pass

    # 3. Fallback via evaluasi DOM JavaScript
    return bool(page.evaluate("""() => {
        const tabs = Array.from(document.querySelectorAll('[role="tab"]'));
        const t = tabs.find(x => (x.innerText || '').toLowerCase().includes('ulasan') || (x.innerText || '').toLowerCase().includes('review'));
        if (t) { t.click(); return true; }
        
        const btns = Array.from(document.querySelectorAll('button'));
        const b = btns.find(x => {
            const txt = (x.innerText || '').toLowerCase();
            const aria = (x.getAttribute('aria-label') || '').toLowerCase();
            return (txt.includes('ulasan') || aria.includes('ulasan') || txt.includes('reviews') || aria.includes('reviews')) && !txt.includes('tulis') && !aria.includes('tulis');
        });
        if (b) { b.click(); return true; }
        return false;
    }"""))

def main():
    parser = argparse.ArgumentParser(description="Scraper Ulasan Google Maps untuk Restoran Halal Bali")
    parser.add_argument("--max", type=int, default=100, help="Jumlah ulasan maksimal per restoran (default: 100)")
    parser.add_argument("--start", type=int, default=1, help="Nomor urut restoran untuk memulai (1-indexed)")
    parser.add_argument("--hotel", "--resto", "--nama", dest="resto_filter", type=str, default="", help="Filter nama restoran tertentu")
    parser.add_argument("--retry-incomplete", action="store_true", help="Menyapu ulang restoran yang ulasannya belum mencapai kuota")
    parser.add_argument("--headless", action="store_true", help="Jalankan di latar belakang tanpa jendela visual")
    args = parser.parse_args()

    max_ulasan_per_resto = args.max
    filter_nama = args.resto_filter.strip().lower()

    print("=" * 65)
    print("🚀 MEMULAI TAHAP 2: SCRAPING ULASAN RESTORAN HALAL BALI")
    print(f"🎯 Target ulasan per restoran : {max_ulasan_per_resto if max_ulasan_per_resto > 0 else 'Semua'}")
    print("=" * 65)

    df_resto = load_clean_resto_master()
    total_master_resto = len(df_resto)
    print(f"📋 Total restoran dalam dataset master: {total_master_resto}")

    # Baca ulasan lama dari file checkpoint jika ada
    existing_review_ids = set()
    resto_review_counts = {}
    last_processed_idx = -1

    if OUTPUT_REVIEW_FILE.exists():
        try:
            df_old = pd.read_csv(OUTPUT_REVIEW_FILE, low_memory=False)
            if "review_id" in df_old.columns:
                existing_review_ids = set(df_old["review_id"].dropna().astype(str).unique())
            if "data_id" in df_old.columns:
                resto_review_counts = df_old["data_id"].dropna().astype(str).value_counts().to_dict()

            # Deteksi restoran terakhir di checkpoint
            if len(df_old) > 0:
                col_name = "nama_restoran" if "nama_restoran" in df_old.columns else ("nama_penginapan" if "nama_penginapan" in df_old.columns else None)
                if col_name:
                    last_resto_name = str(df_old.iloc[-1][col_name]).strip().lower()
                    for idx, r in df_resto.iterrows():
                        nama_r = str(r.get("Nama Restaurant", "")).strip().lower()
                        if nama_r == last_resto_name:
                            last_processed_idx = idx
                            break
            print(f"📋 Checkpoint lama: {len(df_old)} ulasan ditemukan ({len(resto_review_counts)} restoran).")
        except Exception as e:
            print(f"⚠️ Gagal membaca file checkpoint ulasan ({e}), memulai baru.")

    # Tentukan titik mulai
    start_idx = 0
    if args.start > 1:
        start_idx = args.start - 1
        print(f"📍 Parameter --start aktif: Memulai langsung dari restoran ke-{args.start}")
    elif last_processed_idx >= 0 and not args.retry_incomplete and not filter_nama:
        start_idx = last_processed_idx + 1
        print(f"📍 Auto-resume aktif: Melanjutkan dari restoran ke-{start_idx + 1} (setelah '{df_resto.iloc[last_processed_idx].get('Nama Restaurant')}')")

    restos_to_process = []
    for idx in range(start_idx, total_master_resto):
        r_item = df_resto.iloc[idx].to_dict()
        nama_r = str(r_item.get("Nama Restaurant", "")).strip()
        if filter_nama and filter_nama not in nama_r.lower():
            continue
        restos_to_process.append((idx, r_item))

    print(f"🎯 Restoran yang akan diproses pada sesi ini: {len(restos_to_process)}")
    print("=" * 65)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            channel="chrome",
            headless=args.headless,
            args=[
                "--lang=id-ID,id",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox"
            ],
            viewport={"width": 1400, "height": 900},
            locale="id-ID"
        )

        page = context.pages[0] if context.pages else context.new_page()

        # Blokir file gambar & media agar hemat bandwidth dan loading cepat
        def intercept_heavy(route):
            if route.request.resource_type in ["image", "media"]:
                route.abort()
            else:
                route.continue_()

        page.route("**/*", intercept_heavy)

        total_new_reviews = 0

        for r_counter, (global_idx, resto) in enumerate(restos_to_process, start=1):
            nama_resto = str(resto.get("Nama Restaurant", "")).strip()
            alamat = str(resto.get("Alamat", "")).strip()
            kota = str(resto.get("Kota", "") or resto.get("Kecamatan", "") or "Bali").strip()
            no_sertifikat = str(resto.get("No Sertifikat", "")).strip()
            sumber_halal = str(resto.get("Sumber Halal", "BPJPH")).strip()
            url = str(resto.get("url", "")).strip()
            data_id = str(resto.get("data_id", "")).strip()

            resto_display_num = global_idx + 1
            print(f"\n🍽️ [Restoran {resto_display_num}/{total_master_resto}] {nama_resto}")

            # 1. Navigasi ke Google Maps (Direct URL atau Auto Search)
            target_nav_url = ""
            if url and url.startswith("http"):
                target_nav_url = url
            else:
                clean_alm = re.sub(r"\s+", " ", str(alamat)).strip()
                if clean_alm and clean_alm.lower() not in ["nan", "none", ""]:
                    q_search = f"{nama_resto} {clean_alm[:100]}"
                elif kota and kota.lower() not in ["nan", "none", ""]:
                    q_search = f"{nama_resto} {kota} Bali"
                else:
                    q_search = f"{nama_resto} Bali"
                target_nav_url = f"https://www.google.com/maps/search/{urllib.parse.quote(q_search)}?hl=id"
                print(f"   🔍 Mencari tempat: '{q_search}'...")

            try:
                page.goto(target_nav_url, timeout=35000, wait_until="domcontentloaded")
                time.sleep(3.0)

                # Deteksi jika ada kartu pencarian (jika tidak langsung redirect ke profil restoran)
                if "/maps/place/" not in page.url:
                    feed_cards = page.locator("div[role='feed'] > div:has(a[href*='/maps/place/'])")
                    if feed_cards.count() > 0:
                        first_card_link = feed_cards.first.locator("a[href*='/maps/place/']").first
                        first_card_link.click()
                        time.sleep(3.0)

                # Update data_id & koordinat jika ada
                current_url = page.url
                lat_c, lon_c, new_data_id, new_place_id = parse_url_coords_and_id(current_url)

                if new_data_id:
                    data_id = new_data_id
                elif not data_id:
                    data_id = f"resto_{global_idx+1}_{hashlib.md5(nama_resto.encode()).hexdigest()[:8]}"

                # Ambil rating & estimasi total ulasan dari header restoran
                live_rating = ""
                total_est_reviews = 0
                try:
                    r_box = page.locator("div.F7nice").first.inner_text().strip()
                    m_rat = re.search(r"(\d+[.,]\d+)", r_box)
                    if m_rat:
                        live_rating = m_rat.group(1).replace(",", ".")
                    m_cnt = re.search(r"\(([\d.]+)\)", r_box)
                    if m_cnt:
                        total_est_reviews = int(m_cnt.group(1).replace(".", ""))
                except Exception:
                    pass

                # Ambil Telepon & Website restoran
                live_telepon = ""
                live_website = ""
                try:
                    contact_data = page.evaluate(JS_EXTRACT_CONTACT_INFO)
                    if contact_data:
                        live_telepon = clean_csv_cell(contact_data.get("telepon", ""))
                        live_website = clean_csv_cell(contact_data.get("website", ""))
                except Exception:
                    pass

                contact_info_log = []
                if live_telepon:
                    contact_info_log.append(f"📞 {live_telepon}")
                if live_website:
                    contact_info_log.append(f"🌐 {live_website}")
                if contact_info_log:
                    print(f"   📞 Kontak: {' | '.join(contact_info_log)}")

                # 2. Ekstraksi Fasilitas & Layanan Presisi
                fac_set = set()
                fasilitas_str = ""
                try:
                    init_fac = page.evaluate(JS_CHECK_AND_EXPAND_FACILITIES)
                    if init_fac and "inlineItems" in init_fac:
                        for item in init_fac["inlineItems"]:
                            fac_set.add(item)

                    if init_fac and init_fac.get("wasExpanded"):
                        time.sleep(1.2)
                        exp_fac = page.evaluate(JS_COLLECT_EXPANDED_FACILITIES_AND_CLOSE)
                        if exp_fac:
                            for item in exp_fac:
                                fac_set.add(item)
                        time.sleep(0.8)

                    if fac_set:
                        fasilitas_str = ", ".join(sorted(fac_set))
                        print(f"   📋 Fasilitas & Layanan ({len(fac_set)}): {fasilitas_str[:70]}...")
                except Exception as e:
                    pass

                # Hitung target kuota ulasan
                if max_ulasan_per_resto and max_ulasan_per_resto > 0:
                    target_quota = min(total_est_reviews, max_ulasan_per_resto) if total_est_reviews > 0 else max_ulasan_per_resto
                else:
                    target_quota = total_est_reviews if total_est_reviews > 0 else 100

                # 3. Beralih ke Tab Ulasan
                switch_to_reviews_tab(page)

                # Tunggu kartu ulasan muncul
                try:
                    page.wait_for_selector("div.jftiEf", timeout=7000)
                except Exception:
                    pass
                time.sleep(1.0)

                # Progress bar per restoran
                curr_card_count = page.locator("div.jftiEf").count()
                pbar = tqdm(
                    total=target_quota,
                    desc="Scraping Review",
                    unit="ulasan",
                    bar_format="{desc}: {percentage:3.0f}% |{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]",
                    leave=True
                )
                if curr_card_count > 0:
                    pbar.update(min(curr_card_count, target_quota))

                # Multi-Action Infinite Scroll Tangguh
                max_scroll_attempts = max(30, (target_quota // 4) + 20)
                same_count_streak = 0
                last_count = curr_card_count

                for scroll_step in range(max_scroll_attempts):
                    if curr_card_count >= target_quota:
                        break

                    # Action 1: Bawa kartu review terakhir ke viewport
                    try:
                        page.locator("div.jftiEf").last.scroll_into_view_if_needed(timeout=1200)
                    except Exception:
                        pass

                    # Action 2: Kirim simulated mouse wheel
                    page.mouse.move(250, 600)
                    page.mouse.wheel(0, 3000)

                    # Action 3: Container DOM scroll
                    page.evaluate("""() => {
                        const container = document.querySelector('div.m6QErb.DxyBCb.kA9KIf.dS8AEf') || 
                                          document.querySelector('div[role="main"] .m6QErb') ||
                                          document.querySelector('.m6QErb');
                        if (container) {
                            container.scrollBy(0, 3000);
                            container.scrollTop = container.scrollHeight;
                        }
                    }""")

                    time.sleep(1.4)

                    curr_card_count = page.locator("div.jftiEf").count()
                    diff = min(curr_card_count, target_quota) - pbar.n
                    if diff > 0:
                        pbar.update(diff)

                    if curr_card_count == last_count:
                        same_count_streak += 1
                        if same_count_streak >= 8:
                            break
                        # Nudge: scroll sedikit ke atas lalu ke bawah untuk memicu trigger virtualizer
                        page.mouse.wheel(0, -600)
                        time.sleep(0.4)
                        page.mouse.wheel(0, 2500)
                    else:
                        same_count_streak = 0
                        last_count = curr_card_count

                pbar.close()

                # 4. Ekspansi teks ulasan & balasan pemilik
                try:
                    page.evaluate(JS_EXPAND_ALL_REVIEWS)
                except Exception:
                    pass
                time.sleep(0.5)

                # 5. Ekstraksi seluruh kartu ulasan
                raw_items = page.evaluate(JS_EXTRACT_REVIEWS)
                resto_added = 0
                resto_replies = 0
                resto_new_records = []

                for item in raw_items[:target_quota]:
                    teks_ulasan = clean_csv_cell(item["text"] or nama_resto)
                    tanggal_str = clean_csv_cell(item["date"])
                    author = clean_csv_cell(item["author"])
                    has_reply = bool(item["has_reply"])
                    reply_text = clean_csv_cell(item["reply_text"]) if has_reply else ""

                    if has_reply:
                        resto_replies += 1

                    star_match = re.search(r"(\d+(?:[.,]\d+)?)", item["star"])
                    rating_bintang = f"{float(star_match.group(1).replace(',', '.')):.1f}" if star_match else "5.0"

                    rev_match = re.search(r"(\d+)", item["info"].replace(".", ""))
                    reviewer_total = int(rev_match.group(1)) if rev_match else 1

                    like_raw = item.get("likes", 0)
                    likes_count = int(like_raw) if isinstance(like_raw, (int, float)) else 0

                    rev_id = item["raw_id"] if item["raw_id"] else hash_review_id(teks_ulasan, author, tanggal_str)
                    if rev_id in existing_review_ids:
                        continue

                    record = {
                        "review_id": rev_id,
                        "data_id": data_id,
                        "nama_restoran": nama_resto,
                        "kabupaten": kota,
                        "kelompok_sampel": "halal",
                        "kata_kunci": no_sertifikat or "Restoran Halal Bali",
                        "fasilitas": fasilitas_str,
                        "teks_ulasan": teks_ulasan,
                        "teks_terjemahan_google": "",
                        "rating_restoran": live_rating,
                        "rating_bintang": rating_bintang,
                        "tanggal": tanggal_str,
                        "jumlah_like": likes_count,
                        "ada_balasan_pemilik": has_reply,
                        "balasan_pemilik": reply_text,
                        "sumber": "Google Maps",
                        "reviewer_total_ulasan": reviewer_total,
                        "telepon": live_telepon,
                        "website": live_website
                    }

                    existing_review_ids.add(rev_id)
                    resto_new_records.append(record)
                    resto_added += 1

                # Tulis append instan ke checkpoint CSV
                if resto_new_records:
                    write_header = not OUTPUT_REVIEW_FILE.exists()
                    with open(OUTPUT_REVIEW_FILE, mode="a", newline="", encoding="utf-8") as f_out:
                        writer = csv.DictWriter(f_out, fieldnames=CSV_COLUMNS)
                        if write_header:
                            writer.writeheader()
                        writer.writerows(resto_new_records)

                total_new_reviews += resto_added
                print(f"   ✅ Selesai: {resto_added} ulasan baru tersimpan ({resto_replies} memiliki balasan pemilik).")

            except Exception as err:
                print(f"   ❌ Gagal memproses '{nama_resto}': {err}")
                continue

        context.close()

    print("\n" + "=" * 65)
    print(f"🎉 SCRAPING SELESAI!")
    print(f"📊 Total ulasan baru terkumpul: {total_new_reviews}")
    print(f"💾 Checkpoint ulasan tersimpan di: {OUTPUT_REVIEW_FILE}")
    print("=" * 65)

if __name__ == "__main__":
    main()
