import sys
import os
import subprocess

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

def main():
    print("""
=============================================================
   PIPELINE SCRAPING & DATASET ABSA RESTORAN HALAL BALI
=============================================================
1. Jalankan Tahap 2: Scraping Ulasan Restoran dari Google Maps
2. Jalankan Tahap 3: Pemrosesan ABSA & Export Excel Multi-Sheet
3. Jalankan Semua (Tahap 2 -> Tahap 3)
4. Buka / Cek Sesi Login Akun Google (data/browser_profile)
0. Keluar
=============================================================
    """)
    choice = input("Pilih menu (0-4): ").strip()

    script_02 = "02_scrape_ulasan_restoran.py" if os.path.exists("02_scrape_ulasan_restoran.py") else "02_scrape_ulasan_penginapan.py"
    script_03 = "03_proses_dataset_absa.py"

    if choice == "1":
        subprocess.run([sys.executable, script_02])
    elif choice == "2":
        subprocess.run([sys.executable, script_03])
    elif choice == "3":
        print(f"\n--- Menjalankan Tahap 2: Scraping Ulasan Restoran ({script_02}) ---")
        subprocess.run([sys.executable, script_02])
        print(f"\n--- Menjalankan Tahap 3: Pemrosesan & Export Excel ({script_03}) ---")
        subprocess.run([sys.executable, script_03])
        print("\n🎉 Seluruh alur selesai!")
    elif choice == "4":
        subprocess.run([sys.executable, "login_google.py"])
    elif choice == "0":
        print("Selesai.")
    else:
        print("Pilihan tidak valid.")

if __name__ == "__main__":
    main()
