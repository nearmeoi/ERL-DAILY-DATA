import os
import sys
import ctypes

# Pastikan output konsol selalu UTF-8 agar tidak crash di Windows cmd/powershell
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Pastikan proses terhubung ke active desktop station
try:
    user32 = ctypes.windll.user32
    hDesk = user32.OpenInputDesktop(0, False, 0x01FF)
    if hDesk:
        user32.SetThreadDesktop(hDesk)
except Exception:
    pass

import pyautogui
import config

print("=" * 65)
print("  ALAT UJI DETEKSI GAMBAR REFERENSI (ASSETS DETECTOR)")
print("=" * 65)
print("Petunjuk: Pastikan jendela Citrix atau Axapta terlihat di layar.\n")

assets = [
    ("btn_login.png", "Tombol Log On Citrix Web"),
    ("icon_ax.png", "Ikon Dynamics AX Citrix Web"),
    ("menu_stock.png", "Menu Available Stock di Axapta"),
    ("window_stock.png", "Header Jendela Stok Axapta"),
    ("btn_export.png", "Tombol Export ke Excel"),
    ("btn_save.png", "Tombol Save"),
]

for filename, desc in assets:
    img_path = os.path.join(config.ASSETS_DIR, filename)
    if not os.path.exists(img_path):
        print(f"[-] {filename:<16} : BELUM ADA FILE (Simpan ke assets/)")
        continue

    try:
        # Gunakan confidence dari config (misal 0.7)
        pos = pyautogui.locateCenterOnScreen(img_path, confidence=config.DEFAULT_CONFIDENCE)
        if pos:
            print(f"[V] {filename:<16} : TERDETEKSI di koordinat ({pos.x}, {pos.y}) -> ({desc})")
        else:
            print(f"[X] {filename:<16} : TIDAK TERLIHAT di layar saat ini ({desc})")
    except pyautogui.ImageNotFoundException:
        print(f"[X] {filename:<16} : TIDAK TERLIHAT di layar saat ini ({desc})")
    except Exception as e:
        print(f"[!] {filename:<16} : Error teknis: {e}")

print("=" * 65)
print("Catatan: [X] wajar terjadi jika jendela terkait sedang tertutup/di latar belakang.")
print("=" * 65)
