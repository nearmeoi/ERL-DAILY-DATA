import os
import sys
import time
import json
import ctypes
import pyautogui

if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Sambungkan thread ke active desktop Windows
try:
    user32 = ctypes.windll.user32
    hDesk = user32.OpenInputDesktop(0, False, 0x01FF)
    if hDesk:
        user32.SetThreadDesktop(hDesk)
except Exception:
    pass

INPUT_FILE = os.path.join(os.path.dirname(__file__), "rekaman_alur.json")

# Safety
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05

print("=" * 65)
print("             PEMUTAR REKAMAN ALUR (MACRO PLAYER)")
print("=" * 65)

if not os.path.exists(INPUT_FILE):
    print(f"[ERROR] File rekaman '{os.path.basename(INPUT_FILE)}' belum ada!")
    print("Silakan rekam alur kerja Anda terlebih dahulu lewat REKAM_ALUR.bat.")
    sys.exit(1)

with open(INPUT_FILE, "r", encoding="utf-8") as f:
    events = json.load(f)

if not events:
    print("[ERROR] File rekaman kosong!")
    sys.exit(1)

print(f"Total aksi yang akan dijalankan: {len(events)} aksi.")
print("Perhatian: Jangan sentuh mouse atau keyboard saat proses berjalan.")
print("REM DARURAT: Geser kursor ke sudut kiri-atas layar jika ingin STOP paksa.")
print("=" * 65)

for i in range(3, 0, -1):
    print(f"Memulai dalam {i} detik...")
    time.sleep(1)

print("\n🚀 [SEDANG BERJALAN...] Menjalankan rekaman alur...")

last_time = 0
for idx, ev in enumerate(events, 1):
    ev_time = ev.get("time", 0)
    delay = ev_time - last_time
    if delay > 0:
        time.sleep(delay)
    last_time = ev_time

    ev_type = ev.get("type")
    if ev_type == "click":
        btn = ev.get("button", "left")
        x, y = ev.get("x"), ev.get("y")
        print(f"[{idx}/{len(events)}] Klik {btn} di ({x}, {y})")
        pyautogui.click(x, y, button=btn)

    elif ev_type == "scroll":
        dy = ev.get("dy", 0)
        x, y = ev.get("x"), ev.get("y")
        pyautogui.scroll(int(dy * 120), x=x, y=y)

    elif ev_type == "keydown":
        k = ev.get("key", {})
        val = k.get("value")
        if val:
            try:
                pyautogui.keyDown(val)
            except Exception:
                pass

    elif ev_type == "keyup":
        k = ev.get("key", {})
        val = k.get("value")
        if val:
            try:
                pyautogui.keyUp(val)
            except Exception:
                pass

print("\n" + "=" * 65)
print("🎉 [SELESAI] Semua aksi dalam rekaman telah berhasil dijalankan!")
print("=" * 65)
