import os
import sys
import time
import json
import ctypes
from pynput import mouse, keyboard

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

OUTPUT_FILE = os.path.join(os.path.dirname(__file__), "rekaman_alur.json")

events = []
start_time = None
is_recording = False

print("=" * 65)
print("             PEREKAM ALUR OTOMATIS (MACRO RECORDER)")
print("=" * 65)
print("Petunjuk:")
print("1. Tekan tombol [F9] untuk MULAI MEREKAM.")
print("2. Lakukan alur kerja Anda seperti biasa:")
print("   - Klik Edge, Login Citrix, Klik Dynamics AX, Tunggu buka,")
print("   - Klik menu stok, scroll ke bawah, klik Export, simpan file.")
print("3. Setelah selesai, tekan tombol [ESC] untuk BERHENTI & MENYIMPAN.")
print("=" * 65)
print("\nMenunggu Anda menekan [F9] untuk mulai...")


def on_click(x, y, button, pressed):
    global is_recording, start_time, events
    if not is_recording or not pressed:
        return
    
    elapsed = time.time() - start_time
    btn_str = "left" if button == mouse.Button.left else "right"
    events.append({
        "type": "click",
        "time": round(elapsed, 3),
        "x": int(x),
        "y": int(y),
        "button": btn_str
    })
    print(f"  [KLIK {btn_str.upper()}] di ({x}, {y}) pada detik ke-{elapsed:.1f}s")


def on_scroll(x, y, dx, dy):
    global is_recording, start_time, events
    if not is_recording:
        return
    elapsed = time.time() - start_time
    events.append({
        "type": "scroll",
        "time": round(elapsed, 3),
        "x": int(x),
        "y": int(y),
        "dx": dx,
        "dy": dy
    })


def on_press(key):
    global is_recording, start_time, events

    # Tombol F9 untuk mulai rekam
    if key == keyboard.Key.f9 and not is_recording:
        is_recording = True
        start_time = time.time()
        events = []
        print("\n" + "=" * 65)
        print("🔴 [SEDANG MEREKAM...] Lakukan semua aksi Anda sekarang!")
        print("   Tekan [ESC] di keyboard jika sudah selesai.")
        print("=" * 65)
        return

    # Tombol ESC untuk stop dan simpan
    if key == keyboard.Key.esc and is_recording:
        print("\n" + "=" * 65)
        print("⏹️  [REKAMAN DIHENTIKAN]")
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(events, f, indent=2)
        print(f"[BERHASIL] Tersimpan {len(events)} aksi ke: {os.path.basename(OUTPUT_FILE)}")
        print("Sekarang Anda bisa memutarnya kapan saja dengan PUTAR_REKAMAN.bat!")
        print("=" * 65)
        # Hentikan listener
        return False

    if is_recording:
        elapsed = time.time() - start_time
        key_data = None
        if hasattr(key, "char") and key.char is not None:
            key_data = {"key_type": "char", "value": key.char}
        else:
            key_data = {"key_type": "special", "value": str(key).replace("Key.", "")}

        events.append({
            "type": "keydown",
            "time": round(elapsed, 3),
            "key": key_data
        })


def on_release(key):
    global is_recording, start_time, events
    if not is_recording:
        return

    elapsed = time.time() - start_time
    key_data = None
    if hasattr(key, "char") and key.char is not None:
        key_data = {"key_type": "char", "value": key.char}
    else:
        key_data = {"key_type": "special", "value": str(key).replace("Key.", "")}

    events.append({
        "type": "keyup",
        "time": round(elapsed, 3),
        "key": key_data
    })


# Start listener
mouse_listener = mouse.Listener(on_click=on_click, on_scroll=on_scroll)
keyboard_listener = keyboard.Listener(on_press=on_press, on_release=on_release)

mouse_listener.start()
keyboard_listener.start()
keyboard_listener.join()
mouse_listener.stop()
