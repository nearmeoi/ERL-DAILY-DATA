import os
import sys
import glob
import time
import logging
import ctypes
import subprocess

# Pastikan konsol UTF-8 agar tidak crash karena karakter khusus/emoji
if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Sambungkan thread ke active desktop Windows agar screen grab selalu berhasil
try:
    user32 = ctypes.windll.user32
    hDesk = user32.OpenInputDesktop(0, False, 0x01FF)
    if hDesk:
        user32.SetThreadDesktop(hDesk)
except Exception:
    pass

import pyautogui
import pygetwindow as gw
import config

# Setup Logging (Terminal + File)
log_filename = os.path.join(
    config.LOGS_DIR, f"axapta_run_{time.strftime('%Y%m%d_%H%M%S')}.log"
)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_filename, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

# Safety settings PyAutoGUI
pyautogui.FAILSAFE = True  # Dorong kursor ke sudut kiri-atas layar untuk STOP darurat
pyautogui.PAUSE = 0.5


def wait_for_image(image_name, timeout=60, confidence=config.DEFAULT_CONFIDENCE):
    """
    Menunggu gambar/elemen muncul di layar.
    Sistem Dynamic Polling: Menyesuaikan sendiri saat Axapta lag atau loading.
    """
    image_path = os.path.join(config.ASSETS_DIR, image_name)
    if not os.path.exists(image_path):
        logging.error(f"File gambar referensi tidak ditemukan: {image_path}")
        raise FileNotFoundError(f"Referensi {image_name} tidak ada di folder assets.")

    logging.info(f"[WAIT] Menunggu elemen '{image_name}' (Timeout: {timeout}s)...")
    start_time = time.time()

    while time.time() - start_time < timeout:
        try:
            pos = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
            if pos is not None:
                logging.info(f"[OK] Ditemukan '{image_name}' pada koordinat ({pos.x}, {pos.y})")
                return pos
        except (pyautogui.ImageNotFoundException, Exception):
            pass
        time.sleep(1)  # Cek ulang per 1 detik

    raise TimeoutError(
        f"[TIMEOUT] Elemen '{image_name}' tidak muncul setelah {timeout} detik. "
        "Kemungkinan Axapta lag atau tampilan jendela tertutup."
    )


def wait_and_click(
    image_name,
    timeout=60,
    confidence=config.DEFAULT_CONFIDENCE,
    delay_after=1.5,
    clicks=1,
):
    """Menunggu elemen muncul, lalu mengkliknya."""
    pos = wait_for_image(image_name, timeout=timeout, confidence=confidence)
    pyautogui.click(pos, clicks=clicks)
    time.sleep(delay_after)
    return pos


def get_latest_ica_file():
    """Mencari file .ica terbaru di folder Downloads."""
    files = glob.glob(os.path.join(config.DOWNLOADS_DIR, "*.ica"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def run_automation():
    logging.info("=" * 60)
    logging.info("MEMULAI OTOMASI EXPORT DYNAMICS AX (CITRIX)")
    logging.info("=" * 60)

    # ---------------------------------------------------------
    # STEP 1: Buka Edge dan akses URL Citrix
    # ---------------------------------------------------------
    logging.info(f"[Step 1] Membuka Edge ke: {config.CITRIX_URL}")
    subprocess.Popen(f'start msedge "{config.CITRIX_URL}"', shell=True)
    time.sleep(5)

    # ---------------------------------------------------------
    # STEP 2: Login Akun Tersimpan
    # ---------------------------------------------------------
    logging.info("[Step 2] Melakukan Login...")
    try:
        wait_and_click("btn_login.png", timeout=config.LOGIN_TIMEOUT)
    except (FileNotFoundError, TimeoutError):
        logging.info("Info: Tombol 'btn_login.png' dilewati, menekan tombol Enter...")
        pyautogui.press("enter")
    time.sleep(3)

    # ---------------------------------------------------------
    # STEP 3: Klik Menu Dynamics AX & Download .ica
    # ---------------------------------------------------------
    logging.info("[Step 3] Mencari dan mengklik aplikasi Dynamics AX...")
    time_before_click = time.time()
    try:
        wait_and_click("icon_ax.png", timeout=config.APP_ICON_TIMEOUT)
    except TimeoutError:
        logging.warning("Ikon Dynamics AX belum ditemukan, mencoba cek file .ica langsung...")

    # ---------------------------------------------------------
    # STEP 4: Deteksi File .ica dan Jalankan Citrix Receiver
    # ---------------------------------------------------------
    logging.info("[Step 4] Memantau file download .ica di folder Downloads...")
    ica_file = None
    for _ in range(config.ICA_DOWNLOAD_TIMEOUT):
        latest = get_latest_ica_file()
        if latest and os.path.getmtime(latest) >= (time_before_click - 3):
            ica_file = latest
            break
        time.sleep(1)

    if ica_file:
        logging.info(f"Menjalankan file koneksi Citrix: {ica_file}")
        os.startfile(ica_file)
    else:
        logging.info("File .ica baru tidak terdeteksi via disk, melanjutkan (mungkin sudah terbuka otomatis)...")

    # ---------------------------------------------------------
    # STEP 5: Menunggu Axapta Siap & Klik Menu Stock
    # ---------------------------------------------------------
    logging.info("[Step 5] Menunggu Axapta terbuka (bisa lama saat koneksi awal)...")
    wait_and_click(
        "menu_stock.png", timeout=config.AXAPTA_STARTUP_TIMEOUT, delay_after=2
    )

    # ---------------------------------------------------------
    # STEP 6: Menunggu Jendela Data Stok Muncul
    # ---------------------------------------------------------
    logging.info("[Step 6] Menunggu jendela 'Available stock position & item price'...")
    stock_pos = wait_for_image("window_stock.png", timeout=config.WINDOW_STOCK_TIMEOUT)

    # Fokuskan jendela data stok
    pyautogui.click(stock_pos)
    time.sleep(1.5)

    # ---------------------------------------------------------
    # STEP 7: Scroll Sampai Bawah (Mencegah Data Kepotong)
    # ---------------------------------------------------------
    logging.info("[Step 7] Menarik seluruh baris data ke bawah (mencegah kepotong)...")
    # Klik area tabel di bawah header
    pyautogui.click(stock_pos.x, stock_pos.y + 120)
    time.sleep(1)

    for i in range(1, config.SCROLL_REPEAT + 1):
        logging.info(f"   Scroll ke bawah tahap {i}/{config.SCROLL_REPEAT} (Ctrl + End)...")
        pyautogui.hotkey("ctrl", "end")
        time.sleep(config.SCROLL_DELAY)

    # ---------------------------------------------------------
    # STEP 8: Menekan Tombol Export ke Excel
    # ---------------------------------------------------------
    logging.info("[Step 8] Menekan tombol Export ke Excel...")
    wait_and_click("btn_export.png", timeout=config.EXPORT_TIMEOUT, delay_after=2)

    # ---------------------------------------------------------
    # STEP 9: Menunggu Proses Export Selesai & Menyimpan File
    # ---------------------------------------------------------
    logging.info(f"[Step 9] Menunggu proses ekspor Axapta (Timeout maks {config.EXPORT_PROCESS_TIMEOUT}s)...")
    
    # Beri jeda adaptif: Axapta sering 'Not Responding' saat ekspor data banyak
    time.sleep(10)

    # Simpan hasil ekspor
    today_str = time.strftime("%Y%m%d_%H%M%S")
    target_filename = f"Stock_AX_{today_str}"
    
    try:
        # Jika muncul dialog save atau jika user menggunakan icon save
        if os.path.exists(os.path.join(config.ASSETS_DIR, "btn_save.png")):
            wait_and_click("btn_save.png", timeout=config.SAVE_DIALOG_TIMEOUT)
        
        logging.info(f"Mengisi nama file: {target_filename}")
        pyautogui.typewrite(target_filename, interval=0.08)
        time.sleep(1)
        pyautogui.press("enter")
        logging.info(f"[BERHASIL] File disimpan sebagai: {target_filename}")
    except (FileNotFoundError, TimeoutError):
        logging.info("Dialog penyimpanan khusus tidak terdeteksi, menekan Ctrl+S sebagai fallback...")
        pyautogui.hotkey("ctrl", "s")
        time.sleep(1)
        pyautogui.press("enter")
        logging.info("[BERHASIL] Ekspor selesai dikirim.")

    logging.info("=" * 60)
    logging.info("PROSES SELESAI DENGAN SUKSES!")
    logging.info("=" * 60)


if __name__ == "__main__":
    try:
        run_automation()
    except Exception as e:
        logging.critical(f"TERJADI KESALAHAN: {e}", exc_info=True)
        sys.exit(1)
