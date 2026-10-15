import os

# Direktori Dasar
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
LOGS_DIR = os.path.join(BASE_DIR, "logs")
EXPORTS_DIR = os.path.join(BASE_DIR, "exports")
DOWNLOADS_DIR = os.path.join(os.path.expanduser("~"), "Downloads")

# Buat folder jika belum ada
os.makedirs(LOGS_DIR, exist_ok=True)
os.makedirs(EXPORTS_DIR, exist_ok=True)
os.makedirs(ASSETS_DIR, exist_ok=True)

# URL Citrix StoreWeb
CITRIX_URL = "https://citrix.erlangga.co.id/Citrix/StoreWebWeb/"

# Pengaturan Deteksi Gambar
# Default 0.72 sangat ideal untuk remote desktop / Citrix video stream
DEFAULT_CONFIDENCE = 0.72

# Waktu Tunggu / Timeout Anti-Lag (dalam detik)
# Dibuat longgar karena koneksi Citrix & Axapta sering 'Not Responding' saat query besar
LOGIN_TIMEOUT = 25
APP_ICON_TIMEOUT = 35
ICA_DOWNLOAD_TIMEOUT = 30
AXAPTA_STARTUP_TIMEOUT = 120   # Startup awal Axapta dari .ica
WINDOW_STOCK_TIMEOUT = 60      # Menunggu jendela stok muncul
EXPORT_PROCESS_TIMEOUT = 180   # Menunggu Axapta selesai proses 'Exporting data to Excel'
SAVE_DIALOG_TIMEOUT = 45

# Scroll Data Table
SCROLL_REPEAT = 4              # Berapa kali menekan Ctrl+End untuk memuat data sampai akhir
SCROLL_DELAY = 2.5             # Jeda per scroll (detik) agar Axapta selesai merender baris
