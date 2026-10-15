# Otomasi Export Dynamics AX (Citrix)

Sistem otomatisasi untuk membuka Microsoft Edge, login ke Citrix StoreWeb, mengunduh dan menjalankan file `.ica`, membuka Microsoft Dynamics AX (AXAPTA), menarik seluruh data stok tanpa terpotong, lalu mengekspornya secara otomatis setiap hari pukul 07:00 pagi.

---

## 📁 Struktur Folder

```text
AxaptaAutomation/
│
├── .venv/                   # Virtual Environment Python khusus (berisi pyautogui, opencv, dll)
├── assets/                  # Folder potongan screenshot referensi tombol/menu
│   └── README.md            # Daftar screenshot yang perlu diambil
├── exports/                 # Folder hasil penyimpanan file export (bisa disesuaikan di config.py)
├── logs/                    # Riwayat log jalannya otomatisasi setiap hari
│
├── config.py                # Konfigurasi URL, waktu tunggu (timeout anti-lag), dan parameter
├── auto_axapta.py           # Skrip utama otomatisasi
├── test_detection.py        # Skrip uji coba deteksi gambar di layar
├── JALANKAN_SEKARANG.bat     # Eksekusi otomatisasi dengan deteksi gambar (1x klik)
├── UJI_DETEKSI_GAMBAR.bat   # Jalankan tes deteksi gambar pada layar saat ini
│
├── REKAM_ALUR.bat           # PEREKAM ALUR: Rekam klik & ketikan Anda langsung di layar
├── PUTAR_REKAMAN.bat        # PUTAR REKAMAN: Jalankan ulang persis sesuai rekaman Anda
│
├── JADWALKAN_JAM_7_PAGI.bat # Daftarkan jadwal harian 07:00 ke Windows Task Scheduler
└── HAPUS_JADWAL.bat         # Hapus jadwal dari Task Scheduler jika ingin dinonaktifkan
```

---

## 🚀 Langkah Menjalankan Pertama Kali

### 1. Masukkan Screenshot Tombol ke Folder `assets/`
Buka `assets/README.md` untuk melihat daftar file gambar.
Gunakan shortcut **`Windows + Shift + S`** untuk mengambil potongan tombol:
* `btn_login.png` (Tombol Logon di web Citrix)
* `icon_ax.png` (Ikon Dynamics AX di web)
* `menu_stock.png` (Menu Available stock position di Axapta)
* `window_stock.png` (Judul jendela stok Axapta)
* `btn_export.png` (Tombol Export di Axapta)

### 2. Uji Deteksi Gambar
Buka aplikasi Citrix / Axapta Anda di layar, lalu klik dua kali:
👉 **`UJI_DETEKSI_GAMBAR.bat`**
Skrip akan memberitahu apakah tombol-tombol tersebut berhasil dikenali oleh komputer.

### 3. Uji Jalankan Penuh
Klik dua kali:
👉 **`JALANKAN_SEKARANG.bat`**
Perhatikan proses berjalan dari awal hingga akhir.

### 4. Aktifkan Jadwal Jam 07:00 Pagi
Klik kanan pada:
👉 **`JADWALKAN_JAM_7_PAGI.bat`** lalu pilih **"Run as Administrator"**.
Tugas otomatisasi sekarang sudah terdaftar di Windows Task Scheduler.

---

## 🛡️ Fitur Anti-Lag & Keamanan
1. **Dynamic Polling:** Tidak menggunakan jeda mati (*sleep* statis), melainkan terus memantau sampai tombol/jendela muncul (maksimal 120 detik untuk startup Axapta).
2. **Auto-Scroll `Ctrl + End`:** Menekan tombol scroll bawah berulang kali dengan jeda rendering agar database Axapta memuat semua baris tanpa ada yang kepotong.
3. **Emergency Stop (Fail-Safe):** Jika mouse bergerak aneh atau Anda ingin membatalkan seketika, cukup **dorong kursor mouse Anda ke pojok paling kiri-atas layar komputer**, maka skrip akan berhenti seketika.
