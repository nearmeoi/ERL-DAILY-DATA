# 📊 ERL DAILY DATA — Sistem Otomasi Pipeline Cabang Makassar

Aplikasi desktop otomasi pengolahan posisi stok, alokasi FIFO buku masuk, dan rekapitulasi DO Summary harian untuk Cabang Makassar.

---

## 🚀 Fitur Utama

1. **Agregasi Multi-Gudang (STOK)**:
   - Pemilahan otomatis stok Gudang 50 (Pusat), Gudang 56 (Makassar), dan Nasional (Exclude 50, 56, 723).
2. **Karantina Murni 2026 (TRN)**:
   - Filter ketat data karantina tahun berjalan 2026 dan pembuatan sheet Pivot secara otomatis.
3. **Agregasi Deliver Reminder & Remainder (SOOUT & POOUT)**:
   - Pembuatan Sheet2 Pivot SOOUT (Qty > 0) dan POOUT.
4. **Pengisian Cepat Format Cabang**:
   - Injeksi nilai langsung ke 22.400+ baris Format Cabang tanpa formula eksternal yang rentan rusak.
5. **Alokasi Buku Masuk & Ekstraksi FIFO**:
   - Deteksi otomatis penerimaan buku masuk.
   - VLOOKUP hierarki `GL`, `ASM`, `MGR` dari master referensi `SALESUNIT2026.xlsx`.
   - Penyusunan antrean FIFO ke `Sheet3` SOOUT.
   - Pemisahan lembar cetak `MGR1`, `MGR2`, dan pembagian batch 14 SO pada sheet `DELIMIT` untuk Dynamics AX.
   - *Dilewati secara otomatis (adaptif) jika tidak ada buku masuk hari tersebut.*
6. **Pivot Tabular DO Summary External**:
   - Penguraian CSV Text to Columns, Pivot tabular native Excel dengan filter status `1=Open` dan `3=Close`.
7. **Arsitektur Pipeline Modular (1–9 Stages)**:
   - Pelacakan durasi milidetik per tahap layaknya video editing timeline.
   - Error handling diagnostik berstandar enterprise dengan tips solusi praktis.

---

## 💻 Cara Menjalankan

### 1. Langsung Menggunakan File Executable (Tanpa Install Python)
Klik ganda file:
```
Daily Data.exe
```

### 2. Menggunakan Script Launcher
* **GUI Desktop**: Klik ganda `JALANKAN_OTOMASI.bat`
* **Mode Console**: Klik ganda `JALANKAN_CLI.bat`

---

## 📁 Struktur Direktori

```text
├── Daily Data.exe            # Executable mandiri (Standalone)
├── JALANKAN_OTOMASI.bat      # Launcher GUI Desktop
├── JALANKAN_CLI.bat          # Launcher Mode Terminal
├── App/
│   ├── gui.py                # Antarmuka GUI Tkinter dengan timeline tracker
│   ├── main.py               # Engine pipeline 9 tahapan modular
│   └── process_do_summary.py # Engine pengolahan DO Summary
├── REFERENSI/
│   └── SALESUNIT2026.xlsx    # Master hierarki sales responsible cabang
└── RULE/
    ├── DO_AND_DONT.md        # Aturan baku & pedoman operasional harian
    └── README.md             # Penjelasan dokumentasi aturan
```

---

## 📜 Lisensi & Hak Cipta
Hak cipta © 2026 Cabang Makassar. Dikembangkan untuk efisiensi pengolahan data harian.
