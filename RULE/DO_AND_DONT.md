# ATURAN OPERASIONAL & PEDOMAN DO & DON'T
## Sistem Otomasi Posisi Stok, FIFO, & DO Summary Cabang Makassar

Dokumen ini berisi kumpulan aturan baku, logika operasional, serta hal-hal yang **WAJIB DILAKUKAN (DO)** dan **DILARANG DILAKUKAN (DON'T)** dalam pengolahan data harian.

---

## 1. PENYIAPAN FILE & FOLDER HARIAN

### **DO (Harus Dilakukan):**
* Buat satu folder khusus per tanggal, contoh: `07 OKT 2026`, `08 OKT 2026`, dst.
* Letakkan seluruh file harian di dalam folder tanggal tersebut:
  1. `FORMAT POSISISTOK-CABMKS <TANGGAL>.xlsx`
  2. `STOK <TANGGAL>.xlsx`
  3. `TRN <TANGGAL>.xlsx`
  4. `SOOUT <TANGGAL>.xlsx`
  5. `POOUT <TANGGAL>.xlsx`
  6. `buku masuk <TANGGAL>.xlsx` *(jika ada penerimaan buku hari itu)*
  7. `Do_Summary_External_<TANGGAL>.csv` *(di sore hari saat file DO diunduh)*
* Pastikan file master referensi `SALESUNIT2026.xlsx` tetap tersedia di drive `D:\` atau di folder induk proyek.

### **DON'T (Jangan Dilakukan):**
* **JANGAN** membuka file-file Excel yang akan diproses di Microsoft Excel saat aplikasi sedang berjalan. Excel mengunci file (*file lock*) sehingga aplikasi akan mengalami error `PermissionError: Permission denied`.
* **JANGAN** mengubah pola nama file utama (misalnya mengubah format `FORMAT POSISISTOK` menjadi nama lain yang tidak diawali kata tersebut).
* **JANGAN** mencampur file hari kemarin dengan file hari ini di dalam folder yang sama.

---

## 2. PENGOLAHAN FILE STOK (TUGAS 1)

### **DO (Harus Dilakukan):**
* Buat Pivot Table dengan baris: `Item number` dan nilai: `Sum of Available Physical`.
* Filter kolom `Warehouse` sesuai ketentuan:
  * **Pusat**: Gudang **`50`**
  * **Makassar**: Gudang **`56`**
  * **Nasional**: **SEMUA gudang KECUALI `50`, `56`, dan `723`** *(Gunakan filter EXCLUDE)*.
* Gunakan konversi angka yang aman agar data desimal dari AX tidak salah terkonversi.

### **DON'T (Jangan Dilakukan):**
* **JANGAN PERNAH** mencentang *hanya* gudang `50`, `56`, dan `723` untuk stok Nasional. Ini adalah kesalahan fatal yang sering terjadi di Excel saat filter dropdown terbalik antara mencentang (*include*) dan mengecualikan (*exclude*). Akibatnya, stok cabang-cabang lain di seluruh Indonesia justru hilang.
* **JANGAN** memasukkan gudang transit atau gudang `723` ke dalam stok Nasional.

---

## 3. PENGOLAHAN FILE TRN (TUGAS 2)

### **DO (Harus Dilakukan):**
* Filter kolom `Quarantine date` **HANYA tahun 2026**.
* Hapus (*delete*) seluruh baris data dari tahun selain 2026 (2018, 2019, 2020, 2021, 2024, 2025).
* Buat Pivot Table dengan baris: `Item number` dan nilai: `Sum of Quantity`.
* Jika melakukan secara manual di Excel: **WAJIB me-refresh Pivot Table** (`Right Click -> Refresh`) setelah menghapus baris di Sheet1 agar angka tahun lama tidak terbawa.
* Letakkan sheet pivot (`Sheet2`) di posisi tab pertama (paling kiri) dengan header baris 3 (`Row Labels` & `Sum of Quantity`) dan baris `Grand Total` di akhir.

### **DON'T (Jangan Dilakukan):**
* **JANGAN** membiarkan data karantina tahun lampau ikut terhitung ke dalam Format Cabang.
* **JANGAN** lupa me-refresh pivot table jika membuat secara manual di Excel, karena Excel tidak memperbarui pivot secara otomatis saat data sumber dihapus.

---

## 4. PENGOLAHAN FILE SOOUT & POOUT (TUGAS 3 & 4)

### **DO (Harus Dilakukan):**
* **SOOUT**: Buat Pivot Table mengambil `Item number` dan `Sum of Deliver Reminder`.
* Catat bahwa nilai `Deliver Reminder` di raw file AX tersimpan sebagai angka negatif. Saat dipetakan ke file Format Cabang, masukkan nilai **positif (absolut)**.
* **POOUT**: Buat Pivot Table mengambil `Item number` dan `Sum of Deliver remainder`.
* Buang item kosong `(blank)` dari daftar pivot dan sertakan baris `Grand Total` di akhir tabel.

### **DON'T (Jangan Dilakukan):**
* **JANGAN** memasukkan nilai sisa SOOUT bertanda negatif ke file Format Cabang (kolom SOOUT di format cabang harus positif).

---

## 5. PENGISIAN FILE FORMAT CABANG (TUGAS 5)

### **DO (Harus Dilakukan):**
* Isi hasil agregasi pivot ke kolom-kolom yang sesuai pada Format Cabang:
  * Kolom 33 (`STOK_PST`) : Diisi dari Pivot STOK Pusat (WH 50)
  * Kolom 34 (`STOK_NAS`) : Diisi dari Pivot STOK Nasional (Semua WH kecuali 50, 56, 723)
  * Kolom 35 (`STOK_MKS`) : Diisi dari Pivot STOK Makassar (WH 56)
  * Kolom 36 (`TRANS`)    : Diisi dari Pivot TRN tahun 2026
  * Kolom 37 (`SOOUT`)    : Diisi dari Pivot SOOUT (nilai positif)
  * Kolom 38 (`PO OUT`)   : Diisi dari Pivot POOUT
* Jaga keutuhan formula bawaan Format Cabang (kolom 39 s/d 45: `STK+TRN+PO`, `BUFFER`, `KET STOK`, `FAST-SLOW`, dll.).
* Pertahankan relasi link eksternal (`externalLink1.xml.rels`) agar saat file dibuka di Excel tidak muncul pop-up error *Repaired Records*.

### **DON'T (Jangan Dilakukan):**
* **JANGAN** menimpa formula bawaan kolom 39 s/d 45 dengan nilai mati (hardcode).
* **JANGAN** melakukan iterasi mundur dari 1.000.000 baris karena openpyxl bisa lambat atau freeze jika `max_row` Excel bengkak.

---

## 6. KASUS ADA BUKU MASUK & FIFO (TUGAS 6, 7, 8, 9)

### **DO (Harus Dilakukan):**
* **Hanya dijalankan jika ada file Buku Masuk hari ini**. Jika tidak ada, proses ini dilewati dengan bersih (*skip*).
* Di file `SOOUT`:
  1. Hapus transaksi `56-SOI` (hanya pertahankan transaksi penjualan murni `56-SOE`).
  2. Urutkan secara **FIFO**: Urutkan berdasarkan `Item number` (A-Z), lalu `Created date and time` dari tanggal & jam **terlama ke terbaru**.
  3. Sisipkan 3 kolom di kolom ke-6 (Kolom F): `GL`, `ASM`, `MGR`. Isi nilai berdasarkan kode sales responsible yang dicocokkan ke database `SALESUNIT2026`.
  4. Simpan hasil SOE bersih berurutan FIFO ini ke sheet baru `Sheet3` di file SOOUT.
* Buat file print khusus: `BUKMAS <TANGGAL> PRINT.xlsx` berisi 3 sheet:
  * **Sheet `MGR1`**: Memuat data SOE untuk Manager 1 dengan susunan **28 kolom standar AX** (Sales order s/d Customer account).
  * **Sheet `MGR2`**: Memuat data SOE untuk Manager 2 dengan susunan **28 kolom standar AX**.
  * **Sheet `DELIMIT`**: Mengelompokkan nomor SO per manager yang dipisahkan koma (`,`) dengan aturan **TEPAT 14 ITEM PER BATCH** (jangan lebih dan jangan kurang, kecuali batch sisa terakhir).

### **DON'T (Jangan Dilakukan):**
* **JANGAN** memproses file Buku Masuk usang dari hari kemarin jika hari ini tidak ada buku masuk yang diterima.
* **JANGAN** menambahkan kolom kustom di luar 28 kolom standar AX pada sheet `MGR1` dan `MGR2` karena akan menyulitkan proses cetak dan copy-paste AX.
* **JANGAN** membagi batch SO di sheet DELIMIT selain 14 item (harus pas 14 item agar muat dalam batas query filter Dynamics AX).

---

## 7. DO SUMMARY EXTERNAL (PIVOT TABULAR SORE HARI)

### **DO (Harus Dilakukan):**
* Buka file `Do_Summary_External_<TANGGAL>.csv`.
* Lakukan **Text to Columns** pada Kolom A menggunakan delimiter koma (`,`) agar 21 kolom data mentah terurai rapi.
* Buat Pivot Table native Excel pada sheet baru `Sheet1`:
  * **Rows (Hierarki)**: `WILMGR` $\rightarrow$ `WILASM` $\rightarrow$ `WILGL`.
  * **Values**: `Sum of qty` dan `Sum of rp` (format angka ribuan `#,##0`).
  * **Filter (Page Field)**: `mstatus` hanya memilih **`1=Open`** dan **`3=Close`** (sembunyikan `2=Cancel` dan blank).
  * **Report Layout**: Atur ke **Show in Tabular Form** dengan subtotal di bawah setiap grup.
* **Kunci Urutan Daerah Operasional Cabang Makassar**:
  * **MANAGER 1 MAKASSAR**:
    1. ASM 1 MAKASSAR: `GL SD 1` $\rightarrow$ `GL SD 2` $\rightarrow$ `REP PANGKEP`
    2. ASM PARE-PARE: `GL PARE-PARE` $\rightarrow$ `REP SIDRAP` $\rightarrow$ `REP MAMUJU`
    3. ASM KENDARI: `GL SMP/A KENDARI` $\rightarrow$ `REP KOLAKA` $\rightarrow$ `REP BUTON`
  * **MANAGER 2 MAKASSAR**:
    1. ASM 2 MAKASSAR: `GL SMP MAKASSAR` $\rightarrow$ `GL SMA/K MAKASSAR`
    2. ASM BONE: `REP BULUKUMBA` $\rightarrow$ `GL BONE` $\rightarrow$ `REP BANTAENG`
    3. ASM PALOPO: `GL PALOPO` $\rightarrow$ `REP MASAMBA` $\rightarrow$ `REP TORAJA`
    4. GL NB MAKASSAR: `GL NB MAKASSAR`
* Simpan workbook sebagai format `.xlsx`.

### **DON'T (Jangan Dilakukan):**
* **JANGAN** membiarkan urutan daerah disortir berdasarkan abjad bawaan Excel (A-Z), karena urutan regional akan menjadi tidak sesuai struktur organisasi cabang.
* **JANGAN** mengikutkan status `2=Cancel` ke dalam perhitungan Qty dan Rp.

---

## 8. CARA MENJALANKAN APLIKASI (1-KLIK)

### Pilihan 1: Mode GUI Desktop
1. Klik ganda **`JALANKAN_OTOMASI.bat`**.
2. Pada dropdown **Folder Harian**, pilih atau ketik nama foldernya (misal `07 OKT 2026`).
3. Klik tombol hijau **▶ RUN**.
4. Tunggu hingga log menampilkan status sukses dan timeline ringkasan selesai.

### Pilihan 2: Mode Konsol CLI Cepat
1. Klik ganda **`JALANKAN_CLI.bat`**.
2. Ketik nama foldernya (misal `07 OKT 2026`) atau cukup tekan **Enter** untuk langsung memilih folder tanggal terbaru.
3. Seluruh proses akan berjalan otomatis dalam beberapa detik.

---

## 9. ATURAN PENGEMBANGAN SISTEM & GIT VERSION CONTROL

### **DO (Harus Dilakukan):**
* Selalu pastikan `.gitignore` mengabaikan file biner besar (misal `.xlsx` ukuran puluhan MB, folder virtual environment `.venv`, cache build PyInstaller).
* Pastikan source code modular, terdokumentasi, dan lolos uji coba sebelum diajukan ke git.

### **DON'T (DILARANG KERAS):**
* **JANGAN COMMIT TANPA SAYA SURUH (STRICT RULE FOR AI ASSISTANT)**:
  * Asisten AI **DILARANG KERAS** melakukan `git commit` maupun `git push` secara inisiatif sendiri tanpa adanya perintah eksplisit dari pengguna.
  * Setiap perubahan kode lokal harus dikonfirmasi terlebih dahulu. Commit ke repository Git hanya boleh dieksekusi jika pengguna secara tegas memerintahkan: *"commit ke git"*, *"push ke repo"*, atau instruksi serupa.

