# DOKUMENTASI ATURAN & STANDAR OPERASIONAL (RULE)

Folder ini berisi pedoman resmi, aturan operasional, serta hal-hal yang wajib dilakukan (*Do*) dan dilarang dilakukan (*Don't*) dalam sistem otomasi harian Cabang Makassar.

---

## Daftar Berkas:

1. **[`DO_AND_DONT.md`](file:///C:/Users/Akmal/Akmal/RULE/DO_AND_DONT.md)**:
   * Pedoman lengkap Do and Don't untuk seluruh alur:
     - Penyiapan File & Folder Harian
     - Pengolahan STOK (Pusat, Makassar, Nasional Exclude)
     - Pengolahan TRN (Filter 2026 murni & Refresh Pivot)
     - Pengolahan SOOUT & POOUT
     - Pengisian Format Cabang & Proteksi Formula/Link Eksternal
     - Alokasi Buku Masuk, FIFO, & Format Print Delimit 14 Item
     - Pengolahan DO Summary External Tabular Pivot
     - Cara Eksekusi 1-Klik (GUI & CLI)

---

## Ringkasan Cepat Hal Krusial:

| Bagian | Hal yang Sering Salah (DON'T) | Hal yang Benar (DO) |
| :--- | :--- | :--- |
| **STOK Nasional** | Mencentang *hanya* gudang 50, 56, 723 di filter Excel (cabang lain hilang). | Gunakan filter **EXCLUDE**: Semua gudang **KECUALI** 50, 56, 723. |
| **TRN 2026** | Lupa refresh pivot table setelah delete baris non-2026 di Sheet1 (angka tahun 2018-2025 ikut terhitung). | Pastikan Pivot Table hanya memuat data tahun **2026 murni** (392 item unik). |
| **Buku Masuk** | Memproses file buku masuk kemarin saat hari ini tidak ada buku masuk. | Jika hari ini tidak ada buku masuk, **lewati sepenuhnya** langkah alokasi buku masuk. |
| **Delimit SO** | Membuat batch delimit lebih atau kurang dari 14 item. | Bagi tepat **14 item per baris** agar muat dalam query Dynamics AX. |
| **File Lock** | Membuka file Excel di aplikasi Excel saat otomasi dijalankan. | Tutup seluruh file Excel target sebelum menekan tombol jalankan otomasi. |
