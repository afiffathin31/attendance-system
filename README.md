# 🎯 Sistem Absensi Hybrid (Ekstensi In-Meeting & Offline QR Check-in)

Sistem absensi komprehensif berbasis Web yang mengintegrasikan pencatatan kehadiran otomatis untuk **Online Meeting (Google Meet & Zoom)** menggunakan **Ekstensi Browser 1-Klik**, serta **Form Presensi Offline** berbasis **QR Code**, tersinkronisasi langsung ke **Excel (.xlsx)** dan dilengkapi **Dashboard Analitik Grafik Performa**.

---

## ✨ Fitur Utama

1. **🧩 Ekstensi In-Meeting 1-Klik (Google Meet & Zoom)**
   - Widget mengambang elegan muncul otomatis di pojok kanan atas layar rapat.
   - Tombol **"⚡ Rekap Kehadiran Sekarang"**: 1 kali klik langsung menyerap seluruh peserta yang ada di layar rapat ke database dan file Excel.
   - Fitur **"Auto-Sync tiap 5 Menit"**: otomatis memindai dan mencatat peserta yang baru bergabung tanpa perlu klik ulang.
   - 100% Anti-blokir karena berjalan langsung di browser rapat Anda (tidak memerlukan bot terpisah).
   - Filter cerdas otomatis menyaring nama peserta asli dan mengabaikan ikon/menu Google.

2. **📱 Form Presensi Offline & Kiosk QR Code**
   - Form web responsif ramah smartphone (`/index.html`).
   - Karyawan cukup memasukkan **Nama Lengkap** dan memilih **Departemen**.
   - Halaman Display QR Code Kiosk (`/qr.html`) untuk diletakkan di resepsionis / lobi kantor.
   - Waktu kedatangan dicatat dari jam server secara presisi.

3. **📊 Dashboard Analitik & Visualisasi Grafik (Chart.js)**
   - **Rasio Ketepatan Waktu Keseluruhan**: Donut Chart perbandingan On-Time vs Terlambat.
   - **Performa Departemen**: Bar Chart perbandingan kehadiran tepat waktu vs terlambat antar divisi.
   - **Leaderboard Performa Individu**: Horizontal Bar Chart tingkat ketepatan waktu per karyawan.
   - **Tren Kedatangan Harian**: Line Chart fluktuasi kedatangan dari hari ke hari.
   - **Tabel Rekap Real-time**: Filter berdasarkan tanggal, departemen, metode, dan status dengan pencarian instan.

4. **📑 Sinkronisasi File Master Excel (.xlsx) Real-Time**
   - Setiap ada kehadiran baru (baik online maupun offline), file master Excel `data/master_absensi.xlsx` langsung diperbarui secara otomatis menggunakan **ExcelJS**.
   - Terdiri dari 3 Sheet:
     1. **Log Absensi**: Rekap rinci per baris dengan pewarnaan status (Hijau = Tepat Waktu, Merah = Terlambat).
     2. **Performa Individu**: Rekap kehadiran, persentase tepat waktu, dan rata-rata menit keterlambatan per orang.
     3. **Rekap Departemen**: Performa tingkat kehadiran per divisi.
   - Tombol satu-klik **"Unduh Master Excel"** di dashboard admin.

---

## 📁 Struktur Direktori

```text
attendance-system/
├── package.json              # Dependensi ringan (tanpa browser headless)
├── config.json               # Konfigurasi jam kantor, toleransi, password
├── server.js                 # Server Express & REST API
├── services/
│   ├── syncService.js        # Parser nama & sinkronisasi kehadiran rapat
│   ├── excelService.js       # Generator & Synchronizer Excel (.xlsx)
│   └── analyticsService.js   # Mesin hitung analitik & grafik
├── extension/                # Ekstensi Chrome/Edge resmi (Load Unpacked)
│   ├── manifest.json
│   ├── content.js
│   ├── styles.css
│   ├── popup.html
│   └── README.md
├── public/
│   ├── index.html            # Form Absensi Offline (Mobile-friendly)
│   ├── qr.html               # Halaman QR Code Display (Kiosk Mode)
│   ├── admin.html            # Dashboard Admin & Analitik Grafik
│   └── login.html            # Halaman Login Admin
├── pasang-ekstensi.bat       # Script pembantu pasang ekstensi
└── data/
    ├── attendance.json       # Database arsip presensi
    └── master_absensi.xlsx   # File master Excel yang ter-update otomatis
```

---

## 🚀 Panduan Menjalankan Sistem

### 1. Menjalankan Server
```bash
npm start
```

### 2. Memasang Ekstensi Browser (10 Detik)
1. Buka `chrome://extensions` (atau `edge://extensions`).
2. Aktifkan **Developer mode** di pojok kanan atas.
3. Klik **Load unpacked** dan pilih folder `extension/`.
4. Buka Google Meet / Zoom -> Widget absensi langsung muncul!
