# Ringkasan Proyek: Absen WFA

Proyek ini adalah aplikasi web (PWA) untuk sistem absensi *Work From Anywhere* (WFA) yang terintegrasi dengan ekosistem Microsoft 365. Data disimpan dan dikelola menggunakan SharePoint Lists melalui Microsoft Graph API.

## Ringkasan Skrip

| File | Deskripsi Singkat |
| :--- | :--- |
| `sw.js` | **Service Worker**: Mengelola caching aset statis dan fungsionalitas PWA agar aplikasi dapat diakses lebih cepat dan mendukung fallback offline. |
| `js/config.js` | **Konfigurasi**: Berisi ID Azure AD, Site ID SharePoint, ID List, serta pengaturan jam operasional absen dan kebijakan perusahaan. |
| `js/auth.js` | **Autentikasi**: Menggunakan MSAL.js untuk menangani login/logout akun Microsoft, serta manajemen token akses dengan penyesuaian alur untuk perangkat mobile vs desktop. |
| `js/graph.js` | **Graph Service**: Jembatan komunikasi ke Microsoft Graph API. Menangani CRUD data karyawan, pencatatan absen, pengajuan WFA, serta pengiriman email otomatis via Outlook. |
| `js/app.js` | **Main Logic**: Otak utama aplikasi yang mengatur *state* global, navigasi antar tab (Kalender, Absen, Request), dan interaksi UI secara keseluruhan. |
| `js/utils.js` | **Utilities**: Berisi fungsi pembantu seperti format tanggal/jam, perhitungan durasi kerja, notifikasi toast, dan indikator loading. |
| `js/msal-browser.min.js` | **Library**: Pustaka resmi Microsoft untuk proses autentikasi di sisi klien. |

## Perubahan Terbaru (Recent Changes)

*   **Penyelesaian Stuck Verifikasi Login:** Memperbaiki transisi view pada `js/app.js` untuk memastikan loading spinner tertutup setelah login & verifikasi profil karyawan selesai.
*   **Sinkronisasi Skema Kolom (NIP/NRK):** Menyelaraskan field request dengan skema kolom SharePoint List yang menggunakan penamaan `NRK` untuk menghindari error Graph API request.
*   **Dinamisasi & Proteksi Email CC:** Memindahkan email tujuan CC (HRD & MIS) ke `js/config.js` dan menambahkan filtrasi otomatis di `js/graph.js` untuk mencegah duplikasi email CC dengan alamat tujuan (`To`) agar tidak ditolak oleh Graph API.
*   **Liquid & Responsive Layout:** Mengubah `css/style.css` menjadi fully-liquid (100% width) di mobile dan menerapkan sistem multi-kolom grid side-by-side pada layar tablet/desktop, serta memperbesar visualisasi sel kalender agar ramah perangkat desktop.
*   **Filtering Kalender Berdasarkan Atasan:** Menambahkan filter otomatis pada tab Kalender WFA. Karyawan hanya bisa melihat jadwal WFA miliknya sendiri, rekan kerja dengan atasan yang sama, dan bawahan langsungnya (jika pengguna adalah atasan).
*   **Popup Detail Kalender WFA:** Menambahkan fitur popup/bottom sheet interaktif saat mengklik tanggal di kalender untuk menampilkan daftar lengkap karyawan (Nama & NIP/NRK) yang WFA pada hari tersebut.
*   **Alur Persetujuan Implisit (Auto-Approve):** Mengubah status pengajuan WFA baru menjadi `Approved` secara instan agar karyawan bisa langsung melakukan absensi tanpa menunggu verifikasi manual.
*   **Email Notifikasi & Tautan Pembatalan Tunggal:** Menyesuaikan notifikasi email ke atasan dari tipe persetujuan menjadi pemberitahuan/notifikasi dengan tombol tunggal "Batalkan / Tolak Pengajuan" (Reject).
*   **Form Pembatalan Atasan dengan Catatan Opsional:** Memperbarui panel pembatalan/penolakan agar atasan dapat menolak permohonan WFA yang sudah disetujui, dengan kolom catatan alasan penolakan yang bersifat opsional (Pilihan B).
*   **Version Bump Cache PWA:** Menaikkan versi cache di `sw.js` ke `absen-wfa-v2.7` agar pembaruan layout dan script langsung terdeteksi oleh perangkat pengguna.

## Rekomendasi Peningkatan (Improvements)

1.  **Optimasi Query Data (Kritis)**:
    Saat ini, `graphService.getListItems` mengambil *semua* data dari SharePoint lalu memfilternya di sisi klien. Seiring bertambahnya data, aplikasi akan menjadi lambat. Gunakan parameter OData `$filter` di URL API untuk mengambil hanya data yang diperlukan (misal: hanya data bulan ini atau hanya NIP tertentu).

2.  **Keamanan Konfigurasi**:
    Meskipun aplikasi berbasis klien, sebaiknya hindari membiarkan ID teknis terlalu terekspos secara mentah. Pertimbangkan proses *build* (seperti menggunakan Vite atau Webpack) untuk mengelola variabel lingkungan.

3.  **Implementasi Offline Absen**:
    Service worker sudah memiliki *placeholder* untuk sinkronisasi latar belakang. Implementasikan antrean (queue) di `localStorage` sehingga karyawan tetap bisa menekan tombol absen saat sinyal lemah, dan data dikirim otomatis saat koneksi kembali.

4.  **Verifikasi Lokasi (Geolokasi)**:
    Meskipun WFA, perusahaan mungkin butuh data lokasi saat absen. Manfaatkan *Geolocation API* browser untuk mencatat koordinat GPS saat karyawan menekan tombol absen.

5.  **Aktivasi Fitur Admin (Rekap & Data Karyawan)**:
    Saat ini menu Rekap (`view-rekap`) dan Data Karyawan (`view-karyawan`) sudah memiliki markup HTML tetapi belum dihubungkan dengan logika JS di `js/app.js`. Fitur ini perlu diimplementasikan agar admin HRD/MIS dapat memantau absensi langsung di aplikasi.

6.  **Validasi Form yang Lebih Ketat**:
    Tingkatkan validasi pada pengajuan WFA dan input karyawan baru untuk mencegah data duplikat atau format yang salah sebelum dikirim ke server.

7.  **Modularitas Kode**:
    `js/app.js` sudah mulai membesar. Pertimbangkan untuk memecah logika per-tab ke dalam modul terpisah (misal: `calendar.js`, `attendance.js`, `request.js`) agar lebih mudah dipelihara.
