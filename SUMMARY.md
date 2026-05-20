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

## Rekomendasi Peningkatan (Improvements)

1.  **Optimasi Query Data (Kritis)**:
    Saat ini, `graphService.getListItems` mengambil *semua* data dari SharePoint lalu memfilternya di sisi klien. Seiring bertambahnya data, aplikasi akan menjadi lambat. Gunakan parameter OData `$filter` di URL API untuk mengambil hanya data yang diperlukan (misal: hanya data bulan ini atau hanya NIP tertentu).

2.  **Keamanan Konfigurasi**:
    Meskipun aplikasi berbasis klien, sebaiknya hindari membiarkan ID teknis terlalu terekspos secara mentah. Pertimbangkan proses *build* (seperti menggunakan Vite atau Webpack) untuk mengelola variabel lingkungan.

3.  **Implementasi Offline Absen**:
    Service worker sudah memiliki *placeholder* untuk sinkronisasi latar belakang. Implementasikan antrean (queue) di `localStorage` sehingga karyawan tetap bisa menekan tombol absen saat sinyal lemah, dan data dikirim otomatis saat koneksi kembali.

4.  **Verifikasi Lokasi (Geolokasi)**:
    Meskipun WFA, perusahaan mungkin butuh data lokasi saat absen. Manfaatkan *Geolocation API* browser untuk mencatat koordinat GPS saat karyawan menekan tombol absen.

5.  **Validasi Form yang Lebih Ketat**:
    Tingkatkan validasi pada pengajuan WFA dan input karyawan baru untuk mencegah data duplikat atau format yang salah sebelum dikirim ke server.

6.  **Modularitas Kode**:
    `js/app.js` sudah mulai membesar. Pertimbangkan untuk memecah logika per-tab ke dalam modul terpisah (misal: `calendar.js`, `attendance.js`, `request.js`) agar lebih mudah dipelihara.
