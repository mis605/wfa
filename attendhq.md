# AttendHQ — Rangkuman Perubahan & Pengembangan

**Sistem Presensi Head Office PT. GOS INDORAYA**
Berbasis Microsoft 365 · SharePoint Lists · Microsoft Graph API

---

## Latar Belakang

Proyek ini dimulai dari dua versi file yang ditemukan: `wfa-hp.zip` (versi lama, digunakan di HP) dan `wfa-git.zip` (versi GitHub, lebih baru). Perbedaan utama keduanya adalah mekanisme persetujuan: versi lama mengharuskan atasan klik Approve secara manual, sedangkan versi GitHub sudah menggunakan **auto-approve** — status langsung `Approved` saat user submit, atasan hanya perlu intervensi jika ingin menolak.

Seluruh pengembangan dilakukan di atas versi `wfa-git` sebagai base.

---

## Daftar Perubahan

### 1. Rebranding: Absen WFA CO → AttendHQ

Nama aplikasi diganti dari "Absen WFA CO" menjadi **AttendHQ** untuk mencerminkan scope yang spesifik (Head Office) dan terasa lebih modern sebagai nama produk internal.

Perubahan dilakukan di:
- `<title>` browser tab
- Apple PWA meta tag (`apple-mobile-web-app-title`)
- Logo di topbar (`Attend` + `HQ`)
- Judul halaman login
- Subtitle halaman login: "Sistem Presensi Head Office"
- `manifest.json`: `name`, `short_name`, `description`

---

### 2. Perubahan Flow Utama

#### Tab 1 — Kalender (Default setelah login)

Sebelumnya kalender hanya menampilkan jadwal WFA yang disetujui. Sekarang menampilkan **data kehadiran lengkap** dengan informasi tambahan:

- Karyawan yang sudah absen tampil dengan badge berwarna sesuai tipe: hijau (WFO), biru (WFA), oranye (Visit)
- Karyawan yang punya jadwal WFA/Visit approved tapi **belum absen** tampil dengan label kuning bertanda `*`
- Klik pada tanggal membuka modal detail: menampilkan jam masuk, jam keluar, durasi kerja, dan tipe kehadiran per orang
- Filter visibilitas: user biasa hanya melihat sesama peer (atasan sama), tidak termasuk atasannya sendiri
- Akun superuser (dikonfigurasi di `config.js`) dapat melihat semua karyawan

#### Tab 2 — Absen

Alur absen diubah total:

- Sebelumnya: langsung tampil tombol masuk/keluar dengan validasi jadwal WFA
- Sekarang: user pilih tipe kehadiran dulu via **radio button** (WFO / WFA / Visit)
  - **WFO**: langsung boleh absen tanpa syarat
  - **WFA / Visit**: sistem cek apakah ada request approved untuk hari ini. Jika belum ada, muncul popup info dengan tombol langsung ke halaman Request
- Konfirmasi absen menggunakan modal (bukan langsung submit), menampilkan tipe, jam, dan mode sebelum disimpan

#### Tab 3 — Request (sebelumnya hanya WFA)

- Ditambah pilihan tipe via **radio button**: WFA atau Visit (sebelumnya freetext alasan)
- Kolom alasan tetap freetext
- Setelah submit, user langsung diarahkan ke Tab Absen dengan tipe yang baru diajukan sudah terseleksi otomatis — user bisa langsung absen tanpa perlu navigasi manual
- Email notifikasi menyesuaikan tipe (WFA/Visit) di subject dan konten

#### Email Notifikasi

- Subject berubah sesuai tipe: `[WFA Notification]` atau `[Visit Notification]`
- Tombol di email atasan hanya **Batalkan/Tolak** (tidak ada Approve, karena sudah auto-approve)
- CC tetap ke HRD dan MIS seperti sebelumnya
- Saat ditolak, email balasan dikirim ke user dengan catatan atasan

---

### 3. Struktur SharePoint Lists

Perubahan kolom yang diperlukan:

**List Absensi** — tambah kolom baru:
- `Tipe` (Single line of text): diisi otomatis `WFO`, `WFA`, atau `Visit`

**List Permohonan WFA** — tambah kolom baru:
- `Tipe` (Single line of text): diisi `WFA` atau `Visit`

Kolom `Keterangan` di list absensi **tidak diubah** (dipertahankan untuk kompatibilitas data lama).

Kolom `Tanggal` di list absensi diaktifkan **Indexed** di SharePoint untuk mendukung server-side filtering.

---

### 4. Superuser Kalender

Akun tertentu dapat dikonfigurasi untuk melihat semua karyawan di Tab Kalender (tidak terbatas peer/bawahan). Dikonfigurasi di `config.js`:

```js
calendarSuperuserEmails: [
  "mis@gos.co.id",
]
```

Tambah email lain jika perlu (misal HRD, direktur).

---

### 5. Perbaikan Bug: Timezone WIB

Fungsi `getTodayString()` di `utils.js` sebelumnya menggunakan `toISOString()` yang mengembalikan waktu UTC. Akibatnya, user yang membuka app sebelum jam 07:00 WIB mendapat tanggal kemarin, sehingga jadwal WFA tidak terdeteksi dan tombol absen tidak muncul.

Diperbaiki dengan menggunakan `getFullYear()`, `getMonth()`, `getDate()` yang mengikuti waktu lokal device (WIB).

---

### 6. Optimasi Performa: Server-side Filter & Caching

**Sebelumnya:** semua item dari SharePoint list diambil seluruhnya ke browser, baru difilter di JavaScript.

**Sesudahnya:**

- Query absensi menggunakan `$filter` Graph API: `fields/Tanggal ge '...' and fields/Tanggal le '...'` — SharePoint yang memfilter di server
- Jika kolom belum diindex, fallback otomatis ke full fetch dengan warning di console
- **In-memory cache 5 menit** untuk semua query. Cache di-invalidate otomatis saat ada write (absen masuk/keluar, submit request)
- Query kalender menggunakan `Promise.all` untuk fetch absensi dan karyawan secara paralel
- `getPermohonanWfaById` menggunakan endpoint by ID langsung (tidak perlu load semua list)

---

### 7. Kolom Email Atasan di Absensi — Join dari List Karyawan

List absensi tidak memiliki kolom `Email_Atasan`. Daripada menambah kolom baru, data atasan di-**join** dari list karyawan menggunakan `buildNipAtasanMap()` yang membangun map `NIP → emailAtasan` sekali dan di-cache.

Keuntungan: single source of truth, data absensi lama tetap ter-cover, tidak perlu migrasi kolom.

---

### 8. UX Improvements

**Auto-detect status absen (No.6)**
Saat membuka Tab Absen, app langsung query absensi hari ini. Jika sudah absen, radio button otomatis tercentang sesuai tipe dan dashboard langsung tampil — tanpa perlu pilih manual. Guard tambahan: tidak bisa ganti tipe jika sudah absen.

**Notifikasi request ditolak in-app (No.7)**
Setiap login, app mengecek apakah ada request dengan status `Rejected` yang belum pernah dinotifikasi. Jika ada, muncul bottom sheet modal menampilkan tanggal dan catatan atasan. ID yang sudah dinotif disimpan di `localStorage` agar tidak muncul berulang.

**Tombol "Hari Ini" di kalender (No.8)**
Tombol `Hari Ini` muncul di navigasi bulan kalender, hanya tampil saat user sedang melihat bulan selain bulan sekarang. Satu klik langsung kembali ke bulan aktif.

**Progress bar saat submit absen (No.9)**
Di modal konfirmasi absen, saat tombol diklik muncul progress bar animasi (0% → 60% selama request, 100% saat selesai) sebelum modal menutup. Memberikan feedback visual yang jelas di mobile.

---

### 9. Hapus Debug & Error Boundary (No.11 & 12)

**Hapus debug (No.11)**
Semua pemanggilan `debugListColumns` dan `console.log` verbose di `loadUserSession` dihapus dari production code.

**Error boundary (No.12)**
Ditambahkan:
- Fungsi `showFatalError()`: menampilkan halaman error ramah dengan pesan dan tombol "Muat Ulang Aplikasi" saat terjadi crash yang tidak tertangani
- `window.addEventListener('unhandledrejection')` sebagai safety net global untuk promise rejection yang tidak di-catch

---

## File yang Diubah

| File | Keterangan |
|------|-----------|
| `index.html` | Struktur HTML, branding, elemen baru (progress bar, modal notif ditolak, tombol Hari Ini) |
| `js/app.js` | Seluruh logika UI dan flow baru |
| `js/graph.js` | Server-side filter, caching, join emailAtasan, hapus debug |
| `js/utils.js` | Fix timezone `getTodayString()` |
| `js/config.js` | Tambah `calendarSuperuserEmails` |
| `css/style.css` | Style komponen baru (radio tipe, badge, progress bar, error boundary, dll) |
| `manifest.json` | Branding PWA: name, short_name, description |

---

## Konfigurasi Penting di `config.js`

```js
const APP_CONFIG = {
  // Azure AD
  clientId: "...",
  tenantId: "...",

  // SharePoint
  sharepointSiteId: "...",
  listAbsensiId: "...",
  listKaryawanId: "...",
  listPermohonanWfaId: "...",

  // Jam absen
  jamMasukMulai: "07:00",
  jamMasukSelesai: "10:00",
  jamKeluarMulai: "16:00",
  jamKeluarSelesai: "20:00",
  toleransiTerlambat: 15,

  // Email notifikasi
  emailHrd: "hc.info@gos.co.id",
  emailMis: "mis@gos.co.id",

  // Akun yang dapat melihat semua karyawan di kalender
  calendarSuperuserEmails: [
    "mis@gos.co.id",
  ],
};
```

---

## Yang Belum Diimplementasikan (Backlog)

- **Reminder otomatis via Teams** (Cloudflare Worker + Cron + Incoming Webhook): kirim notifikasi ke Teams channel untuk karyawan yang belum absen pada jam tertentu
- **Validasi backend**: auto-approve saat ini dilakukan di frontend; idealnya menggunakan backend/Power Automate untuk keamanan lebih ketat (ditunda, tidak prioritas saat ini)

---

*Dokumen ini dibuat otomatis berdasarkan sesi pengembangan AttendHQ.*
*Last updated: Juni 2026*
