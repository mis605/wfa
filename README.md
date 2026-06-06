# AttendHQ — Sistem Presensi Head Office

Aplikasi presensi digital untuk karyawan **Head Office PT. GOS INDORAYA**, dibangun sebagai **Progressive Web App (PWA)** yang terintegrasi penuh dengan ekosistem **Microsoft 365**. Mendukung tiga tipe kehadiran: WFO, WFA, dan Visit.

---

## ✨ Fitur Utama

| Fitur | Deskripsi |
|-------|-----------|
| 🔐 **SSO Microsoft 365** | Login menggunakan akun email perusahaan via Azure AD / Entra ID |
| 📅 **Kalender Kehadiran** | Tampilan kalender real-time dengan badge tipe kehadiran (WFO/WFA/Visit) dan indikator belum absen |
| 🕒 **Absen Digital** | Absen masuk & keluar dengan pilihan tipe, konfirmasi modal, dan progress bar |
| ✍️ **Pengajuan WFA & Visit** | Pengajuan multi-tanggal langsung dari app; notifikasi email otomatis ke atasan; auto-approve |
| 🔔 **Notifikasi In-app** | Notifikasi otomatis jika pengajuan dibatalkan/ditolak atasan saat login berikutnya. Status tersimpan di SharePoint (`Notified_User`) — tidak muncul lagi di device manapun |
| 👥 **Manajemen Tim** | Atasan dapat menambah, mengedit, dan menghapus anggota tim langsung dari aplikasi. NIK otomatis terisi dari nilai tertinggi yang ada |
| 📲 **PWA** | Bisa diinstal di home screen HP (Android & iOS) seperti aplikasi native |
| ⚡ **Performa Optimal** | Server-side filter via Graph API `$filter` + in-memory cache 5 menit |

---

## 🏗️ Arsitektur

```
AttendHQ (Frontend SPA/PWA)
    │
    ├── Microsoft Authentication Library (MSAL.js)
    │       └── Azure AD / Entra ID
    │
    └── Microsoft Graph API
            ├── SharePoint List: Absensi
            ├── SharePoint List: Karyawan
            └── SharePoint List: Permohonan WFA
```

**Backend-less architecture** — seluruhnya mengandalkan ekosistem Microsoft 365 yang sudah dimiliki organisasi. Tidak membutuhkan server tambahan.

---

## 📂 Struktur File

```
attendhq/
├── index.html              # SPA utama — semua view ada di sini
├── manifest.json           # Konfigurasi PWA (nama, ikon, shortcuts)
├── sw.js                   # Service Worker untuk PWA
├── css/
│   └── style.css           # Styling utama (dark theme)
├── js/
│   ├── app.js              # Logika UI, flow tab, event binding
│   ├── graph.js            # Microsoft Graph API service (absensi, request, email)
│   ├── auth.js             # MSAL authentication handler
│   ├── config.js           # Konfigurasi (Client ID, List ID, jam kerja, superuser)
│   ├── utils.js            # Helper: format tanggal, toast, loading, getTodayString
│   └── msal-browser.min.js # MSAL.js library (lokal, bypass AdBlocker)
└── icons/
    ├── icon.svg
    ├── icon-192.png
    └── icon-512.png
```

---

## ⚙️ Setup & Deployment

### 1. Registrasi App di Azure Portal

1. Buka [Microsoft Entra ID](https://portal.azure.com/) → **App registrations** → **New registration**
2. Beri nama (misal: `AttendHQ`)
3. **Supported account types**: *Accounts in this organizational directory only*
4. **Redirect URI**: pilih tipe **Single-page application (SPA)**, masukkan URL hosting (contoh: `https://org.github.io/attendhq/`)
5. Klik **Register**, salin **Application (client) ID** dan **Directory (tenant) ID**

### 2. Konfigurasi API Permissions

Di halaman app yang baru dibuat → **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**, tambahkan:

- `User.Read`
- `Sites.ReadWrite.All`
- `Mail.Send`

Klik **Grant admin consent** agar user tidak perlu approve satu per satu.

### 3. Persiapan SharePoint Lists

Buat tiga List di SharePoint site organisasi dengan kolom berikut:

**List: Absensi**
| Kolom | Tipe | Keterangan |
|-------|------|-----------|
| Title | Single line | Diisi NIP/NRK (auto oleh sistem) |
| Nama | Single line | Nama karyawan |
| Tanggal | Single line | Format `YYYY-MM-DD` — **aktifkan Indexed** |
| Jam_Masuk | Single line | Format `HH:MM:SS` |
| Jam_Keluar | Single line | Format `HH:MM:SS` |
| Status | Single line | `Tepat Waktu` / `Terlambat Ringan` / `Terlambat` |
| Tipe | Single line | `WFO` / `WFA` / `Visit` |
| Keterangan | Multiple lines | Catatan tambahan (opsional) |

**List: Karyawan**
| Kolom | Tipe | Keterangan |
|-------|------|-----------|
| Title | Single line | NIP/NRK |
| Nama | Single line | Nama lengkap |
| Email | Single line | Email Microsoft 365 |
| Jabatan | Single line | |
| Departemen | Single line | |
| Status_Aktif | Single line | `Aktif` / `Tidak Aktif` |
| Email_Atasan | Single line | Email atasan langsung |

**List: Permohonan WFA**
| Kolom | Tipe | Keterangan |
|-------|------|-----------|
| Title | Single line | NIP/NRK (auto) |
| Nama | Single line | |
| Email_User | Single line | Email pemohon |
| Tanggal_WFA | Single line | Tanggal dipisah koma: `2026-06-01, 2026-06-02` |
| Tipe | Single line | `WFA` / `Visit` |
| Status | Single line | `Approved` / `Rejected` |
| Email_Atasan | Single line | Email atasan |
| Catatan_User | Multiple lines | Alasan pengajuan |
| Catatan_Atasan | Multiple lines | Catatan penolakan (opsional) |
| Notified_User | Single line | `true` jika user sudah melihat notifikasi penolakan — **wajib ada**, diisi sistem otomatis |

> **Tips performa**: Aktifkan **Indexed columns** pada kolom `Tanggal` di list Absensi.
> Caranya: SharePoint list → Settings → Indexed columns → Create a new index → pilih `Tanggal`.

### 4. Konfigurasi `js/config.js`

```javascript
const APP_CONFIG = {
  // Azure AD
  clientId: "MASUKKAN_CLIENT_ID",
  tenantId: "MASUKKAN_TENANT_ID",

  // Redirect URI (otomatis mengikuti URL hosting)
  redirectUri: window.location.origin + window.location.pathname,

  // SharePoint
  sharepointSiteId: "tenant.sharepoint.com,SITE_ID_1,SITE_ID_2",
  listAbsensiId:       "GUID_LIST_ABSENSI",
  listKaryawanId:      "GUID_LIST_KARYAWAN",
  listPermohonanWfaId: "GUID_LIST_PERMOHONAN",

  // Jam kerja (absen masuk FLEKSIBEL, keluar minimal 9 jam setelah masuk)
  durasiKerjaJam:   9,      // minimal jam kerja sebelum bisa absen keluar
  jamKeluarSelesai: "23:59",// batas akhir absen keluar (safety)
  jamAcuanTepat:    "08:00",// acuan jam "Tepat Waktu" untuk status absensi
  toleransiTerlambat: 15,   // menit toleransi setelah jamAcuanTepat

  // Identitas perusahaan
  namaPerusahaan: "PT. GOS INDORAYA",
  logoPerusahaan: "",

  // Email notifikasi (CC pada semua email keluar)
  emailHrd: "hc.info@gos.co.id",
  emailMis: "mis@gos.co.id",

  // Akun yang dapat melihat SEMUA karyawan di Tab Kalender
  calendarSuperuserEmails: [
    "mis@gos.co.id",
  ],
};
```

Cara mendapatkan `sharepointSiteId` dan `listId`:

```
# Site ID
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-path}

# List ID
GET https://graph.microsoft.com/v1.0/sites/{siteId}/lists
```

Gunakan [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) untuk query ini.

### 5. Menjalankan Lokal

```bash
# Python
python3 -m http.server 5500

# Node.js (npx)
npx serve .
```

Buka `http://localhost:5500`. Pastikan `http://localhost:5500` sudah ditambahkan ke **Redirect URI** di Azure Portal.

### 6. Deployment

Upload seluruh folder ke static hosting:

- **GitHub Pages** — push ke branch `gh-pages` atau folder `/docs`
- **Cloudflare Pages** — connect repo, build command kosong, output directory `/`
- **Vercel / Netlify** — drag & drop folder atau connect repo

Setelah deploy, tambahkan URL publik ke **Redirect URI** di Azure Portal (step 1).

---

## 🔑 Cara Mendapatkan SharePoint List ID

1. Buka SharePoint list yang sudah dibuat
2. Klik **Settings** (gear icon) → **List settings**
3. Salin ID dari URL di address bar: `...settings.aspx?List=%7B**GUID-INI**%7D`
4. Decode `%7B` → `{` dan `%7D` → `}`, atau gunakan Graph Explorer

---

## 👥 Hierarki Visibilitas Kalender

| Role | Yang Terlihat di Kalender |
|------|--------------------------|
| Staff biasa | Diri sendiri + rekan satu atasan (peer) |
| Atasan | Diri sendiri + seluruh bawahan langsung |
| Superuser (`calendarSuperuserEmails`) | Semua karyawan |

---

## 📧 Flow Email Notifikasi

```
User submit request WFA/Visit
        │
        ▼
Status otomatis = Approved
        │
        ▼
Email ke Atasan (To) + CC ke HRD, MIS, User
  Subject: [WFA/Visit Notification] ... (Disetujui Otomatis)
  Isi: detail pengajuan + tombol "Batalkan/Tolak"
        │
        ├── Atasan tidak klik → Selesai, user bisa absen
        │
        └── Atasan klik tombol merah → Buka AttendHQ → Form konfirmasi
                │
                ▼
            Status = Rejected
                │
                ▼
            Email ke User (To) + CC ke HRD, MIS
              Subject: [WFA/Visit Status] ... DIBATALKAN/DITOLAK
```

---

## 🗂️ Tipe Kehadiran

| Tipe | Keterangan | Perlu Pengajuan? |
|------|-----------|-----------------|
| WFO | Work From Office — hadir di kantor | Tidak |
| WFA | Work From Anywhere — bekerja dari luar kantor | Ya |
| Visit | Kunjungan lapangan / meeting eksternal | Ya |

---

## 📱 Instalasi PWA

**Android (Chrome)**
1. Buka link AttendHQ di Chrome
2. Ketuk ⋮ → *Add to Home Screen*
3. Ketuk *Tambahkan*

**iPhone/iPad (Safari)**
1. Buka link AttendHQ di Safari
2. Ketuk ikon Share → *Add to Home Screen*
3. Ketuk *Tambahkan*

---

## 🧩 Dependensi

| Library | Versi | Keterangan |
|---------|-------|-----------|
| [MSAL.js](https://github.com/AzureAD/microsoft-authentication-library-for-js) | 3.x | Autentikasi Microsoft 365, disertakan lokal (`msal-browser.min.js`) |
| Microsoft Graph API | v1.0 | Semua operasi data via REST |

Tidak ada npm/bundler. Semua berjalan sebagai vanilla JS ES Modules langsung di browser.

---

## 🔒 Catatan Keamanan

- Seluruh data tersimpan di SharePoint organisasi — tidak ada server pihak ketiga
- Token autentikasi dikelola oleh MSAL.js dengan mekanisme silent refresh
- Auto-approve dilakukan di frontend; untuk keamanan lebih ketat di masa mendatang dapat dimigrasi ke Power Automate atau backend
- Status notifikasi penolakan disimpan di kolom `Notified_User` SharePoint — tidak bergantung pada localStorage sehingga konsisten lintas device dan browser
- `calendarSuperuserEmails` dikonfigurasi di `config.js` — pastikan hanya email yang terpercaya yang ditambahkan

---

## 📋 Backlog

- [ ] Reminder otomatis via Teams (Cloudflare Worker + Cron + Incoming Webhook)
- [ ] Server-side approval via Power Automate (opsional, keamanan lebih ketat)

---

*AttendHQ — PT. GOS INDORAYA · Head Office · 2026*
