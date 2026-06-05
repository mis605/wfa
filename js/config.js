// ============================================================
// CONFIG.JS - Konfigurasi utama aplikasi
// Ganti nilai di bawah sesuai tenant M365 Anda
// ============================================================

const APP_CONFIG = {
  // --- AZURE AD APP REGISTRATION ---
  // Daftar di: https://portal.azure.com > App registrations
  clientId: "c87df5fb-eb2f-40bc-8c33-d0c47dedfe10",
  tenantId: "c1cc324d-9fda-4a48-97ee-ba6157ba9c67",

  // Redirect URI harus sama persis dengan yang didaftarkan di Azure AD
  redirectUri: window.location.origin + window.location.pathname,

  // --- SHAREPOINT LISTS CONFIG ---
  // ID Site SharePoint (contoh: "tenant.sharepoint.com,xxxx-xxxx,yyyy-yyyy")
  sharepointSiteId: "gosgroup.sharepoint.com,75e3dc18-6245-4db0-9977-3b9bfdb582a2,c6ddcd82-3462-4bfb-b375-12e1c12a3e93",
  
  // ID / Nama List di SharePoint
  listAbsensiId: "2400574a-7ace-4c53-a472-afe1a991f68a",
  listKaryawanId: "8d0674a3-05f6-426e-8d45-4cacc750ae4b",
  listPermohonanWfaId: "d8137f93-ff19-4626-9ef5-4cb68bcf59c9",

  // --- PENGATURAN ABSENSI ---
  jamMasukMulai: "07:00",   // Jam mulai bisa absen masuk
  jamMasukSelesai: "10:00", // Jam batas absen masuk (lewat = terlambat)
  jamKeluarMulai: "16:00",  // Jam mulai bisa absen keluar
  jamKeluarSelesai: "20:00",// Jam batas absen keluar

  // Toleransi terlambat dalam menit
  toleransiTerlambat: 15,

  // Nama perusahaan
  namaPerusahaan: "PT. GOS INDORAYA",
  logoPerusahaan: "", // URL logo (opsional)

  // --- EMAIL NOTIFICATION CONFIG ---
  emailHrd: "hc.info@gos.co.id",
  emailMis: "mis@gos.co.id",

  // --- KALENDER SUPERUSER ---
  // Akun yang bisa melihat SEMUA karyawan di Tab Kalender
  // Tambah email lain jika perlu (HRD, direktur, dll)
  calendarSuperuserEmails: [
    "mis@gos.co.id",
  ],
};

// Microsoft Graph API scopes yang dibutuhkan
const GRAPH_SCOPES = [
  "User.Read",
  "Sites.ReadWrite.All",
  "Mail.Send"
];

// MSAL Configuration
const MSAL_CONFIG = {
  auth: {
    clientId: APP_CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${APP_CONFIG.tenantId}`,
    redirectUri: APP_CONFIG.redirectUri,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        if (level === msal.LogLevel.Error) console.error(message);
      },
    },
  },
};

export { APP_CONFIG, MSAL_CONFIG, GRAPH_SCOPES };
