// ============================================================
// CONFIG.JS - Konfigurasi utama aplikasi
// Ganti nilai di bawah sesuai tenant M365 Anda
// ============================================================

const APP_CONFIG = {
  // --- AZURE AD APP REGISTRATION ---
  // Daftar di: https://portal.azure.com > App registrations
  clientId: "efa1afcc-0850-4433-a107-bec0fd0ee282",
  tenantId: "5194178e-ce02-4b0c-8442-1374fd7eca0d",

  // Redirect URI harus sama persis dengan yang didaftarkan di Azure AD
  redirectUri: window.location.origin + window.location.pathname,

  // --- SHAREPOINT LISTS CONFIG ---
  // ID Site SharePoint (contoh: "tenant.sharepoint.com,xxxx-xxxx,yyyy-yyyy")
  sharepointSiteId: "tenant.sharepoint.com,ganti-dengan-site-guid-1,ganti-dengan-site-guid-2",
  
  // ID / Nama List di SharePoint
  listAbsensiId: "Absensi",
  listKaryawanId: "Karyawan",
  listPermohonanWfaId: "PermohonanWfa",

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
