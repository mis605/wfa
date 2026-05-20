// ============================================================
// GRAPH.JS - Service untuk Microsoft Graph API & SharePoint Lists
// ============================================================

import { APP_CONFIG } from './config.js';
import authService from './auth.js';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

class GraphService {

  // Helper: fetch dengan auth token
  async apiFetch(url, options = {}) {
    const token = await authService.getAccessToken();
    
    const defaultHeaders = {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    };

    const response = await fetch(url, {
      ...options,
      headers: { ...defaultHeaders, ...options.headers },
    });

    if (!response.ok) {
      const errorBody = await response.text();
      let errorMsg = `HTTP ${response.status}`;
      try {
        const errorJson = JSON.parse(errorBody);
        errorMsg = errorJson.error?.message || errorMsg;
      } catch {}
      throw new Error(`Graph API Error: ${errorMsg}`);
    }

    if (response.status === 204) return null;
    return response.json();
  }

  // Base URL untuk SharePoint Site
  getSharePointBase() {
    return `${GRAPH_BASE}/sites/${APP_CONFIG.sharepointSiteId}`;
  }

  // Ambil semua item dari SharePoint List
  async getListItems(listId) {
    const url = `${this.getSharePointBase()}/lists/${listId}/items?expand=fields`;
    const data = await this.apiFetch(url);
    return data?.value?.map(item => ({
      listItemId: item.id, // ID internal SharePoint item
      ...item.fields
    })) || [];
  }

  // Tambah item ke SharePoint List
  async addListItem(listId, fields) {
    const url = `${this.getSharePointBase()}/lists/${listId}/items`;
    const body = { fields };
    const response = await this.apiFetch(url, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    return {
      listItemId: response.id,
      ...response.fields
    };
  }

  // Update item di SharePoint List berdasarkan ID internal
  async updateListItem(listId, listItemId, fields) {
    const url = `${this.getSharePointBase()}/lists/${listId}/items/${listItemId}/fields`;
    return this.apiFetch(url, {
      method: 'PATCH',
      body: JSON.stringify(fields)
    });
  }

  // Helper: Timestamp Lokal (WIB) YYYY-MM-DD HH:mm:ss
  getLocalTimestamp() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`;
  }

  // Helper: Cek apakah waktu sekarang dalam rentang yang diizinkan
  isWithinTimeRange(startHHMM, endHHMM) {
    const now = new Date();
    const menitSekarang = now.getHours() * 60 + now.getMinutes();
    const [sh, sm] = startHHMM.split(':').map(Number);
    const [eh, em] = endHHMM.split(':').map(Number);
    return menitSekarang >= sh * 60 + sm && menitSekarang <= eh * 60 + em;
  }

  // ============================================================
  // OPERASI KARYAWAN
  // ============================================================

  async getAllKaryawan() {
    const items = await this.getListItems(APP_CONFIG.listKaryawanId);
    return items.map(item => ({
      id: item.listItemId,
      nip: item.NRK || item.NIP || item.Nip || item.Title || '',
      nama: item.Nama || '',
      email: item.Email || '',
      departemen: item.Departemen || '',
      jabatan: item.Jabatan || '',
      statusAktif: item.Status_Aktif || item.StatusAktif || '',
      emailAtasan: item.Email_Atasan || item.EmailAtasan || '',
      dibuatPada: item.Created || item.Dibuat_Pada || ''
    }));
  }

  async getKaryawanByEmail(email) {
    const items = await this.getAllKaryawan();
    return items.find(k => String(k.email).toLowerCase() === email.toLowerCase()) || null;
  }

  async tambahKaryawan(data) {
    const fields = {
      Title: data.nip,
      NIP: data.nip,
      NRK: data.nip,
      Nama: data.nama,
      Email: data.email,
      Departemen: data.departemen,
      Jabatan: data.jabatan,
      Status_Aktif: 'Aktif',
      StatusAktif: 'Aktif',
      Email_Atasan: data.emailAtasan || '',
      EmailAtasan: data.emailAtasan || ''
    };
    const res = await this.addListItem(APP_CONFIG.listKaryawanId, fields);
    return res.listItemId;
  }

  async updateKaryawan(id, data) {
    const fields = {};
    if (data.nip !== undefined) {
      fields.NIP = data.nip;
      fields.NRK = data.nip;
    }
    if (data.nama !== undefined) fields.Nama = data.nama;
    if (data.email !== undefined) fields.Email = data.email;
    if (data.departemen !== undefined) fields.Departemen = data.departemen;
    if (data.jabatan !== undefined) fields.Jabatan = data.jabatan;
    if (data.statusAktif !== undefined) {
      fields.Status_Aktif = data.statusAktif;
      fields.StatusAktif = data.statusAktif;
    }
    if (data.emailAtasan !== undefined) {
      fields.Email_Atasan = data.emailAtasan;
      fields.EmailAtasan = data.emailAtasan;
    }
    
    await this.updateListItem(APP_CONFIG.listKaryawanId, id, fields);
  }

  // ============================================================
  // OPERASI ABSENSI
  // ============================================================

  async getAbsensiHariIni(nip) {
    const d = new Date();
    const today = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    
    const items = await this.getListItems(APP_CONFIG.listAbsensiId);
    const found = items.find(item => {
      const itemNip = item.NIP || item.Nip || item.NRK || item.Title;
      return String(itemNip) === String(nip) && item.Tanggal === today;
    });
    if (!found) return null;
    
    return {
      id: found.listItemId,
      nip: found.NIP || found.Nip || found.NRK || found.Title || '',
      nama: found.Nama,
      tanggal: found.Tanggal,
      jamMasuk: found.Jam_Masuk || found.JamMasuk,
      jamKeluar: found.Jam_Keluar || found.JamKeluar,
      status: found.Status,
      keterangan: found.Keterangan
    };
  }

  async getAbsensiBulanIni(nip) {
    const now = new Date();
    return this.getAbsensiBulanTertentu(nip, now.getMonth() + 1, now.getFullYear());
  }

  async getAbsensiBulanTertentu(nip, bulan, tahun) {
    const bulanIni = `${tahun}-${String(bulan).padStart(2,'0')}`;
    const items = await this.getListItems(APP_CONFIG.listAbsensiId);
    return items
      .filter(item => {
        const itemNip = item.NIP || item.Nip || item.NRK || item.Title;
        return String(itemNip) === String(nip) && String(item.Tanggal).startsWith(bulanIni);
      })
      .map(item => ({
        id: item.listItemId,
        nip: item.NIP || item.Nip || item.NRK || item.Title || '',
        nama: item.Nama,
        tanggal: item.Tanggal,
        jamMasuk: item.Jam_Masuk || item.JamMasuk,
        jamKeluar: item.Jam_Keluar || item.JamKeluar,
        status: item.Status,
        keterangan: item.Keterangan
      }));
  }

  async getRekapAbsensi(bulan, tahun) {
    const prefix = `${tahun}-${String(bulan).padStart(2,'0')}`;
    const items = await this.getListItems(APP_CONFIG.listAbsensiId);
    return items
      .filter(item => String(item.Tanggal).startsWith(prefix))
      .map(item => ({
        id: item.listItemId,
        nip: item.NIP || item.Nip || item.NRK || item.Title || '',
        nama: item.Nama,
        tanggal: item.Tanggal,
        jamMasuk: item.Jam_Masuk || item.JamMasuk,
        jamKeluar: item.Jam_Keluar || item.JamKeluar,
        status: item.Status,
        keterangan: item.Keterangan
      }));
  }

  async absenMasuk(data) {
    // Validasi jam absen masuk
    if (!this.isWithinTimeRange(APP_CONFIG.jamMasukMulai, APP_CONFIG.jamMasukSelesai)) {
      const [sh, sm] = APP_CONFIG.jamMasukMulai.split(':');
      const [eh, em] = APP_CONFIG.jamMasukSelesai.split(':');
      throw new Error(`Absen masuk hanya bisa dilakukan antara ${sh}:${sm} – ${eh}:${em}.`);
    }
    
    const existing = await this.getAbsensiHariIni(data.nip);
    if (existing) throw new Error('Anda sudah melakukan absen masuk hari ini.');
    
    const now = new Date();
    const tanggal = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
    const jamMasuk = now.toTimeString().substring(0, 8);
    const status = this.hitungStatus(jamMasuk);
    
    const fields = {
      Title: `ABS-${Date.now()}`,
      NIP: data.nip,
      Nama: data.nama,
      Tanggal: tanggal,
      Jam_Masuk: jamMasuk,
      Jam_Keluar: '',
      Status: status,
      Keterangan: data.keterangan || ''
    };
    
    const res = await this.addListItem(APP_CONFIG.listAbsensiId, fields);
    return { id: res.listItemId, status, jamMasuk };
  }

  async absenKeluar(data) {
    if (!this.isWithinTimeRange(APP_CONFIG.jamKeluarMulai, APP_CONFIG.jamKeluarSelesai)) {
      const [sh, sm] = APP_CONFIG.jamKeluarMulai.split(':');
      throw new Error(`Absen keluar hanya bisa dilakukan mulai pukul ${sh}:${sm}.`);
    }
    
    const existing = await this.getAbsensiHariIni(data.nip);
    if (!existing) throw new Error('Anda belum melakukan absen masuk hari ini.');
    if (existing.jamKeluar) throw new Error('Anda sudah melakukan absen keluar hari ini.');
    
    const now = new Date();
    const jamKeluar = now.toTimeString().substring(0, 8);
    
    const fields = {
      Jam_Keluar: jamKeluar,
      Keterangan: data.keterangan || existing.keterangan || ''
    };
    
    await this.updateListItem(APP_CONFIG.listAbsensiId, existing.id, fields);
    return { jamKeluar };
  }

  hitungStatus(jamMasuk) {
    const [h, m] = jamMasuk.split(':').map(Number);
    const menitMasuk = h * 60 + m;
    const [bh, bm] = APP_CONFIG.jamMasukSelesai.split(':').map(Number);
    const batasMenit = bh * 60 + bm;
    const toleransi = APP_CONFIG.toleransiTerlambat;
    
    if (menitMasuk <= batasMenit) return 'Tepat Waktu';
    if (menitMasuk <= batasMenit + toleransi) return 'Terlambat Ringan';
    return 'Terlambat';
  }

  // ============================================================
  // OPERASI PERMOHONAN WFA
  // ============================================================

  async getPermohonanWfa(email) {
    const items = await this.getListItems(APP_CONFIG.listPermohonanWfaId);
    return items
      .filter(item => {
        const itemEmailUser = item.Email_User || item.EmailUser || item.Email || '';
        return String(itemEmailUser).toLowerCase() === email.toLowerCase();
      })
      .map(item => ({
        id: item.listItemId,
        nip: item.NIP || item.Nip || item.NRK || item.Title || '',
        nama: item.Nama || '',
        emailUser: item.Email_User || item.EmailUser || item.Email || '',
        tanggalWfa: item.Tanggal_WFA || item.TanggalWfa || item.Tanggal || '',
        status: item.Status || 'Pending',
        emailAtasan: item.Email_Atasan || item.EmailAtasan || '',
        catatanUser: item.Catatan_User || item.CatatanUser || item.Catatan || '',
        catatanAtasan: item.Catatan_Atasan || item.CatatanAtasan || ''
      }));
  }

  async getPermohonanWfaById(id) {
    const items = await this.getListItems(APP_CONFIG.listPermohonanWfaId);
    const found = items.find(item => String(item.listItemId) === String(id));
    if (!found) return null;
    return {
      id: found.listItemId,
      nip: found.NIP || found.Nip || found.NRK || found.Title || '',
      nama: found.Nama || '',
      emailUser: found.Email_User || found.EmailUser || found.Email || '',
      tanggalWfa: found.Tanggal_WFA || found.TanggalWfa || found.Tanggal || '',
      status: found.Status || 'Pending',
      emailAtasan: found.Email_Atasan || found.EmailAtasan || '',
      catatanUser: found.Catatan_User || found.CatatanUser || found.Catatan || '',
      catatanAtasan: found.Catatan_Atasan || found.CatatanAtasan || ''
    };
  }

  async getApprovedWfaByBulan(bulan, tahun) {
    const items = await this.getListItems(APP_CONFIG.listPermohonanWfaId);
    const prefix = `${tahun}-${String(bulan).padStart(2,'0')}`;
    
    const approvedRequests = items.filter(item => item.Status === 'Approved');
    const result = [];

    approvedRequests.forEach(req => {
      const datesField = req.Tanggal_WFA || req.TanggalWfa || req.Tanggal || '';
      const dates = datesField ? datesField.split(',').map(d => d.trim()) : [];
      dates.forEach(d => {
        if (d.startsWith(prefix)) {
          result.push({
            nip: req.NIP || req.Nip || req.NRK || req.Title || '',
            nama: req.Nama || '',
            tanggal: d
          });
        }
      });
    });

    return result;
  }

  async tambahPermohonanWfa(data) {
    const fields = {
      Title: `REQ-${Date.now()}`,
      NIP: data.nip,
      Nama: data.nama,
      Email_User: data.emailUser,
      Tanggal_WFA: data.tanggalWfa, // String terpisah koma
      Status: 'Pending',
      Email_Atasan: data.emailAtasan,
      Catatan_User: data.catatanUser || '',
      Catatan_Atasan: ''
    };
    
    const res = await this.addListItem(APP_CONFIG.listPermohonanWfaId, fields);
    
    // Kirim email persetujuan ke atasan
    try {
      await this.sendApprovalEmail({
        listItemId: res.listItemId,
        ...fields
      });
    } catch (e) {
      console.error("Gagal mengirim email persetujuan ke atasan:", e);
    }

    return res.listItemId;
  }

  async updateStatusPermohonanWfa(id, status, catatanAtasan = '') {
    const fields = {
      Status: status,
      Catatan_Atasan: catatanAtasan
    };
    await this.updateListItem(APP_CONFIG.listPermohonanWfaId, id, fields);
    
    // Kirim email balasan ke user pemohon
    const req = await this.getPermohonanWfaById(id);
    if (req) {
      try {
        await this.sendResponseEmail(req, status, catatanAtasan);
      } catch (e) {
        console.error("Gagal mengirim email balasan ke pemohon:", e);
      }
    }
  }

  // ============================================================
  // OUTLOOK EMAIL DISPATCH via GRAPH API
  // ============================================================

  async sendMail(message) {
    const url = `${GRAPH_BASE}/me/sendMail`;
    return this.apiFetch(url, {
      method: 'POST',
      body: JSON.stringify({
        message,
        saveToSentItems: 'true'
      })
    });
  }

  async sendApprovalEmail(req) {
    const formattedDates = req.Tanggal_WFA.split(',').map(d => `• ${d.trim()}`).join('<br>');
    const approveUrl = `${APP_CONFIG.redirectUri}?action=approve&id=${req.listItemId}`;
    const rejectUrl = `${APP_CONFIG.redirectUri}?action=reject&id=${req.listItemId}`;

    const mailContent = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px; background-color: #ffffff;">
        <h2 style="color: #2b6cb0; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px;">Pengajuan Work From Anywhere (WFA)</h2>
        <p>Halo,</p>
        <p>Karyawan berikut mengajukan permohonan WFA:</p>
        
        <table style="width: 100%; margin: 20px 0; border-collapse: collapse;">
          <tr style="background-color: #f7fafc;">
            <td style="padding: 10px; font-weight: bold; width: 150px;">Nama Karyawan</td>
            <td style="padding: 10px;">${req.Nama}</td>
          </tr>
          <tr>
            <td style="padding: 10px; font-weight: bold;">NIP</td>
            <td style="padding: 10px;">${req.NIP}</td>
          </tr>
          <tr style="background-color: #f7fafc;">
            <td style="padding: 10px; font-weight: bold; vertical-align: top;">Tanggal WFA</td>
            <td style="padding: 10px; color: #2d3748;">${formattedDates}</td>
          </tr>
          <tr>
            <td style="padding: 10px; font-weight: bold; vertical-align: top;">Catatan/Alasan</td>
            <td style="padding: 10px; color: #718096; font-style: italic;">"${req.Catatan_User || '-'}"</td>
          </tr>
        </table>

        <p style="margin-top: 30px; font-weight: 500;">Silakan pilih respon tindakan untuk pengajuan ini:</p>
        <div style="margin: 20px 0; display: flex; gap: 12px;">
          <a href="${approveUrl}" style="display: inline-block; background-color: #48bb78; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 0.95rem;">Setujui (Approve)</a>
          <a href="${rejectUrl}" style="display: inline-block; background-color: #f56565; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 0.95rem; margin-left: 10px;">Tolak (Reject)</a>
        </div>

        <p style="font-size: 0.85rem; color: #a0aec0; margin-top: 40px; border-top: 1px solid #e2e8f0; padding-top: 15px;">
          Catatan: Tombol di atas akan membuka web aplikasi Absensi WFA CO untuk memproses aksi persetujuan. Anda perlu masuk menggunakan akun Microsoft 365 Anda.
        </p>
      </div>
    `;

    const message = {
      subject: `[WFA Request] Pengajuan WFA - ${req.Nama}`,
      body: {
        contentType: 'HTML',
        content: mailContent
      },
      toRecipients: [
        {
          emailAddress: {
            address: req.Email_Atasan
          }
        }
      ]
    };

    return this.sendMail(message);
  }

  async sendResponseEmail(req, status, catatanAtasan) {
    const isApproved = status === 'Approved';
    const statusLabel = isApproved ? 'DISETUJUI' : 'DITOLAK';
    const color = isApproved ? '#48bb78' : '#f56565';
    const formattedDates = req.tanggalWfa.split(',').map(d => `• ${d.trim()}`).join('<br>');

    const mailContent = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px; background-color: #ffffff;">
        <h2 style="color: ${color}; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px;">Status Pengajuan WFA - ${statusLabel}</h2>
        <p>Halo ${req.nama},</p>
        <p>Pengajuan WFA Anda telah ditinjau oleh atasan dengan status hasil berikut:</p>
        
        <table style="width: 100%; margin: 20px 0; border-collapse: collapse;">
          <tr style="background-color: #f7fafc;">
            <td style="padding: 10px; font-weight: bold; width: 150px;">Status</td>
            <td style="padding: 10px; color: ${color}; font-weight: bold;">${statusLabel}</td>
          </tr>
          <tr>
            <td style="padding: 10px; font-weight: bold; vertical-align: top;">Tanggal Diajukan</td>
            <td style="padding: 10px; color: #2d3748;">${formattedDates}</td>
          </tr>
          <tr style="background-color: #f7fafc;">
            <td style="padding: 10px; font-weight: bold; vertical-align: top;">Catatan Atasan</td>
            <td style="padding: 10px; color: #2d3748;">${catatanAtasan || '-'}</td>
          </tr>
        </table>

        <p style="margin-top: 30px;">
          ${isApproved ? 'Silakan melakukan absensi masuk/pulang di webapp pada tanggal yang disetujui tersebut.' : 'Silakan hubungi atasan Anda jika ada pertanyaan lebih lanjut.'}
        </p>
        
        <p style="font-size: 0.85rem; color: #a0aec0; margin-top: 40px; border-top: 1px solid #e2e8f0; padding-top: 15px;">
          Email ini dikirim secara otomatis oleh Sistem Absensi WFA CO PT. GOS INDORAYA.
        </p>
      </div>
    `;

    const message = {
      subject: `[WFA Status] Pengajuan WFA Anda telah ${statusLabel}`,
      body: {
        contentType: 'HTML',
        content: mailContent
      },
      toRecipients: [
        {
          emailAddress: {
            address: req.emailUser
          }
        }
      ]
    };

    return this.sendMail(message);
  }
}

const graphService = new GraphService();
export default graphService;
