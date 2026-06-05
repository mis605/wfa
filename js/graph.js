// ============================================================
// GRAPH.JS - Service untuk Microsoft Graph API & SharePoint Lists
// ============================================================

import { APP_CONFIG } from './config.js';
import authService from './auth.js';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

class GraphService {

  // Helper: fetch dengan auth token
  async apiFetch(url, options = {}) {
    console.log('[GraphService] apiFetch started for:', url);
    let token;
    try {
      token = await authService.getAccessToken();
      console.log('[GraphService] Token obtained successfully');
    } catch (tokenError) {
      console.error('[GraphService] Failed to obtain token:', tokenError);
      throw tokenError;
    }
    
    const defaultHeaders = {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    };

    let response;
    try {
      response = await fetch(url, {
        ...options,
        headers: { ...defaultHeaders, ...options.headers },
      });
      console.log('[GraphService] Response received. Status:', response.status);
    } catch (fetchError) {
      console.error('[GraphService] Fetch network error:', fetchError);
      throw fetchError;
    }

    if (!response.ok) {
      const errorBody = await response.text();
      let errorMsg = `HTTP ${response.status}`;
      try {
        const errorJson = JSON.parse(errorBody);
        errorMsg = errorJson.error?.message || errorMsg;
      } catch {}
      console.error('[GraphService] Graph API Error response:', errorBody);
      throw new Error(`Graph API Error: ${errorMsg}`);
    }

    if (response.status === 204) return null;
    return response.json();
  }

  getSharePointBase() {
    return `${GRAPH_BASE}/sites/${APP_CONFIG.sharepointSiteId}`;
  }

  async getListItems(listId) {
    console.log('[GraphService] getListItems started for list:', listId);
    const url = `${this.getSharePointBase()}/lists/${listId}/items?expand=fields`;
    const data = await this.apiFetch(url);
    console.log('[GraphService] getListItems count:', data?.value?.length);
    return data?.value?.map(item => ({
      listItemId: item.id,
      ...item.fields
    })) || [];
  }

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

  async updateListItem(listId, listItemId, fields) {
    const url = `${this.getSharePointBase()}/lists/${listId}/items/${listItemId}/fields`;
    return this.apiFetch(url, {
      method: 'PATCH',
      body: JSON.stringify(fields)
    });
  }

  async debugListColumns(listId, listName) {
    console.log(`[Debug] Fetching columns for list ${listName} (${listId})...`);
    const url = `${this.getSharePointBase()}/lists/${listId}/columns`;
    try {
      const data = await this.apiFetch(url);
      if (data && data.value) {
        const cols = data.value.map(c => ({
          name: c.name,
          displayName: c.displayName
        }));
        console.log(`[Debug] Columns for list ${listName}:`, cols);
      }
    } catch (err) {
      console.error(`[Debug] Error fetching columns for ${listName}:`, err);
    }
  }

  // Helper: tanggal lokal WIB (bukan UTC)
  getTodayLocal() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  getLocalTimestamp() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}`;
  }

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

  async getKaryawanByNip(nip) {
    const items = await this.getAllKaryawan();
    return items.find(k => String(k.nip) === String(nip)) || null;
  }

  async tambahKaryawan(data) {
    const fields = {
      Title: data.nip,
      Nama: data.nama,
      Email: data.email,
      Departemen: data.departemen,
      Jabatan: data.jabatan,
      Status_Aktif: 'Aktif',
      Email_Atasan: data.emailAtasan || ''
    };
    const res = await this.addListItem(APP_CONFIG.listKaryawanId, fields);
    return res.listItemId;
  }

  async updateKaryawan(id, data) {
    const fields = {};
    if (data.nip !== undefined) fields.Title = data.nip;
    if (data.nama !== undefined) fields.Nama = data.nama;
    if (data.email !== undefined) fields.Email = data.email;
    if (data.departemen !== undefined) fields.Departemen = data.departemen;
    if (data.jabatan !== undefined) fields.Jabatan = data.jabatan;
    if (data.statusAktif !== undefined) fields.Status_Aktif = data.statusAktif;
    if (data.emailAtasan !== undefined) fields.Email_Atasan = data.emailAtasan;
    await this.updateListItem(APP_CONFIG.listKaryawanId, id, fields);
  }

  // ============================================================
  // OPERASI ABSENSI
  // Kolom baru: Tipe (WFO / WFA / Visit), Email_Atasan
  // ============================================================

  _mapAbsensiItem(item) {
    return {
      id: item.listItemId,
      nip: item.NIP || item.Nip || item.NRK || item.Title || '',
      nama: item.Nama || '',
      tanggal: item.Tanggal || '',
      jamMasuk: item.Jam_Masuk || item.JamMasuk || '',
      jamKeluar: item.Jam_Keluar || item.JamKeluar || '',
      status: item.Status || '',
      tipe: item.Tipe || 'WFO',
      emailAtasan: item.Email_Atasan || item.EmailAtasan || '',
      keterangan: item.Keterangan || ''
    };
  }

  async getAbsensiHariIni(nip) {
    const today = this.getTodayLocal();
    const items = await this.getListItems(APP_CONFIG.listAbsensiId);
    const found = items.find(item => {
      const itemNip = item.NIP || item.Nip || item.NRK || item.Title;
      return String(itemNip) === String(nip) && item.Tanggal === today;
    });
    if (!found) return null;
    return this._mapAbsensiItem(found);
  }

  async getAbsensiBulanIni(nip) {
    const now = new Date();
    return this.getAbsensiBulanTertentu(nip, now.getMonth() + 1, now.getFullYear());
  }

  async getAbsensiBulanTertentu(nip, bulan, tahun) {
    const prefix = `${tahun}-${String(bulan).padStart(2,'0')}`;
    const items = await this.getListItems(APP_CONFIG.listAbsensiId);
    return items
      .filter(item => {
        const itemNip = item.NIP || item.Nip || item.NRK || item.Title;
        return String(itemNip) === String(nip) && String(item.Tanggal).startsWith(prefix);
      })
      .map(item => this._mapAbsensiItem(item));
  }

  // Ambil SEMUA absensi bulan tertentu (untuk kalender — semua karyawan)
  async getAbsensiBulanTertentu_All(bulan, tahun) {
    const prefix = `${tahun}-${String(bulan).padStart(2,'0')}`;
    const items = await this.getListItems(APP_CONFIG.listAbsensiId);
    return items
      .filter(item => String(item.Tanggal || '').startsWith(prefix))
      .map(item => this._mapAbsensiItem(item));
  }

  async absenMasuk(data) {
    if (!this.isWithinTimeRange(APP_CONFIG.jamMasukMulai, APP_CONFIG.jamMasukSelesai)) {
      const [sh, sm] = APP_CONFIG.jamMasukMulai.split(':');
      const [eh, em] = APP_CONFIG.jamMasukSelesai.split(':');
      throw new Error(`Absen masuk hanya bisa dilakukan antara ${sh}:${sm} – ${eh}:${em}.`);
    }
    
    const existing = await this.getAbsensiHariIni(data.nip);
    if (existing) throw new Error('Anda sudah melakukan absen masuk hari ini.');
    
    const now = new Date();
    const tanggal = this.getTodayLocal();
    const jamMasuk = now.toTimeString().substring(0, 8);
    const status = this.hitungStatus(jamMasuk);
    
    const fields = {
      Title: data.nip,
      Nama: data.nama,
      Tanggal: tanggal,
      Jam_Masuk: jamMasuk,
      Jam_Keluar: '',
      Status: status,
      Tipe: data.tipe || 'WFO',
      Email_Atasan: data.emailAtasan || '',
      Keterangan: data.keterangan || ''
    };
    
    const res = await this.addListItem(APP_CONFIG.listAbsensiId, fields);
    return { id: res.listItemId, status, jamMasuk, tipe: data.tipe };
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
  // OPERASI PERMOHONAN WFA / VISIT
  // Kolom baru: Tipe (WFA / Visit)
  // ============================================================

  _mapPermohonanItem(item) {
    return {
      id: item.listItemId,
      nip: item.NIP || item.Nip || item.NRK || item.Title || '',
      nama: item.Nama || '',
      emailUser: item.Email_User || item.EmailUser || item.Email || '',
      tanggalWfa: item.Tanggal_WFA || item.TanggalWfa || item.Tanggal || '',
      tipe: item.Tipe || 'WFA',
      status: item.Status || 'Pending',
      emailAtasan: item.Email_Atasan || item.EmailAtasan || '',
      catatanUser: item.Catatan_User || item.CatatanUser || item.Catatan || '',
      catatanAtasan: item.Catatan_Atasan || item.CatatanAtasan || ''
    };
  }

  async getPermohonanWfa(email) {
    const items = await this.getListItems(APP_CONFIG.listPermohonanWfaId);
    return items
      .filter(item => {
        const itemEmail = item.Email_User || item.EmailUser || item.Email || '';
        return String(itemEmail).toLowerCase() === email.toLowerCase();
      })
      .map(item => this._mapPermohonanItem(item));
  }

  async getPermohonanWfaById(id) {
    const items = await this.getListItems(APP_CONFIG.listPermohonanWfaId);
    const found = items.find(item => String(item.listItemId) === String(id));
    if (!found) return null;
    return this._mapPermohonanItem(found);
  }

  // Ambil semua request approved (WFA + Visit) untuk bulan tertentu — untuk kalender
  async getApprovedRequestByBulan(bulan, tahun) {
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
            tanggal: d,
            tipe: req.Tipe || 'WFA',
            emailAtasan: (req.Email_Atasan || req.EmailAtasan || '').toLowerCase().trim()
          });
        }
      });
    });

    return result;
  }

  // (Backward compat) alias untuk getApprovedRequestByBulan
  async getApprovedWfaByBulan(bulan, tahun) {
    return this.getApprovedRequestByBulan(bulan, tahun);
  }

  async tambahPermohonanWfa(data) {
    const fields = {
      Title: data.nip,
      Nama: data.nama,
      Email_User: data.emailUser,
      Tanggal_WFA: data.tanggalWfa,
      Tipe: data.tipe || 'WFA',
      Status: 'Approved',  // Auto approve
      Email_Atasan: data.emailAtasan,
      Catatan_User: data.catatanUser || '',
      Catatan_Atasan: ''
    };
    
    const res = await this.addListItem(APP_CONFIG.listPermohonanWfaId, fields);
    
    // Kirim email notifikasi ke atasan
    try {
      await this.sendApprovalEmail({
        listItemId: res.listItemId,
        ...fields
      });
    } catch (e) {
      console.error('Gagal mengirim email notifikasi ke atasan:', e);
    }

    return res.listItemId;
  }

  async updateStatusPermohonanWfa(id, status, catatanAtasan = '') {
    const fields = {
      Status: status,
      Catatan_Atasan: catatanAtasan
    };
    await this.updateListItem(APP_CONFIG.listPermohonanWfaId, id, fields);
    
    // Kirim email balasan ke user
    const req = await this.getPermohonanWfaById(id);
    if (req) {
      try {
        await this.sendResponseEmail(req, status, catatanAtasan);
      } catch (e) {
        console.error('Gagal mengirim email balasan ke pemohon:', e);
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
    const tipe = req.Tipe || 'WFA';
    const formattedDates = req.Tanggal_WFA.split(',').map(d => `• ${d.trim()}`).join('<br>');
    const rejectUrl = `${APP_CONFIG.redirectUri}?action=reject&id=${req.listItemId}`;

    const tipeColor = tipe === 'WFA' ? '#4299e1' : '#ed8936';
    const tipeLabel = tipe === 'Visit' ? 'Visit / Kunjungan Lapangan' : 'Work From Anywhere (WFA)';

    const mailContent = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px; background-color: #ffffff;">
        <div style="background: ${tipeColor}; color: white; padding: 12px 20px; border-radius: 6px 6px 0 0; margin: -20px -20px 20px -20px;">
          <h2 style="margin: 0; font-size: 1.1rem;">Pemberitahuan ${tipeLabel}</h2>
        </div>
        <p>Halo,</p>
        <p>Karyawan berikut telah mengajukan permohonan <strong>${tipe}</strong> dan statusnya telah <strong>disetujui secara otomatis oleh sistem</strong>:</p>
        
        <table style="width: 100%; margin: 20px 0; border-collapse: collapse;">
          <tr style="background-color: #f7fafc;">
            <td style="padding: 10px; font-weight: bold; width: 160px;">Nama Karyawan</td>
            <td style="padding: 10px;">${req.Nama}</td>
          </tr>
          <tr>
            <td style="padding: 10px; font-weight: bold;">NIP / NRK</td>
            <td style="padding: 10px;">${req.NRK || req.NIP || req.Title || ''}</td>
          </tr>
          <tr style="background-color: #f7fafc;">
            <td style="padding: 10px; font-weight: bold;">Tipe</td>
            <td style="padding: 10px;">
              <span style="background: ${tipeColor}; color: white; padding: 2px 10px; border-radius: 20px; font-size: 0.85rem; font-weight: bold;">${tipe}</span>
            </td>
          </tr>
          <tr>
            <td style="padding: 10px; font-weight: bold; vertical-align: top;">Tanggal</td>
            <td style="padding: 10px; color: #2d3748;">${formattedDates}</td>
          </tr>
          <tr style="background-color: #f7fafc;">
            <td style="padding: 10px; font-weight: bold; vertical-align: top;">Alasan</td>
            <td style="padding: 10px; color: #718096; font-style: italic;">"${req.Catatan_User || '-'}"</td>
          </tr>
        </table>

        <p style="margin-top: 30px; font-weight: 500; color: #2d3748;">
          Jika Anda menyetujui pengajuan ini, <strong>tidak perlu melakukan tindakan apapun</strong>.
          Namun, jika Anda ingin <strong>membatalkan / menolak</strong> pengajuan ini, klik tombol di bawah:
        </p>
        <div style="margin: 20px 0;">
          <a href="${rejectUrl}" style="display: inline-block; background-color: #f56565; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 0.95rem;">Batalkan / Tolak Pengajuan</a>
        </div>

        <p style="font-size: 0.85rem; color: #a0aec0; margin-top: 40px; border-top: 1px solid #e2e8f0; padding-top: 15px;">
          Catatan: Tombol di atas akan membuka web aplikasi Absensi untuk membatalkan pengajuan. Anda perlu masuk menggunakan akun Microsoft 365.
        </p>
      </div>
    `;

    const userEmail = req.Email_User || req.EmailUser || req.Email || '';
    const toAddress = (req.Email_Atasan || '').toLowerCase().trim();
    const rawCcList = [APP_CONFIG.emailHrd, APP_CONFIG.emailMis, userEmail];

    const ccEmails = new Set();
    const ccRecipients = [];
    for (const rawCc of rawCcList) {
      if (!rawCc) continue;
      const cleanCc = rawCc.toLowerCase().trim();
      if (cleanCc && cleanCc !== toAddress && !ccEmails.has(cleanCc)) {
        ccEmails.add(cleanCc);
        ccRecipients.push({ emailAddress: { address: cleanCc } });
      }
    }

    const message = {
      subject: `[${tipe} Notification] Pengajuan ${tipe} - ${req.Nama} (Disetujui Otomatis)`,
      body: { contentType: 'HTML', content: mailContent },
      toRecipients: [{ emailAddress: { address: req.Email_Atasan } }]
    };

    if (ccRecipients.length > 0) {
      message.ccRecipients = ccRecipients;
    }

    return this.sendMail(message);
  }

  async sendResponseEmail(req, status, catatanAtasan) {
    const isApproved = status === 'Approved';
    const statusLabel = isApproved ? 'DISETUJUI' : 'DIBATALKAN / DITOLAK';
    const color = isApproved ? '#48bb78' : '#f56565';
    const tipe = req.tipe || 'WFA';
    const formattedDates = req.tanggalWfa.split(',').map(d => `• ${d.trim()}`).join('<br>');

    const mailContent = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px; background-color: #ffffff;">
        <h2 style="color: ${color}; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px;">Status Pengajuan ${tipe} - ${statusLabel}</h2>
        <p>Halo ${req.nama},</p>
        <p>Pengajuan <strong>${tipe}</strong> Anda telah ditinjau oleh atasan dengan hasil berikut:</p>
        
        <table style="width: 100%; margin: 20px 0; border-collapse: collapse;">
          <tr style="background-color: #f7fafc;">
            <td style="padding: 10px; font-weight: bold; width: 150px;">Status</td>
            <td style="padding: 10px; color: ${color}; font-weight: bold;">${statusLabel}</td>
          </tr>
          <tr>
            <td style="padding: 10px; font-weight: bold;">Tipe</td>
            <td style="padding: 10px;">${tipe}</td>
          </tr>
          <tr style="background-color: #f7fafc;">
            <td style="padding: 10px; font-weight: bold; vertical-align: top;">Tanggal</td>
            <td style="padding: 10px; color: #2d3748;">${formattedDates}</td>
          </tr>
          <tr>
            <td style="padding: 10px; font-weight: bold; vertical-align: top;">Catatan Atasan</td>
            <td style="padding: 10px; color: #2d3748;">${catatanAtasan || '-'}</td>
          </tr>
        </table>

        <p style="margin-top: 30px;">
          ${isApproved
            ? `Silakan melakukan absensi masuk/pulang di webapp pada tanggal yang disetujui.`
            : `Pengajuan ${tipe} Anda telah dibatalkan/ditolak oleh atasan. Silakan hubungi atasan Anda jika ada pertanyaan.`}
        </p>
        
        <p style="font-size: 0.85rem; color: #a0aec0; margin-top: 40px; border-top: 1px solid #e2e8f0; padding-top: 15px;">
          Email ini dikirim secara otomatis oleh Sistem Absensi ${APP_CONFIG.namaPerusahaan}.
        </p>
      </div>
    `;

    const toAddress = (req.emailUser || '').toLowerCase().trim();
    const rawCcList = [APP_CONFIG.emailHrd, APP_CONFIG.emailMis];
    const ccEmails = new Set();
    const ccRecipients = [];
    for (const rawCc of rawCcList) {
      if (!rawCc) continue;
      const cleanCc = rawCc.toLowerCase().trim();
      if (cleanCc && cleanCc !== toAddress && !ccEmails.has(cleanCc)) {
        ccEmails.add(cleanCc);
        ccRecipients.push({ emailAddress: { address: cleanCc } });
      }
    }

    const message = {
      subject: `[${tipe} Status] Pengajuan ${tipe} Anda telah ${statusLabel}`,
      body: { contentType: 'HTML', content: mailContent },
      toRecipients: [{ emailAddress: { address: req.emailUser } }]
    };

    if (ccRecipients.length > 0) {
      message.ccRecipients = ccRecipients;
    }

    return this.sendMail(message);
  }
}

const graphService = new GraphService();
export default graphService;
