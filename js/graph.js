// ============================================================
// GRAPH.JS - Service untuk Microsoft Graph API & SharePoint Lists
// Revisi: server-side $filter, caching 5 menit, join karyawan untuk emailAtasan
// ============================================================

import { APP_CONFIG } from './config.js';
import authService from './auth.js';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// ============================================================
// SIMPLE IN-MEMORY CACHE (TTL 5 menit)
// ============================================================
const _cache = new Map();
const CACHE_TTL = 5 * 60 * 1000;

function cacheGet(key) {
  const entry = _cache.get(key);
  if (!entry) return null;
  if (Date.now() - entry.ts > CACHE_TTL) { _cache.delete(key); return null; }
  return entry.data;
}

function cacheSet(key, data) {
  _cache.set(key, { data, ts: Date.now() });
  return data;
}

function cacheInvalidate(pattern) {
  for (const key of _cache.keys()) {
    if (key.includes(pattern)) _cache.delete(key);
  }
}

class GraphService {

  async apiFetch(url, options = {}) {
    let token;
    try {
      token = await authService.getAccessToken();
    } catch (err) {
      console.error('[GraphService] Token error:', err);
      throw err;
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
    } catch (err) {
      console.error('[GraphService] Network error:', err);
      throw err;
    }

    if (!response.ok) {
      const errorBody = await response.text();
      let errorMsg = `HTTP ${response.status}`;
      try { errorMsg = JSON.parse(errorBody).error?.message || errorMsg; } catch {}
      console.error('[GraphService] API Error:', errorMsg);
      throw new Error(`Graph API Error: ${errorMsg}`);
    }

    if (response.status === 204) return null;
    return response.json();
  }

  getSharePointBase() {
    return `${GRAPH_BASE}/sites/${APP_CONFIG.sharepointSiteId}`;
  }

  // Ambil semua item tanpa filter — dengan cache
  async getListItems(listId, cacheKey = null) {
    const key = cacheKey || `list:${listId}:all`;
    const cached = cacheGet(key);
    if (cached) return cached;

    const url = `${this.getSharePointBase()}/lists/${listId}/items?expand=fields&$top=5000`;
    const data = await this.apiFetch(url);
    const items = data?.value?.map(item => ({ ...item.fields, listItemId: item.id })) || [];
    return cacheSet(key, items);
  }

  // Ambil items dengan $filter server-side — fallback ke full fetch jika kolom belum diindex
  async getListItemsFiltered(listId, filterQuery, cacheKey) {
    const cached = cacheGet(cacheKey);
    if (cached) return cached;

    const url = `${this.getSharePointBase()}/lists/${listId}/items?expand=fields&$top=5000&$filter=${encodeURIComponent(filterQuery)}`;
    let data;
    try {
      data = await this.apiFetch(url);
    } catch (err) {
      console.warn('[GraphService] $filter failed, fallback to full fetch. Aktifkan index pada kolom SharePoint untuk performa optimal:', err.message);
      return this.getListItems(listId, cacheKey);
    }

    const items = data?.value?.map(item => ({ ...item.fields, listItemId: item.id })) || [];
    return cacheSet(cacheKey, items);
  }

  async addListItem(listId, fields) {
    const url = `${this.getSharePointBase()}/lists/${listId}/items`;
    const response = await this.apiFetch(url, {
      method: 'POST',
      body: JSON.stringify({ fields })
    });
    cacheInvalidate(`list:${listId}`);
    return { listItemId: response.id, ...response.fields };
  }

  async updateListItem(listId, listItemId, fields) {
    const url = `${this.getSharePointBase()}/lists/${listId}/items/${listItemId}/fields`;
    const result = await this.apiFetch(url, {
      method: 'PATCH',
      body: JSON.stringify(fields)
    });
    cacheInvalidate(`list:${listId}`);
    return result;
  }

  // Tanggal lokal WIB (bukan UTC)
  getTodayLocal() {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
  }

  isWithinTimeRange(startHHMM, endHHMM) {
    const now = new Date();
    const cur = now.getHours() * 60 + now.getMinutes();
    const [sh, sm] = startHHMM.split(':').map(Number);
    const [eh, em] = endHHMM.split(':').map(Number);
    return cur >= sh * 60 + sm && cur <= eh * 60 + em;
  }

  _getMonthRange(bulan, tahun) {
    const firstDay = `${tahun}-${String(bulan).padStart(2,'0')}-01`;
    const lastDayDate = new Date(tahun, bulan, 0);
    const lastDay = `${tahun}-${String(bulan).padStart(2,'0')}-${String(lastDayDate.getDate()).padStart(2,'0')}`;
    return { firstDay, lastDay };
  }

  // ============================================================
  // KARYAWAN
  // ============================================================

  async getAllKaryawan() {
    const items = await this.getListItems(APP_CONFIG.listKaryawanId, 'list:karyawan:all');
    return items.map(item => ({
      id: item.listItemId,
      nip: item.NRK || item.NIP || item.Nip || item.Title || '',
      nama: item.Nama || '',
      email: item.Email || '',
      departemen: item.Departemen || '',
      jabatan: item.Jabatan || '',
      statusAktif: item.Status_Aktif || item.StatusAktif || '',
      emailAtasan: item.Email_Atasan || item.EmailAtasan || '',
    }));
  }

  async getKaryawanByEmail(email) {
    const items = await this.getAllKaryawan();
    return items.find(k => String(k.email).toLowerCase() === email.toLowerCase()) || null;
  }

  // Map NIP → emailAtasan dari list karyawan (single source of truth)
  async buildNipAtasanMap() {
    const cached = cacheGet('map:nip-atasan');
    if (cached) return cached;
    const karyawan = await this.getAllKaryawan();
    const map = {};
    karyawan.forEach(k => {
      if (k.nip) map[String(k.nip)] = (k.emailAtasan || '').toLowerCase().trim();
    });
    return cacheSet('map:nip-atasan', map);
  }

  async tambahKaryawan(data) {
    const fields = {
      Title: data.nip, Nama: data.nama, Email: data.email,
      Departemen: data.departemen, Jabatan: data.jabatan,
      Status_Aktif: 'Aktif', Email_Atasan: data.emailAtasan || ''
    };
    const res = await this.addListItem(APP_CONFIG.listKaryawanId, fields);
    cacheInvalidate('list:karyawan'); cacheInvalidate('map:nip-atasan');
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
    cacheInvalidate('list:karyawan'); cacheInvalidate('map:nip-atasan');
  }

  // ============================================================
  // ABSENSI
  // ============================================================

  _mapAbsensiItem(item, nipAtasanMap = null) {
    const nip = item.NIP || item.Nip || item.NRK || item.Title || '';
    // emailAtasan: dari kolom list dulu, fallback ke lookup map karyawan
    const emailAtasanFromList = item.Email_Atasan || item.EmailAtasan || '';
    const emailAtasanFromMap = nipAtasanMap ? (nipAtasanMap[String(nip)] || '') : '';
    return {
      id: item.listItemId, nip, nama: item.Nama || '',
      tanggal: item.Tanggal || '',
      jamMasuk: item.Jam_Masuk || item.JamMasuk || '',
      jamKeluar: item.Jam_Keluar || item.JamKeluar || '',
      status: item.Status || '',
      tipe: item.Tipe || 'WFO',
      emailAtasan: emailAtasanFromList || emailAtasanFromMap,
      keterangan: item.Keterangan || ''
    };
  }

  async getAbsensiHariIni(nip) {
    const today = this.getTodayLocal();
    const cacheKey = `list:${APP_CONFIG.listAbsensiId}:tanggal:${today}`;
    const items = await this.getListItemsFiltered(
      APP_CONFIG.listAbsensiId,
      `fields/Tanggal eq '${today}'`,
      cacheKey
    );
    const found = items.find(item => {
      const itemNip = item.NIP || item.Nip || item.NRK || item.Title;
      return String(itemNip) === String(nip);
    });
    return found ? this._mapAbsensiItem(found) : null;
  }

  async getAbsensiBulanTertentu(nip, bulan, tahun) {
    const prefix = `${tahun}-${String(bulan).padStart(2,'0')}`;
    const { firstDay, lastDay } = this._getMonthRange(bulan, tahun);
    const cacheKey = `list:${APP_CONFIG.listAbsensiId}:bulan:${prefix}`;
    const items = await this.getListItemsFiltered(
      APP_CONFIG.listAbsensiId,
      `fields/Tanggal ge '${firstDay}' and fields/Tanggal le '${lastDay}'`,
      cacheKey
    );
    return items
      .filter(item => String(item.NIP || item.Nip || item.NRK || item.Title) === String(nip))
      .map(item => this._mapAbsensiItem(item));
  }

  // Semua absensi bulan tertentu untuk kalender — join emailAtasan dari list karyawan
  async getAbsensiBulanTertentu_All(bulan, tahun) {
    const prefix = `${tahun}-${String(bulan).padStart(2,'0')}`;
    const { firstDay, lastDay } = this._getMonthRange(bulan, tahun);
    const cacheKey = `list:${APP_CONFIG.listAbsensiId}:bulan:${prefix}`;

    // Fetch paralel: absensi bulan ini + map nip→atasan
    const [items, nipAtasanMap] = await Promise.all([
      this.getListItemsFiltered(
        APP_CONFIG.listAbsensiId,
        `fields/Tanggal ge '${firstDay}' and fields/Tanggal le '${lastDay}'`,
        cacheKey
      ),
      this.buildNipAtasanMap()
    ]);

    return items.map(item => this._mapAbsensiItem(item, nipAtasanMap));
  }

  async absenMasuk(data) {
    if (!this.isWithinTimeRange(APP_CONFIG.jamMasukMulai, APP_CONFIG.jamMasukSelesai)) {
      throw new Error(`Absen masuk hanya bisa dilakukan antara ${APP_CONFIG.jamMasukMulai} – ${APP_CONFIG.jamMasukSelesai}.`);
    }
    const existing = await this.getAbsensiHariIni(data.nip);
    if (existing) throw new Error('Anda sudah melakukan absen masuk hari ini.');

    const now = new Date();
    const tanggal = this.getTodayLocal();
    const jamMasuk = now.toTimeString().substring(0, 8);
    const status = this.hitungStatus(jamMasuk);

    const fields = {
      Title: data.nip, Nama: data.nama, Tanggal: tanggal,
      Jam_Masuk: jamMasuk, Jam_Keluar: '', Status: status,
      Tipe: data.tipe || 'WFO',
      Email_Atasan: data.emailAtasan || '',
      Keterangan: data.keterangan || ''
    };

    const res = await this.addListItem(APP_CONFIG.listAbsensiId, fields);
    // Invalidate cache hari ini supaya absen masuk langsung terefleksi
    cacheInvalidate(`list:${APP_CONFIG.listAbsensiId}:tanggal:${tanggal}`);
    return { id: res.listItemId, status, jamMasuk, tipe: data.tipe };
  }

  async absenKeluar(data) {
    if (!this.isWithinTimeRange(APP_CONFIG.jamKeluarMulai, APP_CONFIG.jamKeluarSelesai)) {
      throw new Error(`Absen keluar hanya bisa dilakukan mulai pukul ${APP_CONFIG.jamKeluarMulai}.`);
    }
    const existing = await this.getAbsensiHariIni(data.nip);
    if (!existing) throw new Error('Anda belum melakukan absen masuk hari ini.');
    if (existing.jamKeluar) throw new Error('Anda sudah melakukan absen keluar hari ini.');

    const now = new Date();
    const jamKeluar = now.toTimeString().substring(0, 8);
    await this.updateListItem(APP_CONFIG.listAbsensiId, existing.id, {
      Jam_Keluar: jamKeluar,
      Keterangan: data.keterangan || existing.keterangan || ''
    });
    // Invalidate cache hari ini
    cacheInvalidate(`list:${APP_CONFIG.listAbsensiId}:tanggal:${this.getTodayLocal()}`);
    return { jamKeluar };
  }

  hitungStatus(jamMasuk) {
    const [h, m] = jamMasuk.split(':').map(Number);
    const menitMasuk = h * 60 + m;
    const [bh, bm] = APP_CONFIG.jamMasukSelesai.split(':').map(Number);
    const batasMenit = bh * 60 + bm;
    if (menitMasuk <= batasMenit) return 'Tepat Waktu';
    if (menitMasuk <= batasMenit + APP_CONFIG.toleransiTerlambat) return 'Terlambat Ringan';
    return 'Terlambat';
  }

  // ============================================================
  // PERMOHONAN WFA / VISIT
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
      catatanAtasan: item.Catatan_Atasan || item.CatatanAtasan || '',
      notifiedUser: item.Notified_User === 'true'
    };
  }

  async getPermohonanWfa(email) {
    // Filter by email di JS — email tidak bisa di-$filter dengan mudah di SharePoint
    const items = await this.getListItems(
      APP_CONFIG.listPermohonanWfaId,
      `list:${APP_CONFIG.listPermohonanWfaId}:all`
    );
    return items
      .filter(item => {
        const itemEmail = item.Email_User || item.EmailUser || item.Email || '';
        return String(itemEmail).toLowerCase() === email.toLowerCase();
      })
      .map(item => this._mapPermohonanItem(item));
  }

  async getPermohonanWfaById(id) {
    // Fetch langsung by ID — efisien, tidak perlu load semua
    const url = `${this.getSharePointBase()}/lists/${APP_CONFIG.listPermohonanWfaId}/items/${id}?expand=fields`;
    try {
      const data = await this.apiFetch(url);
      if (!data) return null;
      return this._mapPermohonanItem({ ...data.fields, listItemId: data.id });
    } catch {
      // Fallback ke full list
      const items = await this.getListItems(APP_CONFIG.listPermohonanWfaId);
      const found = items.find(item => String(item.listItemId) === String(id));
      return found ? this._mapPermohonanItem(found) : null;
    }
  }

  async getApprovedRequestByBulan(bulan, tahun) {
    const prefix = `${tahun}-${String(bulan).padStart(2,'0')}`;
    const cacheKey = `approved-request:${prefix}`;
    const cached = cacheGet(cacheKey);
    if (cached) return cached;

    const items = await this.getListItems(
      APP_CONFIG.listPermohonanWfaId,
      `list:${APP_CONFIG.listPermohonanWfaId}:all`
    );

    const result = [];
    items
      .filter(item => item.Status === 'Approved')
      .forEach(req => {
        const datesField = req.Tanggal_WFA || req.TanggalWfa || req.Tanggal || '';
        datesField.split(',').map(d => d.trim()).forEach(d => {
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

    return cacheSet(cacheKey, result);
  }

  async getApprovedWfaByBulan(bulan, tahun) {
    return this.getApprovedRequestByBulan(bulan, tahun);
  }

  async tambahPermohonanWfa(data) {
    const fields = {
      Title: data.nip, Nama: data.nama,
      Email_User: data.emailUser,
      Tanggal_WFA: data.tanggalWfa,
      Tipe: data.tipe || 'WFA',
      Status: 'Approved',
      Email_Atasan: data.emailAtasan,
      Catatan_User: data.catatanUser || '',
      Catatan_Atasan: ''
    };

    const res = await this.addListItem(APP_CONFIG.listPermohonanWfaId, fields);
    cacheInvalidate(`list:${APP_CONFIG.listPermohonanWfaId}`);
    cacheInvalidate('approved-request:');

    try {
      await this.sendApprovalEmail({ listItemId: res.listItemId, ...fields });
    } catch (e) {
      console.error('Gagal mengirim email notifikasi ke atasan:', e);
    }

    return res.listItemId;
  }

  async updateStatusPermohonanWfa(id, status, catatanAtasan = '') {
    await this.updateListItem(APP_CONFIG.listPermohonanWfaId, id, {
      Status: status,
      Catatan_Atasan: catatanAtasan
    });
    cacheInvalidate(`list:${APP_CONFIG.listPermohonanWfaId}`);
    cacheInvalidate('approved-request:');

    const req = await this.getPermohonanWfaById(id);
    if (req) {
      try { await this.sendResponseEmail(req, status, catatanAtasan); }
      catch (e) { console.error('Gagal mengirim email balasan:', e); }
    }
  }

  // ============================================================
  // EMAIL
  // ============================================================

  async sendMail(message) {
    return this.apiFetch(`${GRAPH_BASE}/me/sendMail`, {
      method: 'POST',
      body: JSON.stringify({ message, saveToSentItems: 'true' })
    });
  }

  async sendApprovalEmail(req) {
    const tipe = req.Tipe || 'WFA';
    const formattedDates = req.Tanggal_WFA.split(',').map(d => `• ${d.trim()}`).join('<br>');
    const rejectUrl = `${APP_CONFIG.redirectUri}?action=reject&id=${req.listItemId}`;
    const tipeColor = tipe === 'WFA' ? '#4299e1' : '#ed8936';
    const tipeLabel = tipe === 'Visit' ? 'Visit / Kunjungan Lapangan' : 'Work From Anywhere (WFA)';

    const mailContent = `
      <div style="font-family:'Segoe UI',sans-serif;max-width:600px;margin:auto;padding:20px;border:1px solid #e2e8f0;border-radius:8px;background:#fff;">
        <div style="background:${tipeColor};color:white;padding:12px 20px;border-radius:6px 6px 0 0;margin:-20px -20px 20px -20px;">
          <h2 style="margin:0;font-size:1.1rem;">Pemberitahuan ${tipeLabel}</h2>
        </div>
        <p>Halo,</p>
        <p>Karyawan berikut telah mengajukan permohonan <strong>${tipe}</strong> dan statusnya telah <strong>disetujui secara otomatis oleh sistem</strong>:</p>
        <table style="width:100%;margin:20px 0;border-collapse:collapse;">
          <tr style="background:#f7fafc;"><td style="padding:10px;font-weight:bold;width:160px;">Nama</td><td style="padding:10px;">${req.Nama}</td></tr>
          <tr><td style="padding:10px;font-weight:bold;">NIP / NRK</td><td style="padding:10px;">${req.NRK||req.NIP||req.Title||''}</td></tr>
          <tr style="background:#f7fafc;"><td style="padding:10px;font-weight:bold;">Tipe</td><td style="padding:10px;"><span style="background:${tipeColor};color:white;padding:2px 10px;border-radius:20px;font-size:0.85rem;font-weight:bold;">${tipe}</span></td></tr>
          <tr><td style="padding:10px;font-weight:bold;vertical-align:top;">Tanggal</td><td style="padding:10px;">${formattedDates}</td></tr>
          <tr style="background:#f7fafc;"><td style="padding:10px;font-weight:bold;vertical-align:top;">Alasan</td><td style="padding:10px;color:#718096;font-style:italic;">"${req.Catatan_User||'-'}"</td></tr>
        </table>
        <p style="margin-top:30px;font-weight:500;">Jika Anda menyetujui, <strong>tidak perlu melakukan tindakan apapun</strong>. Jika ingin <strong>membatalkan/menolak</strong>, klik tombol di bawah:</p>
        <div style="margin:20px 0;">
          <a href="${rejectUrl}" style="display:inline-block;background:#f56565;color:white;padding:12px 24px;text-decoration:none;border-radius:6px;font-weight:bold;">Batalkan / Tolak Pengajuan</a>
        </div>
        <p style="font-size:0.85rem;color:#a0aec0;margin-top:40px;border-top:1px solid #e2e8f0;padding-top:15px;">Tombol di atas akan membuka web aplikasi Absensi. Anda perlu masuk dengan akun Microsoft 365.</p>
      </div>`;

    const userEmail = req.Email_User || req.EmailUser || req.Email || '';
    const toAddress = (req.Email_Atasan || '').toLowerCase().trim();
    const ccRecipients = this._buildCcRecipients([APP_CONFIG.emailHrd, APP_CONFIG.emailMis, userEmail], toAddress);

    const message = {
      subject: `[${tipe} Notification] Pengajuan ${tipe} - ${req.Nama} (Disetujui Otomatis)`,
      body: { contentType: 'HTML', content: mailContent },
      toRecipients: [{ emailAddress: { address: req.Email_Atasan } }]
    };
    if (ccRecipients.length > 0) message.ccRecipients = ccRecipients;
    return this.sendMail(message);
  }

  async sendResponseEmail(req, status, catatanAtasan) {
    const isApproved = status === 'Approved';
    const statusLabel = isApproved ? 'DISETUJUI' : 'DIBATALKAN / DITOLAK';
    const color = isApproved ? '#48bb78' : '#f56565';
    const tipe = req.tipe || 'WFA';
    const formattedDates = req.tanggalWfa.split(',').map(d => `• ${d.trim()}`).join('<br>');

    const mailContent = `
      <div style="font-family:'Segoe UI',sans-serif;max-width:600px;margin:auto;padding:20px;border:1px solid #e2e8f0;border-radius:8px;background:#fff;">
        <h2 style="color:${color};border-bottom:2px solid #e2e8f0;padding-bottom:10px;">Status Pengajuan ${tipe} - ${statusLabel}</h2>
        <p>Halo ${req.nama},</p>
        <p>Pengajuan <strong>${tipe}</strong> Anda telah ditinjau oleh atasan:</p>
        <table style="width:100%;margin:20px 0;border-collapse:collapse;">
          <tr style="background:#f7fafc;"><td style="padding:10px;font-weight:bold;width:150px;">Status</td><td style="padding:10px;color:${color};font-weight:bold;">${statusLabel}</td></tr>
          <tr><td style="padding:10px;font-weight:bold;">Tipe</td><td style="padding:10px;">${tipe}</td></tr>
          <tr style="background:#f7fafc;"><td style="padding:10px;font-weight:bold;vertical-align:top;">Tanggal</td><td style="padding:10px;">${formattedDates}</td></tr>
          <tr><td style="padding:10px;font-weight:bold;vertical-align:top;">Catatan Atasan</td><td style="padding:10px;">${catatanAtasan||'-'}</td></tr>
        </table>
        <p style="margin-top:30px;">${isApproved ? 'Silakan melakukan absensi masuk/pulang di webapp pada tanggal yang disetujui.' : `Pengajuan ${tipe} Anda telah dibatalkan/ditolak. Silakan hubungi atasan Anda jika ada pertanyaan.`}</p>
        <p style="font-size:0.85rem;color:#a0aec0;margin-top:40px;border-top:1px solid #e2e8f0;padding-top:15px;">Email ini dikirim otomatis oleh Sistem Absensi ${APP_CONFIG.namaPerusahaan}.</p>
      </div>`;

    const toAddress = (req.emailUser || '').toLowerCase().trim();
    const ccRecipients = this._buildCcRecipients([APP_CONFIG.emailHrd, APP_CONFIG.emailMis], toAddress);

    const message = {
      subject: `[${tipe} Status] Pengajuan ${tipe} Anda telah ${statusLabel}`,
      body: { contentType: 'HTML', content: mailContent },
      toRecipients: [{ emailAddress: { address: req.emailUser } }]
    };
    if (ccRecipients.length > 0) message.ccRecipients = ccRecipients;
    return this.sendMail(message);
  }

  _buildCcRecipients(rawList, toAddress) {
    const seen = new Set();
    const result = [];
    for (const raw of rawList) {
      if (!raw) continue;
      const clean = raw.toLowerCase().trim();
      if (clean && clean !== toAddress && !seen.has(clean)) {
        seen.add(clean);
        result.push({ emailAddress: { address: clean } });
      }
    }
    return result;
  }
}

  // Tandai request rejected sudah dinotif (persist ke SharePoint kolom Notified_User)
  async markRejectedAsNotified(requestId) {
    await this.updateListItem(APP_CONFIG.listPermohonanWfaId, requestId, {
      Notified_User: 'true'
    });
    cacheInvalidate(`list:${APP_CONFIG.listPermohonanWfaId}`);
  }

  // Generate NIK berikutnya dari NIK tertinggi di list karyawan + 1
  async getNextNik() {
    const karyawan = await this.getAllKaryawan();
    if (karyawan.length === 0) return '';
    const niks = karyawan.map(k => parseInt(k.nip, 10)).filter(n => !isNaN(n));
    if (niks.length === 0) return '';
    return String(Math.max(...niks) + 1);
  }


const graphService = new GraphService();
export default graphService;
