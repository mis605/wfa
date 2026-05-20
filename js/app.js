// ============================================================
// APP.JS - Logika utama & UI controller (SharePoint Lists & 3-Tab)
// ============================================================

import { APP_CONFIG } from './config.js';
import authService from './auth.js';
import graphService from './graph.js';
import {
  formatTanggal, formatJam, formatTanggalPendek,
  hitungDurasi, showToast, setLoading, getMonthYear, getTodayString
} from './utils.js';

// ============================================================
// STATE GLOBAL
// ============================================================
const state = {
  user: null,          // MS user profile
  karyawan: null,      // Data karyawan dari SharePoint List
  absensiHariIni: null,
  calendarDate: new Date(), // Tanggal aktif Kalender (Tab 1)
  pickerDate: new Date(),   // Tanggal aktif Picker Request (Tab 3)
  selectedDates: [],        // Daftar tanggal terpilih ("YYYY-MM-DD")
  approvedWfaList: [],      // Jadwal WFA yang disetujui bulan ini
  currentApprovalRequest: null // Permohonan yang sedang ditinjau atasan
};

// ============================================================
// ROUTER VIEWS
// ============================================================
const views = {
  login: document.getElementById('view-login'),
  calendar: document.getElementById('view-calendar'),
  absen: document.getElementById('view-absen'),
  request: document.getElementById('view-request'),
  approval: document.getElementById('view-approval'),
  loading: document.getElementById('view-loading'),
};

function showView(viewName) {
  const appShell = document.getElementById('app-shell');
  const isAuthView = viewName === 'login' || viewName === 'loading';
  
  // Tampilkan/sembunyikan app-shell
  if (appShell) appShell.classList.toggle('hidden', isAuthView);
  
  // Tampilkan view yang dipilih
  Object.entries(views).forEach(([name, el]) => {
    if (!el) return;
    if (isAuthView) {
      el.classList.toggle('hidden', name !== viewName);
    } else {
      if (name === 'login' || name === 'loading') {
        el.classList.add('hidden');
      } else {
        el.classList.toggle('hidden', name !== viewName);
      }
    }
  });
  
  state.currentView = viewName;
  
  // Update status aktif navigasi bottom-nav
  document.querySelectorAll('[data-nav]').forEach(el => {
    el.classList.toggle('nav__item--active', el.dataset.nav === viewName);
  });
}

// ============================================================
// INISIALISASI APP
// ============================================================
async function initApp() {
  showView('loading');
  
  try {
    await authService.init();
    
    if (!authService.isLoggedIn()) {
      showView('login');
      return;
    }

    await loadUserSession();
    
  } catch (err) {
    console.error('Init error:', err);
    showView('login');
  }
}

async function loadUserSession() {
  showView('loading');
  document.getElementById('loading-text').textContent = 'Memuat profil Microsoft 365...';
  
  try {
    // Ambil profil MS pengguna
    console.log('[App] Loading user session. Fetching user profile...');
    state.user = await authService.getUserProfile();
    console.log('[App] User profile retrieved:', state.user);
    
    // Debug list columns to see exact internal names of columns in SharePoint
    try {
      await graphService.debugListColumns(APP_CONFIG.listKaryawanId, 'Karyawan');
      await graphService.debugListColumns(APP_CONFIG.listAbsensiId, 'Absensi');
      await graphService.debugListColumns(APP_CONFIG.listPermohonanWfaId, 'PermohonanWfa');
    } catch (e) {
      console.warn('[App] Failed to run debugListColumns:', e);
    }
    
    document.getElementById('loading-text').textContent = 'Memverifikasi data karyawan...';
    // Cek data karyawan di list SharePoint
    const userEmail = state.user.mail || state.user.userPrincipalName;
    console.log('[App] Verifying employee status for email:', userEmail);
    
    try {
      state.karyawan = await graphService.getKaryawanByEmail(userEmail);
      console.log('[App] Employee verification result:', state.karyawan);
    } catch (karyawanError) {
      console.error('[App] Error in getKaryawanByEmail:', karyawanError);
      throw karyawanError;
    }
    
    if (!state.karyawan) {
      console.warn('[App] Employee not found in SharePoint list');
      showView('login');
      document.getElementById('login-error').textContent = 
        'Akun Anda belum terdaftar sebagai karyawan di SharePoint. Silakan hubungi admin.';
      document.getElementById('login-error').classList.remove('hidden');
      return;
    }

    if (state.karyawan.statusAktif !== 'Aktif') {
      console.warn('[App] Employee account is not Active:', state.karyawan.statusAktif);
      showView('login');
      document.getElementById('login-error').textContent = 'Akun Anda berstatus Tidak Aktif. Silakan hubungi admin.';
      document.getElementById('login-error').classList.remove('hidden');
      return;
    }

    // Load avatar/foto profil
    const photoUrl = await authService.getUserPhoto();
    const initials = getInitials(state.karyawan.nama);
    
    document.querySelectorAll('.user-avatar, .greeting-avatar').forEach(el => {
      if (photoUrl) {
        el.style.backgroundImage = `url(${photoUrl})`;
        el.textContent = '';
      } else {
        el.style.backgroundImage = 'none';
        el.textContent = initials;
      }
    });

    // Cek URL parameters untuk aksi approval persetujuan atasan
    const params = new URLSearchParams(window.location.search);
    const action = params.get('action');
    const requestId = params.get('id');
    
    if (action && requestId) {
      await loadApprovalRequest(requestId, action);
    } else {
      // Landing page default: Kalender WFA (Tab 1)
      showView('calendar');
      await loadCalendar();
    }
    
  } catch (err) {
    console.error('Session error:', err);
    showToast('Gagal memuat data sesi: ' + err.message, 'error');
    showView('login');
  }
}

function getInitials(nama) {
  if (!nama) return 'U';
  return nama.split(' ').slice(0, 2).map(n => n[0]).join('').toUpperCase();
}

// Live clock helper
function updateClock() {
  const clockEl = document.getElementById('live-clock');
  if (!clockEl) return;
  
  const update = () => {
    const now = new Date();
    clockEl.textContent = now.toLocaleTimeString('id-ID', {
      hour: '2-digit', minute: '2-digit', second: '2-digit'
    });
  };
  update();
  
  if (!window._clockInterval) {
    window._clockInterval = setInterval(update, 1000);
  }
}

function getStatusClass(status) {
  const map = {
    'Tepat Waktu': 'success',
    'Terlambat Ringan': 'warning',
    'Terlambat': 'danger',
    'Selesai': 'success',
    'Belum Absen': 'pending',
  };
  return map[status] || 'pending';
}

// ============================================================
// TAB 1: KALENDER WFA
// ============================================================
async function loadCalendar(offset = 0) {
  if (offset !== 0) {
    state.calendarDate = new Date(
      state.calendarDate.getFullYear(),
      state.calendarDate.getMonth() + offset,
      1
    );
  }
  
  const month = state.calendarDate.getMonth() + 1;
  const year = state.calendarDate.getFullYear();
  const monthLabel = state.calendarDate.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
  
  document.getElementById('calendar-month-label').textContent = monthLabel;
  
  const container = document.getElementById('calendar-days-container');
  container.innerHTML = '<div style="grid-column: span 7; text-align: center; padding: 24px;"><span class="spinner"></span></div>';
  
  try {
    state.approvedWfaList = await graphService.getApprovedWfaByBulan(month, year);
    renderCalendarGrid(month, year);
  } catch (err) {
    container.innerHTML = `<div style="grid-column: span 7; text-align: center; padding: 24px; color: var(--red);">Gagal memuat jadwal kalender: ${err.message}</div>`;
  }
}

function renderCalendarGrid(month, year) {
  const container = document.getElementById('calendar-days-container');
  const daysInMonth = new Date(year, month, 0).getDate();
  
  const firstDayIndex = new Date(year, month - 1, 1).getDay();
  // Senin ke Minggu index mapping
  const startOffset = firstDayIndex === 0 ? 6 : firstDayIndex - 1;
  
  let cellsHtml = '';
  
  // Sel kosong sebelum tanggal 1
  for (let i = 0; i < startOffset; i++) {
    cellsHtml += '<div class="calendar-day-cell calendar-day-cell--empty"></div>';
  }
  
  const today = new Date();
  const todayStr = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
  
  for (let day = 1; day <= daysInMonth; day++) {
    const currentDayStr = `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    
    // Cari daftar WFA yang disetujui untuk hari ini
    const wfaOnDay = state.approvedWfaList.filter(item => item.tanggal === currentDayStr);
    
    const isToday = currentDayStr === todayStr;
    const todayClass = isToday ? 'calendar-day-cell--today' : '';
    
    const wfaHtml = wfaOnDay.map(w => `
      <span class="calendar-wfa-name" title="${w.nama} (${w.nip})">${w.nama}</span>
    `).join('');
    
    cellsHtml += `
      <div class="calendar-day-cell ${todayClass}">
        <span class="calendar-day-number">${day}</span>
        <div class="calendar-wfa-list">
          ${wfaHtml}
        </div>
      </div>
    `;
  }
  
  container.innerHTML = cellsHtml;
}

// ============================================================
// TAB 2: ABSEN WFA
// ============================================================
let absenMode = 'masuk'; // 'masuk' | 'keluar'

async function loadAbsen() {
  const statusCard = document.getElementById('wfa-status-card');
  const statusText = document.getElementById('wfa-status-text');
  const dashboardContent = document.getElementById('absen-dashboard-content');
  
  statusCard.className = 'wfa-status-card';
  statusText.textContent = 'Memeriksa jadwal WFA...';
  dashboardContent.classList.add('hidden');
  
  // Update header info kartu greeting
  document.getElementById('dash-nama').textContent = state.karyawan.nama;
  document.getElementById('dash-jabatan').textContent = `${state.karyawan.jabatan} • ${state.karyawan.departemen}`;
  document.getElementById('dash-tanggal').textContent = formatTanggal(new Date().toISOString());
  
  updateClock();
  
  try {
    const todayStr = getTodayString();
    
    // Verifikasi apakah atasan menyetujui WFA hari ini
    const requests = await graphService.getPermohonanWfa(state.user.mail || state.user.userPrincipalName);
    const approvedForToday = requests.some(req => {
      if (req.status !== 'Approved') return false;
      const dates = req.tanggalWfa ? req.tanggalWfa.split(',').map(d => d.trim()) : [];
      return dates.includes(todayStr);
    });
    
    if (!approvedForToday) {
      statusCard.className = 'wfa-status-card wfa-status-card--none';
      statusText.textContent = 'Anda tidak memiliki jadwal WFA yang disetujui hari ini.';
      return;
    }
    
    statusCard.className = 'wfa-status-card wfa-status-card--approved';
    statusText.textContent = 'Jadwal WFA Anda disetujui untuk hari ini.';
    dashboardContent.classList.remove('hidden');
    
    // Load status jam absen masuk/keluar hari ini
    state.absensiHariIni = await graphService.getAbsensiHariIni(state.karyawan.nip);
    renderAbsenDashboard();
    
  } catch (err) {
    statusCard.className = 'wfa-status-card wfa-status-card--rejected';
    statusText.textContent = 'Gagal memeriksa jadwal WFA: ' + err.message;
  }
}

function renderAbsenDashboard() {
  const a = state.absensiHariIni;
  
  const statusEl = document.getElementById('dash-status');
  const masukEl = document.getElementById('dash-masuk');
  const keluarEl = document.getElementById('dash-keluar');
  const durasiEl = document.getElementById('dash-durasi');
  
  const btnMasuk = document.getElementById('btn-absen-masuk');
  const btnKeluar = document.getElementById('btn-absen-keluar');
  
  if (!a) {
    statusEl.textContent = 'Belum Absen';
    statusEl.className = 'status-badge status--pending';
    masukEl.textContent = '--:--';
    keluarEl.textContent = '--:--';
    durasiEl.textContent = '--';
    
    btnMasuk.disabled = false;
    btnMasuk.classList.remove('btn--disabled');
    btnKeluar.disabled = true;
    btnKeluar.classList.add('btn--disabled');
  } else {
    const isClockedOut = !!a.jamKeluar;
    statusEl.textContent = isClockedOut ? 'Selesai' : a.status;
    statusEl.className = `status-badge status--${getStatusClass(a.status)}`;
    masukEl.textContent = formatJam(a.jamMasuk);
    keluarEl.textContent = isClockedOut ? formatJam(a.jamKeluar) : '--:--';
    durasiEl.textContent = hitungDurasi(a.jamMasuk, a.jamKeluar);
    
    if (!isClockedOut) {
      btnMasuk.disabled = true;
      btnMasuk.classList.add('btn--disabled');
      btnKeluar.disabled = false;
      btnKeluar.classList.remove('btn--disabled');
    } else {
      btnMasuk.disabled = true;
      btnMasuk.classList.add('btn--disabled');
      btnKeluar.disabled = true;
      btnKeluar.classList.add('btn--disabled');
    }
  }
  
  // Set configuration times display
  const infoMasuk = document.getElementById('info-jam-masuk');
  const infoKeluar = document.getElementById('info-jam-keluar');
  if (infoMasuk) infoMasuk.textContent = `${APP_CONFIG.jamMasukMulai}–${APP_CONFIG.jamMasukSelesai}`;
  if (infoKeluar) infoKeluar.textContent = `${APP_CONFIG.jamKeluarMulai}–${APP_CONFIG.jamKeluarSelesai}`;
}

function openAbsenModal(mode) {
  absenMode = mode;
  const modal = document.getElementById('modal-absen');
  const title = document.getElementById('modal-absen-title');
  const textarea = document.getElementById('absen-keterangan');
  
  title.textContent = mode === 'masuk' ? 'Absen Masuk' : 'Absen Keluar';
  textarea.value = '';
  
  modal.classList.remove('hidden');
  modal.classList.add('modal--show');
}

function closeAbsenModal() {
  const modal = document.getElementById('modal-absen');
  modal.classList.remove('modal--show');
  setTimeout(() => modal.classList.add('hidden'), 300);
}

async function submitAbsen() {
  const btn = document.getElementById('btn-submit-absen');
  const keterangan = document.getElementById('absen-keterangan').value.trim();
  
  setLoading(btn, true);
  
  try {
    const payload = {
      nip: state.karyawan.nip,
      nama: state.karyawan.nama,
      keterangan: keterangan
    };
    
    if (absenMode === 'masuk') {
      const result = await graphService.absenMasuk(payload);
      showToast(`✓ Absen masuk berhasil! Status: ${result.status}`, 'success');
    } else {
      const result = await graphService.absenKeluar(payload);
      showToast(`✓ Absen keluar berhasil! Jam: ${result.jamKeluar.substring(0, 5)}`, 'success');
    }
    
    await loadAbsen();
    closeAbsenModal();
  } catch (err) {
    showToast(err.message, 'error');
  } finally {
    setLoading(btn, false, 'Absen Sekarang');
  }
}

// ============================================================
// TAB 3: REQUEST WFA (PENGAJUAN TANGGAL)
// ============================================================
function loadRequestView(offset = 0) {
  if (offset !== 0) {
    state.pickerDate = new Date(
      state.pickerDate.getFullYear(),
      state.pickerDate.getMonth() + offset,
      1
    );
  }
  
  const month = state.pickerDate.getMonth() + 1;
  const year = state.pickerDate.getFullYear();
  const monthLabel = state.pickerDate.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
  
  document.getElementById('picker-month-label').textContent = monthLabel;
  
  // Tampilkan email supervisor
  const managerEmail = state.karyawan.emailAtasan || '';
  const managerEl = document.getElementById('req-manager-email');
  if (managerEl) {
    if (managerEmail) {
      managerEl.textContent = managerEmail;
      managerEl.style.color = '';
    } else {
      managerEl.textContent = 'Belum dikonfigurasi (Hubungi Admin)';
      managerEl.style.color = 'var(--red)';
    }
  }
  
  renderPickerGrid(month, year);
  renderSelectedDates();
}

function renderPickerGrid(month, year) {
  const container = document.getElementById('picker-days-container');
  const daysInMonth = new Date(year, month, 0).getDate();
  
  const firstDayIndex = new Date(year, month - 1, 1).getDay();
  // Senin ke Minggu index mapping
  const startOffset = firstDayIndex === 0 ? 6 : firstDayIndex - 1;
  
  let cellsHtml = '';
  
  // Sel kosong
  for (let i = 0; i < startOffset; i++) {
    cellsHtml += '<div class="picker-day-cell picker-day-cell--empty"></div>';
  }
  
  const today = new Date();
  const todayStr = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
  
  for (let day = 1; day <= daysInMonth; day++) {
    const currentDayStr = `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    const isSelected = state.selectedDates.includes(currentDayStr);
    const isDisabled = currentDayStr < todayStr; // Disable tanggal lampau
    
    let cellClass = 'picker-day-cell';
    if (isSelected) cellClass += ' picker-day-cell--selected';
    if (isDisabled) cellClass += ' picker-day-cell--disabled';
    
    cellsHtml += `
      <div class="${cellClass}" data-date="${currentDayStr}">
        ${day}
      </div>
    `;
  }
  
  container.innerHTML = cellsHtml;
  
  // Click listener pada tanggal pemilih
  container.querySelectorAll('.picker-day-cell:not(.picker-day-cell--empty):not(.picker-day-cell--disabled)').forEach(el => {
    el.addEventListener('click', () => {
      const dateStr = el.dataset.date;
      const idx = state.selectedDates.indexOf(dateStr);
      if (idx > -1) {
        state.selectedDates.splice(idx, 1);
      } else {
        state.selectedDates.push(dateStr);
        state.selectedDates.sort(); // Urutkan secara kronologis
      }
      renderPickerGrid(month, year);
      renderSelectedDates();
    });
  });
}

function renderSelectedDates() {
  const container = document.getElementById('selected-dates-list');
  const btnSubmit = document.getElementById('btn-submit-request');
  
  if (state.selectedDates.length === 0) {
    container.innerHTML = '<span class="empty-text">Belum ada tanggal dipilih. Klik pada kalender di atas.</span>';
    btnSubmit.disabled = true;
    return;
  }
  
  container.innerHTML = state.selectedDates.map(d => `
    <span class="selected-date-tag" data-date="${d}">
      ${formatTanggalPendek(d)}
      <span class="selected-date-tag__remove">&times;</span>
    </span>
  `).join('');
  
  // Event remove tag
  container.querySelectorAll('.selected-date-tag__remove').forEach(el => {
    el.addEventListener('click', (e) => {
      e.stopPropagation();
      const dateStr = el.parentElement.dataset.date;
      const idx = state.selectedDates.indexOf(dateStr);
      if (idx > -1) {
        state.selectedDates.splice(idx, 1);
        const month = state.pickerDate.getMonth() + 1;
        const year = state.pickerDate.getFullYear();
        renderPickerGrid(month, year);
        renderSelectedDates();
      }
    });
  });
  
  validateRequestForm();
}

function validateRequestForm() {
  const btnSubmit = document.getElementById('btn-submit-request');
  const catatan = document.getElementById('req-catatan').value.trim();
  const managerEmail = state.karyawan.emailAtasan || '';
  
  const isValid = state.selectedDates.length > 0 && catatan.length > 0 && managerEmail.length > 0;
  btnSubmit.disabled = !isValid;
}

async function submitWfaRequest() {
  const catatan = document.getElementById('req-catatan').value.trim();
  const managerEmail = state.karyawan.emailAtasan || '';
  
  if (state.selectedDates.length === 0) {
    showToast('Pilih minimal satu tanggal pengajuan.', 'warning');
    return;
  }
  if (!catatan) {
    showToast('Alasan pengajuan WFA wajib diisi.', 'warning');
    return;
  }
  if (!managerEmail) {
    showToast('Atasan Anda belum dikonfigurasi. Hubungi admin.', 'warning');
    return;
  }
  
  const btn = document.getElementById('btn-submit-request');
  setLoading(btn, true);
  
  try {
    const payload = {
      nip: state.karyawan.nip,
      nama: state.karyawan.nama,
      emailUser: state.user.mail || state.user.userPrincipalName,
      tanggalWfa: state.selectedDates.join(', '), // Bentuk string terpisah koma
      emailAtasan: managerEmail,
      catatanUser: catatan
    };
    
    await graphService.tambahPermohonanWfa(payload);
    showToast('✓ Pengajuan WFA berhasil dikirim ke atasan!', 'success');
    
    // Reset Form
    state.selectedDates = [];
    document.getElementById('req-catatan').value = '';
    
    const month = state.pickerDate.getMonth() + 1;
    const year = state.pickerDate.getFullYear();
    renderPickerGrid(month, year);
    renderSelectedDates();
    
    // Redirect ke kalender utama (Tab 1)
    showView('calendar');
    await loadCalendar();
    
  } catch (err) {
    showToast('Gagal memproses pengajuan: ' + err.message, 'error');
  } finally {
    setLoading(btn, false, 'Kirim Pengajuan WFA');
  }
}

// ============================================================
// VIEW PERSATUAN / APPROVAL ATASAN (URL PARSING FLOW)
// ============================================================
async function loadApprovalRequest(requestId, action) {
  showView('approval');
  const loadingState = document.getElementById('approval-loading-state');
  const boxContent = document.getElementById('approval-box-content');
  const titleText = document.getElementById('approval-title-text');
  
  loadingState.classList.remove('hidden');
  boxContent.classList.add('hidden');
  
  try {
    const req = await graphService.getPermohonanWfaById(requestId);
    if (!req) {
      throw new Error('Pengajuan WFA tidak ditemukan.');
    }
    
    state.currentApprovalRequest = req;
    
    // Tulis rincian pengajuan ke layar
    document.getElementById('app-nama').textContent = req.nama;
    document.getElementById('app-nip').textContent = req.nip;
    document.getElementById('app-catatan-karyawan').textContent = req.catatanUser || '-';
    
    // Susun daftar tanggal terpilih
    const datesEl = document.getElementById('app-tanggal-list');
    const dates = req.tanggalWfa ? req.tanggalWfa.split(',') : [];
    datesEl.innerHTML = dates.map(d => `<span class="selected-date-tag">${formatTanggalPendek(d.trim())}</span>`).join('');
    
    const labelAtasan = document.getElementById('app-label-atasan');
    const btnConfirm = document.getElementById('btn-confirm-approval');
    const textarea = document.getElementById('app-catatan-atasan');
    
    if (req.status !== 'Pending') {
      // Jika statusnya sudah diubah (sudah disetujui / ditolak sebelumnya)
      titleText.textContent = `Pengajuan WFA: ${req.status.toUpperCase()}`;
      labelAtasan.textContent = 'Catatan Atasan (Sudah Diproses)';
      textarea.value = req.catatanAtasan || '';
      textarea.disabled = true;
      btnConfirm.disabled = true;
      btnConfirm.textContent = `Sudah ${req.status}`;
      btnConfirm.classList.add('btn--disabled');
      btnConfirm.style.backgroundColor = '';
    } else {
      // Masih pending, persiapkan input & tombol respon
      const isApprove = action === 'approve';
      titleText.textContent = isApprove ? 'Setujui Pengajuan WFA' : 'Tolak Pengajuan WFA';
      labelAtasan.textContent = isApprove ? 'Catatan / Alasan Persetujuan' : 'Catatan / Alasan Penolakan';
      textarea.value = '';
      textarea.disabled = false;
      btnConfirm.disabled = false;
      btnConfirm.textContent = isApprove ? 'Konfirmasi Setujui (Approve)' : 'Konfirmasi Tolak (Reject)';
      btnConfirm.classList.remove('btn--disabled');
      btnConfirm.style.backgroundColor = isApprove ? 'var(--green)' : 'var(--red)';
    }
    
    loadingState.classList.add('hidden');
    boxContent.classList.remove('hidden');
    
  } catch (err) {
    loadingState.classList.add('hidden');
    showToast('Gagal memuat detail pengajuan WFA: ' + err.message, 'error');
    // Jika gagal, redirect ke Kalender utama
    showView('calendar');
    await loadCalendar();
  }
}

async function confirmApproval() {
  const req = state.currentApprovalRequest;
  if (!req) return;
  
  const params = new URLSearchParams(window.location.search);
  const action = params.get('action');
  const status = action === 'approve' ? 'Approved' : 'Rejected';
  const catatanAtasan = document.getElementById('app-catatan-atasan').value.trim();
  
  const btn = document.getElementById('btn-confirm-approval');
  setLoading(btn, true);
  
  try {
    await graphService.updateStatusPermohonanWfa(req.id, status, catatanAtasan);
    showToast(`✓ Pengajuan WFA berhasil ${status === 'Approved' ? 'disetujui' : 'ditolak'}!`, 'success');
    
    // Hapus parameter URL agar tidak terpanggil ulang saat reload
    window.history.replaceState({}, document.title, window.location.pathname);
    
    // Redirect ke Kalender
    showView('calendar');
    await loadCalendar();
  } catch (err) {
    showToast('Gagal memproses persetujuan: ' + err.message, 'error');
  } finally {
    setLoading(btn, false, 'Konfirmasi');
  }
}

function cancelApproval() {
  // Clear URL query parameters
  window.history.replaceState({}, document.title, window.location.pathname);
  
  // Buka kalender
  showView('calendar');
  loadCalendar();
}

// ============================================================
// BIND EVENTS & EVENT LISTENERS
// ============================================================
function bindEvents() {
  // Microsoft Login Button
  document.getElementById('btn-login')?.addEventListener('click', async () => {
    const btn = document.getElementById('btn-login');
    setLoading(btn, true);
    try {
      await authService.login();
      await loadUserSession();
    } catch (err) {
      showToast('Login gagal: ' + err.message, 'error');
      setLoading(btn, false, 'Masuk dengan Microsoft 365');
    }
  });

  // Logout Button
  document.getElementById('btn-logout')?.addEventListener('click', async () => {
    await authService.logout();
    window.location.reload();
  });

  // Navigasi Bottom Nav (3 Tab utama)
  document.querySelectorAll('[data-nav]').forEach(el => {
    el.addEventListener('click', async () => {
      const view = el.dataset.nav;
      showView(view);
      
      if (view === 'calendar') await loadCalendar();
      if (view === 'absen') await loadAbsen();
      if (view === 'request') loadRequestView();
    });
  });

  // Calendar Bulan Navigasi (Tab 1)
  document.getElementById('btn-cal-prev')?.addEventListener('click', () => loadCalendar(-1));
  document.getElementById('btn-cal-next')?.addEventListener('click', () => loadCalendar(1));

  // Request Picker Bulan Navigasi (Tab 3)
  document.getElementById('btn-picker-prev')?.addEventListener('click', () => loadRequestView(-1));
  document.getElementById('btn-picker-next')?.addEventListener('click', () => loadRequestView(1));

  // Input catatan alasan request
  document.getElementById('req-catatan')?.addEventListener('input', validateRequestForm);

  // Submit pengajuan WFA (Tab 3)
  document.getElementById('btn-submit-request')?.addEventListener('click', submitWfaRequest);

  // Absen Masuk & Keluar Buttons (Tab 2)
  document.getElementById('btn-absen-masuk')?.addEventListener('click', () => openAbsenModal('masuk'));
  document.getElementById('btn-absen-keluar')?.addEventListener('click', () => openAbsenModal('keluar'));
  document.getElementById('btn-close-modal')?.addEventListener('click', closeAbsenModal);
  document.getElementById('btn-submit-absen')?.addEventListener('click', submitAbsen);

  // Approval atasan button events
  document.getElementById('btn-confirm-approval')?.addEventListener('click', confirmApproval);
  document.getElementById('btn-cancel-approval')?.addEventListener('click', cancelApproval);

  // Backdrop modal absen click to close
  document.getElementById('modal-absen')?.addEventListener('click', (e) => {
    if (e.target === e.currentTarget) closeAbsenModal();
  });
}

// ============================================================
// ENTRY POINT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
  initApp();
});
