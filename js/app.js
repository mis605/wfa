// ============================================================
// APP.JS - Logika utama & UI controller (Revisi Flow Baru)
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
  user: null,
  karyawan: null,
  absensiHariIni: null,
  tipeAbsenDipilih: null,       // 'WFO' | 'WFA' | 'Visit'
  tipeRequestDipilih: 'WFA',    // 'WFA' | 'Visit'
  absenMode: 'masuk',           // 'masuk' | 'keluar'
  calendarDate: new Date(),
  pickerDate: new Date(),
  selectedDates: [],
  // Data kalender: semua absensi + request pending
  calendarAbsenList: [],        // [{nip, nama, tanggal, tipe, jamMasuk, jamKeluar}]
  calendarPendingList: [],      // [{nip, nama, tanggal, tipe}] - approved request tapi belum absen
  currentApprovalRequest: null
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
  
  if (appShell) appShell.classList.toggle('hidden', isAuthView);
  
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
    state.user = await authService.getUserProfile();
    
    try {
      await graphService.debugListColumns(APP_CONFIG.listKaryawanId, 'Karyawan');
      await graphService.debugListColumns(APP_CONFIG.listAbsensiId, 'Absensi');
      await graphService.debugListColumns(APP_CONFIG.listPermohonanWfaId, 'PermohonanWfa');
    } catch (e) {
      console.warn('[App] Failed to run debugListColumns:', e);
    }
    
    document.getElementById('loading-text').textContent = 'Memverifikasi data karyawan...';
    const userEmail = state.user.mail || state.user.userPrincipalName;
    state.karyawan = await graphService.getKaryawanByEmail(userEmail);
    
    if (!state.karyawan) {
      showView('login');
      document.getElementById('login-error').textContent = 
        'Akun Anda belum terdaftar sebagai karyawan. Silakan hubungi admin.';
      document.getElementById('login-error').classList.remove('hidden');
      return;
    }

    if (state.karyawan.statusAktif !== 'Aktif') {
      showView('login');
      document.getElementById('login-error').textContent = 'Akun Anda berstatus Tidak Aktif. Silakan hubungi admin.';
      document.getElementById('login-error').classList.remove('hidden');
      return;
    }

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

    // Cek URL parameters untuk approval atasan
    const params = new URLSearchParams(window.location.search);
    const action = params.get('action');
    const requestId = params.get('id');
    
    if (action && requestId) {
      await loadApprovalRequest(requestId, action);
    } else {
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
// TAB 1: KALENDER KEHADIRAN
// Menampilkan: siapa sudah absen (WFO/WFA/Visit) + siapa belum absen
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
    const userEmail = (state.karyawan.email || '').toLowerCase().trim();
    const userManagerEmail = (state.karyawan.emailAtasan || '').toLowerCase().trim();
    const userNip = (state.karyawan.nip || '').toLowerCase().trim();

    // Helper: cek apakah item ini visible untuk user login
    const isVisible = (itemManagerEmail, itemNip) => {
      const mgrEmail = (itemManagerEmail || '').toLowerCase().trim();
      const nip = (itemNip || '').toLowerCase().trim();
      // Bawahan langsung user
      const isSubordinate = userEmail && mgrEmail === userEmail;
      // Sesama peer (atasan sama, BUKAN atasan itu sendiri)
      const isSamePeer = userManagerEmail && mgrEmail === userManagerEmail && nip !== userNip;
      // Diri sendiri
      const isSelf = userNip && nip === userNip;
      return isSubordinate || isSamePeer || isSelf;
    };

    // 1. Ambil semua absensi bulan ini
    const allAbsensi = await graphService.getAbsensiBulanTertentu_All(month, year);
    state.calendarAbsenList = allAbsensi.filter(item =>
      isVisible(item.emailAtasan, item.nip)
    );

    // 2. Ambil semua request approved bulan ini (WFA + Visit)
    const allApproved = await graphService.getApprovedRequestByBulan(month, year);
    const approvedFiltered = allApproved.filter(item =>
      isVisible(item.emailAtasan, item.nip)
    );

    // 3. Hitung siapa yang approved tapi belum absen
    // pending = approved request tapi tidak ada record absensi pada tanggal tsb
    state.calendarPendingList = [];
    const todayStr = getTodayString();

    for (const req of approvedFiltered) {
      // Hanya tampilkan pending untuk hari ini atau sebelumnya (bukan masa depan)
      if (req.tanggal > todayStr) continue;
      
      const sudahAbsen = state.calendarAbsenList.some(
        ab => ab.nip === req.nip && ab.tanggal === req.tanggal
      );
      if (!sudahAbsen) {
        state.calendarPendingList.push(req);
      }
    }
    
    renderCalendarGrid(month, year);
  } catch (err) {
    container.innerHTML = `<div style="grid-column: span 7; text-align: center; padding: 24px; color: var(--red);">Gagal memuat kalender: ${err.message}</div>`;
  }
}

function renderCalendarGrid(month, year) {
  const container = document.getElementById('calendar-days-container');
  const daysInMonth = new Date(year, month, 0).getDate();
  const firstDayIndex = new Date(year, month - 1, 1).getDay();
  const startOffset = firstDayIndex === 0 ? 6 : firstDayIndex - 1;
  
  let cellsHtml = '';
  
  for (let i = 0; i < startOffset; i++) {
    cellsHtml += '<div class="calendar-day-cell calendar-day-cell--empty"></div>';
  }
  
  const today = new Date();
  const todayStr = getTodayString();
  
  for (let day = 1; day <= daysInMonth; day++) {
    const currentDayStr = `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    
    const absenOnDay = state.calendarAbsenList.filter(item => item.tanggal === currentDayStr);
    const pendingOnDay = state.calendarPendingList.filter(item => item.tanggal === currentDayStr);
    
    const isToday = currentDayStr === todayStr;
    const todayClass = isToday ? 'calendar-day-cell--today' : '';
    
    // Render nama dengan badge tipe
    const absenHtml = absenOnDay.map(w => {
      const tipeClass = (w.tipe || 'WFO').toLowerCase();
      return `<span class="calendar-entry calendar-entry--${tipeClass}" title="${w.nama} - ${w.tipe || 'WFO'}">${w.nama}</span>`;
    }).join('');

    const pendingHtml = pendingOnDay.map(w => {
      const tipeClass = (w.tipe || 'WFA').toLowerCase();
      return `<span class="calendar-entry calendar-entry--pending" title="${w.nama} - Belum Absen (${w.tipe})">${w.nama}*</span>`;
    }).join('');
    
    cellsHtml += `
      <div class="calendar-day-cell ${todayClass}" data-date="${currentDayStr}" style="cursor: pointer;">
        <span class="calendar-day-number">${day}</span>
        <div class="calendar-wfa-list">
          ${absenHtml}
          ${pendingHtml}
        </div>
      </div>
    `;
  }
  
  container.innerHTML = cellsHtml;
}

function openCalendarDetailModal(dateStr) {
  const modal = document.getElementById('modal-calendar-detail');
  const dateEl = document.getElementById('calendar-detail-date');
  const listEl = document.getElementById('calendar-detail-list');
  
  if (!modal || !dateEl || !listEl) return;
  
  dateEl.textContent = formatTanggal(dateStr + 'T00:00:00');
  
  const absenOnDay = state.calendarAbsenList.filter(item => item.tanggal === dateStr);
  const pendingOnDay = state.calendarPendingList.filter(item => item.tanggal === dateStr);
  
  let html = '';

  if (absenOnDay.length === 0 && pendingOnDay.length === 0) {
    html = `<div style="text-align: center; color: var(--text-muted); padding: 20px; font-style: italic;">Tidak ada data kehadiran pada tanggal ini.</div>`;
  } else {
    // Sudah absen
    if (absenOnDay.length > 0) {
      html += `<div class="detail-group-label">✅ Sudah Absen</div>`;
      html += absenOnDay.map(w => {
        const tipeClass = (w.tipe || 'WFO').toLowerCase();
        const durasi = w.jamMasuk && w.jamKeluar ? hitungDurasi(w.jamMasuk, w.jamKeluar) : '-';
        return `
          <div class="calendar-detail-item">
            <div class="calendar-detail-item__left">
              <strong>${w.nama}</strong>
              <span style="color: var(--text-muted); font-size: 0.72rem; font-family: var(--font-mono);">NIP: ${w.nip}</span>
            </div>
            <div class="calendar-detail-item__right">
              <span class="tipe-badge tipe-badge--${tipeClass}">${w.tipe || 'WFO'}</span>
              <span style="font-size: 0.72rem; color: var(--text-secondary); font-family: var(--font-mono);">
                ${w.jamMasuk ? w.jamMasuk.substring(0,5) : '--'} – ${w.jamKeluar ? w.jamKeluar.substring(0,5) : '--'}
              </span>
              ${durasi !== '-' ? `<span style="font-size: 0.7rem; color: var(--text-muted);">${durasi}</span>` : ''}
            </div>
          </div>
        `;
      }).join('');
    }

    // Belum absen (approved request tapi tidak absen)
    if (pendingOnDay.length > 0) {
      html += `<div class="detail-group-label" style="margin-top: 12px;">⏳ Belum Absen</div>`;
      html += pendingOnDay.map(w => {
        const tipeClass = (w.tipe || 'WFA').toLowerCase();
        return `
          <div class="calendar-detail-item">
            <div class="calendar-detail-item__left">
              <strong>${w.nama}</strong>
              <span style="color: var(--text-muted); font-size: 0.72rem; font-family: var(--font-mono);">NIP: ${w.nip}</span>
            </div>
            <div class="calendar-detail-item__right">
              <span class="tipe-badge tipe-badge--${tipeClass}">${w.tipe}</span>
              <span style="font-size: 0.7rem; color: var(--yellow);">Belum Absen</span>
            </div>
          </div>
        `;
      }).join('');
    }
  }
  
  listEl.innerHTML = html;
  modal.classList.remove('hidden');
  modal.classList.add('modal--show');
}

function closeCalendarDetailModal() {
  const modal = document.getElementById('modal-calendar-detail');
  if (modal) {
    modal.classList.remove('modal--show');
    setTimeout(() => modal.classList.add('hidden'), 300);
  }
}

// ============================================================
// TAB 2: ABSEN
// Flow: pilih tipe (WFO/WFA/Visit) → cek eligibility → absen
// ============================================================
async function loadAbsen() {
  // Update header
  document.getElementById('dash-nama').textContent = state.karyawan.nama;
  document.getElementById('dash-jabatan').textContent = `${state.karyawan.jabatan} • ${state.karyawan.departemen}`;
  document.getElementById('dash-tanggal').textContent = formatTanggal(new Date().toISOString());
  
  updateClock();

  // Reset UI ke awal
  resetAbsenUI();
  
  // Set config jam
  const infoMasuk = document.getElementById('info-jam-masuk');
  const infoKeluar = document.getElementById('info-jam-keluar');
  if (infoMasuk) infoMasuk.textContent = `${APP_CONFIG.jamMasukMulai}–${APP_CONFIG.jamMasukSelesai}`;
  if (infoKeluar) infoKeluar.textContent = `${APP_CONFIG.jamKeluarMulai}–${APP_CONFIG.jamKeluarSelesai}`;
}

function resetAbsenUI() {
  // Reset radio
  document.querySelectorAll('input[name="tipe-absen"]').forEach(r => r.checked = false);
  state.tipeAbsenDipilih = null;

  // Sembunyikan status dan dashboard
  const statusCard = document.getElementById('wfa-status-card');
  const dashboardContent = document.getElementById('absen-dashboard-content');
  statusCard.classList.add('hidden');
  statusCard.className = 'wfa-status-card hidden';
  dashboardContent.classList.add('hidden');
}

async function onTipeAbsenChange(tipe) {
  state.tipeAbsenDipilih = tipe;
  
  const statusCard = document.getElementById('wfa-status-card');
  const statusText = document.getElementById('wfa-status-text');
  const dashboardContent = document.getElementById('absen-dashboard-content');
  
  dashboardContent.classList.add('hidden');
  statusCard.classList.remove('hidden');
  statusCard.className = 'wfa-status-card';
  statusText.textContent = 'Memeriksa...';

  try {
    // Ambil absensi hari ini terlebih dahulu (untuk semua tipe)
    state.absensiHariIni = await graphService.getAbsensiHariIni(state.karyawan.nip);

    if (tipe === 'WFO') {
      // WFO: langsung boleh absen
      statusCard.className = 'wfa-status-card wfa-status-card--approved';
      statusText.textContent = '✅ Kehadiran WFO. Silakan absen.';
      dashboardContent.classList.remove('hidden');
      renderAbsenDashboard();

    } else {
      // WFA / Visit: cek apakah ada request approved untuk hari ini
      const todayStr = getTodayString();
      const requests = await graphService.getPermohonanWfa(
        state.user.mail || state.user.userPrincipalName
      );
      
      const approvedForToday = requests.some(req => {
        if (req.status !== 'Approved') return false;
        if ((req.tipe || 'WFA') !== tipe) return false;
        const dates = req.tanggalWfa ? req.tanggalWfa.split(',').map(d => d.trim()) : [];
        return dates.includes(todayStr);
      });

      if (!approvedForToday) {
        statusCard.className = 'wfa-status-card wfa-status-card--none';
        statusText.textContent = `⚠️ Anda belum memiliki pengajuan ${tipe} yang disetujui untuk hari ini.`;
        // Tampilkan modal info
        openBelumRequestModal(tipe);
        return;
      }

      statusCard.className = 'wfa-status-card wfa-status-card--approved';
      statusText.textContent = `✅ Pengajuan ${tipe} Anda disetujui untuk hari ini.`;
      dashboardContent.classList.remove('hidden');
      renderAbsenDashboard();
    }

  } catch (err) {
    statusCard.className = 'wfa-status-card wfa-status-card--rejected';
    statusText.textContent = 'Gagal memeriksa data: ' + err.message;
  }
}

function openBelumRequestModal(tipe) {
  const modal = document.getElementById('modal-belum-request');
  const text = document.getElementById('belum-request-text');
  text.textContent = `Anda belum memiliki pengajuan ${tipe} yang disetujui untuk hari ini. Ajukan terlebih dahulu melalui halaman Request.`;
  modal.classList.remove('hidden');
  modal.classList.add('modal--show');
}

function closeBelumRequestModal() {
  const modal = document.getElementById('modal-belum-request');
  modal.classList.remove('modal--show');
  setTimeout(() => modal.classList.add('hidden'), 300);
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
}

function openAbsenModal(mode) {
  state.absenMode = mode;
  const modal = document.getElementById('modal-absen');
  
  // Update konfirmasi info
  const now = new Date();
  const jamStr = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
  
  document.getElementById('konfirmasi-tipe').textContent = state.tipeAbsenDipilih || 'WFO';
  document.getElementById('konfirmasi-waktu').textContent = jamStr;
  document.getElementById('konfirmasi-mode').textContent = mode === 'masuk' ? 'Absen Masuk' : 'Absen Keluar';
  document.getElementById('modal-absen-title').textContent = mode === 'masuk' ? 'Konfirmasi Absen Masuk' : 'Konfirmasi Absen Keluar';
  
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
  const tipe = state.tipeAbsenDipilih || 'WFO';
  
  setLoading(btn, true);
  
  try {
    const payload = {
      nip: state.karyawan.nip,
      nama: state.karyawan.nama,
      tipe: tipe,
      emailAtasan: state.karyawan.emailAtasan || ''
    };
    
    if (state.absenMode === 'masuk') {
      const result = await graphService.absenMasuk(payload);
      showToast(`✓ Absen masuk ${tipe} berhasil! Status: ${result.status}`, 'success');
    } else {
      const result = await graphService.absenKeluar(payload);
      showToast(`✓ Absen keluar berhasil! Jam: ${result.jamKeluar.substring(0, 5)}`, 'success');
    }
    
    closeAbsenModal();
    
    // Refresh data absen
    state.absensiHariIni = await graphService.getAbsensiHariIni(state.karyawan.nip);
    renderAbsenDashboard();
    
  } catch (err) {
    showToast(err.message, 'error');
  } finally {
    setLoading(btn, false, 'Konfirmasi & Absen');
  }
}

// ============================================================
// TAB 3: REQUEST WFA / VISIT
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
  
  // Tampilkan email atasan
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
  
  // Set tipe request default
  if (!state.tipeRequestDipilih) state.tipeRequestDipilih = 'WFA';
  const radioWfa = document.getElementById('radio-req-wfa');
  const radioVisit = document.getElementById('radio-req-visit');
  if (state.tipeRequestDipilih === 'WFA' && radioWfa) radioWfa.checked = true;
  if (state.tipeRequestDipilih === 'Visit' && radioVisit) radioVisit.checked = true;

  renderPickerGrid(month, year);
  renderSelectedDates();
}

function renderPickerGrid(month, year) {
  const container = document.getElementById('picker-days-container');
  const daysInMonth = new Date(year, month, 0).getDate();
  const firstDayIndex = new Date(year, month - 1, 1).getDay();
  const startOffset = firstDayIndex === 0 ? 6 : firstDayIndex - 1;
  
  let cellsHtml = '';
  
  for (let i = 0; i < startOffset; i++) {
    cellsHtml += '<div class="picker-day-cell picker-day-cell--empty"></div>';
  }
  
  const todayStr = getTodayString();
  
  for (let day = 1; day <= daysInMonth; day++) {
    const currentDayStr = `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    const isSelected = state.selectedDates.includes(currentDayStr);
    const isDisabled = currentDayStr < todayStr;
    
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
  
  container.querySelectorAll('.picker-day-cell:not(.picker-day-cell--empty):not(.picker-day-cell--disabled)').forEach(el => {
    el.addEventListener('click', () => {
      const dateStr = el.dataset.date;
      const idx = state.selectedDates.indexOf(dateStr);
      if (idx > -1) {
        state.selectedDates.splice(idx, 1);
      } else {
        state.selectedDates.push(dateStr);
        state.selectedDates.sort();
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
  const tipe = state.tipeRequestDipilih;
  
  const isValid = state.selectedDates.length > 0 && catatan.length > 0 && managerEmail.length > 0 && !!tipe;
  btnSubmit.disabled = !isValid;
}

async function submitWfaRequest() {
  const catatan = document.getElementById('req-catatan').value.trim();
  const managerEmail = state.karyawan.emailAtasan || '';
  const tipe = state.tipeRequestDipilih;
  
  if (state.selectedDates.length === 0) {
    showToast('Pilih minimal satu tanggal pengajuan.', 'warning');
    return;
  }
  if (!catatan) {
    showToast('Alasan pengajuan wajib diisi.', 'warning');
    return;
  }
  if (!tipe) {
    showToast('Pilih tipe pengajuan (WFA atau Visit).', 'warning');
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
      tanggalWfa: state.selectedDates.join(', '),
      tipe: tipe,
      emailAtasan: managerEmail,
      catatanUser: catatan
    };
    
    await graphService.tambahPermohonanWfa(payload);
    showToast(`✓ Pengajuan ${tipe} berhasil dikirim! Anda sekarang bisa absen setelah email terkirim.`, 'success');
    
    // Reset form
    state.selectedDates = [];
    document.getElementById('req-catatan').value = '';
    
    const month = state.pickerDate.getMonth() + 1;
    const year = state.pickerDate.getFullYear();
    renderPickerGrid(month, year);
    renderSelectedDates();
    
    // Langsung arahkan ke tab absen, dan refresh status absen secara real-time
    showView('absen');
    await loadAbsen();
    
    // Auto-pilih tipe yang baru disubmit supaya user bisa langsung absen
    const radioEl = document.getElementById(tipe === 'WFA' ? 'radio-wfa' : 'radio-visit');
    if (radioEl) {
      radioEl.checked = true;
      await onTipeAbsenChange(tipe);
    }
    
  } catch (err) {
    showToast('Gagal memproses pengajuan: ' + err.message, 'error');
  } finally {
    setLoading(btn, false, 'Kirim Pengajuan');
  }
}

// ============================================================
// VIEW APPROVAL ATASAN (URL ROUTING)
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
    if (!req) throw new Error('Pengajuan tidak ditemukan.');
    
    state.currentApprovalRequest = req;
    
    document.getElementById('app-nama').textContent = req.nama;
    document.getElementById('app-nip').textContent = req.nip;
    document.getElementById('app-catatan-karyawan').textContent = req.catatanUser || '-';
    
    // Badge tipe
    const tipeEl = document.getElementById('app-tipe');
    const tipe = req.tipe || 'WFA';
    tipeEl.textContent = tipe;
    tipeEl.className = `status-badge tipe-badge tipe-badge--${tipe.toLowerCase()}`;
    
    const datesEl = document.getElementById('app-tanggal-list');
    const dates = req.tanggalWfa ? req.tanggalWfa.split(',') : [];
    datesEl.innerHTML = dates.map(d => `<span class="selected-date-tag">${formatTanggalPendek(d.trim())}</span>`).join('');
    
    const btnConfirm = document.getElementById('btn-confirm-approval');
    const textarea = document.getElementById('app-catatan-atasan');
    
    if (req.status === 'Rejected') {
      titleText.textContent = `Pengajuan ${tipe}: DIBATALKAN`;
      textarea.value = req.catatanAtasan || '';
      textarea.disabled = true;
      btnConfirm.disabled = true;
      btnConfirm.textContent = 'Sudah Ditolak / Dibatalkan';
      btnConfirm.classList.add('btn--disabled');
    } else {
      // Approved (auto) — atasan hanya bisa reject/batalkan
      titleText.textContent = `Batalkan Pengajuan ${tipe}`;
      textarea.value = '';
      textarea.placeholder = 'Contoh: Ada rapat luring penting di kantor...';
      textarea.disabled = false;
      btnConfirm.disabled = false;
      btnConfirm.textContent = 'Konfirmasi Batalkan / Tolak Pengajuan';
      btnConfirm.classList.remove('btn--disabled');
    }
    
    loadingState.classList.add('hidden');
    boxContent.classList.remove('hidden');
    
  } catch (err) {
    loadingState.classList.add('hidden');
    showToast('Gagal memuat detail pengajuan: ' + err.message, 'error');
    showView('calendar');
    await loadCalendar();
  }
}

async function confirmApproval() {
  const req = state.currentApprovalRequest;
  if (!req) return;
  
  const catatanAtasan = document.getElementById('app-catatan-atasan').value.trim();
  const btn = document.getElementById('btn-confirm-approval');
  setLoading(btn, true);
  
  try {
    // Satu-satunya aksi dari email adalah reject
    await graphService.updateStatusPermohonanWfa(req.id, 'Rejected', catatanAtasan);
    showToast('✓ Pengajuan berhasil dibatalkan/ditolak!', 'success');
    
    window.history.replaceState({}, document.title, window.location.pathname);
    showView('calendar');
    await loadCalendar();
  } catch (err) {
    showToast('Gagal memproses: ' + err.message, 'error');
  } finally {
    setLoading(btn, false, 'Konfirmasi Batalkan / Tolak Pengajuan');
  }
}

function cancelApproval() {
  window.history.replaceState({}, document.title, window.location.pathname);
  showView('calendar');
  loadCalendar();
}

// ============================================================
// BIND EVENTS
// ============================================================
function bindEvents() {
  // Login
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

  // Logout
  document.getElementById('btn-logout')?.addEventListener('click', async () => {
    await authService.logout();
    window.location.reload();
  });

  // Bottom Nav
  document.querySelectorAll('[data-nav]').forEach(el => {
    el.addEventListener('click', async () => {
      const view = el.dataset.nav;
      showView(view);
      if (view === 'calendar') await loadCalendar();
      if (view === 'absen') await loadAbsen();
      if (view === 'request') loadRequestView();
    });
  });

  // Calendar navigasi
  document.getElementById('btn-cal-prev')?.addEventListener('click', () => loadCalendar(-1));
  document.getElementById('btn-cal-next')?.addEventListener('click', () => loadCalendar(1));

  // Request picker navigasi
  document.getElementById('btn-picker-prev')?.addEventListener('click', () => loadRequestView(-1));
  document.getElementById('btn-picker-next')?.addEventListener('click', () => loadRequestView(1));

  // Radio tipe absen (Tab 2)
  document.querySelectorAll('input[name="tipe-absen"]').forEach(radio => {
    radio.addEventListener('change', async (e) => {
      await onTipeAbsenChange(e.target.value);
    });
  });

  // Radio tipe request (Tab 3)
  document.querySelectorAll('input[name="tipe-request"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
      state.tipeRequestDipilih = e.target.value;
      validateRequestForm();
    });
  });

  // Input catatan request
  document.getElementById('req-catatan')?.addEventListener('input', validateRequestForm);

  // Submit request
  document.getElementById('btn-submit-request')?.addEventListener('click', submitWfaRequest);

  // Tombol absen masuk/keluar
  document.getElementById('btn-absen-masuk')?.addEventListener('click', () => openAbsenModal('masuk'));
  document.getElementById('btn-absen-keluar')?.addEventListener('click', () => openAbsenModal('keluar'));
  document.getElementById('btn-close-modal')?.addEventListener('click', closeAbsenModal);
  document.getElementById('btn-submit-absen')?.addEventListener('click', submitAbsen);

  // Modal backdrop absen
  document.getElementById('modal-absen')?.addEventListener('click', (e) => {
    if (e.target === e.currentTarget) closeAbsenModal();
  });

  // Modal belum request
  document.getElementById('btn-belum-request-tutup')?.addEventListener('click', closeBelumRequestModal);
  document.getElementById('btn-belum-request-ke-request')?.addEventListener('click', () => {
    closeBelumRequestModal();
    showView('request');
    loadRequestView();
    // Set tipe request sesuai tipe yang dipilih di tab absen
    if (state.tipeAbsenDipilih && state.tipeAbsenDipilih !== 'WFO') {
      state.tipeRequestDipilih = state.tipeAbsenDipilih;
      const radioEl = document.getElementById(
        state.tipeAbsenDipilih === 'WFA' ? 'radio-req-wfa' : 'radio-req-visit'
      );
      if (radioEl) radioEl.checked = true;
    }
  });
  document.getElementById('modal-belum-request')?.addEventListener('click', (e) => {
    if (e.target === e.currentTarget) closeBelumRequestModal();
  });

  // Calendar detail modal
  document.getElementById('calendar-days-container')?.addEventListener('click', (e) => {
    const cell = e.target.closest('.calendar-day-cell');
    if (cell && cell.dataset.date) {
      openCalendarDetailModal(cell.dataset.date);
    }
  });
  document.getElementById('btn-close-calendar-modal')?.addEventListener('click', closeCalendarDetailModal);
  document.getElementById('modal-calendar-detail')?.addEventListener('click', (e) => {
    if (e.target === e.currentTarget) closeCalendarDetailModal();
  });

  // Approval
  document.getElementById('btn-confirm-approval')?.addEventListener('click', confirmApproval);
  document.getElementById('btn-cancel-approval')?.addEventListener('click', cancelApproval);
}

// ============================================================
// ENTRY POINT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
  initApp();
});
