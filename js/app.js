// ============================================================
// APP.JS - Logika utama & UI controller
// Improvements: No.6 auto-detect absen, No.7 notif request ditolak,
//               No.8 tombol Hari Ini di kalender, No.9 loading progress absen,
//               No.11 hapus debug, No.12 error boundary
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
  tipeAbsenDipilih: null,
  tipeRequestDipilih: 'WFA',
  absenMode: 'masuk',
  calendarDate: new Date(),
  pickerDate: new Date(),
  selectedDates: [],
  calendarAbsenList: [],
  calendarPendingList: [],
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
      if (name === 'login' || name === 'loading') el.classList.add('hidden');
      else el.classList.toggle('hidden', name !== viewName);
    }
  });
  state.currentView = viewName;
  document.querySelectorAll('[data-nav]').forEach(el => {
    el.classList.toggle('nav__item--active', el.dataset.nav === viewName);
  });
}

// ============================================================
// NO.12 — ERROR BOUNDARY: tampilkan pesan ramah saat crash fatal
// ============================================================
function showFatalError(message) {
  // Sembunyikan semua view, tampilkan error card
  Object.values(views).forEach(el => el?.classList.add('hidden'));
  const appShell = document.getElementById('app-shell');
  if (appShell) appShell.classList.add('hidden');

  let errorEl = document.getElementById('view-fatal-error');
  if (!errorEl) {
    errorEl = document.createElement('div');
    errorEl.id = 'view-fatal-error';
    errorEl.className = 'fatal-error-view';
    errorEl.innerHTML = `
      <div class="fatal-error-card">
        <div class="fatal-error-icon">⚠️</div>
        <h2 class="fatal-error-title">Terjadi Kesalahan</h2>
        <p class="fatal-error-msg" id="fatal-error-msg"></p>
        <button class="btn-primary" id="btn-fatal-reload" style="margin-top:20px;">
          Muat Ulang Aplikasi
        </button>
      </div>
    `;
    document.getElementById('app').appendChild(errorEl);
    document.getElementById('btn-fatal-reload').addEventListener('click', () => {
      window.location.reload();
    });
  }
  document.getElementById('fatal-error-msg').textContent = message || 'Silakan muat ulang aplikasi.';
  errorEl.classList.remove('hidden');
}

// Global uncaught error handler
window.addEventListener('unhandledrejection', (event) => {
  console.error('[App] Unhandled rejection:', event.reason);
  // Hanya tampilkan fatal error kalau belum ada view yang aktif
  if (!state.user) {
    showFatalError('Gagal memuat aplikasi. Periksa koneksi internet Anda.');
  }
});

// ============================================================
// INISIALISASI APP
// ============================================================
async function initApp() {
  showView('loading');
  try {
    await authService.init();
    if (!authService.isLoggedIn()) { showView('login'); return; }
    await loadUserSession();
  } catch (err) {
    console.error('Init error:', err);
    showFatalError('Gagal menginisialisasi aplikasi: ' + err.message);
  }
}

async function loadUserSession() {
  showView('loading');
  document.getElementById('loading-text').textContent = 'Memuat profil Microsoft 365...';

  try {
    state.user = await authService.getUserProfile();

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
    // Isi info dropdown
    const dropdownNama = document.getElementById('dropdown-nama');
    const dropdownEmail = document.getElementById('dropdown-email');
    if (dropdownNama) dropdownNama.textContent = state.karyawan.nama || '-';
    if (dropdownEmail) dropdownEmail.textContent = state.user.mail || state.user.userPrincipalName || '-';

    document.querySelectorAll('.user-avatar, .greeting-avatar').forEach(el => {
      if (photoUrl) { el.style.backgroundImage = `url(${photoUrl})`; el.textContent = ''; }
      else { el.style.backgroundImage = 'none'; el.textContent = initials; }
    });

    // NO.7 — Cek request yang ditolak sejak terakhir login
    await checkRejectedRequests();

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
    showFatalError('Gagal memuat sesi: ' + err.message);
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
    clockEl.textContent = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  };
  update();
  if (!window._clockInterval) window._clockInterval = setInterval(update, 1000);
}

function getStatusClass(status) {
  const map = {
    'Tepat Waktu': 'success', 'Terlambat Ringan': 'warning',
    'Terlambat': 'danger', 'Selesai': 'success', 'Belum Absen': 'pending',
  };
  return map[status] || 'pending';
}

// ============================================================
// NO.7 — CEK REQUEST DITOLAK (notifikasi in-app)
// ============================================================
async function checkRejectedRequests() {
  try {
    const userEmail = state.user.mail || state.user.userPrincipalName;
    const requests = await graphService.getPermohonanWfa(userEmail);

    // Ambil key dari localStorage untuk track request yang sudah dinotif
    // Semua ID dikonversi ke String untuk menghindari type mismatch (number vs string)
    const notifiedKey = `notified_rejected_${state.karyawan.nip}`;
    const alreadyNotified = JSON.parse(localStorage.getItem(notifiedKey) || '[]')
      .map(id => String(id));

    const newlyRejected = requests.filter(req =>
      req.status === 'Rejected' && !alreadyNotified.includes(String(req.id))
    );

    if (newlyRejected.length > 0) {
      // Tandai sebagai sudah dinotif — simpan sebagai String agar konsisten
      const newNotified = [...alreadyNotified, ...newlyRejected.map(r => String(r.id))];
      localStorage.setItem(notifiedKey, JSON.stringify(newNotified));

      // Tampilkan notifikasi in-app
      showRejectedNotification(newlyRejected);
    }
  } catch (err) {
    // Tidak fatal, cukup log
    console.warn('Gagal cek request ditolak:', err.message);
  }
}

function showRejectedNotification(rejectedList) {
  const modal = document.getElementById('modal-rejected-notif');
  const listEl = document.getElementById('rejected-notif-list');
  if (!modal || !listEl) return;

  listEl.innerHTML = rejectedList.map(req => {
    const dates = req.tanggalWfa ? req.tanggalWfa.split(',').map(d => formatTanggalPendek(d.trim())).join(', ') : '-';
    return `
      <div class="rejected-notif-item">
        <span class="tipe-badge tipe-badge--${(req.tipe||'WFA').toLowerCase()}">${req.tipe || 'WFA'}</span>
        <span class="rejected-notif-tanggal">${dates}</span>
        ${req.catatanAtasan ? `<span class="rejected-notif-catatan">"${req.catatanAtasan}"</span>` : ''}
      </div>
    `;
  }).join('');

  modal.classList.remove('hidden');
  modal.classList.add('modal--show');
}

function closeRejectedNotif() {
  const modal = document.getElementById('modal-rejected-notif');
  if (modal) {
    modal.classList.remove('modal--show');
    setTimeout(() => modal.classList.add('hidden'), 300);
  }
}

// ============================================================
// TAB 1: KALENDER KEHADIRAN
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

  // NO.8 — Tombol "Hari Ini": tampilkan jika bukan bulan sekarang
  const now = new Date();
  const isCurrentMonth = month === now.getMonth() + 1 && year === now.getFullYear();
  const btnHariIni = document.getElementById('btn-cal-today');
  if (btnHariIni) btnHariIni.classList.toggle('hidden', isCurrentMonth);

  const container = document.getElementById('calendar-days-container');
  container.innerHTML = '<div style="grid-column: span 7; text-align: center; padding: 24px;"><span class="spinner"></span></div>';

  try {
    const userEmail = (state.karyawan.email || '').toLowerCase().trim();
    const userManagerEmail = (state.karyawan.emailAtasan || '').toLowerCase().trim();
    const userNip = (state.karyawan.nip || '').toLowerCase().trim();

    const SUPERUSER_EMAILS = (APP_CONFIG.calendarSuperuserEmails || []).map(e => e.toLowerCase().trim());
    const isSuperuser = SUPERUSER_EMAILS.includes(userEmail);

    const isVisible = (itemManagerEmail, itemNip) => {
      if (isSuperuser) return true;
      const mgrEmail = (itemManagerEmail || '').toLowerCase().trim();
      const nip = (itemNip || '').toLowerCase().trim();
      const isSubordinate = userEmail && mgrEmail === userEmail;
      const isSamePeer = userManagerEmail && mgrEmail === userManagerEmail && nip !== userNip;
      const isSelf = userNip && nip === userNip;
      return isSubordinate || isSamePeer || isSelf;
    };

    const [allAbsensi, allApproved] = await Promise.all([
      graphService.getAbsensiBulanTertentu_All(month, year),
      graphService.getApprovedRequestByBulan(month, year)
    ]);

    state.calendarAbsenList = allAbsensi.filter(item => isVisible(item.emailAtasan, item.nip));

    const approvedFiltered = allApproved.filter(item => isVisible(item.emailAtasan, item.nip));
    const todayStr = getTodayString();

    state.calendarPendingList = approvedFiltered.filter(req => {
      if (req.tanggal > todayStr) return false;
      return !state.calendarAbsenList.some(ab => ab.nip === req.nip && ab.tanggal === req.tanggal);
    });

    renderCalendarGrid(month, year);
  } catch (err) {
    container.innerHTML = `
      <div class="calendar-error-state">
        <p>⚠️ Gagal memuat kalender</p>
        <p style="font-size:0.8rem;margin-top:4px;">${err.message}</p>
        <button class="btn-sm" onclick="loadCalendar()" style="margin-top:12px;">Coba Lagi</button>
      </div>`;
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

  const todayStr = getTodayString();

  for (let day = 1; day <= daysInMonth; day++) {
    const currentDayStr = `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    const absenOnDay = state.calendarAbsenList.filter(item => item.tanggal === currentDayStr);
    const pendingOnDay = state.calendarPendingList.filter(item => item.tanggal === currentDayStr);
    const isToday = currentDayStr === todayStr;

    const absenHtml = absenOnDay.map(w => {
      const tipeClass = (w.tipe || 'WFO').toLowerCase();
      return `<span class="calendar-entry calendar-entry--${tipeClass}" title="${w.nama} - ${w.tipe||'WFO'}">${w.nama}</span>`;
    }).join('');

    const pendingHtml = pendingOnDay.map(w =>
      `<span class="calendar-entry calendar-entry--pending" title="${w.nama} - Belum Absen (${w.tipe})">${w.nama}*</span>`
    ).join('');

    cellsHtml += `
      <div class="calendar-day-cell ${isToday ? 'calendar-day-cell--today' : ''}" data-date="${currentDayStr}" style="cursor:pointer;">
        <span class="calendar-day-number">${day}</span>
        <div class="calendar-wfa-list">${absenHtml}${pendingHtml}</div>
      </div>`;
  }

  container.innerHTML = cellsHtml;
}

// NO.8 — Kembali ke bulan hari ini
function goToToday() {
  state.calendarDate = new Date();
  loadCalendar(0);
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
    html = `<div style="text-align:center;color:var(--text-muted);padding:20px;font-style:italic;">Tidak ada data kehadiran pada tanggal ini.</div>`;
  } else {
    if (absenOnDay.length > 0) {
      html += `<div class="detail-group-label">✅ Sudah Absen</div>`;
      html += absenOnDay.map(w => {
        const tipeClass = (w.tipe || 'WFO').toLowerCase();
        const durasi = w.jamMasuk && w.jamKeluar ? hitungDurasi(w.jamMasuk, w.jamKeluar) : null;
        return `
          <div class="calendar-detail-item">
            <div class="calendar-detail-item__left">
              <strong>${w.nama}</strong>
              <span style="color:var(--text-muted);font-size:0.72rem;font-family:var(--font-mono);">NIP: ${w.nip}</span>
            </div>
            <div class="calendar-detail-item__right">
              <span class="tipe-badge tipe-badge--${tipeClass}">${w.tipe||'WFO'}</span>
              <span style="font-size:0.72rem;color:var(--text-secondary);font-family:var(--font-mono);">
                ${w.jamMasuk ? w.jamMasuk.substring(0,5) : '--'} – ${w.jamKeluar ? w.jamKeluar.substring(0,5) : '--'}
              </span>
              ${durasi ? `<span style="font-size:0.7rem;color:var(--text-muted);">${durasi}</span>` : ''}
            </div>
          </div>`;
      }).join('');
    }

    if (pendingOnDay.length > 0) {
      html += `<div class="detail-group-label" style="margin-top:12px;">⏳ Belum Absen</div>`;
      html += pendingOnDay.map(w => `
        <div class="calendar-detail-item">
          <div class="calendar-detail-item__left">
            <strong>${w.nama}</strong>
            <span style="color:var(--text-muted);font-size:0.72rem;font-family:var(--font-mono);">NIP: ${w.nip}</span>
          </div>
          <div class="calendar-detail-item__right">
            <span class="tipe-badge tipe-badge--${(w.tipe||'WFA').toLowerCase()}">${w.tipe}</span>
            <span style="font-size:0.7rem;color:var(--yellow);">Belum Absen</span>
          </div>
        </div>`).join('');
    }
  }

  listEl.innerHTML = html;
  modal.classList.remove('hidden');
  modal.classList.add('modal--show');
}

function closeCalendarDetailModal() {
  const modal = document.getElementById('modal-calendar-detail');
  if (modal) { modal.classList.remove('modal--show'); setTimeout(() => modal.classList.add('hidden'), 300); }
}

// ============================================================
// TAB 2: ABSEN
// NO.6 — Auto-detect status absen hari ini tanpa perlu pilih radio dulu
// ============================================================
async function loadAbsen() {
  document.getElementById('dash-nama').textContent = state.karyawan.nama;
  document.getElementById('dash-jabatan').textContent = `${state.karyawan.jabatan} • ${state.karyawan.departemen}`;
  document.getElementById('dash-tanggal').textContent = formatTanggal(new Date().toISOString());
  updateClock();

  const infoMasuk = document.getElementById('info-jam-masuk');
  const infoKeluar = document.getElementById('info-jam-keluar');
  if (infoMasuk) infoMasuk.textContent = `${APP_CONFIG.jamMasukMulai}–${APP_CONFIG.jamMasukSelesai}`;
  if (infoKeluar) infoKeluar.textContent = `${APP_CONFIG.jamKeluarMulai}–${APP_CONFIG.jamKeluarSelesai}`;

  // NO.6 — Cek apakah sudah absen hari ini, auto-detect tipe & tampilkan dashboard
  try {
    const absenHariIni = await graphService.getAbsensiHariIni(state.karyawan.nip);

    if (absenHariIni) {
      // Sudah absen — auto-select tipe dan tampilkan dashboard langsung
      state.absensiHariIni = absenHariIni;
      state.tipeAbsenDipilih = absenHariIni.tipe || 'WFO';

      // Centang radio sesuai tipe
      const radioId = { WFO: 'radio-wfo', WFA: 'radio-wfa', Visit: 'radio-visit' }[state.tipeAbsenDipilih];
      if (radioId) {
        const radioEl = document.getElementById(radioId);
        if (radioEl) radioEl.checked = true;
      }

      const statusCard = document.getElementById('wfa-status-card');
      const statusText = document.getElementById('wfa-status-text');
      statusCard.classList.remove('hidden');
      statusCard.className = 'wfa-status-card wfa-status-card--approved';
      statusText.textContent = `✅ Anda sudah absen ${state.tipeAbsenDipilih} hari ini.`;
      document.getElementById('absen-dashboard-content').classList.remove('hidden');
      renderAbsenDashboard();
      return;
    }
  } catch (err) {
    console.warn('Auto-detect absen gagal:', err.message);
  }

  // Belum absen — reset ke state awal
  resetAbsenUI();
}

function resetAbsenUI() {
  document.querySelectorAll('input[name="tipe-absen"]').forEach(r => r.checked = false);
  state.tipeAbsenDipilih = null;
  const statusCard = document.getElementById('wfa-status-card');
  statusCard.classList.add('hidden');
  statusCard.className = 'wfa-status-card hidden';
  document.getElementById('absen-dashboard-content').classList.add('hidden');
}

async function onTipeAbsenChange(tipe) {
  // Jika sudah absen dengan tipe lain, jangan izinkan ganti tipe
  if (state.absensiHariIni && state.absensiHariIni.tipe !== tipe) {
    showToast(`Anda sudah absen ${state.absensiHariIni.tipe} hari ini, tidak bisa ganti tipe.`, 'warning');
    // Kembalikan radio ke tipe yang benar
    const radioId = { WFO: 'radio-wfo', WFA: 'radio-wfa', Visit: 'radio-visit' }[state.absensiHariIni.tipe];
    if (radioId) {
      const el = document.getElementById(radioId);
      if (el) el.checked = true;
    }
    return;
  }

  state.tipeAbsenDipilih = tipe;
  const statusCard = document.getElementById('wfa-status-card');
  const statusText = document.getElementById('wfa-status-text');
  const dashboardContent = document.getElementById('absen-dashboard-content');

  dashboardContent.classList.add('hidden');
  statusCard.classList.remove('hidden');
  statusCard.className = 'wfa-status-card';
  statusText.textContent = 'Memeriksa...';

  try {
    state.absensiHariIni = await graphService.getAbsensiHariIni(state.karyawan.nip);

    if (tipe === 'WFO') {
      statusCard.className = 'wfa-status-card wfa-status-card--approved';
      statusText.textContent = '✅ Kehadiran WFO. Silakan absen.';
      dashboardContent.classList.remove('hidden');
      renderAbsenDashboard();
    } else {
      const todayStr = getTodayString();
      const requests = await graphService.getPermohonanWfa(state.user.mail || state.user.userPrincipalName);
      const approvedForToday = requests.some(req => {
        if (req.status !== 'Approved') return false;
        if ((req.tipe || 'WFA') !== tipe) return false;
        const dates = req.tanggalWfa ? req.tanggalWfa.split(',').map(d => d.trim()) : [];
        return dates.includes(todayStr);
      });

      if (!approvedForToday) {
        statusCard.className = 'wfa-status-card wfa-status-card--none';
        statusText.textContent = `⚠️ Anda belum memiliki pengajuan ${tipe} yang disetujui untuk hari ini.`;
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
    masukEl.textContent = '--:--'; keluarEl.textContent = '--:--'; durasiEl.textContent = '--';
    btnMasuk.disabled = false; btnMasuk.classList.remove('btn--disabled');
    btnKeluar.disabled = true; btnKeluar.classList.add('btn--disabled');
  } else {
    const isClockedOut = !!a.jamKeluar;
    statusEl.textContent = isClockedOut ? 'Selesai' : a.status;
    statusEl.className = `status-badge status--${getStatusClass(a.status)}`;
    masukEl.textContent = formatJam(a.jamMasuk);
    keluarEl.textContent = isClockedOut ? formatJam(a.jamKeluar) : '--:--';
    durasiEl.textContent = hitungDurasi(a.jamMasuk, a.jamKeluar);

    if (!isClockedOut) {
      btnMasuk.disabled = true; btnMasuk.classList.add('btn--disabled');
      btnKeluar.disabled = false; btnKeluar.classList.remove('btn--disabled');
    } else {
      btnMasuk.disabled = true; btnMasuk.classList.add('btn--disabled');
      btnKeluar.disabled = true; btnKeluar.classList.add('btn--disabled');
    }
  }
}

// NO.9 — Modal konfirmasi dengan progress bar saat submit
function openAbsenModal(mode) {
  state.absenMode = mode;
  const modal = document.getElementById('modal-absen');
  const now = new Date();
  const jamStr = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });

  document.getElementById('konfirmasi-tipe').textContent = state.tipeAbsenDipilih || 'WFO';
  document.getElementById('konfirmasi-waktu').textContent = jamStr;
  document.getElementById('konfirmasi-mode').textContent = mode === 'masuk' ? 'Absen Masuk' : 'Absen Keluar';
  document.getElementById('modal-absen-title').textContent = mode === 'masuk' ? 'Konfirmasi Absen Masuk' : 'Konfirmasi Absen Keluar';

  // Reset progress bar
  const progressBar = document.getElementById('absen-progress');
  if (progressBar) { progressBar.style.width = '0%'; progressBar.parentElement.classList.add('hidden'); }
  const btn = document.getElementById('btn-submit-absen');
  if (btn) { btn.disabled = false; btn.textContent = 'Konfirmasi & Absen'; }

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

  // NO.9 — Tampilkan progress bar
  const progressWrap = document.getElementById('absen-progress-wrap');
  const progressBar = document.getElementById('absen-progress');
  if (progressWrap) progressWrap.classList.remove('hidden');

  setLoading(btn, true);

  // Animasi progress: 0% → 60% saat request, 60% → 100% saat selesai
  let progress = 0;
  const progressInterval = setInterval(() => {
    if (progress < 60) { progress += 4; if (progressBar) progressBar.style.width = progress + '%'; }
  }, 80);

  try {
    const payload = {
      nip: state.karyawan.nip,
      nama: state.karyawan.nama,
      tipe,
      emailAtasan: state.karyawan.emailAtasan || ''
    };

    if (state.absenMode === 'masuk') {
      const result = await graphService.absenMasuk(payload);
      clearInterval(progressInterval);
      if (progressBar) progressBar.style.width = '100%';
      showToast(`✓ Absen masuk ${tipe} berhasil! Status: ${result.status}`, 'success');
    } else {
      const result = await graphService.absenKeluar(payload);
      clearInterval(progressInterval);
      if (progressBar) progressBar.style.width = '100%';
      showToast(`✓ Absen keluar berhasil! Jam: ${result.jamKeluar.substring(0, 5)}`, 'success');
    }

    setTimeout(() => {
      closeAbsenModal();
    }, 400);

    state.absensiHariIni = await graphService.getAbsensiHariIni(state.karyawan.nip);
    renderAbsenDashboard();

  } catch (err) {
    clearInterval(progressInterval);
    if (progressWrap) progressWrap.classList.add('hidden');
    if (progressBar) progressBar.style.width = '0%';
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
  document.getElementById('picker-month-label').textContent =
    state.pickerDate.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });

  const managerEmail = state.karyawan.emailAtasan || '';
  const managerEl = document.getElementById('req-manager-email');
  if (managerEl) {
    managerEl.textContent = managerEmail || 'Belum dikonfigurasi (Hubungi Admin)';
    managerEl.style.color = managerEmail ? '' : 'var(--red)';
  }

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
  const todayStr = getTodayString();

  let cellsHtml = '';
  for (let i = 0; i < startOffset; i++) cellsHtml += '<div class="picker-day-cell picker-day-cell--empty"></div>';

  for (let day = 1; day <= daysInMonth; day++) {
    const currentDayStr = `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    const isSelected = state.selectedDates.includes(currentDayStr);
    const isDisabled = currentDayStr < todayStr;
    let cls = 'picker-day-cell';
    if (isSelected) cls += ' picker-day-cell--selected';
    if (isDisabled) cls += ' picker-day-cell--disabled';
    cellsHtml += `<div class="${cls}" data-date="${currentDayStr}">${day}</div>`;
  }

  container.innerHTML = cellsHtml;

  container.querySelectorAll('.picker-day-cell:not(.picker-day-cell--empty):not(.picker-day-cell--disabled)').forEach(el => {
    el.addEventListener('click', () => {
      const dateStr = el.dataset.date;
      const idx = state.selectedDates.indexOf(dateStr);
      if (idx > -1) state.selectedDates.splice(idx, 1);
      else { state.selectedDates.push(dateStr); state.selectedDates.sort(); }
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
    </span>`).join('');

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
  const isValid = state.selectedDates.length > 0 && catatan.length > 0 && managerEmail.length > 0 && !!state.tipeRequestDipilih;
  btnSubmit.disabled = !isValid;
}

async function submitWfaRequest() {
  const catatan = document.getElementById('req-catatan').value.trim();
  const managerEmail = state.karyawan.emailAtasan || '';
  const tipe = state.tipeRequestDipilih;

  if (!state.selectedDates.length) { showToast('Pilih minimal satu tanggal pengajuan.', 'warning'); return; }
  if (!catatan) { showToast('Alasan pengajuan wajib diisi.', 'warning'); return; }
  if (!tipe) { showToast('Pilih tipe pengajuan (WFA atau Visit).', 'warning'); return; }
  if (!managerEmail) { showToast('Atasan Anda belum dikonfigurasi. Hubungi admin.', 'warning'); return; }

  const btn = document.getElementById('btn-submit-request');
  setLoading(btn, true);

  try {
    await graphService.tambahPermohonanWfa({
      nip: state.karyawan.nip,
      nama: state.karyawan.nama,
      emailUser: state.user.mail || state.user.userPrincipalName,
      tanggalWfa: state.selectedDates.join(', '),
      tipe,
      emailAtasan: managerEmail,
      catatanUser: catatan
    });

    showToast(`✓ Pengajuan ${tipe} berhasil dikirim!`, 'success');
    state.selectedDates = [];
    document.getElementById('req-catatan').value = '';
    const month = state.pickerDate.getMonth() + 1;
    const year = state.pickerDate.getFullYear();
    renderPickerGrid(month, year);
    renderSelectedDates();

    // Redirect ke tab absen dan auto-pilih tipe
    showView('absen');
    await loadAbsen();
    const radioEl = document.getElementById(tipe === 'WFA' ? 'radio-wfa' : 'radio-visit');
    if (radioEl) { radioEl.checked = true; await onTipeAbsenChange(tipe); }

  } catch (err) {
    showToast('Gagal memproses pengajuan: ' + err.message, 'error');
  } finally {
    setLoading(btn, false, 'Kirim Pengajuan');
  }
}

// ============================================================
// VIEW APPROVAL ATASAN
// ============================================================
async function loadApprovalRequest(requestId, action) {
  showView('approval');
  const loadingState = document.getElementById('approval-loading-state');
  const boxContent = document.getElementById('approval-box-content');

  loadingState.classList.remove('hidden');
  boxContent.classList.add('hidden');

  try {
    const req = await graphService.getPermohonanWfaById(requestId);
    if (!req) throw new Error('Pengajuan tidak ditemukan.');
    state.currentApprovalRequest = req;

    document.getElementById('app-nama').textContent = req.nama;
    document.getElementById('app-nip').textContent = req.nip;
    document.getElementById('app-catatan-karyawan').textContent = req.catatanUser || '-';

    const tipeEl = document.getElementById('app-tipe');
    const tipe = req.tipe || 'WFA';
    tipeEl.textContent = tipe;
    tipeEl.className = `status-badge tipe-badge tipe-badge--${tipe.toLowerCase()}`;

    const datesEl = document.getElementById('app-tanggal-list');
    const dates = req.tanggalWfa ? req.tanggalWfa.split(',') : [];
    datesEl.innerHTML = dates.map(d => `<span class="selected-date-tag">${formatTanggalPendek(d.trim())}</span>`).join('');

    const btnConfirm = document.getElementById('btn-confirm-approval');
    const textarea = document.getElementById('app-catatan-atasan');
    const titleText = document.getElementById('approval-title-text');

    if (req.status === 'Rejected') {
      titleText.textContent = `Pengajuan ${tipe}: DIBATALKAN`;
      textarea.value = req.catatanAtasan || '';
      textarea.disabled = true;
      btnConfirm.disabled = true;
      btnConfirm.textContent = 'Sudah Ditolak / Dibatalkan';
      btnConfirm.classList.add('btn--disabled');
    } else {
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

// ============================================================
// KELOLA TIM — Modal manajemen anggota tim atasan
// ============================================================
let _timState = {
  list: [],          // [{id, nip, nama, email, jabatan, departemen, statusAktif}]
  editTarget: null,  // item sedang diedit
  hapusTarget: null  // item sedang dihapus
};

async function openTimModal() {
  const modal = document.getElementById('modal-tim');
  if (!modal) return;
  modal.classList.remove('hidden');
  modal.classList.add('modal--show');

  document.getElementById('tim-loading').classList.remove('hidden');
  document.getElementById('tim-content').classList.add('hidden');

  try {
    const myEmail = (state.karyawan.email || '').toLowerCase().trim();
    const allKaryawan = await graphService.getAllKaryawan();

    // Anggota tim = karyawan yang emailAtasan-nya adalah email user login
    _timState.list = allKaryawan.filter(k =>
      (k.emailAtasan || '').toLowerCase().trim() === myEmail
    );

    renderTimList();
    document.getElementById('tim-loading').classList.add('hidden');
    document.getElementById('tim-content').classList.remove('hidden');
  } catch (err) {
    showToast('Gagal memuat tim: ' + err.message, 'error');
    closeTimModal();
  }
}

function closeTimModal() {
  const modal = document.getElementById('modal-tim');
  if (modal) { modal.classList.remove('modal--show'); setTimeout(() => modal.classList.add('hidden'), 300); }
}

function renderTimList() {
  const container = document.getElementById('tim-list');
  if (!container) return;

  if (_timState.list.length === 0) {
    container.innerHTML = '<div class="tim-empty">Belum ada anggota tim yang terdaftar dengan Anda sebagai atasan.</div>';
    return;
  }

  container.innerHTML = _timState.list.map(k => {
    const initials = k.nama.split(' ').slice(0, 2).map(n => n[0]).join('').toUpperCase();
    const isAktif = k.statusAktif === 'Aktif';
    const avatarClass = isAktif ? 'tim-item__avatar' : 'tim-item__avatar tim-item__avatar--inactive';
    return `
      <div class="tim-item" data-id="${k.id}">
        <div class="${avatarClass}">${initials}</div>
        <div class="tim-item__info">
          <div class="tim-item__name">${k.nama}${!isAktif ? ' <span style="font-size:0.68rem;color:var(--text-muted);">(Nonaktif)</span>' : ''}</div>
          <div class="tim-item__meta">${k.jabatan || '-'} · ${k.departemen || '-'}</div>
          <div class="tim-item__meta" style="font-family:var(--font-mono);font-size:0.68rem;">${k.email}</div>
        </div>
        <div class="tim-item__actions">
          <button class="btn-tim-edit" data-id="${k.id}" title="Edit">✏️</button>
          <button class="btn-tim-hapus" data-id="${k.id}" title="Hapus">🗑️</button>
        </div>
      </div>`;
  }).join('');

  container.querySelectorAll('.btn-tim-edit').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const item = _timState.list.find(k => String(k.id) === String(btn.dataset.id));
      if (item) openFormAnggota('edit', item);
    });
  });

  container.querySelectorAll('.btn-tim-hapus').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const item = _timState.list.find(k => String(k.id) === String(btn.dataset.id));
      if (item) openHapusAnggota(item);
    });
  });
}

function openFormAnggota(mode, item = null) {
  _timState.editTarget = item;
  const modal = document.getElementById('modal-form-anggota');
  const title = document.getElementById('form-anggota-title');
  const errEl = document.getElementById('form-anggota-error');

  title.textContent = mode === 'edit' ? 'Edit Anggota Tim' : 'Tambah Anggota Tim';
  errEl.classList.add('hidden');

  document.getElementById('form-nip').value = item?.nip || '';
  document.getElementById('form-nama').value = item?.nama || '';
  document.getElementById('form-email').value = item?.email || '';
  document.getElementById('form-jabatan').value = item?.jabatan || '';
  document.getElementById('form-departemen').value = item?.departemen || '';
  document.getElementById('form-status').value = item?.statusAktif || 'Aktif';

  // NIP & email tidak bisa diubah saat edit (identitas)
  document.getElementById('form-nip').disabled = mode === 'edit';
  document.getElementById('form-email').disabled = mode === 'edit';

  modal.classList.remove('hidden');
  modal.classList.add('modal--show');
}

function closeFormAnggota() {
  const modal = document.getElementById('modal-form-anggota');
  if (modal) { modal.classList.remove('modal--show'); setTimeout(() => modal.classList.add('hidden'), 300); }
  _timState.editTarget = null;
}

async function simpanAnggota() {
  const btn = document.getElementById('btn-simpan-anggota');
  const errEl = document.getElementById('form-anggota-error');
  errEl.classList.add('hidden');

  const nip = document.getElementById('form-nip').value.trim();
  const nama = document.getElementById('form-nama').value.trim();
  const email = document.getElementById('form-email').value.trim().toLowerCase();
  const jabatan = document.getElementById('form-jabatan').value.trim();
  const departemen = document.getElementById('form-departemen').value.trim();
  const statusAktif = document.getElementById('form-status').value;

  if (!nip || !nama || !email) {
    errEl.textContent = 'NIP, Nama, dan Email wajib diisi.';
    errEl.classList.remove('hidden');
    return;
  }

  if (!email.includes('@')) {
    errEl.textContent = 'Format email tidak valid.';
    errEl.classList.remove('hidden');
    return;
  }

  setLoading(btn, true);

  try {
    const myEmail = state.karyawan.email || state.user.mail || state.user.userPrincipalName;

    if (_timState.editTarget) {
      // MODE EDIT
      await graphService.updateKaryawan(_timState.editTarget.id, {
        nama, jabatan, departemen, statusAktif,
        emailAtasan: myEmail
      });
      showToast(`✓ Data ${nama} berhasil diperbarui.`, 'success');
    } else {
      // MODE TAMBAH — cek duplikasi NIP/email dulu
      const existing = await graphService.getAllKaryawan();
      const dupNip = existing.find(k => String(k.nip) === String(nip));
      const dupEmail = existing.find(k => k.email?.toLowerCase() === email);

      if (dupNip) {
        errEl.textContent = `NIP ${nip} sudah terdaftar (${dupNip.nama}).`;
        errEl.classList.remove('hidden');
        setLoading(btn, false, 'Simpan');
        return;
      }
      if (dupEmail) {
        errEl.textContent = `Email ${email} sudah terdaftar (${dupEmail.nama}).`;
        errEl.classList.remove('hidden');
        setLoading(btn, false, 'Simpan');
        return;
      }

      await graphService.tambahKaryawan({
        nip, nama, email, jabatan, departemen,
        statusAktif,
        emailAtasan: myEmail
      });
      showToast(`✓ ${nama} berhasil ditambahkan ke tim.`, 'success');
    }

    closeFormAnggota();
    await openTimModal(); // refresh list

  } catch (err) {
    errEl.textContent = 'Gagal menyimpan: ' + err.message;
    errEl.classList.remove('hidden');
  } finally {
    setLoading(btn, false, 'Simpan');
  }
}

function openHapusAnggota(item) {
  _timState.hapusTarget = item;
  document.getElementById('hapus-nama-target').textContent = item.nama;
  const modal = document.getElementById('modal-hapus-anggota');
  modal.classList.remove('hidden');
  modal.classList.add('modal--show');
}

function closeHapusAnggota() {
  const modal = document.getElementById('modal-hapus-anggota');
  if (modal) { modal.classList.remove('modal--show'); setTimeout(() => modal.classList.add('hidden'), 300); }
  _timState.hapusTarget = null;
}

async function konfirmasiHapusAnggota() {
  const item = _timState.hapusTarget;
  if (!item) return;

  const btn = document.getElementById('btn-konfirmasi-hapus');
  setLoading(btn, true);

  try {
    await graphService.updateKaryawan(item.id, { statusAktif: 'Tidak Aktif' });
    showToast(`${item.nama} dinonaktifkan dari daftar karyawan.`, 'success');
    closeHapusAnggota();
    await openTimModal(); // refresh
  } catch (err) {
    showToast('Gagal menghapus: ' + err.message, 'error');
  } finally {
    setLoading(btn, false, 'Ya, Hapus');
  }
}

function bindEvents() {
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

  // Dropdown toggle — klik avatar buka/tutup menu
  document.getElementById('topbar-avatar')?.addEventListener('click', (e) => {
    e.stopPropagation();
    document.getElementById('user-dropdown')?.classList.toggle('hidden');
  });

  // Tutup dropdown saat klik di luar
  document.addEventListener('click', () => {
    document.getElementById('user-dropdown')?.classList.add('hidden');
  });

  document.getElementById('btn-menu-logout')?.addEventListener('click', async () => {
    await authService.logout();
    window.location.reload();
  });

  document.getElementById('btn-menu-tim')?.addEventListener('click', () => {
    document.getElementById('user-dropdown')?.classList.add('hidden');
    openTimModal();
  });

  document.querySelectorAll('[data-nav]').forEach(el => {
    el.addEventListener('click', async () => {
      const view = el.dataset.nav;
      showView(view);
      if (view === 'calendar') await loadCalendar();
      if (view === 'absen') await loadAbsen();
      if (view === 'request') loadRequestView();
    });
  });

  document.getElementById('btn-cal-prev')?.addEventListener('click', () => loadCalendar(-1));
  document.getElementById('btn-cal-next')?.addEventListener('click', () => loadCalendar(1));
  // NO.8 — Tombol Hari Ini
  document.getElementById('btn-cal-today')?.addEventListener('click', goToToday);

  document.getElementById('btn-picker-prev')?.addEventListener('click', () => loadRequestView(-1));
  document.getElementById('btn-picker-next')?.addEventListener('click', () => loadRequestView(1));

  document.querySelectorAll('input[name="tipe-absen"]').forEach(radio => {
    radio.addEventListener('change', async (e) => { await onTipeAbsenChange(e.target.value); });
  });

  document.querySelectorAll('input[name="tipe-request"]').forEach(radio => {
    radio.addEventListener('change', (e) => { state.tipeRequestDipilih = e.target.value; validateRequestForm(); });
  });

  document.getElementById('req-catatan')?.addEventListener('input', validateRequestForm);
  document.getElementById('btn-submit-request')?.addEventListener('click', submitWfaRequest);

  document.getElementById('btn-absen-masuk')?.addEventListener('click', () => openAbsenModal('masuk'));
  document.getElementById('btn-absen-keluar')?.addEventListener('click', () => openAbsenModal('keluar'));
  document.getElementById('btn-close-modal')?.addEventListener('click', closeAbsenModal);
  document.getElementById('btn-submit-absen')?.addEventListener('click', submitAbsen);
  document.getElementById('modal-absen')?.addEventListener('click', (e) => { if (e.target === e.currentTarget) closeAbsenModal(); });

  document.getElementById('btn-belum-request-tutup')?.addEventListener('click', closeBelumRequestModal);
  document.getElementById('btn-belum-request-ke-request')?.addEventListener('click', () => {
    closeBelumRequestModal();
    showView('request');
    loadRequestView();
    if (state.tipeAbsenDipilih && state.tipeAbsenDipilih !== 'WFO') {
      state.tipeRequestDipilih = state.tipeAbsenDipilih;
      const radioEl = document.getElementById(state.tipeAbsenDipilih === 'WFA' ? 'radio-req-wfa' : 'radio-req-visit');
      if (radioEl) radioEl.checked = true;
    }
  });
  document.getElementById('modal-belum-request')?.addEventListener('click', (e) => { if (e.target === e.currentTarget) closeBelumRequestModal(); });

  document.getElementById('calendar-days-container')?.addEventListener('click', (e) => {
    const cell = e.target.closest('.calendar-day-cell');
    if (cell && cell.dataset.date) openCalendarDetailModal(cell.dataset.date);
  });
  document.getElementById('btn-close-calendar-modal')?.addEventListener('click', closeCalendarDetailModal);
  document.getElementById('modal-calendar-detail')?.addEventListener('click', (e) => { if (e.target === e.currentTarget) closeCalendarDetailModal(); });

  // NO.7 — Close rejected notif
  document.getElementById('btn-close-rejected-notif')?.addEventListener('click', closeRejectedNotif);
  document.getElementById('btn-close-rejected-notif-ok')?.addEventListener('click', closeRejectedNotif);
  document.getElementById('modal-rejected-notif')?.addEventListener('click', (e) => { if (e.target === e.currentTarget) closeRejectedNotif(); });

  document.getElementById('btn-confirm-approval')?.addEventListener('click', confirmApproval);
  document.getElementById('btn-cancel-approval')?.addEventListener('click', cancelApproval);

  // ---- KELOLA TIM ----
  document.getElementById('btn-close-tim')?.addEventListener('click', closeTimModal);
  document.getElementById('modal-tim')?.addEventListener('click', (e) => { if (e.target === e.currentTarget) closeTimModal(); });

  document.getElementById('btn-tambah-anggota')?.addEventListener('click', () => openFormAnggota('tambah'));

  document.getElementById('btn-close-form-anggota')?.addEventListener('click', closeFormAnggota);
  document.getElementById('modal-form-anggota')?.addEventListener('click', (e) => { if (e.target === e.currentTarget) closeFormAnggota(); });
  document.getElementById('btn-simpan-anggota')?.addEventListener('click', simpanAnggota);

  document.getElementById('btn-batal-hapus')?.addEventListener('click', closeHapusAnggota);
  document.getElementById('modal-hapus-anggota')?.addEventListener('click', (e) => { if (e.target === e.currentTarget) closeHapusAnggota(); });
  document.getElementById('btn-konfirmasi-hapus')?.addEventListener('click', konfirmasiHapusAnggota);
}

// ============================================================
// ENTRY POINT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
  initApp();
});
