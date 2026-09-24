// ============================================================
// UTILS.JS - Utility: Kamera, Geolokasi, Helper
// ============================================================

// ============================================================
// HELPER FUNCTIONS
// ============================================================

function formatTanggal(dateStr) {
  if (!dateStr) return '-';
  const d = new Date(dateStr);
  return d.toLocaleDateString('id-ID', {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
  });
}

function formatJam(timeStr) {
  if (!timeStr) return '-';
  return String(timeStr).substring(0, 5);
}

function formatTanggalPendek(dateStr) {
  if (!dateStr) return '-';
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('id-ID', {
    day: '2-digit', month: 'short', year: 'numeric'
  });
}

function hitungDurasi(jamMasuk, jamKeluar) {
  if (!jamMasuk || !jamKeluar) return '-';
  const [hm, mm] = jamMasuk.split(':').map(Number);
  const [hk, mk] = jamKeluar.split(':').map(Number);
  const diffMenit = (hk * 60 + mk) - (hm * 60 + mm);
  if (diffMenit < 0) return '-';
  const jam = Math.floor(diffMenit / 60);
  const menit = diffMenit % 60;
  return `${jam}j ${menit}m`;
}

function showToast(message, type = 'info', duration = 3500) {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast--${type}`;
  
  const icons = {
    success: '✓',
    error: '✕',
    warning: '⚠',
    info: 'ℹ',
  };

  toast.innerHTML = `
    <span class="toast__icon">${icons[type] || icons.info}</span>
    <span class="toast__msg">${message}</span>
  `;

  container.appendChild(toast);
  
  // Animate in
  requestAnimationFrame(() => toast.classList.add('toast--show'));
  
  setTimeout(() => {
    toast.classList.remove('toast--show');
    toast.addEventListener('transitionend', () => toast.remove(), { once: true });
  }, duration);
}

function setLoading(element, loading, originalText = '') {
  if (!element) return;
  if (loading) {
    element.disabled = true;
    element.dataset.originalText = element.textContent;
    element.innerHTML = '<span class="spinner"></span>';
  } else {
    element.disabled = false;
    element.textContent = originalText || element.dataset.originalText || 'Submit';
  }
}

function getTodayString() {
  return new Date().toLocaleDateString('en-CA');
}

export {
  formatTanggal, formatJam, formatTanggalPendek,
  hitungDurasi, showToast, setLoading, getTodayString
};
