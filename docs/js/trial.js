/* ข้อจำกัดเวอร์ชันทดลอง (DEMO) + แถบแจ้งด้านบน
   - เวลาทดลอง 5 ชั่วโมง (นับตาม IP + อุปกรณ์) จัดการโดย js/trial-gate.js → window.DemoTrial
   - ไฟล์นี้: ระดับความชัดที่เปิดให้ลอง, ขนาดไฟล์สูงสุด, แถบแจ้ง และกล่องแจ้ง 🔒 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.DemoLimits = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /** ระดับที่ใช้ได้ในเวอร์ชันทดลอง — ระดับอื่นแสดงไว้แต่ล็อก 🔒 */
  const ALLOWED_KEYS = Object.freeze(['balanced', 'smallest']);
  const MAX_FILE_MB = 30;
  const MAX_FILE_BYTES = MAX_FILE_MB * 1024 * 1024;

  function isLocked(key) {
    return ALLOWED_KEYS.indexOf(key) === -1;
  }

  /** true เมื่อไฟล์ใหญ่เกินที่เวอร์ชันทดลองรองรับ */
  function isTooLarge(file) {
    return Boolean(file) && Number(file.size) > MAX_FILE_BYTES;
  }

  /** true เมื่อหมดเวลาทดลอง (trial-gate.js แสดงหน้าจอเต็มให้แล้ว — ผู้เรียกแค่หยุดการทำงาน) */
  function trialExpired(win) {
    const gate = win && win.DemoTrial;
    return Boolean(gate && typeof gate.isExpired === 'function' && gate.isExpired());
  }

  function lockedMessage(contact) {
    return `🔒 ระดับนี้มีในเวอร์ชันเต็ม — ${contact || ''}`.trim();
  }

  function tooLargeMessage() {
    return `เวอร์ชันทดลองรองรับไฟล์ไม่เกิน ${MAX_FILE_MB} MB ต่อไฟล์\nเวอร์ชันเต็มไม่จำกัดขนาดไฟล์`;
  }

  return { ALLOWED_KEYS, MAX_FILE_MB, MAX_FILE_BYTES, isLocked, isTooLarge, trialExpired, lockedMessage, tooLargeMessage };
});

(function () {
  'use strict';
  if (typeof window === 'undefined' || typeof document === 'undefined') return;

  const D = window.DemoLimits;
  const cfg = window.DEMO_TRIAL || {};
  const hours = Number(cfg.hours) || 5;
  const contact = cfg.contact || window.DEMO_CONTACT || '';
  let toast;
  let toastTimer;

  /** กล่องแจ้งเล็กๆ เมื่อกดระดับที่ล็อก */
  D.showLocked = function showLocked() {
    if (!toast) {
      toast = document.createElement('div');
      toast.className = 'demo-toast';
      toast.setAttribute('role', 'status');
      toast.setAttribute('aria-live', 'polite');
      document.body.appendChild(toast);
    }
    toast.textContent = D.lockedMessage(contact);
    toast.classList.add('is-show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => toast.classList.remove('is-show'), 4000);
  };

  function renderBanner() {
    const banner = document.createElement('div');
    banner.className = 'demo-banner';
    banner.setAttribute('role', 'note');
    const a = document.createElement('span');
    a.className = 'demo-banner-main';
    a.textContent = `🧪 เวอร์ชันทดลองใช้ฟรี ${hours} ชั่วโมง (1 ครั้งต่อคน)`;
    const b = document.createElement('span');
    b.className = 'demo-banner-contact';
    b.textContent = contact;
    banner.append(a, b);
    document.body.insertBefore(banner, document.body.firstChild);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', renderBanner);
  else renderBanner();
})();
