/* ระบบทดลองใช้ฟรี 3 วัน (ต่อเบราว์เซอร์) — เก็บเวลาเริ่มทดลองไว้ใน localStorage เท่านั้น
   ไม่มีการส่งข้อมูลหรือไฟล์ใดๆ ออกจากเครื่อง */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.DemoTrialCore = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const KEY_START = 'demo_rpp_trial_start';
  const HOUR = 3600 * 1000;
  const DAY = 24 * HOUR;

  /**
   * คืนสถานะทดลองใช้: { allowed, expired, remainingMs, days, hours, storage }
   * storage = false เมื่อใช้ localStorage ไม่ได้ (อนุญาตให้ใช้งานได้ตามปกติ)
   */
  function getStatus(storage, now, trialDays) {
    const total = (trialDays || 3) * DAY;
    let start;
    try {
      const raw = storage.getItem(KEY_START);
      start = raw === null ? NaN : Number(raw);
      // ไม่มีค่า / ค่าเสีย / ค่าอยู่ในอนาคต → เริ่มนับใหม่จากตอนนี้
      if (!Number.isFinite(start) || start <= 0 || start > now) {
        start = now;
        storage.setItem(KEY_START, String(start));
      }
    } catch (_error) {
      return { allowed: true, expired: false, remainingMs: total, days: 0, hours: 0, storage: false };
    }
    const remainingMs = Math.max(0, start + total - now);
    const expired = remainingMs <= 0;
    return {
      allowed: !expired,
      expired,
      remainingMs,
      days: Math.floor(remainingMs / DAY),
      hours: Math.floor((remainingMs % DAY) / HOUR),
      storage: true,
    };
  }

  function getStorage() {
    try {
      return window.localStorage;
    } catch (_error) {
      return null;
    }
  }

  return { KEY_START, DAY, HOUR, getStatus, getStorage };
});

(function () {
  'use strict';
  if (typeof window === 'undefined' || typeof document === 'undefined') return;

  const core = window.DemoTrialCore;
  const days = Number(window.DEMO_TRIAL_DAYS) || 3;
  const contact = window.DEMO_CONTACT || '';
  const price = window.DEMO_PRICE || '';
  const nullStorage = { getItem() { throw new Error('no storage'); }, setItem() { throw new Error('no storage'); } };

  function status() {
    return core.getStatus(core.getStorage() || nullStorage, Date.now(), days);
  }

  let banner;
  let modal;

  function renderBanner() {
    const s = status();
    if (!banner) {
      banner = document.createElement('div');
      banner.className = 'demo-banner';
      banner.setAttribute('role', 'note');
      document.body.insertBefore(banner, document.body.firstChild);
    }
    const left = s.expired
      ? `🧪 เวอร์ชันทดลองใช้ฟรี ${days} วัน · หมดระยะทดลองแล้ว`
      : `🧪 เวอร์ชันทดลองใช้ฟรี ${days} วัน · เหลืออีก ${s.days} วัน ${s.hours} ชม.`;
    banner.classList.toggle('is-expired', s.expired);
    banner.innerHTML = '';
    const a = document.createElement('span');
    a.className = 'demo-banner-main';
    a.textContent = left;
    const b = document.createElement('span');
    b.className = 'demo-banner-contact';
    b.textContent = contact;
    banner.append(a, b);
  }

  function showExpired() {
    if (!modal) {
      modal = document.createElement('dialog');
      modal.className = 'dialog demo-expired';
      modal.innerHTML =
        '<form method="dialog" class="dialog-card">' +
        '<div class="demo-expired-icon" aria-hidden="true">⏰</div>' +
        '<h3></h3><p class="demo-expired-body"></p>' +
        '<div class="demo-price"></div><p class="demo-contact"></p>' +
        '<div class="dialog-actions"><button class="btn btn-primary" value="ok">ตกลง</button></div>' +
        '</form>';
      modal.querySelector('h3').textContent = `หมดระยะทดลองใช้ ${days} วันแล้ว`;
      modal.querySelector('.demo-expired-body').textContent =
        'ขอบคุณที่ทดลองใช้งานโปรแกรมย่อไฟล์ PowerPoint หากต้องการใช้งานต่อ สามารถซื้อสิทธิ์ใช้งานได้ตามรายละเอียดด้านล่าง';
      modal.querySelector('.demo-price').textContent = price;
      modal.querySelector('.demo-contact').textContent = contact;
      modal.addEventListener('click', (event) => {
        if (event.target === modal) modal.close();
      });
      document.body.appendChild(modal);
    }
    if (typeof modal.showModal === 'function') {
      if (!modal.open) modal.showModal();
    } else {
      window.alert(`หมดระยะทดลองใช้ ${days} วันแล้ว\n\n${price}\n${contact}`);
    }
  }

  /** เรียกก่อนบีบอัด/ดาวน์โหลด — คืน false (และแสดงหน้าต่างแจ้ง) เมื่อหมดระยะทดลอง */
  function allow() {
    const s = status();
    renderBanner();
    if (s.allowed) return true;
    showExpired();
    return false;
  }

  window.DemoTrial = { allow, status };

  function init() {
    renderBanner();
    setInterval(renderBanner, 60 * 1000);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
