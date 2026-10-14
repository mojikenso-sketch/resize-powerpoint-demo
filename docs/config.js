/* ตั้งค่าเวอร์ชันทดลอง (DEMO) — แก้ข้อความติดต่อ/ราคา/ลิงก์ Trial Gate ได้ที่ไฟล์นี้ไฟล์เดียว */
window.DEMO_CONTACT = 'สนใจซื้อสิทธิ์ใช้งาน ติดต่อ: Fastwork / LINE: (ใส่ภายหลัง)';
window.DEMO_PRICE = 'สิทธิ์ใช้งานรายปี 990 บาท/ผู้ใช้/ปี · ขายขาด 1,990 บาท/ผู้ใช้';

/* ทดลองใช้ฟรี 5 ชั่วโมง 1 ครั้งต่อคน (นับตาม IP + อุปกรณ์) — ใช้ร่วมกับ js/trial-gate.js
   gateUrl = URL ของ Web App "Trial Gate" (Google Apps Script) ที่ลงท้ายด้วย /exec
   ถ้ายังเป็น PASTE_... ระบบจะไม่จำกัดเวลา (เปิดใช้ได้ตามปกติ) */
window.DEMO_TRIAL = {
  gateUrl: 'https://script.google.com/macros/s/AKfycbwQzhsI1OIN-YUumWLE0n1150c9hTPC--Npm19k44AIDjneJfn2SFXnTYxsW9OGbO0/exec',
  app: 'resize-ppt',
  hours: 5,
  contact: window.DEMO_CONTACT,
  priceHtml: 'สิทธิ์ใช้งานรายปี 990 บาท/ผู้ใช้/ปี · ขายขาด 1,990 บาท/ผู้ใช้',
};
