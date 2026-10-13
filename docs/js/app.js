(function () {
  'use strict';

  const R = window.PptxResizer;
  // ต้องตรงกับ version.txt บนเซิร์ฟเวอร์ — ถ้าไม่ตรง แปลว่าแท็บนี้ยังเป็นโค้ดเก่า
  const BUILD = '6';
  // เวอร์ชัน DEMO (GitHub Pages): ไม่มี usage.php — นับสถิติเฉพาะในเบราว์เซอร์นี้ด้วย localStorage
  const USAGE_STORAGE_KEY = 'demo_rpp_usage';
  const VISIT_SESSION_KEY = 'resize-powerpoint-visit-counted-v1';
  const $ = (id) => document.getElementById(id);
  const reducedMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  const PRESET_STYLE = {
    smallest: { accent: '#10B981', icon: 'i-send' },
    compact: { accent: '#0891B2', icon: 'i-laptop' },
    balanced: { accent: '#3B63FB', icon: 'i-scale', badge: 'แนะนำ' },
    clear: { accent: '#8B5CF6', icon: 'i-monitor' },
    high: { accent: '#DB2777', icon: 'i-gem' },
  };
  const RING_LENGTH = 163.36;

  const els = {
    dropZone: $('dropZone'),
    fileInput: $('fileInput'),
    pickButton: $('pickButton'),
    fileName: $('fileName'),
    fileMeta: $('fileMeta'),
    qualityGrid: $('qualityGrid'),
    helperTile: $('helperTile'),
    qualityHint: $('qualityHint'),
    qualitySummary: $('qualitySummary'),
    outputName: $('outputName'),
    originalSize: $('originalSize'),
    estimatedSize: $('estimatedSize'),
    reduction: $('reduction'),
    ringFill: $('ringFill'),
    details: $('details'),
    progress: $('progress'),
    progressBar: $('progressBar'),
    status: $('status'),
    statusWrap: $('statusWrap'),
    resetButton: $('resetButton'),
    downloadButton: $('downloadButton'),
    dialog: $('messageDialog'),
    dialogTitle: $('dialogTitle'),
    dialogBody: $('dialogBody'),
    confetti: $('confetti'),
    usageStats: $('usageStats'),
    visitCount: $('visitCount'),
    downloadCount: $('downloadCount'),
  };

  const state = {
    file: null,
    token: 0,
    pickToken: 0,
    selectedKey: null,
    results: {},
    analyzing: false,
    failed: false,
    downloaded: false,
  };

  const cards = {};

  // ---------------------------------------------------------------------------
  // Small helpers
  // ---------------------------------------------------------------------------

  function svgIcon(id, className = 'ico') {
    return `<svg class="${className}" aria-hidden="true"><use href="#${id}"/></svg>`;
  }

  function renderUsageCounts(counts) {
    const visits = Math.max(0, Math.trunc(Number(counts.visits) || 0));
    const downloads = Math.max(0, Math.trunc(Number(counts.downloads) || 0));
    els.visitCount.textContent = visits.toLocaleString('th-TH');
    els.downloadCount.textContent = downloads.toLocaleString('th-TH');
    els.usageStats.dataset.state = 'ready';
  }

  async function requestUsage(eventName = null) {
    let counts = { visits: 0, downloads: 0 };
    try {
      const stored = JSON.parse(localStorage.getItem(USAGE_STORAGE_KEY) || 'null');
      if (stored && typeof stored === 'object') {
        counts = {
          visits: Math.max(0, Math.trunc(Number(stored.visits) || 0)),
          downloads: Math.max(0, Math.trunc(Number(stored.downloads) || 0)),
        };
      }
      if (eventName === 'visit') counts.visits += 1;
      else if (eventName === 'download') counts.downloads += 1;
      if (eventName) localStorage.setItem(USAGE_STORAGE_KEY, JSON.stringify(counts));
    } catch (_error) {
      throw new Error('usage counter storage unavailable');
    }
    return counts;
  }

  async function initializeUsageCounter() {
    let shouldCountVisit = true;
    try {
      shouldCountVisit = sessionStorage.getItem(VISIT_SESSION_KEY) !== 'done';
      if (shouldCountVisit) sessionStorage.setItem(VISIT_SESSION_KEY, 'pending');
    } catch (_error) {
      // บางเบราว์เซอร์ปิด sessionStorage; ยังนับการเปิดหน้านี้ได้ตามปกติ
    }

    try {
      const counts = await requestUsage(shouldCountVisit ? 'visit' : null);
      renderUsageCounts(counts);
      if (shouldCountVisit) {
        try {
          sessionStorage.setItem(VISIT_SESSION_KEY, 'done');
        } catch (_error) {
          // แสดงตัวเลขต่อได้แม้ sessionStorage ใช้งานไม่ได้
        }
      }
    } catch (_error) {
      els.usageStats.dataset.state = 'unavailable';
      els.usageStats.title = 'ไม่สามารถโหลดสถิติการใช้งานได้ในขณะนี้';
      if (shouldCountVisit) {
        try {
          sessionStorage.removeItem(VISIT_SESSION_KEY);
        } catch (_storageError) {
          // ไม่มีข้อมูลส่วนบุคคลหรือข้อมูลไฟล์ถูกเก็บในตัวนับนี้
        }
      }
    }
  }

  function recordDownloadUsage() {
    requestUsage('download')
      .then(renderUsageCounts)
      .catch(() => {
        els.usageStats.dataset.state = 'unavailable';
        els.usageStats.title = 'ดาวน์โหลดสำเร็จ แต่บันทึกสถิติไม่สำเร็จ';
      });
  }

  /** นับตัวเลขขึ้น/ลงแบบนุ่มนวล — ค่าสุดท้ายตรงกับค่าจริงเสมอ */
  function animateNumber(el, to, format) {
    const from = typeof el._value === 'number' ? el._value : null;
    el._value = to;
    if (el._raf) cancelAnimationFrame(el._raf);

    if (from === null || from === to || reducedMotion) {
      el.textContent = format(to);
      return;
    }

    const start = performance.now();
    const duration = 650;
    const step = (now) => {
      const t = Math.min(1, (now - start) / duration);
      const eased = 1 - Math.pow(1 - t, 3);
      el.textContent = format(from + (to - from) * eased);
      if (t < 1) el._raf = requestAnimationFrame(step);
      else {
        el._raf = null;
        el.textContent = format(to);
      }
    };
    el._raf = requestAnimationFrame(step);
  }

  function setText(el, text) {
    if (el._raf) cancelAnimationFrame(el._raf);
    el._raf = null;
    el._value = null;
    el.textContent = text;
  }

  function setRing(percent) {
    const value = Math.max(0, Math.min(100, percent || 0));
    els.ringFill.style.strokeDashoffset = String(RING_LENGTH * (1 - value / 100));
  }

  // ---------------------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------------------

  function buildCards() {
    R.QUALITY_PRESETS.forEach((preset, index) => {
      const style = PRESET_STYLE[preset.key] || { accent: '#3B63FB', icon: 'i-images' };
      const card = document.createElement('button');
      card.type = 'button';
      card.className = 'quality-card';
      card.setAttribute('role', 'radio');
      card.setAttribute('aria-checked', 'false');
      card.dataset.key = preset.key;
      card.disabled = true;
      card.style.setProperty('--accent', style.accent);

      const meter = Array.from({ length: R.QUALITY_PRESETS.length }, (_v, i) => `<i class="${i <= index ? 'on' : ''}"></i>`).join('');
      card.innerHTML = `
        <span class="q-top">
          <span class="q-icon">${svgIcon(style.icon)}</span>
          <span class="q-heading">
            <span class="q-title"><span class="q-name"></span>${style.badge ? `<span class="q-badge">${style.badge}</span>` : ''}</span>
            <span class="q-spec"></span>
          </span>
          <span class="q-check">${svgIcon('i-check')}</span>
        </span>
        <span class="q-use"></span>
        <span class="q-meter" aria-hidden="true"><span>ความชัด</span>${meter}</span>
        <span class="q-result">
          <span class="q-estimate">เลือกไฟล์ก่อน</span>
          <span class="q-bar" aria-hidden="true"><i></i></span>
        </span>`;
      card.querySelector('.q-name').textContent = preset.title;
      card.querySelector('.q-spec').textContent = R.specification(preset);
      card.querySelector('.q-use').textContent = preset.useCase;
      card.title = preset.description;
      card.addEventListener('click', () => selectPreset(preset.key));
      card.addEventListener('keydown', onCardKeydown);
      els.qualityGrid.insertBefore(card, els.helperTile);
      cards[preset.key] = card;
    });
  }

  function onCardKeydown(event) {
    const keys = R.QUALITY_PRESETS.map((p) => p.key);
    const index = keys.indexOf(event.currentTarget.dataset.key);
    let next = null;
    if (event.key === 'ArrowRight' || event.key === 'ArrowDown') next = keys[(index + 1) % keys.length];
    if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') next = keys[(index - 1 + keys.length) % keys.length];
    if (next && !cards[next].disabled) {
      event.preventDefault();
      selectPreset(next);
      cards[next].focus();
    }
  }

  /** ป้ายเปอร์เซ็นต์บนการ์ด — ไม่แสดง "-0.0%" เมื่อไฟล์ไม่เล็กลง */
  function reductionLabel(result) {
    if (result.resizedSize >= result.originalSize) return { text: 'ไม่เล็กลง', warn: true };
    const pct = result.reductionPercent;
    return { text: pct < 0.1 ? '-<0.1%' : `-${pct.toFixed(1)}%`, warn: false };
  }

  /** kind: '' | 'loading' | 'ready' | 'error'; label: { text, warn } */
  function setCardEstimate(key, text, kind, label) {
    const card = cards[key];
    const el = card.querySelector('.q-estimate');
    el.className = `q-estimate${kind ? ` ${kind}` : ''}`;
    el.textContent = text;
    if (label) {
      const pct = document.createElement('span');
      pct.className = `pct${label.warn ? ' warn' : ''}`;
      pct.textContent = label.text;
      el.appendChild(pct);
    }

    const bar = card.querySelector('.q-bar');
    bar.classList.toggle('loading', kind === 'loading');
    if (kind !== 'ready') bar.style.setProperty('--w', '0%');
  }

  function setCardBar(key, fraction) {
    const bar = cards[key].querySelector('.q-bar');
    bar.classList.remove('loading');
    const width = Math.max(2, Math.min(100, fraction * 100));
    requestAnimationFrame(() => bar.style.setProperty('--w', `${width}%`));
  }

  function setCardsEnabled(enabled) {
    for (const card of Object.values(cards)) card.disabled = !enabled;
  }

  function refreshSelection() {
    for (const [key, card] of Object.entries(cards)) {
      const selected = key === state.selectedKey;
      card.setAttribute('aria-checked', String(selected));
      card.tabIndex = selected || (!state.selectedKey && key === R.DEFAULT_PRESET_KEY) ? 0 : -1;
    }
  }

  function setProgress(percent) {
    const value = Math.max(0, Math.min(100, percent));
    els.progressBar.style.width = `${value}%`;
    els.progress.setAttribute('aria-valuenow', String(Math.round(value)));
  }

  /** statusState: idle | busy | done | error */
  function setStatus(text, statusState) {
    els.status.textContent = text;
    if (statusState) els.statusWrap.dataset.state = statusState;
  }

  function updateSteps() {
    const ready = Boolean(state.file && state.selectedKey && state.results[state.selectedKey]);
    let current = 1;
    if (state.file) current = ready ? 3 : 2;
    if (ready && state.downloaded) current = 4;

    document.body.dataset.step = String(current);
    for (const n of [1, 2, 3]) {
      const done = n < current;
      const active = n === current;
      document.querySelectorAll(`[data-card="${n}"]`).forEach((el) => {
        el.classList.toggle('is-done', done);
        el.classList.toggle('is-active', active);
      });
      document.querySelectorAll(`.stp[data-for="${n}"]`).forEach((el) => {
        el.classList.toggle('is-done', done);
        el.classList.toggle('is-active', active);
      });
      document.querySelectorAll(`.stp-line[data-for="${n}"]`).forEach((el) => {
        el.classList.toggle('is-done', n <= current);
      });
    }
  }

  function renderDetails(result) {
    const items = [
      ['รูปทั้งหมด', result.totalImages],
      ['บีบอัด JPEG', result.compressedImages],
      ['คงรูปเดิม', result.skippedImages],
    ];

    els.details.textContent = '';
    items.forEach(([label, value], index) => {
      if (index) {
        const sep = document.createElement('span');
        sep.className = 'sep';
        sep.textContent = '  •  ';
        els.details.appendChild(sep);
      }
      const chip = document.createElement('span');
      chip.className = 'chip';
      chip.style.animationDelay = `${index * 0.05}s`;
      chip.append(`${label} `);
      const b = document.createElement('b');
      b.textContent = value.toLocaleString('en-US');
      chip.appendChild(b);
      els.details.appendChild(chip);
    });

    if (result.resizedSize >= result.originalSize) {
      const sep = document.createElement('span');
      sep.className = 'sep';
      sep.textContent = '  —  ';
      const chip = document.createElement('span');
      chip.className = 'chip warn';
      chip.textContent = 'ระดับนี้ไม่ทำให้ไฟล์เล็กลง';
      els.details.append(sep, chip);
    }
    els.details.hidden = false;
  }

  function updateSummary() {
    if (!state.file || !state.selectedKey) {
      els.qualitySummary.textContent = 'ยังไม่ได้เลือกระดับความชัด';
      setText(els.estimatedSize, '—');
      setText(els.reduction, '—');
      setRing(0);
      els.details.hidden = true;
      return;
    }

    const preset = R.PRESET_BY_KEY[state.selectedKey];
    els.qualitySummary.textContent =
      `${preset.title}  •  ด้านยาวสูงสุด ${preset.maxImageSide.toLocaleString('en-US')} px  •  JPEG Quality ${preset.jpegQuality}`;

    const result = state.results[state.selectedKey];
    if (!result) {
      const text = state.failed ? 'คำนวณไม่สำเร็จ' : 'กำลังคำนวณ';
      setText(els.estimatedSize, text);
      setText(els.reduction, text);
      setRing(0);
      els.details.hidden = true;
      return;
    }

    animateNumber(els.estimatedSize, result.resizedSize, (v) => R.formatBytes(Math.round(v)));
    animateNumber(els.reduction, result.reductionPercent, (v) => `${v.toFixed(1)}%`);
    setRing(result.reductionPercent);
    renderDetails(result);
  }

  function updateButtons() {
    els.downloadButton.disabled = !(state.file && state.selectedKey && state.results[state.selectedKey]);
    els.resetButton.disabled = !state.file;
    updateSteps();
  }

  function showDialog(title, body, kind) {
    els.dialogTitle.textContent = title;
    els.dialogTitle.className = kind || '';
    els.dialogBody.textContent = body;
    els.dialog.classList.remove('is-success', 'is-error', 'is-info');
    els.dialog.classList.add(`is-${kind || 'info'}`);

    if (typeof els.dialog.showModal === 'function') {
      if (els.dialog.open) els.dialog.close();
      // restart the entrance animations
      void els.dialog.offsetWidth;
      els.dialog.showModal();
      if (kind === 'success') launchConfetti();
    } else {
      window.alert(`${title}\n\n${body}`);
    }
  }

  // ---------------------------------------------------------------------------
  // Effects
  // ---------------------------------------------------------------------------

  function addRipple(event) {
    const target = event.currentTarget;
    if (target.disabled || reducedMotion) return;
    const rect = target.getBoundingClientRect();
    const size = Math.max(rect.width, rect.height) * 2.2;
    const ripple = document.createElement('span');
    ripple.className = 'ripple';
    ripple.style.width = ripple.style.height = `${size}px`;
    ripple.style.left = `${event.clientX - rect.left - size / 2}px`;
    ripple.style.top = `${event.clientY - rect.top - size / 2}px`;
    target.appendChild(ripple);
    ripple.addEventListener('animationend', () => ripple.remove(), { once: true });
  }

  let confettiRaf = null;
  function launchConfetti() {
    if (reducedMotion) return;
    const canvas = els.confetti;
    const ctx = canvas.getContext('2d');
    const dpr = Math.min(2, window.devicePixelRatio || 1);
    const width = window.innerWidth;
    const height = window.innerHeight;
    canvas.width = width * dpr;
    canvas.height = height * dpr;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    const colors = ['#E53935', '#F57C00', '#FACC15', '#16A34A', '#0284C7', '#4338CA', '#9333EA'];
    const pieces = [];
    const spawn = (x, y, angle) => {
      for (let i = 0; i < 70; i += 1) {
        const a = angle + (Math.random() - 0.5) * 1.1;
        const speed = 7 + Math.random() * 9;
        pieces.push({
          x, y,
          vx: Math.cos(a) * speed,
          vy: Math.sin(a) * speed,
          w: 6 + Math.random() * 6,
          h: 8 + Math.random() * 8,
          r: Math.random() * Math.PI,
          vr: (Math.random() - 0.5) * 0.35,
          color: colors[(Math.random() * colors.length) | 0],
          round: Math.random() < 0.3,
        });
      }
    };
    spawn(0, height * 0.65, -Math.PI / 3.2);
    spawn(width, height * 0.65, -Math.PI + Math.PI / 3.2);

    const start = performance.now();
    if (confettiRaf) cancelAnimationFrame(confettiRaf);
    const frame = (now) => {
      const elapsed = now - start;
      ctx.clearRect(0, 0, width, height);
      for (const p of pieces) {
        p.vy += 0.28;
        p.vx *= 0.985;
        p.vy *= 0.985;
        p.x += p.vx;
        p.y += p.vy;
        p.r += p.vr;
        ctx.save();
        ctx.globalAlpha = Math.max(0, 1 - elapsed / 2800);
        ctx.translate(p.x, p.y);
        ctx.rotate(p.r);
        ctx.fillStyle = p.color;
        if (p.round) {
          ctx.beginPath();
          ctx.arc(0, 0, p.w / 2, 0, Math.PI * 2);
          ctx.fill();
        } else {
          ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h * Math.abs(Math.cos(p.r * 2)) + 2);
        }
        ctx.restore();
      }
      if (elapsed < 2800) confettiRaf = requestAnimationFrame(frame);
      else {
        ctx.clearRect(0, 0, width, height);
        confettiRaf = null;
      }
    };
    confettiRaf = requestAnimationFrame(frame);
  }

  // ---------------------------------------------------------------------------
  // Actions
  // ---------------------------------------------------------------------------

  function selectPreset(key) {
    if (!state.file || !cards[key] || cards[key].disabled) return;
    if (state.selectedKey !== key) state.downloaded = false;
    state.selectedKey = key;
    refreshSelection();
    updateSummary();
    updateButtons();
  }

  function resetUi() {
    state.token += 1;
    state.pickToken += 1;
    state.file = null;
    state.selectedKey = null;
    state.results = {};
    state.analyzing = false;
    state.failed = false;
    state.downloaded = false;

    els.fileInput.value = '';
    els.dropZone.classList.remove('has-file');
    els.fileName.textContent = 'ยังไม่ได้เลือกไฟล์';
    els.fileName.removeAttribute('title');
    els.fileMeta.textContent = 'รองรับไฟล์ .pptx';
    els.outputName.textContent = 'เลือกไฟล์เพื่อแสดงชื่อไฟล์ Output';
    setText(els.originalSize, '—');
    els.qualityHint.textContent = 'การ์ดจะเปิดใช้งานหลังเลือกไฟล์ และแสดงขนาดปลายทางของแต่ละระดับ';
    for (const key of Object.keys(cards)) setCardEstimate(key, 'เลือกไฟล์ก่อน');
    setCardsEnabled(false);
    refreshSelection();
    setProgress(0);
    els.progress.classList.remove('is-active');
    setStatus('เริ่มต้นโดยเลือกไฟล์ PowerPoint', 'idle');
    updateSummary();
    updateButtons();
  }

  /** คืน true เมื่อบนเซิร์ฟเวอร์มีเวอร์ชันใหม่กว่าโค้ดที่แท็บนี้กำลังรัน */
  async function isOutdated() {
    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 2500);
      const response = await fetch(`version.txt?t=${Date.now()}`, { cache: 'no-store', signal: controller.signal });
      clearTimeout(timer);
      if (!response.ok) return false;
      const latest = (await response.text()).trim();
      return Boolean(latest) && latest !== BUILD;
    } catch (_error) {
      return false; // เซิร์ฟเวอร์ไม่ตอบ — ใช้งานต่อได้ตามปกติ
    }
  }

  function promptReload() {
    showDialog(
      'มีเวอร์ชันใหม่ของเว็บ',
      'หน้านี้เปิดค้างไว้ตั้งแต่ก่อนการอัปเดต\nกด "ตกลง" เพื่อโหลดหน้าใหม่ แล้วเลือกไฟล์อีกครั้ง',
      'info',
    );
    els.dialog.addEventListener('close', () => window.location.reload(), { once: true });
  }

  async function handleFile(file) {
    if (!file) return;
    if (window.DemoTrial && !window.DemoTrial.allow()) return; // DEMO: หมดระยะทดลอง

    if (await isOutdated()) {
      promptReload();
      return;
    }

    if (!/\.pptx$/i.test(file.name)) {
      showDialog('ไม่สามารถใช้ไฟล์นี้ได้', 'กรุณาเลือกไฟล์นามสกุล .pptx', 'error');
      return;
    }
    if (file.size === 0) {
      showDialog('ไม่สามารถใช้ไฟล์นี้ได้', 'ไฟล์ว่างเปล่า (0 ไบต์)', 'error');
      return;
    }

    // ตรวจไฟล์ใหม่ให้ผ่านก่อน แล้วค่อยยกเลิกงานเดิม — ไฟล์เสียจะไม่ทำให้งานที่กำลังคำนวณค้าง
    state.pickToken += 1;
    const pickToken = state.pickToken;
    const isStalePick = () => pickToken !== state.pickToken;
    const previousStatus = { text: els.status.textContent, state: els.statusWrap.dataset.state };
    const busyBefore = state.analyzing;

    let buffer;
    try {
      if (!busyBefore) setStatus('กำลังอ่านไฟล์...', 'busy');
      buffer = await file.arrayBuffer();
      if (isStalePick()) return;
      await R.loadPresentation(buffer);
      if (isStalePick()) return;
    } catch (error) {
      if (isStalePick()) return;
      if (!busyBefore) {
        if (state.file) setStatus(previousStatus.text, previousStatus.state);
        else setStatus('เริ่มต้นโดยเลือกไฟล์ PowerPoint', 'idle');
      }
      showDialog('ไม่สามารถใช้ไฟล์นี้ได้', error && error.message ? error.message : String(error), 'error');
      return;
    }

    state.token += 1;
    const token = state.token;
    const isCancelled = () => token !== state.token;

    state.file = file;
    state.results = {};
    state.failed = false;
    state.analyzing = true;
    state.downloaded = false;
    state.selectedKey = R.DEFAULT_PRESET_KEY;

    els.dropZone.classList.remove('has-file');
    void els.dropZone.offsetWidth; // replay the file-in animation
    els.dropZone.classList.add('has-file');
    els.fileName.textContent = file.name;
    els.fileName.title = file.name;
    els.fileMeta.textContent = `${R.formatBytes(file.size)}   •   แก้ไขล่าสุด ${new Date(file.lastModified).toLocaleString('th-TH')}`;
    els.outputName.textContent = R.makeOutputName(file.name);
    setText(els.originalSize, '');
    animateNumber(els.originalSize, file.size, (v) => R.formatBytes(Math.round(v)));
    els.qualityHint.textContent = 'กำลังคำนวณขนาดจริงของทุกระดับ สามารถเลือกระดับไว้ก่อนได้';
    for (const key of Object.keys(cards)) setCardEstimate(key, 'รอคำนวณ', 'loading');
    setCardsEnabled(true);
    refreshSelection();
    setProgress(0);
    els.progress.classList.add('is-active');
    setStatus('กำลังสร้างไฟล์จำลองทุกระดับความชัด...', 'busy');
    updateSummary();
    updateButtons();

    try {
      await R.analyzeAllPresets(buffer, {
        isCancelled,
        onProgress: (info) => {
          if (isCancelled()) return;
          onAnalyzeProgress(info);
        },
        onPresetReady: (key, result) => {
          if (isCancelled()) return;
          state.results[key] = result;
          setCardEstimate(key, `คาดว่าจะได้ ${R.formatBytes(result.resizedSize)}`, 'ready', reductionLabel(result));
          // แถบเทียบขนาดระหว่างระดับ (ระดับที่ใหญ่ที่สุด = เต็มแถบ)
          const largest = Math.max(...Object.values(state.results).map((r) => r.resizedSize), 1);
          for (const [readyKey, readyResult] of Object.entries(state.results)) {
            setCardBar(readyKey, readyResult.resizedSize / largest);
          }
          updateSummary();
          updateButtons();
        },
      });
      if (isCancelled()) return;
      state.analyzing = false;
      setProgress(100);
      els.progress.classList.remove('is-active');
      setStatus('คำนวณครบทุกระดับแล้ว เลือกระดับที่ต้องการและดาวน์โหลดไฟล์', 'done');
      els.qualityHint.textContent = 'เปรียบเทียบความชัดและขนาดไฟล์ได้จากตัวเลขบนแต่ละการ์ด';
      updateButtons();
    } catch (error) {
      if (isCancelled() || (error && error.name === 'CancelledError')) return;
      console.error(error);
      state.analyzing = false;
      state.failed = true;
      setProgress(0);
      els.progress.classList.remove('is-active');
      setStatus('ไม่สามารถคำนวณขนาดไฟล์ได้', 'error');
      for (const key of Object.keys(cards)) {
        if (!state.results[key]) setCardEstimate(key, 'คำนวณไม่สำเร็จ', 'error');
      }
      updateSummary();
      updateButtons();
      showDialog('ไม่สามารถวิเคราะห์ไฟล์ได้', error && error.message ? error.message : String(error), 'error');
    }
  }

  function onAnalyzeProgress(info) {
    // 0-5% อ่านไฟล์, 5-70% ประมวลผลรูป, 70-100% ประกอบไฟล์ทุกระดับ
    if (info.phase === 'open') {
      setProgress(0);
      setStatus('กำลังเปิดไฟล์...', 'busy');
    } else if (info.phase === 'read') {
      setProgress((info.current / Math.max(1, info.total)) * 5);
      setStatus(`กำลังอ่านข้อมูลในไฟล์ ${info.current}/${info.total}`, 'busy');
    } else if (info.phase === 'images') {
      setProgress(5 + ((info.current - 1) / Math.max(1, info.total)) * 65);
      setStatus(`กำลังวิเคราะห์รูป ${info.current}/${info.total}  •  ${info.name}`, 'busy');
      for (const key of Object.keys(cards)) {
        if (!state.results[key]) setCardEstimate(key, 'กำลังคำนวณ...', 'loading');
      }
    } else if (info.phase === 'build') {
      const within = (info.percent || 0) / 100;
      setProgress(70 + ((info.presetIndex - 1 + within) / info.totalPresets) * 30);
      setStatus(`กำลังสร้างไฟล์ระดับ ${info.preset.title} (${info.presetIndex}/${info.totalPresets})`, 'busy');
    }
  }

  function download() {
    if (window.DemoTrial && !window.DemoTrial.allow()) return; // DEMO: หมดระยะทดลอง
    const result = state.results[state.selectedKey];
    if (!state.file || !result) {
      showDialog('ยังคำนวณไม่เสร็จ', 'กรุณารอให้ระบบคำนวณขนาดของระดับที่เลือกเสร็จก่อน', 'error');
      return;
    }

    const preset = R.PRESET_BY_KEY[result.presetKey];
    const outputName = R.makeOutputName(state.file.name);
    const url = URL.createObjectURL(result.blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = outputName;
    link.rel = 'noopener';
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    recordDownloadUsage();

    state.downloaded = true;
    updateSteps();
    setStatus(`ดาวน์โหลดไฟล์เรียบร้อย: ${outputName}`, 'done');
    showDialog(
      'ดาวน์โหลดเรียบร้อย',
      `ระดับ: ${preset.title}\nไฟล์: ${outputName}\n\n` +
        `${R.formatBytes(result.originalSize)} → ${R.formatBytes(result.resizedSize)}\n` +
        (result.resizedSize >= result.originalSize
          ? 'ขนาดไม่เล็กลง (ไฟล์นี้แทบไม่มีรูปที่บีบอัดได้)'
          : `ลดลง ${result.reductionPercent.toFixed(1)}%`),
      'success',
    );
  }

  // ---------------------------------------------------------------------------
  // Events
  // ---------------------------------------------------------------------------

  function bindEvents() {
    els.fileInput.addEventListener('change', () => {
      const file = els.fileInput.files && els.fileInput.files[0];
      // allow re-selecting the same file later
      els.fileInput.value = '';
      handleFile(file);
    });

    // ปุ่มจริง (โฟกัสด้วยคีย์บอร์ดได้) เปิดตัวเลือกไฟล์; คลิกพื้นที่ drop zone ก็เปิดได้เช่นกัน
    els.pickButton.addEventListener('click', () => els.fileInput.click());
    els.dropZone.addEventListener('click', (event) => {
      if (event.target.closest('button, input')) return;
      els.fileInput.click();
    });

    els.downloadButton.addEventListener('click', download);
    els.resetButton.addEventListener('click', resetUi);

    document.querySelectorAll('.btn').forEach((btn) => btn.addEventListener('pointerdown', addRipple));
    for (const card of Object.values(cards)) card.addEventListener('pointerdown', addRipple);

    // Prevent the browser from opening a file dropped outside the drop zone.
    window.addEventListener('dragover', (event) => event.preventDefault());
    window.addEventListener('drop', (event) => {
      event.preventDefault();
      dragDepth = 0;
      els.dropZone.classList.remove('dragover');
      const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      if (file) handleFile(file);
    });

    let dragDepth = 0;
    els.dropZone.addEventListener('dragenter', (event) => {
      event.preventDefault();
      dragDepth += 1;
      els.dropZone.classList.add('dragover');
    });
    els.dropZone.addEventListener('dragleave', () => {
      dragDepth = Math.max(0, dragDepth - 1);
      if (!dragDepth) els.dropZone.classList.remove('dragover');
    });

    // Close the dialog when clicking the backdrop.
    els.dialog.addEventListener('click', (event) => {
      if (event.target === els.dialog) els.dialog.close();
    });
  }

  function checkSupport() {
    const missing = [];
    if (typeof JSZip === 'undefined') missing.push('JSZip');
    if (!window.Blob || !Blob.prototype.arrayBuffer) missing.push('Blob.arrayBuffer');
    if (!HTMLCanvasElement.prototype.toBlob) missing.push('canvas.toBlob');
    if (missing.length) {
      showDialog('เบราว์เซอร์ไม่รองรับ', `กรุณาใช้ Google Chrome หรือ Microsoft Edge รุ่นล่าสุด\n(ขาด: ${missing.join(', ')})`, 'error');
      els.pickButton.disabled = true;
      els.fileInput.disabled = true;
    }
  }

  buildCards();
  bindEvents();
  resetUi();
  checkSupport();
  initializeUsageCounter();
})();
