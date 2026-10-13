/*
 * PPTX resize engine — port of PPTX_Resizer_GUI.py (resize_pptx) to the browser.
 * Uses JSZip for the package and Canvas for safe JPEG re-encoding.
 *
 * Important: image part names, XML relationships, OLE objects, and non-JPEG
 * images are deliberately left untouched. Rewriting those package parts can
 * make PowerPoint repair the file or discard slide content.
 */
(function (global) {
  'use strict';

  // ---------------------------------------------------------------------------
  // Compression profiles
  // ---------------------------------------------------------------------------

  const QUALITY_PRESETS = Object.freeze([
    { key: 'smallest', title: 'เล็กที่สุด', maxImageSide: 960, jpegQuality: 28,
      description: 'บีบอัดสูงสุดตาม Logic เดิม', useCase: 'เหมาะกับส่งอีเมลและแชร์ผ่านแชต' },
    { key: 'compact', title: 'ประหยัดพื้นที่', maxImageSide: 1280, jpegQuality: 42,
      description: 'ไฟล์เล็ก แต่รายละเอียดเพิ่มขึ้น', useCase: 'เหมาะกับดูบนโน้ตบุ๊กและประชุมทั่วไป' },
    { key: 'balanced', title: 'สมดุล', maxImageSide: 1600, jpegQuality: 58,
      description: 'ขนาดไฟล์และความคมชัดสมดุลกัน', useCase: 'แนะนำสำหรับงาน PowerPoint ทั่วไป' },
    { key: 'clear', title: 'คมชัด', maxImageSide: 1920, jpegQuality: 72,
      description: 'รักษาตัวอักษรและรายละเอียดได้ดี', useCase: 'เหมาะกับจอใหญ่และห้องประชุม' },
    { key: 'high', title: 'คมชัดสูง', maxImageSide: 2560, jpegQuality: 85,
      description: 'ลดขนาดแบบรักษาคุณภาพมากที่สุด', useCase: 'เหมาะกับภาพรายละเอียดสูงและงานพิมพ์' },
  ]);

  const PRESET_BY_KEY = Object.fromEntries(QUALITY_PRESETS.map((p) => [p.key, p]));
  const DEFAULT_PRESET_KEY = 'balanced';

  const RASTER_EXTENSIONS = new Set(['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff', '.webp']);
  const SAFE_REENCODE_EXTENSIONS = new Set(['.jpg', '.jpeg']);
  const TEXT_PACKAGE_EXTENSIONS = new Set(['.xml', '.rels', '.vml']);
  const MIME_BY_EXTENSION = {
    '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.gif': 'image/gif',
    '.bmp': 'image/bmp', '.tif': 'image/tiff', '.tiff': 'image/tiff', '.webp': 'image/webp',
  };
  const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  const CONTENT_TYPES = '[Content_Types].xml';
  const PRESENTATION_XML = 'ppt/presentation.xml';
  const MEDIA_PREFIX = 'ppt/media/';

  const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

  class CancelledError extends Error {
    constructor() {
      super('cancelled');
      this.name = 'CancelledError';
    }
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

  function formatBytes(value) {
    const units = ['B', 'KB', 'MB', 'GB'];
    let size = Number(value);
    for (const unit of units) {
      if (size < 1024 || unit === units[units.length - 1]) {
        return `${size.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ${unit}`;
      }
      size /= 1024;
    }
    return `${value} B`;
  }

  function specification(preset) {
    return `${preset.maxImageSide.toLocaleString('en-US')} px  |  JPEG ${preset.jpegQuality}`;
  }

  function basename(path) {
    return path.slice(path.lastIndexOf('/') + 1);
  }

  function splitExtension(name) {
    const index = name.lastIndexOf('.');
    return index > 0 ? [name.slice(0, index), name.slice(index)] : [name, ''];
  }

  function extensionOf(path) {
    return splitExtension(basename(path))[1].toLowerCase();
  }

  /** สร้างชื่อ output โดยเติม _Resize ต่อท้ายชื่อเดิม */
  function makeOutputName(fileName) {
    const [stem, extension] = splitExtension(fileName);
    return `${stem}_Resize${extension || '.pptx'}`;
  }

  function compareNames(a, b) {
    if (a === b) return 0;
    return a < b ? -1 : 1;
  }

  /** Normalise a path inside the package; returns null when it escapes the root. */
  function normalizePackagePath(segments) {
    const result = [];
    for (const segment of segments) {
      if (!segment || segment === '.') continue;
      if (segment === '..') {
        if (!result.length) return null;
        result.pop();
      } else {
        result.push(segment);
      }
    }
    return result.join('/');
  }

  function isUnsafeEntryName(name) {
    const normalized = name.replace(/\\/g, '/');
    if (normalized.startsWith('/') || /^[a-zA-Z]:/.test(normalized)) return true;
    return normalizePackagePath(normalized.split('/')) === null;
  }

  /** แปลง Target ใน .rels เป็น path จริงภายใน package */
  function resolveRelationshipTarget(sourcePart, target) {
    let normalizedTarget = target.replace(/\\/g, '/');
    try {
      normalizedTarget = decodeURIComponent(normalizedTarget);
    } catch (_error) {
      // keep the raw target
    }
    const base = normalizedTarget.startsWith('/') ? [] : sourcePart.split('/').slice(0, -1);
    return normalizePackagePath(base.concat(normalizedTarget.split('/')));
  }

  function parseXml(bytes) {
    try {
      const text = new TextDecoder('utf-8').decode(bytes);
      const doc = new DOMParser().parseFromString(text, 'application/xml');
      if (doc.getElementsByTagName('parsererror').length) return null;
      return doc;
    } catch (_error) {
      return null;
    }
  }

  function nextFrame() {
    return new Promise((resolve) => setTimeout(resolve, 0));
  }

  // ---------------------------------------------------------------------------
  // Package loading
  // ---------------------------------------------------------------------------

  async function loadPresentation(buffer) {
    if (typeof JSZip === 'undefined') {
      throw new Error('โหลดไลบรารี JSZip ไม่สำเร็จ กรุณารีเฟรชหน้าเว็บ');
    }

    const head = new Uint8Array(buffer, 0, Math.min(8, buffer.byteLength));
    const isCompoundFile = head.length >= 4 && head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0;
    if (isCompoundFile) {
      throw new Error('ไฟล์นี้ถูกตั้งรหัสผ่าน หรือเป็นไฟล์ .ppt รุ่นเก่า\nกรุณาเปิดด้วย PowerPoint แล้วบันทึกเป็น .pptx แบบไม่ใส่รหัสผ่านก่อน');
    }

    let zip;
    try {
      zip = await JSZip.loadAsync(buffer);
    } catch (_error) {
      throw new Error('ไฟล์ PPTX เสียหาย หรือไม่ใช่ไฟล์ PowerPoint ที่ถูกต้อง');
    }

    zip.forEach((_relativePath, entry) => {
      if (isUnsafeEntryName(entry.name)) {
        throw new Error(`พบ path ที่ไม่ปลอดภัยในไฟล์ PPTX: ${entry.name}`);
      }
    });

    if (!zip.file(PRESENTATION_XML)) {
      throw new Error('ไฟล์ไม่ใช่ PPTX ที่สมบูรณ์: ไม่พบ ppt/presentation.xml');
    }
    if (!zip.file(CONTENT_TYPES)) {
      throw new Error('ไฟล์ไม่ใช่ PPTX ที่สมบูรณ์: ไม่พบ [Content_Types].xml');
    }
    return zip;
  }

  // ---------------------------------------------------------------------------
  // Image compression
  // ---------------------------------------------------------------------------

  function targetSize(width, height, maxSide) {
    const longestSide = Math.max(width, height);
    if (longestSide <= maxSide) return [width, height];
    const scale = maxSide / longestSide;
    return [Math.max(1, Math.floor(width * scale)), Math.max(1, Math.floor(height * scale))];
  }

  /**
   * PowerPoint และ Pillow ไม่หมุนภาพตาม EXIF Orientation แต่เบราว์เซอร์หมุนให้เสมอ
   * จึงรีเซ็ต Orientation เป็น 1 (สำเนาใหม่) ก่อน decode เพื่อให้ภาพออกมาเหมือนต้นฉบับ
   */
  function neutralizeJpegOrientation(bytes) {
    if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return bytes;

    let offset = 2;
    while (offset + 4 <= bytes.length) {
      if (bytes[offset] !== 0xff) return bytes;
      const marker = bytes[offset + 1];
      if (marker === 0xff) {
        offset += 1; // fill byte
        continue;
      }
      if (marker === 0xda || marker === 0xd9) return bytes; // start of scan / end of image
      const segmentLength = (bytes[offset + 2] << 8) | bytes[offset + 3];
      if (segmentLength < 2) return bytes;

      const isExif = marker === 0xe1 && segmentLength >= 16 &&
        bytes[offset + 4] === 0x45 && bytes[offset + 5] === 0x78 && bytes[offset + 6] === 0x69 &&
        bytes[offset + 7] === 0x66 && bytes[offset + 8] === 0 && bytes[offset + 9] === 0;
      if (isExif) {
        const tiff = offset + 10;
        const segmentEnd = Math.min(bytes.length, offset + 2 + segmentLength);
        const little = bytes[tiff] === 0x49 && bytes[tiff + 1] === 0x49;
        const big = bytes[tiff] === 0x4d && bytes[tiff + 1] === 0x4d;
        if (!little && !big) return bytes;
        const u16 = (p) => (little ? bytes[p] | (bytes[p + 1] << 8) : (bytes[p] << 8) | bytes[p + 1]);
        const u32 = (p) => (little
          ? (bytes[p] | (bytes[p + 1] << 8) | (bytes[p + 2] << 16)) + bytes[p + 3] * 0x1000000
          : bytes[p] * 0x1000000 + ((bytes[p + 1] << 16) | (bytes[p + 2] << 8) | bytes[p + 3]));

        const ifd = tiff + u32(tiff + 4);
        if (ifd + 2 > segmentEnd) return bytes;
        const count = u16(ifd);
        for (let i = 0; i < count; i += 1) {
          const entry = ifd + 2 + i * 12;
          if (entry + 12 > segmentEnd) break;
          if (u16(entry) === 0x0112) {
            if (u16(entry + 8) === 1) return bytes;
            const copy = new Uint8Array(bytes);
            copy[entry + 8] = little ? 1 : 0;
            copy[entry + 9] = little ? 0 : 1;
            return copy;
          }
        }
        return bytes;
      }
      offset += 2 + segmentLength;
    }
    return bytes;
  }

  async function decodeImage(bytes, name) {
    const type = MIME_BY_EXTENSION[extensionOf(name)] || '';
    const blob = new Blob([neutralizeJpegOrientation(bytes)], { type });

    if (typeof createImageBitmap === 'function') {
      try {
        return await createImageBitmap(blob);
      } catch (_error) {
        // fall through to <img> decoding
      }
    }

    const url = URL.createObjectURL(blob);
    try {
      const image = new Image();
      image.decoding = 'async';
      await new Promise((resolve, reject) => {
        image.onload = resolve;
        image.onerror = () => reject(new Error('decode failed'));
        image.src = url;
      });
      if (!image.naturalWidth || !image.naturalHeight) throw new Error('empty image');
      return image;
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  function sourceSize(source) {
    return [source.naturalWidth || source.width, source.naturalHeight || source.height];
  }

  /** ย่อทั้งภาพแบบรักษาสัดส่วน (step-down เพื่อให้ภาพไม่แตกเมื่อย่อมาก ๆ) */
  function drawScaled(ctx, source, width, height) {
    let [currentWidth, currentHeight] = sourceSize(source);
    let current = source;
    const temps = [];

    while (currentWidth > width * 2 && currentHeight > height * 2) {
      const nextWidth = Math.max(width, Math.floor(currentWidth / 2));
      const nextHeight = Math.max(height, Math.floor(currentHeight / 2));
      const temp = document.createElement('canvas');
      temp.width = nextWidth;
      temp.height = nextHeight;
      const tempCtx = temp.getContext('2d');
      tempCtx.imageSmoothingEnabled = true;
      tempCtx.imageSmoothingQuality = 'high';
      tempCtx.drawImage(current, 0, 0, nextWidth, nextHeight);
      temps.push(temp);
      current = temp;
      currentWidth = nextWidth;
      currentHeight = nextHeight;
    }

    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(current, 0, 0, width, height);

    for (const temp of temps) {
      temp.width = 0;
      temp.height = 0;
    }
  }

  /** คืนค่า JPEG bytes ที่บีบอัดแล้ว หรือ null เมื่อแปลงไม่ได้ */
  async function encodeJpeg(source, preset) {
    const [sourceWidth, sourceHeight] = sourceSize(source);
    const [width, height] = targetSize(sourceWidth, sourceHeight, preset.maxImageSide);

    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d', { alpha: false });
    if (!ctx) return null;

    // ภาพโปร่งใสใช้พื้นหลังสีขาวเหมือน Logic เดิม
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, width, height);
    drawScaled(ctx, source, width, height);

    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', preset.jpegQuality / 100));
    canvas.width = 0;
    canvas.height = 0;
    if (!blob || blob.type !== 'image/jpeg') return null;
    return new Uint8Array(await blob.arrayBuffer());
  }

  /** เข้ารหัสภาพเดียวสำหรับทุก preset (decode ครั้งเดียว); null เมื่อเปิดภาพไม่ได้ */
  async function compressImageForPresets(bytes, name) {
    let source;
    try {
      source = await decodeImage(bytes, name);
    } catch (_error) {
      return null; // ภาพที่เบราว์เซอร์อ่านไม่ได้จะถูกเก็บไว้เหมือนเดิม
    }

    try {
      const output = {};
      for (const preset of QUALITY_PRESETS) {
        try {
          output[preset.key] = await encodeJpeg(source, preset);
        } catch (_error) {
          output[preset.key] = null;
        }
      }
      return output;
    } finally {
      if (source && typeof source.close === 'function') source.close();
    }
  }

  // ---------------------------------------------------------------------------
  // Package rewriting
  // ---------------------------------------------------------------------------

  /** ประกอบไฟล์กลับเป็น PPTX */
  async function zipPackage(files, onProgress) {
    const output = new JSZip();
    const names = [...files.keys()].sort((a, b) => {
      if (a === CONTENT_TYPES) return -1;
      if (b === CONTENT_TYPES) return 1;
      return compareNames(a, b);
    });

    for (const name of names) {
      const isText = TEXT_PACKAGE_EXTENSIONS.has(extensionOf(name));
      output.file(name, files.get(name), {
        binary: true,
        createFolders: false,
        compression: 'DEFLATE',
        compressionOptions: { level: isText ? 9 : 6 },
      });
    }

    return output.generateAsync(
      { type: 'blob', mimeType: PPTX_MIME, compression: 'DEFLATE', compressionOptions: { level: 9 } },
      (metadata) => onProgress && onProgress(metadata.percent),
    );
  }

  /** ตรวจสอบโครงสร้างพื้นฐานของ PPTX */
  function relationshipSourcePart(relsName) {
    if (relsName === '_rels/.rels') return '';
    const marker = '/_rels/';
    const markerIndex = relsName.lastIndexOf(marker);
    if (markerIndex < 0 || !relsName.endsWith('.rels')) return null;
    return `${relsName.slice(0, markerIndex)}/${relsName.slice(markerIndex + marker.length, -5)}`;
  }

  /** ตรวจ CRC, XML และ internal relationships ก่อนอนุญาตให้ดาวน์โหลด */
  async function validatePptx(blob) {
    let zip;
    try {
      zip = await JSZip.loadAsync(blob, { checkCRC32: true });
    } catch (_error) {
      throw new Error('ไม่สามารถสร้างไฟล์ PPTX output ที่สมบูรณ์ได้');
    }
    if (!zip.file(CONTENT_TYPES) || !zip.file(PRESENTATION_XML)) {
      throw new Error('Output ที่สร้างขึ้นมีโครงสร้าง PPTX ไม่ครบถ้วน');
    }

    const names = [];
    zip.forEach((_relativePath, entry) => {
      if (!entry.dir) names.push(entry.name);
    });
    const nameSet = new Set(names);

    for (const name of names) {
      if (!TEXT_PACKAGE_EXTENSIONS.has(extensionOf(name)) && name !== CONTENT_TYPES) continue;
      const bytes = await zip.file(name).async('uint8array');
      const doc = parseXml(bytes);
      if (!doc) throw new Error(`Output มี XML ไม่สมบูรณ์: ${name}`);

      if (!name.toLowerCase().endsWith('.rels')) continue;
      const sourcePart = relationshipSourcePart(name);
      if (sourcePart === null) continue;

      const relationships = Array.from(doc.getElementsByTagNameNS(RELATIONSHIPS_NS, 'Relationship'));
      for (const relationship of relationships) {
        if ((relationship.getAttribute('TargetMode') || '').toLowerCase() === 'external') continue;
        const target = (relationship.getAttribute('Target') || '').split(/[?#]/, 1)[0];
        if (!target) continue;
        const targetPath = resolveRelationshipTarget(sourcePart, target);
        if (!targetPath || !nameSet.has(targetPath)) {
          throw new Error(`Output อ้างอิงข้อมูลที่หายไป: ${name} → ${target}`);
        }
      }
    }
  }

  /**
   * ตรวจทุก XML/RELS ที่ถูกเขียนใหม่ว่ายังเป็น XML ที่ถูกต้อง — ถ้าเสียให้ error ทันที
   * ดีกว่าปล่อยไฟล์ที่ PowerPoint ต้อง Repair และทำให้สไลด์ว่าง
   */
  function assertRewrittenXmlIsValid(files, pristineFiles) {
    for (const [name, data] of files) {
      if (pristineFiles.get(name) === data) continue;
      if (!TEXT_PACKAGE_EXTENSIONS.has(extensionOf(name))) continue;
      if (!parseXml(data)) {
        throw new Error(`ตรวจพบ XML เสียหลังแปลง: ${name}\nระบบยกเลิกเพื่อไม่ให้ได้ไฟล์ที่ต้อง Repair — กรุณาแจ้งผู้พัฒนา`);
      }
    }
  }

  async function buildPresetPackage(context, preset, onZipProgress) {
    const { baseFiles, imageNames, encodedImages } = context;
    const files = new Map(baseFiles);
    let compressedImages = 0;
    let skippedImages = 0;

    for (const name of imageNames) {
      const original = baseFiles.get(name);
      const encoded = encodedImages.get(name);
      const data = encoded ? encoded[preset.key] : null;

      // ใช้ไฟล์ใหม่เฉพาะเมื่อเล็กกว่าไฟล์เดิม
      if (!data || data.length >= original.length) {
        skippedImages += 1;
        continue;
      }

      // เก็บชื่อ part และ Content-Type เดิมเสมอ จึงไม่มีการแก้ slide XML/.rels
      files.set(name, data);
      compressedImages += 1;
    }

    assertRewrittenXmlIsValid(files, context.pristineFiles);
    const blob = await zipPackage(files, onZipProgress);
    await validatePptx(blob);

    const originalSize = context.originalSize;
    const resizedSize = blob.size;
    const savedBytes = Math.max(0, originalSize - resizedSize);

    return {
      presetKey: preset.key,
      blob,
      originalSize,
      resizedSize,
      savedBytes,
      reductionPercent: originalSize > 0 ? (savedBytes * 100) / originalSize : 0,
      totalImages: context.totalImages,
      compressedImages,
      convertedImages: 0,
      skippedImages: context.totalImages - compressedImages,
      flattenedOleObjects: 0,
    };
  }

  /**
   * วิเคราะห์ไฟล์และสร้างไฟล์จำลองจริงของทุกระดับความชัด
   * options: { onProgress(info), onPresetReady(key, result), isCancelled() }
   */
  async function analyzeAllPresets(buffer, options = {}) {
    const onProgress = options.onProgress || (() => {});
    const checkCancelled = () => {
      if (options.isCancelled && options.isCancelled()) throw new CancelledError();
    };

    onProgress({ phase: 'open', fraction: 0 });
    const zip = await loadPresentation(buffer);
    checkCancelled();

    const entries = [];
    zip.forEach((_relativePath, entry) => {
      if (!entry.dir) entries.push(entry);
    });

    const baseFiles = new Map();
    for (let index = 0; index < entries.length; index += 1) {
      baseFiles.set(entries[index].name, await entries[index].async('uint8array'));
      if (index % 20 === 0) {
        onProgress({ phase: 'read', current: index + 1, total: entries.length });
        checkCancelled();
      }
    }

    const allImageNames = [...baseFiles.keys()]
      .filter((name) => name.startsWith(MEDIA_PREFIX) && !name.slice(MEDIA_PREFIX.length).includes('/'))
      .filter((name) => RASTER_EXTENSIONS.has(extensionOf(name)))
      .sort(compareNames);
    // PNG/GIF/BMP/TIFF/WebP และภาพที่มี transparency คง bytes เดิมทั้งหมด
    // เพื่อไม่เปลี่ยนชนิดไฟล์หรือ relationship ภายใน PPTX
    const imageNames = allImageNames.filter((name) => SAFE_REENCODE_EXTENSIONS.has(extensionOf(name)));

    const encodedImages = new Map();
    for (let index = 0; index < imageNames.length; index += 1) {
      const name = imageNames[index];
      onProgress({ phase: 'images', current: index + 1, total: imageNames.length, name: basename(name) });
      encodedImages.set(name, await compressImageForPresets(baseFiles.get(name), name));
      checkCancelled();
    }

    const pristineFiles = new Map(baseFiles);
    await nextFrame();
    checkCancelled();

    const context = {
      baseFiles,
      pristineFiles,
      imageNames,
      encodedImages,
      totalImages: allImageNames.length,
      originalSize: buffer.byteLength,
    };

    const results = {};
    for (let index = 0; index < QUALITY_PRESETS.length; index += 1) {
      const preset = QUALITY_PRESETS[index];
      onProgress({ phase: 'build', presetIndex: index + 1, totalPresets: QUALITY_PRESETS.length, preset, percent: 0 });
      const result = await buildPresetPackage(context, preset, (percent) =>
        onProgress({ phase: 'build', presetIndex: index + 1, totalPresets: QUALITY_PRESETS.length, preset, percent }),
      );
      checkCancelled();
      results[preset.key] = result;
      if (options.onPresetReady) options.onPresetReady(preset.key, result);
    }
    return results;
  }

  global.PptxResizer = {
    QUALITY_PRESETS,
    PRESET_BY_KEY,
    DEFAULT_PRESET_KEY,
    CancelledError,
    formatBytes,
    specification,
    makeOutputName,
    loadPresentation,
    analyzeAllPresets,
  };
})(window);
