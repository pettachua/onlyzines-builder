// ============ UTILITIES ============
// Normalize asset URLs (passthrough â R2 pub URLs work directly)
function normalizeAssetUrl(url) {
  return url;
}

// ============ IMAGE COMPRESSION ============
// Compress images on import to keep save payloads under API limits
function compressImage(dataUrl, maxDimension, quality) {
  maxDimension = maxDimension || 1600;
  quality = quality || 0.82;
  
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      let w = img.naturalWidth;
      let h = img.naturalHeight;
      
      // If already small enough and it's a JPEG, skip compression
      if (w <= maxDimension && h <= maxDimension && dataUrl.startsWith('data:image/jpeg')) {
        // Check if the data URL is already reasonably sized (< 500KB)
        if (dataUrl.length < 500000) {
          resolve(dataUrl);
          return;
        }
      }
      
      // Scale down if needed
      if (w > maxDimension || h > maxDimension) {
        const scale = maxDimension / Math.max(w, h);
        w = Math.round(w * scale);
        h = Math.round(h * scale);
      }
      
      const canvas = document.createElement('canvas');
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, w, h);
      
      // Try JPEG first (smaller), fall back to PNG for transparency
      let compressed = canvas.toDataURL('image/jpeg', quality);
      
      // If original was PNG and might have transparency, keep as PNG but still resized
      if (dataUrl.startsWith('data:image/png') && compressed.length > dataUrl.length * 0.9) {
        compressed = canvas.toDataURL('image/png');
      }
      
      resolve(compressed);
    };
    img.onerror = () => resolve(dataUrl); // On error, return original
    img.src = dataUrl;
  });
}

// ============ R2 IMAGE UPLOAD ============
// In-flight dedup: if the same data URL is already being uploaded, reuse the promise
const _uploadInflight = new Map();

function _uploadKey(dataUrl) {
  const start = dataUrl.indexOf(',') + 1;
  const mid = start + Math.floor((dataUrl.length - start) / 2);
  return dataUrl.length + '_' + dataUrl.substring(mid, mid + 16);
}

// Uploads a compressed data URL to R2 via the backend, returns the public URL
// Includes 30s timeout via AbortController and one automatic retry
async function uploadImageToR2(dataUrl) {
  // Dedup: if this exact data URL is already uploading, piggyback on it
  if (dataUrl.startsWith('data:')) {
    const key = _uploadKey(dataUrl);
    if (_uploadInflight.has(key)) return _uploadInflight.get(key);
    const promise = _uploadImageToR2Inner(dataUrl).finally(() => _uploadInflight.delete(key));
    _uploadInflight.set(key, promise);
    return promise;
  }
  return dataUrl; // already a URL, nothing to upload
}

async function _uploadImageToR2Inner(dataUrl) {
  const MAX_RETRIES = 1;
  const TIMEOUT_MS = 30000;

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), TIMEOUT_MS);

    try {
      const res = await apiAdapter.fetch(`${API_BASE}/api/upload`, {
        method: 'POST',
        body: JSON.stringify({ dataUrl }),
        signal: controller.signal
      });
      clearTimeout(timeoutId);

      if (!res || !res.ok) {
        console.error(`R2 upload failed (attempt ${attempt + 1}):`, res?.status);
        if (attempt < MAX_RETRIES) continue; // retry
        updateSaveStatus('â  Image stored locally (cloud upload failed)');
        return dataUrl;
      }
      const data = await res.json();
      if (!data || !data.url) {
        console.error(`R2 upload returned no URL (attempt ${attempt + 1})`);
        if (attempt < MAX_RETRIES) continue; // retry
        updateSaveStatus('â  Image stored locally (invalid response)');
        return dataUrl;
      }
      return data.url;
    } catch (err) {
      clearTimeout(timeoutId);
      const isTimeout = err.name === 'AbortError';
      console.error(`R2 upload ${isTimeout ? 'timed out' : 'error'} (attempt ${attempt + 1}):`, err);
      if (attempt < MAX_RETRIES) continue; // retry
      updateSaveStatus(isTimeout ? 'â  Image upload timed out â stored locally' : 'â  Image stored locally (cloud upload failed)');
      return dataUrl;
    }
  }
  return dataUrl; // fallback
}

// ============ SAVE & EXPORT ============
function saveJSON() {
  const title = document.querySelector('.topbar-title')?.textContent || 'Untitled Zine';
  const data = { version: '20', title, pages: state.pages };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = title.replace(/[^a-z0-9]/gi, '_') + '.json'; a.click();
}

// html2canvas v1.4.1 does NOT render CSS filter: (brightness, contrast, saturate).
// Before capture, bake any per-image CSS filters into the pixel data so html2canvas
// receives pre-filtered images with no CSS filter to ignore.
// MUST run before pdfFixImages — while images are still <img> elements.
async function preFilterImages(container) {
  var imgs = container.querySelectorAll('img');
  for (var i = 0; i < imgs.length; i++) {
    var img = imgs[i];
    var f = img.style.filter;
    if (!f) continue;
    if (!/brightness|contrast|saturate/.test(f)) continue;
    if (!img.complete || img.naturalWidth === 0) continue;
    try {
      var c = document.createElement('canvas');
      c.width = img.naturalWidth;
      c.height = img.naturalHeight;
      var ctx = c.getContext('2d');
      ctx.filter = f;
      ctx.drawImage(img, 0, 0, c.width, c.height);
      img.src = c.toDataURL('image/png');
      img.style.filter = '';
    } catch (e) {
      try {
        var corsImg = new Image();
        corsImg.crossOrigin = 'anonymous';
        await new Promise(function(resolve, reject) {
          corsImg.onload = resolve;
          corsImg.onerror = reject;
          var bust = img.src.indexOf('?') === -1 ? '?' : '&';
          corsImg.src = img.src + bust + '_cors=' + Date.now();
        });
        var c2 = document.createElement('canvas');
        c2.width = corsImg.naturalWidth;
        c2.height = corsImg.naturalHeight;
        var ctx2 = c2.getContext('2d');
        ctx2.filter = f;
        ctx2.drawImage(corsImg, 0, 0, c2.width, c2.height);
        img.src = c2.toDataURL('image/png');
        img.style.filter = '';
      } catch (e2) {
        console.warn('preFilterImages: could not bake filter for', img.src && img.src.substring(0, 60), e2);
      }
    }
  }
}
// html2canvas does NOT support object-fit on <img>. Before capture,
// convert any <img> using object-fit into a <div> with background-image.
function pdfFixImages(container) {
  container.querySelectorAll('img').forEach(img => {
    const fit = img.style.objectFit;
    if (!fit || (fit !== 'cover' && fit !== 'contain')) return;
    const div = document.createElement('div');
    div.style.width = img.style.width || '100%';
    div.style.height = img.style.height || '100%';
    div.style.backgroundImage = 'url("' + img.src + '")';
    div.style.backgroundSize = fit;
    div.style.backgroundPosition = 'center';
    div.style.backgroundRepeat = 'no-repeat';
    if (img.style.opacity) div.style.opacity = img.style.opacity;
    if (img.style.transform) { div.style.transform = img.style.transform; div.style.transformOrigin = img.style.transformOrigin; }
    ['webkitMaskImage','maskImage','webkitMaskSize','maskSize'].forEach(p => { if (img.style[p]) div.style[p] = img.style[p]; });
    img.parentNode.replaceChild(div, img);
  });
  container.querySelectorAll('svg').forEach(svg => svg.remove());
  container.querySelectorAll('[style*="clip-path"]').forEach(el => {
    el.style.clipPath = 'none';
    el.style.webkitClipPath = 'none';
  });
}

// Lazy-load a script, returns a promise.  Resolves when loaded, rejects on error.
function loadScript(url) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = url;
    s.onload = resolve;
    s.onerror = () => reject(new Error('Failed to load ' + url));
    document.head.appendChild(s);
  });
}

// Ensure PDF libraries are loaded (only fetched once)
async function ensurePDFLibs() {
  if (window.html2canvas && window.jspdf) return;
  const status = document.getElementById('pdfStatus');
  if (status) status.textContent = 'Loading PDF libraries...';
  try {
    await Promise.all([
      window.html2canvas ? Promise.resolve() : loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js'),
      window.jspdf ? Promise.resolve() : loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js')
    ]);
  } catch (e) {
    throw new Error('Could not load PDF libraries. Check your internet connection and try again.');
  }
  // Verify they actually loaded
  if (!window.html2canvas) throw new Error('html2canvas failed to initialize');
  if (!window.jspdf || !window.jspdf.jsPDF) throw new Error('jsPDF failed to initialize');
}

// Ensure html2canvas is loaded (lighter than ensurePDFLibs — skips jsPDF)
async function ensureHtml2Canvas() {
  if (window.html2canvas) return;
  await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
  if (!window.html2canvas) throw new Error('html2canvas failed to load');
}

// Capture the cover page (page 0) as a JPEG data URL for publish-time cover generation.
// This is a standalone version of the capturePage helper inside savePDF, used by publishIssue.
async function captureCoverImage() {
  await ensureHtml2Canvas();
  const page = state.pages[0];
  if (!page) return null;

  const width = PAGE_W;
  const height = PAGE_H;

  const temp = document.createElement('div');
  temp.style.cssText = 'position:fixed;left:-9999px;top:0;width:' + width + 'px;height:' + height + 'px;overflow:hidden;pointer-events:none;background:' + (page.paper || '#f5f3ee');
  temp.innerHTML = renderPageElement(page);
  document.body.appendChild(temp);

  // Clean up selection artifacts
  temp.querySelectorAll('.handle, .crop-handle').forEach(el => el.remove());
  temp.querySelectorAll('.selected').forEach(el => el.classList.remove('selected'));
  temp.querySelectorAll('.multi-selected').forEach(el => el.classList.remove('multi-selected'));
  const pageDiv = temp.querySelector('.page');
  if (pageDiv) {
    pageDiv.classList.remove('active-page');
    temp.querySelectorAll('.active-page-indicator').forEach(el => el.remove());
  }

  // Force CORS reload: set crossOrigin and cache-bust all external images
  // so html2canvas can access them after pdfFixImages converts to background-image
  const imgs = temp.querySelectorAll('img');
  const cacheBust = '_cors=' + Date.now();
  imgs.forEach(img => {
    if (img.src && (img.src.startsWith('http://') || img.src.startsWith('https://'))) {
      const originalSrc = img.src;
      img.crossOrigin = 'anonymous';
      const sep = originalSrc.includes('?') ? '&' : '?';
      img.src = originalSrc + sep + cacheBust;
    }
  });

  // Wait for images to reload with CORS headers
  if (imgs.length > 0) {
    await Promise.all(Array.from(imgs).map(img => {
      if (img.complete && img.naturalWidth > 0) return Promise.resolve();
      return new Promise(resolve => {
        img.onload = resolve;
        img.onerror = () => resolve();
        setTimeout(resolve, 5000);
      });
    }));
  }

  await preFilterImages(temp);
  pdfFixImages(temp);
  await new Promise(r => requestAnimationFrame(() => setTimeout(r, 100)));

  let canvas;
  try {
    canvas = await html2canvas(temp, {
      scale: 2, width, height,
      backgroundColor: page.paper || '#f5f3ee',
      useCORS: true, allowTaint: false, logging: false
    });
  } catch (err) {
    console.error('Cover capture error:', err);
    document.body.removeChild(temp);
    return null;
  }
  document.body.removeChild(temp);

  return canvas.toDataURL('image/jpeg', 0.85);
}

async function savePDF() {
  const btnPDF = document.getElementById('btnPDF');
  const origHTML = btnPDF ? btnPDF.innerHTML : '';
  // Save editor selection state BEFORE try so catch can restore it
  const savedSelected = state.selected;
  const savedMultiSelected = state.multiSelected;
  const savedActivePage = state.activePage;
  const savedPositionMode = state.imagePositionMode;
  state.selected = null;
  state.multiSelected = [];
  state.activePage = null;
  state.imagePositionMode = null;
  try {
    if (btnPDF) { btnPDF.innerHTML = 'â³ Loading...'; btnPDF.disabled = true; }

    await ensurePDFLibs();

    const { jsPDF } = window.jspdf;
    const PW = PAGE_W;
    const PH = PAGE_H;
    const SW = SPREAD_W;

    // --- Helper: render one builder page into an offscreen div and capture to canvas ---
    async function capturePage(page, width, height) {
      const temp = document.createElement('div');
      temp.style.cssText = 'position:fixed;left:-9999px;top:0;width:' + width + 'px;height:' + height + 'px;overflow:hidden;pointer-events:none;background:' + (page.paper || '#f5f3ee');
      temp.innerHTML = renderPageElement(page);
      document.body.appendChild(temp);

      // Strip any selection handles and outlines that might have leaked
      temp.querySelectorAll('.handle, .crop-handle').forEach(el => el.remove());
      temp.querySelectorAll('.selected').forEach(el => el.classList.remove('selected'));
      temp.querySelectorAll('.multi-selected').forEach(el => el.classList.remove('multi-selected'));

      // Strip texture filter class (html2canvas can't render SVG filters)
      const pageDiv = temp.querySelector('.page');
      let textureType = null;
      if (pageDiv) {
        // Remove active-page highlight (blue outline) from export
        pageDiv.classList.remove('active-page');
        // Remove any indicator overlays
        temp.querySelectorAll('.active-page-indicator').forEach(el => el.remove());
        const cl = pageDiv.className;
        const m = cl.match(/paper-(smooth|matte|newsprint|cardstock|riso|distressed|patina|saigon)/);
        if (m) {
          textureType = m[1];
          pageDiv.classList.remove('paper-' + textureType);
        }
      }

      // Wait for images
      const imgs = temp.querySelectorAll('img');
      if (imgs.length > 0) {
        await Promise.all(Array.from(imgs).map(img => {
          if (img.complete && img.naturalWidth > 0) return Promise.resolve();
          return new Promise(resolve => {
            img.onload = resolve;
            img.onerror = () => { console.warn('PDF: img load failed', img.src?.substring(0,60)); resolve(); };
            setTimeout(resolve, 5000);
          });
        }));
      }

      await preFilterImages(temp);
      pdfFixImages(temp);
      await new Promise(r => requestAnimationFrame(() => setTimeout(r, 100)));

      let canvas;
      try {
        canvas = await html2canvas(temp, {
          scale: 2, width: width, height: height,
          backgroundColor: page.paper || '#f5f3ee',
          useCORS: true, allowTaint: true, logging: false
        });
      } catch (err) {
        console.error('html2canvas error:', err);
        // Fallback: blank canvas
        canvas = document.createElement('canvas');
        canvas.width = width * 2; canvas.height = height * 2;
      }
      document.body.removeChild(temp);

      // Apply texture overlay onto the captured canvas
      if (textureType && textureType !== 'smooth') {
        pdfApplyTexture(canvas, textureType);
      }

      return canvas;
    }

    // --- Helper: apply texture overlay to a canvas ---
    function pdfApplyTexture(canvas, textureType) {
      const ctx = canvas.getContext('2d');
      const w = canvas.width;
      const h = canvas.height;
      // Texture parameters: strength controls visibility, freq controls grain size
      const params = {
        matte:      { strength: 0.04, freq: 1,  warm: false, variance: 30 },
        newsprint:  { strength: 0.06, freq: 3,  warm: true,  variance: 50 },
        cardstock:  { strength: 0.07, freq: 2,  warm: false, variance: 40 },
        riso:       { strength: 0.10, freq: 1,  warm: false, variance: 60 },
        distressed: { strength: 0.12, freq: 2,  warm: true,  variance: 70 },
        patina:     { strength: 0.07, freq: 1,  warm: false, variance: 45 },
        saigon:     { strength: 0.05, freq: 4,  warm: true,  variance: 35 }
      };
      const p = params[textureType] || params.matte;
      // Build texture on separate canvas
      const texCanvas = document.createElement('canvas');
      texCanvas.width = w;
      texCanvas.height = h;
      const tCtx = texCanvas.getContext('2d');
      const imgData = tCtx.createImageData(w, h);
      const d = imgData.data;
      for (let i = 0; i < d.length; i += 4) {
        const px = (i / 4) % w;
        const py = Math.floor((i / 4) / w);
        let noise;
        if (p.freq > 1) {
          // Quantized grain for coarser texture
          const qx = Math.floor(px / p.freq) * p.freq;
          const qy = Math.floor(py / p.freq) * p.freq;
          noise = ((Math.sin(qx * 12.9898 + qy * 78.233) * 43758.5453) % 1);
          if (noise < 0) noise = -noise;
        } else {
          noise = Math.random();
        }
        // Variable darkness: base dark + variance
        const v = Math.floor(noise * p.variance);
        d[i]     = p.warm ? v + 15 : v;  // R
        d[i + 1] = p.warm ? v + 8  : v;  // G
        d[i + 2] = v;                     // B
        // Alpha varies per pixel for organic feel
        const baseAlpha = p.strength * 255;
        d[i + 3] = Math.floor(baseAlpha * (0.5 + noise * 0.5));
      }
      tCtx.putImageData(imgData, 0, 0);
      // Overlay onto main canvas (source-over = normal blend)
      ctx.globalCompositeOperation = 'source-over';
      ctx.drawImage(texCanvas, 0, 0);
    }

    // --- Build PDF ---
    // Page 1: Cover (single 400Ã600)
    const pdf = new jsPDF({ orientation: 'portrait', unit: 'px', format: [PW, PH] });

    if (btnPDF) btnPDF.innerHTML = 'â³ Cover...';
    const coverCanvas = await capturePage(state.pages[0], PW, PH);
    pdf.addImage(coverCanvas.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, PW, PH);

    // Pages 2+: Spreads (pairs of pages, 800Ã600)
    const innerPages = state.pages.slice(1);
    for (let i = 0; i < innerPages.length; i += 2) {
      const leftPage = innerPages[i];
      const rightPage = innerPages[i + 1]; // may be undefined if odd count

      pdf.addPage([SW, PH], 'landscape');

      const spreadIdx = Math.floor(i / 2) + 1;
      if (btnPDF) btnPDF.innerHTML = `â³ Spread ${spreadIdx}...`;

      // Capture left page
      const leftCanvas = await capturePage(leftPage, PW, PH);

      // Compose spread canvas (800Ã600 at 2x = 1600Ã1200)
      const spreadCanvas = document.createElement('canvas');
      spreadCanvas.width = SW * 2;
      spreadCanvas.height = PH * 2;
      const sCtx = spreadCanvas.getContext('2d');

      // Draw left page
      sCtx.drawImage(leftCanvas, 0, 0, PW * 2, PH * 2);

      // Draw right page (or blank)
      if (rightPage) {
        const rightCanvas = await capturePage(rightPage, PW, PH);
        sCtx.drawImage(rightCanvas, PW * 2, 0, PW * 2, PH * 2);
      } else {
        // Fill right half with default paper color
        sCtx.fillStyle = '#f5f3ee';
        sCtx.fillRect(PW * 2, 0, PW * 2, PH * 2);
      }

      pdf.addImage(spreadCanvas.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, SW, PH);
    }

    const title = document.querySelector('.topbar-title')?.textContent || 'Untitled Zine';
    pdf.save(title.replace(/[^a-z0-9]/gi, '_') + '.pdf');

    // Restore editor selection state
    state.selected = savedSelected;
    state.multiSelected = savedMultiSelected;
    state.activePage = savedActivePage;
    state.imagePositionMode = savedPositionMode;

    if (btnPDF) { btnPDF.innerHTML = origHTML; btnPDF.disabled = false; }
  } catch (err) {
    console.error('PDF export failed:', err);
    alert('PDF export failed: ' + (err.message || err));
    // Restore editor selection state on error too
    state.selected = savedSelected;
    state.multiSelected = savedMultiSelected;
    state.activePage = savedActivePage;
    state.imagePositionMode = savedPositionMode;
    if (btnPDF) { btnPDF.innerHTML = origHTML; btnPDF.disabled = false; }
  }
}

// ============ PASSWORD & INIT ============
function checkPassword() {
  if (document.getElementById('pwInput').value === PASSWORD) {
    startBuilder();
  } else {
    document.getElementById('pwError').style.display = 'block';
  }
}

function startBuilder() {
  if (window.__BUILDER_UI_STARTED__) return;
  window.__BUILDER_UI_STARTED__ = true;
  document.getElementById('gate').classList.add('hidden');
  document.getElementById('app').classList.remove('hidden');
  sessionStorage.setItem('oz_auth', '1');
  init();
  initIcons();
  initDeleteSpreadDelegate();
  builder.emit('ready');
  bootstrapPersistence();
}

document.getElementById('pwInput').addEventListener('keypress', e => {
  if (e.key === 'Enter') checkPassword();
});


// Check if coming from platform with token (bypass password gate)
function hasValidPlatformToken() {
  const params = new URLSearchParams(window.location.search);
  return params.get('issueId') && params.get('token');
}

// Auto-start: either from platform token OR previous session
if (hasValidPlatformToken() || sessionStorage.getItem('oz_auth') === '1') {
  document.addEventListener('DOMContentLoaded', () => { 
    startBuilder();
  });
  if (document.readyState !== 'loading') { 
    startBuilder();
  }
}

// ============================================================================
// API ADAPTER - Handles ALL backend communication (isolated from editor)
// ============================================================================

const apiAdapter = {
  tokens: {
    access: null,
    refresh: null
  },

  issueId: null,
  zineId: null,
  version: null,

  // ============ SESSION LIFECYCLE ============
  // Proactive refresh + sleep/wake handling, per Pete & Clive's guardrails:
  //   A: single refresh in flight (already via _refreshPromise below)
  //   B: clean degradation — once terminal, stop retrying
  // Modal only fires on terminal auth failure (refresh endpoint 401), not on
  // network blips, 5xx, or ordinary 401s that successfully re-auth.
  _refreshTimer: null,
  _accessExpiresAt: null,   // ms since epoch
  _sessionExpired: false,   // terminal auth-failure flag
  _lifecycleInitialized: false,
  _onSessionExpired: null,  // UI hook set by caller; fired once on terminal failure

  sessionExpired() { return this._sessionExpired; },

  // Decode the JWT exp claim (ms) without verifying signature — we trust our own token.
  _decodeAccessExpiry(accessToken) {
    try {
      const payload = JSON.parse(atob(accessToken.split('.')[1].replace(/-/g, '+').replace(/_/g, '/')));
      return payload && typeof payload.exp === 'number' ? payload.exp * 1000 : null;
    } catch (e) {
      return null;
    }
  },

  _scheduleProactiveRefresh() {
    this._cancelProactiveRefresh();
    if (!this.tokens.access) return;
    const expMs = this._decodeAccessExpiry(this.tokens.access);
    this._accessExpiresAt = expMs;
    if (!expMs) return;
    // Fire 5 min before expiry; if already past that threshold, fire soon.
    const delay = Math.max(1000, expMs - Date.now() - 5 * 60 * 1000);
    this._refreshTimer = setTimeout(() => {
      this._refreshTimer = null;
      // Don't await; the single-flight mutex handles any overlap.
      this.refreshAccessToken();
    }, delay);
  },

  _cancelProactiveRefresh() {
    if (this._refreshTimer) {
      clearTimeout(this._refreshTimer);
      this._refreshTimer = null;
    }
  },

  // On tab wake: timers may be stale, so recompute from actual expiry.
  async _handleVisibilityChange() {
    if (document.visibilityState !== 'visible') return;
    if (this._sessionExpired || !this.tokens.access) return;
    const expMs = this._accessExpiresAt || this._decodeAccessExpiry(this.tokens.access);
    if (!expMs) return;
    // If within 5 min of expiry (or already expired), refresh now.
    if (Date.now() >= expMs - 5 * 60 * 1000) {
      await this.refreshAccessToken();
    } else {
      // Re-arm the timer in case it was lost during sleep.
      this._scheduleProactiveRefresh();
    }
  },

  _markSessionExpired() {
    if (this._sessionExpired) return;
    this._sessionExpired = true;
    this._cancelProactiveRefresh();
    try { if (typeof this._onSessionExpired === 'function') this._onSessionExpired(); } catch (e) {}
  },

  initLifecycle() {
    if (this._lifecycleInitialized) return;
    this._lifecycleInitialized = true;
    try {
      document.addEventListener('visibilitychange', () => this._handleVisibilityChange());
    } catch (e) {}
  },
  
  // Build API URLs for this issue (flat routes â backend doesn't use nested zine routes)
  issueUrl() {
    return `${API_BASE}/api/publisher/issues/${this.issueId}`;
  },
  
  saveUrl() {
    return `${API_BASE}/api/publisher/issues/${this.issueId}/save`;
  },
  
  publishUrl() {
    return `${API_BASE}/api/publisher/issues/${this.issueId}/publish`;
  },
  
  setTokens(access, refresh) {
    this.tokens.access = access;
    this.tokens.refresh = refresh;
    // Fresh tokens mean we're authenticated again — clear any terminal-expired state.
    this._sessionExpired = false;
    // Persist to sessionStorage so tokens survive page refresh within the tab
    try {
      sessionStorage.setItem('oz_builder_tokens', JSON.stringify({ access, refresh }));
    } catch (e) { /* sessionStorage unavailable */ }
    this.initLifecycle();
    this._scheduleProactiveRefresh();
  },
  
  loadTokensFromSession() {
    try {
      const stored = JSON.parse(sessionStorage.getItem('oz_builder_tokens'));
      if (stored && stored.access) {
        this.tokens.access = stored.access;
        this.tokens.refresh = stored.refresh;
        this.initLifecycle();
        this._scheduleProactiveRefresh();
        return true;
      }
    } catch (e) {}
    return false;
  },
  
  clearSessionTokens() {
    try { sessionStorage.removeItem('oz_builder_tokens'); } catch (e) {}
    try { sessionStorage.removeItem('oz_builder_issueId'); } catch (e) {}
    try { sessionStorage.removeItem('oz_builder_zineId'); } catch (e) {}
  },
  
  // Mutex: only one refresh in flight at a time. All concurrent callers await the same promise.
  _refreshPromise: null,

  async refreshAccessToken() {
    if (!this.tokens.refresh) return false;
    if (this._sessionExpired) return false; // Guardrail B: don't retry-storm after terminal failure
    // If a refresh is already in flight, piggyback on it (Guardrail A)
    if (this._refreshPromise) return this._refreshPromise;

    this._refreshPromise = (async () => {
      try {
        const res = await fetch(`${API_BASE}/api/auth/refresh`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ refreshToken: this.tokens.refresh })
        });
        if (!res.ok) {
          // 401 from the refresh endpoint = refresh token is terminally dead.
          // Anything else (5xx, timeout) = transient; let the next fetch try again.
          if (res.status === 401) {
            this._markSessionExpired();
          }
          return false;
        }
        const data = await res.json();
        this.tokens.access = data.tokens?.accessToken || data.accessToken;
        if (data.tokens?.refreshToken) this.tokens.refresh = data.tokens.refreshToken;
        // Update sessionStorage with refreshed tokens
        try {
          sessionStorage.setItem('oz_builder_tokens', JSON.stringify({ access: this.tokens.access, refresh: this.tokens.refresh }));
        } catch (e) {}
        // Reschedule the proactive refresh for the new token's lifetime
        this._scheduleProactiveRefresh();
        return true;
      } catch (e) {
        // Network error / fetch threw — transient, not terminal. Let next call try again.
        console.error('Token refresh failed:', e);
        return false;
      } finally {
        this._refreshPromise = null;
      }
    })();

    return this._refreshPromise;
  },

  async fetch(url, options = {}) {
    if (!this.tokens.access) return null;
    // Guardrail B: if session is already terminally expired, don't attempt further requests.
    if (this._sessionExpired) return null;

    // Clone options to avoid mutating the caller's object on retry
    const makeHeaders = () => ({
      ...options.headers,
      'Authorization': `Bearer ${this.tokens.access}`,
      'Content-Type': 'application/json'
    });

    let res = await fetch(url, { ...options, headers: makeHeaders() });

    // If 401, try refreshing token once
    if (res.status === 401) {
      const refreshed = await this.refreshAccessToken();
      if (refreshed) {
        // Retry with fresh token, preserving original signal if present
        res = await fetch(url, { ...options, headers: makeHeaders() });
      }
    }

    return res;
  },
  
  async load(issueId) {
    try {
      const res = await this.fetch(this.issueUrl());
      if (!res || !res.ok) {
        console.error('Failed to load issue, status:', res?.status);
        return null;
      }
      const data = await res.json();
      
      // Backend returns { issue, zine, builderState }
      // builderState has: { version, project: { name }, pages, roles }
      const bs = data.builderState;
      
      if (bs && bs.pages && bs.pages.length > 0) {
        // Paper name â hex color mapping (matches backend PAPER_COLORS)
        const PAPER_HEX = {
          cotton: '#fdfbf7', cream: '#f8f4e8', bright: '#ffffff',
          kraft: '#d4c4a8', newsprint: '#f0ebe0', blush: '#fdf2f0',
          sage: '#e8ede5', sky: '#e8f1f8'
        };
        
        // Map backend page format to builder format
        const pages = bs.pages.map((p, i) => ({
          // First page must be 'cover' â builder uses this ID for single-page rendering
          id: i === 0 ? 'cover' : (p.id || `p${i}`),
          name: p.name || (i === 0 ? 'Cover' : `Page ${i}`),
          section: p.section,
          // Convert paper name to hex if needed (builder uses hex for CSS background)
          paper: p.paper?.startsWith('#') ? p.paper : (PAPER_HEX[p.paper] || '#fdfbf7'),
          texture: p.texture,
          elements: (p.elements || []).map(el => {
            // Backend stores full element data in block.data, which includes the original builder format
            // If it has 't' field, it's already in builder format; if 'type', needs mapping
            const mapped = el.t ? el : { ...el, t: el.type || el.t };
            // Normalize legacy R2 URLs to custom domain
            if (mapped.src) mapped.src = normalizeAssetUrl(mapped.src);
            if (mapped.revealMask) mapped.revealMask = normalizeAssetUrl(mapped.revealMask);
            return mapped;
          })
        }));
        
        // Safety: ensure at least cover + 6 pages (3 spreads). Pad if API returned fewer.
        if (pages.length > 0) {
          // Ensure first page is always 'cover'
          pages[0].id = 'cover';
          pages[0].name = pages[0].name || 'Cover';
          // Pad to minimum 7 pages (cover + 6)
          const minPages = 7;
          while (pages.length < minPages) {
            const idx = pages.length;
            pages.push({
              id: `p${idx}`,
              name: `Page ${idx}`,
              paper: '#fdfbf7',
              texture: undefined,
              elements: []
            });
          }
        }
        
        return {
          state: {
            pages: pages,
            title: bs.project?.name || data.issue?.title || 'Untitled Zine'
          },
          version: bs.version || '13.1',
          publishedAt: data.issue?.publishedAt || null,
          issueNumber: data.issue?.issueNumber || null,
          zine: data.zine || null,
          isPublished: data.issue?.isPublished || false,
          hasUnpublishedChanges: data.issue?.hasUnpublishedChanges || false,
          updatedAt: data.issue?.updatedAt || null, // used to compare against any preserved local draft
        };
      }
      
      // No pages yet â return null so builder starts fresh
      return null;
    } catch (e) {
      console.error('Load error:', e);
      return null;
    }
  },
  
  async save(issueId, builderState) {
    try {
      // Backend expects: PUT /issues/:id/save with { builderState: { version, project, pages, roles } }
      // Hex â paper name mapping (reverse of load)
      
      const builderStatePayload = {
        version: '13.1',
        project: {
          name: builderState.title || 'Untitled Zine',
          description: ''
        },
        pages: builderState.pages.map((p, i) => ({
          id: p.id === 'cover' ? 'cover' : (p.id || `p${i + 1}`),
          name: p.name || (i === 0 ? 'Cover' : `Page ${i + 1}`),
          section: p.section || (i === 0 ? 'cover' : 'editorial'),
          paper: HEX_TO_PAPER[p.paper] || p.paper || 'cotton',
          texture: p.texture || 'smooth',
          deckled: p.deckled || false,
          elements: (p.elements || []).map(el => ({
            ...el,
            type: el.t || el.type || 'unknown',
          }))
        })),
        roles: {
          display: { f: 'Bebas Neue', s: 72 },
          header: { f: 'Playfair Display', s: 36 },
          subhead: { f: 'Work Sans', s: 18 },
          copy: { f: 'EB Garamond', s: 14 },
          caption: { f: 'Inter', s: 10 },
          scrawl: { f: 'Homemade Apple', s: 16 },
        }
      };
      
      const res = await this.fetch(this.saveUrl(), {
        method: 'PUT',
        body: JSON.stringify({ builderState: builderStatePayload })
      });
      if (!res || !res.ok) {
        const errData = await res.json().catch(() => ({}));
        console.error('Save failed:', res?.status, errData);
        return false;
      }
      const data = await res.json();
      console.log('Save success:', data.issue?.id);
      return true;
    } catch (e) {
      console.error('Save error:', e);
      return false;
    }
  }
};

// Local storage adapter for standalone mode
const localAdapter = {
  load(key) {
    try {
      const data = localStorage.getItem(`oz_${key}`);
      return data ? JSON.parse(data) : null;
    } catch (e) {
      return null;
    }
  },

  save(key, state) {
    try {
      localStorage.setItem(`oz_${key}`, JSON.stringify(state));
      return true;
    } catch (e) {
      return false;
    }
  }
};

// ============================================================================
// DRAFT STORE — Pete & Clive's "preserve draft FIRST, then decide UI" layer.
// Writes the builder state to localStorage keyed by issueId so the work survives
// any auth failure, tab crash, or accidental close. Cleared on successful save
// or publish so we never resurrect zombie drafts.
// ============================================================================
const draftStore = {
  _key(issueId) { return `oz_draft_${issueId}`; },

  preserve(issueId, state) {
    if (!issueId || !state) return false;
    try {
      localStorage.setItem(this._key(issueId), JSON.stringify({
        savedAt: Date.now(),
        state
      }));
      return true;
    } catch (e) {
      // QuotaExceeded on mobile or with very large drafts — not fatal.
      console.warn('Draft preservation failed:', e);
      return false;
    }
  },

  load(issueId) {
    if (!issueId) return null;
    try {
      const raw = localStorage.getItem(this._key(issueId));
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null;
    }
  },

  clear(issueId) {
    if (!issueId) return;
    try { localStorage.removeItem(this._key(issueId)); } catch (e) {}
  }
};

// ============================================================================
// EVENT BRIDGE - Connects editor to persistence (glue layer)
// ============================================================================
let persistenceState = {
  mode: 'standalone', // 'standalone' or 'api'
  issueId: null,
  isDirty: false,
  isSaving: false,
  isLoading: false,
  saveTimer: null,
  statusEl: null,
  _dirtyDuringSave: false,  // tracks if new edits arrived while a save was in flight
  isPublished: false,
  hasUnpublishedChanges: false,
};

function updateSaveStatus(status) {
  let el = document.getElementById('saveStatus');
  if (!el) {
    // Create status element in topbar
    const topbar = document.querySelector('.topbar-actions');
    if (topbar) {
      el = document.createElement('span');
      el.id = 'saveStatus';
      el.style.cssText = 'font-size:10px;color:#666;margin-right:8px;';
      topbar.insertBefore(el, topbar.firstChild);
    }
  }
  if (el) el.textContent = status;
}

// ============================================================================
// SESSION-EXPIRED MODAL — shown ONLY on terminal auth failure (refresh token
// dead), NOT on network blips, 5xx, or ordinary 401s that re-auth cleanly.
// Draft is already preserved by the caller; this is just the UI surface.
// ============================================================================
let _sessionExpiredShown = false;
function showSessionExpiredModal() {
  if (_sessionExpiredShown) return; // idempotent — no modal spam
  _sessionExpiredShown = true;
  const overlay = document.getElementById('sessionExpiredOverlay');
  if (!overlay) {
    // Fallback for the rare case the markup isn't in the page (defensive).
    alert('Your session expired — log in again to save your work. Your changes are preserved locally.');
    window.location.href = 'https://onlyzines.com/';
    return;
  }
  overlay.classList.remove('hidden');
  const loginBtn = document.getElementById('sessionExpiredLoginBtn');
  const cancelBtn = document.getElementById('sessionExpiredCancelBtn');
  if (loginBtn) loginBtn.onclick = () => { window.location.href = 'https://onlyzines.com/'; };
  if (cancelBtn) cancelBtn.onclick = () => { overlay.classList.add('hidden'); _sessionExpiredShown = false; };
}

// ============================================================================
// DRAFT RESTORE TOAST — on builder load, if we find a preserved draft for the
// current issue that is newer than the server's last-saved timestamp, offer
// to restore it. No toast if draft is older than server (server is fresher).
// ============================================================================
function offerDraftRestore(issueId, serverUpdatedAtMs) {
  const draft = draftStore.load(issueId);
  if (!draft || !draft.state) return;
  // If server state is strictly newer, discard the draft silently.
  if (typeof serverUpdatedAtMs === 'number' && draft.savedAt <= serverUpdatedAtMs) {
    draftStore.clear(issueId);
    return;
  }
  const toast = document.getElementById('draftRestoreToast');
  if (!toast) return;
  toast.classList.remove('hidden');
  const restoreBtn = document.getElementById('draftRestoreYesBtn');
  const discardBtn = document.getElementById('draftRestoreNoBtn');
  if (restoreBtn) restoreBtn.onclick = () => {
    try { builder.loadState(draft.state); } catch (e) { console.error('Draft restore failed:', e); }
    draftStore.clear(issueId);
    toast.classList.add('hidden');
    // Mark dirty so the restored state gets pushed to backend on next save.
    persistenceState.isDirty = true;
    updateSaveStatus('Restored — save to keep');
  };
  if (discardBtn) discardBtn.onclick = () => {
    draftStore.clear(issueId);
    toast.classList.add('hidden');
  };
}

// Scan all elements for data URLs and upload to R2 before save.
// Returns number of replacements made. Mutates the state in place.
async function uploadPendingDataUrls(builderState) {
  const uploads = [];
  for (const page of (builderState.pages || [])) {
    for (const el of (page.elements || [])) {
      if (el.src && typeof el.src === 'string' && el.src.startsWith('data:')) {
        uploads.push(uploadImageToR2(el.src).then(url => {
          if (url && url !== el.src) el.src = url;
        }));
      }
      if (el.dataUrl && typeof el.dataUrl === 'string' && el.dataUrl.startsWith('data:')) {
        uploads.push(uploadImageToR2(el.dataUrl).then(url => {
          if (url && url !== el.dataUrl) el.dataUrl = url;
        }));
      }
      if (el.revealMask && typeof el.revealMask === 'string' && el.revealMask.startsWith('data:')) {
        uploads.push(uploadImageToR2(el.revealMask).then(url => {
          if (url && url !== el.revealMask) el.revealMask = url;
        }));
      }
      if (el.originalSrc && typeof el.originalSrc === 'string' && el.originalSrc.startsWith('data:')) {
        uploads.push(uploadImageToR2(el.originalSrc).then(url => {
          if (url && url !== el.originalSrc) el.originalSrc = url;
        }));
      }
    }
  }
  if (uploads.length > 0) {
    await Promise.all(uploads);
  }
  return uploads.length;
}

async function saveNow() {
  if (persistenceState.isSaving || persistenceState.isLoading) return;
  if (!persistenceState.isDirty) return;

  persistenceState.isSaving = true;
  updateSaveStatus('Saving...');

  const builderState = builder.getState();

  // Safety net: upload any remaining data URLs to R2 before save
  const pendingCount = await uploadPendingDataUrls(builderState);
  if (pendingCount > 0) {
    console.log(`Uploaded ${pendingCount} pending data URL(s) to R2 before save`);
  }

  // Visibility: warn if any data URLs survived upload attempts
  const _remaining = [];
  for (const page of (builderState.pages || [])) {
    for (const el of (page.elements || [])) {
      if (el.src && typeof el.src === 'string' && el.src.startsWith('data:')) _remaining.push({ id: el.id, field: 'src' });
      if (el.dataUrl && typeof el.dataUrl === 'string' && el.dataUrl.startsWith('data:')) _remaining.push({ id: el.id, field: 'dataUrl' });
      if (el.revealMask && typeof el.revealMask === 'string' && el.revealMask.startsWith('data:')) _remaining.push({ id: el.id, field: 'revealMask' });
      if (el.originalSrc && typeof el.originalSrc === 'string' && el.originalSrc.startsWith('data:')) _remaining.push({ id: el.id, field: 'originalSrc' });
    }
  }
  if (_remaining.length > 0) {
    const sample = _remaining.slice(0, 2).map(r => `${r.id}.${r.field}`).join(', ');
    console.warn(`[upload] ${_remaining.length} data URL(s) remain after upload attempts (${sample})`);
  }

  let success = false;

  if (persistenceState.mode === 'api' && persistenceState.issueId) {
    success = await apiAdapter.save(persistenceState.issueId, builderState);
    // Layered draft preservation:
    //   success → clear any stale draft so we never resurrect zombies on next load
    //   failure WITH terminal auth expiry → stash the draft & surface the modal
    if (success) {
      draftStore.clear(persistenceState.issueId);
    } else if (apiAdapter.sessionExpired()) {
      draftStore.preserve(persistenceState.issueId, builderState);
      showSessionExpiredModal();
    }
  } else {
    success = localAdapter.save('draft', builderState);
  }

  persistenceState.isSaving = false;
  // Only clear dirty flag if save succeeded. If new edits arrived during the save
  // (isDirty was re-set by scheduleAutosave), preserve the dirty flag.
  if (success && !persistenceState._dirtyDuringSave) {
    persistenceState.isDirty = false;
  }
  persistenceState._dirtyDuringSave = false;

  updateSaveStatus(success ? 'Saved' : 'Save failed');

  if (success) {
    // If this is a published issue, mark that the live version is now behind
    if (persistenceState.isPublished) {
      persistenceState.hasUnpublishedChanges = true;
      updatePublishBar();
    }
    // Clear "Saved" after 2 seconds
    setTimeout(() => {
      if (!persistenceState.isDirty) updateSaveStatus('');
    }, 2000);
  }
}

function scheduleAutosave() {
  if (persistenceState.isLoading) return; // Don't save while loading

  persistenceState.isDirty = true;
  // Track if edits arrive while a save is in flight
  if (persistenceState.isSaving) persistenceState._dirtyDuringSave = true;
  updateSaveStatus('Unsaved changes');
  
  clearTimeout(persistenceState.saveTimer);
  persistenceState.saveTimer = setTimeout(saveNow, 1500);
}

async function bootstrapPersistence() {
  // Wire the apiAdapter's session-expiry hook to the UI.
  // When a refresh ultimately fails (401 from /api/auth/refresh), the adapter
  // calls this once; draft is preserved in saveNow/publishIssue error paths.
  apiAdapter._onSessionExpired = showSessionExpiredModal;

  // Parse URL for API mode
  const params = new URLSearchParams(window.location.search);
  const issueId = params.get('issueId');
  const zineId = params.get('zineId');
  const token = params.get('token');
  const refreshToken = params.get('refreshToken');
  
  if (issueId && token) {
    // API mode from URL params - clean URL
    window.history.replaceState({}, '', window.location.pathname);
    
    persistenceState.mode = 'api';
    persistenceState.issueId = issueId;
    persistenceState.zineId = zineId;
    apiAdapter.setTokens(token, refreshToken);
    apiAdapter.issueId = issueId;
    apiAdapter.zineId = zineId;
    
    // Also persist issueId/zineId to sessionStorage for page refresh
    try {
      sessionStorage.setItem('oz_builder_issueId', issueId);
      sessionStorage.setItem('oz_builder_zineId', zineId || '');
    } catch (e) {}
    
  } else if (apiAdapter.loadTokensFromSession()) {
    // Page was refreshed â restore from sessionStorage
    const storedIssueId = sessionStorage.getItem('oz_builder_issueId');
    const storedZineId = sessionStorage.getItem('oz_builder_zineId');
    
    if (storedIssueId) {
      persistenceState.mode = 'api';
      persistenceState.issueId = storedIssueId;
      persistenceState.zineId = storedZineId;
      apiAdapter.issueId = storedIssueId;
      apiAdapter.zineId = storedZineId;
    }
  }
  
  if (persistenceState.mode === 'api' && persistenceState.issueId) {
    // Load from API
    persistenceState.isLoading = true;
    updateSaveStatus('Loading...');
    
    // Show loading overlay to block interaction
    const overlay = document.getElementById('builderLoadingOverlay');
    if (overlay) overlay.classList.remove('hidden');
    
    const loaded = await apiAdapter.load(persistenceState.issueId);
    
    if (loaded && loaded.state && loaded.state.pages) {
      builder.loadState(loaded.state);
      updateSaveStatus('Loaded');

      // If a local draft was preserved from a prior session (usually because of
      // a session-expiry event), offer to restore it — but only if it's newer
      // than what the server has. This is the other half of the Pete & Clive
      // "preserve first, decide UI later" guardrail.
      try {
        const serverMs = loaded.updatedAt ? new Date(loaded.updatedAt).getTime() : null;
        offerDraftRestore(persistenceState.issueId, serverMs);
      } catch (e) { /* non-critical */ }

      // Track publish state for the Publish/Update button
      apiAdapter.publishedAt = loaded.publishedAt || null;
      apiAdapter.issueNumber = loaded.issueNumber || null;
      apiAdapter.zineSlug = loaded.zine?.slug || null;
      apiAdapter.zineVisibility = loaded.zine?.visibility || 'UNLISTED';
      apiAdapter.publisherHandle = null;
      apiAdapter.publicUrl = null;

      // Draft/live separation state
      persistenceState.isPublished = loaded.isPublished || false;
      persistenceState.hasUnpublishedChanges = loaded.hasUnpublishedChanges || false;
      updatePublishBar();
      
      // Fetch publisher handle for public URL (zine response doesn't include it)
      try {
        const acctRes = await apiAdapter.fetch(`${API_BASE}/api/publisher/account`);
        if (acctRes && acctRes.ok) {
          const acctData = await acctRes.json();
          apiAdapter.publisherHandle = acctData.publisher?.handle || null;
        }
      } catch (e) { /* non-critical */ }
      
      if (apiAdapter.publishedAt && apiAdapter.publisherHandle && apiAdapter.zineSlug && apiAdapter.issueNumber) {
        apiAdapter.publicUrl = 'https://onlyzines.com/' + apiAdapter.publisherHandle + '/' + apiAdapter.zineSlug + '/' + apiAdapter.issueNumber;
      }
    } else {
      updateSaveStatus('New issue');
    }
    
    // Update publish button based on published state
    updatePublishButton();
    
    persistenceState.isLoading = false;
    
    // Hide loading overlay
    const overlayEl = document.getElementById('builderLoadingOverlay');
    if (overlayEl) overlayEl.classList.add('hidden');
    
    // Clear status after delay
    setTimeout(() => updateSaveStatus(''), 2000);
  }
  
  if (persistenceState.mode !== 'api') {
    // Standalone mode - try loading local draft
    persistenceState.mode = 'standalone';
    const localDraft = localAdapter.load('draft');
    if (localDraft && localDraft.pages) {
      builder.loadState(localDraft);
    }
  }
  
  // Start listening for changes
  builder.on('change', scheduleAutosave);
}

function updatePublishButton() {
  const btn = document.getElementById('btnPublish');
  if (!btn) return;

  // Phase B.1: btn now contains <img><span>...</span> from initIcons. Update the span text so
  // we don't wipe the icon. Fallback to textContent if initIcons hasn't run (e.g. very early load).
  const span = btn.querySelector('span');
  const label = apiAdapter.publishedAt ? 'Update' : 'Publish';
  if (span) {
    span.textContent = label;
  } else {
    btn.textContent = label;
  }
  btn.title = apiAdapter.publishedAt ? 'Update your published zine' : 'Publish this issue';
}

// Phase B: Print modal. Reuses .publish-modal-overlay/.publish-modal container from showPublishModal.
// "Print flat" is a thin wrapper around the existing savePDF() — no forked export logic.
// "Print & fold (True Zine)" is explicitly disabled until Phase C ships the fold imposition.
function openPrintModal() {
  const overlay = document.createElement('div');
  overlay.className = 'publish-modal-overlay';
  overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };
  overlay.innerHTML = `
    <div class="publish-modal">
      <h3>Print</h3>
      <div class="print-modal-options">
        <button class="print-modal-option" id="printFlatBtn" type="button">
          <div class="print-modal-option-title">Print flat</div>
          <div class="print-modal-option-desc">Printed as you see it.</div>
        </button>
        <button class="print-modal-option" id="printFoldBtn" type="button" disabled aria-disabled="true" title="Coming soon — fold imposition lands in the next update">
          <div class="print-modal-option-title">Print &amp; fold (True Zine)</div>
          <div class="print-modal-option-desc">Re-ordered for you to fold into a zine.</div>
          <span class="print-modal-option-soon">Coming soon</span>
        </button>
      </div>
      <button class="publish-modal-close" onclick="this.closest('.publish-modal-overlay').remove()">Done</button>
    </div>
  `;
  document.body.appendChild(overlay);
  // Print flat → close modal, then call existing savePDF() unchanged.
  const flatBtn = overlay.querySelector('#printFlatBtn');
  flatBtn.addEventListener('click', () => {
    overlay.remove();
    savePDF();
  });
}

function getPublicUrl() {
  if (apiAdapter.publicUrl) return apiAdapter.publicUrl;
  const handle = apiAdapter.publisherHandle;
  const slug = apiAdapter.zineSlug;
  const num = apiAdapter.issueNumber;
  if (handle && slug && num) {
    return `https://onlyzines.com/${handle}/${slug}/${num}`;
  }
  return null;
}

function showPublishModal(title, message, url) {
  const overlay = document.createElement('div');
  overlay.className = 'publish-modal-overlay';
  overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };

  let urlHtml = '';
  if (url) {
    urlHtml = `
      <div class="publish-modal-url">
        <input type="text" value="${url}" readonly id="publishUrlInput" onclick="this.select()">
        <button onclick="navigator.clipboard.writeText('${url}').then(()=>{this.textContent='Copied!';setTimeout(()=>this.textContent='Copy',1500)})">Copy</button>
      </div>
    `;
  }

  // Visibility indicator — only show on successful publish (when we have a URL)
  let visibilityHtml = '';
  if (url && apiAdapter.zineId) {
    const vis = apiAdapter.zineVisibility || 'UNLISTED';
    const isPublic = vis === 'PUBLIC';
    const statusText = isPublic
      ? 'Public — anyone can read this zine'
      : 'Private — collectors must request access to read';
    const toggleLabel = isPublic ? 'Make Private' : 'Make Public';
    visibilityHtml = `
      <div class="publish-modal-visibility">
        <span class="visibility-dot ${isPublic ? 'public' : 'unlisted'}"></span>
        <span class="visibility-text" id="visibilityStatus">${statusText}</span>
      </div>
      <button class="publish-modal-toggle" id="visibilityToggle" onclick="toggleZineVisibility(this)">${toggleLabel}</button>
    `;
  }

  overlay.innerHTML = `
    <div class="publish-modal">
      <h3>${title}</h3>
      <p>${message}</p>
      ${urlHtml}
      ${visibilityHtml}
      <button class="publish-modal-close" onclick="this.closest('.publish-modal-overlay').remove()">Done</button>
    </div>
  `;
  document.body.appendChild(overlay);
}

async function toggleZineVisibility(btn) {
  if (!apiAdapter.zineId) return;
  const current = apiAdapter.zineVisibility || 'UNLISTED';
  const next = current === 'PUBLIC' ? 'UNLISTED' : 'PUBLIC';

  btn.disabled = true;
  btn.textContent = 'Updating...';

  try {
    const res = await apiAdapter.fetch(`${API_BASE}/api/publisher/zines/${apiAdapter.zineId}`, {
      method: 'PATCH',
      body: JSON.stringify({ visibility: next }),
    });
    if (!res || !res.ok) throw new Error('Failed to update visibility');

    apiAdapter.zineVisibility = next;
    const isPublic = next === 'PUBLIC';
    const dot = document.querySelector('.visibility-dot');
    const status = document.getElementById('visibilityStatus');
    if (dot) { dot.className = 'visibility-dot ' + (isPublic ? 'public' : 'unlisted'); }
    if (status) {
      status.textContent = isPublic
        ? 'Public — anyone can read this zine'
        : 'Private — collectors must request access to read';
    }
    btn.textContent = isPublic ? 'Make Private' : 'Make Public';
    btn.disabled = false;
  } catch (e) {
    console.error('Visibility toggle error:', e);
    btn.textContent = 'Failed — try again';
    btn.disabled = false;
    setTimeout(() => {
      const isPublic = (apiAdapter.zineVisibility || 'UNLISTED') === 'PUBLIC';
      btn.textContent = isPublic ? 'Make Private' : 'Make Public';
    }, 2000);
  }
}

async function publishIssue() {
  if (persistenceState.mode !== 'api' || !persistenceState.issueId) {
    alert('Please save your zine first. Publishing is only available when editing from the platform.');
    return;
  }
  
  const btn = document.getElementById('btnPublish');
  const originalText = btn.textContent;
  const isAlreadyPublished = !!apiAdapter.publishedAt;
  
  btn.textContent = isAlreadyPublished ? 'Saving...' : 'Publishing...';
  btn.disabled = true;
  
  try {
    // Always save pending changes first
    if (persistenceState.isDirty) {
      await saveNow();
      if (persistenceState.isDirty) {
        throw new Error('Could not save changes. Please check your connection and try again.');
      }
    }
    
    // Capture cover snapshot and upload to R2
    // Capture cover image and upload to R2
    btn.textContent = 'Capturing cover...';
    let coverImageUrl = null;
    try {
      const coverDataUrl = await captureCoverImage();
      if (coverDataUrl) {
        btn.textContent = 'Uploading cover...';
        const uploaded = await uploadImageToR2(coverDataUrl);
        // Only use it if upload succeeded (returned a URL, not a data: fallback)
        if (uploaded && !uploaded.startsWith('data:')) {
          coverImageUrl = uploaded;
        }
      }
    } catch (e) {
      console.warn('Cover capture failed, publishing without cover:', e);
    }

    if (isAlreadyPublished) {
      // Already published â push a new snapshot to live
      const res = await apiAdapter.fetch(apiAdapter.publishUrl(), {
        method: 'POST',
        body: JSON.stringify({ coverImageUrl })
      });

      if (!res || !res.ok) {
        let errMsg = 'Failed to update live issue';
        try {
          const errData = await res.json();
          errMsg = errData?.error?.message || errData?.message || errMsg;
        } catch (e) {}
        throw new Error(errMsg);
      }

      const data = await res.json();
      const url = data.url ? 'https://onlyzines.com' + data.url : getPublicUrl();
      if (url) apiAdapter.publicUrl = url;

      persistenceState.hasUnpublishedChanges = false;
      updatePublishBar();

      showPublishModal('Live Updated â', 'Your changes are now live.', url || null);
      btn.textContent = 'Update';
      btn.disabled = false;
    } else {
      // First publish â call the publish endpoint
      const res = await apiAdapter.fetch(apiAdapter.publishUrl(), {
        method: 'POST',
        body: JSON.stringify({ coverImageUrl })
      });
      
      if (!res || !res.ok) {
        let errMsg = 'Failed to publish';
        try {
          const errData = await res.json();
          errMsg = errData?.error?.message || errData?.message || errData?.error || errMsg;
        } catch (parseErr) {
          try { errMsg = await res.text() || errMsg; } catch(e) {}
        }
        throw new Error(errMsg);
      }
      
      const data = await res.json();
      apiAdapter.publishedAt = data.issue?.publishedAt || new Date().toISOString();

      const url = data.url ? 'https://onlyzines.com' + data.url : getPublicUrl();
      if (url) apiAdapter.publicUrl = url;

      // Issue is now published and in sync
      persistenceState.isPublished = true;
      persistenceState.hasUnpublishedChanges = false;
      updatePublishBar();

      showPublishModal('Published!', 'Your zine is now live.', url);

      updatePublishButton();
      btn.disabled = false;
    }
    
  } catch (err) {
    console.error('Publish error:', err);
    // Preserve-then-decide per Pete & Clive's guardrail.
    if (apiAdapter.sessionExpired()) {
      // Terminal auth failure — stash the draft locally before any UI, then surface
      // the session-expired modal instead of the misleading "Publish Failed" dialog.
      try {
        const liveState = builder.getState();
        draftStore.preserve(persistenceState.issueId, liveState);
      } catch (e) { /* preservation is best-effort */ }
      showSessionExpiredModal();
    } else {
      // Non-auth failure (network blip, 5xx, validation, etc.) — keep the existing
      // generic modal. Those errors are genuinely "publish failed."
      showPublishModal('Publish Failed', err.message, null);
    }
    btn.textContent = originalText;
    btn.disabled = false;
  }
}

function updatePublishBar() {
  const bar = document.getElementById('publishBar');
  const label = document.getElementById('publishBarLabel');
  const btn = document.getElementById('publishBarBtn');
  if (!bar) return;
  const visible = persistenceState.isPublished && persistenceState.hasUnpublishedChanges;
  bar.classList.toggle('visible', visible);
  // Reset to default state if we're re-showing
  if (visible && label) label.textContent = 'Unpublished changes';
  if (visible && btn) { btn.textContent = 'Update Live Issue'; btn.disabled = false; btn.style.display = ''; }
}

async function updateLiveIssue() {
  if (persistenceState.mode !== 'api' || !persistenceState.issueId) return;

  const label = document.getElementById('publishBarLabel');
  const btn = document.getElementById('publishBarBtn');
  if (btn) { btn.disabled = true; btn.textContent = 'Publishingâ¦'; }

  try {
    // Save any pending draft changes first
    if (persistenceState.isDirty) {
      await saveNow();
      if (persistenceState.isDirty) throw new Error('Could not save draft â check your connection.');
    }

    // Capture cover image and upload to R2
    if (btn) btn.textContent = 'Capturing cover...';
    let coverImageUrl = null;
    try {
      const coverDataUrl = await captureCoverImage();
      if (coverDataUrl) {
        if (btn) btn.textContent = 'Uploading cover...';
        const uploaded = await uploadImageToR2(coverDataUrl);
        if (uploaded && !uploaded.startsWith('data:')) {
          coverImageUrl = uploaded;
        }
      }
    } catch (e) {
      console.warn('Cover capture failed, updating without cover:', e);
    }

    const res = await apiAdapter.fetch(apiAdapter.publishUrl(), {
      method: 'POST',
      body: JSON.stringify(coverImageUrl ? { coverImageUrl } : {})
    });

    if (!res || !res.ok) {
      let errMsg = 'Update failed';
      try { const d = await res.json(); errMsg = d?.error?.message || errMsg; } catch (e) {}
      throw new Error(errMsg);
    }

    const data = await res.json();
    const url = data.url ? 'https://onlyzines.com' + data.url : getPublicUrl();
    if (url) apiAdapter.publicUrl = url;

    persistenceState.hasUnpublishedChanges = false;
    if (label) label.textContent = 'Live issue updated â';
    if (btn) btn.style.display = 'none';
    setTimeout(updatePublishBar, 2500);
  } catch (err) {
    if (label) label.textContent = err.message || 'Update failed â try again';
    if (btn) { btn.disabled = false; btn.textContent = 'Retry'; }
  }
}

// Expose for debugging
window.builder = builder;
window.apiAdapter = apiAdapter;
