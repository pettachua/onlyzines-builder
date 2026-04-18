// ============================================================================
// BUILDER FACADE - The ONLY way external code can interact with the editor
// API code can ONLY use these methods. No direct DOM or state manipulation.
// ============================================================================
const builder = {
  _listeners: {},
  
  // Get serialized state for saving
  getState() {
    const _gs0 = performance.now();
    const _gsSwaps = _extractBlobs(state.pages);
    const pages = JSON.parse(JSON.stringify(state.pages));
    _restoreBlobs(_gsSwaps);
    const _gs1 = performance.now();
    // Filter out elements that are entirely off-canvas (outside page bounds)
    pages.forEach(page => {
      page.elements = (page.elements || []).filter(el => {
        // Keep element if any part of it is within the page bounds
        return el.x < PAGE_W && el.y < PAGE_H && (el.x + el.w) > 0 && (el.y + el.h) > 0;
      });
    });
    // Restore blob refs → real data URLs in the cloned copy so save payload has actual data
    _rehydrateBlobs(pages);
    const _gs2 = performance.now();
    if ((_gs2 - _gs0) > 5) {
      console.warn(`[getState perf] serialize=${(_gs1-_gs0).toFixed(1)}ms  rehydrate=${(_gs2-_gs1).toFixed(1)}ms  total=${(_gs2-_gs0).toFixed(1)}ms`);
    }
    return {
      pages,
      title: document.querySelector('.topbar-title')?.textContent || 'Untitled Zine'
    };
  },
  
  // Load state atomically (safe for API use)
  loadState(newState, skipRender = false) {
    if (!newState || !newState.pages) return false;
    try {
      state.pages = JSON.parse(JSON.stringify(newState.pages));
      
      // Existing data — preserve exact page count from save. Do NOT pad.
      // (New-zine defaults are set in init(), not here.)
      if (state.pages.length > 0) {
        state.pages[0].id = 'cover';
        state.pages[0].name = 'Cover';
      }
      // Renumber all pages so display names are always sequential
      for (let i = 1; i < state.pages.length; i++) {
        state.pages[i].name = 'Page ' + i;
      }
      
      state.selected = null;
      state.multiSelected = [];
      state.imagePositionMode = null;
      // Migrate legacy crop/zoom properties to new image-position model
      for (const page of state.pages) {
        for (const el of (page.elements || [])) {
          if (el.t === 'image' && el.imgW !== undefined) {
            el.imgOffsetX = el.cropX || 0;
            el.imgOffsetY = el.cropY || 0;
            delete el.cropX; delete el.cropY; delete el.cropScale;
            delete el.imgW; delete el.imgH;
          }
        }
      }
      // Capture any data URL blobs so history snapshots can use ref keys
      _populateBlobStore(state.pages);
      // Reset history baseline to the loaded state (no saveHistory — no change event)
      const _baseSwaps = _extractBlobs(state.pages);
      const _baseSnap = JSON.stringify(state.pages);
      _restoreBlobs(_baseSwaps);
      state.history = [_baseSnap];
      state.historyIndex = 0;
      _historyBytes = _baseSnap.length * 2;
      // Load fonts used in this zine on demand
      loadFontsForPages(state.pages);
      if (newState.title) {
        const titleEl = document.querySelector('.topbar-title');
        if (titleEl) titleEl.textContent = newState.title;
      }
      if (!skipRender) {
        // Don't call saveHistory here to avoid triggering autosave during load
        render();
      }
      return true;
    } catch (e) {
      console.error('Failed to load state:', e);
      return false;
    }
  },
  
  // Event system
  on(event, handler) {
    if (!this._listeners[event]) this._listeners[event] = [];
    this._listeners[event].push(handler);
  },
  
  off(event, handler) {
    if (!this._listeners[event]) return;
    this._listeners[event] = this._listeners[event].filter(h => h !== handler);
  },
  
  emit(event, data) {
    if (!this._listeners[event]) return;
    this._listeners[event].forEach(h => h(data));
  },
  
  // Status display
  setStatus(msg) {
    console.log('[Builder Status]', msg);
  }
};


// ============ INITIALIZATION ============
function init() {
  // Double-init guard - prevents corruption from accidental double initialization
  if (window.__BUILDER_STARTED__) return;
  window.__BUILDER_STARTED__ = true;
  
  state.pages = [
    {id:'cover', name:'Cover', paper:'#fdfbf7', elements:[]},
    {id:'p1', name:'Page 1', paper:'#fdfbf7', elements:[]},
    {id:'p2', name:'Page 2', paper:'#fdfbf7', elements:[]},
    {id:'p3', name:'Page 3', paper:'#fdfbf7', elements:[]},
    {id:'p4', name:'Page 4', paper:'#fdfbf7', elements:[]},
    {id:'p5', name:'Page 5', paper:'#fdfbf7', elements:[]},
    {id:'p6', name:'Page 6', paper:'#fdfbf7', elements:[]}
  ];
  saveHistory();
  render();
}

function generateId() {
  // Use crypto.randomUUID when available for collision-free IDs, with fallback
  if (typeof crypto !== 'undefined' && crypto.randomUUID) {
    return 'el_' + crypto.randomUUID().replace(/-/g, '').substring(0, 12);
  }
  return 'el_' + Date.now() + '_' + Math.random().toString(36).substr(2,6);
}

function getPage(id) {
  return state.pages.find(p => p.id === (id || state.currentPage));
}

function getActivePage() {
  return state.pages.find(p => p.id === state.activePage) || getPage();
}

function getSelectedElement() {
  if (!state.selected) return null;
  for (const page of state.pages) {
    const el = page.elements.find(e => e.id === state.selected);
    if (el) return el;
  }
  return null;
}

function findElementById(id) {
  if (!id) return null;
  for (const page of state.pages) {
    const el = page.elements.find(e => e.id === id);
    if (el) return el;
  }
  return null;
}

function countWords(text) {
  if (!text) return 0;
  return text.trim().split(/\s+/).filter(w => w).length;
}

function generateLoremIpsum(wordCount) {
  let result = '';
  while (countWords(result) < wordCount) {
    result += LOREM + ' ';
  }
  return result.split(/\s+/).slice(0, wordCount).join(' ');
}


// ============ HISTORY ============

// --- Blob externalization helpers (Phase 3A: el.src only) ---

function _blobKey(dataUrl) {
  const start = dataUrl.indexOf(',') + 1;
  const mid = start + Math.floor((dataUrl.length - start) / 2);
  return 'blob_' + dataUrl.substring(mid, mid + 16) + '_' + dataUrl.length;
}

function _extractBlobs(pages) {
  const swaps = [];
  for (const p of pages) {
    for (const el of (p.elements || [])) {
      if (el.src && typeof el.src === 'string' && el.src.startsWith('data:')) {
        const key = _blobKey(el.src);
        _blobStore.set(key, el.src);
        swaps.push({ el, dataUrl: el.src });
        el.src = key;
      }
    }
  }
  return swaps;
}

function _restoreBlobs(swaps) {
  for (const s of swaps) { s.el.src = s.dataUrl; }
}

function _rehydrateBlobs(pages) {
  for (const p of pages) {
    for (const el of (p.elements || [])) {
      if (el.src && typeof el.src === 'string' && el.src.startsWith('blob_')) {
        const real = _blobStore.get(el.src);
        if (real) {
          el.src = real;
        } else {
          console.warn('[blobStore] missing ref:', el.src, '— element', el.id);
        }
      }
    }
  }
}

function _populateBlobStore(pages) {
  for (const p of pages) {
    for (const el of (p.elements || [])) {
      if (el.src && typeof el.src === 'string' && el.src.startsWith('data:')) {
        _blobStore.set(_blobKey(el.src), el.src);
      }
    }
  }
}

// Prune _blobStore entries that are no longer referenced by any history snapshot.
// Called periodically (every 10 saves) to prevent unbounded growth.
let _blobPruneCounter = 0;
function _pruneBlobStore() {
  if (_blobStore.size === 0) return;
  _blobPruneCounter++;
  if (_blobPruneCounter < 10) return; // only prune every 10 saves
  _blobPruneCounter = 0;
  // Collect all blob refs that appear in current history
  const liveRefs = new Set();
  for (const snapshot of state.history) {
    // Scan for blob_XXXX_YYYY keys (fast string search, no parse needed)
    let idx = 0;
    while ((idx = snapshot.indexOf('blob_', idx)) !== -1) {
      const end = snapshot.indexOf('"', idx);
      if (end !== -1) liveRefs.add(snapshot.substring(idx, end));
      idx = end !== -1 ? end : idx + 5;
    }
  }
  // Also keep refs for current live state data URLs
  for (const p of state.pages) {
    for (const el of (p.elements || [])) {
      if (el.src && typeof el.src === 'string' && el.src.startsWith('data:')) {
        liveRefs.add(_blobKey(el.src));
      }
    }
  }
  let pruned = 0;
  for (const key of _blobStore.keys()) {
    if (!liveRefs.has(key)) { _blobStore.delete(key); pruned++; }
  }
  if (pruned > 0) console.log(`[blobStore] pruned ${pruned} stale entries, ${_blobStore.size} remain`);
}

// --- End blob externalization helpers ---

function saveHistory() {
  const _sh0 = performance.now();
  const _swaps = _extractBlobs(state.pages);
  const snapshot = JSON.stringify(state.pages);
  _restoreBlobs(_swaps);
  const _sh1 = performance.now();
  const snapshotSize = snapshot.length * 2; // rough byte estimate (UTF-16)
  const _shDur = _sh1 - _sh0;
  if (_shDur > 5) {
    let _dataUrlCount = 0, _elCount = 0;
    for (const p of state.pages) { for (const el of (p.elements||[])) { _elCount++; if (el.src && el.src.startsWith('data:')) _dataUrlCount++; } }
    console.warn(`[saveHistory perf] stringify=${_shDur.toFixed(1)}ms  size=${(snapshot.length/1024).toFixed(0)}KB  pages=${state.pages.length}  els=${_elCount}  dataUrls=${_dataUrlCount}  historyLen=${state.history.length}`);
  }

  if (state.historyIndex < state.history.length - 1) {
    // Recalculate bytes for trimmed entries
    const removed = state.history.slice(state.historyIndex + 1);
    removed.forEach(s => { _historyBytes -= s.length * 2; });
    state.history = state.history.slice(0, state.historyIndex + 1);
  }

  state.history.push(snapshot);
  _historyBytes += snapshotSize;

  // Enforce max entries (reduce cap if snapshots are large)
  const maxEntries = snapshotSize > 2 * 1024 * 1024 ? 20 : 50;
  while (state.history.length > maxEntries) {
    const removed = state.history.shift();
    _historyBytes -= removed.length * 2;
  }

  // Enforce total memory budget
  while (_historyBytes > HISTORY_MAX_BYTES && state.history.length > 1) {
    const removed = state.history.shift();
    _historyBytes -= removed.length * 2;
  }

  state.historyIndex = state.history.length - 1;

  // Periodically prune orphaned blob store entries to prevent memory growth
  _pruneBlobStore();

  // Emit change event for persistence layer
  if (typeof builder !== 'undefined' && builder.emit) {
    builder.emit('change');
  }
}

function undo() {
  if (state.historyIndex > 0) {
    state.historyIndex--;
    state.pages = JSON.parse(state.history[state.historyIndex]);
    _rehydrateBlobs(state.pages);
    _stateGeneration++;
    state.selected = null;
    state.selectedSpread = null;
    // currentPage may point to a page that no longer exists after undo — fall back to cover
    if (!state.pages.find(p => p.id === state.currentPage)) {
      state.currentPage = 'cover';
      state.activePage = 'cover';
    }
    _needsPagesPanelUpdate = true;
    render();
  }
}

function redo() {
  if (state.historyIndex < state.history.length - 1) {
    state.historyIndex++;
    state.pages = JSON.parse(state.history[state.historyIndex]);
    _rehydrateBlobs(state.pages);
    _stateGeneration++;
    state.selected = null;
    state.selectedSpread = null;
    // currentPage may point to a page that no longer exists after redo — fall back to cover
    if (!state.pages.find(p => p.id === state.currentPage)) {
      state.currentPage = 'cover';
      state.activePage = 'cover';
    }
    _needsPagesPanelUpdate = true;
    render();
  }
}

function setLayoutStyle(id) {
  state.layoutStyle = id;
  renderPagesPanel();
}

function selectPage(id) {
  state.currentPage = id;
  state.activePage = id;
  state.selected = null;
  render();
}

function applyLayout(id) {
  const _al0 = performance.now();
  const layout = LAYOUTS.find(l => l.id === id);
  if (!layout) return;
  const page = getActivePage();
  if (!page) return;

  saveHistory();
  const _al1 = performance.now();

  // Collect existing content from the page
  const existingImages = page.elements.filter(el => el.t === 'image' && el.src);
  const existingText = page.elements.filter(el => el.t === 'text' && el.txt && el.txt.trim());

  // Use pre-computed set from constants.js (avoids rebuilding on every call)
  function isPlaceholderText(txt) {
    if (!txt || !txt.trim()) return true;
    const lower = txt.toLowerCase().trim();
    if (_ALL_LAYOUT_DEFAULTS.has(lower)) return true;
    if (lower.startsWith('lorem ipsum')) return true;
    if (lower.startsWith('lorem:')) return true;
    return false;
  }
  
  // If page is empty (no real content), just apply layout as-is with empty placeholders
  const hasContent = existingImages.length > 0 || existingText.some(t => !isPlaceholderText(t.txt));
  
  if (!hasContent) {
    // No user content — apply layout fresh (original behavior)
    page.elements = layout.elements.map((el, idx) => ({
      ...el,
      id: generateId(),
      z: idx + 1
    }));
    _needsPagesPanelUpdate = true;
    const _al2 = performance.now();
    render();
    const _al3 = performance.now();
    if ((_al3 - _al0) > 16) console.warn(`[applyLayout perf] total=${(_al3-_al0).toFixed(1)}ms  saveHistory=${(_al1-_al0).toFixed(1)}ms  logic=${(_al2-_al1).toFixed(1)}ms  render=${(_al3-_al2).toFixed(1)}ms  layout=${id}  hasContent=false`);
    return;
  }
  
  // Smart redistribution: match existing content to layout slots
  const imagePool = [...existingImages]; // images with actual src data
  const textPool = existingText.filter(t => !isPlaceholderText(t.txt)); // only user-edited text
  
  const newElements = [];
  let zIndex = 1;
  
  layout.elements.forEach(slot => {
    const newEl = {
      ...slot,
      id: generateId(),
      z: zIndex++
    };
    
    if (slot.t === 'image' && imagePool.length > 0) {
      // Fill image slot with existing image content
      const img = imagePool.shift();
      newEl.src = img.src;
      newEl.fitMode = img.fitMode || 'cover';
      newEl.flipH = img.flipH;
      newEl.flipV = img.flipV;
      newEl.opacity = img.opacity;
      newEl.deepEtched = img.deepEtched;
      newEl.revealMask = img.revealMask;
      // Reset image position for new layout dimensions
      newEl.imgOffsetX = 0;
      newEl.imgOffsetY = 0;
    } else if (slot.t === 'text' && textPool.length > 0) {
      // Fill text slot with existing text content
      const txt = textPool.shift();
      newEl.txt = txt.txt;
      // Keep the layout's font/size/style for visual consistency
      // but preserve user's text content
      // If user had custom colors, keep them
      if (txt.color && txt.color !== '#000000' && txt.color !== '#111111' && txt.color !== '#333333' && txt.color !== '#555555' && txt.color !== '#666666') {
        newEl.color = txt.color;
      }
    }
    // Shapes and other types: use layout defaults
    
    newElements.push(newEl);
  });
  
  // Handle overflow: remaining images that didn't fit into slots
  const overflowX = 280;
  let overflowY = 20;
  
  imagePool.forEach(img => {
    newElements.push({
      ...img,
      id: generateId(),
      x: overflowX,
      y: overflowY,
      w: Math.min(img.w || 100, 100),
      h: Math.min(img.h || 100, 100),
      z: zIndex++
    });
    overflowY += Math.min(img.h || 100, 100) + 10;
    if (overflowY > 550) { overflowY = 20; } // wrap if too tall
  });
  
  // Handle overflow: remaining text that didn't fit
  textPool.forEach(txt => {
    newElements.push({
      ...txt,
      id: generateId(),
      x: overflowX,
      y: overflowY,
      w: Math.min(txt.w || 100, 100),
      h: Math.min(txt.h || 40, 40),
      z: zIndex++
    });
    overflowY += Math.min(txt.h || 40, 40) + 10;
    if (overflowY > 550) { overflowY = 20; }
  });
  
  page.elements = newElements;
  _needsPagesPanelUpdate = true;
  const _al2 = performance.now();
  render();
  const _al3 = performance.now();
  if ((_al3 - _al0) > 16) console.warn(`[applyLayout perf] total=${(_al3-_al0).toFixed(1)}ms  saveHistory=${(_al1-_al0).toFixed(1)}ms  logic=${(_al2-_al1).toFixed(1)}ms  render=${(_al3-_al2).toFixed(1)}ms  layout=${id}  hasContent=true  imgs=${existingImages.length}  texts=${existingText.length}`);
}

function addPage() {
  try {
    // 30 spread cap (cover + 60 pages = 61 entries)
    if (state.pages.length >= 61) {
      alert('Maximum of 30 spreads reached. Delete a spread to add more.');
      return;
    }
    saveHistory();
    const num = state.pages.length;
    // Use unique timestamp + random suffix to guarantee unique IDs even if called rapidly
    const ts = Date.now();
    const rand = Math.floor(Math.random() * 10000);
    const newId = 'p' + ts + '_' + rand;
    const newId2 = 'p' + (ts + 1) + '_' + rand;
    // Always add a spread (pair of pages)
    state.pages.push({
      id: newId,
      name: 'Page ' + num,
      section: 'editorial',
      paper: '#fdfbf7',
      elements: []
    });
    state.pages.push({
      id: newId2,
      name: 'Page ' + (num + 1),
      section: 'editorial',
      paper: '#fdfbf7',
      elements: []
    });
    // Navigate to the newly added spread and select it in the sidebar
    state.currentPage = newId;
    state.activePage = newId;
    const newSpreadIdx = Math.floor((state.pages.length - 2) / 2);
    state.selectedSpread = newSpreadIdx;
    _needsPagesPanelUpdate = true;
    render();
    // Scroll sidebar to the new spread and flash it
    requestAnimationFrame(() => {
      const thumbs = document.querySelectorAll('.thumb-spread');
      const newThumb = thumbs[newSpreadIdx];
      if (newThumb) {
        newThumb.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        newThumb.classList.add('thumb-spread-new');
        setTimeout(() => newThumb.classList.remove('thumb-spread-new'), 700);
      }
    });
  } catch (err) {
    console.error('[addPage] ERROR:', err);
    alert('Error adding page: ' + err.message);
  }
}

function deletePage(pageId) {
  // Don't delete the cover
  if (pageId === 'cover') return;
  
  // Find the page index
  const idx = state.pages.findIndex(p => p.id === pageId);
  if (idx < 1) return; // Don't delete if not found or if it's the cover (index 0)
  
  // Must have at least 2 pages (cover + 1 page)
  if (state.pages.length <= 2) {
    alert('Cannot delete the last page. Your zine must have at least one page.');
    return;
  }
  
  saveHistory();
  
  // Remove the page
  state.pages.splice(idx, 1);
  
  // Renumber remaining pages
  for (let i = 1; i < state.pages.length; i++) {
    state.pages[i].name = 'Page ' + i;
  }
  
  // Update current page if we deleted the active one
  if (state.currentPage === pageId || state.activePage === pageId) {
    // Select the previous page, or the first page if we deleted page 1
    const newIdx = Math.max(1, idx - 1);
    if (state.pages[newIdx]) {
      state.currentPage = state.pages[newIdx].id;
      state.activePage = state.pages[newIdx].id;
    } else {
      state.currentPage = state.pages[1].id;
      state.activePage = state.pages[1].id;
    }
  }

  _needsPagesPanelUpdate = true;
  render();
}

// ============ DELETE-SPREAD VALIDATION ============
// Single source of truth: returns { allowed:true } or { allowed:false, reason:string }
function canDeleteSpread(spreadIdx) {
  const totalSpreads = Math.floor((state.pages.length - 1) / 2);
  const leftIdx = spreadIdx * 2 + 1;

  // Must keep at least 1 spread (cover + 2 pages)
  if (totalSpreads <= 1) {
    return { allowed: false, reason: 'Cannot delete the last spread.' };
  }
  // Spread index must point to a real page
  if (leftIdx < 1 || leftIdx >= state.pages.length) {
    return { allowed: false, reason: 'Invalid spread index.' };
  }
  // Defensive: if pages array has corrupt even length, heal it before allowing delete
  if ((state.pages.length - 1) % 2 !== 0) {
    console.warn('[canDeleteSpread] Corrupt page count (' + state.pages.length + '), healing with empty page');
    state.pages.push({ id: 'p_heal_' + Date.now(), name: 'Page ' + state.pages.length, section: 'editorial', paper: '#fdfbf7', elements: [] });
  }
  return { allowed: true };
}

function deleteSpread(spreadIdx) {
  console.log('[deleteSpread] called with spreadIdx:', spreadIdx);
  try {
    const check = canDeleteSpread(spreadIdx);
    if (!check.allowed) {
      console.log('[deleteSpread] BLOCKED —', check.reason);
      alert(check.reason);
      return;
    }

    const leftIdx = spreadIdx * 2 + 1;
    const rightIdx = leftIdx + 1;
    console.log('[deleteSpread] guard passed, proceeding with delete');

    saveHistory();

    // Remove both pages (right first so indices don't shift)
    if (rightIdx < state.pages.length) {
      state.pages.splice(rightIdx, 1);
    }
    if (leftIdx < state.pages.length) {
      state.pages.splice(leftIdx, 1);
    }

    // Renumber display names only — IDs stay stable
    for (let i = 1; i < state.pages.length; i++) {
      state.pages[i].name = 'Page ' + i;
    }

    // Stay on nearest remaining spread
    const newLeftIdx = Math.min(leftIdx, state.pages.length - 1);
    const validIdx = Math.max(1, newLeftIdx);
    if (state.pages[validIdx]) {
      state.currentPage = state.pages[validIdx].id;
      state.activePage = state.pages[validIdx].id;
    } else {
      state.currentPage = 'cover';
      state.activePage = 'cover';
    }
    state.selectedSpread = null;
    state.selected = null;

    _needsPagesPanelUpdate = true;
    render();
    scheduleAutosave();
  } catch (err) {
    console.error('[deleteSpread] ERROR:', err);
    alert('Error deleting spread: ' + err.message);
  }
}

// ============ SPREAD SELECTION & REORDERING ============

function selectSpread(spreadIdx, addToSelection = false) {
  // Single selection only (shift for multi-select not implemented yet)
  state.selectedSpread = spreadIdx;
  
  // Also select the first page of this spread
  const pageIdx = spreadIdx * 2 + 1;
  if (state.pages[pageIdx]) {
    state.currentPage = state.pages[pageIdx].id;
    state.activePage = state.pages[pageIdx].id;
  }
  state.selected = null;
  render();
}

function moveSpread(fromSpreadIdx, toSpreadIdx, dropBelow) {
  saveHistory();
  
  // Get the pages from the dragged spread
  const fromPageIdx = fromSpreadIdx * 2 + 1;
  const leftPage = state.pages[fromPageIdx];
  const rightPage = state.pages[fromPageIdx + 1];
  
  if (!leftPage) return;
  
  // Remove the dragged pages
  const pagesToMove = rightPage ? [leftPage, rightPage] : [leftPage];
  state.pages.splice(fromPageIdx, pagesToMove.length);
  
  // Calculate new insert position
  let toPageIdx = toSpreadIdx * 2 + 1;
  if (fromSpreadIdx < toSpreadIdx) {
    // If we removed pages before the target, adjust the index
    toPageIdx -= pagesToMove.length;
  }
  if (dropBelow) {
    toPageIdx += 2;
  }
  
  // Insert at new position
  state.pages.splice(toPageIdx, 0, ...pagesToMove);
  
  // Renumber all pages
  for (let i = 1; i < state.pages.length; i++) {
    state.pages[i].name = 'Page ' + i;
  }
  
  // Update selected spread
  state.selectedSpread = Math.floor((toPageIdx - 1) / 2);

  _needsPagesPanelUpdate = true;
  render();
}

// Old HTML5 drag API stubs removed — spread reordering now uses mousedown/mousemove/mouseup

// ============ SPREAD COPY/PASTE ============
function copySpread() {
  if (state.selectedSpread === null) return;
  
  const pageIdx = state.selectedSpread * 2 + 1;
  const leftPage = state.pages[pageIdx];
  const rightPage = state.pages[pageIdx + 1];
  
  state.spreadClipboard = {
    left: leftPage ? JSON.parse(JSON.stringify(leftPage)) : null,
    right: rightPage ? JSON.parse(JSON.stringify(rightPage)) : null
  };
  state.lastCopied = 'spread';
}

function pasteSpread() {
  if (!state.spreadClipboard) return;
  
  saveHistory();
  
  // Insert after currently selected spread (or at end if none selected)
  let insertIdx;
  if (state.selectedSpread !== null) {
    insertIdx = (state.selectedSpread + 1) * 2 + 1;
  } else {
    insertIdx = state.pages.length;
  }
  
  const newPages = [];
  
  if (state.spreadClipboard.left) {
    newPages.push({
      ...state.spreadClipboard.left,
      id: 'p' + Date.now(),
      elements: state.spreadClipboard.left.elements.map(el => ({
        ...el,
        id: generateId()
      }))
    });
  }
  
  if (state.spreadClipboard.right) {
    newPages.push({
      ...state.spreadClipboard.right,
      id: 'p' + (Date.now() + 1),
      elements: state.spreadClipboard.right.elements.map(el => ({
        ...el,
        id: generateId()
      }))
    });
  }
  
  // Insert the new pages
  state.pages.splice(insertIdx, 0, ...newPages);
  
  // Renumber all pages
  for (let i = 1; i < state.pages.length; i++) {
    state.pages[i].name = 'Page ' + i;
  }
  
  // Select the newly pasted spread
  state.selectedSpread = Math.floor((insertIdx - 1) / 2);
  if (state.pages[insertIdx]) {
    state.currentPage = state.pages[insertIdx].id;
    state.activePage = state.pages[insertIdx].id;
  }

  _needsPagesPanelUpdate = true;
  render();
}

function handlePageDrop(e, pageId) {
  e.preventDefault();
  
  // Clear any drop target highlights
  document.querySelectorAll('.drop-target').forEach(el => el.classList.remove('drop-target'));
  
  // Handle sticker drops
  const stickerData = e.dataTransfer?.getData('sticker');
  if (stickerData) {
    const page = getPage(pageId);
    if (!page) return;
    const pageEl = e.target.closest('.page') || document.querySelector(`.page[data-page="${pageId}"]`);
    if (!pageEl) return;
    const rect = pageEl.getBoundingClientRect();
    const dropX = (e.clientX - rect.left) - 60;
    const dropY = (e.clientY - rect.top) - 60;
    saveHistory();
    page.elements.push({
      id: generateId(), t: 'image',
      x: Math.max(0, dropX), y: Math.max(0, dropY),
      w: 120, h: 120,
      src: `stickers/${stickerData}`, sticker: true, fitMode: 'contain',
      z: page.elements.length + 1
    });
    _needsPagesPanelUpdate = true;
    render();
    scheduleAutosave();
    return;
  }
  
  const file = e.dataTransfer?.files?.[0];
  if (!file?.type?.startsWith('image/')) return;
  
  const page = getPage(pageId);
  if (!page) return;
  
  // Check if dropped on an empty image placeholder
  // First try direct target, then use elementsFromPoint as fallback for spread mode
  let dropTarget = e.target.closest('.element.image.empty');
  if (!dropTarget) {
    const elementsAtPoint = document.elementsFromPoint(e.clientX, e.clientY);
    dropTarget = elementsAtPoint.find(el => el.classList?.contains('element') && el.classList?.contains('image') && el.classList?.contains('empty')) || null;
  }
  
  if (dropTarget) {
    // Snap into existing placeholder
    const elId = dropTarget.dataset.id;
    // In spread mode, the element may belong to a different page than the one passed in
    const targetPageId = dropTarget.dataset.page || pageId;
    const targetPage = getPage(targetPageId) || page;
    const el = targetPage.elements.find(x => x.id === elId);
    if (el) {
      saveHistory();
      const reader = new FileReader();
      reader.onload = async ev => {
        const compressed = await compressImage(ev.target.result);
        // Display immediately with local data URL
        el.src = compressed;
        el.imgOffsetX = 0;
        el.imgOffsetY = 0;
        // Locked composition: compute cover-fit inner image coordinates
        const tmpImg = new Image();
        tmpImg.onload = () => {
          const cs = Math.max(el.w / tmpImg.naturalWidth, el.h / tmpImg.naturalHeight);
          const iw = Math.round(tmpImg.naturalWidth * cs);
          const ih = Math.round(tmpImg.naturalHeight * cs);
          el.innerX = Math.round((el.w - iw) / 2);
          el.innerY = Math.round((el.h - ih) / 2);
          el.innerW = iw;
          el.innerH = ih;
          _needsPagesPanelUpdate = true;
          render();
          // Upload to R2 in background, swap when done
          const gen = _stateGeneration;
          uploadImageToR2(compressed).then(function(r2Url) {
            if (_stateGeneration !== gen) return;
            if (r2Url && r2Url !== compressed && el.src === compressed) {
              el.src = r2Url;
              var imgEl = document.querySelector('[data-id="' + el.id + '"] img');
              if (imgEl) imgEl.src = r2Url;
              scheduleAutosave();
            }
          });
        };
        tmpImg.src = compressed;
      };
      reader.readAsDataURL(file);
      return;
    }
  }
  
  // No placeholder - create new image element
  const pageEl = e.target.closest('.page') || document.querySelector(`.page[data-page="${pageId}"]`);
  if (!pageEl) return;
  const rect = pageEl.getBoundingClientRect();
  const dropX = e.clientX - rect.left;
  const dropY = e.clientY - rect.top;
  
  saveHistory();
  const reader = new FileReader();
  reader.onload = async ev => {
    const compressed = await compressImage(ev.target.result);
    const img = new Image();
    img.onload = async () => {
      let w = img.naturalWidth;
      let h = img.naturalHeight;

      const maxSize = Math.min(300, rect.width - 20, rect.height - 20);
      if (w > maxSize || h > maxSize) {
        const scale = maxSize / Math.max(w, h);
        w = Math.round(w * scale);
        h = Math.round(h * scale);
      }

      const x = Math.max(0, Math.min(dropX - w/2, rect.width - w));
      const y = Math.max(0, Math.min(dropY - h/2, rect.height - h));

      // Locked composition: compute cover-fit inner image coordinates
      const coverScale = Math.max(w / img.naturalWidth, h / img.naturalHeight);
      const iw = Math.round(img.naturalWidth * coverScale);
      const ih = Math.round(img.naturalHeight * coverScale);
      const newEl = {
        id: generateId(),
        t: 'image',
        x: x,
        y: y,
        w: w,
        h: h,
        src: compressed,
        innerX: Math.round((w - iw) / 2),
        innerY: Math.round((h - ih) / 2),
        innerW: iw,
        innerH: ih,
        z: page.elements.length + 1
      };
      page.elements.push(newEl);
      _needsPagesPanelUpdate = true;
      render();
      // Upload to R2 in background, swap when done
      const gen = _stateGeneration;
      uploadImageToR2(compressed).then(function(r2Url) {
        if (_stateGeneration !== gen) return;
        if (r2Url && r2Url !== compressed && newEl.src === compressed) {
          newEl.src = r2Url;
          var imgEl = document.querySelector('[data-id="' + newEl.id + '"] img');
          if (imgEl) imgEl.src = r2Url;
          scheduleAutosave();
        }
      });
    };
    img.src = compressed;
  };
  reader.readAsDataURL(file);
}

function uploadToElement(elId, pageId) {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/*';
  input.onchange = e => {
    const file = e.target.files[0];
    if (!file) return;
    const page = getPage(pageId);
    const el = page?.elements?.find(x => x.id === elId);
    if (!el) return;
    saveHistory();
    const reader = new FileReader();
    reader.onload = async ev => {
      const compressed = await compressImage(ev.target.result);
      // Display immediately with local data URL
      el.src = compressed;
      _needsPagesPanelUpdate = true;
      render();
      // Upload to R2 in background, swap when done
      const gen = _stateGeneration;
      uploadImageToR2(compressed).then(function(r2Url) {
        if (_stateGeneration !== gen) return;
        if (r2Url && r2Url !== compressed && el.src === compressed) {
          el.src = r2Url;
          var imgEl = document.querySelector('[data-id="' + el.id + '"] img');
          if (imgEl) imgEl.src = r2Url;
          scheduleAutosave();
        }
      });
    };
    reader.readAsDataURL(file);
  };
  input.click();
}

// ============ DRAG & RESIZE ============
function startDrag(e, elId, pageId) {
  const page = getPage(pageId);
  const el = page?.elements?.find(x => x.id === elId);
  if (!el) return;
  
  dragState = {
    active: true,
    type: 'move',
    element: el,
    pageId: pageId,
    startX: e.clientX,
    startY: e.clientY,
    origX: el.x,
    origY: el.y,
    historySaved: false,
    hasMoved: false
  };
}

function startGroupDrag(e, pageId) {
  const page = getPage(pageId);
  if (!page) return;
  
  // Get all selected elements with their original positions
  const elements = page.elements.filter(el => state.multiSelected.includes(el.id));
  if (elements.length === 0) return;
  
  dragState = {
    active: true,
    type: 'groupMove',
    elements: elements.map(el => ({
      el: el,
      origX: el.x,
      origY: el.y
    })),
    pageId: pageId,
    startX: e.clientX,
    startY: e.clientY,
    historySaved: false
  };
}

function startResize(e, elId, pageId, corner) {
  const page = getPage(pageId);
  const el = page?.elements?.find(x => x.id === elId);
  if (!el) return;
  
  dragState = {
    active: true,
    type: 'resize',
    corner: corner,
    element: el,
    pageId: pageId,
    startX: e.clientX,
    startY: e.clientY,
    origX: el.x,
    origY: el.y,
    origW: el.w,
    origH: el.h,
    // Locked image composition: store original inner values for proportional scaling
    origInnerX: el.innerX,
    origInnerY: el.innerY,
    origInnerW: el.innerW,
    origInnerH: el.innerH,
    historySaved: false
  };
}

// Crop mode: drag a corner handle to resize the FRAME (crop window).
// The image stays fixed — only the visible area changes.
// For corners that move the frame origin (NW, NE top, SW left),
// innerX/innerY compensate so the image appears stationary.
function startCropDrag(e, el, elId, pageId, corner) {
  const startX = e.clientX;
  const startY = e.clientY;
  const origX = el.x;
  const origY = el.y;
  const origW = el.w;
  const origH = el.h;
  const origIX = el.innerX;
  const origIY = el.innerY;
  let historySaved = false;
  const MIN = 20; // minimum frame dimension

  const div = document.querySelector(`[data-id="${elId}"]`);
  // Get page-relative offset for DPS spreads
  const xOff = parseInt(div?.dataset.xoffset || '0');

  function onMove(ev) {
    ev.preventDefault();
    if (!historySaved) { saveHistory(); historySaved = true; }
    const dx = ev.clientX - startX;
    const dy = ev.clientY - startY;

    if (corner === 'se') {
      // SE: frame grows/shrinks from bottom-right. Origin stays. Image stays.
      el.w = Math.max(MIN, origW + dx);
      el.h = Math.max(MIN, origH + dy);
      // innerX/innerY unchanged — image doesn't move
    } else if (corner === 'sw') {
      // SW: left edge + bottom edge move
      const newW = Math.max(MIN, origW - dx);
      const actualDx = origW - newW; // how much left edge actually moved right
      el.x = origX + actualDx;
      el.w = newW;
      el.h = Math.max(MIN, origH + dy);
      // Compensate: frame moved right by actualDx, so shift image left by same amount
      el.innerX = origIX - actualDx;
    } else if (corner === 'ne') {
      // NE: right edge + top edge move
      const newH = Math.max(MIN, origH - dy);
      const actualDy = origH - newH; // how much top edge actually moved down
      el.w = Math.max(MIN, origW + dx);
      el.y = origY + actualDy;
      el.h = newH;
      // Compensate: frame moved down by actualDy, so shift image up by same amount
      el.innerY = origIY - actualDy;
    } else { // nw
      // NW: left edge + top edge move
      const newW = Math.max(MIN, origW - dx);
      const newH = Math.max(MIN, origH - dy);
      const actualDx = origW - newW;
      const actualDy = origH - newH;
      el.x = origX + actualDx;
      el.y = origY + actualDy;
      el.w = newW;
      el.h = newH;
      // Compensate both axes
      el.innerX = origIX - actualDx;
      el.innerY = origIY - actualDy;
    }

    // Live DOM update — frame position/size + inner image position
    if (div) {
      div.style.left = (el.x + xOff) + 'px';
      div.style.top = el.y + 'px';
      div.style.width = el.w + 'px';
      div.style.height = el.h + 'px';
      const img = div.querySelector('.img-wrap img');
      if (img) {
        img.style.left = el.innerX + 'px';
        img.style.top = el.innerY + 'px';
        // width/height NOT changed — image stays same size
      }
    }
  }

  function onUp() {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    scheduleAutosave();
    render(); // full render to sync handles, thumbnails, etc.
  }

  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

function startRotate(e, elId, pageId) {
  const page = getPage(pageId);
  const el = page?.elements?.find(x => x.id === elId);
  if (!el) return;
  
  const div = document.querySelector(`[data-id="${elId}"]`);
  if (!div) return;
  const rect = div.getBoundingClientRect();
  
  dragState = {
    active: true,
    type: 'rotate',
    element: el,
    pageId: pageId,
    centerX: rect.left + rect.width / 2,
    centerY: rect.top + rect.height / 2,
    historySaved: false
  };
}


document.addEventListener('mousemove', e => {
  if (!dragState.active) return;
  
  // Handle group move
  if (dragState.type === 'groupMove') {
    if (!dragState.historySaved) {
      saveHistory();
      dragState.historySaved = true;
    }
    
    const dx = e.clientX - dragState.startX;
    const dy = e.clientY - dragState.startY;
    
    dragState.elements.forEach(({el, origX, origY}) => {
      el.x = origX + dx;
      el.y = origY + dy;
      const div = document.querySelector(`[data-id="${el.id}"]`);
      if (div) {
        const xOff = parseInt(div.dataset.xoffset || '0');
        div.style.left = (el.x + xOff) + 'px';
        div.style.top = el.y + 'px';
      }
    });
    return;
  }
  
  if (!dragState.element) return;

  // Require minimum movement before actually dragging (prevents accidental micro-drags)
  if (!dragState.hasMoved) {
    const moveDist = Math.abs(e.clientX - dragState.startX) + Math.abs(e.clientY - dragState.startY);
    if (moveDist < 4) return;
    dragState.hasMoved = true;
  }

  // Save history once
  if (!dragState.historySaved) {
    saveHistory();
    dragState.historySaved = true;
  }

  const el = dragState.element;
  const div = document.querySelector(`[data-id="${el.id}"]`);

  if (dragState.type === 'move') {
    el.x = dragState.origX + (e.clientX - dragState.startX);
    el.y = dragState.origY + (e.clientY - dragState.startY);
    if (div) {
      const xOff = parseInt(div.dataset.xoffset || '0');
      div.style.left = (el.x + xOff) + 'px';
      div.style.top = el.y + 'px';
    }
  } else if (dragState.type === 'resize') {
    const dx = e.clientX - dragState.startX;
    const dy = e.clientY - dragState.startY;
    const aspect = dragState.origW / dragState.origH;
    const shiftHeld = e.shiftKey;
    
    if (dragState.corner === 'se') {
      el.w = Math.max(20, dragState.origW + dx);
      el.h = Math.max(20, dragState.origH + dy);
      if (shiftHeld) el.h = el.w / aspect;
    } else if (dragState.corner === 'sw') {
      el.w = Math.max(20, dragState.origW - dx);
      el.x = dragState.origX + dragState.origW - el.w;
      el.h = Math.max(20, dragState.origH + dy);
      if (shiftHeld) el.h = el.w / aspect;
    } else if (dragState.corner === 'ne') {
      el.w = Math.max(20, dragState.origW + dx);
      el.h = Math.max(20, dragState.origH - dy);
      el.y = dragState.origY + dragState.origH - el.h;
      if (shiftHeld) {
        el.h = el.w / aspect;
        el.y = dragState.origY + dragState.origH - el.h;
      }
    } else if (dragState.corner === 'nw') {
      el.w = Math.max(20, dragState.origW - dx);
      el.x = dragState.origX + dragState.origW - el.w;
      el.h = Math.max(20, dragState.origH - dy);
      el.y = dragState.origY + dragState.origH - el.h;
      if (shiftHeld) {
        el.h = el.w / aspect;
        el.y = dragState.origY + dragState.origH - el.h;
      }
    }

    // Locked composition: scale inner image proportionally with frame
    if (dragState.origInnerW !== undefined) {
      const fx = el.w / dragState.origW;
      const fy = el.h / dragState.origH;
      el.innerX = Math.round(dragState.origInnerX * fx);
      el.innerY = Math.round(dragState.origInnerY * fy);
      el.innerW = Math.round(dragState.origInnerW * fx);
      el.innerH = Math.round(dragState.origInnerH * fy);
    }

    if (div) {
      const xOff = parseInt(div.dataset.xoffset || '0');
      div.style.left = (el.x + xOff) + 'px';
      div.style.top = el.y + 'px';
      div.style.width = el.w + 'px';
      div.style.height = el.h + 'px';
      // Live-update inner image position during resize
      if (el.innerW !== undefined) {
        const innerImg = div.querySelector('.img-wrap img');
        if (innerImg) {
          innerImg.style.left = el.innerX + 'px';
          innerImg.style.top = el.innerY + 'px';
          innerImg.style.width = el.innerW + 'px';
          innerImg.style.height = el.innerH + 'px';
        }
      }
    }
  } else if (dragState.type === 'rotate') {
    const angle = Math.atan2(e.clientY - dragState.centerY, e.clientX - dragState.centerX) * 180 / Math.PI + 90;
    el.rotation = Math.round(angle);
    if (div) div.style.transform = `rotate(${el.rotation}deg)`;
  }
});

document.addEventListener('mouseup', (e) => {
  if (dragState.active) {
    const wasActive = dragState.historySaved;
    const movedElement = dragState.element;
    const wasMoving = dragState.type === 'move';
    let didCrossPage = false;
    
    // DPS cross-page logic: transfer element only if its CENTER lands on another page
    if (wasMoving && movedElement) {
      const currentPageId = dragState.pageId;
      const sourcePageDiv = document.querySelector(`.page[data-page="${currentPageId}"]`);
      
      if (sourcePageDiv) {
        const sourceRect = sourcePageDiv.getBoundingClientRect();
        // Element center in screen coords (using source page rect since el.x is page-relative)
        const elCenterScreenX = sourceRect.left + movedElement.x + movedElement.w / 2;
        const elCenterScreenY = sourceRect.top + movedElement.y + movedElement.h / 2;
        
        // Check all page divs to find which one the center is over
        let targetPageId = null;
        let targetRect = null;
        document.querySelectorAll('.page').forEach(pDiv => {
          const r = pDiv.getBoundingClientRect();
          if (elCenterScreenX >= r.left && elCenterScreenX <= r.right &&
              elCenterScreenY >= r.top && elCenterScreenY <= r.bottom) {
            targetPageId = pDiv.dataset.page;
            targetRect = r;
          }
        });
        
        if (targetPageId && currentPageId && targetPageId !== currentPageId) {
          if (!dragState.historySaved) saveHistory();
          
          const sourcePage = getPage(currentPageId);
          const targetPage = getPage(targetPageId);
          
          if (sourcePage && targetPage) {
            const idx = sourcePage.elements.findIndex(el => el.id === movedElement.id);
            if (idx !== -1) sourcePage.elements.splice(idx, 1);
            
            // Convert page-relative coords from source page to target page.
            // Screen position of element = sourceRect.left + el.x
            // New page-relative position = screenPosition - targetRect.left
            movedElement.x = (sourceRect.left + movedElement.x) - targetRect.left;
            movedElement.y = (sourceRect.top + movedElement.y) - targetRect.top;
            // Clamp to target page bounds so element doesn't land outside visible area
            movedElement.x = Math.max(-movedElement.w + 20, Math.min(PAGE_W - 20, movedElement.x));
            movedElement.y = Math.max(-movedElement.h + 20, Math.min(PAGE_H - 20, movedElement.y));
            
            targetPage.elements.push(movedElement);
            state.activePage = targetPageId;
            state.currentPage = targetPageId;
            state.selected = movedElement.id;
            didCrossPage = true;
          }
        }
      }
    }
    
    const wasResize = dragState.type === 'resize';
    dragState = { active: false };
    if (didCrossPage) _needsPagesPanelUpdate = true;
    if (wasActive || didCrossPage) {
      render();
    }
  }
});


// Helper to find which page contains an element
function findPageContainingElement(elementId) {
  for (const page of state.pages) {
    if (page.elements.find(el => el.id === elementId)) {
      return page.id;
    }
  }
  return null;
}

function addSticker(packId, file) {
  const page = getActivePage();
  if (!page) return;
  saveHistory();
  const src = `stickers/${packId}/${file}`;
  page.elements.push({
    id: generateId(), t: 'image', x: 100, y: 100, w: 120, h: 120,
    src: src, sticker: true, fitMode: 'contain', z: page.elements.length + 1
  });
  _needsPagesPanelUpdate = true;
  render();
}

function addRippedFrame(frameStyleId) {
  const page = getActivePage();
  if (!page) return;
  const frameDef = RIPPED_FRAMES.find(f => f.id === frameStyleId);
  if (!frameDef) return;
  saveHistory();
  const isPolaroid = frameDef.type === 'polaroid';
  const isPhotoFrame = frameDef.type === 'photo-frame';
  const newEl = {
    id: generateId(), t: 'image',
    x: 60, y: 60,
    w: isPolaroid ? 180 : isPhotoFrame ? 180 : 200,
    h: isPolaroid ? 210 : isPhotoFrame ? 220 : 200,
    rippedFrame: frameStyleId, src: null, fitMode: 'cover', z: page.elements.length + 1
  };
  page.elements.push(newEl);
  _needsPagesPanelUpdate = true;
  render();
  scheduleAutosave();
}

function trackColorUsage(color) {
  if (!color || color === 'transparent') return;
  color = color.toLowerCase();
  colorUsage[color] = (colorUsage[color] || 0) + 1;
  updateFrequentColors();
}

function updateFrequentColors() {
  // Get top used colors that aren't already in savedColors
  const sortedColors = Object.entries(colorUsage)
    .sort((a, b) => b[1] - a[1])
    .map(([color]) => color)
    .filter(c => !savedColors.includes(c))
    .slice(0, 4); // Top 4 frequent colors
  
  // Add frequent colors to the beginning of savedColors if used more than 3 times
  sortedColors.forEach(color => {
    if (colorUsage[color] >= 3 && !savedColors.includes(color)) {
      savedColors.unshift(color);
      if (savedColors.length > 12) savedColors.pop();
    }
  });
}


// ============ ELEMENT ACTIONS ============
function addText() {
  const page = getActivePage();
  if (!page) return;
  saveHistory();
  const newId = generateId();
  page.elements.push({
    id: newId, t: 'text', x: 50, y: 100, w: 140, h: 30,
    txt: 'Text', fontFamily: 'Inter,sans-serif', fontSize: 16, color: '#111111', z: page.elements.length + 1
  });
  state.selected = newId;
  _needsPagesPanelUpdate = true;
  render();
  // Auto-enter edit mode so user can immediately type/paste
  requestAnimationFrame(() => {
    const textEl = document.querySelector(`[data-id="${newId}"] .text-content`);
    if (textEl) {
      textEl.contentEditable = 'true';
      textEl.focus();
      // Select all text so user can immediately replace by typing/pasting
      const range = document.createRange();
      range.selectNodeContents(textEl);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }
  });
}

function addTextBox() {
  const page = getActivePage();
  if (!page) return;
  saveHistory();
  const newId = generateId();
  page.elements.push({
    id: newId, t: 'text', x: 30, y: 150, w: 180, h: 200,
    txt: LOREM.substring(0, 350), fontFamily: 'Inter,sans-serif', fontSize: 11, color: '#111111', isBox: true, z: page.elements.length + 1
  });
  state.selected = newId;
  _needsPagesPanelUpdate = true;
  render();
  // Auto-enter edit mode so user can immediately type/paste
  requestAnimationFrame(() => {
    const textEl = document.querySelector(`[data-id="${newId}"] .text-content`);
    if (textEl) {
      textEl.contentEditable = 'true';
      textEl.focus();
      // Select all text so user can immediately replace by typing/pasting
      const range = document.createRange();
      range.selectNodeContents(textEl);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }
  });
}

function addImage() {
  const page = getActivePage();
  if (!page) return;
  saveHistory();
  page.elements.push({
    id: generateId(), t: 'image', x: 50, y: 50, w: 150, h: 150, src: null, z: page.elements.length + 1
  });
  _needsPagesPanelUpdate = true;
  render();
}

function addShape(type) {
  const page = getActivePage();
  if (!page) return;
  saveHistory();
  page.elements.push({
    id: generateId(), t: 'shape', x: 80, y: 80, w: 100, h: 100,
    shape: type, color: '#4A3F2A', z: page.elements.length + 1
  });
  _needsPagesPanelUpdate = true;
  render();
}

function addGraphic(id) {
  const g = GRAPHICS.find(x => x.id === id);
  if (!g) return;
  const page = getActivePage();
  if (!page) return;
  saveHistory();
  page.elements.push({
    id: generateId(), t: 'graphic', x: 100, y: 100, w: g.w, h: g.h, cls: g.cls, z: page.elements.length + 1
  });
  _needsPagesPanelUpdate = true;
  render();
}

function deleteElement() {
  // Handle multi-selection delete
  if (state.multiSelected.length > 0) {
    saveHistory();
    for (const page of state.pages) {
      page.elements = page.elements.filter(e => !state.multiSelected.includes(e.id));
    }
    state.multiSelected = [];
    _needsPagesPanelUpdate = true;
    render();
    return;
  }
  
  // Single selection delete
  const el = getSelectedElement();
  if (!el) return;
  saveHistory();
  for (const page of state.pages) {
    const idx = page.elements.findIndex(e => e.id === state.selected);
    if (idx >= 0) {
      page.elements.splice(idx, 1);
      break;
    }
  }
  state.selected = null;
  _needsPagesPanelUpdate = true;
  render();
}

function uploadImage() {
  const el = getSelectedElement();
  if (!el) return;
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/*';
  input.onchange = e => {
    const file = e.target.files[0];
    if (!file) return;
    saveHistory();
    const reader = new FileReader();
    reader.onload = async ev => {
      const compressed = await compressImage(ev.target.result);
      // Display immediately with local data URL
      el.src = compressed;
      _needsPagesPanelUpdate = true;
      render();
      // Upload to R2 in background, swap when done
      const gen = _stateGeneration;
      uploadImageToR2(compressed).then(function(r2Url) {
        if (_stateGeneration !== gen) return;
        if (r2Url && r2Url !== compressed && el.src === compressed) {
          el.src = r2Url;
          var imgEl = document.querySelector('[data-id="' + el.id + '"] img');
          if (imgEl) imgEl.src = r2Url;
          scheduleAutosave();
        }
      });
    };
    reader.readAsDataURL(file);
  };
  input.click();
}

// Shared debounce guard for color picker slider interactions
let _colorHistoryTimer = null;
let _colorHistorySaved = false;
function _colorHistoryGate() {
  if (!_colorHistorySaved) { saveHistory(); _colorHistorySaved = true; }
  clearTimeout(_colorHistoryTimer);
  _colorHistoryTimer = setTimeout(() => { _colorHistorySaved = false; }, 600);
}

// Debounce guard for rapid font browsing (clicking through font list)
let _fontHistoryTimer = null;
let _fontHistorySaved = false;
function _fontHistoryGate() {
  if (!_fontHistorySaved) { saveHistory(); _fontHistorySaved = true; }
  clearTimeout(_fontHistoryTimer);
  _fontHistoryTimer = setTimeout(() => { _fontHistorySaved = false; }, 800);
}

function setPageColor(c) {
  const page = getActivePage();
  if (!page) return;
  _colorHistoryGate();
  page.paper = c;
  trackColorUsage(c);
  // During color picker drag, update DOM directly instead of full render()
  if (window._colorPickerActive) {
    const pageDiv = document.querySelector(`.page[data-page="${page.id}"]`);
    if (pageDiv) pageDiv.style.background = c;
    window._pageColorDirty = true;
    return;
  }
  render();
}

function setPageTexture(texture) {
  const page = getActivePage();
  if (page) {
    saveHistory();
    page.texture = texture;
    _needsPagesPanelUpdate = true;
    render();
  }
}

// Text tools
let _fontRenderTimer = null;
function setFont(f) {
  ensureFontLoaded(f);
  const el = getSelectedElement();
  if (el && el.t === 'text') {
    // Apply to selected text element — debounced for rapid font browsing
    _fontHistoryGate();
    el.fontFamily = f;
    // Update DOM directly instead of full renderStage()
    const textDiv = document.querySelector(`[data-id="${el.id}"] .text-content`);
    if (textDiv) textDiv.style.fontFamily = f;
    updateFontListActive(f);
    // Defer single render() until browsing settles
    clearTimeout(_fontRenderTimer);
    _fontRenderTimer = setTimeout(() => { _needsPagesPanelUpdate = true; render(); }, 150);
  } else {
    // No text element selected — drop a new text element with this font onto the active page
    const page = getActivePage();
    if (!page) return;
    saveHistory();
    const newId = generateId();
    page.elements.push({
      id: newId, t: 'text', x: 80, y: 100, w: 200, h: 40,
      txt: 'Text', fontFamily: f, fontSize: 24, color: '#111111', z: page.elements.length + 1
    });
    state.selected = newId;
    _needsPagesPanelUpdate = true;
    render();
  }
}

let _fontSizeRenderTimer = null;
let _fontSizeHistorySaved = false;
let _fontSizeHistoryTimer = null;
function setFontSize(s) {
  const el = getSelectedElement();
  if (!el) return;
  // Debounce saveHistory during rapid input (arrow hold, scroll, typing)
  if (!_fontSizeHistorySaved) { saveHistory(); _fontSizeHistorySaved = true; }
  clearTimeout(_fontSizeHistoryTimer);
  _fontSizeHistoryTimer = setTimeout(() => { _fontSizeHistorySaved = false; }, 600);
  el.fontSize = parseInt(s);
  // Update DOM directly instead of full render()
  const textDiv = document.querySelector(`[data-id="${el.id}"] .text-content`);
  if (textDiv) textDiv.style.fontSize = el.fontSize + 'px';
  // Defer single render() until input settles
  clearTimeout(_fontSizeRenderTimer);
  _fontSizeRenderTimer = setTimeout(() => { _needsPagesPanelUpdate = true; render(); }, 150);
}
function toggleBold() { const el = getSelectedElement(); if (el) { saveHistory(); el.bold = !el.bold; _needsPagesPanelUpdate = true; render(); }}
function toggleItalic() { const el = getSelectedElement(); if (el) { saveHistory(); el.italic = !el.italic; _needsPagesPanelUpdate = true; render(); }}
function setAlign(a) { const el = getSelectedElement(); if (el) { saveHistory(); el.align = a; _needsPagesPanelUpdate = true; render(); }}
function setTextColor(c) {
  const el = getSelectedElement();
  if (!el) return;
  _colorHistoryGate();
  el.color = c;
  trackColorUsage(c);
  // During color picker drag, update DOM directly instead of full render()
  if (window._colorPickerActive) {
    const textDiv = document.querySelector(`[data-id="${el.id}"] .text-content`);
    if (textDiv) textDiv.style.color = c;
    window._textColorDirty = true;
    return;
  }
  render();
}
function fillLorem() {
  const el = getSelectedElement();
  if (!el?.isBox) return;
  const target = parseInt(document.getElementById('targetWords')?.value);
  if (!target) return;
  saveHistory();
  el.targetWords = target;
  el.txt = generateLoremIpsum(target);
  _needsPagesPanelUpdate = true;
  render();
}

// Image tools
function setBorder(b) { const el = getSelectedElement(); if (el) { saveHistory(); el.border = b === 'none' ? null : b; _needsPagesPanelUpdate = true; render(); }}
function setFrameWidth(w) { const el = getSelectedElement(); if (el) { saveHistory(); el.frameWidth = parseInt(w); _needsPagesPanelUpdate = true; render(); }}

// Shape tools
function setShapeType(t) { const el = getSelectedElement(); if (el) { saveHistory(); el.shape = t; render(); }}
function setShapeColor(c) {
  const el = getSelectedElement();
  if (!el) return;
  _colorHistoryGate();
  el.color = c;
  trackColorUsage(c);
  // During color picker drag, update DOM directly instead of full render()
  if (window._colorPickerActive) {
    const shapeDiv = document.querySelector(`[data-id="${el.id}"] .shape-inner`);
    if (shapeDiv) shapeDiv.style.background = c;
    window._shapeColorDirty = true;
    return;
  }
  render();
}


// Phase A — clipboard quick-action buttons live in the Tools sidebar and share
// their handlers with the keyboard shortcuts (⌘X/⌘C/⌘V). This function syncs
// the buttons' disabled state with current selection + clipboard contents so the
// affordance matches what the action will actually do. Called from render() and
// from doCopy (since Copy mutates clipboard without otherwise triggering a render).
function updateQuickActions() {
  const cutBtn = document.getElementById('qaCut');
  const copyBtn = document.getElementById('qaCopy');
  const pasteBtn = document.getElementById('qaPaste');
  if (!cutBtn || !copyBtn || !pasteBtn) return;
  const hasSelection = !!state.selected || (Array.isArray(state.multiSelected) && state.multiSelected.length > 0);
  const hasClipboard = !!state.clipboard && (state.lastCopied === 'element' || state.lastCopied === 'multi');
  cutBtn.disabled = !hasSelection;
  copyBtn.disabled = !hasSelection;
  pasteBtn.disabled = !hasClipboard;
}

function doCopy() {
  // Handle multi-selection copy
  if (state.multiSelected.length > 0) {
    const page = getActivePage();
    if (!page) return;
    const elements = page.elements.filter(e => state.multiSelected.includes(e.id));
    state.clipboard = JSON.parse(JSON.stringify(elements));
    state.lastCopied = 'multi';
    document.getElementById('ctxMenu').classList.remove('show');
    updateQuickActions();
    return;
  }

  const el = getSelectedElement();
  if (el) {
    state.clipboard = JSON.parse(JSON.stringify(el));
    state.lastCopied = 'element';
  }
  document.getElementById('ctxMenu').classList.remove('show');
  updateQuickActions();
}

function doCut() {
  // Handle multi-selection cut
  if (state.multiSelected.length > 0) {
    doCopy();
    deleteElement();
    document.getElementById('ctxMenu').classList.remove('show');
    return;
  }
  
  const el = getSelectedElement();
  if (el) {
    state.clipboard = JSON.parse(JSON.stringify(el));
    state.lastCopied = 'element';
    deleteElement();
  }
  document.getElementById('ctxMenu').classList.remove('show');
}

function doPaste() {
  if (!state.clipboard) return;
  const page = getActivePage();
  if (!page) return;
  saveHistory();
  
  // Handle multi-element paste
  if (state.lastCopied === 'multi' && Array.isArray(state.clipboard)) {
    const newIds = [];
    state.clipboard.forEach((el, i) => {
      const newId = generateId();
      page.elements.push({
        ...el,
        id: newId,
        x: el.x + 20,
        y: el.y + 20,
        z: page.elements.length + 1
      });
      newIds.push(newId);
    });
    state.multiSelected = newIds;
    state.selected = null;
    _needsPagesPanelUpdate = true;
    render();
    document.getElementById('ctxMenu').classList.remove('show');
    return;
  }
  
  // Single element paste
  page.elements.push({
    ...state.clipboard,
    id: generateId(),
    x: state.clipboard.x + 20,
    y: state.clipboard.y + 20,
    z: page.elements.length + 1
  });
  _needsPagesPanelUpdate = true;
  render();
  document.getElementById('ctxMenu').classList.remove('show');
}

function doDuplicate() {
  doCopy();
  doPaste();
}

function bringToFront() {
  const page = getActivePage();
  if (!page) return;
  
  // Handle multi-selection
  if (state.multiSelected.length > 0) {
    saveHistory();
    const maxZ = Math.max(...page.elements.map(e => e.z || 0));
    state.multiSelected.forEach((id, i) => {
      const el = page.elements.find(e => e.id === id);
      if (el) el.z = maxZ + 1 + i;
    });
    render();
    document.getElementById('ctxMenu').classList.remove('show');
    return;
  }
  
  // Single selection
  const el = getSelectedElement();
  if (!el) return;
  saveHistory();
  el.z = Math.max(...page.elements.map(e => e.z || 0)) + 1;
  render();
  document.getElementById('ctxMenu').classList.remove('show');
}

function sendToBack() {
  const page = getActivePage();
  if (!page) return;
  
  // Handle multi-selection
  if (state.multiSelected.length > 0) {
    saveHistory();
    // Set all selected elements to z = 0, then shift others up
    const minZ = state.multiSelected.length;
    page.elements.forEach(el => {
      if (state.multiSelected.includes(el.id)) {
        el.z = state.multiSelected.indexOf(el.id);
      } else {
        el.z = (el.z || 0) + minZ;
      }
    });
    render();
    document.getElementById('ctxMenu').classList.remove('show');
    return;
  }
  
  // Single selection
  const el = getSelectedElement();
  if (el) { saveHistory(); el.z = 0; render(); }
  document.getElementById('ctxMenu').classList.remove('show');
}

// ============ IMAGE POSITION MODE ============
// Figma-style: double-click image to reposition it inside its frame.
// No zoom, no crop-resize, no competing render paths.

function enterImagePositionMode(elId) {
  const el = findElementById(elId);
  if (!el || el.t !== 'image' || !el.src) return;

  // Legacy image: auto-upgrade to locked model on first crop-mode entry
  if (el.innerW === undefined) {
    const img = document.querySelector(`[data-id="${elId}"] img`);
    if (img && img.naturalWidth && img.naturalHeight) {
      const coverScale = Math.max(el.w / img.naturalWidth, el.h / img.naturalHeight);
      const iw = Math.round(img.naturalWidth * coverScale);
      const ih = Math.round(img.naturalHeight * coverScale);
      // Apply any existing imgOffset so the upgrade preserves current pan
      const ox = el.imgOffsetX || 0;
      const oy = el.imgOffsetY || 0;
      el.innerX = Math.round((el.w - iw) / 2 + ox);
      el.innerY = Math.round((el.h - ih) / 2 + oy);
      el.innerW = iw;
      el.innerH = ih;
      // Clear legacy offset — now fully managed by innerX/Y
      delete el.imgOffsetX;
      delete el.imgOffsetY;
      saveHistory();
      scheduleAutosave();
    }
  }

  state.imagePositionMode = elId;
  state.selected = elId;
  render();
}

function exitImagePositionMode() {
  if (!state.imagePositionMode) return;
  state.imagePositionMode = null;
  render();
}

function resetImagePosition() {
  const el = getSelectedElement();
  if (!el || el.t !== 'image') return;
  saveHistory();
  el.imgOffsetX = 0;
  el.imgOffsetY = 0;
  render();
}

function setFitMode(mode) {
  const el = getSelectedElement();
  if (!el || el.t !== 'image') return;
  saveHistory();
  el.fitMode = mode;
  render();
}

function flipImage(dir) {
  const el = getSelectedElement();
  if (!el || el.t !== 'image') return;
  saveHistory();
  if (dir === 'h') {
    el.flipH = !el.flipH;
  } else {
    el.flipV = !el.flipV;
  }
  _needsPagesPanelUpdate = true;
  render();
}

function setImageOpacity(val) {
  const el = getSelectedElement();
  if (!el || el.t !== 'image') return;
  // Save history on first change (debounced via the slider)
  if (!el._opacitySaving) {
    saveHistory();
    el._opacitySaving = true;
    setTimeout(() => { delete el._opacitySaving; }, 500);
  }
  el.opacity = parseInt(val) / 100;
  const imgEl = document.querySelector(`[data-id="${el.id}"] img`);
  if (imgEl) imgEl.style.opacity = el.opacity;
  // Update slider display
  const span = document.querySelector('.slider-row span');
  if (span) span.textContent = val + '%';
  // Notify persistence layer
  builder.emit('change');
}

// Image adjustment slider: stash pre-change values on first tick, live-preview only
let _adjustPre = null;
function setImageAdjust(prop, val) {
  const el = getSelectedElement();
  if (!el || el.t !== 'image') return;
  // Stash original values before first mutation (lightweight, no stringify)
  if (!_adjustPre) {
    _adjustPre = { id: el.id, exposure: el.exposure, contrast: el.contrast, saturation: el.saturation };
  }
  el[prop] = parseInt(val);
  const imgEl = document.querySelector(`[data-id="${el.id}"] img`);
  if (imgEl) {
    const b = el.exposure ?? 100;
    const c = el.contrast ?? 100;
    const s = el.saturation ?? 100;
    imgEl.style.filter = (b === 100 && c === 100 && s === 100) ? '' : `brightness(${b/100}) contrast(${c/100}) saturate(${s/100})`;
  }
}

// Called once on slider release (onchange) — restore pre-values, snapshot, re-apply, render
function commitImageAdjust() {
  if (!_adjustPre) return;
  var el = getSelectedElement();
  if (el && el.id === _adjustPre.id) {
    // Save current (post-drag) values
    var postEx = el.exposure, postCt = el.contrast, postSa = el.saturation;
    // Temporarily restore pre-drag values so saveHistory captures the old state
    el.exposure = _adjustPre.exposure;
    el.contrast = _adjustPre.contrast;
    el.saturation = _adjustPre.saturation;
    saveHistory();
    // Re-apply post-drag values
    el.exposure = postEx;
    el.contrast = postCt;
    el.saturation = postSa;
  }
  _adjustPre = null;
  _needsPagesPanelUpdate = true;
  render();
  builder.emit('change');
}

// ============ DRAW TOOL ============
function toggleDraw() {
  if (drawState.active) {
    finishDrawing();
  } else {
    drawState.active = true;
    state.selected = null;
    document.getElementById('drawBar').classList.add('show');
    render();
    // Verify canvas was created and events are bound
    const c = document.getElementById('drawCanvas');
    console.log('[toggleDraw] canvas exists:', !!c, 'drawState.canvas:', !!drawState.canvas);
    if (!c) console.error('[toggleDraw] WARNING: No drawCanvas found in DOM after render!');
  }
}

function setDrawTool(tool) {
  drawState.tool = tool;
  document.querySelectorAll('.draw-btn[data-tool]').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tool === tool);
  });
}

function setDrawColor(color) {
  drawState.color = color;
  const picker = document.getElementById('drawColorPicker');
  if (picker && picker.value !== color) picker.value = color;
}

function startDrawStroke(e) {
  console.log('[startDrawStroke] called, target:', e.target.tagName, e.target.id);
  // Re-acquire canvas from DOM in case render() replaced it
  const freshCanvas = document.getElementById('drawCanvas');
  if (freshCanvas && freshCanvas !== drawState.canvas) {
    drawState.canvas = freshCanvas;
    drawState.ctx = freshCanvas.getContext('2d', { willReadFrequently: true });
  }
  if (!drawState.canvas) { console.error('[startDrawStroke] No canvas!'); return; }
  drawState.drawing = true;
  const rect = drawState.canvas.getBoundingClientRect();
  const x = e.clientX - rect.left;
  const y = e.clientY - rect.top;
  const ctx = drawState.ctx;
  ctx.strokeStyle = drawState.color;
  ctx.lineWidth = parseInt(document.getElementById('drawSizeSlider').value);
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  if (drawState.tool === 'pencil') ctx.globalAlpha = 0.6;
  else if (drawState.tool === 'chalk') { ctx.globalAlpha = 0.5; ctx.lineWidth *= 2; }
  else ctx.globalAlpha = 1;
  ctx.beginPath();
  ctx.moveTo(x, y);
}

function continueDrawStroke(e) {
  if (!drawState.drawing || !drawState.canvas) return;
  const rect = drawState.canvas.getBoundingClientRect();
  let x = e.clientX - rect.left;
  let y = e.clientY - rect.top;
  if (drawState.tool === 'chalk') { x += (Math.random() - 0.5) * 3; y += (Math.random() - 0.5) * 3; }
  drawState.ctx.lineTo(x, y);
  drawState.ctx.stroke();
  drawState.ctx.beginPath();
  drawState.ctx.moveTo(x, y);
}

function endDrawStroke() { drawState.drawing = false; }

function finishDrawing() {
  if (drawState.canvas && drawState.ctx) {
    const imageData = drawState.ctx.getImageData(0, 0, drawState.canvas.width, drawState.canvas.height);
    const data = imageData.data;
    const w = drawState.canvas.width;
    const h = drawState.canvas.height;
    
    // Find bounding box of actual content
    let minX = w, minY = h, maxX = 0, maxY = 0;
    let hasContent = false;
    
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const idx = (y * w + x) * 4;
        if (data[idx + 3] > 0) { // Check alpha channel
          hasContent = true;
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
      }
    }
    
    if (hasContent) {
      // Add small padding
      const pad = 5;
      minX = Math.max(0, minX - pad);
      minY = Math.max(0, minY - pad);
      maxX = Math.min(w - 1, maxX + pad);
      maxY = Math.min(h - 1, maxY + pad);
      
      const cropW = maxX - minX + 1;
      const cropH = maxY - minY + 1;
      
      // Create cropped canvas
      const croppedCanvas = document.createElement('canvas');
      croppedCanvas.width = cropW;
      croppedCanvas.height = cropH;
      const croppedCtx = croppedCanvas.getContext('2d');
      croppedCtx.drawImage(drawState.canvas, minX, minY, cropW, cropH, 0, 0, cropW, cropH);
      
      saveHistory();
      const page = getPage();
      const dataUrl = croppedCanvas.toDataURL('image/png');
      const drawEl = {
        id: generateId(),
        t: 'drawing',
        x: minX,
        y: minY,
        w: cropW,
        h: cropH,
        dataUrl: dataUrl,
        z: page.elements.length + 1
      };
      page.elements.push(drawEl);

      // Upload drawing to R2 in background, swap in permanent URL
      const gen = _stateGeneration;
      uploadImageToR2(dataUrl).then(url => {
        if (_stateGeneration !== gen) return; // State was undone/redone, discard
        if (url && url !== dataUrl && drawEl.dataUrl === dataUrl) {
          drawEl.dataUrl = url;
        }
        scheduleAutosave();
      });
    }
  }
  drawState.active = false;
  document.getElementById('drawBar').classList.remove('show');
  _needsPagesPanelUpdate = true;
  render();
}

function cancelDrawing() {
  drawState.active = false;
  document.getElementById('drawBar').classList.remove('show');
  render();
}

// ============ REVEAL BRUSH ============
function openRevealBrush() {
  const el = getSelectedElement();
  if (!el?.src || el.t !== 'image') {
    alert('Please select an image first');
    return;
  }
  
  // Deselect to prevent drag interference
  state.selected = null;
  
  revealState.active = true;
  revealState.quickMode = false;
  revealState.element = el;
  revealState.maskHistory = [];
  revealState.maskHistoryIndex = -1;
  
  document.getElementById('revealBar').classList.add('show');
  document.getElementById('revealSize').oninput = function() {
    document.getElementById('revealSizeLabel').textContent = this.value;
  };
  
  render();
  setTimeout(() => setupRevealCanvas(), 50);
}

function setupRevealCanvas() {
  if (!revealState.element) return;
  
  const el = revealState.element;
  
  // Find the DOM element for this image
  const domEl = document.querySelector(`.element[data-id="${el.id}"]`);
  if (!domEl) return;
  
  // Determine the container and offset
  // In spread mode, elements are in .spread-elements (above .page), so we attach to .spread
  // On cover, elements are directly in .page, so we attach to .page
  let container;
  let canvasLeft = 0;
  let canvasTop = 0;
  const spread = domEl.closest('.spread');
  
  if (spread) {
    container = spread;
    // In spread mode, figure out if element is on the right page (offset by 400px)
    const xOffset = parseInt(domEl.dataset.xoffset || '0');
    canvasLeft = xOffset;
  } else {
    container = domEl.closest('.page');
  }
  if (!container) return;
  
  // Remove existing reveal canvas if any
  const existing = document.getElementById('revealCanvasOverlay');
  if (existing) existing.remove();
  
  // Create canvas overlay that covers the page area
  const canvas = document.createElement('canvas');
  canvas.id = 'revealCanvasOverlay';
  canvas.className = 'reveal-canvas-overlay';
  canvas.width = PAGE_W;
  canvas.height = PAGE_H;
  canvas.style.cssText = `
    position: absolute;
    top: ${canvasTop}px;
    left: ${canvasLeft}px;
    width: ${PAGE_W}px;
    height: ${PAGE_H}px;
    z-index: 9999;
    cursor: crosshair;
    pointer-events: auto;
  `;
  
  container.appendChild(canvas);
  
  const ctx = canvas.getContext('2d');
  revealState.canvas = canvas;
  revealState.ctx = ctx;
  revealState.pageOffset = { x: el.x, y: el.y };
  revealState.elSize = { w: el.w, h: el.h };
  
  // Create a separate mask canvas at element size
  const maskCanvas = document.createElement('canvas');
  maskCanvas.width = el.w;
  maskCanvas.height = el.h;
  revealState.maskCanvas = maskCanvas;
  revealState.maskCtx = maskCanvas.getContext('2d', { willReadFrequently: true });
  
  // Load existing mask if any
  if (el.revealMask) {
    const img = new Image();
    img.crossOrigin = 'anonymous'; // Required for clean canvas (getImageData on mask)
    img.onerror = function() {
      console.error('Reveal: mask image failed to load (CORS)');
    };
    img.onload = () => {
      revealState.maskCtx.drawImage(img, 0, 0, el.w, el.h);
      drawRevealPreview();
      saveRevealHistory();
    };
    img.src = el.revealMask.startsWith('data:') ? el.revealMask : el.revealMask + (el.revealMask.includes('?') ? '&' : '?') + '_cors=1';
  } else {
    drawRevealPreview();
    saveRevealHistory();
  }
  
  // Bind events with event capture to prevent propagation
  canvas.addEventListener('mousedown', revealMouseDown, true);
  canvas.addEventListener('mousemove', revealMouseMove, true);
  canvas.addEventListener('mouseup', revealMouseUp, true);
  canvas.addEventListener('mouseleave', revealMouseUp, true);
}

function drawRevealPreview() {
  if (!revealState.ctx || !revealState.maskCanvas || !revealState.element) return;
  
  const ctx = revealState.ctx;
  const el = revealState.element;
  
  // Clear the overlay canvas
  ctx.clearRect(0, 0, PAGE_W, PAGE_H);

  // Draw semi-transparent dark overlay on the whole page to dim it
  ctx.fillStyle = 'rgba(0, 0, 0, 0.3)';
  ctx.fillRect(0, 0, PAGE_W, PAGE_H);
  
  // Cut out the element area (make it clear so we can see the page beneath)
  ctx.clearRect(el.x, el.y, el.w, el.h);
  
  // Now draw the reveal preview - show the masked image at high opacity
  // This simulates what the final reveal will look like
  const img = document.querySelector(`.element[data-id="${el.id}"] img`);
  if (img && revealState.maskCanvas) {
    // Create a temp canvas to composite the reveal preview
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = el.w;
    tempCanvas.height = el.h;
    const tempCtx = tempCanvas.getContext('2d');
    
    // Draw the image
    tempCtx.drawImage(img, 0, 0, el.w, el.h);
    
    // Apply the mask using destination-in
    tempCtx.globalCompositeOperation = 'destination-in';
    tempCtx.drawImage(revealState.maskCanvas, 0, 0);
    
    // Draw this preview on our overlay, slightly offset and with a glow to show it's the "pop forward" layer
    ctx.save();
    ctx.shadowColor = '#e91e63';
    ctx.shadowBlur = 10;
    ctx.drawImage(tempCanvas, el.x, el.y);
    ctx.restore();
  }
  
  // Draw border around paintable area
  ctx.strokeStyle = '#e91e63';
  ctx.lineWidth = 2;
  ctx.setLineDash([5, 5]);
  ctx.strokeRect(el.x, el.y, el.w, el.h);
  ctx.setLineDash([]);
}

function revealMouseDown(e) {
  e.preventDefault();
  e.stopPropagation();
  
  if (!revealState.active && !revealState.quickMode) return;
  revealState.drawing = true;
  revealState.lastBrushPos = null;
  revealPaint(e);
}

function revealMouseMove(e) {
  e.preventDefault();
  e.stopPropagation();
  
  if (!revealState.drawing) return;
  revealPaint(e);
}

function revealMouseUp(e) {
  e.preventDefault();
  e.stopPropagation();
  
  if (revealState.drawing) {
    revealState.drawing = false;
    revealState.lastBrushPos = null;
    saveRevealHistory();
    saveRevealMaskToElement();
  }
}

function revealPaint(e) {
  if (!revealState.maskCtx || !revealState.canvas) return;
  
  const rect = revealState.canvas.getBoundingClientRect();
  // Convert screen coordinates to canvas coordinates (account for display scaling)
  const scaleX = revealState.canvas.width / rect.width;
  const scaleY = revealState.canvas.height / rect.height;
  const canvasX = (e.clientX - rect.left) * scaleX;
  const canvasY = (e.clientY - rect.top) * scaleY;
  
  // Convert to element-local coordinates
  const el = revealState.element;
  const x = canvasX - el.x;
  const y = canvasY - el.y;
  
  // Only paint within element bounds
  if (x < 0 || x > el.w || y < 0 || y > el.h) return;
  
  const size = parseInt(document.getElementById('revealSize')?.value) || 30;
  const radius = size / 2;

  // Paint white on the mask (white = revealed)
  const maskCtx = revealState.maskCtx;
  maskCtx.fillStyle = '#ffffff';

  // Interpolate from last position to current for smooth strokes
  const last = revealState.lastBrushPos || { x, y };
  const dx = x - last.x;
  const dy = y - last.y;
  const dist = Math.sqrt(dx * dx + dy * dy);
  const step = Math.max(1, radius / 2);
  const steps = Math.max(1, Math.ceil(dist / step));

  for (let i = 0; i <= steps; i++) {
    const t = steps === 1 ? 1 : i / steps;
    const px = last.x + dx * t;
    const py = last.y + dy * t;
    maskCtx.beginPath();
    maskCtx.arc(px, py, radius, 0, Math.PI * 2);
    maskCtx.fill();
  }

  revealState.lastBrushPos = { x, y };

  // Update preview
  drawRevealPreview();
}

function saveRevealHistory() {
  if (!revealState.maskCanvas) return;
  
  revealState.maskHistory = revealState.maskHistory.slice(0, revealState.maskHistoryIndex + 1);
  revealState.maskHistory.push(revealState.maskCanvas.toDataURL());
  revealState.maskHistoryIndex = revealState.maskHistory.length - 1;
  
  if (revealState.maskHistory.length > 30) {
    revealState.maskHistory.shift();
    revealState.maskHistoryIndex--;
  }
}

function revealUndo() {
  if (revealState.maskHistoryIndex > 0) {
    revealState.maskHistoryIndex--;
    restoreRevealMask();
  }
}

function revealRedo() {
  if (revealState.maskHistoryIndex < revealState.maskHistory.length - 1) {
    revealState.maskHistoryIndex++;
    restoreRevealMask();
  }
}

function restoreRevealMask() {
  const data = revealState.maskHistory[revealState.maskHistoryIndex];
  if (!data || !revealState.maskCtx) return;
  
  const img = new Image();
  img.onload = () => {
    revealState.maskCtx.clearRect(0, 0, revealState.maskCanvas.width, revealState.maskCanvas.height);
    revealState.maskCtx.drawImage(img, 0, 0);
    drawRevealPreview();
    saveRevealMaskToElement();
  };
  img.src = data;
}

function saveRevealMaskToElement() {
  if (!revealState.element || !revealState.maskCanvas) return;
  
  // Check if mask has any content
  const ctx = revealState.maskCtx;
  const imageData = ctx.getImageData(0, 0, revealState.maskCanvas.width, revealState.maskCanvas.height);
  let hasContent = false;
  for (let i = 3; i < imageData.data.length; i += 4) {
    if (imageData.data[i] > 0) { hasContent = true; break; }
  }
  
  if (hasContent) {
    const maskDataUrl = revealState.maskCanvas.toDataURL();
    revealState.element.revealMask = maskDataUrl;

    // Upload mask to R2 in background, swap in permanent URL
    const revealEl = revealState.element;
    const gen = _stateGeneration;
    uploadImageToR2(maskDataUrl).then(url => {
      if (_stateGeneration !== gen) return; // State was undone/redone, discard
      if (url && url !== maskDataUrl && revealEl.revealMask === maskDataUrl) {
        revealEl.revealMask = url;
      }
      scheduleAutosave();
    });
  } else {
    delete revealState.element.revealMask;
  }

  // Re-render just the element, not the whole page (to keep our canvas)
  const pageEl = document.querySelector(`.element[data-id="${revealState.element.id}"]`);
  if (pageEl) {
    // Update just the reveal layer inside the element
    updateRevealLayer(revealState.element, pageEl);
  }
}

function updateRevealLayer(el, pageEl) {
  // Remove old reveal layer
  const oldLayer = pageEl.querySelector('.reveal-layer');
  if (oldLayer) oldLayer.remove();
  
  if (!el.revealMask) return;
  
  // Create new reveal layer
  const layer = document.createElement('div');
  layer.className = 'reveal-layer';
  layer.style.cssText = 'position:absolute;inset:0;z-index:999;pointer-events:none;overflow:hidden;';
  
  const img = pageEl.querySelector('img');
  if (img) {
    const revealImg = img.cloneNode();
    revealImg.style.cssText = img.style.cssText + `-webkit-mask-image:url(${el.revealMask});mask-image:url(${el.revealMask});-webkit-mask-size:100% 100%;mask-size:100% 100%;`;
    layer.appendChild(revealImg);
  }
  
  pageEl.appendChild(layer);
}

function clearRevealMask() {
  if (revealState.element) {
    saveHistory();
    delete revealState.element.revealMask;
  }
  if (revealState.maskCtx && revealState.maskCanvas) {
    revealState.maskCtx.clearRect(0, 0, revealState.maskCanvas.width, revealState.maskCanvas.height);
    saveRevealHistory();
  }
  if (revealState.ctx) {
    drawRevealPreview();
  }
  render();
  if (revealState.active) {
    setTimeout(setupRevealCanvas, 50);
  }
}

function closeRevealBrush() {
  revealState.active = false;
  revealState.quickMode = false;
  
  // Remove overlay canvas
  const overlay = document.getElementById('revealCanvasOverlay');
  if (overlay) overlay.remove();
  
  revealState.canvas = null;
  revealState.ctx = null;
  revealState.maskCanvas = null;
  revealState.maskCtx = null;
  revealState.element = null;
  
  document.getElementById('revealBar').classList.remove('show');
  document.getElementById('revealIndicator').classList.remove('show');
  
  render();
}

// Quick reveal mode with R key
function startQuickReveal() {
  const el = getSelectedElement();
  if (!el?.src || el.t !== 'image') return;
  
  // Deselect to prevent drag
  state.selected = null;
  
  revealState.quickMode = true;
  revealState.element = el;
  revealState.maskHistory = [];
  revealState.maskHistoryIndex = -1;
  
  document.getElementById('revealIndicator').classList.add('show');
  render();
  setTimeout(() => setupRevealCanvas(), 50);
}

function endQuickReveal() {
  if (!revealState.quickMode) return;
  
  revealState.quickMode = false;
  
  // Remove overlay canvas
  const overlay = document.getElementById('revealCanvasOverlay');
  if (overlay) overlay.remove();
  
  revealState.canvas = null;
  revealState.ctx = null;
  revealState.maskCanvas = null;
  revealState.maskCtx = null;
  
  document.getElementById('revealIndicator').classList.remove('show');
  
  // Don't clear element - keep the mask
  if (!revealState.active) {
    revealState.element = null;
  }
  
  render();
}


// ============ CUT OUT (BACKGROUND REMOVAL) ============
function openCutout() {
  const el = getSelectedElement();
  if (!el?.src) return;
  document.getElementById('ctxMenu').classList.remove('show');
  
  const canvas = document.getElementById('cutoutCanvas');
  const overlay = document.getElementById('cutoutOverlay');
  
  // Reset state
  cutoutState = {
    element: el,
    ctx: null,
    img: null,
    mask: null,
    maskCtx: null,
    points: [],
    drawing: false,
    scaleX: 1,
    scaleY: 1,
    rectStart: null,
    maskHistory: [],
    maskHistoryIndex: -1
  };
  
  const img = new Image();
  img.crossOrigin = 'anonymous'; // Required for clean canvas (getImageData/toDataURL)
  img.onerror = function() {
    console.error('Cutout: image failed to load (CORS)');
  };
  img.onload = function() {
    // Size canvas to fit screen
    const maxW = window.innerWidth * 0.7;
    const maxH = window.innerHeight * 0.55;
    let w = img.width, h = img.height;
    if (w > maxW) { h *= maxW / w; w = maxW; }
    if (h > maxH) { w *= maxH / h; h = maxH; }
    
    // Set canvas dimensions
    canvas.width = img.width;
    canvas.height = img.height;
    canvas.style.width = w + 'px';
    canvas.style.height = h + 'px';
    canvas.style.cursor = 'crosshair';
    
    // Store references
    cutoutState.ctx = canvas.getContext('2d');
    cutoutState.img = img;
    cutoutState.scaleX = img.width / w;
    cutoutState.scaleY = img.height / h;
    
    // Draw original image
    cutoutState.ctx.drawImage(img, 0, 0);
    
    // Create mask canvas - uses ALPHA for masking
    // Start fully opaque (keep everything), tools make transparent (removes)
    cutoutState.mask = document.createElement('canvas');
    cutoutState.mask.width = img.width;
    cutoutState.mask.height = img.height;
    cutoutState.maskCtx = cutoutState.mask.getContext('2d', { willReadFrequently: true });
    cutoutState.maskCtx.fillStyle = 'white';
    cutoutState.maskCtx.fillRect(0, 0, img.width, img.height);

    // Save initial mask state
    saveMaskHistory();

    // Show overlay
    overlay.classList.add('show');
    
    // Setup mode change listener
    const modeSelect = document.getElementById('cutoutMode');
    const sizeSlider = document.getElementById('cutoutSize');
    const sizeLabel = document.getElementById('cutoutSizeLabel');
    const hint = document.getElementById('cutoutHint');
    
    function updateHint() {
      const mode = modeSelect.value;
      if (mode === 'wand') {
        hint.textContent = '🪄 Magic Wand: Click on background to remove similar colors. Adjust tolerance with slider.';
        sizeLabel.textContent = sizeSlider.value + ' (tolerance)';
      } else if (mode === 'brush') {
        hint.textContent = '🖌️ Brush: Paint over areas to remove them.';
        sizeLabel.textContent = sizeSlider.value + 'px';
      } else {
        hint.textContent = '⬛ Rectangle: Drag to select rectangular area to remove.';
        sizeLabel.textContent = sizeSlider.value + 'px';
      }
    }

    modeSelect.onchange = updateHint;
    sizeSlider.oninput = () => {
      const mode = modeSelect.value;
      sizeLabel.textContent = sizeSlider.value + (mode === 'wand' ? ' (tolerance)' : 'px');
    };
    
    updateHint();
  };
  const imgSrc = el.src;
  img.src = imgSrc.startsWith('data:') ? imgSrc : imgSrc + (imgSrc.includes('?') ? '&' : '?') + '_cors=1';
}

// Mouse event handlers - attached directly to canvas in HTML
function cutoutMouseDown(e) {
  if (!cutoutState.img) return;
  const canvas = document.getElementById('cutoutCanvas');
  const rect = canvas.getBoundingClientRect();
  const x = (e.clientX - rect.left) * cutoutState.scaleX;
  const y = (e.clientY - rect.top) * cutoutState.scaleY;
  
  const mode = document.getElementById('cutoutMode').value;
  
  if (mode === 'wand') {
    // Magic wand - flood fill similar colors
    magicWandSelect(Math.round(x), Math.round(y));
    return;
  }

  cutoutState.drawing = true;

  if (mode === 'rect') {
    cutoutState.rectStart = { x, y };
  } else {
    // Brush mode - start drawing
    cutoutState.lastBrushPos = { x, y };
    handleCutoutDraw(e);
  }
}

function cutoutMouseMove(e) {
  if (!cutoutState.img) return;
  
  const mode = document.getElementById('cutoutMode').value;
  
  if (mode === 'rect' && cutoutState.drawing && cutoutState.rectStart) {
    // Show rectangle preview
    const canvas = document.getElementById('cutoutCanvas');
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) * cutoutState.scaleX;
    const y = (e.clientY - rect.top) * cutoutState.scaleY;
    drawRectPreview(x, y);
  } else if (mode === 'brush' && cutoutState.drawing) {
    handleCutoutDraw(e);
  }
}

function cutoutMouseUp(e) {
  if (!cutoutState.img) return;

  const mode = document.getElementById('cutoutMode').value;

  if (mode === 'rect' && cutoutState.drawing && cutoutState.rectStart) {
    const canvas = document.getElementById('cutoutCanvas');
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) * cutoutState.scaleX;
    const y = (e.clientY - rect.top) * cutoutState.scaleY;
    
    // Apply rectangle erase
    const rx = Math.min(cutoutState.rectStart.x, x);
    const ry = Math.min(cutoutState.rectStart.y, y);
    const rw = Math.abs(x - cutoutState.rectStart.x);
    const rh = Math.abs(y - cutoutState.rectStart.y);
    
    if (rw > 5 && rh > 5) {
      cutoutState.maskCtx.save();
      cutoutState.maskCtx.globalCompositeOperation = 'destination-out';
      cutoutState.maskCtx.fillStyle = 'white';
      cutoutState.maskCtx.fillRect(rx, ry, rw, rh);
      cutoutState.maskCtx.restore();
      
      saveMaskHistory();
      drawBrushPreview();
    }
    
    cutoutState.rectStart = null;
  } else if (mode === 'brush' && cutoutState.drawing) {
    // Save history after brush stroke ends
    saveMaskHistory();
  }
  
  cutoutState.drawing = false;
  cutoutState.lastBrushPos = null;
}

function drawRectPreview(currentX, currentY) {
  const { ctx, img, mask, rectStart } = cutoutState;
  if (!ctx || !img || !rectStart) return;
  
  // Draw current mask state
  drawBrushPreview();
  
  // Draw rectangle outline
  const rx = Math.min(rectStart.x, currentX);
  const ry = Math.min(rectStart.y, currentY);
  const rw = Math.abs(currentX - rectStart.x);
  const rh = Math.abs(currentY - rectStart.y);
  
  ctx.strokeStyle = '#ff4444';
  ctx.lineWidth = 2;
  ctx.setLineDash([5, 5]);
  ctx.strokeRect(rx, ry, rw, rh);
  ctx.setLineDash([]);
  
  // Fill with semi-transparent red to show what will be removed
  ctx.fillStyle = 'rgba(255, 0, 0, 0.3)';
  ctx.fillRect(rx, ry, rw, rh);
}

function handleCutoutDraw(e) {
  const canvas = document.getElementById('cutoutCanvas');
  const rect = canvas.getBoundingClientRect();
  const x = (e.clientX - rect.left) * cutoutState.scaleX;
  const y = (e.clientY - rect.top) * cutoutState.scaleY;
  
  // BRUSH MODE - erase from mask to REMOVE
  const size = parseInt(document.getElementById('cutoutSize').value) || 20;

  // Use destination-out to erase (make transparent) on mask
  cutoutState.maskCtx.save();
  cutoutState.maskCtx.globalCompositeOperation = 'destination-out';
  cutoutState.maskCtx.fillStyle = 'white';

  // Interpolate from last position to current for smooth strokes
  const last = cutoutState.lastBrushPos || { x, y };
  const dx = x - last.x;
  const dy = y - last.y;
  const dist = Math.sqrt(dx * dx + dy * dy);
  const step = Math.max(1, size / 4);
  const steps = Math.max(1, Math.ceil(dist / step));

  for (let i = 0; i <= steps; i++) {
    const t = steps === 1 ? 1 : i / steps;
    const px = last.x + dx * t;
    const py = last.y + dy * t;
    cutoutState.maskCtx.beginPath();
    cutoutState.maskCtx.arc(px, py, size, 0, Math.PI * 2);
    cutoutState.maskCtx.fill();
  }

  cutoutState.maskCtx.restore();
  cutoutState.lastBrushPos = { x, y };

  // Update preview
  drawBrushPreview();
}

// Mask history functions for undo/redo
function saveMaskHistory() {
  if (!cutoutState.mask) return;
  
  // Remove any redo states
  cutoutState.maskHistory = cutoutState.maskHistory.slice(0, cutoutState.maskHistoryIndex + 1);
  
  // Save current mask as ImageData
  const imageData = cutoutState.maskCtx.getImageData(0, 0, cutoutState.mask.width, cutoutState.mask.height);
  cutoutState.maskHistory.push(imageData);
  cutoutState.maskHistoryIndex = cutoutState.maskHistory.length - 1;
  
  // Limit history size
  if (cutoutState.maskHistory.length > 50) {
    cutoutState.maskHistory.shift();
    cutoutState.maskHistoryIndex--;
  }
}

function cutoutUndo() {
  if (cutoutState.maskHistoryIndex > 0) {
    cutoutState.maskHistoryIndex--;
    const imageData = cutoutState.maskHistory[cutoutState.maskHistoryIndex];
    cutoutState.maskCtx.putImageData(imageData, 0, 0);
    drawBrushPreview();
  }
}

function cutoutRedo() {
  if (cutoutState.maskHistoryIndex < cutoutState.maskHistory.length - 1) {
    cutoutState.maskHistoryIndex++;
    const imageData = cutoutState.maskHistory[cutoutState.maskHistoryIndex];
    cutoutState.maskCtx.putImageData(imageData, 0, 0);
    drawBrushPreview();
  }
}

function drawBrushPreview() {
  const { ctx, img, mask } = cutoutState;
  if (!ctx || !img || !mask) return;
  
  const canvas = ctx.canvas;
  
  // Draw checkerboard background (shows transparency)
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  const checkSize = 16;
  for (let cx = 0; cx < canvas.width; cx += checkSize) {
    for (let cy = 0; cy < canvas.height; cy += checkSize) {
      ctx.fillStyle = (Math.floor(cx/checkSize) + Math.floor(cy/checkSize)) % 2 === 0 ? '#999' : '#666';
      ctx.fillRect(cx, cy, checkSize, checkSize);
    }
  }
  
  // Create temp canvas with masked image
  const temp = document.createElement('canvas');
  temp.width = img.width;
  temp.height = img.height;
  const tc = temp.getContext('2d');
  
  // Draw image
  tc.drawImage(img, 0, 0);
  
  // Apply mask - destination-in keeps only where mask is opaque (white)
  tc.globalCompositeOperation = 'destination-in';
  tc.drawImage(mask, 0, 0);
  
  // Draw masked result over checkerboard
  ctx.drawImage(temp, 0, 0);
}

function clearCutout() {
  cutoutState.points = [];
  cutoutState.rectStart = null;

  // Reset mask to all white (keep everything)
  if (cutoutState.maskCtx && cutoutState.mask) {
    cutoutState.maskCtx.fillStyle = 'white';
    cutoutState.maskCtx.fillRect(0, 0, cutoutState.mask.width, cutoutState.mask.height);
    saveMaskHistory();
  }

  // Reset preview to original image
  if (cutoutState.ctx && cutoutState.img) {
    cutoutState.ctx.clearRect(0, 0, cutoutState.ctx.canvas.width, cutoutState.ctx.canvas.height);
    cutoutState.ctx.drawImage(cutoutState.img, 0, 0);
  }
}

// Magic Wand - flood fill to remove similar colors
function magicWandSelect(startX, startY) {
  const { img, mask, maskCtx } = cutoutState;
  if (!img || !mask || !maskCtx) return;
  
  // Get tolerance from slider (repurposed size slider)
  const tolerance = parseInt(document.getElementById('cutoutSize').value) || 20;
  
  // Create temp canvas to read original image pixels
  const tempCanvas = document.createElement('canvas');
  tempCanvas.width = img.width;
  tempCanvas.height = img.height;
  const tempCtx = tempCanvas.getContext('2d', { willReadFrequently: true });
  tempCtx.drawImage(img, 0, 0);

  let imageData;
  try {
    imageData = tempCtx.getImageData(0, 0, img.width, img.height);
  } catch (e) {
    console.warn('Magic wand: cannot read image pixels (cross-origin image). Use Brush mode instead.');
    alert('Magic wand is not available for this image (cross-origin). Please use Brush mode instead.');
    return;
  }
  const data = imageData.data;
  const width = img.width;
  const height = img.height;
  
  // Get target color at click point
  const targetIdx = (startY * width + startX) * 4;
  const targetR = data[targetIdx];
  const targetG = data[targetIdx + 1];
  const targetB = data[targetIdx + 2];
  
  // Create visited array
  const visited = new Uint8Array(width * height);
  
  // Pixels to make transparent in mask
  const toRemove = [];
  
  // Flood fill using queue
  const queue = [[startX, startY]];
  let qHead = 0;
  visited[startY * width + startX] = 1;

  while (qHead < queue.length) {
    const [x, y] = queue[qHead++];
    const idx = (y * width + x) * 4;
    
    // Check color similarity
    const r = data[idx];
    const g = data[idx + 1];
    const b = data[idx + 2];
    
    const diff = Math.abs(r - targetR) + Math.abs(g - targetG) + Math.abs(b - targetB);
    
    if (diff <= tolerance * 3) {
      toRemove.push([x, y]);
      
      // Add neighbors
      const neighbors = [[x-1, y], [x+1, y], [x, y-1], [x, y+1]];
      for (const [nx, ny] of neighbors) {
        if (nx >= 0 && nx < width && ny >= 0 && ny < height) {
          const nIdx = ny * width + nx;
          if (!visited[nIdx]) {
            visited[nIdx] = 1;
            queue.push([nx, ny]);
          }
        }
      }
    }
  }
  
  // Remove selected pixels from mask
  if (toRemove.length > 0) {
    const maskData = maskCtx.getImageData(0, 0, mask.width, mask.height);
    const md = maskData.data;
    
    for (const [x, y] of toRemove) {
      const idx = (y * mask.width + x) * 4;
      md[idx + 3] = 0; // Set alpha to 0 (transparent in mask = remove from image)
    }
    
    maskCtx.putImageData(maskData, 0, 0);
    saveMaskHistory();
    drawBrushPreview();
  }
}

function applyCutout() {
  const { element, img, mask } = cutoutState;

  if (!element || !img || !mask) {
    closeCutout();
    return;
  }

  saveHistory();

  // Store original for restore
  if (!element.originalSrc) {
    element.originalSrc = element.src;
  }

  // Create output canvas — apply mask to image
  const output = document.createElement('canvas');
  output.width = img.width;
  output.height = img.height;
  const ctx = output.getContext('2d');
  ctx.drawImage(img, 0, 0);
  ctx.globalCompositeOperation = 'destination-in';
  ctx.drawImage(mask, 0, 0);

  // Save with transparency — use data URL immediately for responsive UI,
  // then upload to R2 in background and swap in the permanent URL
  const dataUrl = output.toDataURL('image/png');
  element.src = dataUrl;
  element.border = null;
  element.deepEtched = true; // Mark as deep-etched for transparent background

  closeCutout();
  render();

  // Upload derived image to R2 (non-blocking)
  const gen = _stateGeneration;
  uploadImageToR2(dataUrl).then(url => {
    if (_stateGeneration !== gen) return;
    if (url && url !== dataUrl && element.src === dataUrl) { element.src = url; }
    scheduleAutosave();
  });
}

function closeCutout() {
  document.getElementById('cutoutOverlay').classList.remove('show');
  cutoutState = {
    element: null,
    ctx: null,
    img: null,
    mask: null,
    maskCtx: null,
    points: [],
    drawing: false,
    scaleX: 1,
    scaleY: 1,
    rectStart: null,
    maskHistory: [],
    maskHistoryIndex: -1
  };
}

function restoreBackground() {
  const el = getSelectedElement();
  if (!el?.originalSrc) return;

  saveHistory();
  el.src = el.originalSrc;
  el.deepEtched = false;
  render();
}


// ============ KEYBOARD SHORTCUTS ============
document.addEventListener('keydown', e => {
  // Check if deep-etch overlay is open
  const cutoutOpen = document.getElementById('cutoutOverlay').classList.contains('show');
  const isTyping = document.activeElement?.tagName === 'INPUT' || document.activeElement?.isContentEditable;
  
  // R key for quick reveal mode — disabled while reveal brush is being stabilised
  // if (e.key === 'r' && !isTyping && !cutoutOpen && !e.metaKey && !e.ctrlKey) {
  //   if (!revealState.quickMode && !revealState.active) {
  //     startQuickReveal();
  //   }
  //   return;
  // }
  
  if ((e.metaKey || e.ctrlKey) && e.key === 'a') {
    // If user is typing in any editable field, let browser handle select-all naturally
    if (isTyping) return;
    // Otherwise block it (prevents accidentally selecting everything on page)
    e.preventDefault();
    return;
  }
  
  if ((e.metaKey || e.ctrlKey) && e.key === 'z' && !e.shiftKey) { 
    e.preventDefault(); 
    if (cutoutOpen) cutoutUndo();
    else if (revealState.active || revealState.quickMode) revealUndo();
    else undo(); 
    return; 
  }
  if ((e.metaKey || e.ctrlKey) && (e.key === 'y' || (e.shiftKey && e.key === 'z'))) { 
    e.preventDefault(); 
    if (cutoutOpen) cutoutRedo();
    else if (revealState.active || revealState.quickMode) revealRedo();
    else redo();
    return; 
  }
  // ⌘⇧E / Ctrl+Shift+E — hidden export of the raw .zine JSON for internal/debug use.
  // Matched before ⌘S so the shift-modifier form doesn't fall through to the autosave no-op.
  if ((e.metaKey || e.ctrlKey) && e.shiftKey && (e.key === 'e' || e.key === 'E')) {
    e.preventDefault(); saveJSON(); return;
  }
  // ⌘S / Ctrl+S — swallowed silently. Autosave handles persistence; the visible Save
  // button was removed in the Phase A Tools reorg because a manual save contradicts
  // the autosave model shipped in the Phase 1+2 session bundle.
  if ((e.metaKey || e.ctrlKey) && e.key === 's') { e.preventDefault(); return; }
  
  // Delete - handle both single and multi selection
  if ((e.key === 'Delete' || e.key === 'Backspace') && (state.selected || state.multiSelected.length > 0) && !isTyping) { 
    e.preventDefault(); 
    deleteElement(); 
  }
  
  if (e.key === 'Escape') { 
    if (cutoutOpen) closeCutout();
    else if (revealState.active) closeRevealBrush();
    else if (revealState.quickMode) endQuickReveal();
    else if (state.imagePositionMode) exitImagePositionMode();
    else {
      state.selected = null;
      state.multiSelected = [];
      state.selectedSpread = null;
      if (drawState.active) cancelDrawing(); 
      render(); 
    }
  }
  if (e.key === 'Enter' && state.imagePositionMode && !isTyping) {
    e.preventDefault();
    exitImagePositionMode();
  }
  
  // Cut: Ctrl+X - handle both single and multi selection
  if ((e.metaKey || e.ctrlKey) && e.key === 'x' && !isTyping) {
    e.preventDefault();
    if (state.selected || state.multiSelected.length > 0) {
      doCut();
    }
  }
  
  // Copy: if element selected, copy element; if spread selected (no element), copy spread
  if ((e.metaKey || e.ctrlKey) && e.key === 'c' && !isTyping) {
    if (state.selected || state.multiSelected.length > 0) {
      doCopy();
    } else if (state.selectedSpread !== null) {
      copySpread();
    }
  }
  
  // Paste: use lastCopied to determine what to paste
  if ((e.metaKey || e.ctrlKey) && e.key === 'v' && !isTyping) {
    if (state.lastCopied === 'spread' && state.spreadClipboard) {
      pasteSpread();
    } else if ((state.lastCopied === 'element' || state.lastCopied === 'multi') && state.clipboard) {
      doPaste();
    }
  }
  
  if ((e.metaKey || e.ctrlKey) && e.key === 'd' && !isTyping) { 
    e.preventDefault(); 
    if (state.selected || state.multiSelected.length > 0) {
      doDuplicate(); 
    }
  }
});

// Keyup for quick reveal mode
document.addEventListener('keyup', e => {
  if (e.key === 'r' && revealState.quickMode) {
    endQuickReveal();
  }
});

// B2 FIX: Prevent browser from opening files when drop misses canvas
document.addEventListener('dragover', e => { e.preventDefault(); }, false);
document.addEventListener('drop', e => {
  // Only prevent if not already handled by a page drop handler
  if (!e._handled) { e.preventDefault(); }
}, false);

