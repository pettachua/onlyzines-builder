function initIcons() {
  const saveBtn = document.getElementById('btnSave');
  const pdfBtn = document.getElementById('btnPDF');
  if (saveBtn) {
    saveBtn.innerHTML = '<img src="' + DADA_ICONS.save + '"><span>Save</span>';
  }
  if (pdfBtn) {
    pdfBtn.innerHTML = '<img src="' + DADA_ICONS.pdf + '"><span>PDF</span>';
  }
}

// Cache for targeted thumbnail updates — avoids full innerHTML rebuild when only content changed
const _thumbCache = { pageIds: null, layoutStyle: null, fps: {} };
let _layoutsAccordionOpen = false;

function _thumbFP(page) {
  // Lightweight fingerprint: page background + all element props with large strings truncated
  return (page.paper || '') + '|' + (page.texture || '') + '|' + JSON.stringify(page.elements || [], function(k, v) {
    return typeof v === 'string' && v.length > 200 ? v.length + ':' + v.slice(0, 32) : v;
  });
}

function renderPagesPanel() {
  const panel = document.getElementById('pagesPanel');
  const cover = state.pages[0];

  // Ensure fonts used across all pages are requested before thumbnails render
  if (typeof loadFontsForPages === 'function') loadFontsForPages(state.pages);

  // --- Targeted update path: same page structure, patch only changed thumbnails ---
  const currentIds = state.pages.map(function(p) { return p.id; }).join(',');
  if (_thumbCache.pageIds === currentIds
      && _thumbCache.layoutStyle === state.layoutStyle
      && panel.querySelector('.thumb-content[data-page-id]')) {
    for (var pi = 0; pi < state.pages.length; pi++) {
      var page = state.pages[pi];
      var fp = _thumbFP(page);
      if (_thumbCache.fps[page.id] !== fp) {
        var thumbEl = panel.querySelector('.thumb-content[data-page-id="' + page.id + '"]');
        if (thumbEl) {
          thumbEl.innerHTML = renderSimpleThumbnail(page);
          // Update wrapper background in case paper color changed
          var wrapper = thumbEl.parentElement;
          if (wrapper) wrapper.style.background = page.paper || '#f5f3ee';
        }
        _thumbCache.fps[page.id] = fp;
      }
    }
    // Update active/selected classes (cheap DOM toggles)
    var coverEl = panel.querySelector('.thumb-cover');
    if (coverEl) coverEl.classList.toggle('active', state.currentPage === 'cover');
    panel.querySelectorAll('.thumb-spread').forEach(function(spreadEl) {
      var idx = parseInt(spreadEl.dataset.spreadIdx);
      var leftId = state.pages[idx * 2 + 1] ? state.pages[idx * 2 + 1].id : null;
      var rightId = state.pages[idx * 2 + 2] ? state.pages[idx * 2 + 2].id : null;
      spreadEl.classList.toggle('active', state.currentPage === leftId || state.currentPage === rightId);
      spreadEl.classList.toggle('selected-spread', state.selectedSpread === idx);
    });
    panel.querySelectorAll('.thumb-page[data-page-id]').forEach(function(pageEl) {
      pageEl.classList.toggle('active-thumb', state.activePage === pageEl.dataset.pageId);
    });
    return;
  }

  // --- Full rebuild path (first render or structural change) ---
  // Preserve scroll position so panel doesn't jump to top on add/delete
  const scrollTop = panel.scrollTop;

  let html = '';

  // Pages accordion
  const spreadCount = Math.floor((state.pages.length - 1) / 2);
  html += '<div class="accordion open">';
  html += `<div class="accordion-head" onclick="this.parentElement.classList.toggle('open')">Pages <span style="font-weight:400;color:#999;margin-left:4px">(${spreadCount} spread${spreadCount !== 1 ? 's' : ''} + cover)</span> <span>▼</span></div>`;
  html += '<div class="accordion-body" id="spreadsContainer">';

  // Cover thumbnail (not draggable)
  const coverTexture = cover.texture ? `paper-${cover.texture}` : '';
  html += `<div class="thumb-cover ${coverTexture} ${state.currentPage==='cover'?'active':''}" style="background:${cover.paper||'#f5f3ee'}" onclick="selectPage('cover')">`;
  html += `<div class="thumb-content" data-page-id="cover">${renderSimpleThumbnail(cover)}</div>`;
  html += '<div class="thumb-label">Cover</div></div>';

  // Spread thumbnails
  for (let i = 1; i < state.pages.length; i += 2) {
    const left = state.pages[i];
    const right = state.pages[i + 1];
    const isActive = state.currentPage === left?.id || state.currentPage === right?.id;
    const spreadIdx = Math.floor(i / 2);
    const isSelectedSpread = state.selectedSpread === spreadIdx;

    html += `<div class="thumb-spread ${isActive?'active':''} ${isSelectedSpread?'selected-spread':''}"
      data-spread-idx="${spreadIdx}"
      onmousedown="spreadMouseDown(event, ${spreadIdx})">`;
    if (left) {
      const leftTexture = left.texture ? `paper-${left.texture}` : '';
      html += `<div class="thumb-page ${state.activePage===left.id?'active-thumb':''} ${leftTexture}" data-page-id="${left.id}" style="background:${left.paper||'#f5f3ee'}" onclick="event.stopPropagation();selectPage('${left.id}')">`;
      let leftThumb = '';
      try { leftThumb = renderSimpleThumbnail(left); } catch(e) { console.warn('[thumb] left render error:', e); }
      html += `<div class="thumb-content" data-page-id="${left.id}">${leftThumb}</div>`;
      html += '</div>';
    }
    if (right) {
      const rightTexture = right.texture ? `paper-${right.texture}` : '';
      html += `<div class="thumb-page ${state.activePage===right.id?'active-thumb':''} ${rightTexture}" data-page-id="${right.id}" style="background:${right.paper||'#f5f3ee'}" onclick="event.stopPropagation();selectPage('${right.id}')">`;
      let rightThumb = '';
      try { rightThumb = renderSimpleThumbnail(right); } catch(e) { console.warn('[thumb] right render error:', e); }
      html += `<div class="thumb-content" data-page-id="${right.id}">${rightThumb}</div>`;
      html += '</div>';
    }
    html += `<button class="thumb-delete" data-spread-idx="${spreadIdx}" title="Delete spread (Cmd+Z to undo)" aria-label="Delete spread ${spreadIdx + 1}">×</button>`;
    html += `<div class="thumb-label">${left?.name||''}${right?' / '+right.name:''}</div>`;
    html += '</div>';
  }

  html += '</div></div>';

  // Layouts accordion
  html += '<div class="accordion' + (_layoutsAccordionOpen ? ' open' : '') + '">';
  html += '<div class="accordion-head" onclick="this.parentElement.classList.toggle(\'open\');_layoutsAccordionOpen=this.parentElement.classList.contains(\'open\')">Layouts <span>▼</span></div>';
  html += '<div class="accordion-body">';

  // Style tabs
  html += '<div class="style-tabs">';
  STYLE_TABS.forEach(tab => {
    html += `<div class="style-tab ${state.layoutStyle===tab.id?'active':''} ${tab.empty?'empty':''}" onclick="_layoutsAccordionOpen=true;setLayoutStyle('${tab.id}')">${tab.name}</div>`;
  });
  html += '</div>';

  // Layout grid — only show when a style tab is selected
  if (!state.layoutStyle) {
    html += '<div style="padding:15px;text-align:center;color:#999;font-size:10px">Choose a style above to browse layouts</div>';
  } else {
    const currentStyle = STYLE_TABS.find(t => t.id === state.layoutStyle);
    if (currentStyle?.empty) {
      html += '<div style="padding:15px;text-align:center;color:#888;font-size:10px;font-style:italic">Templates coming soon</div>';
    } else {
      const layouts = state.layoutStyle === 'all' ? LAYOUTS : LAYOUTS.filter(l => l.style === state.layoutStyle);
      html += '<div class="layout-grid">';
      layouts.forEach(layout => {
        html += `<div class="layout-item" onclick="applyLayout('${layout.id}')">`;
        html += `<div class="layout-preview">${renderLayoutPreview(layout)}</div>`;
        html += `<div class="layout-name">${layout.name}</div>`;
        html += '</div>';
      });
      html += '</div>';
    }
  }

  html += '</div></div>';

  panel.innerHTML = html;
  panel.scrollTop = scrollTop;

  // Populate fingerprint cache after full rebuild
  _thumbCache.pageIds = currentIds;
  _thumbCache.layoutStyle = state.layoutStyle;
  _thumbCache.fps = {};
  for (var ci = 0; ci < state.pages.length; ci++) {
    _thumbCache.fps[state.pages[ci].id] = _thumbFP(state.pages[ci]);
  }
  // Delete handlers are now delegated — see initDeleteSpreadDelegate() below.
  // No per-button attachment needed here.
}

// ============ DELEGATED DELETE-SPREAD HANDLER ============
// Attached once at init. Survives innerHTML rebuilds because the listener
// lives on #pagesPanel, not on individual .thumb-delete buttons.
let _deleteSpreadDelegateAttached = false;
function initDeleteSpreadDelegate() {
  if (_deleteSpreadDelegateAttached) return;
  const panel = document.getElementById('pagesPanel');
  if (!panel) return;
  _deleteSpreadDelegateAttached = true;

  panel.addEventListener('click', function(e) {
    const btn = e.target.closest('.thumb-delete');
    if (!btn) return;
    e.stopPropagation();
    const idx = parseInt(btn.dataset.spreadIdx, 10);
    console.log('[delete-delegate] click on spread idx:', idx);
    if (isNaN(idx)) return;
    deleteSpread(idx);
  });

  // Also intercept mousedown on delete buttons to prevent spread drag interference.
  // This replaces the inline onmousedown="event.stopPropagation()" on the button HTML.
  panel.addEventListener('mousedown', function(e) {
    if (e.target.closest('.thumb-delete')) {
      e.stopPropagation();
    }
  });
}

function spreadMouseDown(e, spreadIdx) {
  // Don't start drag if clicking delete button or page thumbnails
  if (e.target.closest('.thumb-delete')) return;
  if (e.target.classList.contains('thumb-page')) return;
  
  // Select the spread
  state.selectedSpread = spreadIdx;
  const pageIdx = spreadIdx * 2 + 1;
  if (state.pages[pageIdx]) {
    state.currentPage = state.pages[pageIdx].id;
    state.activePage = state.pages[pageIdx].id;
  }
  state.selected = null;
  
  // Update visual selection state directly without destroying DOM
  document.querySelectorAll('.thumb-spread').forEach((el, idx) => {
    el.classList.toggle('selected-spread', idx === spreadIdx);
    el.classList.toggle('active', 
      state.currentPage === state.pages[pageIdx]?.id || 
      state.activePage === state.pages[pageIdx]?.id
    );
  });
  
  // Setup for potential drag
  spreadDragState.startY = e.clientY;
  spreadDragState.spreadIdx = spreadIdx;
  spreadDragState.element = e.target.closest('.thumb-spread');
  spreadDragState.isDragging = false;
  
  document.addEventListener('mousemove', spreadMouseMove);
  document.addEventListener('mouseup', spreadMouseUp);
  // Safety: clean up listeners if mouse leaves window entirely (prevents phantom drags)
  document.addEventListener('mouseleave', spreadMouseUpSafety);
  // DO NOT call render() here - it destroys the DOM mid-click and swallows delete/add clicks
}

function spreadMouseUpSafety(e) {
  // Only fire if mouse actually left the document (not just an element)
  if (e.relatedTarget === null) {
    document.removeEventListener('mouseleave', spreadMouseUpSafety);
    spreadMouseUp(e);
  }
}

function spreadMouseMove(e) {
  const dy = Math.abs(e.clientY - spreadDragState.startY);
  
  // Start dragging after 5px movement
  if (!spreadDragState.isDragging && dy > 5) {
    spreadDragState.isDragging = true;
    if (spreadDragState.element) {
      spreadDragState.element.classList.add('dragging');
    }
  }
  
  if (!spreadDragState.isDragging) return;
  
  // Find which spread we're over
  const spreads = document.querySelectorAll('.thumb-spread');
  spreads.forEach(spread => {
    spread.classList.remove('drag-target-above', 'drag-target-below');
    
    if (spread === spreadDragState.element) return;
    
    const rect = spread.getBoundingClientRect();
    if (e.clientY >= rect.top && e.clientY <= rect.bottom) {
      const midY = rect.top + rect.height / 2;
      if (e.clientY < midY) {
        spread.classList.add('drag-target-above');
      } else {
        spread.classList.add('drag-target-below');
      }
    }
  });
}

function spreadMouseUp(e) {
  document.removeEventListener('mousemove', spreadMouseMove);
  document.removeEventListener('mouseup', spreadMouseUp);
  document.removeEventListener('mouseleave', spreadMouseUpSafety);
  
  if (!spreadDragState.isDragging) {
    spreadDragState.element = null;
    spreadDragState.spreadIdx = null;
    return;
  }
  
  // Find drop target
  const spreads = document.querySelectorAll('.thumb-spread');
  let targetSpreadIdx = null;
  let dropBelow = false;
  
  spreads.forEach(spread => {
    if (spread.classList.contains('drag-target-above')) {
      targetSpreadIdx = parseInt(spread.dataset.spreadIdx);
      dropBelow = false;
    } else if (spread.classList.contains('drag-target-below')) {
      targetSpreadIdx = parseInt(spread.dataset.spreadIdx);
      dropBelow = true;
    }
    spread.classList.remove('drag-target-above', 'drag-target-below', 'dragging');
  });
  
  const fromSpreadIdx = spreadDragState.spreadIdx;
  
  // Reset state
  spreadDragState.isDragging = false;
  spreadDragState.element = null;
  spreadDragState.spreadIdx = null;
  
  // Perform the move if valid
  if (targetSpreadIdx !== null && targetSpreadIdx !== fromSpreadIdx) {
    moveSpread(fromSpreadIdx, targetSpreadIdx, dropBelow);
  }
}



// ============ TOOLS PANEL ============
let _lastToolsFingerprint = '';

function renderToolsPanel() {
  // Skip re-render while color picker slider is being dragged — prevents destroying the slider mid-drag
  if (window._colorPickerActive) return;
  // Skip re-render if user is actively typing in a tools panel input (hex color, font size, etc.)
  const activeEl = document.activeElement;
  if (activeEl && activeEl.closest('#toolsPanel') && (activeEl.tagName === 'INPUT' || activeEl.tagName === 'SELECT')) return;
  const panel = document.getElementById('toolsPanel');
  const el = getSelectedElement();

  // Phase 4 — fingerprint gate: skip full innerHTML rebuild when inputs haven't changed
  const pg = getActivePage();
  const fp = (state.selected || '') + '::'
    + (el ? [el.t, el.fontFamily, el.fontSize, el.bold, el.italic, el.align,
             el.color, el.opacity, el.exposure, el.contrast, el.saturation,
             el.border, el.fitMode, el.shape, el.rippedFrame,
             el.deepEtched ? 1 : 0, el.revealMask ? 1 : 0,
             el.isBox ? 1 : 0, el.targetWords,
             el.txt ? el.txt.length : 0].join('|') : '')
    + '::' + (pg ? pg.id + '|' + (pg.paper || '') + '|' + (pg.texture || '') : '')
    + '::' + activeStickerTab
    + '::' + (drawState.active ? 1 : 0)
    + '::' + savedColors.join('|');
  if (fp === _lastToolsFingerprint) return;
  _lastToolsFingerprint = fp;
  
  // Clear manual section toggles when selection changes
  const currentSelId = el ? el.id : null;
  if (currentSelId !== _lastSelectedForToggles) {
    state.sectionToggles = {};
    _lastSelectedForToggles = currentSelId;
  }
  
  const isText = el && el.t === 'text';
  const isImage = el && el.t === 'image';
  const isShape = el && el.t === 'shape';
  const isGraphic = el && el.t === 'graphic';
  const isDrawing = el && el.t === 'drawing';
  
  // Determine which section should be expanded based on selection
  const expandedSection = isText ? 'text' : isImage ? 'image' : (isShape || isGraphic || isDrawing) ? 'elements' : null;
  
  let html = '';
  
  // TEXT SECTION
  html += renderTextSection(el, isText, expandedSection === 'text');
  
  // IMAGE SECTION
  html += renderImageSection(el, isImage, expandedSection === 'image');
  
  // ELEMENTS SECTION (Shapes + Graphics + Draw combined)
  html += renderElementsSection(el, isShape, expandedSection === 'elements');
  
  // STICKERS SECTION
  html += renderStickersSection(expandedSection);
  
  // FRAMES SECTION
  html += renderFramesSection(expandedSection);
  
  // PAGE SECTION
  const activePage = getActivePage();
  if (activePage) {
    let pageCollapsed;
    if (state.sectionToggles['sec-page'] !== undefined) {
      pageCollapsed = state.sectionToggles['sec-page'] === 'collapsed' ? 'collapsed' : '';
    } else {
      pageCollapsed = 'collapsed';
    }
    html += `<div class="tool-section ${pageCollapsed}" id="sec-page">`;
    html += `<div class="tool-title" onclick="toggleSection('sec-page')"><span class="tool-icon"><img src="${DADA_ICONS.page}" style="background:#e8e4ce"></span><span class="tool-name">Page</span></div>`;
    html += `<div class="tool-content">`;
    
    // Background color
    html += `<div class="tool-row"><label class="tool-label">Background Color</label></div>`;
    html += renderColorPicker(activePage.paper || '#f5f3ee', 'setPageColor', false);
    
    // Paper texture
    html += `<div class="tool-row" style="margin-top:12px"><label class="tool-label">Paper Texture</label>`;
    html += `<select class="tool-input" onchange="setPageTexture(this.value)" style="cursor:pointer">`;
    const textures = [
      {id: 'smooth', name: 'Smooth'},
      {id: 'matte', name: 'Matte'},
      {id: 'newsprint', name: 'Newsprint'},
      {id: 'cardstock', name: 'Cardstock'},
      {id: 'riso', name: 'Risograph'},
      {id: 'distressed', name: 'Distressed'},
      {id: 'patina', name: 'Ink Patina'},
      {id: 'saigon', name: 'Linen Stock'}
    ];
    textures.forEach(t => {
      const selected = (activePage.texture || 'smooth') === t.id ? 'selected' : '';
      html += `<option value="${t.id}" ${selected}>${t.name}</option>`;
    });
    html += `</select></div>`;
    
    html += `</div></div>`;
  }
  
  // DELETE BUTTON (when element selected)
  if (el) {
    html += `<div class="divider"></div>`;
    html += `<button class="tool-btn" style="color:#c00;border-color:#fcc" onclick="deleteElement()">Delete Element</button>`;
  }
  
  panel.innerHTML = html;

  // Re-attach the persistent font list (built once, survives innerHTML rebuilds)
  const activeFamily = (isText && el) ? (el.fontFamily || 'Inter,sans-serif') : null;
  initFontList(activeFamily);
}

function renderTextSection(el, isText, isExpanded) {
  const textEl = isText ? el : null;
  const disabled = !isText ? 'disabled' : '';
  // Manual toggle overrides auto-logic
  let collapsed;
  if (state.sectionToggles['sec-text'] !== undefined) {
    collapsed = state.sectionToggles['sec-text'] === 'collapsed' ? 'collapsed' : '';
  } else {
    collapsed = (isExpanded === true) ? '' : 'collapsed';
  }
  
  // Build summary for collapsed view
  const summary = textEl ? `${(textEl.fontFamily || 'Inter').split(',')[0]}, ${textEl.fontSize || 12}px` : '';
  
  let html = `<div class="tool-section ${collapsed}" id="sec-text">`;
  html += `<div class="tool-title" onclick="toggleSection('sec-text')"><span class="tool-icon"><img src="${DADA_ICONS.text}" style="background:#d1d1cf"></span><span class="tool-name">Text</span>${summary ? `<span class="tool-summary">${summary}</span>` : ''}</div>`;
  html += `<div class="tool-content">`;
  
  // Insert buttons - ALWAYS enabled
  html += `<div class="tool-grid c2" style="margin-bottom:10px">`;
  html += `<button onclick="addText()">+ Text</button>`;
  html += `<button onclick="addTextBox()">+ Text Box</button>`;
  html += `</div>`;
  
  // Word count (if text box)
  if (textEl?.isBox) {
    const wordCount = countWords(textEl.txt || '');
    html += `<div style="background:#f5f5f5;padding:6px;border-radius:4px;font-size:10px;margin-bottom:10px">📝 ${wordCount}${textEl.targetWords ? ' / ' + textEl.targetWords : ''} words</div>`;
    html += `<div class="tool-row"><label class="tool-label">Target Words</label>`;
    html += `<div style="display:flex;gap:4px"><input type="number" class="tool-input" style="width:60px" value="${textEl.targetWords||''}" id="targetWords"><button class="btn" onclick="fillLorem()">Fill</button></div></div>`;
  }
  
  // Font selector - greyed when no text selected
  const fontDisabled = !isText ? 'section-disabled' : '';
  html += `<div class="${fontDisabled}">`;
  html += `<div class="tool-row"><label class="tool-label">Font ${!isText ? '<span style="font-weight:normal;color:#999">(select text)</span>' : ''}</label><div class="font-list" id="persistent-font-list"></div></div>`;
  
  // Size
  html += `<div class="tool-row"><label class="tool-label">Size</label>`;
  html += `<input type="number" class="tool-input" style="width:60px" value="${textEl?.fontSize || 12}" oninput="setFontSize(this.value)" onchange="setFontSize(this.value)" ${disabled ? 'disabled' : ''}></div>`;
  
  // Style (Bold/Italic)
  html += `<div class="tool-row"><label class="tool-label">Style</label><div class="format-row">`;
  html += `<button class="format-btn ${textEl?.bold ? 'active' : ''}" onclick="toggleBold()" ${disabled ? 'disabled' : ''} aria-label="Bold" aria-pressed="${!!textEl?.bold}"><b>B</b></button>`;
  html += `<button class="format-btn ${textEl?.italic ? 'active' : ''}" onclick="toggleItalic()" ${disabled ? 'disabled' : ''} aria-label="Italic" aria-pressed="${!!textEl?.italic}"><i>I</i></button>`;
  html += `</div></div>`;
  
  // Alignment
  html += `<div class="tool-row"><label class="tool-label">Align</label><div class="format-row">`;
  ['left','center','right','justify'].forEach(a => {
    const icon = a === 'left' ? '≡' : a === 'center' ? '☰' : a === 'right' ? '≡' : '⊞';
    const active = textEl && (textEl.align || 'left') === a ? 'active' : '';
    html += `<button class="format-btn ${active}" onclick="setAlign('${a}')" title="${a}" aria-label="Align ${a}" ${disabled ? 'disabled' : ''}>${icon}</button>`;
  });
  html += `</div></div>`;
  
  // Color with wheel
  html += `<div class="tool-row"><label class="tool-label">Color</label>`;
  html += renderColorPicker(textEl?.color || '#111111', 'setTextColor', disabled);
  html += `</div>`;
  html += `</div>`; // close fontDisabled div
  
  html += `</div></div>`;
  return html;
}

// ============ PERSISTENT FONT LIST ============
// The font list DOM is built once and re-appended on each renderToolsPanel() call
// to avoid rebuilding 89 styled font items on every render cycle.
let _fontListNode = null;
let _fontListObserver = null;

function buildFontListHTML(activeFamily) {
  let html = '';
  FONTS.forEach(f => {
    const active = activeFamily && activeFamily === f.family ? 'active' : '';
    html += `<div class="font-item ${active}" onclick="setFont('${f.family}')" style="font-family:sans-serif" data-family="${f.family}">${f.name}</div>`;
  });
  return html;
}

function initFontList(activeFamily) {
  const container = document.getElementById('persistent-font-list');
  if (!container) return;

  if (_fontListNode && _fontListNode.children.length > 0) {
    // _fontListNode is the orphaned old container (destroyed by innerHTML).
    // Move its children into the new container — no rebuild, just pointer moves.
    while (_fontListNode.firstChild) {
      container.appendChild(_fontListNode.firstChild);
    }
  } else {
    // First time — build the font items directly into the container
    container.innerHTML = buildFontListHTML(activeFamily);
  }
  // Store reference to current container (will become orphan on next innerHTML rebuild)
  _fontListNode = container;
  updateFontListActive(activeFamily);

  // Set up IntersectionObserver for lazy font preview loading.
  // Must recreate each time because the observer root (container) changes on every
  // renderToolsPanel() innerHTML rebuild — the old root becomes an orphaned DOM node.
  if (_fontListObserver) _fontListObserver.disconnect();
  _fontListObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const item = entry.target;
        const family = item.dataset.family;
        if (family) {
          ensureFontLoaded(family);
          item.style.fontFamily = family;
          _fontListObserver.unobserve(item);
        }
      }
    });
  }, { root: container, rootMargin: '50px' });
  // Observe only items that haven't loaded their preview yet
  container.querySelectorAll('.font-item').forEach(item => {
    if (item.style.fontFamily === 'sans-serif') {
      _fontListObserver.observe(item);
    }
  });
}

function updateFontListActive(activeFamily) {
  if (!_fontListNode) return;
  const prev = _fontListNode.querySelector('.font-item.active');
  if (prev) prev.classList.remove('active');
  if (activeFamily) {
    const items = _fontListNode.querySelectorAll('.font-item');
    for (let i = 0; i < items.length; i++) {
      if (items[i].getAttribute('onclick') === "setFont('" + activeFamily + "')") {
        items[i].classList.add('active');
        break;
      }
    }
  }
}

function renderImageSection(el, isImage, isExpanded) {
  const imgEl = isImage ? el : null;
  const disabled = !isImage ? 'disabled' : '';
  const hasImage = imgEl?.src;
  // If nothing selected (isExpanded undefined/null), keep open. Otherwise collapse non-matching sections
  let collapsed;
  if (state.sectionToggles['sec-image'] !== undefined) {
    collapsed = state.sectionToggles['sec-image'] === 'collapsed' ? 'collapsed' : '';
  } else {
    collapsed = (isExpanded === true) ? '' : 'collapsed';
  }
  
  // Build summary
  let summary = '';
  if (imgEl) {
    const parts = [];
    if (imgEl.rippedFrame) {
      const fd = RIPPED_FRAMES.find(f => f.id === imgEl.rippedFrame);
      if (fd) parts.push(fd.name);
    }
    if (imgEl.border && imgEl.border.startsWith('shadow')) parts.push(imgEl.border.replace('shadow-','shadow '));
    if (imgEl.border === 'float') parts.push('float');
    summary = parts.join(', ');
  }
  
  let html = `<div class="tool-section ${collapsed}" id="sec-image">`;
  html += `<div class="tool-title" onclick="toggleSection('sec-image')"><span class="tool-icon"><img src="${DADA_ICONS.image}" style="background:#c5c2af"></span><span class="tool-name">Image</span>${summary ? `<span class="tool-summary">${summary}</span>` : ''}</div>`;
  html += `<div class="tool-content">`;
  
  // Insert button - ALWAYS enabled
  html += `<button class="tool-btn" style="margin-bottom:10px" onclick="addImage()">+ Image</button>`;
  
  // Options - greyed when no image selected
  const optionsDisabled = !isImage ? 'section-disabled' : '';
  html += `<div class="${optionsDisabled}">`;
  if (!isImage) {
    html += `<p style="font-size:9px;color:#999;margin-bottom:8px">Select an image to edit</p>`;
  }

  // Fit/Fill toggle
  const currentFitMode = imgEl?.fitMode || 'cover';
  const isFit = currentFitMode === 'contain';
  html += `<div class="tool-row"><label class="tool-label">Image Fit</label><div class="tool-grid c2">`;
  html += `<button class="${!isFit ? 'active' : ''}" onclick="setFitMode('cover')" title="Fill frame completely (may crop edges)">Fill</button>`;
  html += `<button class="${isFit ? 'active' : ''}" onclick="setFitMode('contain')" title="Show whole image (may letterbox)">Fit</button>`;
  html += `</div></div>`;
  
  // Flip buttons
  html += `<div class="tool-row"><label class="tool-label">Flip</label><div class="tool-grid c2">`;
  html += `<button onclick="flipImage('h')">↔ Horiz</button>`;
  html += `<button onclick="flipImage('v')">↕ Vert</button>`;
  html += `</div></div>`;
  
  // Reset image position button
  html += `<button class="tool-btn" onclick="resetImagePosition()" style="margin-top:6px">↺ Reset Position</button>`;
  
  // Opacity slider
  html += `<div class="tool-row" style="margin-top:8px"><label class="tool-label">Opacity</label>`;
  html += `<div class="slider-row"><input type="range" min="10" max="100" value="${Math.round((imgEl?.opacity ?? 1) * 100)}" oninput="setImageOpacity(this.value)"><span>${Math.round((imgEl?.opacity ?? 1) * 100)}%</span></div></div>`;

  // Image Adjustments
  html += `<div class="tool-row" style="margin-top:4px"><label class="tool-label">Exposure</label>`;
  html += `<div class="slider-row"><input type="range" min="50" max="150" value="${imgEl?.exposure ?? 100}" oninput="setImageAdjust('exposure',this.value)" onchange="commitImageAdjust()"></div></div>`;

  html += `<div class="tool-row"><label class="tool-label">Contrast</label>`;
  html += `<div class="slider-row"><input type="range" min="50" max="150" value="${imgEl?.contrast ?? 100}" oninput="setImageAdjust('contrast',this.value)" onchange="commitImageAdjust()"></div></div>`;

  html += `<div class="tool-row"><label class="tool-label">Saturation</label>`;
  html += `<div class="slider-row"><input type="range" min="0" max="200" value="${imgEl?.saturation ?? 100}" oninput="setImageAdjust('saturation',this.value)" onchange="commitImageAdjust()"></div></div>`;

  // Shadow
  html += `<div class="divider"></div>`;
  html += `<div class="tool-row"><label class="tool-label">Shadow</label><div class="tool-grid c2">`;
  [['shadow-left','Left'],['shadow-right','Right'],['shadow-bottom','Bottom'],['float','Float']].forEach(([b,n]) => {
    const active = imgEl?.border === b ? 'active' : '';
    html += `<button class="${active}" onclick="setBorder('${b}')">${n}</button>`;
  });
  html += `</div></div>`;

  // Cut Out
  html += `<div class="divider"></div>`;
  html += `<button class="tool-btn" onclick="openCutout()">✂️ Cut Out / Remove BG</button>`;
  if (imgEl?.deepEtched) {
    html += `<button class="tool-btn" onclick="restoreBackground()" style="margin-top:6px">↩️ Restore Original</button>`;
  }
  
  // Reveal Brush
  html += `<div class="divider"></div>`;
  html += `<button class="tool-btn" onclick="openRevealBrush()">✨ Reveal Brush</button>`;
  html += `<p style="font-size:9px;color:#888;margin-top:4px">Paint to reveal image above other elements (like text). Tip: Hold R for quick access.</p>`;
  if (imgEl?.revealMask) {
    html += `<button class="tool-btn" onclick="clearRevealMask()" style="margin-top:6px">Clear Reveal Mask</button>`;
  }

  html += `</div>`; // close optionsDisabled div
  
  html += `</div></div>`;
  return html;
}

function renderElementsSection(el, isShape, isExpanded) {
  const shapeEl = isShape ? el : null;
  const isGraphic = el && el.t === 'graphic';
  const isDrawing = el && el.t === 'drawing';
  const isElementType = isShape || isGraphic || isDrawing;
  const disabled = !isShape ? 'disabled' : '';
  
  // Collapsed logic
  let collapsed;
  if (state.sectionToggles['sec-elements'] !== undefined) {
    collapsed = state.sectionToggles['sec-elements'] === 'collapsed' ? 'collapsed' : '';
  } else {
    collapsed = (isExpanded === true) ? '' : 'collapsed';
  }
  
  // Summary
  let summary = '';
  if (shapeEl) summary = shapeEl.shape;
  else if (isGraphic) summary = 'graphic';
  else if (isDrawing) summary = 'drawing';
  
  let html = `<div class="tool-section ${collapsed}" id="sec-elements">`;
  html += `<div class="tool-title" onclick="toggleSection('sec-elements')"><span class="tool-icon"><img src="${DADA_ICONS.elements}" style="background:#d4d1cc"></span><span class="tool-name">Elements</span>${summary ? `<span class="tool-summary">${summary}</span>` : ''}</div>`;
  html += `<div class="tool-content">`;
  
  // SHAPES subsection
  html += `<div class="tool-row"><label class="tool-label">Shapes</label><div class="tool-grid c5">`;
  SHAPES.forEach(s => {
    const active = shapeEl?.shape === s.id ? 'active' : '';
    html += `<button class="${active}" onclick="addShape('${s.id}')">${s.icon}</button>`;
  });
  html += `</div></div>`;
  
  // Shape color picker — always visible under Shapes
  html += `<div class="tool-row"><label class="tool-label">Shape Color</label>`;
  html += renderColorPicker(shapeEl?.color || '#4A3F2A', 'setShapeColor', !isShape);
  html += `</div>`;
  
  // DRAW subsection
  html += `<div class="divider"></div>`;
  html += `<div class="tool-row"><label class="tool-label">Drawing</label></div>`;
  html += `<button class="tool-btn" onclick="toggleDraw()">${drawState.active ? '✏️ Exit Draw Mode' : '✏️ Start Drawing'}</button>`;
  
  html += `</div></div>`;
  return html;
}

function renderStickersSection(expandedSection) {
  let collapsed;
  if (state.sectionToggles['sec-stickers'] !== undefined) {
    collapsed = state.sectionToggles['sec-stickers'] === 'collapsed' ? 'collapsed' : '';
  } else {
    collapsed = expandedSection ? 'collapsed' : 'collapsed';
  }
  let html = `<div class="tool-section ${collapsed}" id="sec-stickers">`;
  html += `<div class="tool-title" onclick="toggleSection('sec-stickers')"><span class="tool-icon" style="background:#e8ddd4;width:24px;height:24px;border-radius:4px;display:inline-flex;align-items:center;justify-content:center;font-size:14px">✦</span><span class="tool-name">Stickers</span></div>`;
  html += `<div class="tool-content">`;
  
  // Tabs
  html += `<div class="sticker-tabs">`;
  STICKER_PACKS.forEach(pack => {
    const active = activeStickerTab === pack.id ? 'active' : '';
    html += `<button class="${active}" onclick="switchStickerTab('${pack.id}')">${pack.label}</button>`;
  });
  html += `</div>`;
  
  // Grid for active tab
  const pack = STICKER_PACKS.find(p => p.id === activeStickerTab);
  if (pack) {
    html += `<div class="sticker-grid-wrap"><div class="sticker-grid">`;
    pack.files.forEach(file => {
      const src = `stickers/${pack.id}/${file}`;
      html += `<img class="sticker-thumb" src="${src}" loading="lazy" draggable="true" onclick="addSticker('${pack.id}','${file}')" ondragstart="event.dataTransfer.setData('sticker','${pack.id}/${file}')" alt="" title="${file.replace('.png','')}">`;
    });
    html += `</div></div>`;
  }
  
  html += `</div></div>`;
  return html;
}

function switchStickerTab(tabId) {
  activeStickerTab = tabId;
  renderToolsPanel();
}

function renderFramesSection(expandedSection) {
  let collapsed;
  if (state.sectionToggles['sec-frames'] !== undefined) {
    collapsed = state.sectionToggles['sec-frames'] === 'collapsed' ? 'collapsed' : '';
  } else {
    collapsed = 'collapsed';
  }
  let html = `<div class="tool-section ${collapsed}" id="sec-frames">`;
  html += `<div class="tool-title" onclick="toggleSection('sec-frames')"><span class="tool-icon" style="background:#d4cfc4;width:24px;height:24px;border-radius:4px;display:inline-flex;align-items:center;justify-content:center;font-size:13px">▨</span><span class="tool-name">Frames</span></div>`;
  html += `<div class="tool-content">`;
  html += `<p style="font-size:9px;color:#888;margin-bottom:8px">Ripped paper frames — drop an image inside</p>`;
  html += `<div class="frame-grid">`;
  RIPPED_FRAMES.forEach(frame => {
    html += `<div class="frame-thumb" onclick="addRippedFrame('${frame.id}')" title="${frame.name}">`;
    if (frame.type === 'torn' || frame.type === 'photo-frame') {
      html += `<img src="${getFrameSrc(frame)}" style="width:100%;height:100%;object-fit:contain;" loading="lazy">`;
    } else if (frame.type === 'polaroid') {
      const txtCol = frame.color === '#1a1a1a' ? '#666' : '#bbb';
      html += `<svg viewBox="0 0 100 120" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="2" width="96" height="116" rx="2" fill="${frame.color}" stroke="#ccc" stroke-width="1"/><rect x="8" y="8" width="84" height="80" fill="#d4d0c8"/><text x="50" y="106" text-anchor="middle" font-size="8" fill="${txtCol}">${frame.name}</text></svg>`;
    }
    html += `</div>`;
  });
  html += `</div>`;
  html += `</div></div>`;
  return html;
}


// Saved colors (persisted in localStorage if available)

// Color usage tracking - tracks how often each color is used
let colorUsage = {};


// Current color picker state
let colorPickerState = {
  callback: null,
  hue: 0,
  sat: 100,
  val: 100
};

function renderColorPicker(currentColor, callback, disabled) {
  if (disabled) {
    return `<div class="color-picker-container" style="opacity:0.5;pointer-events:none">
      <div class="color-gradient" style="background:#ccc"></div>
      <p style="font-size:10px;color:#999;text-align:center">Select element to edit color</p>
    </div>`;
  }
  
  // Parse current color to HSV
  const hsv = hexToHsv(currentColor || '#000000');
  
  const pickerId = 'picker_' + callback;
  
  let html = `<div class="color-picker-container" data-callback="${callback}" data-picker-id="${pickerId}">`;
  
  // Saturation/Brightness gradient
  const hueColor = hsvToHex(hsv.h, 100, 100);
  html += `<div class="color-gradient" id="${pickerId}_gradient" 
    style="background: linear-gradient(to right, #fff, ${hueColor}), linear-gradient(to top, #000, transparent); background-blend-mode: multiply;"
    onmousedown="startGradientDrag(event, '${callback}')"
    ontouchstart="startGradientDrag(event, '${callback}')">`;
  
  // Handle position
  const handleX = hsv.s;
  const handleY = 100 - hsv.v;
  html += `<div class="color-gradient-handle" id="${pickerId}_handle" style="left:${handleX}%;top:${handleY}%;background:${currentColor}"></div>`;
  html += `</div>`;
  
  // Hue slider
  html += `<div class="hue-slider-wrap">`;
  html += `<input type="range" class="hue-slider" id="${pickerId}_hue" min="0" max="360" value="${hsv.h}" 
    oninput="updateHue('${callback}', this.value)"
    onmousedown="window._colorPickerActive=true;event.stopPropagation()"
    onmouseup="window._colorPickerActive=false;if(window._textColorDirty||window._pageColorDirty||window._shapeColorDirty){window._textColorDirty=false;window._pageColorDirty=false;window._shapeColorDirty=false;render();}renderToolsPanel()"
    ontouchstart="window._colorPickerActive=true;event.stopPropagation()"
    ontouchend="window._colorPickerActive=false;if(window._textColorDirty||window._pageColorDirty||window._shapeColorDirty){window._textColorDirty=false;window._pageColorDirty=false;window._shapeColorDirty=false;render();}renderToolsPanel()">`;
  html += `</div>`;
  
  // Hex input
  html += `<div class="color-inputs-row">`;
  html += `<input type="text" class="color-hex-input" id="${pickerId}_hex" value="${currentColor}" 
    onchange="setColorFromHex('${callback}', this.value)" 
    onkeydown="if(event.key==='Enter')setColorFromHex('${callback}', this.value)">`;
  html += `</div>`;
  
  // Saved colors
  html += `<div class="saved-colors-row">`;
  html += `<span class="saved-colors-label">Saved:</span>`;
  savedColors.forEach((c, i) => {
    html += `<div class="color-swatch" style="background:${c};width:18px;height:18px" onclick="${callback}('${c}')" title="${c}"></div>`;
  });
  html += `<button class="add-color-btn" onclick="saveCurrentColor('${callback}', '${currentColor}')" title="Save color">+</button>`;
  html += `</div>`;
  
  html += `</div>`;
  return html;
}

function hexToHsv(hex) {
  let r = parseInt(hex.slice(1,3), 16) / 255;
  let g = parseInt(hex.slice(3,5), 16) / 255;
  let b = parseInt(hex.slice(5,7), 16) / 255;
  
  let max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0, v = max;
  let d = max - min;
  s = max === 0 ? 0 : d / max;
  
  if (max !== min) {
    switch (max) {
      case r: h = (g - b) / d + (g < b ? 6 : 0); break;
      case g: h = (b - r) / d + 2; break;
      case b: h = (r - g) / d + 4; break;
    }
    h /= 6;
  }
  
  return { h: Math.round(h * 360), s: Math.round(s * 100), v: Math.round(v * 100) };
}

function hsvToHex(h, s, v) {
  h = h / 360;
  s = s / 100;
  v = v / 100;
  
  let r, g, b;
  const i = Math.floor(h * 6);
  const f = h * 6 - i;
  const p = v * (1 - s);
  const q = v * (1 - f * s);
  const t = v * (1 - (1 - f) * s);
  
  switch (i % 6) {
    case 0: r = v; g = t; b = p; break;
    case 1: r = q; g = v; b = p; break;
    case 2: r = p; g = v; b = t; break;
    case 3: r = p; g = q; b = v; break;
    case 4: r = t; g = p; b = v; break;
    case 5: r = v; g = p; b = q; break;
  }
  
  const toHex = x => {
    const hex = Math.round(x * 255).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  };
  
  return '#' + toHex(r) + toHex(g) + toHex(b);
}

function startGradientDrag(e, callback) {
  e.preventDefault();
  window._colorPickerActive = true;
  const gradient = e.target.closest('.color-gradient');
  if (!gradient) return;
  
  const pickerId = gradient.id.replace('_gradient', '');
  const hueSlider = document.getElementById(pickerId + '_hue');
  const hue = parseInt(hueSlider?.value || 0);
  
  function updateFromPosition(clientX, clientY) {
    const rect = gradient.getBoundingClientRect();
    let x = (clientX - rect.left) / rect.width * 100;
    let y = (clientY - rect.top) / rect.height * 100;
    x = Math.max(0, Math.min(100, x));
    y = Math.max(0, Math.min(100, y));
    
    const sat = x;
    const val = 100 - y;
    const color = hsvToHex(hue, sat, val);
    
    // Update handle
    const handle = document.getElementById(pickerId + '_handle');
    if (handle) {
      handle.style.left = x + '%';
      handle.style.top = y + '%';
      handle.style.background = color;
    }
    
    // Update hex input
    const hexInput = document.getElementById(pickerId + '_hex');
    if (hexInput) hexInput.value = color;
    
    // Call the callback
    window[callback](color);
  }
  
  const touch = e.touches ? e.touches[0] : e;
  updateFromPosition(touch.clientX, touch.clientY);
  
  function onMove(ev) {
    ev.preventDefault();
    const t = ev.touches ? ev.touches[0] : ev;
    updateFromPosition(t.clientX, t.clientY);
  }
  
  function onUp() {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    document.removeEventListener('touchmove', onMove);
    document.removeEventListener('touchend', onUp);
    document.removeEventListener('mouseleave', onUp);
    window._colorPickerActive = false;
    // If color functions skipped render() during drag, do a single render now
    if (window._textColorDirty || window._pageColorDirty || window._shapeColorDirty) { window._textColorDirty = false; window._pageColorDirty = false; window._shapeColorDirty = false; render(); }
    renderToolsPanel();
  }

  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
  document.addEventListener('touchmove', onMove, { passive: false });
  document.addEventListener('touchend', onUp);
  // Safety: if mouse leaves window, end the drag to prevent stuck flag
  document.addEventListener('mouseleave', onUp);
}

function updateHue(callback, hue) {
  const container = document.querySelector(`[data-callback="${callback}"]`);
  if (!container) return;
  
  const pickerId = container.dataset.pickerId;
  const gradient = document.getElementById(pickerId + '_gradient');
  const handle = document.getElementById(pickerId + '_handle');
  const hexInput = document.getElementById(pickerId + '_hex');
  
  // Get current sat/val from handle position
  let sat = 100, val = 100;
  if (handle && gradient) {
    const rect = gradient.getBoundingClientRect();
    sat = parseFloat(handle.style.left) || 100;
    val = 100 - (parseFloat(handle.style.top) || 0);
  }
  
  const hueColor = hsvToHex(parseInt(hue), 100, 100);
  if (gradient) {
    gradient.style.background = `linear-gradient(to right, #fff, ${hueColor}), linear-gradient(to top, #000, transparent)`;
    gradient.style.backgroundBlendMode = 'multiply';
  }
  
  const color = hsvToHex(parseInt(hue), sat, val);
  if (handle) handle.style.background = color;
  if (hexInput) hexInput.value = color;
  
  window[callback](color);
}

function setColorFromHex(callback, hex) {
  // Validate and normalize hex
  hex = hex.trim();
  if (!hex.startsWith('#')) hex = '#' + hex;
  if (!/^#[0-9A-Fa-f]{6}$/.test(hex)) return;
  
  window[callback](hex);
  renderToolsPanel(); // Re-render to update picker position
}

function saveCurrentColor(callback, color) {
  if (!savedColors.includes(color)) {
    savedColors.push(color);
    if (savedColors.length > 12) savedColors.shift(); // Keep max 12
    renderToolsPanel();
  }
}

function toggleSection(sectionId) {
  const section = document.getElementById(sectionId);
  if (section) {
    const isNowCollapsed = !section.classList.contains('collapsed');
    section.classList.toggle('collapsed');
    state.sectionToggles[sectionId] = isNowCollapsed ? 'collapsed' : 'open';
  }
}

// ============ CONTEXT MENU ============
function showContextMenu(x, y) {
  const menu = document.getElementById('ctxMenu');
  menu.style.left = x + 'px';
  menu.style.top = y + 'px';
  menu.classList.add('show');
  
  const el = getSelectedElement();
  document.getElementById('ctxCutout').style.display = (el?.t === 'image' && el?.src) ? 'block' : 'none';
}

document.addEventListener('mousedown', e => {
  // Close context menu if clicking outside it
  if (!e.target.closest('.ctx-menu')) {
    document.getElementById('ctxMenu').classList.remove('show');
  }
});

document.addEventListener('click', e => {
  // Close context menu if clicking outside it
  if (!e.target.closest('.ctx-menu')) {
    document.getElementById('ctxMenu').classList.remove('show');
  }
});
