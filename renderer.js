// ============ MAIN RENDER ============
function render() {
  // Skip re-render while draw mode is active and canvas already exists — stage rebuild would wipe drawn content.
  // The first render after entering draw mode is allowed through (canvas not yet in DOM) to create the canvas.
  if (drawState.active && document.getElementById('drawCanvas')) return;
  try {
    const _r0 = performance.now();
    let _r1 = _r0;
    if (_needsPagesPanelUpdate) {
      renderPagesPanel();
      _needsPagesPanelUpdate = false;
      _r1 = performance.now();
    }
    renderStage();
    const _r2 = performance.now();
    renderToolsPanel();
    const _r3 = performance.now();
    const _total = _r3 - _r0;
    if (_total > 16) {
      console.warn(`[render perf] total=${_total.toFixed(1)}ms  pages=${(_r1-_r0).toFixed(1)}ms  stage=${(_r2-_r1).toFixed(1)}ms  tools=${(_r3-_r2).toFixed(1)}ms  spreads=${Math.floor((state.pages.length-1)/2)}`);
    }
  } catch (err) {
    console.error('[render] ERROR:', err);
  }
}

function renderSimpleThumbnail(page) {
  if (!page) return '';

  // Sidebar is ~250px, spread has 2 pages, each page thumbnail is ~120px wide
  // Page is 400x600, so scale to fit 120px width = 0.3 scale
  // But thumb-content has bottom:12px for label, so actual height is less
  // Using 0.28 scale for safety
  const scale = 0.28;

  let html = '';

  // Render elements scaled down
  const sorted = [...(page.elements || [])].sort((a, b) => (a.z || 0) - (b.z || 0));
  sorted.forEach(el => {
    const x = el.x * scale;
    const y = el.y * scale;
    const w = el.w * scale;
    const h = el.h * scale;
    const rotation = el.rotation ? `transform:rotate(${el.rotation}deg);` : '';
    const style = `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;overflow:hidden;${rotation}`;

    if (el.t === 'image') {
      if (el.src) {
        const bg = el.deepEtched ? 'transparent' : '#e0e0e0';
        const hasFlip = el.flipH || el.flipV;
        let imgStyle = '';
        let thumbWrapStyle = '';
        if (el.innerW !== undefined) {
          // Locked composition: absolutely positioned, scaled to thumbnail
          const ix = el.innerX * scale;
          const iy = el.innerY * scale;
          const iw = el.innerW * scale;
          const ih = el.innerH * scale;
          imgStyle = `position:absolute;left:${ix}px;top:${iy}px;width:${iw}px;height:${ih}px;`;
          // Flip on wrapper (matching canvas .img-wrap behavior) so transform-origin is frame center
          if (hasFlip) thumbWrapStyle = `position:absolute;top:0;left:0;width:100%;height:100%;overflow:hidden;transform:scale(${el.flipH?-1:1},${el.flipV?-1:1});transform-origin:center;`;
        } else {
          // Legacy: browser-managed object-fit
          const fitMode = el.deepEtched ? 'contain' : (el.fitMode || 'cover');
          const ox = el.imgOffsetX || 0;
          const oy = el.imgOffsetY || 0;
          imgStyle = `width:100%;height:100%;object-fit:${fitMode};`;
          if (ox !== 0 || oy !== 0) imgStyle += `object-position:calc(50% + ${ox}px) calc(50% + ${oy}px);`;
          // Legacy images: flip directly on img (no offset, so transform-origin doesn't matter)
          if (hasFlip) {
            let imgTransform = '';
            if (el.flipH) imgTransform += 'scaleX(-1) ';
            if (el.flipV) imgTransform += 'scaleY(-1) ';
            imgStyle += `transform:${imgTransform.trim()};`;
          }
        }
        const ex = el.exposure ?? 100;
        const c = el.contrast ?? 100;
        const s = el.saturation ?? 100;
        if (!(ex === 100 && c === 100 && s === 100)) {
          imgStyle += `filter:brightness(${ex/100}) contrast(${c/100}) saturate(${s/100});`;
        }

        if (el.rippedFrame) {
          const frameDef = RIPPED_FRAMES.find(f => f.id === el.rippedFrame);
          if (frameDef && frameDef.type === 'torn') {
            // Torn paper: apply mask-image to clip the photo
            const frameSrc = getFrameSrc(frameDef);
            const maskStyle = `-webkit-mask-image:url(${frameSrc});mask-image:url(${frameSrc});-webkit-mask-size:100% 100%;mask-size:100% 100%;-webkit-mask-repeat:no-repeat;mask-repeat:no-repeat;`;
            html += `<div style="${style}background:transparent;overflow:visible"><div style="width:100%;height:100%;${maskStyle}"><img src="${el.src}" style="${imgStyle}"></div></div>`;
          } else if (frameDef && frameDef.type === 'photo-frame' && frameDef.window) {
            // Photo-frame: clip photo to window area, overlay frame PNG on top
            const fw = frameDef.window;
            html += `<div style="${style}background:transparent;overflow:hidden"><div style="position:relative;display:block;width:100%;height:100%;isolation:isolate;"><div style="position:absolute;top:${fw.top*100}%;left:${fw.left*100}%;width:${fw.width*100}%;height:${fw.height*100}%;-webkit-clip-path:inset(0);clip-path:inset(0);overflow:hidden;z-index:1;"><img src="${el.src}" style="width:100%;height:100%;object-fit:cover;display:block;"></div><img src="${getFrameSrc(frameDef)}" style="position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:2;"></div></div>`;
          } else if (frameDef && frameDef.type === 'polaroid') {
            // CSS polaroid: colored background with photo inset
            html += `<div style="${style}background:${frameDef.color};overflow:hidden;${frameDef.shadow?'box-shadow:0 1px 4px rgba(0,0,0,.15);':''}"><div style="position:absolute;top:4%;left:4%;width:92%;height:76%;overflow:hidden;"><img src="${el.src}" style="${imgStyle}"></div></div>`;
          } else {
            html += thumbWrapStyle
              ? `<div style="${style}background:${bg}"><div style="${thumbWrapStyle}"><img src="${el.src}" style="${imgStyle}"></div></div>`
              : `<div style="${style}background:${bg}"><img src="${el.src}" style="${imgStyle}"></div>`;
          }
        } else {
          html += thumbWrapStyle
            ? `<div style="${style}background:${bg}"><div style="${thumbWrapStyle}"><img src="${el.src}" style="${imgStyle}"></div></div>`
            : `<div style="${style}background:${bg}"><img src="${el.src}" style="${imgStyle}"></div>`;
        }
      } else {
        html += `<div style="${style}background:#ddd;border:1px dashed #bbb"></div>`;
      }
    } else if (el.t === 'text') {
      // Use a minimum 6px font so text blocks are visible in thumbnails
      const fontSize = Math.max(6, (el.fontSize || 12) * scale);
      // Strip HTML tags for cleaner thumbnail rendering
      const plainText = (el.txt || '').replace(/<[^>]*>/g, ' ').substring(0, 200);
      // Use authoritative element font directly — no quoting (matches canvas renderer line 427)
      const ff = el.fontFamily || 'Inter,sans-serif';
      // Text uses overflow:visible to match canvas behavior (text can flow beyond its box)
      const textStyle = `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;overflow:visible;padding:${6*scale}px;box-sizing:border-box;white-space:pre-wrap;word-break:break-word;${rotation}`;
      html += `<div style="${textStyle}font-family:${ff};font-size:${fontSize}px;color:${el.color||'#111'};text-align:${el.align||'left'};line-height:1.4;${el.bold?'font-weight:700;':''}${el.italic?'font-style:italic;':''}">${plainText}</div>`;
    } else if (el.t === 'shape') {
      let shapeStyle = `${style}background:${el.color||'#4A3F2A'};`;
      if (el.shape === 'circle') shapeStyle += 'border-radius:50%;';
      html += `<div style="${shapeStyle}"></div>`;
    } else if (el.t === 'drawing' && el.dataUrl) {
      html += `<div style="${style}"><img src="${el.dataUrl}" style="width:100%;height:100%"></div>`;
    } else if (el.t === 'graphic' && el.src) {
      html += `<div style="${style}"><img src="${el.src}" style="width:100%;height:100%;object-fit:contain"></div>`;
    }
  });

  return html;
}

// Keep for compatibility
function renderLiveThumbnail(page, scale) {
  return renderSimpleThumbnail(page);
}

function renderThumbnail(page, scale) {
  return renderSimpleThumbnail(page);
}

function renderLayoutPreview(layout) {
  if (!layout.preview) return '';
  let svg = '<svg viewBox="0 0 100 100" style="width:100%;height:100%"><rect width="100" height="100" fill="#f5f3ee"/>';
  layout.preview.forEach(p => {
    if (p.t === 'rect') {
      svg += `<rect x="${p.x}" y="${p.y}" width="${p.w}" height="${p.h}" fill="#d5d5d5" rx="1"/>`;
    } else if (p.t === 'text') {
      svg += `<text x="${p.x}" y="${p.y + (p.size||5)}" font-family="${p.font||'Inter'}" font-size="${p.size||5}" fill="#444">${p.txt}</text>`;
    } else if (p.t === 'circle') {
      svg += `<circle cx="${p.x}" cy="${p.y}" r="${p.r}" fill="#666"/>`;
    }
  });
  svg += '</svg>';
  return svg;
}

// ============ STAGE RENDERING ============
function renderStage() {
  const stage = document.getElementById('stage');
  const page = getPage();
  if (!page) { stage.innerHTML = ''; return; }
  
  let html = '';
  // Cover page shows alone, other pages show as spreads
  if (page.id === 'cover') {
    html = renderPageElement(page);
  } else {
    const idx = state.pages.findIndex(p => p.id === state.currentPage);
    let leftIdx = idx % 2 === 0 ? idx - 1 : idx;
    let rightIdx = leftIdx + 1;
    if (leftIdx < 1) leftIdx = 1;
    
    const leftPage = state.pages[leftIdx];
    const rightPage = state.pages[rightIdx];
    
    html = '<div class="spread">';
    // Page backgrounds only (no elements inside)
    if (leftPage) {
      const leftTexture = leftPage.texture ? `paper-${leftPage.texture}` : '';
      html += `<div class="page ${leftTexture}" data-page="${leftPage.id}" style="background:${leftPage.paper}">`;
      html += '</div>';
    }
    if (rightPage) {
      const rightTexture = rightPage.texture ? `paper-${rightPage.texture}` : '';
      html += `<div class="page ${rightTexture}" data-page="${rightPage.id}" style="background:${rightPage.paper}">`;
      html += '</div>';
    }
    
    // Shared element layer — all elements from both pages rendered here
    // with spread-relative coordinates (right page offset by 400px)
    html += '<div class="spread-elements" style="position:absolute;inset:0;pointer-events:none;">';
    
    // Left page elements
    if (leftPage) {
      const leftSorted = [...(leftPage.elements || [])].sort((a, b) => (a.z || 0) - (b.z || 0));
      var leftMaxZ = 1;
      leftSorted.forEach(el => {
        if ((el.z || 1) > leftMaxZ) leftMaxZ = el.z || 1;
        // Elements render at their natural coordinates (left page starts at 0)
        html += renderElement(el, leftPage.id);
      });
      leftSorted.forEach(el => {
        if (el.t === 'image' && el.revealMask && el.src) {
          html += renderRevealLayer(el, leftMaxZ);
        }
      });
    }

    // Right page elements (offset by 400px)
    if (rightPage) {
      const rightSorted = [...(rightPage.elements || [])].sort((a, b) => (a.z || 0) - (b.z || 0));
      var rightMaxZ = 1;
      rightSorted.forEach(el => {
        if ((el.z || 1) > rightMaxZ) rightMaxZ = el.z || 1;
        html += renderElement(el, rightPage.id, PAGE_W); // offset for right page
      });
      rightSorted.forEach(el => {
        if (el.t === 'image' && el.revealMask && el.src) {
          html += renderRevealLayer(el, rightMaxZ, PAGE_W); // offset for right page
        }
      });
    }
    
    html += '</div>'; // end spread-elements

    // Draw canvas — rendered at spread level (ABOVE .spread-elements) so it reliably receives mouse events.
    // Positioned over the active page using left offset.
    if (drawState.active) {
      if (leftPage && leftPage.id === state.currentPage) {
        html += `<canvas id="drawCanvas" width="${PAGE_W}" height="${PAGE_H}" style="position:absolute;left:0;top:0;width:${PAGE_W}px;height:${PAGE_H}px;cursor:crosshair;z-index:100;pointer-events:auto" aria-label="Drawing canvas"></canvas>`;
      } else if (rightPage && rightPage.id === state.currentPage) {
        html += `<canvas id="drawCanvas" width="${PAGE_W}" height="${PAGE_H}" style="position:absolute;left:${PAGE_W}px;top:0;width:${PAGE_W}px;height:${PAGE_H}px;cursor:crosshair;z-index:100;pointer-events:auto" aria-label="Drawing canvas"></canvas>`;
      }
    }

    html += '<div class="gutter-line"></div>';
    // Active page indicator
    if (leftPage && state.activePage === leftPage.id) {
      html += '<div class="active-page-indicator"></div>';
    } else if (rightPage && state.activePage === rightPage.id) {
      html += '<div class="active-page-indicator right"></div>';
    }
    html += '</div>';
  }
  
  stage.innerHTML = html;
  bindStageEvents();
}

function renderPageElement(page) {
  const isActive = page.id === state.activePage;
  const textureClass = page.texture ? `paper-${page.texture}` : '';
  let html = `<div class="page ${isActive?'active-page':''} ${textureClass}" data-page="${page.id}" style="background:${page.paper}">`;
  
  if (drawState.active && page.id === state.currentPage) {
    html += `<canvas id="drawCanvas" width="${PAGE_W}" height="${PAGE_H}" style="position:absolute;inset:0;cursor:crosshair;z-index:50" aria-label="Drawing canvas"></canvas>`;
  }
  
  // Find the max z-index used by elements
  const maxZ = Math.max(...(page.elements || []).map(el => el.z || 1), 1);
  
  const sorted = [...(page.elements || [])].sort((a, b) => (a.z || 0) - (b.z || 0));
  sorted.forEach(el => {
    html += renderElement(el, page.id);
  });
  
  // Add reveal layers at page level (ABOVE all other elements)
  // Each reveal layer gets z-index of maxZ + 100 + its own z-index (to maintain relative order)
  sorted.forEach(el => {
    if (el.t === 'image' && el.revealMask && el.src) {
      html += renderRevealLayer(el, maxZ);
    }
  });
  
  html += '</div>';
  return html;
}

function renderRevealLayer(el, maxZ, xOffset) {
  xOffset = xOffset || 0;
  // Build the same image style as the original
  const flipH = el.flipH ? -1 : 1;
  const flipV = el.flipV ? -1 : 1;
  const opacity = el.opacity !== undefined ? el.opacity : 1;
  let imgStyle = '';
  if (el.innerW !== undefined) {
    // Locked composition
    imgStyle = `position:absolute;left:${el.innerX}px;top:${el.innerY}px;width:${el.innerW}px;height:${el.innerH}px;`;
  } else {
    // Legacy
    const fitMode = el.fitMode || 'cover';
    const ox = el.imgOffsetX || 0;
    const oy = el.imgOffsetY || 0;
    imgStyle = `object-fit:${fitMode};width:100%;height:100%;`;
    if (ox !== 0 || oy !== 0) imgStyle += `object-position:calc(50% + ${ox}px) calc(50% + ${oy}px);`;
  }
  if (opacity !== 1) imgStyle += `opacity:${opacity};`;
  const ex = el.exposure ?? 100;
  const ct = el.contrast ?? 100;
  const sa = el.saturation ?? 100;
  if (!(ex === 100 && ct === 100 && sa === 100)) {
    imgStyle += `filter:brightness(${ex/100}) contrast(${ct/100}) saturate(${sa/100});`;
  }

  // Reveal layer z-index: maxZ + 100 + element's own z to maintain stacking between multiple reveals
  const revealZ = maxZ + 100 + (el.z || 1);

  // Position the reveal layer at the same spot as the element
  // Flip is applied via an inner wrapper with center origin
  let wrapStyle = `position:absolute;left:${el.x + xOffset}px;top:${el.y}px;width:${el.w}px;height:${el.h}px;z-index:${revealZ};pointer-events:none;overflow:hidden;`;
  if (el.rotation) wrapStyle += `transform:rotate(${el.rotation}deg);`;

  let flipWrapStyle = el.innerW !== undefined ? 'position:absolute;top:0;left:0;width:100%;height:100%;overflow:hidden;' : '';
  if (flipH === -1 || flipV === -1) {
    flipWrapStyle += `transform:scale(${flipH},${flipV});transform-origin:center;width:100%;height:100%;`;
  }

  return `<div class="reveal-layer-page" style="${wrapStyle}">
    <div style="${flipWrapStyle}"><img src="${el.src}" style="${imgStyle}-webkit-mask-image:url(${el.revealMask});mask-image:url(${el.revealMask});-webkit-mask-size:100% 100%;mask-size:100% 100%;"></div>
  </div>`;
}

function renderElement(el, pageId, xOffset) {
  xOffset = xOffset || 0;
  const isSelected = el.id === state.selected;
  const isMultiSelected = state.multiSelected.includes(el.id);
  const ptrEvents = drawState.active ? 'none' : 'auto';
  let style = `left:${el.x + xOffset}px;top:${el.y}px;width:${el.w}px;height:${el.h}px;z-index:${el.z||1};pointer-events:${ptrEvents};`;
  if (el.rotation) style += `transform:rotate(${el.rotation}deg);`;
  
  let classes = `element ${el.t}`;
  if (isSelected) classes += ' selected';
  if (isMultiSelected) classes += ' multi-selected';
  if (state.imagePositionMode === el.id) classes += ' crop-mode';
  if (el.t === 'image' && !el.src) classes += ' empty';
  if (el.t === 'image' && el.deepEtched) classes += ' deep-etched';
  if (el.t === 'image' && el.sticker) classes += ' sticker';
  if (el.t === 'image' && el.rippedFrame) {
    classes += ' ripped-frame';
    const _fd = RIPPED_FRAMES.find(f => f.id === el.rippedFrame);
    if (_fd && _fd.type === 'polaroid') {
      classes += ' polaroid-frame';
      style += `background:${_fd.color};`;
    } else if (_fd && _fd.type === 'photo-frame') {
      style += `overflow:hidden;background:transparent;`;
    }
  }
  if (el.t === 'image' && el.border) classes += ` border-${el.border}`;
  if (el.t === 'shape') classes += ` ${el.shape}`;
  
  let inner = '';
  if (el.t === 'image') {
    if (el.src) {
      // Build image style
      const flipH = el.flipH ? -1 : 1;
      const flipV = el.flipV ? -1 : 1;
      const opacity = el.opacity !== undefined ? el.opacity : 1;
      const ex = el.exposure ?? 100;
      const ct = el.contrast ?? 100;
      const sa = el.saturation ?? 100;

      let imgStyle = '';
      let wrapStyle = '';

      if (el.innerW !== undefined) {
        // Locked composition: absolutely positioned image inside frame
        imgStyle = `position:absolute;left:${el.innerX}px;top:${el.innerY}px;width:${el.innerW}px;height:${el.innerH}px;`;
        wrapStyle = 'position:absolute;top:0;left:0;width:100%;height:100%;overflow:hidden;';
      } else {
        // Legacy: browser-managed object-fit
        const fitMode = el.fitMode || 'cover';
        const ox = el.imgOffsetX || 0;
        const oy = el.imgOffsetY || 0;
        imgStyle = `object-fit:${fitMode};width:100%;height:100%;`;
        if (ox !== 0 || oy !== 0) imgStyle += `object-position:calc(50% + ${ox}px) calc(50% + ${oy}px);`;
      }

      if (opacity !== 1) imgStyle += `opacity:${opacity};`;
      if (!(ex === 100 && ct === 100 && sa === 100)) {
        imgStyle += `filter:brightness(${ex/100}) contrast(${ct/100}) saturate(${sa/100});`;
      }

      // Flip is applied on the img-wrap with center origin so image mirrors in place
      if (flipH === -1 || flipV === -1) {
        wrapStyle += `transform:scale(${flipH},${flipV});transform-origin:center;`;
      }

      inner = `<div class="img-wrap" style="${wrapStyle}"><img src="${el.src}" style="${imgStyle}"></div>`;
      if (el.rippedFrame) {
        const frameDef = RIPPED_FRAMES.find(f => f.id === el.rippedFrame);
        if (frameDef && frameDef.type === 'torn') {
          const maskStyle = `-webkit-mask-image:url(${getFrameSrc(frameDef)});mask-image:url(${getFrameSrc(frameDef)});-webkit-mask-size:100% 100%;mask-size:100% 100%;-webkit-mask-repeat:no-repeat;mask-repeat:no-repeat;`;
          inner = `<div class="img-wrap" style="${wrapStyle}${maskStyle}"><img src="${el.src}" style="${imgStyle}"></div>`;
        } else if (frameDef && frameDef.type === 'photo-frame' && frameDef.window) {
          // Photo-frame: PNG overlay on top, user photo clipped to measured window area
          // isolation:isolate prevents bleed into parent stacking context
          // clip-path + -webkit-clip-path for robust cross-browser clipping
          const w = frameDef.window;
          inner = `<div class="img-wrap" style="${wrapStyle}position:relative;display:block;width:100%;height:100%;overflow:hidden;isolation:isolate;pointer-events:auto;"><div style="position:absolute;top:${w.top*100}%;left:${w.left*100}%;width:${w.width*100}%;height:${w.height*100}%;-webkit-clip-path:inset(0);clip-path:inset(0);overflow:hidden;z-index:1;"><img src="${el.src}" style="width:100%;height:100%;object-fit:cover;display:block;${opacity !== 1 ? 'opacity:'+opacity+';' : ''}${!(ex === 100 && ct === 100 && sa === 100) ? 'filter:brightness('+(ex/100)+') contrast('+(ct/100)+') saturate('+(sa/100)+');' : ''}"></div><img src="${getFrameSrc(frameDef)}" style="position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:2;" loading="lazy"></div>`;
        } else if (frameDef && frameDef.type === 'polaroid') {
          inner = `<div class="img-wrap" style="${wrapStyle}"><img src="${el.src}" style="${imgStyle}"></div>`;
        }
      }
    } else {
      // Empty placeholder
      if (el.rippedFrame) {
        const frameDef = RIPPED_FRAMES.find(f => f.id === el.rippedFrame);
        if (frameDef && frameDef.type === 'torn') {
          const maskStyle = `-webkit-mask-image:url(${getFrameSrc(frameDef)});mask-image:url(${getFrameSrc(frameDef)});-webkit-mask-size:100% 100%;mask-size:100% 100%;-webkit-mask-repeat:no-repeat;mask-repeat:no-repeat;`;
          inner = `<div class="frame-placeholder" style="background:#e0e0e0;${maskStyle}"><span style="font-size:9px;color:#999">+ Image</span></div>`;
        } else if (frameDef && frameDef.type === 'photo-frame' && frameDef.window) {
          const w = frameDef.window;
          inner = `<div style="position:relative;width:100%;height:100%;"><div style="position:absolute;top:${w.top*100}%;left:${w.left*100}%;width:${w.width*100}%;height:${w.height*100}%;background:#e0e0e0;z-index:1;display:flex;align-items:center;justify-content:center;"><span style="font-size:9px;color:#999">+ Image</span></div><img src="${getFrameSrc(frameDef)}" style="position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:2;" loading="lazy"></div>`;
        } else if (frameDef && frameDef.type === 'polaroid') {
          inner = `<div class="frame-placeholder" style="background:#e0e0e0"><span style="font-size:9px;color:#999">+ Image</span></div>`;
        } else {
          inner = '<span style="font-size:9px;color:#999">+ Image</span>';
        }
      } else {
        inner = '<span style="font-size:9px;color:#999">+ Image</span>';
      }
    }
  } else if (el.t === 'text') {
    let textStyle = `font-family:${el.fontFamily||'Inter,sans-serif'};font-size:${el.fontSize||12}px;color:${el.color||'#111'};text-align:${el.align||'left'};`;
    if (el.bold) textStyle += 'font-weight:700;';
    if (el.italic) textStyle += 'font-style:italic;';
    inner = `<div class="text-content" style="${textStyle}">${el.txt || ''}</div>`;
  } else if (el.t === 'shape') {
    const tex = el.texture || 'medium';
    inner = `<div class="shape-inner tex-${tex}" style="background:${el.color||'#4A3F2A'}"></div>`;
  } else if (el.t === 'graphic') {
    inner = `<div class="graphic-inner ${el.cls}"></div>`;
  } else if (el.t === 'drawing' && el.dataUrl) {
    inner = `<img src="${el.dataUrl}">`;
  }
  
  let handles = '';
  if (state.imagePositionMode === el.id) {
    // Crop mode: 4 corner L-bracket handles (resize frame, not image)
    handles = '<div class="crop-handle crop-handle-corner crop-handle-nw"></div>' +
              '<div class="crop-handle crop-handle-corner crop-handle-ne"></div>' +
              '<div class="crop-handle crop-handle-corner crop-handle-se"></div>' +
              '<div class="crop-handle crop-handle-corner crop-handle-sw"></div>';
  } else if (isSelected) {
    handles = '<div class="handle handle-nw"></div><div class="handle handle-ne"></div><div class="handle handle-se"></div><div class="handle handle-sw"></div><div class="handle handle-rotate"></div>';
  }
  
  
  return `<div class="${classes}" data-id="${el.id}" data-page="${pageId}" data-xoffset="${xOffset}" style="${style}">${inner}${handles}</div>`;
}

// ============ STAGE EVENT BINDING ============
function bindStageEvents() {
  // Page clicks
  document.querySelectorAll('.page').forEach(pageEl => {
    pageEl.addEventListener('mousedown', e => {
      // Only start marquee if clicking directly on page background
      if (e.target === pageEl || e.target.classList.contains('page')) {
        if (state.imagePositionMode) { exitImagePositionMode(); return; }
        if (revealState.active || revealState.quickMode) return;
        if (drawState.active) return;
        
        e.preventDefault();
        
        // Start marquee selection
        const rect = pageEl.getBoundingClientRect();
        marqueeState.active = true;
        marqueeState.startX = e.clientX - rect.left;
        marqueeState.startY = e.clientY - rect.top;
        marqueeState.currentX = marqueeState.startX;
        marqueeState.currentY = marqueeState.startY;
        marqueeState.pageId = pageEl.dataset.page;
        
        // Clear current selection
        state.selected = null;
        state.multiSelected = [];
        state.activePage = pageEl.dataset.page;
        // Immediately remove handles/selected class from DOM
        document.querySelectorAll('.element.selected').forEach(el => {
          el.classList.remove('selected');
          el.querySelectorAll('.handle').forEach(h => h.remove());
        });
        
        // Create marquee element
        const marquee = document.createElement('div');
        marquee.className = 'marquee-box';
        marquee.id = 'marqueeBox';
        pageEl.appendChild(marquee);
        
        updateMarqueeBox();
        
        document.addEventListener('mousemove', onMarqueeMove);
        document.addEventListener('mouseup', onMarqueeUp);
      }
    });
    pageEl.addEventListener('click', e => {
      if (e.target === pageEl || e.target.classList.contains('page')) {
        // Exit any active text editing
        const activeEditable = document.querySelector('.text-content[contenteditable="true"]');
        if (activeEditable) activeEditable.blur();
        // Finish drawing if draw mode is active
        if (drawState.active) { finishDrawing(); return; }
        if (marqueeState.active) return; // Don't deselect if just finishing marquee
        if (state.imagePositionMode) { exitImagePositionMode(); return; }
        state.activePage = pageEl.dataset.page;
        state.selected = null;
        state.multiSelected = [];
        render();
      }
    });
    pageEl.addEventListener('dragover', e => {
      e.preventDefault();
      // Check if hovering over an empty image placeholder
      const target = e.target.closest('.element.image.empty');
      // Clear previous highlights
      document.querySelectorAll('.drop-target').forEach(el => el.classList.remove('drop-target'));
      // Highlight if over empty placeholder
      if (target) {
        target.classList.add('drop-target');
      }
    });
    pageEl.addEventListener('dragleave', e => {
      // Clear highlights when leaving page
      document.querySelectorAll('.drop-target').forEach(el => el.classList.remove('drop-target'));
    });
    pageEl.addEventListener('drop', e => { e._handled = true; handlePageDrop(e, pageEl.dataset.page); });
  });
  
  // B1 FIX: In spread mode, elements are in .spread-elements overlay above .page divs.
  // Without these handlers, drops on image placeholders in spreads don't register.
  document.querySelectorAll('.spread-elements').forEach(spreadEl => {
    spreadEl.addEventListener('dragover', e => {
      e.preventDefault();
      const target = e.target.closest('.element.image.empty');
      document.querySelectorAll('.drop-target').forEach(el => el.classList.remove('drop-target'));
      if (target) {
        target.classList.add('drop-target');
      }
    });
    spreadEl.addEventListener('dragleave', e => {
      document.querySelectorAll('.drop-target').forEach(el => el.classList.remove('drop-target'));
    });
    spreadEl.addEventListener('drop', e => {
      e._handled = true;
      // Determine which page the drop is on based on x position
      const spreadRect = spreadEl.closest('.spread').getBoundingClientRect();
      const dropX = e.clientX - spreadRect.left;
      // Check if dropped on an empty image placeholder first
      const dropTarget = e.target.closest('.element.image.empty');
      let pageId;
      if (dropTarget) {
        pageId = dropTarget.dataset.page;
      } else {
        // Use x position to determine left or right page
        const spread = spreadEl.closest('.spread');
        const pages = spread.querySelectorAll('.page');
        if (dropX > PAGE_W && pages[1]) {
          pageId = pages[1].dataset.page;
        } else if (pages[0]) {
          pageId = pages[0].dataset.page;
        }
      }
      if (pageId) handlePageDrop(e, pageId);
    });
  });

  // Click on grey stage background (outside the canvas/spread) → deselect everything
  const stageEl = document.getElementById('stage');
  if (stageEl) {
    stageEl.addEventListener('mousedown', e => {
      // Only if clicking the stage itself, not a child (spread, page, element, etc.)
      if (e.target === stageEl) {
        const activeEditable = document.querySelector('.text-content[contenteditable="true"]');
        if (activeEditable) activeEditable.blur();
        if (state.imagePositionMode) { exitImagePositionMode(); }
        state.selected = null;
        state.multiSelected = [];
        render();
      }
    });
  }

  // Click on empty space within .spread-elements overlay → deselect
  // (The overlay sits above .page divs, so page click handlers don't fire)
  document.querySelectorAll('.spread-elements').forEach(spreadEl => {
    spreadEl.addEventListener('mousedown', e => {
      // Only if clicking the overlay itself, not an element inside it
      if (e.target === spreadEl) {
        const activeEditable = document.querySelector('.text-content[contenteditable="true"]');
        if (activeEditable) activeEditable.blur();
        if (state.imagePositionMode) { exitImagePositionMode(); return; }
        if (drawState.active) return;
        state.selected = null;
        state.multiSelected = [];
        // Determine left vs right page based on click position within the spread
        const spread = spreadEl.closest('.spread');
        const spreadPages = spread ? spread.querySelectorAll('.page') : [];
        if (spreadPages.length === 2) {
          const rect = spread.getBoundingClientRect();
          const clickX = e.clientX - rect.left;
          state.activePage = clickX >= rect.width / 2 ? spreadPages[1].dataset.page : spreadPages[0].dataset.page;
        } else if (spreadPages.length === 1) {
          state.activePage = spreadPages[0].dataset.page;
        }
        render();
      }
    });
  });

  // Elements
  document.querySelectorAll('.element').forEach(el => bindElementEvents(el));
  
  // Draw canvas — remove before re-adding to prevent listener stacking across render cycles
  if (drawState.active) {
    const canvas = document.getElementById('drawCanvas');
    console.log('[bindStageEvents] drawState.active, canvas found:', !!canvas);
    if (canvas) {
      drawState.canvas = canvas;
      drawState.ctx = canvas.getContext('2d');
      canvas.removeEventListener('mousedown', startDrawStroke);
      canvas.removeEventListener('mousemove', continueDrawStroke);
      canvas.removeEventListener('mouseup', endDrawStroke);
      canvas.removeEventListener('mouseleave', endDrawStroke);
      canvas.addEventListener('mousedown', startDrawStroke);
      canvas.addEventListener('mousemove', continueDrawStroke);
      canvas.addEventListener('mouseup', endDrawStroke);
      canvas.addEventListener('mouseleave', endDrawStroke);
      console.log('[bindStageEvents] draw events bound, canvas size:', canvas.width, 'x', canvas.height);
    } else {
      console.error('[bindStageEvents] WARNING: drawState.active but no drawCanvas in DOM!');
    }
  }
}

// Marquee selection functions
function onMarqueeMove(e) {
  if (!marqueeState.active) return;
  
  const pageEl = document.querySelector(`.page[data-page="${marqueeState.pageId}"]`);
  if (!pageEl) return;
  
  const rect = pageEl.getBoundingClientRect();
  marqueeState.currentX = Math.max(0, Math.min(PAGE_W, e.clientX - rect.left));
  marqueeState.currentY = Math.max(0, Math.min(PAGE_H, e.clientY - rect.top));
  
  updateMarqueeBox();
  updateMarqueeSelection();
}

function onMarqueeUp(e) {
  document.removeEventListener('mousemove', onMarqueeMove);
  document.removeEventListener('mouseup', onMarqueeUp);

  // Remove marquee box
  const marquee = document.getElementById('marqueeBox');
  if (marquee) marquee.remove();

  // Ignore tiny marquees (likely just a click, not an intentional selection drag)
  const marqueeW = Math.abs(marqueeState.currentX - marqueeState.startX);
  const marqueeH = Math.abs(marqueeState.currentY - marqueeState.startY);
  if (marqueeW < 5 && marqueeH < 5) {
    // This was just a click on empty space — deselect everything
    state.selected = null;
    state.multiSelected = [];
    marqueeState.active = false;
    render();
    return;
  }

  // Finalize selection
  if (state.multiSelected.length === 1) {
    // If only one element selected, switch to single selection mode
    state.selected = state.multiSelected[0];
    state.multiSelected = [];
  } else if (state.multiSelected.length === 0) {
    // No elements selected - this was just a click
    state.selected = null;
  }
  
  marqueeState.active = false;
  render();
}

function updateMarqueeBox() {
  const marquee = document.getElementById('marqueeBox');
  if (!marquee) return;
  
  const x = Math.min(marqueeState.startX, marqueeState.currentX);
  const y = Math.min(marqueeState.startY, marqueeState.currentY);
  const w = Math.abs(marqueeState.currentX - marqueeState.startX);
  const h = Math.abs(marqueeState.currentY - marqueeState.startY);
  
  marquee.style.left = x + 'px';
  marquee.style.top = y + 'px';
  marquee.style.width = w + 'px';
  marquee.style.height = h + 'px';
}

function updateMarqueeSelection() {
  const page = state.pages.find(p => p.id === marqueeState.pageId);
  if (!page) return;
  
  const mx1 = Math.min(marqueeState.startX, marqueeState.currentX);
  const my1 = Math.min(marqueeState.startY, marqueeState.currentY);
  const mx2 = Math.max(marqueeState.startX, marqueeState.currentX);
  const my2 = Math.max(marqueeState.startY, marqueeState.currentY);
  
  // Find elements that intersect with marquee
  const selected = [];
  (page.elements || []).forEach(el => {
    const ex1 = el.x;
    const ey1 = el.y;
    const ex2 = el.x + el.w;
    const ey2 = el.y + el.h;
    
    // Check intersection
    if (mx1 < ex2 && mx2 > ex1 && my1 < ey2 && my2 > ey1) {
      selected.push(el.id);
    }
  });
  
  state.multiSelected = selected;
  
  // Update visual feedback
  document.querySelectorAll('.element').forEach(el => {
    el.classList.remove('multi-selected');
    if (selected.includes(el.dataset.id)) {
      el.classList.add('multi-selected');
    }
  });
}

function bindElementEvents(div) {
  const elId = div.dataset.id;
  const pageId = div.dataset.page;
  
  div.addEventListener('mousedown', e => {
    if (e.target.classList.contains('handle')) return;
    e.stopPropagation();
    
    const el = findElementById(elId);
    
    // If in image-position mode, handle pan (Figma-style reposition)
    if (state.imagePositionMode === elId && el?.t === 'image') {
      // Don't pan if clicking on a crop handle
      if (e.target.classList.contains('crop-handle') || e.target.closest('.crop-handle')) return;
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX;
      const startY = e.clientY;

      if (el.innerW !== undefined) {
        // Locked model: pan via innerX/innerY
        const startIX = el.innerX;
        const startIY = el.innerY;
        let historySaved = false;

        function onMoveLocked(ev) {
          ev.preventDefault();
          if (!historySaved) { saveHistory(); historySaved = true; }
          el.innerX = startIX + (ev.clientX - startX);
          el.innerY = startIY + (ev.clientY - startY);
          const img = div.querySelector('.img-wrap img');
          if (img) {
            img.style.left = el.innerX + 'px';
            img.style.top = el.innerY + 'px';
          }
        }
        function onUpLocked() {
          document.removeEventListener('mousemove', onMoveLocked);
          document.removeEventListener('mouseup', onUpLocked);
          scheduleAutosave();
          renderPagesPanel();
        }
        document.addEventListener('mousemove', onMoveLocked);
        document.addEventListener('mouseup', onUpLocked);
      } else {
        // Legacy model: pan via imgOffsetX/imgOffsetY
        const startOX = el.imgOffsetX || 0;
        const startOY = el.imgOffsetY || 0;

        function onMoveLegacy(ev) {
          ev.preventDefault();
          el.imgOffsetX = startOX + (ev.clientX - startX);
          el.imgOffsetY = startOY + (ev.clientY - startY);
          const img = div.querySelector('img');
          if (img) {
            img.style.objectPosition = `calc(50% + ${el.imgOffsetX}px) calc(50% + ${el.imgOffsetY}px)`;
          }
        }
        function onUpLegacy() {
          document.removeEventListener('mousemove', onMoveLegacy);
          document.removeEventListener('mouseup', onUpLegacy);
          saveHistory();
          renderPagesPanel();
        }
        document.addEventListener('mousemove', onMoveLegacy);
        document.addEventListener('mouseup', onUpLegacy);
      }
      return;
    }
    
    // Check if clicking on a multi-selected element
    if (state.multiSelected.includes(elId)) {
      // Start group drag
      startGroupDrag(e, pageId);
      return;
    }
    
    // Normal selection - clear multi-selection
    state.multiSelected = [];
    state.selected = elId;
    state.activePage = pageId;
    state.currentPage = pageId;
    
    // Always render tools immediately on selection so options appear without needing to drag
    renderToolsPanel();

    // Immediate DOM update: show handles on selected element without full render()
    document.querySelectorAll('.element.selected').forEach(el => {
      el.classList.remove('selected');
      el.querySelectorAll('.handle').forEach(h => h.remove());
    });
    div.classList.add('selected');
    if (!div.querySelector('.handle')) {
      div.insertAdjacentHTML('beforeend',
        '<div class="handle handle-nw"></div><div class="handle handle-ne"></div>' +
        '<div class="handle handle-se"></div><div class="handle handle-sw"></div>' +
        '<div class="handle handle-rotate"></div>');
      // Bind resize/rotate events to new handles
      div.querySelectorAll('.handle').forEach(handle => {
        handle.addEventListener('mousedown', he => {
          he.stopPropagation();
          const type = handle.className.includes('rotate') ? 'rotate' :
                       handle.className.includes('nw') ? 'nw' :
                       handle.className.includes('ne') ? 'ne' :
                       handle.className.includes('sw') ? 'sw' : 'se';
          if (type === 'rotate') startRotate(he, elId, pageId);
          else startResize(he, elId, pageId, type);
        });
      });
    }

    if (!e.target.classList.contains('text-content')) {
      startDrag(e, elId, pageId);
    } else {
      // If text is already in edit mode, don't drag - let browser handle text selection
      if (e.target.contentEditable === 'true') {
        return;
      }
      // If element is already selected, enter edit mode on single click (no double-click needed)
      if (state.selected === elId) {
        e.target.contentEditable = 'true';
        e.target.focus();
        return;
      }
      // First click on unselected text - select element, start drag
      startDrag(e, elId, pageId);
    }
  });
  
  div.addEventListener('dblclick', e => {
    e.stopPropagation();
    e.preventDefault();
    const el = findElementById(elId);
    
    // Double-click on image - toggle image-position mode
    if (el?.t === 'image' && el.src) {
      if (state.imagePositionMode === elId) exitImagePositionMode();
      else enterImagePositionMode(elId);
      return;
    }
    
    if (el?.t === 'text' && e.target.classList.contains('text-content')) {
      const textEl = e.target;
      textEl.contentEditable = 'true';
      textEl.focus();
    } else if (el?.t === 'text') {
      state.selected = elId;
      render();
      setTimeout(() => {
        const textEl = document.querySelector(`[data-id="${elId}"] .text-content`);
        if (textEl) {
          textEl.contentEditable = 'true';
          textEl.focus();
          const range = document.createRange();
          range.selectNodeContents(textEl);
          const sel = window.getSelection();
          sel.removeAllRanges();
          sel.addRange(range);
        }
      }, 20);
    }
  });
  
  div.addEventListener('contextmenu', e => {
    e.preventDefault();
    e.stopPropagation();
    
    // If right-clicking on a multi-selected element, keep the multi-selection
    if (state.multiSelected.includes(elId)) {
      // Keep multi-selection, show context menu
    } else {
      // Single selection
      state.selected = elId;
      state.multiSelected = [];
    }
    render();
    showContextMenu(e.clientX, e.clientY);
  });
  
  // Handles
  div.querySelectorAll('.handle').forEach(handle => {
    handle.addEventListener('mousedown', e => {
      e.stopPropagation();
      const type = handle.className.includes('rotate') ? 'rotate' :
                   handle.className.includes('nw') ? 'nw' :
                   handle.className.includes('ne') ? 'ne' :
                   handle.className.includes('sw') ? 'sw' : 'se';
      if (type === 'rotate') startRotate(e, elId, pageId);
      else startResize(e, elId, pageId, type);
    });
  });

  // Crop handles — resize the frame (crop window), image stays fixed
  div.querySelectorAll('.crop-handle').forEach(handle => {
    handle.addEventListener('mousedown', e => {
      e.stopPropagation();
      e.preventDefault();
      const el = findElementById(elId);
      if (!el || el.innerW === undefined) return;
      const corner = handle.className.includes('nw') ? 'nw' :
                     handle.className.includes('ne') ? 'ne' :
                     handle.className.includes('sw') ? 'sw' : 'se';
      startCropDrag(e, el, elId, pageId, corner);
    });
  });

  // Text content — single source of truth: blur saves to history, input syncs live only
  const textContent = div.querySelector('.text-content');
  if (textContent) {
    // Track the text before editing starts, so we can compare on blur
    let textBeforeEdit = null;

    textContent.addEventListener('focus', () => {
      const el = findElementById(elId);
      if (el) textBeforeEdit = el.txt;
    });

    // Live sync on input (no history save — that happens on blur)
    textContent.addEventListener('input', () => {
      const el = findElementById(elId);
      if (el) {
        el.txt = textContent.innerHTML;
      }
    });

    // Save to history only on blur, and only if text actually changed
    textContent.addEventListener('blur', () => {
      const el = findElementById(elId);
      if (el && el.t === 'text') {
        const newText = textContent.innerHTML;
        el.txt = newText;
        // Only create a history entry if text actually changed during this edit session
        if (textBeforeEdit !== null && textBeforeEdit !== newText) {
          saveHistory();
          renderPagesPanel();
        }
        textContent.contentEditable = 'false';
        textBeforeEdit = null;
      }
    });

    // Exit text editing when pressing Escape
    textContent.addEventListener('keydown', e => {
      if (e.key === 'Escape') {
        textContent.blur();
      }
    });

    // Paste as plain text only - strip all HTML formatting
    // Uses Selection/Range API instead of deprecated execCommand
    textContent.addEventListener('paste', e => {
      e.preventDefault();
      const text = (e.clipboardData || window.clipboardData).getData('text/plain');
      const sel = window.getSelection();
      if (!sel || sel.rangeCount === 0) return;
      const range = sel.getRangeAt(0);
      range.deleteContents();
      const textNode = document.createTextNode(text);
      range.insertNode(textNode);
      // Move cursor to end of inserted text
      range.setStartAfter(textNode);
      range.setEndAfter(textNode);
      sel.removeAllRanges();
      sel.addRange(range);
      // Sync element state (live sync, history saved on blur)
      const el = findElementById(elId);
      if (el) el.txt = textContent.innerHTML;
    });
  }
  
  // Empty image click
  if (div.classList.contains('empty')) {
    div.addEventListener('click', e => {
      e.stopPropagation();
      uploadToElement(elId, pageId);
    });
  }
}
