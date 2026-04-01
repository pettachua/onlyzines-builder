// ============ FONT LOADER ============
// On-demand Google Fonts loading with dedup.
// Shared utility — can be included in builder and reader.

const _loadedFonts = new Set();

// Weight/italic specs for fonts that need them.
// Fonts NOT listed here load with default weight only.
const FONT_SPECS = {
  'Playfair Display': ':wght@400;700',
  'Inter': ':wght@400;500;600',
  'Work Sans': ':wght@400;500',
  'Londrina Solid': ':wght@400;900',
  'Kalam': ':wght@400;700',
  'DM Mono': ':wght@300;400;500',
  'Libre Baskerville': ':ital,wght@0,400;0,700;1,400',
  'DM Serif Display': ':ital@0;1',
  'Bodoni Moda': ':ital,wght@0,400;0,700;0,900;1,400',
  'Cormorant Garamond': ':ital,wght@0,300;0,400;0,700;1,300;1,400',
  'Oswald': ':wght@300;400;500;700',
  'Orbitron': ':wght@400;700;900',
  'Space Mono': ':wght@400;700',
  'Chivo': ':wght@400;700;900',
  'Passion One': ':wght@400;700;900',
  'Dancing Script': ':wght@400;700',
  'Spectral SC': ':wght@400;700',
  'Scheherazade New': ':wght@400;700',
  'Nanum Myeongjo': ':wght@400;700',
  'Bodoni Moda SC': ':ital,wght@0,400;0,700;0,900;1,400'
};

/**
 * Ensure a Google Font family is loaded. No-op if already loaded.
 * @param {string} familyWithFallback — e.g. "Playfair Display,serif" or "Playfair Display"
 */
let _fontRerenderScheduled = false;

function ensureFontLoaded(familyWithFallback) {
  const family = familyWithFallback.split(',')[0].trim();
  if (!family || family === 'sans-serif' || family === 'serif' || family === 'monospace' || family === 'cursive') return;
  if (_loadedFonts.has(family)) return;
  _loadedFonts.add(family);

  const spec = FONT_SPECS[family] || '';
  const encoded = encodeURIComponent(family).replace(/%20/g, '+');
  const link = document.createElement('link');
  link.rel = 'stylesheet';
  link.href = `https://fonts.googleapis.com/css2?family=${encoded}${spec}&display=swap`;
  link.onload = function() {
    if (_fontRerenderScheduled) return;
    _fontRerenderScheduled = true;
    setTimeout(function() {
      _fontRerenderScheduled = false;
      if (typeof _thumbCache !== 'undefined') _thumbCache.pageIds = null;
      if (typeof renderPagesPanel === 'function') renderPagesPanel();
    }, 0);
  };
  document.head.appendChild(link);
}

/**
 * Load fonts for a set of pages (scans all text elements).
 * @param {Array} pages — array of page objects with .elements[]
 */
function loadFontsForPages(pages) {
  const seen = new Set();
  for (const page of pages) {
    for (const el of (page.elements || [])) {
      if ((el.t === 'text' || el.type === 'text') && el.fontFamily) {
        const family = el.fontFamily.split(',')[0].trim();
        if (!seen.has(family)) {
          seen.add(family);
          ensureFontLoaded(el.fontFamily);
        }
      }
    }
  }
}
