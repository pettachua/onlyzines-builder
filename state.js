// ============ STATE ============

let state = {
  pages: [],
  currentPage: 'cover',
  activePage: 'cover',
  selected: null,
  multiSelected: [], // Array of element IDs for marquee selection
  clipboard: null,
  spreadClipboard: null,
  selectedSpread: null,
  lastCopied: null,
  imagePositionMode: null,
  layoutStyle: null,
  history: [],
  historyIndex: -1,
  sectionToggles: {} // tracks manual open/close of toolbar sections
};

let marqueeState = {
  active: false,
  startX: 0,
  startY: 0,
  currentX: 0,
  currentY: 0,
  pageId: null
};

let drawState = {
  active: false,
  tool: 'ink',
  color: '#333333',
  size: 3,
  canvas: null,
  ctx: null,
  drawing: false
};

let dragState = {
  active: false,
  type: null,
  element: null,
  startX: 0,
  startY: 0,
  origX: 0,
  origY: 0,
  origW: 0,
  origH: 0
};

let cutoutState = {
  element: null,
  ctx: null,
  img: null,
  mask: null,
  maskCtx: null,
  points: [],
  drawing: false,
  rectStart: null,
  maskHistory: [],
  maskHistoryIndex: -1
};

let revealState = {
  active: false,
  quickMode: false, // true when holding R key
  element: null,
  canvas: null,
  ctx: null,
  drawing: false,
  maskHistory: [],
  maskHistoryIndex: -1
};

let spreadDragState = {
  isDragging: false,
  startY: 0,
  spreadIdx: null,
  element: null,
  placeholder: null
};

let activeStickerTab = 'paint';

let savedColors = ['#B33A2B', '#4A3F2A', '#E85C2A', '#4A6FA5', '#6F7D4E', '#E6B7C2', '#F5E31B', '#1F2DBF'];

let _historyBytes = 0;

let _stateGeneration = 0;

let _needsPagesPanelUpdate = true; // gate for renderPagesPanel() inside render()

// Blob store: holds raw data URL strings keyed by content-derived ref.
// History snapshots store short ref keys instead of multi-KB data URLs.
const _blobStore = new Map();

let _lastSelectedForToggles = undefined;
