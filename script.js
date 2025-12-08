const apiKeyInput = document.querySelector('#apiKey');
const sheetUrlInput = document.querySelector('#sheetUrl');
const publishedIdInput = document.querySelector('#publishedId');
const loadBtn = document.querySelector('#loadBtn');
const clearCacheBtn = document.querySelector('#clearCacheBtn');
const statusEl = document.querySelector('#status');
const cacheState = document.querySelector('#cacheState');
const sheetFrame = document.querySelector('#sheetFrame');
const dimensionBadge = document.querySelector('#dimensionBadge');
const sheetBadge = document.querySelector('#sheetBadge');
const sheetSelect = document.querySelector('#sheetSelect');
const sectionsContainer = document.querySelector('#sectionsContainer');
const sectionStatus = document.querySelector('#sectionStatus');
const sectionsFileInput = document.querySelector('#sectionsFile');
const mappingFileInput = document.querySelector('#mappingFile');
const snowCanvas = document.querySelector('#snowCanvas');

const STORAGE_KEY = 'gs_api_key';
const STORAGE_SHEET = 'gs_last_sheet';
const STORAGE_PUB = 'gs_published_sheet';
const STORAGE_CSV = 'gs_csv_text';
const STORAGE_SECTION_JSON = 'gs_section_json';
const STORAGE_MAP = 'gs_csv_map';

// Optional CORS proxy to bypass policy when serving statically. Example:
const CORS_PROXY = 'https://cors.isomorphic-fetch.workers.dev/?u=';

let sheetsMeta = [];
let currentSheetId = '';
let currentPublishedId = '';
let sectionJsonStore = {};
let mappingById = {};
let sheetRowsCache = {};

const setStatus = (msg, type = 'info') => {
  const color = type === 'error' ? 'danger' : type === 'warn' ? 'warn' : 'muted';
  statusEl.innerHTML = `<span class="${color}">${msg}</span>`;
};

const setCacheBadge = () => {
  const hasKey = !!localStorage.getItem(STORAGE_KEY);
  const hasSheet = !!localStorage.getItem(STORAGE_SHEET);
  const hasPub = !!localStorage.getItem(STORAGE_PUB);
  cacheState.textContent = `Cache: ${hasKey ? 'API key' : 'none'}${hasSheet ? ' + sheet' : ''}${hasPub ? ' + pub' : ''}`;
};

const loadSectionJsonCache = () => {
  try {
    const raw = localStorage.getItem(STORAGE_SECTION_JSON);
    sectionJsonStore = raw ? JSON.parse(raw) : {};
  } catch {
    sectionJsonStore = {};
  }
};

const saveSectionJsonCache = () => {
  localStorage.setItem(STORAGE_SECTION_JSON, JSON.stringify(sectionJsonStore));
};

const loadCache = () => {
  const cachedKey = localStorage.getItem(STORAGE_KEY);
  if (cachedKey) apiKeyInput.value = cachedKey;
  const cachedSheet = localStorage.getItem(STORAGE_SHEET);
  if (cachedSheet) sheetUrlInput.value = cachedSheet;
  const cachedPub = localStorage.getItem(STORAGE_PUB);
  if (cachedPub) publishedIdInput.value = cachedPub;
  loadSectionJsonCache();
  setCacheBadge();
};

const parseSheetId = (input) => {
  if (!input) return null;
  const urlMatch = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)(?:\/|$)/);
  const id = urlMatch ? urlMatch[1] : null;
  const gidMatch = input.match(/[?&#]gid=([0-9]+)/);
  const gid = gidMatch ? gidMatch[1] : '0';
  if (id) return { id, gid };
  if (/^[a-zA-Z0-9-_]{20,}$/.test(input)) return { id: input, gid: '0' };
  return null;
};

const fetchJson = async (url) => {
  const tryFetch = async (target) => {
    const res = await fetch(target);
    if (!res.ok) {
      const msg = await res.text();
      throw new Error(`Request failed ${res.status}: ${msg}`);
    }
    return res.json();
  };
  if (CORS_PROXY) {
    try {
      return await tryFetch(`${CORS_PROXY}${encodeURIComponent(url)}`);
    } catch (err) {
      console.warn('CORS proxy fetch failed, retrying direct', err);
    }
  }
  return tryFetch(url);
};

const buildEmbedUrl = ({ id, gid, publishedId }) => {
  // If published ID is provided, use the /d/e/{publishedId} route.
  if (publishedId) {
    return `https://docs.google.com/spreadsheets/d/e/${publishedId}/pubhtml?gid=${gid}&single=true&widget=true&headers=false`;
  }
  return `https://docs.google.com/spreadsheets/d/${id}/pubhtml?gid=${gid}&single=true&widget=true&headers=false`;
};

const updateDropdown = (options = [], activeGid = '') => {
  sheetSelect.innerHTML = '';
  if (!options.length) {
    const opt = document.createElement('option');
    opt.value = activeGid || '';
    opt.textContent = activeGid ? `gid ${activeGid}` : 'Chua tai';
    sheetSelect.appendChild(opt);
    return;
  }
  options.forEach(({ gid, title }) => {
    const opt = document.createElement('option');
    opt.value = String(gid);
    opt.textContent = `${title} (gid ${gid})`;
    sheetSelect.appendChild(opt);
  });
  const match = activeGid && options.find(o => String(o.gid) === String(activeGid));
  sheetSelect.value = match ? String(activeGid) : String(options[0].gid);
};

const embedSheet = ({ id, gid, title, publishedId }) => {
  currentSheetId = id;
  currentPublishedId = publishedId || currentPublishedId || '';
  const embedUrl = buildEmbedUrl({ id, gid, publishedId: currentPublishedId });
  sheetFrame.src = embedUrl;
  sheetBadge.textContent = `Sheet: ${title || `gid ${gid}`}`;
  dimensionBadge.textContent = `Sheet ID: ${id}`;
};

const parseCsv = (text) => {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if (!lines.length) return [];
  const [headerLine, ...rows] = lines;
  const headers = headerLine.split(',').map(h => h.trim().toLowerCase());
  const idxName = headers.indexOf('name');
  const idxId = headers.indexOf('id');
  if (idxName === -1 || idxId === -1) return [];
  return rows.map(line => {
    const parts = line.split(',').map(p => p.trim());
    return {
      name: parts[idxName],
      id: parts[idxId],
    };
  }).filter(r => r.name && r.id);
};

const sectionKey = ({ id, name }) => `${id}::${name}`;

const parseMappingCsv = (text) => {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  if (!lines.length) return [];
  const [headerLine, ...rows] = lines;
  const headers = headerLine.split(',').map(h => h.replace(/^\ufeff/, '').trim().toLowerCase());
  const idxId = headers.indexOf('id');
  const idxCol = headers.indexOf('column');
  const idxJson = headers.indexOf('json_key');
  const idxType = headers.indexOf('type');
  if (idxId === -1 || idxCol === -1 || idxJson === -1) return [];
  return rows.map(line => {
    const parts = line.split(',').map(p => p.trim());
    return {
      id: parts[idxId],
      column: parts[idxCol],
      jsonKey: parts[idxJson],
      type: idxType >= 0 ? parts[idxType] : '',
    };
  }).filter(r => r.id && r.column && r.jsonKey);
};

const buildMappingIndex = (entries) => {
  mappingById = {};
  entries.forEach(e => {
    if (!mappingById[e.id]) mappingById[e.id] = [];
    mappingById[e.id].push(e);
  });
};

const buildTreeNode = (value, path, labelOverride) => {
  const li = document.createElement('li');
  const labelText = labelOverride ?? path.split('.').slice(-1)[0];
  if (value && typeof value === 'object') {
    const isArray = Array.isArray(value);
    const header = document.createElement('div');
    header.className = 'pair';
    const keySpan = document.createElement('span');
    keySpan.className = 'key';
    keySpan.textContent = labelText;
    keySpan.dataset.path = path;
    const typeSpan = document.createElement('span');
    typeSpan.className = 'muted';
    typeSpan.textContent = isArray ? '[ ]' : '{ }';
    header.appendChild(keySpan);
    header.appendChild(typeSpan);
    li.appendChild(header);
    const ul = document.createElement('ul');
    const entries = isArray ? value.map((v, i) => [i, v]) : Object.entries(value);
    entries.forEach(([k, v]) => {
      ul.appendChild(buildTreeNode(v, path ? `${path}.${k}` : String(k), String(k)));
    });
    li.appendChild(ul);
  } else {
    const row = document.createElement('div');
    row.className = 'pair';
    const keySpan = document.createElement('span');
    keySpan.className = 'key';
    keySpan.textContent = labelText;
    keySpan.dataset.path = path;
    const input = document.createElement('input');
    input.className = 'value';
    input.value = value ?? '';
    input.dataset.path = path;
    row.appendChild(keySpan);
    row.appendChild(input);
    li.appendChild(row);
  }
  return li;
};

const renderJsonTree = (obj, container, sectionName) => {
  container.innerHTML = '';
  const tree = document.createElement('div');
  tree.className = 'tree';
  const ul = document.createElement('ul');
  const root = buildTreeNode(obj, sectionName, sectionName);
  ul.appendChild(root);
  tree.appendChild(ul);
  container.appendChild(tree);
};

const highlightTree = (container, highlights = []) => {
  if (!highlights.length) return;
  const map = new Map(highlights.map(h => [h.path, h.type || 'existing']));
  container.querySelectorAll('.key[data-path]').forEach(el => {
    const typ = map.get(el.dataset.path);
    if (!typ) return;
    el.classList.add(typ === 'new' ? 'highlight-new' : 'highlight');
  });
};

const setSectionStatus = (msg, type = 'info') => {
  const color = type === 'error' ? 'danger' : type === 'warn' ? 'warn' : '';
  sectionStatus.textContent = msg;
  sectionStatus.className = `badge${color ? ' ' + color : ''}`;
};

const createSection = ({ name, id }) => {
  const details = document.createElement('details');
  details.className = 'section';
  const summary = document.createElement('summary');
  summary.innerHTML = `<span>${id} - ${name}</span>`;
  details.appendChild(summary);

  const body = document.createElement('div');
  body.className = 'body';

  const cacheKey = sectionKey({ id, name });
  const textarea = document.createElement('textarea');
  const cachedText = sectionJsonStore[cacheKey];
  textarea.value = typeof cachedText === 'string' && cachedText.trim() ? cachedText : '{}';

  const jsonHeader = document.createElement('div');
  jsonHeader.className = 'muted';
  jsonHeader.textContent = 'Input JSON';

  const applyBtn = document.createElement('button');
  applyBtn.className = 'apply';
  applyBtn.type = 'button';
  applyBtn.textContent = 'Apply From Sheet';

  const copyBtn = document.createElement('button');
  copyBtn.className = 'copy';
  copyBtn.type = 'button';
  copyBtn.textContent = 'Copy JSON';

  const actionRow = document.createElement('div');
  actionRow.className = 'action-row';
  actionRow.appendChild(applyBtn);
  actionRow.appendChild(copyBtn);

  const treeMount = document.createElement('div');

  let lastHighlights = [];

  const writeCache = (val) => {
    sectionJsonStore[cacheKey] = val;
    saveSectionJsonCache();
  };

  const parseAndRender = (highlightPaths = null) => {
    try {
      const parsed = textarea.value.trim() ? JSON.parse(textarea.value) : {};
      const normalized = textarea.value.trim() ? JSON.stringify(parsed, null, 2) : '{}';
      textarea.value = normalized;
      renderJsonTree(parsed, treeMount, name);
      if (highlightPaths !== null) lastHighlights = highlightPaths;
      highlightTree(treeMount, lastHighlights);
      textarea.dataset.error = '';
      writeCache(normalized);
    } catch (err) {
      textarea.dataset.error = err.message;
    }
  };

  textarea.addEventListener('input', parseAndRender);
  const commitTreeInput = (target) => {
    const path = target.dataset.path;
    const raw = textarea.value.trim() ? JSON.parse(textarea.value) : {};
    const segments = path.split('.').slice(1); // remove section name prefix
    let cursor = raw;
    for (let i = 0; i < segments.length - 1; i++) {
      const key = segments[i];
      if (!(key in cursor) || typeof cursor[key] !== 'object') cursor[key] = {};
      cursor = cursor[key];
    }
    const last = segments[segments.length - 1];
    cursor[last] = target.value;
    const normalized = JSON.stringify(raw, null, 2);
    textarea.value = normalized;
    renderJsonTree(raw, treeMount, name);
    highlightTree(treeMount, lastHighlights);
    target.classList.add('edited');
    setTimeout(() => target.classList.remove('edited'), 1000);
    writeCache(normalized);
  };

  treeMount.addEventListener('keydown', (e) => {
    const target = e.target;
    if (!target.classList.contains('value')) return;
    if (e.key === 'Enter') {
      e.preventDefault();
      commitTreeInput(target);
    }
  });

  treeMount.addEventListener('blur', (e) => {
    const target = e.target;
    if (!target.classList.contains('value')) return;
    commitTreeInput(target);
  }, true);

  applyBtn.addEventListener('click', () => {
    applyMappingFromSheet({ id, name, textarea, treeMount });
  });

  copyBtn.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(textarea.value);
      setSectionStatus(`Đã copy JSON của ${id}`, 'info');
    } catch (err) {
      console.error(err);
      setSectionStatus('Lỗi copy vào clipboard.', 'error');
    }
  });

  parseAndRender();
  body.appendChild(actionRow);
  body.appendChild(jsonHeader);
  body.appendChild(textarea);
  body.appendChild(treeMount);
  details.appendChild(body);
  return details;
};

const renderSections = (rows) => {
  sectionsContainer.innerHTML = '';
  if (!rows.length) {
    setSectionStatus('Khong co du lieu tu CSV', 'warn');
    return;
  }
  rows.forEach(({ name, id }) => {
    const sec = createSection({ name, id });
    sectionsContainer.appendChild(sec);
    sec.addEventListener('toggle', () => {
      if (!sec.open) return;
      sectionsContainer.querySelectorAll('details.section').forEach(other => {
        if (other !== sec) other.open = false;
      });
    });
  });
  setSectionStatus(`Da tao ${rows.length} section`, 'info');
};

const handleCsvText = (text) => {
  localStorage.setItem(STORAGE_CSV, text);
  const rows = parseCsv(text);
  renderSections(rows);
  if (!rows.length) setSectionStatus('CSV khong hop le (can cot name,id)', 'error');
};

const renderMappingTable = () => {
  if (!mappingById || !Object.keys(mappingById).length) {
    setSectionStatus('Khong co mapping', 'warn');
    return;
  }
  const container = document.createElement('div');
  container.className = 'mapping-panel';
  const table = document.createElement('table');
  table.className = 'map-table';
  table.innerHTML = '<thead><tr><th>ID</th><th>Column</th><th>Type</th><th>json_key</th></tr></thead>';
  const tbody = document.createElement('tbody');
  Object.entries(mappingById).forEach(([id, entries]) => {
    entries.forEach((row, idx) => {
      const tr = document.createElement('tr');
      const idCell = idx === 0 ? `<td rowspan="${entries.length}">${id}</td>` : '';
      tr.innerHTML = `${idCell}<td>${row.column}</td><td>${row.type || '-'}</td><td>${row.jsonKey}</td>`;
      tbody.appendChild(tr);
    });
  });
  table.appendChild(tbody);
  container.appendChild(table);
  const mappingHost = document.querySelector('#mappingHost');
  if (mappingHost) {
    mappingHost.innerHTML = '';
    mappingHost.appendChild(container);
  }
};

const readLocalFile = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(reader.result.toString());
  reader.onerror = reject;
  reader.readAsText(file);
});

const autoFetchConfig = async () => {
  try {
    const sectionsUrl = CORS_PROXY ? `${CORS_PROXY}${encodeURIComponent('Configs/sections-config.csv')}` : 'Configs/sections-config.csv';
    const mappingUrl = CORS_PROXY ? `${CORS_PROXY}${encodeURIComponent('Configs/config_mapping.csv')}` : 'Configs/config_mapping.csv';
    const [sectionsRes, mapRes] = await Promise.all([
      fetch(sectionsUrl, { cache: 'no-cache' }),
      fetch(mappingUrl, { cache: 'no-cache' }),
    ]);
    if (!sectionsRes.ok || !mapRes.ok) throw new Error('fetch config failed');
    const [sectionsText, mapText] = await Promise.all([sectionsRes.text(), mapRes.text()]);
    localStorage.setItem(STORAGE_MAP, mapText);
    localStorage.setItem(STORAGE_CSV, sectionsText);
    buildMappingIndex(parseMappingCsv(mapText));
    handleCsvText(sectionsText);
    renderMappingTable();
    setSectionStatus('Da load config tu file du an.', 'info');
    return true;
  } catch (err) {
    console.warn('Auto fetch config failed, fallback to cache', err);
    return false;
  }
};

const loadProjectConfig = async () => {
  const cachedMap = localStorage.getItem(STORAGE_MAP);
  const cachedCsv = localStorage.getItem(STORAGE_CSV);
  if (cachedMap) buildMappingIndex(parseMappingCsv(cachedMap));
  if (cachedCsv) {
    handleCsvText(cachedCsv);
    renderMappingTable();
    setSectionStatus('Da load config tu cache. Chon file hoac nhan Default de cap nhat.', 'info');
  } else {
    setSectionStatus('Chon file sections/mapping hoac nhan Default de tai config.', 'warn');
  }
};

const loadSheetView = async () => {
  const apiKey = apiKeyInput.value.trim();
  const sheetInput = sheetUrlInput.value.trim();
  const publishedId = publishedIdInput.value.trim();
  const parsed = parseSheetId(sheetInput);
  const gidFromInput = parsed?.gid || '0';

  if (!parsed && !publishedId) return setStatus('Can URL/ID sheet hoac published sheet ID.', 'error');

  if (apiKey) localStorage.setItem(STORAGE_KEY, apiKey);
  if (sheetInput) localStorage.setItem(STORAGE_SHEET, sheetInput);
  if (publishedId) localStorage.setItem(STORAGE_PUB, publishedId);
  setCacheBadge();

  const embedFallback = () => {
    sheetsMeta = [];
    updateDropdown([], gidFromInput);
    const embedId = parsed?.id || publishedId || '';
    if (!embedId) {
      setStatus('Khong co ID de nhung sheet.', 'error');
      return;
    }
    embedSheet({ id: embedId, gid: gidFromInput, title: '', publishedId });
    setStatus('Dang hien thi web view sheet. Neu khong thay du lieu, hay publish sheet hoac dang nhap Google.');
  };

  if (!apiKey || !parsed?.id) {
    embedFallback();
    setStatus('Khong co API key hoac sheet ID, khong the lay danh sach tab. Dang hien thi gid hien tai.', 'warn');
    return;
  }

  try {
    setStatus('Dang lay danh sach sheet tabs...');
    const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${parsed.id}?fields=sheets.properties(sheetId,title)&key=${apiKey}`;
    const meta = await fetchJson(metaUrl);
    sheetsMeta = (meta.sheets || []).map(s => ({
      gid: String(s.properties.sheetId),
      title: s.properties.title || `Sheet ${s.properties.sheetId}`,
    }));

    if (!sheetsMeta.length) {
      embedFallback();
      setStatus('Khong tim thay tab nao, dung gid hien tai.', 'warn');
      return;
    }

    const target = sheetsMeta.find(s => String(s.gid) === String(parsed.gid)) || sheetsMeta[0];
    updateDropdown(sheetsMeta, target.gid);
    embedSheet({ id: parsed.id, gid: target.gid, title: target.title, publishedId });
    setStatus('Dang hien thi web view sheet. Neu khong thay du lieu, hay publish sheet hoac dang nhap Google.');
  } catch (err) {
    console.error(err);
    embedFallback();
    setStatus(`Loi lay danh sach sheet: ${err.message}`, 'error');
  }
};

loadBtn.addEventListener('click', loadSheetView);
sheetUrlInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') loadSheetView(); });
apiKeyInput.addEventListener('input', () => {
  localStorage.setItem(STORAGE_KEY, apiKeyInput.value.trim());
  setCacheBadge();
});
clearCacheBtn.addEventListener('click', () => {
  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(STORAGE_SHEET);
  localStorage.removeItem(STORAGE_PUB);
  localStorage.removeItem(STORAGE_CSV);
  localStorage.removeItem(STORAGE_SECTION_JSON);
  apiKeyInput.value = '';
  sheetUrlInput.value = '';
  publishedIdInput.value = '';
  sectionJsonStore = {};
  setCacheBadge();
  sheetBadge.textContent = 'Sheet: -';
  dimensionBadge.textContent = 'Embed';
  sheetSelect.innerHTML = '<option value="">Chua tai</option>';
  setStatus('Cache da duoc xoa.');
  sectionsContainer.innerHTML = '';
  setSectionStatus('Chua tai', 'warn');
});

sheetSelect.addEventListener('change', () => {
  if (!currentSheetId) return;
  const gid = sheetSelect.value || '0';
  const meta = sheetsMeta.find(s => String(s.gid) === String(gid));
  embedSheet({ id: currentSheetId, gid, title: meta?.title, publishedId: currentPublishedId });
  setStatus(`Dang hien thi sheet: ${meta?.title || `gid ${gid}`}`);
});

loadCache();
loadProjectConfig();

sectionsFileInput?.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const text = await readLocalFile(file);
    handleCsvText(text);
    setSectionStatus('Đã load sections-config từ file.', 'info');
  } catch (err) {
    console.error(err);
    setSectionStatus('Lỗi đọc sections-config.', 'error');
  }
});

// Snow effect
const initSnow = () => {
  if (!snowCanvas) return;
  const ctx = snowCanvas.getContext('2d');
  let width = window.innerWidth;
  let height = window.innerHeight;
  let flakes = [];

  const resize = () => {
    width = window.innerWidth;
    height = window.innerHeight;
    snowCanvas.width = width;
    snowCanvas.height = height;
  };
  resize();
  window.addEventListener('resize', resize);

  const createFlakes = (count = 120) => {
    flakes = Array.from({ length: count }, () => ({
      x: Math.random() * width,
      y: Math.random() * height,
      r: Math.random() * 2 + 1,
      d: Math.random() * 0.8 + 0.2,
      sway: Math.random() * 1,
    }));
  };
  createFlakes();

  const draw = () => {
    ctx.clearRect(0, 0, width, height);
    ctx.fillStyle = 'rgba(255,255,255,0.8)';
    flakes.forEach(f => {
      ctx.beginPath();
      ctx.arc(f.x, f.y, f.r, 0, Math.PI * 2);
      ctx.fill();
    });
    update();
    requestAnimationFrame(draw);
  };

  const update = () => {
    flakes.forEach(f => {
      f.y += f.d * 2;
      f.x += Math.sin(f.y * 0.01) * f.sway;
      if (f.y > height) {
        f.y = -5;
        f.x = Math.random() * width;
      }
    });
  };

  draw();
};

initSnow();

mappingFileInput?.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const text = await readLocalFile(file);
    localStorage.setItem(STORAGE_MAP, text);
    buildMappingIndex(parseMappingCsv(text));
    renderMappingTable();
    setSectionStatus('Đã load mapping từ file.', 'info');
  } catch (err) {
    console.error(err);
    setSectionStatus('Lỗi đọc mapping.', 'error');
  }
});

const fetchSheetRows = async (gid) => {
  const cacheKey = `${currentSheetId || ''}:${gid}`;
  if (sheetRowsCache[cacheKey]) return sheetRowsCache[cacheKey];
  const apiKey = apiKeyInput.value.trim();
  if (!apiKey || !currentSheetId) throw new Error('Thiếu API key hoặc sheet ID.');
  const title = sheetsMeta.find(s => String(s.gid) === String(gid))?.title;
  if (!title) throw new Error('Không tìm thấy tab hiện tại.');
  const range = `'${title}'`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${currentSheetId}/values/${encodeURIComponent(range)}?majorDimension=ROWS&key=${apiKey}`;
  const data = await fetchJson(url);
  const values = data.values || [];
  if (!values.length) return { headers: [], rows: [] };

  // Find header row that contains ID column
  let headerRowIdx = values.findIndex(r => r.some(cell => (cell || '').trim().toLowerCase() === 'id'));
  if (headerRowIdx === -1) headerRowIdx = 0;
  const headers = (values[headerRowIdx] || []).map(h => (h || '').trim());
  const headerIdx = headers.reduce((acc, h, idx) => {
    acc[h.toLowerCase()] = idx;
    return acc;
  }, {});

  const rows = values.slice(headerRowIdx + 1).map(row => {
    const obj = {};
    headers.forEach((h, idx) => { obj[h.toLowerCase()] = row[idx]; });
    return obj;
  });
  const result = { headers, headerIdx, rows };
  sheetRowsCache[cacheKey] = result;
  return result;
};

const castValue = (raw, type) => {
  if (raw === undefined || raw === null) return undefined;
  const val = String(raw).trim();
  if (val === '') return undefined;
  if (type?.toLowerCase() === 'int') return Number.parseInt(val, 10);
  if (type?.toLowerCase() === 'float') return Number.parseFloat(val);
  return val;
};

const setByPath = (obj, path, value) => {
  const segments = path.split('.');
  let cursor = obj;
  for (let i = 0; i < segments.length - 1; i++) {
    const k = segments[i];
    if (!cursor[k] || typeof cursor[k] !== 'object') cursor[k] = {};
    cursor = cursor[k];
  }
  cursor[segments[segments.length - 1]] = value;
};

const applyMappingFromSheet = async ({ id, name, textarea, treeMount }) => {
  const mappings = mappingById[id] || [];
  if (!mappings.length) {
    setSectionStatus(`Không có mapping cho ID ${id}`, 'warn');
    return;
  }
  const gid = sheetSelect.value || '0';
  const apiKey = apiKeyInput.value.trim();
  if (!apiKey || !currentSheetId) {
    setSectionStatus('Thiếu API key hoặc sheet ID để đọc sheet.', 'error');
    return;
  }
  try {
    const sheetData = await fetchSheetRows(gid);
    const idColIdx = sheetData.headerIdx?.['id'];
    if (idColIdx === undefined) throw new Error('Sheet không có cột ID');
    const targetRow = sheetData.rows.find(r => (r.id || '').trim() === id);
    if (!targetRow) {
      setSectionStatus(`Không tìm thấy dòng có ID ${id}`, 'warn');
      return;
    }

    const currentJson = textarea.value.trim() ? JSON.parse(textarea.value) : {};
    const highlightMeta = [];

    mappings.forEach(entry => {
      const colVal = targetRow[entry.column.toLowerCase()];
      if (colVal === undefined) return;
      const keys = entry.jsonKey.split('|');
      const types = (entry.type || '').split('|');
      const parts = colVal.split(/[|;]/); // support both | and ; separators from sheet value
      keys.forEach((keyPath, idx) => {
        const t = types[idx] || types[0] || 'string';
        const v = castValue(parts[idx] ?? colVal, t);
        if (v === undefined) return;
        // detect whether field existed before
        const segs = keyPath.split('.');
        let cursor = currentJson;
        for (let i = 0; i < segs.length - 1; i++) {
          const k = segs[i];
          if (!(k in cursor) || typeof cursor[k] !== 'object') {
            cursor[k] = {};
          }
          cursor = cursor[k];
        }
        const lastKey = segs[segs.length - 1];
        const existed = Object.prototype.hasOwnProperty.call(cursor, lastKey);
        setByPath(currentJson, keyPath, v);
        highlightMeta.push({ path: `${name}.${keyPath}`, type: existed ? 'existing' : 'new' });
      });
    });

    const normalized = JSON.stringify(currentJson, null, 2);
    textarea.value = normalized;
    renderJsonTree(currentJson, treeMount, name);
    highlightTree(treeMount, highlightMeta);
    lastHighlights = highlightMeta;
    sectionJsonStore[`${id}::${name}`] = normalized;
    saveSectionJsonCache();
    setSectionStatus(`Đã áp dụng từ sheet cho ID ${id}`, 'info');
  } catch (err) {
    console.error(err);
    setSectionStatus(`Lỗi áp dụng từ sheet: ${err.message}`, 'error');
  }
};
