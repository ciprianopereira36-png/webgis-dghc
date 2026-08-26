/* ================================================================
   WEB GIS DGHC – script.js
   DGHC Habitasaun Timor-Leste
   Versão 1.0 – 2025
   ================================================================ */

/* ─────────────────────────────────────────────────────────────────
   COLUMN MAPPING CONFIGURATION
   Sesuaikan nama kolom Excel di sini jika struktur berbeda.
   Masukkan semua kemungkinan nama kolom (case-insensitive).
───────────────────────────────────────────────────────────────── */
const COLUMN_MAP = {
  id:             ['nu uma', 'id', 'id_uma', 'house_id', 'no', 'nomor', 'no uma', 'id uma'],
  latitude:       ['latitude', 'lat', 'y', 'lintang'],
  longitude:      ['longitude', 'long', 'lon', 'x', 'bujur'],
  municipality:   ['municipality', 'munisipiu', 'municipio', 'kabupaten', 'munisípiu'],
  post:           ['administrative post', 'postu administrativu', 'postu', 'kecamatan', 'admin post'],
  suco:           ['suco', 'suku', 'desa', 'kelurahan'],
  aldeia:         ['aldeia', 'dusun', 'kampung', 'lingkungan'],
  status:         ['status', 'kondisi', 'kategori'],
  program:        ['program', 'programa', 'projetu'],
  year:           ['year', 'tinan', 'tahun', 'ano'],
};

/* ─────────────────────────────────────────────────────────────────
   IMAGE PATH CONFIGURATION
   Sesuaikan folder foto jika berbeda.
───────────────────────────────────────────────────────────────── */
const IMAGE_CONFIG = {
  // Folder foto. Gunakan '' (kosong) jika foto di root.
  folder: 'images',
  // Ekstensi yang dicoba secara berurutan
  extensions: ['jpg', 'jpeg', 'png', 'webp'],
  // Padding ID (mis: 1 → '01')
  padLength: 2,
};

/* ═══════════════════════════════════════════════════════════════
   APLIKASI STATE
═══════════════════════════════════════════════════════════════ */
const APP = {
  allData:      [],   // semua data dari Excel
  filtered:     [],   // data setelah filter
  map:          null, // instance Leaflet
  clusterGroup: null, // MarkerClusterGroup
  markers:      {},   // { id: marker }
  charts:       {},   // instance Chart.js
  userMarker:   null, // marker lokasi pengguna
  routeControl: null, // instance routing
  currentView:  'dashboard',
  pagination: { page: 1, perPage: 15 },
  sort:       { col: 'no', dir: 'asc' },
  tableSearch: '',
};

/* ═══════════════════════════════════════════════════════════════
   UTILS
═══════════════════════════════════════════════════════════════ */
function sanitize(str) {
  if (str == null) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function escAttr(str) {
  return String(str ?? '').replace(/'/g, '&#39;').replace(/"/g, '&quot;');
}

/** Detect column by COLUMN_MAP key */
function detectCol(headers, mapKey) {
  const names = COLUMN_MAP[mapKey];
  const lc = headers.map(h => String(h).toLowerCase().trim());
  for (const n of names) {
    const idx = lc.indexOf(n.toLowerCase());
    if (idx !== -1) return headers[idx];
  }
  return null;
}

/** Parse DMS string like "  9° 3'5.45\"S" → -9.051514 */
function parseDMS(dmsStr) {
  if (dmsStr == null) return NaN;
  const s = String(dmsStr).trim();
  // Try decimal first
  const dec = parseFloat(s);
  if (!isNaN(dec) && !s.includes('°')) return dec;

  // DMS pattern: deg° min' sec"[NSEW]  or  -deg min sec
  const m = s.match(/(\d+)[°\s]\s*(\d+)['']\s*([\d.]+)[""]?\s*([NSEWnsew]?)/);
  if (m) {
    let deg = parseFloat(m[1]);
    let min = parseFloat(m[2]);
    let sec = parseFloat(m[3]);
    let dir = m[4].toUpperCase();
    let dd = deg + min / 60 + sec / 3600;
    if (dir === 'S' || dir === 'W') dd = -dd;
    return dd;
  }
  // Try plain number with direction suffix: "-9.051"
  const m2 = s.match(/([-\d.]+)\s*([NSEWnsew]?)/);
  if (m2) {
    let dd = parseFloat(m2[1]);
    if (m2[2].toUpperCase() === 'S' || m2[2].toUpperCase() === 'W') dd = -Math.abs(dd);
    return dd;
  }
  return NaN;
}

/** Validate coordinate for Timor-Leste region (warning, not hard block) */
function validateCoord(lat, lng) {
  if (isNaN(lat) || isNaN(lng)) return 'invalid';
  if (lat < -90 || lat > 90 || lng < -180 || lng > 180) return 'out_of_range';
  // TL warning zone
  if (lat < -11 || lat > -7 || lng < 123 || lng > 128) return 'outside_tl';
  return 'ok';
}

/** Build image path candidates for a given ID */
function getImagePath(id) {
  if (id == null) return null;
  const idStr = String(id).trim();
  const padded = idStr.padStart(IMAGE_CONFIG.padLength, '0');
  const folder = IMAGE_CONFIG.folder ? IMAGE_CONFIG.folder.replace(/\/$/, '') + '/' : '';
  const candidates = [];
  for (const ext of IMAGE_CONFIG.extensions) {
    if (idStr !== padded) candidates.push(`${folder}${padded}.${ext}`);
    candidates.push(`${folder}${idStr}.${ext}`);
  }
  return candidates;
}

/** Return CSS status class */
function statusClass(status) {
  if (!status) return 'status-other';
  const s = String(status).toLowerCase();
  if (s.includes('dignu') || s.includes('layak')) return 'status-dignu';
  if (s.includes('vulnerable') || s.includes('rentan')) return 'status-vulnerable';
  if (s.includes('interven') || s.includes('butuh')) return 'status-intervensi';
  return 'status-other';
}
function markerClass(status) {
  if (!status) return 'marker-default';
  const s = String(status).toLowerCase();
  if (s.includes('dignu') || s.includes('layak')) return 'marker-dignu';
  if (s.includes('vulnerable') || s.includes('rentan')) return 'marker-vulnerable';
  if (s.includes('interven') || s.includes('butuh')) return 'marker-intervensi';
  return 'marker-other';
}

function formatNum(n) { return Number(n).toLocaleString('pt-TL'); }

/* ═══════════════════════════════════════════════════════════════
   NOTIFICATION
═══════════════════════════════════════════════════════════════ */
let notifTimer = null;
function showNotif(msg, type = 'info', duration = 5000) {
  const bar = document.getElementById('notification-bar');
  bar.className = `notification-bar ${type}`;
  bar.innerHTML = `<i class="fa-solid fa-${type === 'success' ? 'circle-check' : type === 'error' ? 'circle-xmark' : type === 'warning' ? 'triangle-exclamation' : 'circle-info'}"></i> ${sanitize(msg)}`;
  bar.classList.remove('hidden');
  clearTimeout(notifTimer);
  notifTimer = setTimeout(() => bar.classList.add('hidden'), duration);
}

/* ═══════════════════════════════════════════════════════════════
   LOADING
═══════════════════════════════════════════════════════════════ */
function showLoading(msg = 'Karga dadus...') {
  document.getElementById('loading-message').textContent = msg;
  document.getElementById('loading-overlay').classList.remove('hidden');
}
function hideLoading() {
  document.getElementById('loading-overlay').classList.add('hidden');
}

/* ═══════════════════════════════════════════════════════════════
   MAP INIT
═══════════════════════════════════════════════════════════════ */
function initMap() {
  // Center on Timor-Leste (Dili region)
  APP.map = L.map('map', {
    center: [-8.874, 125.727],
    zoom: 9,
    zoomControl: false,
  });

  // Zoom control
  L.control.zoom({ position: 'topright' }).addTo(APP.map);

  // Scale
  L.control.scale({ imperial: false }).addTo(APP.map);

  // Base layers
  const osm = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
    maxZoom: 19,
  });

  const satellite = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
    attribution: 'Tiles © Esri — Esri, i-cubed, USDA, USGS, AEX, GeoEye, Getmapping, Aerogrid, IGN, IGP, UPR-EGP',
    maxZoom: 19,
  });

  const cartoLight = L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', {
    attribution: '© OpenStreetMap contributors © CARTO',
    maxZoom: 19,
  });

  osm.addTo(APP.map);

  L.control.layers({
    'OpenStreetMap':  osm,
    'Satellite':      satellite,
    'Carto Light':    cartoLight,
  }, {}, { position: 'topright' }).addTo(APP.map);

  // Cluster group
  APP.clusterGroup = L.markerClusterGroup({
    chunkedLoading: true,
    maxClusterRadius: 60,
    showCoverageOnHover: false,
  });
  APP.map.addLayer(APP.clusterGroup);
}

/* ═══════════════════════════════════════════════════════════════
   EXCEL READER
═══════════════════════════════════════════════════════════════ */
function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

async function readDefaultExcel() {
  try {
    showLoading('Mbaca ficheiru Excel...');
    const resp = await fetch('data/Coordinate_Uma.xlsx');
    if (!resp.ok) throw new Error('File not found');
    const buf = await resp.arrayBuffer();
    const data = new Uint8Array(buf);
    const wb = XLSX.read(data, { type: 'array' });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
    return json;
  } catch (err) {
    // Try root path
    try {
      const resp2 = await fetch('Coordinate_Uma.xlsx');
      if (!resp2.ok) throw new Error('Not found at root either');
      const buf2 = await resp2.arrayBuffer();
      const data2 = new Uint8Array(buf2);
      const wb2 = XLSX.read(data2, { type: 'array' });
      const ws2 = wb2.Sheets[wb2.SheetNames[0]];
      return XLSX.utils.sheet_to_json(ws2, { defval: '' });
    } catch (err2) {
      console.warn('Default Excel not found, waiting for user upload.', err2);
      return null;
    }
  }
}

/* ═══════════════════════════════════════════════════════════════
   PARSE & MAP DATA
═══════════════════════════════════════════════════════════════ */
function processRawRows(rows) {
  if (!rows || rows.length === 0) return [];
  const headers = Object.keys(rows[0]);

  // Detect column names
  const colId   = detectCol(headers, 'id');
  const colLat  = detectCol(headers, 'latitude');
  const colLng  = detectCol(headers, 'longitude');
  const colMun  = detectCol(headers, 'municipality');
  const colPost = detectCol(headers, 'post');
  const colSuco = detectCol(headers, 'suco');
  const colAlde = detectCol(headers, 'aldeia');
  const colStat = detectCol(headers, 'status');
  const colProg = detectCol(headers, 'program');
  const colYear = detectCol(headers, 'year');

  if (!colLat || !colLng) {
    showNotif('⚠ Kolom Latitude/Longitude la hetan. Favor verifika estrutura Excel.', 'warning', 8000);
    return [];
  }

  let invalid = 0;
  let outsideTL = 0;
  const parsed = [];

  rows.forEach((row, idx) => {
    const rawLat = row[colLat] ?? '';
    const rawLng = row[colLng] ?? '';
    const lat = parseDMS(rawLat);
    const lng = parseDMS(rawLng);
    const validity = validateCoord(lat, lng);

    if (validity === 'invalid' || validity === 'out_of_range') {
      invalid++;
      return; // skip
    }
    if (validity === 'outside_tl') outsideTL++;

    const id = colId ? (row[colId] ?? (idx + 1)) : (idx + 1);
    parsed.push({
      _rowIdx:      idx,
      id:           id,
      latitude:     lat,
      longitude:    lng,
      municipality: colMun  ? (row[colMun]  ?? '') : '',
      post:         colPost ? (row[colPost] ?? '') : '',
      suco:         colSuco ? (row[colSuco] ?? '') : '',
      aldeia:       colAlde ? (row[colAlde] ?? '') : '',
      status:       colStat ? (row[colStat] ?? '') : '',
      program:      colProg ? (row[colProg] ?? '') : '',
      year:         colYear ? (row[colYear] ?? '') : '',
      imagePaths:   getImagePath(id),
    });
  });

  if (invalid > 0)    showNotif(`⚠ ${invalid} dadus iha koordinat la válidu (la hatudu iha mapa).`, 'warning', 7000);
  if (outsideTL > 0)  showNotif(`ℹ ${outsideTL} koordinat la iha area Timor-Leste (verifika dadus).`, 'info', 5000);

  return parsed;
}

/* ═══════════════════════════════════════════════════════════════
   LOAD DATA INTO APP
═══════════════════════════════════════════════════════════════ */
async function loadData(rows) {
  showLoading('Prosesa koordinat...');
  await sleep(40);

  const parsed = processRawRows(rows);
  APP.allData = parsed;
  APP.filtered = [...parsed];
  APP.pagination.page = 1;

  showLoading('Kria marker mapa...');
  await sleep(40);

  populateFilters();
  renderMarkers(APP.filtered);

  showLoading('Atualiza KPI no gráfiku...');
  await sleep(40);

  updateKPI(APP.filtered);
  renderCharts(APP.filtered);
  renderTable(APP.filtered);
  renderGaleria(APP.filtered);

  hideLoading();
  showNotif(`✅ Dadus karga ho susesu: ${formatNum(parsed.length)} uma.`, 'success', 5000);
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

/* ═══════════════════════════════════════════════════════════════
   MARKERS
═══════════════════════════════════════════════════════════════ */
function buildMarkerIcon(status) {
  const cls = markerClass(status);
  return L.divIcon({
    className: cls,
    html: `<div class="custom-marker"><i class="fa-solid fa-house"></i></div>`,
    iconSize: [34, 34],
    iconAnchor: [17, 34],
    popupAnchor: [0, -36],
  });
}

function renderMarkers(data) {
  APP.clusterGroup.clearLayers();
  APP.markers = {};

  data.forEach(house => {
    const marker = L.marker([house.latitude, house.longitude], {
      icon: buildMarkerIcon(house.status),
    });
    marker.bindPopup(() => buildPopupHTML(house), {
      maxWidth: 300,
      className: 'house-popup',
    });
    marker.on('popupopen', () => {
      // Lazy-load image verification happens in popup
    });
    APP.markers[house.id] = marker;
    APP.clusterGroup.addLayer(marker);
  });
}

function buildPopupHTML(house) {
  const imgPaths = house.imagePaths || [];
  const imgSrc = imgPaths.length ? escAttr(imgPaths[0]) : '';
  const imgHtml = imgSrc
    ? `<img class="popup-photo" src="${imgSrc}" alt="Foto Uma ${escAttr(house.id)}" loading="lazy"
         onerror="this.style.display='none';this.nextElementSibling.style.display='flex'"  />
       <div class="popup-photo-placeholder" style="display:none">
         <i class="fa-solid fa-image"></i><span>Foto la disponível</span>
       </div>`
    : `<div class="popup-photo-placeholder">
         <i class="fa-solid fa-image"></i><span>Foto la disponível</span>
       </div>`;

  const sc = statusClass(house.status);
  const sid = escAttr(house.id);

  return `<div class="popup-card">
    ${imgHtml}
    <div class="popup-body">
      <div class="popup-id">UMA-${sanitize(String(house.id).padStart(3,'0'))}</div>
      ${house.municipality ? `<div class="popup-row"><strong>Munisípiu:</strong><span>${sanitize(house.municipality)}</span></div>` : ''}
      ${house.post         ? `<div class="popup-row"><strong>Postu:</strong><span>${sanitize(house.post)}</span></div>` : ''}
      ${house.suco         ? `<div class="popup-row"><strong>Suco:</strong><span>${sanitize(house.suco)}</span></div>` : ''}
      ${house.aldeia       ? `<div class="popup-row"><strong>Aldeia:</strong><span>${sanitize(house.aldeia)}</span></div>` : ''}
      <div class="popup-row"><strong>Lat/Lng:</strong><span>${house.latitude.toFixed(5)}, ${house.longitude.toFixed(5)}</span></div>
      ${house.status ? `<div><span class="popup-status ${sc}">${sanitize(house.status)}</span></div>` : ''}
    </div>
    <div class="popup-actions">
      <button class="btn-popup btn-detail"   onclick="openDetail('${sid}')"><i class="fa-solid fa-eye"></i> Detallu</button>
      <button class="btn-popup btn-navigate" onclick="navigateTo('${sid}')"><i class="fa-solid fa-route"></i> Nabegasaun</button>
    </div>
  </div>`;
}

/* ═══════════════════════════════════════════════════════════════
   FILTERS
═══════════════════════════════════════════════════════════════ */
function populateFilters() {
  const fields = ['municipality', 'post', 'suco', 'aldeia', 'status', 'program', 'year'];
  const ids = {
    municipality: 'filter-municipality', post: 'filter-post', suco: 'filter-suco',
    aldeia: 'filter-aldeia', status: 'filter-status', program: 'filter-program', year: 'filter-year',
  };
  const labels = {
    municipality: 'Hotu-hotu Munisípiu', post: 'Hotu-hotu Postu',
    suco: 'Hotu-hotu Suco', aldeia: 'Hotu-hotu Aldeia',
    status: 'Hotu-hotu Status', program: 'Hotu-hotu Programa', year: 'Hotu-hotu Tinan',
  };

  fields.forEach(field => {
    const sel = document.getElementById(ids[field]);
    const values = [...new Set(APP.allData.map(d => d[field]).filter(Boolean))].sort();
    sel.innerHTML = `<option value="">${labels[field]}</option>`;
    values.forEach(v => {
      const opt = document.createElement('option');
      opt.value = v; opt.textContent = v;
      sel.appendChild(opt);
    });
  });
}

function applyFilters() {
  const mun   = document.getElementById('filter-municipality').value;
  const post  = document.getElementById('filter-post').value;
  const suco  = document.getElementById('filter-suco').value;
  const aldeia= document.getElementById('filter-aldeia').value;
  const status= document.getElementById('filter-status').value;
  const prog  = document.getElementById('filter-program').value;
  const year  = document.getElementById('filter-year').value;
  const q     = document.getElementById('search-input').value.trim().toLowerCase();

  APP.filtered = APP.allData.filter(d => {
    if (mun    && d.municipality !== mun)   return false;
    if (post   && d.post !== post)           return false;
    if (suco   && d.suco !== suco)           return false;
    if (aldeia && d.aldeia !== aldeia)       return false;
    if (status && d.status !== status)       return false;
    if (prog   && d.program !== prog)        return false;
    if (year   && String(d.year) !== year)   return false;
    if (q) {
      const haystack = [d.id, d.municipality, d.post, d.suco, d.aldeia, d.status]
        .join(' ').toLowerCase();
      if (!haystack.includes(q)) return false;
    }
    return true;
  });

  APP.pagination.page = 1;
  renderMarkers(APP.filtered);
  updateKPI(APP.filtered);
  renderCharts(APP.filtered);
  renderTable(APP.filtered);
  renderGaleria(APP.filtered);
}

/* ═══════════════════════════════════════════════════════════════
   KPI
═══════════════════════════════════════════════════════════════ */
function updateKPI(data) {
  document.getElementById('kpi-total').textContent = formatNum(data.length);

  const dignu = data.filter(d => String(d.status).toLowerCase().includes('dignu') ||
                                  String(d.program).toLowerCase().includes('dignu')).length;
  document.getElementById('kpi-dignu').textContent = formatNum(dignu);

  const vuln = data.filter(d => String(d.status).toLowerCase().includes('vulnerable') ||
                                 String(d.status).toLowerCase().includes('rentan')).length;
  document.getElementById('kpi-vulnerable').textContent = formatNum(vuln);

  const munSet = new Set(data.map(d => d.municipality).filter(Boolean));
  document.getElementById('kpi-municipality').textContent = formatNum(munSet.size || (data.length > 0 ? '—' : '0'));
}

/* ═══════════════════════════════════════════════════════════════
   CHARTS
═══════════════════════════════════════════════════════════════ */
function countBy(data, field) {
  const map = {};
  data.forEach(d => {
    const k = d[field] || 'N/D';
    map[k] = (map[k] || 0) + 1;
  });
  return map;
}

function buildBarData(countObj, limit = 12) {
  const entries = Object.entries(countObj).sort((a, b) => b[1] - a[1]).slice(0, limit);
  return { labels: entries.map(e => e[0]), values: entries.map(e => e[1]) };
}

const CHART_COLORS = [
  '#3b82f6','#22c55e','#f59e0b','#ef4444','#8b5cf6',
  '#06b6d4','#f97316','#84cc16','#ec4899','#14b8a6',
  '#a855f7','#6366f1',
];

function renderCharts(data) {
  renderBarChart('chart-municipality', data, 'municipality', 'Uma ba Munisípiu');
  renderPieChart('chart-status', data, 'status', 'Status Uma');
  // Relatóriu charts
  renderBarChart('chart-r-municipality', data, 'municipality', 'Uma ba Munisípiu');
  renderPieChart('chart-r-status', data, 'status', 'Status Uma');
  renderBarChart('chart-r-program', data, 'program', 'Uma ba Programa');
  renderBarChart('chart-r-year', data, 'year', 'Uma ba Tinan');
}

function renderBarChart(canvasId, data, field, label) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const countObj = countBy(data, field);
  const { labels, values } = buildBarData(countObj, 13);

  if (APP.charts[canvasId]) APP.charts[canvasId].destroy();
  APP.charts[canvasId] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label,
        data: values,
        backgroundColor: CHART_COLORS.slice(0, labels.length),
        borderRadius: 5,
        borderSkipped: false,
      }],
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { font: { size: 10 } } },
        y: { beginAtZero: true, ticks: { font: { size: 10 } } },
      },
    },
  });
}

function renderPieChart(canvasId, data, field, label) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const countObj = countBy(data, field);
  const labels = Object.keys(countObj);
  const values = Object.values(countObj);

  if (APP.charts[canvasId]) APP.charts[canvasId].destroy();
  APP.charts[canvasId] = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: CHART_COLORS.slice(0, labels.length),
        borderWidth: 2,
        borderColor: '#fff',
      }],
    },
    options: {
      responsive: true,
      plugins: { legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 10 } } },
    },
  });
}

/* ═══════════════════════════════════════════════════════════════
   TABLE
═══════════════════════════════════════════════════════════════ */
function renderTable(data) {
  const q = APP.tableSearch.toLowerCase();
  let rows = q
    ? data.filter(d => [d.id, d.municipality, d.post, d.suco, d.aldeia, d.status]
        .join(' ').toLowerCase().includes(q))
    : data;

  // Sort
  rows = [...rows].sort((a, b) => {
    let av = a[APP.sort.col] ?? a._rowIdx;
    let bv = b[APP.sort.col] ?? b._rowIdx;
    if (APP.sort.col === 'no') { av = a._rowIdx; bv = b._rowIdx; }
    if (typeof av === 'string') av = av.toLowerCase();
    if (typeof bv === 'string') bv = bv.toLowerCase();
    if (av < bv) return APP.sort.dir === 'asc' ? -1 : 1;
    if (av > bv) return APP.sort.dir === 'asc' ? 1 : -1;
    return 0;
  });

  const total = rows.length;
  const { page, perPage } = APP.pagination;
  const start = (page - 1) * perPage;
  const pageRows = rows.slice(start, start + perPage);

  const tbody = document.getElementById('table-body');
  if (!pageRows.length) {
    tbody.innerHTML = `<tr><td colspan="11" style="text-align:center;color:#94a3b8;padding:24px">Dadus la hetan.</td></tr>`;
  } else {
    tbody.innerHTML = pageRows.map((house, i) => {
      const imgPaths = house.imagePaths || [];
      const imgSrc = imgPaths.length ? escAttr(imgPaths[0]) : '';
      const thumb = imgSrc
        ? `<img src="${imgSrc}" class="thumb-img" alt="foto" loading="lazy"
             onerror="this.style.display='none';this.nextElementSibling.style.display='flex'"
             onclick="goToMarker('${escAttr(house.id)}')" />
           <div class="thumb-placeholder" style="display:none"><i class="fa-solid fa-image"></i></div>`
        : `<div class="thumb-placeholder"><i class="fa-solid fa-image"></i></div>`;

      const sid = escAttr(house.id);
      return `<tr>
        <td>${start + i + 1}</td>
        <td><strong>${sanitize(String(house.id).padStart(3,'0'))}</strong></td>
        <td title="${sanitize(house.municipality)}">${sanitize(house.municipality) || '—'}</td>
        <td title="${sanitize(house.post)}">${sanitize(house.post) || '—'}</td>
        <td title="${sanitize(house.suco)}">${sanitize(house.suco) || '—'}</td>
        <td title="${sanitize(house.aldeia)}">${sanitize(house.aldeia) || '—'}</td>
        <td><span class="badge ${statusClass(house.status)}">${sanitize(house.status) || 'N/D'}</span></td>
        <td>${house.latitude.toFixed(5)}</td>
        <td>${house.longitude.toFixed(5)}</td>
        <td>${thumb}</td>
        <td style="white-space:nowrap">
          <button class="tbl-btn view" onclick="openDetail('${sid}')"><i class="fa-solid fa-eye"></i></button>
          <button class="tbl-btn zoom" onclick="goToMarker('${sid}')"><i class="fa-solid fa-magnifying-glass-location"></i></button>
          <button class="tbl-btn nav"  onclick="navigateTo('${sid}')"><i class="fa-solid fa-route"></i></button>
        </td>
      </tr>`;
    }).join('');
  }

  renderPagination(total, page, perPage);
}

function renderPagination(total, page, perPage) {
  const bar = document.getElementById('pagination-bar');
  const totalPages = Math.ceil(total / perPage);
  if (totalPages <= 1) { bar.innerHTML = `<span class="page-info">${formatNum(total)} rekord</span>`; return; }

  let html = '';
  html += `<button class="page-btn" onclick="changePage(1)" ${page===1?'disabled':''}><i class="fa-solid fa-angles-left"></i></button>`;
  html += `<button class="page-btn" onclick="changePage(${page-1})" ${page===1?'disabled':''}><i class="fa-solid fa-angle-left"></i></button>`;

  const range = 2;
  for (let p = Math.max(1, page - range); p <= Math.min(totalPages, page + range); p++) {
    html += `<button class="page-btn ${p===page?'active':''}" onclick="changePage(${p})">${p}</button>`;
  }

  html += `<button class="page-btn" onclick="changePage(${page+1})" ${page===totalPages?'disabled':''}><i class="fa-solid fa-angle-right"></i></button>`;
  html += `<button class="page-btn" onclick="changePage(${totalPages})" ${page===totalPages?'disabled':''}><i class="fa-solid fa-angles-right"></i></button>`;
  html += `<span class="page-info">${formatNum(total)} rekord · Halaman ${page}/${totalPages}</span>`;
  bar.innerHTML = html;
}

window.changePage = function(p) {
  APP.pagination.page = p;
  renderTable(APP.filtered);
};

/* ═══════════════════════════════════════════════════════════════
   GALERIA
═══════════════════════════════════════════════════════════════ */
function renderGaleria(data) {
  const grid = document.getElementById('galeria-grid');
  if (!data.length) {
    grid.innerHTML = `<p style="color:#94a3b8;padding:16px">Dadus la hetan.</p>`;
    return;
  }

  grid.innerHTML = data.map(house => {
    const imgPaths = house.imagePaths || [];
    const imgSrc = imgPaths.length ? escAttr(imgPaths[0]) : '';
    const sid = escAttr(house.id);
    const imgHtml = imgSrc
      ? `<img class="galeria-img" src="${imgSrc}" alt="Uma ${sanitize(house.id)}" loading="lazy"
           onerror="this.style.display='none';this.nextElementSibling.style.display='flex'" />
         <div class="galeria-img-placeholder" style="display:none"><i class="fa-solid fa-image"></i></div>`
      : `<div class="galeria-img-placeholder"><i class="fa-solid fa-image"></i></div>`;

    return `<div class="galeria-item" onclick="openDetail('${sid}')">
      ${imgHtml}
      <div class="galeria-caption">
        UMA-${sanitize(String(house.id).padStart(3,'0'))}
        <span>${sanitize(house.municipality || house.suco || 'N/D')}</span>
      </div>
    </div>`;
  }).join('');
}

/* ═══════════════════════════════════════════════════════════════
   DETAIL MODAL
═══════════════════════════════════════════════════════════════ */
window.openDetail = function(id) {
  const house = APP.allData.find(d => String(d.id) === String(id));
  if (!house) return;

  const imgPaths = house.imagePaths || [];
  const mainSrc = imgPaths.length ? escAttr(imgPaths[0]) : '';
  const sid = escAttr(id);

  const photoHtml = mainSrc
    ? `<img id="modal-main-photo" class="modal-photo" src="${mainSrc}" alt="Foto Uma"
         onerror="this.style.display='none';document.getElementById('modal-photo-placeholder').style.display='flex'" />
       <div id="modal-photo-placeholder" class="modal-photo-placeholder" style="display:none">
         <i class="fa-solid fa-image"></i><span>Foto la disponível</span>
       </div>`
    : `<div id="modal-photo-placeholder" class="modal-photo-placeholder">
         <i class="fa-solid fa-image"></i><span>Foto la disponível</span>
       </div>`;

  const fields = [
    { label: 'ID Uma',            val: `UMA-${String(house.id).padStart(3,'0')}` },
    { label: 'Munisípiu',         val: house.municipality },
    { label: 'Postu Administrativu', val: house.post },
    { label: 'Suco',              val: house.suco },
    { label: 'Aldeia',            val: house.aldeia },
    { label: 'Status',            val: house.status },
    { label: 'Programa',          val: house.program },
    { label: 'Tinan',             val: house.year },
    { label: 'Latitude',          val: house.latitude.toFixed(6) },
    { label: 'Longitude',         val: house.longitude.toFixed(6) },
  ];

  const infoHtml = fields.map(f => f.val
    ? `<div class="modal-info-item">
        <label>${sanitize(f.label)}</label>
        <span>${sanitize(String(f.val))}</span>
       </div>` : ''
  ).join('');

  const gmapsUrl = `https://www.google.com/maps/dir/?api=1&destination=${house.latitude},${house.longitude}`;

  document.getElementById('modal-body-content').innerHTML = `
    ${photoHtml}
    <div class="modal-title"><i class="fa-solid fa-house"></i> UMA-${sanitize(String(house.id).padStart(3,'0'))}</div>
    <div class="modal-info-grid">${infoHtml}</div>
    <div class="modal-actions">
      <button class="btn-modal zoom"  onclick="goToMarker('${sid}');closeModal()"><i class="fa-solid fa-magnifying-glass-location"></i> Zoom Mapa</button>
      <button class="btn-modal green" onclick="navigateTo('${sid}');closeModal()"><i class="fa-solid fa-route"></i> Nabegasaun</button>
      <a class="btn-modal gmaps" href="${gmapsUrl}" target="_blank" rel="noopener noreferrer"><i class="fa-brands fa-google"></i> Google Maps</a>
    </div>`;

  document.getElementById('modal-detail').classList.remove('hidden');
};

window.closeModal = function() {
  document.getElementById('modal-detail').classList.add('hidden');
};

/* ═══════════════════════════════════════════════════════════════
   MAP ACTIONS
═══════════════════════════════════════════════════════════════ */
window.goToMarker = function(id) {
  // Switch to dashboard view which has the map
  switchView('dashboard');
  const house = APP.allData.find(d => String(d.id) === String(id));
  if (!house) return;
  setTimeout(() => {
    APP.map.setView([house.latitude, house.longitude], 16);
    const marker = APP.markers[house.id];
    if (marker) {
      APP.clusterGroup.zoomToShowLayer(marker, () => {
        marker.openPopup();
      });
    }
  }, 100);
};

window.navigateTo = function(id) {
  const house = APP.allData.find(d => String(d.id) === String(id));
  if (!house) return;
  switchView('dashboard');

  if (!navigator.geolocation) {
    fallbackNavigate(house);
    return;
  }
  showLoading('Hetan lokasaun ita-nia...');
  navigator.geolocation.getCurrentPosition(
    pos => {
      hideLoading();
      startRoute(pos.coords.latitude, pos.coords.longitude, house);
    },
    err => {
      hideLoading();
      showNotif('Lokasaun la bele asesu. Abre Google Maps...', 'warning', 4000);
      fallbackNavigate(house);
    },
    { timeout: 8000, enableHighAccuracy: true }
  );
};

function fallbackNavigate(house) {
  const url = `https://www.google.com/maps/dir/?api=1&destination=${house.latitude},${house.longitude}`;
  window.open(url, '_blank');
}

function startRoute(fromLat, fromLng, house) {
  // Clear existing route
  clearRoute();

  // Show user location marker
  if (APP.userMarker) APP.map.removeLayer(APP.userMarker);
  APP.userMarker = L.marker([fromLat, fromLng], {
    icon: L.divIcon({
      className: '',
      html: `<div class="user-location-marker"></div>`,
      iconSize: [20, 20], iconAnchor: [10, 10],
    }),
    zIndexOffset: 1000,
  }).addTo(APP.map).bindPopup('📍 Lokasaun Ita-Nia').openPopup();

  APP.routeControl = L.Routing.control({
    waypoints: [
      L.latLng(fromLat, fromLng),
      L.latLng(house.latitude, house.longitude),
    ],
    routeWhileDragging: false,
    showAlternatives: false,
    fitSelectedRoutes: true,
    lineOptions: {
      styles: [{ color: '#3b82f6', weight: 4, opacity: 0.8 }],
    },
    createMarker: function(i, wp) {
      if (i === 1) return L.marker(wp.latLng, { icon: buildMarkerIcon(house.status) });
      return null; // user marker already added
    },
    router: L.Routing.osrmv1({
      serviceUrl: 'https://router.project-osrm.org/route/v1',
    }),
  }).addTo(APP.map);

  document.getElementById('btn-clear-route').classList.remove('hidden');
  showNotif('🧭 Rota hahú husi lokasaun ita-nia ba uma ne\'ebé hili.', 'info', 5000);
}

function clearRoute() {
  if (APP.routeControl) {
    try { APP.map.removeControl(APP.routeControl); } catch(e) {}
    APP.routeControl = null;
  }
  if (APP.userMarker) {
    APP.map.removeLayer(APP.userMarker);
    APP.userMarker = null;
  }
  document.getElementById('btn-clear-route').classList.add('hidden');
}

/* ═══════════════════════════════════════════════════════════════
   SEARCH (debounced)
═══════════════════════════════════════════════════════════════ */
let searchTimer = null;
function onSearch(e) {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    const q = e.target.value.trim().toLowerCase();
    if (q.length >= 1) {
      // Try exact ID match first → zoom
      const exact = APP.allData.find(d => String(d.id).toLowerCase() === q ||
        `uma-${String(d.id).padStart(3,'0')}`.toLowerCase() === q);
      if (exact) {
        APP.filtered = [exact];
      } else {
        applyFilters();
        return;
      }
      // zoom to found
      switchView('dashboard');
      setTimeout(() => {
        APP.map.setView([exact.latitude, exact.longitude], 16);
        const m = APP.markers[exact.id];
        if (m) { APP.clusterGroup.zoomToShowLayer(m, () => m.openPopup()); }
      }, 150);
      updateKPI(APP.filtered);
      renderMarkers(APP.filtered);
      renderTable(APP.filtered);
    } else {
      applyFilters();
    }
  }, 320);
}

/* ═══════════════════════════════════════════════════════════════
   VIEW SWITCHING
═══════════════════════════════════════════════════════════════ */
function switchView(viewName) {
  APP.currentView = viewName;
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));

  const target = document.getElementById(`view-${viewName}`);
  if (target) target.classList.add('active');

  const navItem = document.querySelector(`.nav-item[data-view="${viewName}"]`);
  if (navItem) navItem.classList.add('active');

  // On mobile: close sidebar after navigation
  if (window.innerWidth < 768) {
    document.getElementById('sidebar').classList.remove('mobile-open');
  }

  // Invalidate map size after view switch
  if (viewName === 'dashboard') setTimeout(() => APP.map && APP.map.invalidateSize(), 120);
}

/* ═══════════════════════════════════════════════════════════════
   SIDEBAR TOGGLE
═══════════════════════════════════════════════════════════════ */
function initSidebar() {
  const sidebar = document.getElementById('sidebar');
  const wrapper = document.getElementById('main-wrapper');
  const toggle  = document.getElementById('sidebar-toggle');

  toggle.addEventListener('click', () => {
    if (window.innerWidth < 768) {
      sidebar.classList.toggle('mobile-open');
    } else {
      sidebar.classList.toggle('collapsed');
      wrapper.classList.toggle('expanded');
      setTimeout(() => APP.map && APP.map.invalidateSize(), 260);
    }
  });
}

/* ═══════════════════════════════════════════════════════════════
   MY LOCATION
═══════════════════════════════════════════════════════════════ */
function initMyLocation() {
  document.getElementById('btn-my-location').addEventListener('click', () => {
    if (!navigator.geolocation) {
      showNotif('Browser la suporta geolokasaun.', 'error');
      return;
    }
    showLoading('Hetan lokasaun ita-nia...');
    navigator.geolocation.getCurrentPosition(
      pos => {
        hideLoading();
        const { latitude: lat, longitude: lng } = pos.coords;
        switchView('dashboard');
        setTimeout(() => {
          APP.map.setView([lat, lng], 14);
          if (APP.userMarker) APP.map.removeLayer(APP.userMarker);
          APP.userMarker = L.marker([lat, lng], {
            icon: L.divIcon({
              className: '',
              html: `<div class="user-location-marker"></div>`,
              iconSize: [20, 20], iconAnchor: [10, 10],
            }),
          }).addTo(APP.map).bindPopup('📍 Lokasaun Ita-Nia').openPopup();
          showNotif('📍 Lokasaun ita-nia hatudu ona iha mapa.', 'success');
        }, 150);
      },
      err => {
        hideLoading();
        showNotif('Lokasaun pengguna la bele asesu. Aktiva GPS/permission lokasaun iha browser.', 'warning', 6000);
      },
      { timeout: 8000, enableHighAccuracy: true }
    );
  });
}

/* ═══════════════════════════════════════════════════════════════
   IMPORT EXCEL
═══════════════════════════════════════════════════════════════ */
function initImport() {
  document.getElementById('excel-input').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx','xls','csv'].includes(ext)) {
      showNotif('❌ Formatu ficheiru la válidu. Uza XLSX, XLS, ka CSV.', 'error');
      e.target.value = '';
      return;
    }

    showLoading('Lee ficheiru Excel...');
    try {
      let rows;
      if (ext === 'csv') {
        rows = await readCSV(file);
      } else {
        rows = await readExcel(file);
      }
      showLoading(`Prosesa ${rows.length} liña dadus...`);
      await sleep(60);
      await loadData(rows);
    } catch (err) {
      hideLoading();
      showNotif('❌ Ficheiru la bele lee. Verifika formatu XLSX/XLS/CSV.', 'error');
      console.error(err);
    }
    e.target.value = '';
  });
}

function readCSV(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'string' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        resolve(XLSX.utils.sheet_to_json(ws, { defval: '' }));
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsText(file);
  });
}

/* ═══════════════════════════════════════════════════════════════
   TABLE SORT
═══════════════════════════════════════════════════════════════ */
function initTableSort() {
  document.querySelectorAll('#data-table th[data-col]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      if (APP.sort.col === col) {
        APP.sort.dir = APP.sort.dir === 'asc' ? 'desc' : 'asc';
      } else {
        APP.sort.col = col;
        APP.sort.dir = 'asc';
      }
      document.querySelectorAll('#data-table th').forEach(t => t.classList.remove('sort-asc','sort-desc'));
      th.classList.add(`sort-${APP.sort.dir}`);
      APP.pagination.page = 1;
      renderTable(APP.filtered);
    });
  });
}

/* ═══════════════════════════════════════════════════════════════
   TABLE SEARCH (separate from map search)
═══════════════════════════════════════════════════════════════ */
let tableSearchTimer = null;
function initTableSearch() {
  document.getElementById('table-search').addEventListener('input', (e) => {
    clearTimeout(tableSearchTimer);
    tableSearchTimer = setTimeout(() => {
      APP.tableSearch = e.target.value.trim();
      APP.pagination.page = 1;
      renderTable(APP.filtered);
    }, 280);
  });
}

/* ═══════════════════════════════════════════════════════════════
   INIT — BOOT SEQUENCE
═══════════════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', async () => {
  // Init map
  initMap();

  // Init UI controls
  initSidebar();
  initMyLocation();
  initImport();
  initTableSort();
  initTableSearch();

  // Nav items
  document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      const view = item.dataset.view;
      // "mapa" redirects to dashboard (same map)
      if (view === 'mapa') {
        switchView('dashboard');
        setTimeout(() => {
          document.querySelector('.map-card').scrollIntoView({ behavior: 'smooth', block: 'start' });
          APP.map && APP.map.invalidateSize();
        }, 150);
      } else {
        switchView(view);
      }
    });
  });

  // Filter listeners
  ['filter-municipality','filter-post','filter-suco','filter-aldeia',
   'filter-status','filter-program','filter-year'].forEach(id => {
    document.getElementById(id).addEventListener('change', applyFilters);
  });

  // Search input
  document.getElementById('search-input').addEventListener('input', onSearch);

  // Reset filter
  document.getElementById('btn-reset-filter').addEventListener('click', () => {
    ['filter-municipality','filter-post','filter-suco','filter-aldeia',
     'filter-status','filter-program','filter-year'].forEach(id => {
      document.getElementById(id).value = '';
    });
    document.getElementById('search-input').value = '';
    APP.filtered = [...APP.allData];
    APP.pagination.page = 1;
    renderMarkers(APP.filtered);
    updateKPI(APP.filtered);
    renderCharts(APP.filtered);
    renderTable(APP.filtered);
    renderGaleria(APP.filtered);
  });

  // Clear route button
  document.getElementById('btn-clear-route').addEventListener('click', clearRoute);

  // Modal close
  document.getElementById('modal-close-btn').addEventListener('click', closeModal);
  document.getElementById('modal-detail').addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-detail')) closeModal();
  });

  // ESC key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeModal();
  });

  // Load default Excel
  showLoading('Karga dadus Excel default...');
  const rows = await readDefaultExcel();
  if (rows) {
    await loadData(rows);
  } else {
    hideLoading();
    showNotif('ℹ Klik "Import Excel" hodi karga ficheiru Coordinate_Uma.xlsx.', 'info', 10000);
    updateKPI([]);
    renderCharts([]);
    renderTable([]);
    renderGaleria([]);
  }
});
