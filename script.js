/* State */
const state = {
  devices: [],
  links: [],
  configs: [],
  iconTheme: 'flat',
  layout: 'force'
};

/* Helpers */
const $ = (sel, ctx = document) => ctx.querySelector(sel);
const $$ = (sel, ctx = document) => Array.from(ctx.querySelectorAll(sel));

function saveState() {
  localStorage.setItem('starweb_state', JSON.stringify({
    devices: state.devices,
    links: state.links,
    configs: state.configs
  }));
}

function loadState() {
  try {
    const raw = localStorage.getItem('starweb_state');
    if (!raw) return;
    const parsed = JSON.parse(raw);
    state.devices = parsed.devices || [];
    state.links = parsed.links || [];
    state.configs = parsed.configs || [];
  } catch (e) {
    console.warn('Failed to load state', e);
  }
}

function nowIsoDate() {
  return new Date().toISOString();
}

function upsertDevice(device) {
  const idx = state.devices.findIndex(d => d.ip === device.ip || (device.hostname && d.hostname === device.hostname));
  if (idx >= 0) {
    state.devices[idx] = { ...state.devices[idx], ...device };
  } else {
    state.devices.push(device);
  }
}

function removeDeviceByIp(ip) {
  state.devices = state.devices.filter(d => d.ip !== ip);
}

function renderTable() {
  const tbody = $('#deviceTable tbody');
  tbody.innerHTML = '';
  for (const d of state.devices) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${d.ip || ''}</td>
      <td>${d.hostname || ''}</td>
      <td>${d.type || ''}</td>
      <td>${d.os || ''}</td>
      <td>${d.date || ''}</td>
      <td>
        <button data-action="edit" data-ip="${d.ip}">Edit</button>
        <button data-action="delete" data-ip="${d.ip}" class="danger">Delete</button>
      </td>
    `;
    tbody.appendChild(tr);
  }
}

function promptDevice(existing) {
  const ip = prompt('IP address', existing?.ip || '');
  if (!ip) return null;
  const hostname = prompt('Hostname (optional)', existing?.hostname || '');
  const type = prompt('Device type (router/switch/server/etc)', existing?.type || '');
  const os = prompt('OS version', existing?.os || '');
  const date = existing?.date || nowIsoDate();
  return { ip, hostname, type, os, date };
}

/* Fake Data */
const DEVICE_TYPES = ['router', 'switch', 'firewall', 'server', 'workstation', 'ap'];
const OS_VERSIONS = ['IOS 15.2', 'NX-OS 10.2', 'PAN-OS 11', 'Ubuntu 22.04', 'Windows 11', 'ESXi 8'];

function randomIp() {
  return `10.${Math.floor(Math.random()*256)}.${Math.floor(Math.random()*256)}.${Math.floor(Math.random()*256)}`;
}

function randomChoice(arr) { return arr[Math.floor(Math.random()*arr.length)]; }

function generateFakeData(count = 20) {
  const newDevices = [];
  for (let i = 0; i < count; i++) {
    const ip = randomIp();
    const type = randomChoice(DEVICE_TYPES);
    const os = randomChoice(OS_VERSIONS);
    const hostname = `${type}-${ip.split('.').slice(-2).join('')}`;
    newDevices.push({ ip, hostname, type, os, date: nowIsoDate() });
  }
  // simple links between neighbors
  const newLinks = [];
  for (let i = 1; i < newDevices.length; i++) {
    newLinks.push({ from_ip: newDevices[i-1].ip, to_ip: newDevices[i].ip, protocol: 'ethernet' });
  }
  for (const d of newDevices) upsertDevice(d);
  state.links.push(...newLinks);
}

/* CSV / XLSX Ingestion */
function parseCsvText(text, filename = 'data.csv') {
  return new Promise(resolve => {
    Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      complete: (res) => resolve({ data: res.data, meta: res.meta, filename })
    });
  });
}

async function parseSpreadsheet(file) {
  const buf = await file.arrayBuffer();
  const workbook = XLSX.read(buf);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  return { data: json, filename: file.name };
}

function buildMappingUI(columns) {
  const mapping = $('#mapping');
  mapping.innerHTML = '';
  const required = ['from_ip','from_port','from_mac','to_ip','to_port','to_mac','protocol','hostname'];
  for (const field of required) {
    const wrap = document.createElement('div');
    const label = document.createElement('label');
    label.textContent = field;
    const select = document.createElement('select');
    const none = document.createElement('option');
    none.value = ''; none.textContent = '(none)';
    select.appendChild(none);
    for (const col of columns) {
      const opt = document.createElement('option');
      opt.value = col; opt.textContent = col;
      const colLower = col.toLowerCase();
      const fieldLower = field.toLowerCase();
      if (colLower === fieldLower) opt.selected = true;
      // alias: handle common typo form_mac -> from_mac
      if (fieldLower === 'from_mac' && colLower === 'form_mac') opt.selected = true;
      select.appendChild(opt);
    }
    wrap.appendChild(label);
    wrap.appendChild(select);
    mapping.appendChild(wrap);
  }
  const applyBtn = document.createElement('button');
  applyBtn.textContent = 'Apply Mapping & Ingest';
  applyBtn.addEventListener('click', () => applyMapping());
  mapping.appendChild(applyBtn);
}

function applyMapping() {
  const selects = $$('#mapping select');
  const fields = ['from_ip','from_port','from_mac','to_ip','to_port','to_mac','protocol','hostname'];
  const mapping = {};
  fields.forEach((f, i) => mapping[f] = selects[i].value);
  const data = window.__lastParsedRows || [];
  const links = [];
  for (const row of data) {
    const link = {
      from_ip: mapping.from_ip ? row[mapping.from_ip] : '',
      from_port: mapping.from_port ? row[mapping.from_port] : '',
      from_mac: mapping.from_mac ? row[mapping.from_mac] : (row['form_mac'] || ''),
      to_ip: mapping.to_ip ? row[mapping.to_ip] : '',
      to_port: mapping.to_port ? row[mapping.to_port] : '',
      to_mac: mapping.to_mac ? row[mapping.to_mac] : '',
      protocol: mapping.protocol ? row[mapping.protocol] : '',
      hostname: mapping.hostname ? row[mapping.hostname] : ''
    };
    if (link.from_ip && link.to_ip) {
      links.push(link);
      // create placeholder devices
      upsertDevice({ ip: link.from_ip, hostname: link.hostname || '', type: '', os: '', date: nowIsoDate() });
      upsertDevice({ ip: link.to_ip, hostname: '', type: '', os: '', date: nowIsoDate() });
      // persist the original row as a source line for export
      try { state.configs.push(JSON.stringify(row)); } catch (e) {}
    }
  }
  state.links.push(...links);
  saveState();
  renderTable();
  renderViz();
}

/* Config parsing (very forgiving) */
function parseConfigText(text) {
  // Extract ip, hostname, device type, os heuristically
  const lines = text.split(/\n+/).map(l => l.trim()).filter(Boolean);
  const added = [];
  for (const line of lines) {
    const ipMatch = line.match(/(\b\d{1,3}(?:\.\d{1,3}){3}\b)/);
    const hostMatch = line.match(/host(?:name)?\s*[:=\-]?\s*([A-Za-z0-9_.\-]+)/i);
    const typeMatch = line.match(/(router|switch|firewall|server|ap|workstation)/i);
    const osMatch = line.match(/(IOS XE|IOS|NX-OS|PAN-OS|JunOS|Windows\s?\d+|Ubuntu\s?\d+\.\d+|ESXi\s?\d+)/i);
    if (ipMatch) {
      const device = {
        ip: ipMatch[1],
        hostname: hostMatch ? hostMatch[1] : '',
        type: typeMatch ? typeMatch[1].toLowerCase() : '',
        os: osMatch ? osMatch[1] : '',
        date: nowIsoDate()
      };
      upsertDevice(device);
      added.push(device);
      state.configs.push(line);
    }
  }
  return added;
}

/* Visualization */
function positionNodes(layout, nodes, edges, canvas) {
  const width = canvas.clientWidth;
  const height = canvas.clientHeight;
  const positions = new Map();
  if (layout === 'grid') {
    const cols = Math.ceil(Math.sqrt(nodes.length));
    const cellW = width / (cols + 1);
    const cellH = height / (cols + 1);
    nodes.forEach((n, i) => {
      const r = Math.floor(i / cols) + 1;
      const c = (i % cols) + 1;
      positions.set(n.ip, { x: c * cellW, y: r * cellH });
    });
  } else if (layout === 'circular') {
    const r = Math.min(width, height) * 0.35;
    nodes.forEach((n, i) => {
      const t = (i / nodes.length) * Math.PI * 2;
      positions.set(n.ip, { x: width/2 + Math.cos(t)*r, y: height/2 + Math.sin(t)*r });
    });
  } else if (layout === 'hierarchical') {
    const levels = 4;
    const byType = new Map();
    nodes.forEach(n => {
      const lvl = ['router','firewall'].includes(n.type) ? 1 : n.type === 'switch' ? 2 : 3;
      if (!byType.has(lvl)) byType.set(lvl, []);
      byType.get(lvl).push(n);
    });
    for (let lvl = 1; lvl <= levels; lvl++) {
      const row = byType.get(lvl) || [];
      const y = (lvl / (levels+1)) * height;
      row.forEach((n, i) => {
        const x = ((i+1) / (row.length+1)) * width;
        positions.set(n.ip, { x, y });
      });
    }
  } else {
    // force-like but simple
    const center = { x: width/2, y: height/2 };
    nodes.forEach((n, i) => {
      const angle = (i / nodes.length) * Math.PI * 2;
      positions.set(n.ip, { x: center.x + Math.cos(angle)* (width*0.3), y: center.y + Math.sin(angle)*(height*0.3) });
    });
  }
  return positions;
}

function deviceIconChar(device) {
  const theme = state.iconTheme;
  const type = (device.type || '').toLowerCase();
  if (theme === 'emoji') {
    if (type === 'router') return '📡';
    if (type === 'switch') return '🔀';
    if (type === 'firewall') return '🧱';
    if (type === 'server') return '🗄️';
    if (type === 'ap') return '📶';
    return '💻';
  }
  if (theme === 'glyph') {
    if (type === 'router') return 'R';
    if (type === 'switch') return 'S';
    if (type === 'firewall') return 'F';
    if (type === 'server') return 'SRV';
    if (type === 'ap') return 'AP';
    return 'PC';
  }
  // flat theme uses no char, color only
  return '';
}

function renderViz() {
  const canvas = $('#vizCanvas');
  canvas.innerHTML = '';
  const nodes = state.devices;
  const edges = state.links;
  const positions = positionNodes(state.layout, nodes, edges, canvas);

  // draw edges
  for (const e of edges) {
    const a = positions.get(e.from_ip);
    const b = positions.get(e.to_ip);
    if (!a || !b) continue;
    const dx = b.x - a.x;
    const dy = b.y - a.y;
    const len = Math.hypot(dx, dy);
    const angle = Math.atan2(dy, dx) * 180 / Math.PI;
    const el = document.createElement('div');
    el.className = 'edge';
    el.style.width = `${len}px`;
    el.style.left = `${a.x}px`;
    el.style.top = `${a.y}px`;
    el.style.transform = `rotate(${angle}deg)`;
    canvas.appendChild(el);
  }

  // draw nodes
  for (const n of nodes) {
    const p = positions.get(n.ip);
    if (!p) continue;
    const el = document.createElement('div');
    el.className = `node ${state.iconTheme}`;
    el.style.left = `${p.x - 28}px`;
    el.style.top = `${p.y - 28}px`;
    el.title = `${n.hostname || n.ip} (${n.type || 'device'})\nOS: ${n.os || 'n/a'}\n${n.ip}`;
    el.textContent = deviceIconChar(n);
    const label = document.createElement('div');
    label.className = 'node-label';
    label.textContent = n.hostname || n.ip;
    el.appendChild(label);
    canvas.appendChild(el);
  }
}

/* Export */
function toCsv(rows) {
  if (!rows.length) return '';
  const headers = Object.keys(rows[0]);
  const escape = (v) => {
    if (v == null) return '';
    const s = String(v);
    if (s.includes(',') || s.includes('"') || s.includes('\n')) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };
  const lines = [headers.join(',')];
  for (const r of rows) {
    lines.push(headers.map(h => escape(r[h])).join(','));
  }
  return lines.join('\n');
}

function download(filename, content, type = 'text/plain') {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* Wire UI */
function bindEvents() {
  document.body.addEventListener('click', (e) => {
    const target = e.target;
    if (target.matches('button[data-action="edit"]')) {
      const ip = target.getAttribute('data-ip');
      const existing = state.devices.find(d => d.ip === ip);
      const d = promptDevice(existing);
      if (d) { upsertDevice(d); saveState(); renderTable(); renderViz(); }
    }
    if (target.matches('button[data-action="delete"]')) {
      const ip = target.getAttribute('data-ip');
      removeDeviceByIp(ip);
      saveState(); renderTable(); renderViz();
    }
  });

  $('#addDeviceBtn').addEventListener('click', () => {
    const d = promptDevice();
    if (d) { upsertDevice(d); saveState(); renderTable(); renderViz(); }
  });
  $('#genFakeBtn').addEventListener('click', () => { generateFakeData(25); saveState(); renderTable(); renderViz(); });
  $('#clearAllBtn').addEventListener('click', () => {
    if (confirm('Clear all inventory and links?')) {
      state.devices = []; state.links = []; state.configs = []; saveState(); renderTable(); renderViz();
    }
  });

  // hero scroll
  $$('button[data-scroll-target]').forEach(btn => btn.addEventListener('click', () => {
    const id = btn.getAttribute('data-scroll-target');
    document.getElementById(id)?.scrollIntoView({ behavior: 'smooth' });
  }));

  // ingest file
  $('#fileInput').addEventListener('change', async (ev) => {
    const file = ev.target.files?.[0];
    if (!file) return;
    if (file.type.includes('sheet') || file.name.endsWith('.xlsx')) {
      const { data } = await parseSpreadsheet(file);
      window.__lastParsedRows = data;
      buildMappingUI(Object.keys(data[0] || {}));
    } else {
      const text = await file.text();
      const { data } = await parseCsvText(text, file.name);
      window.__lastParsedRows = data;
      buildMappingUI(Object.keys(data[0] || {}));
    }
  });

  // parse config
  $('#parseConfigBtn').addEventListener('click', () => {
    const text = $('#configInput').value;
    const added = parseConfigText(text);
    saveState(); renderTable(); renderViz();
    const preview = $('#exportPreview');
    preview.textContent = `Added ${added.length} devices from configs.`;
  });

  // layout and icon theme
  $('#layoutSelect').addEventListener('change', (e) => { state.layout = e.target.value; renderViz(); });
  $('#iconThemeSelect').addEventListener('change', (e) => { state.iconTheme = e.target.value; renderViz(); });

  // exports
  $('#exportJsonBtn').addEventListener('click', () => {
    const payload = { devices: state.devices, links: state.links, generatedAt: nowIsoDate() };
    const text = JSON.stringify(payload, null, 2);
    $('#exportPreview').textContent = text;
    download('inventory.json', text, 'application/json');
  });
  $('#exportCsvBtn').addEventListener('click', () => {
    const csv = toCsv(state.devices.map(d => ({ ip: d.ip, hostname: d.hostname, type: d.type, os: d.os, date: d.date })));
    $('#exportPreview').textContent = csv;
    download('inventory.csv', csv, 'text/csv');
  });
  $('#exportConfigsBtn').addEventListener('click', () => {
    const lines = state.configs.join('\n');
    $('#exportPreview').textContent = lines;
    download('configs.txt', lines, 'text/plain');
  });
}

/* Init */
function init() {
  loadState();
  renderTable();
  bindEvents();
  renderViz();
}

document.addEventListener('DOMContentLoaded', init);

