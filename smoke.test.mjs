import { readFileSync } from 'fs';
import { JSDOM } from 'jsdom';

// Load HTML and strip external resources
const rawHtml = readFileSync('./index.html', 'utf-8');
const sanitizedHtml = rawHtml
  .replace(/<link[^>]*>/gi, '')
  .replace(new RegExp('<script[^>]*src=[^>]*><\\/script>', 'gi'), '');
const dom = new JSDOM(sanitizedHtml, {
  url: 'http://localhost/',
  runScripts: 'dangerously',
  resources: 'usable',
  pretendToBeVisual: true,
});

// Stub fetch of external script tags (we'll inject globals manually)
global.window = dom.window;
global.document = dom.window.document;
global.Blob = dom.window.Blob;
global.URL = dom.window.URL;
dom.window.alert = () => {};
dom.window.confirm = () => true;

// Provide Papa and XLSX in global scope expected by script.js
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
global.Papa = Papa;
global.XLSX = XLSX;

// Load our app script
const scriptText = readFileSync('./script.js', 'utf-8');
const scriptEl = document.createElement('script');
scriptEl.textContent = scriptText;
document.body.appendChild(scriptEl);

// Fire DOMContentLoaded to trigger app init
document.dispatchEvent(new dom.window.Event('DOMContentLoaded', { bubbles: true }));
await new Promise(r => setTimeout(r, 150));

// 1) Add device through UI
const addBtn = document.querySelector('#addDeviceBtn');
// stub prompts
let promptCalls = 0;
dom.window.prompt = (msg, defVal) => {
  promptCalls++;
  if (promptCalls === 1) return '10.10.10.1';
  if (promptCalls === 2) return 'router-1010';
  if (promptCalls === 3) return 'router';
  if (promptCalls === 4) return 'IOS 15.2';
  return defVal || '';
};
addBtn.click();
await new Promise(r => setTimeout(r, 100));

// 2) Generate fake data
document.querySelector('#genFakeBtn').click();
await new Promise(r => setTimeout(r, 100));

// 3) Export JSON
document.querySelector('#exportJsonBtn').click();
await new Promise(r => setTimeout(r, 50));
const preview = document.querySelector('#exportPreview').textContent || '';

// 4) Basic assertions
const rows = document.querySelectorAll('#deviceTable tbody tr').length;
if (rows < 2) {
  console.error('Expected at least 2 devices, got', rows);
  process.exit(1);
}
if (!preview.includes('devices')) {
  console.error('Export preview missing devices JSON');
  process.exit(1);
}

console.log('Smoke test passed. Devices:', rows);
