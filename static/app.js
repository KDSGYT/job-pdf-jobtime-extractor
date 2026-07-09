const form = document.querySelector('#uploadForm');
const fileInput = document.querySelector('#pdfFile');
const fileName = document.querySelector('#fileName');
const statusBox = document.querySelector('#statusBox');
const summaryGrid = document.querySelector('#summaryGrid');
const tableWrap = document.querySelector('#tableWrap');
const resultsBody = document.querySelector('#resultsBody');
const filters = document.querySelector('#filters');
const searchInput = document.querySelector('#searchInput');
const typeFilter = document.querySelector('#typeFilter');
const downloadBtn = document.querySelector('#downloadBtn');
const extractBtn = document.querySelector('#extractBtn');

let currentRows = [];
let currentFile = null;
let sortColumn = '';
let sortDirection = 'asc';

const timeColumns = new Set(['start', 'end', 'duration', 'job_start', 'job_end', 'job_duration']);

fileInput.addEventListener('change', () => {
  currentFile = fileInput.files[0] || null;
  fileName.textContent = currentFile ? currentFile.name : 'No file selected';
  downloadBtn.disabled = !currentFile || currentRows.length === 0;
});

function setStatus(message, isError = false) {
  statusBox.textContent = message;
  statusBox.classList.toggle('error', isError);
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, ch => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
  }[ch]));
}

function badgeClass(type) {
  if (type === 'Split') return 'badge split';
  if (type === 'STBY') return 'badge stby';
  return 'badge';
}

function filteredRows() {
  const query = searchInput.value.trim().toLowerCase();
  const type = typeFilter.value;
  const words = query.split(/\s+/).filter(Boolean);
  return currentRows.filter(row => {
    if (type && row.record_type !== type) return false;
    if (!words.length) return true;
    const haystack = Object.values(row).map(value => String(value ?? '').toLowerCase()).join(' ');
    return words.every(word => haystack.includes(word));
  });
}

function timeToMinutes(value) {
  const [hours, minutes] = String(value || '').split(':');
  if (!/^\d+$/.test(hours || '') || !/^\d+$/.test(minutes || '')) return -1;
  return Number(hours) * 60 + Number(minutes);
}

function sortValue(row, column) {
  if (timeColumns.has(column)) return timeToMinutes(row[column]);
  if (column === 'page') return Number(row[column] || 0);
  return String(row[column] ?? '').toLowerCase();
}

function sortedRows(rows) {
  if (!sortColumn) return rows;
  const direction = sortDirection === 'desc' ? -1 : 1;
  return [...rows].sort((a, b) => {
    const av = sortValue(a, sortColumn);
    const bv = sortValue(b, sortColumn);
    if (av < bv) return -1 * direction;
    if (av > bv) return 1 * direction;
    return 0;
  });
}

function renderSortHeaders() {
  document.querySelectorAll('th[data-sort]').forEach(th => {
    th.classList.remove('sorted-asc', 'sorted-desc');
    if (th.dataset.sort === sortColumn) th.classList.add(`sorted-${sortDirection}`);
  });
}

function renderRows() {
  const rows = sortedRows(filteredRows());
  resultsBody.innerHTML = rows.map(row => `
    <tr>
      <td><strong>${escapeHtml(row.job_number)}</strong></td>
      <td>${escapeHtml(row.valid_days)}</td>
      <td><span class="${badgeClass(row.record_type)}">${escapeHtml(row.record_type)}</span></td>
      <td>${escapeHtml(row.start)}</td>
      <td>${escapeHtml(row.end)}</td>
      <td><strong>${escapeHtml(row.duration)}</strong></td>
      <td>${escapeHtml(row.from_location)}</td>
      <td>${escapeHtml(row.to_location)}</td>
      <td>${escapeHtml(row.job_start)} → ${escapeHtml(row.job_end)}</td>
      <td>${escapeHtml(row.on_duty_location)}</td>
      <td>${escapeHtml(row.page)}</td>
    </tr>
  `).join('');
  renderSortHeaders();
  setStatus(`Showing ${rows.length} of ${currentRows.length} extracted records.`);
}

function renderSummary(counts) {
  document.querySelector('#jobsCount').textContent = counts.jobs;
  document.querySelector('#recordsCount').textContent = counts.records;
  document.querySelector('#pdCount').textContent = counts.pd_time_records;
  document.querySelector('#stbyCount').textContent = counts.stby_records;
  document.querySelector('#splitCount').textContent = counts.split_records;
  summaryGrid.hidden = false;
  filters.hidden = false;
  tableWrap.hidden = false;
}

async function postPdf(url) {
  const data = new FormData();
  data.append('pdf', currentFile);
  return fetch(url, { method: 'POST', body: data });
}

form.addEventListener('submit', async event => {
  event.preventDefault();
  if (!currentFile) return setStatus('Please choose a PDF first.', true);
  extractBtn.disabled = true;
  downloadBtn.disabled = true;
  setStatus('Reading PDF and extracting records...');
  try {
    const response = await postPdf('/api/extract');
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || 'Extraction failed.');
    currentRows = payload.rows || [];
    renderSummary(payload.counts);
    renderRows();
    downloadBtn.disabled = currentRows.length === 0;
  } catch (error) {
    currentRows = [];
    resultsBody.innerHTML = '';
    setStatus(error.message, true);
  } finally {
    extractBtn.disabled = false;
  }
});

async function downloadCsv() {
  if (!currentFile) return;
  setStatus('Preparing CSV download...');
  const response = await postPdf('/api/extract.csv');
  if (!response.ok) {
    setStatus(await response.text(), true);
    return;
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const base = currentFile.name.replace(/\.pdf$/i, '').replace(/\s+/g, '_') || 'job_times';
  a.href = url;
  a.download = `${base}_pd_stby_split.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  setStatus('CSV downloaded.');
}

downloadBtn.addEventListener('click', downloadCsv);
searchInput.addEventListener('input', renderRows);
typeFilter.addEventListener('change', renderRows);
document.querySelectorAll('th[data-sort]').forEach(th => {
  th.addEventListener('click', () => {
    const nextColumn = th.dataset.sort;
    sortDirection = sortColumn === nextColumn && sortDirection === 'asc' ? 'desc' : 'asc';
    sortColumn = nextColumn;
    renderRows();
  });
});
