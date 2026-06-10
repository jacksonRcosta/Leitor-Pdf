const FUNCTION_URL = '/.netlify/functions/converter';
const MAX_SIZE_MB  = 4;

// ── Elementos ─────────────────────────────────────────────────────────────────
const dropZone       = document.getElementById('drop-zone');
const fileInput      = document.getElementById('file-input');
const fileInfo       = document.getElementById('file-info');
const fileName       = document.getElementById('file-name');
const fileSize       = document.getElementById('file-size');
const removeFile     = document.getElementById('remove-file');
const convertBtn     = document.getElementById('convert-btn');
const uploadSection  = document.getElementById('upload-section');
const progressSection= document.getElementById('progress-section');
const progressMsg    = document.getElementById('progress-msg');
const resultSection  = document.getElementById('result-section');
const resultTbody    = document.getElementById('result-tbody');
const downloadAllBtn = document.getElementById('download-all-btn');
const newConvBtn     = document.getElementById('new-conversion-btn');
const errorSection   = document.getElementById('error-section');
const errorMsg       = document.getElementById('error-msg');
const retryBtn       = document.getElementById('retry-btn');

let selectedFile  = null;
let zipBase64     = null;

// ── Upload ────────────────────────────────────────────────────────────────────
dropZone.addEventListener('click',    () => fileInput.click());
dropZone.addEventListener('keydown',  e => e.key === 'Enter' && fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave',() => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) handleFile(f);
});
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});
removeFile.addEventListener('click', resetUpload);

function handleFile(file) {
  if (!file.name.toLowerCase().endsWith('.pdf')) {
    showError('O arquivo selecionado não é um PDF.');
    return;
  }
  if (file.size > MAX_SIZE_MB * 1024 * 1024) {
    showError(`O arquivo excede o limite de ${MAX_SIZE_MB} MB.`);
    return;
  }
  selectedFile = file;
  fileName.textContent = file.name;
  fileSize.textContent = formatSize(file.size);
  dropZone.classList.add('hidden');
  fileInfo.classList.remove('hidden');
  convertBtn.disabled = false;
}

function formatSize(bytes) {
  return bytes < 1024 * 1024
    ? `${(bytes / 1024).toFixed(1)} KB`
    : `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function resetUpload() {
  selectedFile = null;
  zipBase64    = null;
  fileInput.value = '';
  dropZone.classList.remove('hidden');
  fileInfo.classList.add('hidden');
  convertBtn.disabled = true;
}

// ── Conversão ─────────────────────────────────────────────────────────────────
convertBtn.addEventListener('click', async () => {
  if (!selectedFile) return;
  showSection('progress');
  progressMsg.textContent = 'Lendo PDF...';

  try {
    const b64 = await fileToBase64(selectedFile);
    progressMsg.textContent = 'Processando pedidos...';

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 55000);

    let res;
    try {
      res = await fetch(FUNCTION_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ pdf: b64 }),
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timeout);
    }

    const text = await res.text();

    if (!text || !text.trim()) {
      throw new Error(
        res.status === 413
          ? 'Arquivo muito grande. O limite é 4 MB.'
          : res.status === 504 || res.status === 524
          ? 'Tempo de processamento excedido. Tente com um PDF menor.'
          : `Sem resposta do servidor (HTTP ${res.status}). Verifique o deploy da função.`
      );
    }

    let data;
    try {
      data = JSON.parse(text);
    } catch {
      console.error('Resposta não-JSON:', text.slice(0, 300));
      throw new Error('Resposta inválida do servidor. Verifique os logs da função no Netlify.');
    }

    if (!res.ok) {
      throw new Error(data.erro || `Erro ${res.status}`);
    }

    zipBase64 = data.zip;
    renderResult(data.pedidos);
    showSection('result');

  } catch (err) {
    if (err.name === 'AbortError') {
      showError('Tempo limite atingido (55s). O PDF pode ser muito complexo ou grande.');
    } else {
      showError(err.message || 'Erro inesperado. Tente novamente.');
    }
  }
});

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = () => resolve(reader.result.split(',')[1]);
    reader.onerror = () => reject(new Error('Falha ao ler o arquivo.'));
    reader.readAsDataURL(file);
  });
}

// ── Resultado ─────────────────────────────────────────────────────────────────
function renderResult(pedidos) {
  resultTbody.innerHTML = '';
  pedidos.forEach(p => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="badge">${esc(p.loja)}</span></td>
      <td>${esc(p.razao_social)}</td>
      <td>${esc(p.num_pedido) || '—'}</td>
      <td>${esc(p.data_compra) || '—'}</td>
      <td>${esc(p.data_entrega) || '—'}</td>
      <td>${p.total_itens}</td>
    `;
    resultTbody.appendChild(tr);
  });
}

downloadAllBtn.addEventListener('click', () => {
  if (!zipBase64) return;
  const name = selectedFile
    ? selectedFile.name.replace(/\.pdf$/i, '') + '_convertido.zip'
    : 'pedidos_convertidos.zip';
  downloadBase64(zipBase64, 'application/zip', name);
});

newConvBtn.addEventListener('click', () => {
  resetUpload();
  resultTbody.innerHTML = '';
  showSection('upload');
});

retryBtn.addEventListener('click', () => showSection('upload'));

// ── Helpers ───────────────────────────────────────────────────────────────────
function showSection(name) {
  uploadSection.classList.add('hidden');
  progressSection.classList.add('hidden');
  resultSection.classList.add('hidden');
  errorSection.classList.add('hidden');
  if (name === 'upload')    uploadSection.classList.remove('hidden');
  if (name === 'progress')  progressSection.classList.remove('hidden');
  if (name === 'result')    resultSection.classList.remove('hidden');
  if (name === 'error')     errorSection.classList.remove('hidden');
}

function showError(msg) {
  errorMsg.textContent = msg;
  showSection('error');
}

function downloadBase64(b64, mime, filename) {
  const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  const blob  = new Blob([bytes], { type: mime });
  const url   = URL.createObjectURL(blob);
  const a     = document.createElement('a');
  a.href      = url;
  a.download  = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function esc(str) {
  if (!str) return '';
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// Exibir upload na carga inicial
showSection('upload');
