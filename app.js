/* JANスキャナー: 商品マスタ検索 + 収集リスト作成 */

const DB_NAME = 'jan-scanner-db';
const DB_STORE = 'kv';

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(DB_STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function idbGet(key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readonly');
    const req = tx.objectStore(DB_STORE).get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function idbSet(key, value) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readwrite');
    tx.objectStore(DB_STORE).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

const state = {
  sheets: [],       // [{ name, header:[...], rows:[{...}] }]
  janMap: new Map(), // normalizedJan -> [{ sheet, row }]
  collected: [],     // [{ sheet, row }]
  applications: [],  // [{ 区分, 実行日付, 入力日付, 商品コード, 小売, 仕切, 備考, _display:{商品名, JAN} }]
  html5QrCode: null,
  scanning: false,
  lastMatch: null,   // 直近にマッチした商品（申請行作成の対象）
  appTemplate: null, // { fileLabel, sheetName, header:[...] }
};

const $ = (id) => document.getElementById(id);

function toast(msg) {
  const el = $('toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.remove('show'), 1800);
}

function normalizeJan(raw) {
  if (raw === null || raw === undefined) return null;
  const digits = String(raw).replace(/[^0-9]/g, '');
  if (!digits) return null;
  return String(parseInt(digits, 10));
}

function buildJanMap() {
  state.janMap.clear();
  for (const sheet of state.sheets) {
    for (const row of sheet.rows) {
      const jan = normalizeJan(row['JAN']);
      if (!jan) continue;
      if (!state.janMap.has(jan)) state.janMap.set(jan, []);
      state.janMap.get(jan).push({ sheet: sheet.name, row });
    }
  }
}

function candidateKey(c) {
  const r = c.row;
  return [c.sheet, r['商品名'], r['卸'] ?? r['卸（ランク1）'] ?? ''].join('|');
}

// マスタ内の同一JAN重複行のうち、内容が実質同一のものは1件にまとめる
function dedupeCandidates(list) {
  const seen = new Set();
  const out = [];
  for (const c of list) {
    const k = candidateKey(c);
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(c);
  }
  return out;
}

// 同一JANに複数商品が紐づく場合があるため配列で返す
function lookupCandidates(rawCode) {
  const jan = normalizeJan(rawCode);
  if (!jan) return [];
  let list = state.janMap.get(jan);
  if (!list) {
    // UPC-A(12桁) <-> EAN-13(先頭0付き13桁) の揺れを吸収
    if (jan.length === 12) list = state.janMap.get('0' + jan);
    else if (jan.length === 13 && jan.startsWith('0')) list = state.janMap.get(jan.slice(1));
  }
  return list ? dedupeCandidates(list) : [];
}

/* ---------- マスタ読み込み ---------- */

async function loadWorkbookFromArrayBuffer(buf, fileLabel) {
  const wb = XLSX.read(buf, { type: 'array' });
  const sheets = wb.SheetNames.map(name => {
    const ws = wb.Sheets[name];
    const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
    const header = json.length ? Object.keys(json[0]) : [];
    return { name, header, rows: json };
  });
  state.sheets = sheets;
  buildJanMap();

  const rowCount = sheets.reduce((sum, s) => sum + s.rows.length, 0);
  const meta = { fileLabel, loadedAt: new Date().toISOString(), rowCount, sheetNames: sheets.map(s => s.name) };
  await idbSet('sheets', sheets);
  await idbSet('meta', meta);
  renderMasterInfo(meta);
  showMainScreens();
}

async function tryLoadSavedMaster() {
  const sheets = await idbGet('sheets');
  const meta = await idbGet('meta');
  if (sheets && sheets.length) {
    state.sheets = sheets;
    buildJanMap();
    renderMasterInfo(meta);
    showMainScreens();
    return true;
  }
  return false;
}

function renderMasterInfo(meta) {
  if (!meta) return;
  const d = new Date(meta.loadedAt);
  const dateStr = `${d.getFullYear()}/${d.getMonth()+1}/${d.getDate()} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
  $('masterInfo').textContent =
    `${meta.fileLabel} を読込済み（${dateStr}） / シート: ${meta.sheetNames.join('・')} / 全${meta.rowCount}件`;
  $('headerStatus').textContent = `${meta.rowCount}件読込済`;
  $('reloadMasterBtn').classList.remove('hidden');
}

function showMainScreens() {
  $('scanCard').classList.remove('hidden');
  $('listCard').classList.remove('hidden');
}

$('masterFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  $('setupStatus').textContent = '読み込み中...';
  try {
    const buf = await file.arrayBuffer();
    await loadWorkbookFromArrayBuffer(buf, file.name);
    $('setupStatus').textContent = '読み込み完了しました。';
    toast('マスタデータを読み込みました');
  } catch (err) {
    console.error(err);
    $('setupStatus').textContent = '読み込みに失敗しました: ' + err.message;
  }
});

$('reloadMasterBtn').addEventListener('click', () => {
  if (confirm('現在のマスタデータを新しいファイルで置き換えます。よろしいですか？')) {
    $('masterFile').value = '';
    $('masterFile').click();
  }
});

/* ---------- スキャン ---------- */

function beep() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain); gain.connect(ctx.destination);
    osc.frequency.value = 880;
    gain.gain.setValueAtTime(0.15, ctx.currentTime);
    osc.start();
    osc.stop(ctx.currentTime + 0.12);
  } catch (e) { /* no-op */ }
  if (navigator.vibrate) navigator.vibrate(80);
}

let lastScanCode = null;
let lastScanTime = 0;

function onScanSuccess(decodedText) {
  const now = Date.now();
  if (decodedText === lastScanCode && (now - lastScanTime) < 2000) return; // 連続重複防止
  lastScanCode = decodedText;
  lastScanTime = now;
  beep();
  handleCode(decodedText);
}

function onScanFailure() { /* 毎フレーム呼ばれるため無視 */ }

$('startScanBtn').addEventListener('click', async () => {
  $('reader').classList.remove('hidden');
  $('startScanBtn').classList.add('hidden');
  $('stopScanBtn').classList.remove('hidden');
  state.html5QrCode = new Html5Qrcode('reader', {
    formatsToSupport: [
      Html5QrcodeSupportedFormats.EAN_13,
      Html5QrcodeSupportedFormats.EAN_8,
      Html5QrcodeSupportedFormats.UPC_A,
      Html5QrcodeSupportedFormats.UPC_E,
      Html5QrcodeSupportedFormats.CODE_128,
    ],
    verbose: false,
  });
  try {
    await state.html5QrCode.start(
      { facingMode: 'environment' },
      { fps: 10, qrbox: { width: 280, height: 160 } },
      onScanSuccess,
      onScanFailure
    );
    state.scanning = true;
  } catch (err) {
    console.error(err);
    toast('カメラを起動できませんでした: ' + err);
    $('startScanBtn').classList.remove('hidden');
    $('stopScanBtn').classList.add('hidden');
    $('reader').classList.add('hidden');
  }
});

$('stopScanBtn').addEventListener('click', stopScanning);

async function stopScanning() {
  if (state.html5QrCode && state.scanning) {
    try { await state.html5QrCode.stop(); await state.html5QrCode.clear(); } catch (e) {}
  }
  state.scanning = false;
  $('reader').classList.add('hidden');
  $('startScanBtn').classList.remove('hidden');
  $('stopScanBtn').classList.add('hidden');
}

$('manualSearchBtn').addEventListener('click', () => {
  const v = $('manualJan').value.trim();
  if (v) handleCode(v);
});
$('manualJan').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') $('manualSearchBtn').click();
});

function renderHitBox(hit, code) {
  const box = $('matchBox');
  box.className = 'found';
  const r = hit.row;
  box.innerHTML = `
    <div class="name">${escapeHtml(r['商品名'] || '')} <span class="badge">${escapeHtml(hit.sheet)}</span></div>
    <div class="meta">JAN: ${escapeHtml(String(code))} ／ 規格: ${escapeHtml(r['規格'] || '-')} ／ 卸: ${escapeHtml(String(r['卸'] ?? r['卸（ランク1）'] ?? '-'))} ／ ロケ: ${escapeHtml(r['ロケ'] || '-')}</div>
    <div style="margin-top:8px; display:flex; gap:8px;">
      <button id="addToListBtn">リストに追加</button>
      <button id="setAppTargetBtn" class="secondary">申請対象にする</button>
    </div>
  `;
  $('addToListBtn').addEventListener('click', () => addToList(hit, code));
  $('setAppTargetBtn').addEventListener('click', () => setAppTarget(hit, code));
}

function setAppTarget(hit, code) {
  state.lastMatch = { hit, code };
  const r = hit.row;
  $('appFormTarget').textContent = `対象商品：${r['商品名'] || ''}（商品コード: ${r['品目ｺｰﾄﾞ'] || '未取得'} ／ JAN: ${code}）`;
  toast('申請対象に設定しました');
}

function handleCode(code) {
  const candidates = lookupCandidates(code);
  const box = $('matchBox');
  box.classList.remove('hidden');

  if (candidates.length === 0) {
    box.className = 'notfound';
    box.innerHTML = `
      <div class="name">マスタに見つかりませんでした</div>
      <div class="meta">JAN: ${escapeHtml(String(code))}（新規商品の可能性があります）</div>
    `;
    return;
  }

  if (candidates.length === 1) {
    renderHitBox(candidates[0], code);
    return;
  }

  // 同一JANに複数商品が登録されているため選択させる
  box.className = 'notfound';
  box.innerHTML = `
    <div class="name">⚠ このJANには複数の商品が登録されています</div>
    <div class="meta">JAN: ${escapeHtml(String(code))} ／ 該当する商品を選んでください</div>
    <div id="candidateList" style="margin-top:8px; display:flex; flex-direction:column; gap:6px;"></div>
  `;
  const listEl = $('candidateList');
  candidates.forEach((c, i) => {
    const r = c.row;
    const btn = document.createElement('button');
    btn.className = 'secondary';
    btn.style.textAlign = 'left';
    btn.textContent = `${r['商品名'] || ''}（${c.sheet} / 卸${r['卸'] ?? r['卸（ランク1）'] ?? '-'}）`;
    btn.addEventListener('click', () => renderHitBox(c, code));
    listEl.appendChild(btn);
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

/* ---------- 収集リスト ---------- */

function addToList(hit, jan) {
  const normJan = normalizeJan(jan);
  if (state.collected.some(c => normalizeJan(c.row['JAN']) === normJan)) {
    toast('すでにリストにあります');
    return;
  }
  state.collected.push({ sheet: hit.sheet, row: hit.row });
  persistCollected();
  renderList();
  toast('リストに追加しました');
}

function removeFromList(idx) {
  state.collected.splice(idx, 1);
  persistCollected();
  renderList();
}

function persistCollected() {
  localStorage.setItem('jan-scanner-collected', JSON.stringify(state.collected));
}
function restoreCollected() {
  try {
    const raw = localStorage.getItem('jan-scanner-collected');
    if (raw) state.collected = JSON.parse(raw);
  } catch (e) { state.collected = []; }
}

function renderList() {
  const tbody = document.querySelector('#listTable tbody');
  tbody.innerHTML = '';
  state.collected.forEach((c, idx) => {
    const r = c.row;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><button class="small danger" data-idx="${idx}">削除</button></td>
      <td>${escapeHtml(r['JAN'])}</td>
      <td>${escapeHtml(r['商品名'] || '')}</td>
      <td>${escapeHtml(r['規格'] || '-')}</td>
      <td>${escapeHtml(String(r['卸'] ?? r['卸（ランク1）'] ?? '-'))}</td>
      <td>${escapeHtml(r['ロケ'] || '-')}</td>
      <td>${escapeHtml(c.sheet)}</td>
    `;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll('button[data-idx]').forEach(btn => {
    btn.addEventListener('click', () => removeFromList(Number(btn.dataset.idx)));
  });
  $('listCount').textContent = state.collected.length;
}

$('clearListBtn').addEventListener('click', () => {
  if (state.collected.length && !confirm('収集リストを空にします。よろしいですか？')) return;
  state.collected = [];
  persistCollected();
  renderList();
});

$('copyListBtn').addEventListener('click', async () => {
  if (!state.collected.length) { toast('リストが空です'); return; }
  const bySheet = new Map();
  for (const c of state.collected) {
    if (!bySheet.has(c.sheet)) bySheet.set(c.sheet, []);
    bySheet.get(c.sheet).push(c.row);
  }
  let text = '';
  for (const [sheetName, rows] of bySheet) {
    const header = state.sheets.find(s => s.name === sheetName)?.header || Object.keys(rows[0]);
    text += `=== ${sheetName} ===\n`;
    text += header.join('\t') + '\n';
    for (const row of rows) {
      text += header.map(h => row[h] ?? '').join('\t') + '\n';
    }
    text += '\n';
  }
  try {
    await navigator.clipboard.writeText(text);
    toast('コピーしました。夢ぷりんと/Excelに貼り付けてください');
  } catch (e) {
    // クリップボードAPIが使えない環境向けフォールバック
    const ta = document.createElement('textarea');
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    toast('コピーしました');
  }
});

/* ---------- 申請テンプレート ---------- */

let pendingAppWorkbook = null; // シート選択待ちの間だけ保持

$('appTemplateFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  pendingAppWorkbook = { wb, fileLabel: file.name };

  const select = $('appSheetSelect');
  select.innerHTML = '';
  wb.SheetNames.forEach((name) => {
    const opt = document.createElement('option');
    opt.value = name;
    opt.textContent = name;
    select.appendChild(opt);
  });
  // それらしい名前を推測して初期選択
  const guess = wb.SheetNames.find(n => n.includes('登録') || n.includes('入力'));
  if (guess) select.value = guess;

  $('appSheetPicker').classList.remove('hidden');
  $('appTemplateStatus').textContent = 'シートを選択して「これに決定」を押してください';
});

$('appSheetConfirmBtn').addEventListener('click', async () => {
  if (!pendingAppWorkbook) return;
  const sheetName = $('appSheetSelect').value;
  const ws = pendingAppWorkbook.wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const header = (rows[0] || []).map(h => (h === undefined ? '' : String(h)));

  state.appTemplate = { fileLabel: pendingAppWorkbook.fileLabel, sheetName, header };
  await idbSet('appTemplate', state.appTemplate);

  $('appTemplateStatus').textContent = `「${sheetName}」を申請一覧表として設定しました（${header.filter(h => h).length}列）`;
  $('appSheetPicker').classList.add('hidden');
  $('appFormCard').classList.remove('hidden');
  $('appListCard').classList.remove('hidden');
  toast('申請テンプレートを設定しました');
});

async function tryLoadSavedAppTemplate() {
  const tpl = await idbGet('appTemplate');
  if (tpl) {
    state.appTemplate = tpl;
    $('appTemplateStatus').textContent = `「${tpl.sheetName}」（${tpl.fileLabel}）を申請一覧表として使用中`;
    $('appFormCard').classList.remove('hidden');
    $('appListCard').classList.remove('hidden');
  }
}

/* ---------- 申請リスト ---------- */

function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
$('appNyuryokuDate').value = todayStr();

function dateInputToExcelStr(v) {
  if (!v) return '';
  const [y, m, d] = v.split('-');
  return `${y}/${Number(m)}/${Number(d)}`;
}

$('addAppBtn').addEventListener('click', () => {
  if (!state.lastMatch) {
    toast('先にスキャン結果から「申請対象にする」を選んでください');
    return;
  }
  const r = state.lastMatch.hit.row;
  const entry = {
    区分: $('appKubun').value,
    実行日付: dateInputToExcelStr($('appJikkoDate').value),
    入力日付: dateInputToExcelStr($('appNyuryokuDate').value),
    商品コード: r['品目ｺｰﾄﾞ'] || '',
    小売: $('appKouri').value,
    仕切: $('appShikiri').value,
    備考: $('appBiko').value,
    _display: { 商品名: r['商品名'] || '', JAN: state.lastMatch.code },
  };
  state.applications.push(entry);
  persistApplications();
  renderAppList();
  toast('申請リストに追加しました');
  $('appKouri').value = '';
  $('appShikiri').value = '';
  $('appBiko').value = '';
});

function removeApplication(idx) {
  state.applications.splice(idx, 1);
  persistApplications();
  renderAppList();
}

function persistApplications() {
  localStorage.setItem('jan-scanner-applications', JSON.stringify(state.applications));
}
function restoreApplications() {
  try {
    const raw = localStorage.getItem('jan-scanner-applications');
    if (raw) state.applications = JSON.parse(raw);
  } catch (e) { state.applications = []; }
}

function renderAppList() {
  const tbody = document.querySelector('#appTable tbody');
  tbody.innerHTML = '';
  state.applications.forEach((a, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><button class="small danger" data-idx="${idx}">削除</button></td>
      <td>${escapeHtml(a.区分)}</td>
      <td>${escapeHtml(a._display?.商品名 || '')}</td>
      <td>${escapeHtml(a.商品コード || '(未登録)')}</td>
      <td>${escapeHtml(a.実行日付)}</td>
      <td>${escapeHtml(a.小売)}</td>
      <td>${escapeHtml(a.仕切)}</td>
      <td>${escapeHtml(a.備考)}</td>
    `;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll('button[data-idx]').forEach(btn => {
    btn.addEventListener('click', () => removeApplication(Number(btn.dataset.idx)));
  });
  $('appListCount').textContent = state.applications.length;
}

$('clearAppBtn').addEventListener('click', () => {
  if (state.applications.length && !confirm('申請リストを空にします。よろしいですか？')) return;
  state.applications = [];
  persistApplications();
  renderAppList();
});

$('copyAppBtn').addEventListener('click', async () => {
  if (!state.applications.length) { toast('申請リストが空です'); return; }
  if (!state.appTemplate) { toast('申請テンプレートが未設定です'); return; }
  const header = state.appTemplate.header;
  let text = '';
  for (const a of state.applications) {
    text += header.map(h => (h && a[h] !== undefined) ? a[h] : '').join('\t') + '\n';
  }
  try {
    await navigator.clipboard.writeText(text);
    toast('コピーしました。表の一番下の空き行に貼り付けてください');
  } catch (e) {
    const ta = document.createElement('textarea');
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    toast('コピーしました');
  }
});

/* ---------- 初期化 ---------- */

(async function init() {
  restoreCollected();
  renderList();
  restoreApplications();
  renderAppList();
  await tryLoadSavedAppTemplate();
  const loaded = await tryLoadSavedMaster();
  if (!loaded) {
    $('setupCard').classList.remove('hidden');
  }
})();
