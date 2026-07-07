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
  itfMap: new Map(), // normalizedITF(外箱コード) -> [{ sheet, row }]
  collected: [],     // [{ sheet, row }]
  html5QrCode: null,
  scanning: false,
  corrections: [],   // [{ 区分:'JAN'|'ITF', 商品名, メーカー, 品目コード, 旧コード, 新コード, 検出日 }]
  correctionPendingCode: null, // 訂正申請フロー中に保持する「見つからなかったコード」
  correctionSelectedHit: null, // 訂正申請フローで検索して選んだ商品
  currentShelf: null, // ロケーション登録で今スキャンした商品を追加する先の棚(ロケ番号)
  mode: 'price', // 'price' | 'location' | 'correction'
  locations: [],     // [{ 商品名, メーカー, 品目コード, 旧ロケ, 新ロケ, 登録日 }]
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
  state.itfMap.clear();
  for (const sheet of state.sheets) {
    for (const row of sheet.rows) {
      const jan = normalizeJan(row['JAN']);
      if (jan) {
        if (!state.janMap.has(jan)) state.janMap.set(jan, []);
        state.janMap.get(jan).push({ sheet: sheet.name, row });
      }
      // ITF(外箱・段ボールの集合包装コード)からも同じ商品を引けるようにする
      const itf = normalizeJan(row['ITF']);
      if (itf) {
        if (!state.itfMap.has(itf)) state.itfMap.set(itf, []);
        state.itfMap.get(itf).push({ sheet: sheet.name, row });
      }
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

// 同一JAN/ITFに複数商品が紐づく場合があるため配列で返す
function lookupCandidates(rawCode) {
  const jan = normalizeJan(rawCode);
  if (!jan) return [];
  let list = state.janMap.get(jan);
  if (!list) {
    // UPC-A(12桁) <-> EAN-13(先頭0付き13桁) の揺れを吸収
    if (jan.length === 12) list = state.janMap.get('0' + jan);
    else if (jan.length === 13 && jan.startsWith('0')) list = state.janMap.get(jan.slice(1));
  }
  if (!list) {
    // JANで見つからない場合は外箱のITFコードとして検索
    list = state.itfMap.get(jan);
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
  applyModeVisibility();
  // バーコードリーダーですぐ読み取れるよう、JAN入力欄にフォーカスしておく
  setTimeout(() => $('manualJan').focus(), 0);
}

/* ---------- モード切り替え（プライス発行／ロケ登録／申請） ---------- */

const MODES = ['price', 'location', 'correction'];

function setMode(mode) {
  state.mode = mode;
  localStorage.setItem('jan-scanner-mode', mode);
  MODES.forEach(m => $(`modeBtn-${m}`).classList.toggle('active', m === mode));
  clearMatchAndInput();
  applyModeVisibility();
}

MODES.forEach(m => {
  $(`modeBtn-${m}`).addEventListener('click', () => setMode(m));
});

// モードと中身の有無に応じて、各出力リストの表示/非表示をまとめて切り替える
function applyModeVisibility() {
  $('shelfCard').classList.toggle('hidden', state.mode !== 'location');
  $('listCard').classList.toggle('hidden', !(state.mode === 'price' && state.collected.length > 0));
  $('locationListCard').classList.toggle('hidden', !(state.mode === 'location' && state.locations.length > 0));
  $('correctionListCard').classList.toggle('hidden', !(state.mode === 'correction' && state.corrections.length > 0));
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
      Html5QrcodeSupportedFormats.ITF, // 外箱・段ボールの集合包装コード
    ],
    // 対応端末ではブラウザ標準のバーコード検出エンジンを使い、精度と速度を上げる
    experimentalFeatures: { useBarCodeDetectorIfSupported: true },
    verbose: false,
  });
  try {
    await state.html5QrCode.start(
      { facingMode: 'environment' },
      {
        fps: 15,
        // 1次元バーコードは横長なので、枠(170x110)に対して横幅を広めに取る
        qrbox: { width: 150, height: 80 },
        disableFlip: true,
        aspectRatio: 4 / 3,
      },
      onScanSuccess,
      onScanFailure
    );
    state.scanning = true;
    setupTorchToggle();
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
  $('torchBtn').classList.add('hidden');
}

// ライト(トーチ)に対応した端末だけボタンを表示する
function setupTorchToggle() {
  const btn = $('torchBtn');
  btn.classList.add('hidden');
  btn.dataset.on = '0';
  try {
    const track = state.html5QrCode?.getRunningTrackCameraCapabilities?.();
    const caps = track?.torchFeature?.() ?? null;
    if (caps && caps.isSupported && caps.isSupported()) {
      btn.classList.remove('hidden');
      btn.onclick = async () => {
        const turnOn = btn.dataset.on !== '1';
        try {
          await caps.apply(turnOn);
          btn.dataset.on = turnOn ? '1' : '0';
          btn.textContent = turnOn ? 'ライトを消す' : 'ライトを点ける';
        } catch (e) { toast('ライトを切り替えられませんでした'); }
      };
    }
  } catch (e) { /* 未対応端末は無視 */ }
}

// バーコードリーダー(USB/Bluetooth)は入力欄にJANを打ち込んだ後Enterを送るキーボードとして動作する。
// 毎回手で消さなくても続けてスキャンできるよう、検索後は自動的に欄を空にしてフォーカスを戻す。
function processManualCode() {
  const v = $('manualJan').value.trim();
  if (!v) return;
  handleCode(v);
  $('manualJan').value = '';
  $('manualJan').focus();
}
$('manualSearchBtn').addEventListener('click', processManualCode);
$('manualJan').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    processManualCode();
  }
});

/* ---------- 商品名・メーカーでのキーワード検索 ---------- */

let keywordSearchTimer = null;
$('keywordSearch').addEventListener('input', () => {
  clearTimeout(keywordSearchTimer);
  keywordSearchTimer = setTimeout(() => runKeywordSearch($('keywordSearch').value), 250);
});

// 全角/半角の表記ゆれ（"140g"と"140ｇ"など）を吸収するための正規化
function normalizeForSearch(s) {
  return String(s).normalize('NFKC').toLowerCase();
}

// スペース区切りの複数キーワードによるAND検索（商品名・メーカー・規格をまたいで一致すればOK）
function searchMasterRows(query, limit) {
  const terms = normalizeForSearch(query).split(/[\s　]+/).filter(Boolean);
  if (!terms.length) return [];
  const hits = [];
  outer:
  for (const sheet of state.sheets) {
    for (const row of sheet.rows) {
      const haystack = normalizeForSearch(
        [row['商品名'], row['メーカー'], row['規格']].filter(Boolean).join(' ')
      );
      if (terms.every(t => haystack.includes(t))) {
        hits.push({ sheet: sheet.name, row });
        if (hits.length >= limit) break outer;
      }
    }
  }
  return hits;
}

function runKeywordSearch(query) {
  const resultsEl = $('keywordResults');
  const q = query.trim();
  if (q.length < 2) {
    resultsEl.innerHTML = '';
    return;
  }
  const hits = searchMasterRows(q, 30);
  if (!hits.length) {
    resultsEl.innerHTML = `<p class="muted">該当する商品が見つかりません</p>`;
    return;
  }
  resultsEl.innerHTML = '';
  hits.forEach((hit) => {
    const r = hit.row;
    const btn = document.createElement('button');
    btn.className = 'secondary';
    btn.style.textAlign = 'left';
    btn.textContent = `${r['商品名'] || ''}（${r['メーカー'] || ''} / ${r['規格'] || '-'} / 卸${r['卸'] ?? r['卸（ランク1）'] ?? '-'}）`;
    btn.addEventListener('click', () => {
      renderHitBox(hit, r['JAN']);
      $('matchBox').scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
    resultsEl.appendChild(btn);
  });
  if (hits.length >= 30) {
    const note = document.createElement('p');
    note.className = 'muted';
    note.textContent = '候補が多いため上位30件のみ表示しています。もう少し詳しく入力してください。';
    resultsEl.appendChild(note);
  }
}

function renderHitBox(hit, code) {
  const box = $('matchBox');
  box.className = 'found';
  const r = hit.row;
  const isItf = normalizeJan(code) !== normalizeJan(r['JAN']);
  const codeLabel = isItf ? `ITF(外箱): ${escapeHtml(String(code))} ／ JAN: ${escapeHtml(String(r['JAN'] ?? '-'))}` : `JAN: ${escapeHtml(String(code))}`;
  let actionHtml = '';
  if (state.mode === 'price') {
    actionHtml = `<button id="addToListBtn">リストに追加</button>`;
  } else if (state.mode === 'location') {
    const shelfBtnLabel = state.currentShelf
      ? `この棚(${escapeHtml(state.currentShelf)})に追加`
      : '商品を追加する（棚が未選択）';
    actionHtml = `<button id="addLocationOpenBtn" class="secondary">${shelfBtnLabel}</button>`;
  }
  box.innerHTML = `
    <div class="name">${escapeHtml(r['商品名'] || '')} <span class="badge">${escapeHtml(hit.sheet)}</span></div>
    <div class="meta">${codeLabel} ／ 規格: ${escapeHtml(r['規格'] || '-')} ／ 卸: ${escapeHtml(String(r['卸'] ?? r['卸（ランク1）'] ?? '-'))} ／ ロケ: ${escapeHtml(r['ロケ'] || '-')}</div>
    ${actionHtml ? `<div style="margin-top:8px; display:flex; gap:8px;">${actionHtml}</div>` : ''}
  `;
  if (state.mode === 'price') {
    $('addToListBtn').addEventListener('click', () => {
      addToList(hit, code);
      clearMatchAndInput();
    });
  } else if (state.mode === 'location') {
    $('addLocationOpenBtn').addEventListener('click', () => addProductToCurrentShelf(hit));
  }
}

// 追加済みかどうか一目で分かるよう、表示と入力欄をクリアして次のスキャンに備える
function clearMatchAndInput() {
  const box = $('matchBox');
  box.classList.add('hidden');
  box.innerHTML = '';
  $('manualJan').value = '';
  $('manualJan').focus();
}

function handleCode(code) {
  const candidates = lookupCandidates(code);
  const box = $('matchBox');
  box.classList.remove('hidden');

  if (candidates.length === 0) {
    box.className = 'notfound';
    const correctionBtnHtml = state.mode === 'correction'
      ? `<div style="margin-top:8px;"><button id="openCorrectionBtn" class="secondary">JAN/ITFコードの訂正申請</button></div>`
      : '';
    box.innerHTML = `
      <div class="name">マスタに見つかりませんでした</div>
      <div class="meta">JAN: ${escapeHtml(String(code))}（新規商品の可能性、またはメーカー都合のコード変更の可能性があります）</div>
      ${correctionBtnHtml}
    `;
    if (state.mode === 'correction') {
      $('openCorrectionBtn').addEventListener('click', () => startCorrectionFlow(code));
    }
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
  // 重複判定は「スキャンしたコード」ではなく商品そのもの(実際のJAN)で行う
  // ※ITF(外箱コード)経由で追加された場合でもJANが同じなら重複扱いにするため
  const normJan = normalizeJan(hit.row['JAN']) ?? normalizeJan(jan);
  if (state.collected.some(c => (normalizeJan(c.row['JAN']) ?? '') === normJan)) {
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
  applyModeVisibility();
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

// プライスカード用に出力する列（この順番で固定）
const PRICE_CARD_COLUMNS = [
  { label: 'JANコード', get: (r) => r['JAN'] ?? '' },
  { label: 'メーカー', get: (r) => r['メーカー'] ?? '' },
  { label: '商品名', get: (r) => r['商品名'] ?? '' },
  { label: '規格', get: (r) => r['規格'] ?? '' },
  { label: 'ケース', get: (r) => r['ケース'] ?? '' },
  { label: 'ボール', get: (r) => r['ボール'] ?? '' },
  { label: '卸', get: (r) => r['卸'] ?? r['卸（ランク1）'] ?? '' },
  { label: 'ロケ', get: (r) => r['ロケ'] ?? '' },
  { label: '出荷', get: (r) => r['出荷'] ?? '' },
  { label: '切替日付', get: (r) => r['切替日付'] ?? '' },
];

function buildPriceCardRows() {
  return state.collected.map(c => PRICE_CARD_COLUMNS.map(col => col.get(c.row)));
}

$('copyListBtn').addEventListener('click', async () => {
  if (!state.collected.length) { toast('リストが空です'); return; }
  const header = PRICE_CARD_COLUMNS.map(c => c.label);
  const rows = buildPriceCardRows();
  let text = header.join('\t') + '\n';
  for (const row of rows) {
    text += row.join('\t') + '\n';
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

$('saveExcelBtn').addEventListener('click', () => {
  if (!state.collected.length) { toast('リストが空です'); return; }
  const header = PRICE_CARD_COLUMNS.map(c => c.label);
  const rows = buildPriceCardRows();
  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'プライスカード');
  const d = new Date();
  const stamp = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}_${String(d.getHours()).padStart(2,'0')}${String(d.getMinutes()).padStart(2,'0')}`;
  XLSX.writeFile(wb, `プライスカード_${stamp}.xlsx`);
  toast('Excelファイルを保存しました');
});

function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

/* ---------- ロケーション登録：棚を選ぶ → 商品を追加する ---------- */

function persistCurrentShelf() {
  if (state.currentShelf) localStorage.setItem('jan-scanner-current-shelf', state.currentShelf);
  else localStorage.removeItem('jan-scanner-current-shelf');
}

function renderCurrentShelf() {
  $('currentShelfDisplay').textContent = `現在の棚：${state.currentShelf || '未選択'}`;
}

$('setShelfBtn').addEventListener('click', () => {
  const v = $('shelfInput').value.trim();
  if (!v) { toast('棚（ロケ番号）を入力してください'); return; }
  state.currentShelf = v;
  persistCurrentShelf();
  renderCurrentShelf();
  $('shelfInput').value = '';
  toast(`棚を「${v}」に設定しました`);
});

// 棚が選択済みなら、スキャンした商品をワンタップでその棚に登録する
function addProductToCurrentShelf(hit) {
  if (!state.currentShelf) {
    toast('先に棚を選んでください');
    $('shelfCard').classList.remove('hidden');
    $('shelfCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
    $('shelfInput').focus();
    return;
  }
  const r = hit.row;
  const entry = {
    商品名: r['商品名'] || '',
    メーカー: r['メーカー'] || '',
    品目コード: r['品目ｺｰﾄﾞ'] || '',
    旧ロケ: r['ロケ'] || '',
    新ロケ: state.currentShelf,
    登録日: todayStr(),
  };
  state.locations.push(entry);
  persistLocations();
  renderLocationList();
  toast(`「${entry.商品名}」を棚(${state.currentShelf})に登録しました`);
  clearMatchAndInput();
}

function removeLocation(idx) {
  state.locations.splice(idx, 1);
  persistLocations();
  renderLocationList();
}

function persistLocations() {
  localStorage.setItem('jan-scanner-locations', JSON.stringify(state.locations));
}
function restoreLocations() {
  try {
    const raw = localStorage.getItem('jan-scanner-locations');
    if (raw) state.locations = JSON.parse(raw);
  } catch (e) { state.locations = []; }
}

function renderLocationList() {
  applyModeVisibility();
  const tbody = document.querySelector('#locationTable tbody');
  tbody.innerHTML = '';
  state.locations.forEach((loc, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><button class="small danger" data-idx="${idx}">削除</button></td>
      <td>${escapeHtml(loc.商品名)}</td>
      <td>${escapeHtml(loc.品目コード)}</td>
      <td>${escapeHtml(loc.旧ロケ)}</td>
      <td>${escapeHtml(loc.新ロケ)}</td>
    `;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll('button[data-idx]').forEach(btn => {
    btn.addEventListener('click', () => removeLocation(Number(btn.dataset.idx)));
  });
  $('locationListCount').textContent = state.locations.length;
}

$('clearLocationBtn').addEventListener('click', () => {
  if (state.locations.length && !confirm('ロケーション登録リストを空にします。よろしいですか？')) return;
  state.locations = [];
  persistLocations();
  renderLocationList();
});

const LOCATION_COLUMNS = ['商品名', 'メーカー', '品目コード', '旧ロケ', '新ロケ', '登録日'];

$('copyLocationBtn').addEventListener('click', async () => {
  if (!state.locations.length) { toast('リストが空です'); return; }
  let text = LOCATION_COLUMNS.join('\t') + '\n';
  for (const loc of state.locations) {
    text += LOCATION_COLUMNS.map(col => loc[col] ?? '').join('\t') + '\n';
  }
  try {
    await navigator.clipboard.writeText(text);
    toast('コピーしました');
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

$('saveLocationExcelBtn').addEventListener('click', () => {
  if (!state.locations.length) { toast('リストが空です'); return; }
  const rows = state.locations.map(loc => LOCATION_COLUMNS.map(col => loc[col] ?? ''));
  const ws = XLSX.utils.aoa_to_sheet([LOCATION_COLUMNS, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'ロケーション登録');
  const d = new Date();
  const stamp = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}_${String(d.getHours()).padStart(2,'0')}${String(d.getMinutes()).padStart(2,'0')}`;
  XLSX.writeFile(wb, `ロケーション登録_${stamp}.xlsx`);
  toast('Excelファイルを保存しました');
});

/* ---------- JAN/ITFコード訂正申請 ---------- */

function startCorrectionFlow(code) {
  state.correctionPendingCode = code;
  state.correctionSelectedHit = null;
  // 13桁ならJAN、14桁ならITFの可能性が高いので初期値として推測する
  const digits = String(code).replace(/[^0-9]/g, '');
  $('correctionKubun').value = digits.length === 14 ? 'ITF' : 'JAN';
  $('correctionTargetInfo').textContent = `見つからなかったコード: ${code}`;
  $('correctionSearch').value = '';
  $('correctionSearchResults').innerHTML = '';
  $('correctionConfirm').classList.add('hidden');
  $('correctionCard').classList.remove('hidden');
  $('correctionCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  setTimeout(() => $('correctionSearch').focus(), 100);
}

let correctionSearchTimer = null;
$('correctionSearch').addEventListener('input', () => {
  clearTimeout(correctionSearchTimer);
  correctionSearchTimer = setTimeout(() => runCorrectionSearch($('correctionSearch').value), 250);
});

function runCorrectionSearch(query) {
  const resultsEl = $('correctionSearchResults');
  const q = query.trim();
  if (q.length < 2) {
    resultsEl.innerHTML = '';
    return;
  }
  const hits = searchMasterRows(q, 30);
  resultsEl.innerHTML = '';
  if (!hits.length) {
    resultsEl.innerHTML = `<p class="muted">該当する商品が見つかりません</p>`;
    return;
  }
  hits.forEach((hit) => {
    const r = hit.row;
    const btn = document.createElement('button');
    btn.className = 'secondary';
    btn.style.textAlign = 'left';
    btn.textContent = `${r['商品名'] || ''}（${r['メーカー'] || ''} / 品目コード:${r['品目ｺｰﾄﾞ'] || '-'}）`;
    btn.addEventListener('click', () => selectCorrectionCandidate(hit));
    resultsEl.appendChild(btn);
  });
}

function selectCorrectionCandidate(hit) {
  state.correctionSelectedHit = hit;
  const r = hit.row;
  $('correctionConfirmName').textContent = `${r['商品名'] || ''}（${r['メーカー'] || ''} / 品目コード:${r['品目ｺｰﾄﾞ'] || '-'}）`;
  updateCorrectionOldCode();
  $('correctionNewCode').value = state.correctionPendingCode;
  $('correctionConfirm').classList.remove('hidden');
}

function updateCorrectionOldCode() {
  if (!state.correctionSelectedHit) return;
  const r = state.correctionSelectedHit.row;
  const kubun = $('correctionKubun').value;
  $('correctionOldCode').value = (kubun === 'ITF' ? r['ITF'] : r['JAN']) ?? '(未登録)';
}
$('correctionKubun').addEventListener('change', updateCorrectionOldCode);

$('cancelCorrectionBtn').addEventListener('click', () => {
  $('correctionCard').classList.add('hidden');
});

$('addCorrectionBtn').addEventListener('click', () => {
  if (!state.correctionSelectedHit) { toast('商品を選択してください'); return; }
  const r = state.correctionSelectedHit.row;
  const kubun = $('correctionKubun').value;
  const entry = {
    区分: kubun,
    商品名: r['商品名'] || '',
    メーカー: r['メーカー'] || '',
    品目コード: r['品目ｺｰﾄﾞ'] || '',
    旧コード: (kubun === 'ITF' ? r['ITF'] : r['JAN']) ?? '',
    新コード: $('correctionNewCode').value.trim(),
    検出日: todayStr(),
  };
  state.corrections.push(entry);
  persistCorrections();
  renderCorrectionList();
  toast('訂正申請リストに追加しました');

  // 一連の流れを終え、次のスキャンに備える
  $('correctionCard').classList.add('hidden');
  clearMatchAndInput();
});

function removeCorrection(idx) {
  state.corrections.splice(idx, 1);
  persistCorrections();
  renderCorrectionList();
}

function persistCorrections() {
  localStorage.setItem('jan-scanner-corrections', JSON.stringify(state.corrections));
}
function restoreCorrections() {
  try {
    const raw = localStorage.getItem('jan-scanner-corrections');
    if (raw) state.corrections = JSON.parse(raw);
  } catch (e) { state.corrections = []; }
}

function renderCorrectionList() {
  applyModeVisibility();
  const tbody = document.querySelector('#correctionTable tbody');
  tbody.innerHTML = '';
  state.corrections.forEach((c, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><button class="small danger" data-idx="${idx}">削除</button></td>
      <td>${escapeHtml(c.区分)}</td>
      <td>${escapeHtml(c.商品名)}</td>
      <td>${escapeHtml(c.品目コード)}</td>
      <td>${escapeHtml(c.旧コード)}</td>
      <td>${escapeHtml(c.新コード)}</td>
    `;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll('button[data-idx]').forEach(btn => {
    btn.addEventListener('click', () => removeCorrection(Number(btn.dataset.idx)));
  });
  $('correctionListCount').textContent = state.corrections.length;
}

$('clearCorrectionBtn').addEventListener('click', () => {
  if (state.corrections.length && !confirm('コード訂正申請リストを空にします。よろしいですか？')) return;
  state.corrections = [];
  persistCorrections();
  renderCorrectionList();
});

const CORRECTION_COLUMNS = ['区分', '商品名', 'メーカー', '品目コード', '旧コード', '新コード', '検出日'];

$('copyCorrectionBtn').addEventListener('click', async () => {
  if (!state.corrections.length) { toast('リストが空です'); return; }
  let text = CORRECTION_COLUMNS.join('\t') + '\n';
  for (const c of state.corrections) {
    text += CORRECTION_COLUMNS.map(col => c[col] ?? '').join('\t') + '\n';
  }
  try {
    await navigator.clipboard.writeText(text);
    toast('コピーしました');
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

$('saveCorrectionExcelBtn').addEventListener('click', () => {
  if (!state.corrections.length) { toast('リストが空です'); return; }
  const rows = state.corrections.map(c => CORRECTION_COLUMNS.map(col => c[col] ?? ''));
  const ws = XLSX.utils.aoa_to_sheet([CORRECTION_COLUMNS, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'コード訂正申請');
  const d = new Date();
  const stamp = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}_${String(d.getHours()).padStart(2,'0')}${String(d.getMinutes()).padStart(2,'0')}`;
  XLSX.writeFile(wb, `コード訂正申請_${stamp}.xlsx`);
  toast('Excelファイルを保存しました');
});

/* ---------- 初期化 ---------- */

(async function init() {
  state.mode = localStorage.getItem('jan-scanner-mode') || 'price';
  MODES.forEach(m => $(`modeBtn-${m}`).classList.toggle('active', m === state.mode));

  restoreCollected();
  renderList();
  restoreCorrections();
  renderCorrectionList();
  restoreLocations();
  renderLocationList();
  state.currentShelf = localStorage.getItem('jan-scanner-current-shelf') || null;
  renderCurrentShelf();
  const loaded = await tryLoadSavedMaster();
  if (!loaded) {
    $('setupCard').classList.remove('hidden');
  }
})();
