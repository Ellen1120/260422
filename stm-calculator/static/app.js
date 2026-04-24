'use strict';

// ── 상태 ─────────────────────────────────────────────────
let _products = [];
let _selected = { productId: null, strength: null };
let _calcResult = null;
let _sortedSolutions = [];
let _pollTimer = null;

// ── 초기화 ────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  loadProducts();
});

async function loadProducts() {
  try {
    const res = await fetch('/api/products');
    _products = await res.json();
    renderProductSelect();
  } catch (e) {
    console.error('제품 목록 로드 실패', e);
  }
}

function renderProductSelect() {
  const sel = document.getElementById('sel-product');
  sel.innerHTML = '<option value="">-- 제품 선택 --</option>';
  _products.forEach(p => {
    const opt = document.createElement('option');
    opt.value = p.id;
    opt.textContent = p.code_no || p.name;
    sel.appendChild(opt);
  });
  if (_products.length === 0) {
    const opt = document.createElement('option');
    opt.disabled = true;
    opt.textContent = '파싱된 제품 없음 (관리 메뉴에서 파싱하세요)';
    sel.appendChild(opt);
  }
}

// ── Step 1: 제품 선택 ─────────────────────────────────────
function onProductChange() {
  const productId = document.getElementById('sel-product').value;
  _selected.productId = productId;
  _selected.strength = null;

  hide('row-strength'); hide('row-tests'); hide('row-batches');
  document.getElementById('btn-calc').disabled = true;
  hide('result-section');

  if (!productId) return;

  const product = _products.find(p => p.id === productId);
  if (!product) return;

  document.getElementById('hint-stm').textContent = `${product.name}  (${product.stm_file})`;

  const selStr   = document.getElementById('sel-strength');
  const lblFixed = document.getElementById('lbl-strength-fixed');

  if (product.strengths.length === 1) {
    selStr.style.display   = 'none';
    lblFixed.style.display = '';
    lblFixed.textContent   = product.strengths[0];
    show('row-strength');
    _selected.strength = product.strengths[0];
    onStrengthChange(product);
    return;
  }

  selStr.style.display   = '';
  lblFixed.style.display = 'none';
  selStr.innerHTML = '';
  product.strengths.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    selStr.appendChild(opt);
  });
  selStr.onchange = () => onStrengthChange(product);
  show('row-strength');
  selStr.value = product.strengths[0];
  onStrengthChange(product);
}

// ── Step 2: 함량 선택 ─────────────────────────────────────
function onStrengthChange(product) {
  const lblFixed = document.getElementById('lbl-strength-fixed');
  if (lblFixed.style.display !== 'none') {
    _selected.strength = lblFixed.textContent.trim();
  } else {
    _selected.strength = document.getElementById('sel-strength').value;
  }
  hide('row-tests');
  document.getElementById('btn-calc').disabled = true;
  hide('result-section');

  if (!_selected.strength) return;

  const container = document.getElementById('test-checkboxes');
  container.innerHTML = '';
  product.test_items.forEach(name => {
    const lbl = document.createElement('label');
    const cb  = document.createElement('input');
    cb.type = 'checkbox'; cb.value = name;
    cb.addEventListener('change', syncCalcButton);
    lbl.append(cb, ' ' + name);
    container.appendChild(lbl);
  });
  show('row-tests');
  show('row-batches');
}

function syncCalcButton() {
  const any = document.querySelectorAll('#test-checkboxes input:checked').length > 0;
  document.getElementById('btn-calc').disabled = !any;
}

// ── 이론량 계산 ──────────────────────────────────────────
async function calculate() {
  const testItems = Array.from(
    document.querySelectorAll('#test-checkboxes input:checked')
  ).map(cb => cb.value);

  const batches = parseInt(document.getElementById('inp-batches').value) || 1;
  const btn = document.getElementById('btn-calc');
  btn.disabled = true;
  btn.textContent = '계산 중…';

  try {
    const res = await fetch('/api/calculate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        product_id: _selected.productId,
        strength:   _selected.strength,
        test_items: testItems,
        batch_count: batches,
      }),
    });
    if (!res.ok) {
      const err = await res.json();
      alert('오류: ' + (err.detail || '알 수 없는 오류'));
      return;
    }
    _calcResult = await res.json();
    renderResult(_calcResult);
  } catch (e) {
    alert('서버 오류: ' + e.message);
  } finally {
    btn.disabled = false;
    btn.textContent = '이론량 계산';
    syncCalcButton();
  }
}

// ── 결과 렌더링 ──────────────────────────────────────────
function renderResult(data) {
  document.getElementById('result-subtitle').textContent =
    `${data.product_name}  ·  ${data.strength}  ·  ${data.batch_count}배치`;

  const docNoEl = document.getElementById('result-docno');
  if (data.doc_no) {
    docNoEl.textContent = data.doc_no;
    docNoEl.style.display = '';
  } else {
    docNoEl.style.display = 'none';
  }

  renderSolutionTable(data.solutions, data.dissolution_medium);
  renderGlasswareTable(data.glassware);
  renderPipetteTable(data.pipettes || []);
  renderFilterTable(data.filters || []);

  show('result-section');
  document.getElementById('result-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// 솔루션 표에서 제외할 이름 패턴
const _HIDDEN_SOL = /^(standard\s+stock\s+solution|standard\s+solution|sample\s+solution)/i;

// ── 용액 테이블 ───────────────────────────────────────────
function renderSolutionTable(solutions, dm) {
  const tbody = document.querySelector('#tbl-solutions tbody');
  tbody.innerHTML = '';

  const visible = solutions.filter(s => !_HIDDEN_SOL.test((s.solution_name || '').trim()));

  if (!visible.length) {
    tbody.innerHTML = '<tr><td colspan="3" class="empty-msg">조제 정보 없음</td></tr>';
    return;
  }

  // 원본 solutions를 visible로 교체해 이후 로직에 반영
  solutions = visible;

  const solutionPriority = name => {
    const n = (name || '').toLowerCase();
    if (/mobile\s*phase/i.test(n)) return 0;
    if (/\bbuffer\b/i.test(n)) return 1;
    if (/dissolution/i.test(n)) return 2;
    return 3;
  };

  _sortedSolutions = [...solutions].sort((a, b) =>
    solutionPriority(a.solution_name) - solutionPriority(b.solution_name)
  );

  _sortedSolutions.forEach((s, idx) => {
    const tr = document.createElement('tr');

    const isDissolutionMedium = /dissolution/i.test(s.solution_name);
    let effectiveVolumeMl, breakdownHtml = '';

    if (isDissolutionMedium && dm && dm.total_medium_ml) {
      effectiveVolumeMl = dm.total_medium_ml;
      const parts = [`검액 ${fmt(dm.sample_medium_ml)} mL`];
      if (dm.standard_medium_ml_once > 0) parts.push(`표준액 ${fmt(dm.standard_medium_ml_once)} mL`);
      breakdownHtml = `<div class="diss-breakdown-inline">(${parts.join(' + ')})</div>`;
    } else {
      effectiveVolumeMl = s.theoretical_volume_ml;
    }

    const theoretical = effectiveVolumeMl != null
      ? fmt(effectiveVolumeMl) + ' mL' : '-';
    const placeholderVal = effectiveVolumeMl != null ? String(effectiveVolumeMl) : '';

    const inputId   = `prep-input-${idx}`;
    const reagentId = `reagent-out-${idx}`;

    tr.innerHTML = `
      <td>${esc(s.solution_name)}</td>
      <td class="num bold">
        ${theoretical}
        ${breakdownHtml}
      </td>
      <td>
        <div class="prep-input-cell">
          <input type="number" id="${inputId}" min="0" step="1"
                 placeholder="${placeholderVal}"
                 data-idx="${idx}"
                 data-theoretical="${placeholderVal}"
                 oninput="onPrepAmountInput(this)"
                 style="width:110px" />
          <span class="unit">mL</span>
          <div class="reagent-list" id="${reagentId}"></div>
        </div>
      </td>
    `;
    tbody.appendChild(tr);

    const outEl = tr.querySelector('.reagent-list');
    if (outEl) _renderReagents(outEl, s, effectiveVolumeMl);
  });
}

function onPrepAmountInput(input) {
  const idx          = parseInt(input.dataset.idx);
  const theoreticalMl = parseFloat(input.dataset.theoretical);
  const prepMl       = parseFloat(input.value);
  const sol          = _sortedSolutions[idx];
  const outEl        = document.getElementById(`reagent-out-${idx}`);

  if (input.value.trim() !== '') {
    input.classList.add('prep-input-modified');
  } else {
    input.classList.remove('prep-input-modified');
  }

  const refMl = (prepMl > 0) ? prepMl : theoreticalMl;
  _renderReagents(outEl, sol, refMl);
}

function _renderReagents(outEl, sol, volumeMl) {
  if (!outEl) return;
  if (!volumeMl || volumeMl <= 0 || !sol.volume_per_batch_ml) {
    outEl.innerHTML = sol.ingredients && sol.ingredients.length
      ? '<span class="reagent-placeholder">조제량 정보 없음</span>'
      : '<span class="reagent-placeholder">-</span>';
    return;
  }
  if (!sol.ingredients || !sol.ingredients.length) {
    outEl.innerHTML = '<span class="reagent-placeholder">-</span>';
    return;
  }

  const scale = volumeMl / sol.volume_per_batch_ml;
  const items = sol.ingredients.map(ing => {
    const scaledAmt = roundSig(ing.amount * scale, 4);
    const trackingHtml = ing.tracking_no
      ? `<span class="reagent-tracking">${esc(ing.tracking_no)}</span>`
      : '';
    return `<div class="reagent-item">
      <span class="reagent-name">${esc(ing.name)}${trackingHtml}</span>
      <span class="reagent-amount">${scaledAmt} ${esc(ing.unit)}</span>
    </div>`;
  });

  outEl.innerHTML = items.join('');
}

// ── 초자 테이블 ───────────────────────────────────────────
function renderGlasswareTable(glassware) {
  const tbody = document.querySelector('#tbl-glassware tbody');
  tbody.innerHTML = '';

  if (!glassware.length) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-msg">초자 정보 없음</td></tr>';
    return;
  }

  const sorted = [...glassware].sort((a, b) =>
    (a.type + a.size).localeCompare(b.type + b.size)
  );
  sorted.forEach(g => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${esc(g.type)}</td>
      <td>${esc(g.size)}</td>
      <td class="num">${g.count_per_batch}개</td>
      <td class="num bold">${g.total_count}개</td>
    `;
    tbody.appendChild(tr);
  });
}


// ── 피펫/메스 실린더 테이블 ──────────────────────────────
function renderPipetteTable(pipettes) {
  const block = document.getElementById('block-pipettes');
  const tbody = document.querySelector('#tbl-pipettes tbody');
  tbody.innerHTML = '';

  if (!pipettes.length) {
    block.style.display = 'none';
    return;
  }

  block.style.display = '';
  pipettes.forEach(p => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${esc(p.type)}</td>
      <td class="num bold">${p.volume_ml} mL</td>
      <td class="num">1개</td>
    `;
    tbody.appendChild(tr);
  });
}

// ── 필터 테이블 ──────────────────────────────────────────
function renderFilterTable(filters) {
  const block = document.getElementById('block-filters');
  const tbody = document.querySelector('#tbl-filters tbody');
  tbody.innerHTML = '';

  if (!filters.length) {
    block.style.display = 'none';
    return;
  }

  block.style.display = '';

  const sorted = [...filters].sort((a, b) =>
    (a.material + a.manufacturer).localeCompare(b.material + b.manufacturer)
  );

  sorted.forEach(f => {
    const kind = f.filter_type === 'membrane' ? 'Membrane filter' : 'Syringe filter';
    const size = f.size_um ? `${f.size_um} µm` : '-';
    const mat  = f.material || '-';
    const mfr  = f.manufacturer || '-';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${esc(kind)} <span class="th-hint">${esc(size)}</span></td>
      <td>${esc(mat)}</td>
      <td>${esc(mfr)}</td>
      <td class="num">${f.count_per_batch}개</td>
      <td class="num bold">${f.total_count}개</td>
    `;
    tbody.appendChild(tr);
  });
}

// ── 관리 패널 ────────────────────────────────────────────
function toggleAdmin() {
  const panel = document.getElementById('admin-panel');
  panel.classList.toggle('hidden');
}

async function startParse() {
  document.getElementById('btn-parse').disabled = true;
  document.getElementById('parse-status').textContent = '파싱 시작 중…';
  document.getElementById('parse-log').classList.remove('hidden');
  document.getElementById('parse-log').textContent = '';
  try {
    await fetch('/api/parse', { method: 'POST' });
    if (_pollTimer) clearInterval(_pollTimer);
    _pollTimer = setInterval(pollParseStatus, 1500);
  } catch (e) {
    document.getElementById('parse-status').textContent = '오류: ' + e.message;
    document.getElementById('btn-parse').disabled = false;
  }
}

async function pollParseStatus() {
  try {
    const res  = await fetch('/api/parse/status');
    const data = await res.json();
    const logEl = document.getElementById('parse-log');

    if (data.log && data.log.length) {
      logEl.textContent = data.log.join('\n');
      logEl.scrollTop = logEl.scrollHeight;
    }

    if (!data.parsing) {
      clearInterval(_pollTimer);
      document.getElementById('btn-parse').disabled = false;
      document.getElementById('parse-status').textContent =
        data.error ? '오류: ' + data.error : `완료: ${data.product_count}개 제품`;
      if (!data.error) await loadProducts();
    } else {
      document.getElementById('parse-status').innerHTML =
        '<span class="spinning">⟳</span> 파싱 중…';
    }
  } catch (_) {}
}

// ── 유틸 ─────────────────────────────────────────────────
function show(id) { const el = document.getElementById(id); el.classList.remove('hidden'); el.style.display = ''; }
function hide(id) { const el = document.getElementById(id); el.style.display = 'none'; }

function esc(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function fmt(n) {
  return Number(n) % 1 === 0 ? String(n) : Number(n).toFixed(1);
}

function roundSig(n, digits) {
  if (n === 0) return 0;
  const d = Math.ceil(Math.log10(Math.abs(n)));
  const p = digits - d;
  const m = Math.pow(10, p);
  return Math.round(n * m) / m;
}
