'use strict';

// ── 상태 ─────────────────────────────────────────────────
let _products = [];
let _selected = { productId: null, strength: null };
let _calcResult = null;
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
    opt.textContent = p.name;
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

  document.getElementById('hint-stm').textContent = product.stm_file;

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

  renderDissolutionMedium(data.dissolution_medium);
  renderSolutionTable(data.solutions);
  renderGlasswareTable(data.glassware);
  renderPrepDetails(data.solutions);

  show('result-section');
  document.getElementById('result-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ── 시험액 필요량 (용출) ─────────────────────────────────
function renderDissolutionMedium(dm) {
  if (!dm || !dm.volume_per_vessel_ml) {
    hide('diss-medium-block');
    return;
  }

  const sampleMl = dm.sample_medium_ml;
  const stdMl    = dm.standard_medium_ml_once || 0;
  const totalMl  = dm.total_medium_ml;

  const parts = [`검액 ${fmt(sampleMl)} mL`];
  if (stdMl > 0) parts.push(`표준액 ${fmt(stdMl)} mL`);
  const breakdown = parts.join(' + ');

  document.getElementById('diss-medium-content').innerHTML = `
    <div class="diss-total-display">
      <span class="diss-big-total">${fmt(totalMl)} mL</span>
      <span class="diss-breakdown">(${breakdown})</span>
    </div>
  `;
  show('diss-medium-block');
}

// ── 용액 테이블 ───────────────────────────────────────────
function renderSolutionTable(solutions) {
  const tbody = document.querySelector('#tbl-solutions tbody');
  tbody.innerHTML = '';

  if (!solutions.length) {
    tbody.innerHTML = '<tr><td colspan="3" class="empty-msg">조제 정보 없음</td></tr>';
    return;
  }

  solutions.forEach((s, idx) => {
    const tr = document.createElement('tr');
    const theoretical = s.theoretical_volume_ml != null
      ? fmt(s.theoretical_volume_ml) + ' mL' : '-';

    const inputId   = `prep-input-${idx}`;
    const reagentId = `reagent-out-${idx}`;

    tr.innerHTML = `
      <td>${esc(s.solution_name)}</td>
      <td class="num bold">${theoretical}</td>
      <td>
        <div class="prep-input-cell">
          <input type="number" id="${inputId}" min="0" step="1"
                 placeholder="직접 입력 (mL)"
                 data-idx="${idx}"
                 oninput="onPrepAmountInput(this)"
                 style="width:110px" />
          <span class="unit">mL</span>
          <div class="reagent-list" id="${reagentId}"></div>
        </div>
      </td>
    `;
    tbody.appendChild(tr);

    // 이론량 기준으로 시약 자동 계산 표시
    const outEl = document.getElementById(reagentId);
    _renderReagents(outEl, s, s.theoretical_volume_ml);
  });
}

function onPrepAmountInput(input) {
  const idx   = parseInt(input.dataset.idx);
  const prepMl = parseFloat(input.value);
  const sol   = _calcResult.solutions[idx];
  const outEl = document.getElementById(`reagent-out-${idx}`);

  // 입력이 비어있으면 이론량으로 복귀
  const refMl = (prepMl > 0) ? prepMl : sol.theoretical_volume_ml;
  _renderReagents(outEl, sol, refMl);
}

function _renderReagents(outEl, sol, volumeMl) {
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

// ── 조제 절차 ─────────────────────────────────────────────
function renderPrepDetails(solutions) {
  const container = document.getElementById('prep-details');
  container.innerHTML = '';
  solutions.forEach(s => {
    if (!s.preparation_text) return;
    const card = document.createElement('div');
    card.className = 'prep-card';
    card.innerHTML = `
      <div class="prep-card-header">
        <span class="badge">${esc(s.test_item)}</span>
        <strong>${esc(s.solution_name)}</strong>
      </div>
      <div class="prep-card-body">${esc(s.preparation_text).replace(/\n/g, '<br>')}</div>
    `;
    container.appendChild(card);
  });
  if (!container.children.length) {
    container.innerHTML = '<p style="color:#94a3b8;font-size:13px">조제 절차 정보 없음</p>';
  }
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
