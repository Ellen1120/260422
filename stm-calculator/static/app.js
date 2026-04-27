'use strict';

// ── 상태 ─────────────────────────────────────────────────
let _products = [];
// strengthConfigs: [{strength: str, testItemBatches: [{name, batch_count}]}]
let _selected = { productId: null, strengthConfigs: [] };
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
  _selected.productId = productId || null;
  _selected.strengthConfigs = [];

  hide('row-strength'); hide('row-tests');
  document.getElementById('btn-calc').disabled = true;
  hide('result-section');

  if (!productId) return;
  const product = _products.find(p => p.id === productId);
  if (!product) return;

  document.getElementById('hint-stm').textContent = `${product.name}  (${product.stm_file})`;

  const selStr   = document.getElementById('sel-strength');
  const lblFixed = document.getElementById('lbl-strength-fixed');
  const multiDiv = document.getElementById('strength-multi');
  selStr.style.display = 'none';

  if (product.strengths.length === 1) {
    // 단일 함량: 자동 선택 후 시험항목 표시
    _selected.strengthConfigs = [{ strength: product.strengths[0], testItemBatches: [] }];
    lblFixed.textContent = product.strengths[0];
    lblFixed.style.display = '';
    multiDiv.style.display = 'none';
    show('row-strength');
    _renderSingleTestItems(product);
  } else {
    // 복수 함량: 함량별 독립 시험항목 카드
    lblFixed.style.display = 'none';
    multiDiv.innerHTML = '';
    product.strengths.forEach(s => {
      const div = document.createElement('div');
      div.className = 'str-cfg';
      div.dataset.strength = s;
      div.innerHTML = `
        <div class="str-cfg-hdr" onclick="toggleStrCfg(this)">
          <span class="str-cfg-check"></span>
          <span class="str-cfg-name">${esc(s)}</span>
        </div>
        <div class="str-cfg-body hidden">
          <div class="checkbox-group str-test-items"></div>
        </div>
      `;
      const itemsDiv = div.querySelector('.str-test-items');
      product.test_items.forEach(name => {
        const lbl = document.createElement('label');
        lbl.innerHTML = `
          <input type="checkbox" value="${esc(name)}" onchange="onStrTestChange(this)">
          <span>${esc(name)}</span>
          <span class="batch-inline hidden">
            <input type="number" class="batch-count-input" value="1" min="1" max="100"
                   oninput="onStrBatchInput(this)" onclick="event.stopPropagation()">
            배치
          </span>
        `;
        itemsDiv.appendChild(lbl);
      });
      multiDiv.appendChild(div);
    });
    multiDiv.style.display = '';
    show('row-strength');
    hide('row-tests');
  }
}

// ── 복수 함량 카드 토글 ───────────────────────────────────
function toggleStrCfg(hdrEl) {
  const cfgEl   = hdrEl.closest('.str-cfg');
  const strength = cfgEl.dataset.strength;
  const body    = cfgEl.querySelector('.str-cfg-body');
  const isOn    = cfgEl.classList.toggle('selected');
  body.classList.toggle('hidden', !isOn);
  if (isOn) {
    if (!_selected.strengthConfigs.find(sc => sc.strength === strength)) {
      _selected.strengthConfigs.push({ strength, testItemBatches: [] });
    }
  } else {
    _selected.strengthConfigs = _selected.strengthConfigs.filter(sc => sc.strength !== strength);
  }
  syncCalcButton();
}

// 복수 함량 시험항목 체크 변경
function onStrTestChange(cb) {
  const batchSpan = cb.closest('label').querySelector('.batch-inline');
  if (cb.checked) batchSpan.classList.remove('hidden');
  else { batchSpan.classList.add('hidden'); batchSpan.querySelector('input').value = 1; }
  _syncStrCfg(cb.closest('.str-cfg'));
  syncCalcButton();
}

// 복수 함량 배치 수 변경
function onStrBatchInput(inp) {
  _syncStrCfg(inp.closest('.str-cfg'));
}

function _syncStrCfg(cfgEl) {
  const strength = cfgEl.dataset.strength;
  const checked  = cfgEl.querySelectorAll('input[type="checkbox"]:checked');
  const tibs = [...checked].map(cb => ({
    name: cb.value,
    batch_count: parseInt(cb.closest('label').querySelector('.batch-count-input').value) || 1,
  }));
  const cfg = _selected.strengthConfigs.find(sc => sc.strength === strength);
  if (cfg) cfg.testItemBatches = tibs;
}

// ── Step 2: 단일 함량 시험항목 렌더 ──────────────────────
function _renderSingleTestItems(product) {
  const container = document.getElementById('test-checkboxes');
  container.innerHTML = '';
  product.test_items.forEach(name => {
    const lbl = document.createElement('label');
    lbl.innerHTML = `
      <input type="checkbox" value="${esc(name)}" onchange="onSingleTestChange(this)">
      <span>${esc(name)}</span>
      <span class="batch-inline hidden">
        <input type="number" class="batch-count-input" value="1" min="1" max="100"
               oninput="_syncSingleCfg()" onclick="event.stopPropagation()">
        배치
      </span>
    `;
    container.appendChild(lbl);
  });
  show('row-tests');
}

function onSingleTestChange(cb) {
  const batchSpan = cb.closest('label').querySelector('.batch-inline');
  if (cb.checked) batchSpan.classList.remove('hidden');
  else { batchSpan.classList.add('hidden'); batchSpan.querySelector('input').value = 1; }
  _syncSingleCfg();
  syncCalcButton();
}

function _syncSingleCfg() {
  const checked = document.querySelectorAll('#test-checkboxes input[type="checkbox"]:checked');
  const tibs = [...checked].map(cb => ({
    name: cb.value,
    batch_count: parseInt(cb.closest('label').querySelector('.batch-count-input').value) || 1,
  }));
  if (_selected.strengthConfigs.length > 0) {
    _selected.strengthConfigs[0].testItemBatches = tibs;
  }
}

function syncCalcButton() {
  const ok = _selected.strengthConfigs.some(sc => sc.testItemBatches.length > 0);
  document.getElementById('btn-calc').disabled = !ok;
}

// ── 이론량 계산 ──────────────────────────────────────────
async function calculate() {
  const btn = document.getElementById('btn-calc');
  btn.disabled = true;
  btn.textContent = '계산 중…';

  const validConfigs = _selected.strengthConfigs.filter(sc => sc.testItemBatches.length > 0);

  try {
    const res = await fetch('/api/calculate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        product_id: _selected.productId,
        strength_configs: validConfigs.map(sc => ({
          strength:   sc.strength,
          test_items: sc.testItemBatches,
        })),
      }),
    });
    if (!res.ok) {
      const err = await res.json();
      const detail = err.detail;
      const msg = Array.isArray(detail)
        ? detail.map(d => d.msg || JSON.stringify(d)).join('; ')
        : (detail || '알 수 없는 오류');
      alert('오류: ' + msg);
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
  const cfgs = data.strength_configs || [];
  let subtitleStr;
  if (cfgs.length === 0) {
    subtitleStr = data.strength || '';
  } else if (cfgs.length === 1) {
    const sc = cfgs[0];
    const testsPart = (sc.test_items || []).map(ti => `${ti.name} ${ti.batch_count}배치`).join(', ');
    subtitleStr = `${sc.strength}  |  ${testsPart}`;
  } else {
    subtitleStr = cfgs.map(sc => {
      const testsPart = (sc.test_items || []).map(ti => `${ti.name} ${ti.batch_count}배치`).join(', ');
      return `${sc.strength}: ${testsPart}`;
    }).join('  /  ');
  }
  document.getElementById('result-subtitle').textContent =
    `${data.product_name}  ·  ${subtitleStr}`;

  const docNoEl = document.getElementById('result-docno');
  if (data.doc_no) {
    docNoEl.textContent = data.doc_no;
    docNoEl.style.display = '';
  } else {
    docNoEl.style.display = 'none';
  }

  renderSolutionTable(data.solutions, data.dissolution_medium);
  renderGlasswareTable(data.glassware, data.strength_configs);
  renderPipetteTable(data.pipettes || []);
  renderStandardsTable(data.standards || []);
  renderColumnTable(data.columns || []);
  renderFilterTable(data.filters || [], data.strength_configs);

  show('result-section');
  document.getElementById('result-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// 솔루션 표에서 제외할 이름 패턴 (영문/한글 공통)
const _HIDDEN_SOL = /(?:standard\s+(?:stock\s+)?solution|sample\s+solution|표준원액|표준액|검액)/i;

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
function renderGlasswareTable(glassware, strengthConfigs) {
  const tbody = document.querySelector('#tbl-glassware tbody');
  tbody.innerHTML = '';
  const summaryDiv = document.getElementById('gw-summary');
  summaryDiv.innerHTML = '';

  if (!glassware.length) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-msg">초자 정보 없음</td></tr>';
    return;
  }

  // ── 요약 집계: (type, size) 기준 합산 ──
  const sumMap = new Map();
  for (const g of glassware) {
    const key = `${g.type}||${g.size}`;
    if (!sumMap.has(key)) sumMap.set(key, { type: g.type, size: g.size, total: 0 });
    sumMap.get(key).total += g.total_count;
  }
  const sumSorted = [...sumMap.values()].sort((a, b) => {
    const t = (a.type || '').localeCompare(b.type || '');
    return t !== 0 ? t : (a.size || '').localeCompare(b.size || '');
  });

  // 요약 테이블 렌더링
  const sumTbl = document.createElement('table');
  sumTbl.id = 'tbl-glassware-summary';
  sumTbl.innerHTML = `
    <thead><tr><th>초자명</th><th>규격</th><th>합계 수량</th></tr></thead>
    <tbody>
      ${sumSorted.map(s => `
        <tr>
          <td>${esc(s.type)}</td>
          <td>${esc(s.size)}</td>
          <td class="num bold">${s.total}개</td>
        </tr>
      `).join('')}
    </tbody>
  `;
  summaryDiv.appendChild(sumTbl);

  // 상세 내역 토글 버튼
  const tblGw = document.getElementById('tbl-glassware');
  const toggleBtn = document.createElement('button');
  toggleBtn.className = 'btn btn-outline gw-detail-toggle';
  toggleBtn.style.cssText = 'margin-top:12px; font-size:0.85rem;';
  toggleBtn.textContent = '▶ 상세 내역 보기';
  summaryDiv.appendChild(toggleBtn);

  // 상세 테이블 소제목
  const detailLabel = document.createElement('p');
  detailLabel.className = 'block-desc';
  detailLabel.style.cssText = 'margin-top:16px; margin-bottom:6px; font-weight:600; color:var(--gray-6); display:none;';
  detailLabel.textContent = '상세 내역 (용액 조제별)';
  summaryDiv.appendChild(detailLabel);

  tblGw.style.display = 'none';
  toggleBtn.onclick = () => {
    const isHidden = tblGw.style.display === 'none';
    tblGw.style.display = isHidden ? '' : 'none';
    detailLabel.style.display = isHidden ? '' : 'none';
    toggleBtn.textContent = isHidden ? '▼ 상세 내역 숨기기' : '▶ 상세 내역 보기';
  };

  // ── 상세 테이블 ──
  // 표준원액(stock) → 표준액(standard) → 검액(sample) → 기타 순서로 정렬
  const _gwPrepPriority = src => {
    const s = (src || '').toLowerCase();
    if (/stock/.test(s)) return 0;
    if (/standard/.test(s)) return 1;
    if (/sample/.test(s)) return 2;
    return 3;
  };
  const multiStrength = (strengthConfigs || []).length > 1;
  const sorted = [...glassware].sort((a, b) => {
    const pa = _gwPrepPriority(a.source_prep), pb = _gwPrepPriority(b.source_prep);
    if (pa !== pb) return pa - pb;
    const sp = (a.source_prep || '').localeCompare(b.source_prep || '');
    if (sp !== 0) return sp;
    const st = (a.test_item || '').localeCompare(b.test_item || '');
    if (st !== 0) return st;
    const ss = (a.strength || '').localeCompare(b.strength || '');
    if (ss !== 0) return ss;
    return (a.type + a.size).localeCompare(b.type + b.size);
  });
  sorted.forEach(g => {
    const hints = [];
    if (g.test_item) hints.push(esc(g.test_item));
    if (multiStrength && g.strength) hints.push(`(${esc(g.strength)})`);
    const hintHtml = hints.length ? ` <span class="th-hint">· ${hints.join(' ')}</span>` : '';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="sources-cell">${esc(g.source_prep || '')}${hintHtml}</td>
      <td>${esc(g.type)}</td>
      <td>${esc(g.size)}</td>
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

// ── 표준품 테이블 ────────────────────────────────────────
function renderStandardsTable(standards) {
  const block = document.getElementById('block-standards');
  const tbody = document.querySelector('#tbl-standards tbody');
  tbody.innerHTML = '';

  if (!standards || !standards.length) {
    block.style.display = 'none';
    return;
  }

  block.style.display = '';
  standards.forEach(s => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="bold">${esc(s.name || '')}</td>
      <td>${esc(s.consumable_type || '')}</td>
      <td>${esc(s.location || '')}</td>
    `;
    tbody.appendChild(tr);
  });
}

// ── 컬럼 테이블 ──────────────────────────────────────────
function renderColumnTable(columns) {
  const block = document.getElementById('block-columns');
  const container = document.getElementById('col-groups');
  container.innerHTML = '';
  if (!columns.length) { block.style.display = 'none'; return; }
  block.style.display = '';

  // Name + Spec 기준으로 그룹핑
  const groups = [];
  const keyMap = new Map();
  columns.forEach(c => {
    const key = `${c.name}||${c.spec}`;
    if (!keyMap.has(key)) {
      const g = { name: c.name, spec: c.spec, rows: [] };
      keyMap.set(key, g);
      groups.push(g);
    }
    keyMap.get(key).rows.push(c);
  });

  groups.forEach(g => {
    const header = document.createElement('p');
    header.style.cssText = 'margin:12px 0 6px; font-weight:600; color:#1e40af;';
    header.textContent = `${g.name}  ·  ${g.spec}`;
    container.appendChild(header);

    const table = document.createElement('table');
    table.innerHTML = `
      <thead>
        <tr>
          <th>Column No.</th>
          <th>Storage Location</th>
          <th>Test Item</th>
          <th>비고</th>
        </tr>
      </thead>
      <tbody></tbody>
    `;
    const tbody = table.querySelector('tbody');
    g.rows.forEach(c => {
      const tr = document.createElement('tr');
      const remarkCell = c.remark
        ? `<td style="color:#b45309; font-weight:600;">${esc(c.remark)}</td>`
        : '<td></td>';
      tr.innerHTML = `
        <td class="bold">${esc(c.column_no)}</td>
        <td>${esc(c.location)}</td>
        <td>${esc(c.test_item || '')}</td>
        ${remarkCell}
      `;
      tbody.appendChild(tr);
    });
    container.appendChild(table);
  });
}

// ── 필터 테이블 ──────────────────────────────────────────
function renderFilterTable(filters, strengthConfigs) {
  const block = document.getElementById('block-filters');
  const tbody = document.querySelector('#tbl-filters tbody');
  tbody.innerHTML = '';

  if (!filters.length) {
    block.style.display = 'none';
    return;
  }

  block.style.display = '';
  const multiStrength = (strengthConfigs || []).length > 1;

  const sorted = [...filters].sort((a, b) => {
    const sp = (a.source_prep || '').localeCompare(b.source_prep || '');
    if (sp !== 0) return sp;
    const st = (a.test_item || '').localeCompare(b.test_item || '');
    if (st !== 0) return st;
    const ss = (a.strength || '').localeCompare(b.strength || '');
    if (ss !== 0) return ss;
    return (a.material + a.manufacturer).localeCompare(b.material + b.manufacturer);
  });

  sorted.forEach(f => {
    const isFalcon = f.filter_type === 'centrifuge';
    const kind = isFalcon ? '원심분리 팔콘'
               : f.filter_type === 'membrane' ? 'Membrane filter'
               : 'Syringe filter';
    const size = isFalcon ? '50 mL' : (f.size_um ? `${f.size_um} µm` : '-');
    const mat  = f.material || '-';
    const mfr  = f.manufacturer || '-';
    const hints = [];
    if (f.test_item) hints.push(esc(f.test_item));
    if (multiStrength && f.strength) hints.push(`(${esc(f.strength)})`);
    const hintHtml = hints.length ? ` <span class="th-hint">· ${hints.join(' ')}</span>` : '';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="sources-cell">${esc(f.source_prep || '')}${hintHtml}</td>
      <td>${esc(kind)} <span class="th-hint">${esc(size)}</span></td>
      <td>${esc(mat)}</td>
      <td>${esc(mfr)}</td>
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
