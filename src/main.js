// 앱 오케스트레이션 — 상태, 이벤트, 액션
import { CONFIG, loadMapping, saveMapping, loadChangeLog, saveChangeLog, redirectUri } from './config.js';
import { initAuth, getAccount, login, logout } from './auth.js';
import { telemetry } from './graph.js';
import { ctx, resolveAll, readListItems, readExcelViaWorkbook, readExcelViaDownload } from './sources.js';
import { buildDiff, autoMap, resolveKeyColumns, nameKey, normalize, display } from './diff.js';
import { applyChanges } from './sync.js';
import { renderSetup, renderDiff, renderChanges, renderResult, renderBench, esc } from './ui.js';

const state = {
  account: null,
  tab: 'setup',
  busy: false,
  error: null,
  connected: false,
  keyError: null,

  sheets: [],
  sheet: null,
  excelMethod: 'workbook',
  excelHeaders: [],
  excelCache: null, // 마지막으로 읽은 {headers, rows, ...}

  mapping: loadMapping() || { keyExcel: '', keyList: '', pairs: [] },

  diff: null,
  selected: new Set(),

  writeOpts: { batchSize: CONFIG.defaults.batchSize, concurrency: CONFIG.defaults.concurrency, dryRun: true },
  applying: false,
  progress: { done: 0, total: 0 },
  lastApply: null,
  lastAppliedRows: null,

  // ④ 적용 결과 그리드
  postApply: null,
  resultScope: 'changed',

  changeLog: loadChangeLog(),
  readBench: null,
  telemetry: telemetry.calls,
};

const view = document.getElementById('view');
const $ = (sel) => document.querySelector(sel);

// ------------------------------------------------------------------ 렌더

function render() {
  const map = { setup: renderSetup, diff: renderDiff, changes: renderChanges, result: renderResult, bench: renderBench };
  const err = state.error
    ? `<div class="alert err"><b>오류</b> — ${esc(state.error.message || state.error)}
       ${state.error.detail ? `<pre>${esc(String(state.error.detail).slice(0, 800))}</pre>` : ''}</div>`
    : '';
  const busy = state.busy ? `<div class="alert info">처리 중… <span class="muted">${esc(state.busyLabel || '')}</span></div>` : '';
  view.innerHTML = err + busy + map[state.tab](state);

  document.querySelectorAll('#tabs button').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === state.tab)));
  $('#who').textContent = state.account ? `${state.account.name || ''} <${state.account.username}>` : '로그인 필요';
  $('#btn-login').hidden = !!state.account;
  $('#btn-logout').hidden = !state.account;

  wireCompareSync();
}

/**
 * 비교 그리드의 스크롤 동기화.
 * - 세로: 키 열 + 전 패널 + 후 패널 3개를 함께 이동 (행이 항상 같은 줄에 놓이도록)
 * - 가로: 전/후 두 패널만 함께 이동 (같은 컬럼을 나란히 두고 비교하도록)
 * render() 가 innerHTML 을 갈아끼우므로 매 렌더마다 다시 연결한다.
 */
function wireCompareSync() {
  const root = document.getElementById('cmp-grid');
  if (!root) return;
  const panes = [...root.querySelectorAll('.cmp-pane')];
  const all = [root.querySelector('.cmp-key'), ...panes].filter(Boolean);
  let lock = false;

  for (const el of all) {
    el.addEventListener('scroll', () => {
      if (lock) return;
      lock = true;
      for (const other of all) if (other !== el) other.scrollTop = el.scrollTop;
      if (panes.includes(el)) for (const p of panes) if (p !== el) p.scrollLeft = el.scrollLeft;
      requestAnimationFrame(() => {
        lock = false;
      });
    });
  }
}

/** 체크박스/옵션 변경 시 전체 재렌더 없이 툴바만 갱신 (details 열림 상태 보존) */
function syncToolbar() {
  if (!state.diff) return;
  const selCount = state.diff.changed.filter((r) => state.selected.has(r.key)).length;
  document.querySelectorAll('[data-act="sel"]').forEach((cb) => (cb.checked = state.selected.has(cb.dataset.key)));
  const label = document.querySelector('[data-sel-count]');
  if (label) label.textContent = `${selCount} / ${state.diff.changed.length} 선택됨`;
  const btn = document.querySelector('[data-act="apply"]');
  if (btn) {
    btn.disabled = state.applying || !selCount;
    btn.textContent = state.writeOpts.dryRun ? '드라이런 실행' : `선택 ${selCount}건 SharePoint 에 적용`;
  }
}

async function guard(label, fn) {
  state.busy = true;
  state.busyLabel = label;
  state.error = null;
  render();
  try {
    await fn();
  } catch (e) {
    console.error(e);
    state.error = e;
  } finally {
    state.busy = false;
    state.busyLabel = '';
    render();
  }
}

// ------------------------------------------------------------------ 액션

async function actConnect() {
  await guard('사이트 · 리스트 · 파일 리졸브 중', async () => {
    await resolveAll();
    const wb = await readExcelViaWorkbook(state.sheet);
    state.sheets = wb.sheets;
    state.sheet = wb.sheet;
    state.excelHeaders = wb.headers;
    state.excelCache = wb;
    state.connected = true;

    // 키 컬럼(CONFIG.keyColumn)은 사용자가 고르지 않는다 — 양쪽에서 이름으로 자동 매칭.
    const { keyExcel, keyList } = resolveKeyColumns(wb.headers, ctx.columns, CONFIG.keyColumn);
    state.mapping.keyExcel = keyExcel;
    state.mapping.keyList = keyList;
    state.keyError = keyExcel && keyList
      ? null
      : `키 컬럼 '${CONFIG.keyColumn}' 을 ` +
        [!keyExcel && `Excel 시트('${wb.sheet}')`, !keyList && `리스트('${ctx.listTitle}')`]
          .filter(Boolean)
          .join(' 과 ') +
        ` 에서 찾지 못했습니다. 아래에서 직접 지정하세요.`;

    const exclude = { keyExcel, keyList };
    if (!state.mapping.pairs.length) {
      state.mapping.pairs = autoMap(wb.headers, ctx.columns.filter((c) => !c.readOnly), exclude);
    }
    // 키는 매칭 기준이므로 갱신 대상에서 제외
    state.mapping.pairs = state.mapping.pairs.filter(
      (p) => !(keyExcel && nameKey(p.excel) === nameKey(keyExcel)) && p.list !== keyList
    );
    saveMapping(state.mapping);
  });
}

function mappedFieldNames() {
  return [state.mapping.keyList, ...state.mapping.pairs.map((p) => p.list)].filter(Boolean);
}

async function readExcel() {
  const m = state.excelMethod;
  if (m === 'workbook') return { primary: await readExcelViaWorkbook(state.sheet), both: null };
  if (m === 'download') return { primary: await readExcelViaDownload(state.sheet), both: null };
  const wb = await readExcelViaWorkbook(state.sheet);
  const dl = await readExcelViaDownload(state.sheet);
  return { primary: wb, both: { wb, dl } };
}

async function actRunDiff() {
  await guard('리스트 · Excel 읽는 중', async () => {
    const list = await readListItems(mappedFieldNames());
    const { primary: excel, both } = await readExcel();
    state.excelCache = excel;
    state.excelHeaders = excel.headers;

    state.diff = buildDiff(excel, list, state.mapping);
    state.selected = new Set(state.diff.changed.map((r) => r.key));
    state.lastApply = null;
    state.lastRead = { list, excel, both };
    state.tab = 'diff';
    saveMapping(state.mapping);
  });
}

async function actApply() {
  const rows = state.diff.changed.filter((r) => state.selected.has(r.key));
  if (!rows.length) return;
  if (!state.writeOpts.dryRun && !confirm(`SharePoint 리스트 '${ctx.listTitle}' 의 ${rows.length}개 항목을 실제로 갱신합니다. 진행할까요?`))
    return;

  state.applying = true;
  state.progress = { done: 0, total: Math.ceil(rows.length / state.writeOpts.batchSize) };
  state.error = null;
  render();

  try {
    const result = await applyChanges(rows, {
      ...state.writeOpts,
      onProgress: (done, total) => {
        state.progress = { done, total };
        const p = document.querySelector('progress');
        if (p) {
          p.value = done;
          p.max = total;
        }
      },
    });
    state.lastApply = result;

    const entry = {
      at: new Date().toISOString(),
      user: state.account?.username || '',
      dryRun: result.dryRun,
      okCount: result.okCount,
      failCount: result.failCount,
      wallMs: result.wallMs,
      itemsPerSec: result.itemsPerSec,
      batchSize: result.batchSize,
      concurrency: result.concurrency,
      items: result.results.map((r) => {
        const src = rows.find((x) => x.key === r.key);
        return {
          key: r.key,
          itemId: r.itemId,
          ok: r.ok,
          status: r.status,
          cells: (src?.diffCells || []).map((c) => ({ label: c.label, before: c.beforeText, after: c.afterText })),
        };
      }),
    };
    state.changeLog = [entry, ...state.changeLog];
    saveChangeLog(state.changeLog);
    state.lastAppliedRows = rows;

    // 실제로 쓴 경우에만 SharePoint 를 다시 읽어 반영 결과를 검증한다.
    if (!result.dryRun && result.okCount) {
      state.postApply = await buildPostApply(rows, result);
      state.resultScope = 'changed';
      state.tab = 'result';
    }
  } catch (e) {
    console.error(e);
    state.error = e;
  } finally {
    state.applying = false;
    render();
  }
}

/**
 * 적용 직후 SharePoint 를 다시 읽어, 쓴 값이 실제로 반영됐는지 셀 단위로 대조한다.
 * 반영 확인(ok) / 기대값 불일치(ng) 를 표시해 타입 변환 문제를 드러낸다.
 */
async function buildPostApply(appliedRows, result) {
  const list = await readListItems(mappedFieldNames());

  const colOf = (name) => ctx.columns.find((c) => c.name === name);
  const keyCol = colOf(state.mapping.keyList) || { type: 'text' };
  const cols = state.mapping.pairs
    .filter((p) => p.excel && p.list)
    .map((p) => ({ field: p.list, label: colOf(p.list)?.displayName || p.list, col: colOf(p.list) }));

  // 이번에 쓰기 성공한 항목의 기대값 (키 정규화 기준)
  const intendedByKey = new Map();
  for (const r of result.results.filter((x) => x.ok)) {
    const src = appliedRows.find((x) => x.key === r.key);
    if (src) intendedByKey.set(r.key, new Map(src.diffCells.map((c) => [c.field, c])));
  }

  let okCells = 0;
  let ngCells = 0;

  const rows = list.items.map((it) => {
    const nkey = normalize(it.fields[state.mapping.keyList], keyCol);
    const intended = intendedByKey.get(nkey);
    const values = {};
    const marks = {};
    const intendedText = {};

    for (const c of cols) {
      const raw = it.fields[c.field];
      values[c.field] = display(raw, c.col);
      const want = intended?.get(c.field);
      if (!want) continue;
      const match = normalize(raw, c.col) === normalize(want.after, c.col);
      marks[c.field] = match ? 'ok' : 'ng';
      intendedText[c.field] = want.afterText;
      match ? okCells++ : ngCells++;
    }

    return {
      key: display(it.fields[state.mapping.keyList], keyCol),
      itemId: it.id,
      changed: !!intended,
      values,
      marks,
      intended: intendedText,
    };
  });

  rows.sort((a, b) => Number(b.changed) - Number(a.changed)); // 변경분을 위로

  return {
    at: new Date().toLocaleString('ko-KR'),
    readMs: list.ms,
    cols: cols.map(({ field, label }) => ({ field, label })),
    rows,
    appliedCount: intendedByKey.size,
    okCells,
    ngCells,
  };
}

async function actReadBench() {
  await guard('읽기 벤치마크 실행 중 (4가지 방식 순차 측정)', async () => {
    const fields = mappedFieldNames();
    const listFull = await readListItems(null);
    const listSelect = fields.length ? await readListItems(fields) : listFull;
    const wb = await readExcelViaWorkbook(state.sheet);
    const dl = await readExcelViaDownload(state.sheet);

    state.readBench = {
      at: new Date().toLocaleString('ko-KR'),
      listFull: { ms: listFull.ms, count: listFull.count, pages: listFull.pages },
      listSelect: { ms: listSelect.ms, count: listSelect.count, pages: listSelect.pages },
      wb: { ms: wb.ms, rows: wb.rows.length },
      dl: { ms: dl.ms, rows: dl.rows.length, bytes: dl.bytes, downloadMs: dl.downloadMs, parseMs: dl.parseMs },
    };
  });
}

function exportCsv() {
  const lines = [['적용시각', '사용자', '드라이런', '키', 'ItemId', '필드', '전', '후', '결과'].join(',')];
  const q = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
  for (const e of state.changeLog)
    for (const it of e.items)
      for (const c of it.cells)
        lines.push([e.at, e.user, e.dryRun, it.key, it.itemId, c.label, c.before, c.after, it.ok ? 'OK' : `HTTP ${it.status}`].map(q).join(','));

  const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `kufri-changes-${new Date().toISOString().slice(0, 19).replace(/[:T]/g, '')}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ------------------------------------------------------------------ 이벤트

document.getElementById('tabs').addEventListener('click', (e) => {
  const b = e.target.closest('button[data-tab]');
  if (!b) return;
  state.tab = b.dataset.tab;
  render();
});

$('#btn-login').addEventListener('click', () => login());
$('#btn-logout').addEventListener('click', () => logout());

view.addEventListener('click', (e) => {
  const el = e.target.closest('[data-act]');
  if (!el) return;
  const act = el.dataset.act;
  const i = Number(el.dataset.i);

  switch (act) {
    case 'connect':
      return actConnect();
    case 'run-diff':
      return actRunDiff();
    case 'apply':
      return actApply();
    case 'run-read-bench':
      return actReadBench();

    case 'pair-add':
      state.mapping.pairs.push({ excel: '', list: '' });
      return render();
    case 'pair-del':
      state.mapping.pairs.splice(i, 1);
      saveMapping(state.mapping);
      return render();
    case 'automap':
      state.mapping.pairs = autoMap(state.excelHeaders, ctx.columns.filter((c) => !c.readOnly), state.mapping);
      saveMapping(state.mapping);
      return render();

    case 'sel':
      el.checked ? state.selected.add(el.dataset.key) : state.selected.delete(el.dataset.key);
      return syncToolbar();
    case 'sel-all':
      state.selected = new Set(state.diff.changed.map((r) => r.key));
      return syncToolbar();
    case 'sel-none':
      state.selected.clear();
      return syncToolbar();

    case 'refresh-result':
      if (!state.lastAppliedRows || !state.lastApply) return;
      return guard('적용 결과 재조회 중', async () => {
        state.postApply = await buildPostApply(state.lastAppliedRows, state.lastApply);
      });

    case 'export-csv':
      return exportCsv();
    case 'clear-log':
      if (!confirm('적용 이력을 모두 삭제할까요?')) return;
      state.changeLog = [];
      saveChangeLog([]);
      return render();
    case 'clear-telemetry':
      telemetry.reset();
      return render();
  }
});

view.addEventListener('change', (e) => {
  const el = e.target.closest('[data-act]');
  if (!el) return;
  const act = el.dataset.act;
  const i = Number(el.dataset.i);

  switch (act) {
    case 'sheet':
      state.sheet = el.value;
      return actConnect();
    case 'excel-method':
      state.excelMethod = el.value;
      return;
    case 'key-excel':
      state.mapping.keyExcel = el.value;
      saveMapping(state.mapping);
      return render();
    case 'key-list':
      state.mapping.keyList = el.value;
      saveMapping(state.mapping);
      return render();
    case 'pair-excel':
      state.mapping.pairs[i].excel = el.value;
      saveMapping(state.mapping);
      return;
    case 'pair-list':
      state.mapping.pairs[i].list = el.value;
      saveMapping(state.mapping);
      return render();

    case 'batch-size':
      state.writeOpts.batchSize = Math.min(20, Math.max(1, Number(el.value) || 20));
      el.value = state.writeOpts.batchSize;
      return;
    case 'concurrency':
      state.writeOpts.concurrency = Math.min(16, Math.max(1, Number(el.value) || 4));
      el.value = state.writeOpts.concurrency;
      return;
    case 'dry-run':
      state.writeOpts.dryRun = el.checked;
      return syncToolbar();

    case 'result-scope':
      state.resultScope = el.value;
      return render();
  }
});

// ------------------------------------------------------------------ 부트

(async () => {
  try {
    state.account = await initAuth();
  } catch (e) {
    console.error(e);
    state.error = new Error(
      `MSAL 초기화 실패: ${e.message} — 앱 등록의 SPA 리디렉션 URI 가 '${redirectUri()}' 와 일치하는지 확인하세요.`
    );
  }
  render();
})();
