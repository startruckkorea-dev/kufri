// 앱 오케스트레이션 — 상태, 이벤트, 액션
import { CONFIG, loadMapping, saveMapping, loadChangeLog, saveChangeLog, redirectUri } from './config.js';
import { initAuth, getAccount, login, logout } from './auth.js';
import { telemetry } from './graph.js';
import { ctx, resolveAll, readListItems, readExcelViaWorkbook, readExcelViaDownload } from './sources.js';
import { buildDiff, autoMap, resolveKeyColumns, nameKey, normalize, display } from './diff.js';
import { applyChanges } from './sync.js';
import { readItemVersions, buildRestorePlan, buildRestorePlanFromLog, buildPointInTimePlan } from './restore.js';
import { renderSetup, renderDiff, renderChanges, renderResult, renderRollback, renderBench, esc } from './ui.js';

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

  // ⑤ 롤백
  rollbackMode: 'log', // 'log' = 적용 이력 기준, 'pit' = 시점 기준 전체 스캔
  rollbackEntryIdx: 0,
  pitCutoff: startOfTodayValue(),
  rollback: null,
  rollbackResult: null,

  changeLog: loadChangeLog(),
  readBench: null,
  telemetry: telemetry.calls,
};

const view = document.getElementById('view');
const $ = (sel) => document.querySelector(sel);

/** datetime-local 입력용 값 (로컬 시간대 기준 오늘 00:00) */
function startOfTodayValue() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
}

// ------------------------------------------------------------------ 렌더

function render() {
  const map = {
    setup: renderSetup,
    diff: renderDiff,
    changes: renderChanges,
    result: renderResult,
    rollback: renderRollback,
    bench: renderBench,
  };
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
  for (const root of document.querySelectorAll('.cmp')) {
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

    logApply(rows, result);
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

/** 적용 / 롤백 공통 이력 기록 */
function logApply(rows, result, note = null) {
  const entry = {
    at: new Date().toISOString(),
    user: state.account?.username || '',
    note,
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
        // 롤백 행은 beforeText 대신 currentText 를 쓴다
        cells: (src?.diffCells || []).map((c) => ({
          label: c.label,
          before: c.beforeText ?? c.currentText,
          after: c.afterText,
        })),
      };
    }),
  };
  state.changeLog = [entry, ...state.changeLog];
  saveChangeLog(state.changeLog);
}

/** 롤백 대상 컬럼 — 매핑된 컬럼(키 제외) */
function restoreCols() {
  const colOf = (n) => ctx.columns.find((c) => c.name === n);
  return state.mapping.pairs
    .filter((p) => p.excel && p.list)
    .map((p) => ({ field: p.list, label: colOf(p.list)?.displayName || p.list, col: colOf(p.list) }));
}

/** 버전 기록에서 복원 지점을 찾아 계획을 세운다. 여기서는 아직 아무것도 쓰지 않는다. */
async function actRbPrepare() {
  await guard('버전 기록 조회 중', async () => {
    const entry = state.changeLog[state.rollbackEntryIdx];
    if (!entry) throw new Error('선택한 적용 이력을 찾을 수 없습니다.');
    if (entry.dryRun) throw new Error('드라이런은 실제로 쓰지 않았으므로 되돌릴 것이 없습니다.');

    const cols = restoreCols();
    const itemIds = entry.items.filter((i) => i.ok).map((i) => i.itemId);
    const { versions, errors, ms } = await readItemVersions(itemIds, { concurrency: state.writeOpts.concurrency });

    const usable = [...versions.values()].filter((v) => v.length >= 2).length;
    let plan;
    let source = 'version';

    if (usable === 0) {
      // 버전 관리가 꺼져 있거나 이전 버전이 없다 → 적용 이력의 '전' 값으로 대체
      const list = await readListItems(mappedFieldNames());
      plan = buildRestorePlanFromLog(entry, list.items, cols, state.mapping.keyList);
      source = 'log';
    } else {
      plan = buildRestorePlan(entry, versions, cols, entry.user);
    }

    state.rollback = { plan, source, versionMs: ms, at: new Date().toLocaleString('ko-KR'), errors };
    state.rollbackResult = null;
  });
}

/**
 * 시점 기준 스캔 — 적용 이력이 없어도 되돌릴 수 있는 경로.
 * 리스트 전체 항목의 버전 기록을 훑어 기준 시각 직전 버전을 찾는다. 아직 쓰지 않는다.
 */
async function actRbScan() {
  await guard('전체 항목 버전 기록 스캔 중', async () => {
    const cutoffMs = Date.parse(state.pitCutoff);
    if (!Number.isFinite(cutoffMs)) throw new Error('기준 시각이 올바르지 않습니다.');

    const cols = restoreCols();
    if (!cols.length) throw new Error('매핑된 컬럼이 없습니다. ① 탭에서 매핑을 먼저 설정하세요.');

    const list = await readListItems(mappedFieldNames());
    const { versions, errors, ms } = await readItemVersions(
      list.items.map((i) => i.id),
      { concurrency: state.writeOpts.concurrency }
    );

    const plan = buildPointInTimePlan(
      list.items,
      versions,
      cols,
      cutoffMs,
      state.mapping.keyList,
      state.account?.username
    );

    state.rollback = {
      plan,
      source: 'pit',
      versionMs: list.ms + ms,
      at: new Date().toLocaleString('ko-KR'),
      errors,
    };
    state.rollbackResult = null;
  });
}

/** 복원 실행 — 일반 적용과 같은 $batch 경로를 타므로 쓰기 성능도 함께 측정된다. */
async function actRbApply() {
  const rows = state.rollback?.plan.rows.filter((r) => r.diffCells.length) || [];
  if (!rows.length) return;
  if (
    !confirm(
      `SharePoint 리스트 '${ctx.listTitle}' 의 ${rows.length}개 항목을 적용 이전 값으로 되돌립니다.\n` +
        `이 복원 작업도 새 버전을 만듭니다 (원래 값이 사라지지는 않습니다).\n\n진행할까요?`
    )
  )
    return;

  await guard('복원 쓰기 중', async () => {
    const result = await applyChanges(rows, {
      batchSize: state.writeOpts.batchSize,
      concurrency: state.writeOpts.concurrency,
      dryRun: false,
    });
    state.rollbackResult = result;
    logApply(rows, result, 'rollback');

    // 복원 결과도 ④ 탭에서 셀 단위로 검증한다
    state.lastApply = result;
    state.lastAppliedRows = rows;
    state.postApply = await buildPostApply(rows, result);
    state.resultScope = 'changed';
    state.tab = 'result';
  });
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
  if (state.tab === 'rollback') ensureRollbackEntry();
  render();
});

/** 롤백 선택값 보정 — 이력이 없으면 시점 기준 방식으로 넘긴다 */
function ensureRollbackEntry() {
  const idx = state.changeLog.findIndex((x) => !x.dryRun);
  if (idx < 0) {
    state.rollbackMode = 'pit'; // 되돌릴 이력이 없으면 시점 기준밖에 방법이 없다
    return;
  }
  const cur = state.changeLog[state.rollbackEntryIdx];
  if (!cur || cur.dryRun) state.rollbackEntryIdx = idx;
}

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
    case 'rb-prepare':
      return actRbPrepare();
    case 'rb-scan':
      return actRbScan();
    case 'rb-apply':
      return actRbApply();

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

    case 'rb-entry':
      state.rollbackEntryIdx = Number(el.value);
      state.rollback = null;
      state.rollbackResult = null;
      return render();
    case 'rb-mode':
      state.rollbackMode = el.value;
      state.rollback = null;
      state.rollbackResult = null;
      return render();
    case 'rb-cutoff':
      state.pitCutoff = el.value; // 입력 포커스를 유지하려고 재렌더하지 않는다
      return;
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
