// 화면 렌더링 — 순수 함수. 상태를 받아 HTML 문자열을 만든다.
import { CONFIG } from './config.js';
import { ctx, writableColumns } from './sources.js';

export const esc = (v) =>
  String(v ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

export const ms = (v) => (v == null ? '—' : v >= 1000 ? (v / 1000).toFixed(2) + ' s' : Math.round(v) + ' ms');
export const n2 = (v) => (v == null ? '—' : Number(v).toFixed(2));
export const kb = (b) => (b == null ? '—' : b >= 1048576 ? (b / 1048576).toFixed(2) + ' MB' : Math.round(b / 1024) + ' KB');

const stat = (label, value, cls = '') => `
  <div class="stat ${cls}"><div class="label">${esc(label)}</div><div class="value">${value}</div></div>`;

const alert = (kind, html) => `<div class="alert ${kind}">${html}</div>`;

/** 그리드 한 칸의 표시 폭 (전/후 두 패널의 열이 정확히 겹치도록 고정) */
const COL_W = 160;

// ------------------------------------------------------------------ ① 연결 · 매핑

export function renderSetup(s) {
  if (!s.account) {
    return `<div class="card"><div class="empty">
      상단의 <b>Microsoft 계정으로 로그인</b> 을 눌러 시작하세요.<br />
      <span class="muted">로그인한 계정에 부여된 SharePoint 권한 범위에서만 동작합니다.</span>
    </div></div>`;
  }

  const target = `
    <div class="card">
      <h2>대상 <span class="sub">src/config.js · .env 로 변경 가능</span></h2>
      <dl class="kv">
        <dt>사이트</dt><dd>${esc(CONFIG.hostname + CONFIG.sitePath)}</dd>
        <dt>리스트</dt><dd>${esc(CONFIG.listName)}${ctx.listId ? ` <span class="badge ok">연결됨</span>` : ''}</dd>
        <dt>파일 폴더</dt><dd>Shared Documents/${esc(CONFIG.fileFolder)}</dd>
        <dt>파일</dt><dd>${ctx.fileItem ? esc(ctx.fileItem.name) + ` <span class="muted">(${kb(ctx.fileItem.size)})</span>` : esc(CONFIG.fileBaseName) + ' <span class="muted">(확장자 탐색 예정)</span>'}</dd>
      </dl>
      <div class="row mt">
        <button class="btn primary" data-act="connect" ${s.busy ? 'disabled' : ''}>
          ${s.connected ? '스키마 다시 읽기' : '연결 및 스키마 로드'}
        </button>
        ${s.connected ? `<span class="badge ok">리스트 컬럼 ${ctx.columns.length}개 · Excel 헤더 ${s.excelHeaders.length}개</span>` : ''}
      </div>
    </div>`;

  if (!s.connected) return target;

  const allCols = ctx.columns;
  // 키 컬럼은 매칭 기준이므로 갱신 매핑 후보에서 뺀다.
  const cols = writableColumns().filter((c) => c.name !== s.mapping.keyList);
  const opt = (list, sel, blank = '— 선택 —') =>
    `<option value="">${blank}</option>` +
    list
      .map((c) => `<option value="${esc(c.name)}" ${c.name === sel ? 'selected' : ''}>${esc(c.displayName || c.name)} (${esc(c.type)})</option>`)
      .join('');
  const headerOpts = (list, sel) =>
    `<option value="">— 선택 —</option>` +
    list.map((h) => `<option value="${esc(h)}" ${h === sel ? 'selected' : ''}>${esc(h)}</option>`).join('');
  const optH = (sel) => headerOpts(s.excelHeaders, sel);
  const optHPair = (sel) => headerOpts(s.excelHeaders.filter((h) => h !== s.mapping.keyExcel), sel);

  const mappingRows = s.mapping.pairs
    .map(
      (p, i) => `
      <tr>
        <td><select data-act="pair-excel" data-i="${i}">${optHPair(p.excel)}</select></td>
        <td class="muted">→</td>
        <td><select data-act="pair-list" data-i="${i}">${opt(cols, p.list)}</select></td>
        <td>${p.list ? `<span class="badge">${esc(allCols.find((c) => c.name === p.list)?.type || '')}</span>` : ''}</td>
        <td><button class="btn sm danger" data-act="pair-del" data-i="${i}">삭제</button></td>
      </tr>`
    )
    .join('');

  return (
    target +
    `<div class="card">
      <h2>Excel 읽기 방식 <span class="sub">읽기 속도 비교 대상</span></h2>
      <div class="row">
        <label class="field"><span>워크시트</span>
          <select data-act="sheet">${s.sheets.map((n) => `<option ${n === s.sheet ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select>
        </label>
        <label class="field"><span>비교 시 사용할 방식</span>
          <select data-act="excel-method">
            <option value="workbook" ${s.excelMethod === 'workbook' ? 'selected' : ''}>Workbook API (Graph)</option>
            <option value="download" ${s.excelMethod === 'download' ? 'selected' : ''}>다운로드 + SheetJS</option>
            <option value="both" ${s.excelMethod === 'both' ? 'selected' : ''}>둘 다 실행 후 비교</option>
          </select>
        </label>
      </div>
    </div>

    <div class="card">
      <h2>컬럼 매핑 <span class="sub">Excel 헤더 → SharePoint 리스트 내부 컬럼명</span></h2>
      ${
        s.keyError
          ? alert('warn', `${esc(s.keyError)}`) +
            `<div class="row">
              <label class="field"><span>키 (Excel)</span><select data-act="key-excel">${optH(s.mapping.keyExcel)}</select></label>
              <label class="field"><span>키 (List)</span><select data-act="key-list">${opt(allCols, s.mapping.keyList)}</select></label>
            </div>`
          : `<div class="row">
              <span class="badge ok">키 자동 매칭됨</span>
              <span class="muted">기준 키</span> <b class="mono">${esc(CONFIG.keyColumn)}</b>
              <span class="muted">·</span>
              <span class="muted">Excel</span> <span class="mono">${esc(s.mapping.keyExcel)}</span>
              <span class="muted">→ List</span> <span class="mono">${esc(allCols.find((c) => c.name === s.mapping.keyList)?.displayName || s.mapping.keyList)}</span>
              <span class="muted">(내부명 ${esc(s.mapping.keyList)})</span>
            </div>
            <div class="row"><span class="muted">키는 매칭 기준이므로 갱신 대상에서 제외됩니다.</span></div>`
      }
      <div class="scroll mt">
        <table>
          <thead><tr><th style="width:32%">Excel 헤더</th><th style="width:24px"></th><th style="width:32%">List 컬럼</th><th>타입</th><th style="width:70px"></th></tr></thead>
          <tbody>${mappingRows || `<tr><td colspan="5" class="empty">매핑이 없습니다. 자동 매핑을 누르거나 행을 추가하세요.</td></tr>`}</tbody>
        </table>
      </div>
      <div class="row mt">
        <button class="btn" data-act="pair-add">행 추가</button>
        <button class="btn" data-act="automap">이름으로 자동 매핑</button>
        <div style="flex:1"></div>
        <button class="btn primary" data-act="run-diff"
          ${s.busy || !s.mapping.keyExcel || !s.mapping.keyList || !s.mapping.pairs.some((p) => p.excel && p.list) ? 'disabled' : ''}>
          데이터 읽기 &amp; 비교 실행
        </button>
      </div>
    </div>`
  );
}

// ------------------------------------------------------------------ ② 비교 (전/후)

export function renderDiff(s) {
  if (!s.diff) return `<div class="card"><div class="empty">먼저 <b>① 연결 · 매핑</b> 에서 비교를 실행하세요.</div></div>`;
  const d = s.diff;
  const u = d.summary;

  const warn = [];
  if (d.warnings.dupExcelKeys.length)
    warn.push(`Excel 에 중복 키 ${d.warnings.dupExcelKeys.length}건 (첫 행만 사용): <span class="mono">${esc(d.warnings.dupExcelKeys.slice(0, 5).join(', '))}</span>`);
  if (d.warnings.dupListKeys.length)
    warn.push(`List 에 중복 키 ${d.warnings.dupListKeys.length}건 (첫 항목만 갱신): <span class="mono">${esc(d.warnings.dupListKeys.slice(0, 5).join(', '))}</span>`);

  const selCount = d.changed.filter((r) => s.selected.has(r.key)).length;

  return `
    <div class="card">
      <h2>비교 요약 <span class="sub">Excel 기준 · 값이 다른 항목만 갱신</span></h2>
      <div class="stats">
        ${stat('Excel 행', u.excelRows)}
        ${stat('List 항목', u.listItems)}
        ${stat('매칭됨', u.matched)}
        ${stat('변경 행', u.changedRows, 'hi')}
        ${stat('변경 셀', u.changedCells, 'hi')}
        ${stat('동일', u.unchanged)}
        ${stat('Excel 에만', u.excelOnly)}
        ${stat('List 에만', u.listOnly)}
      </div>
      ${s.lastRead ? renderReadLine(s.lastRead) : ''}
      ${warn.length ? alert('warn', warn.join('<br />')) : ''}
      ${u.excelOnly || u.listOnly ? alert('info', `이번 PoC 범위는 <b>변경 갱신만</b> 입니다. Excel 에만 있는 ${u.excelOnly}행(신규 후보), List 에만 있는 ${u.listOnly}건(삭제 후보)은 <b>건드리지 않습니다.</b>`) : ''}
    </div>

    <div class="card">
      <h2>쓰기 옵션 <span class="sub">$batch 묶음 + 배치 동시 전송</span></h2>
      <div class="row">
        <label class="field"><span>배치 크기</span>
          <input type="number" min="1" max="20" value="${s.writeOpts.batchSize}" data-act="batch-size" style="width:70px" />
          <span class="muted">(Graph 상한 20)</span></label>
        <label class="field"><span>동시 배치 수</span>
          <input type="number" min="1" max="16" value="${s.writeOpts.concurrency}" data-act="concurrency" style="width:70px" /></label>
        <label class="field"><input type="checkbox" data-act="dry-run" ${s.writeOpts.dryRun ? 'checked' : ''} />
          <span>드라이런 (실제로 쓰지 않음)</span></label>
      </div>
      <div class="row mt">
        <button class="btn sm" data-act="sel-all">전체 선택</button>
        <button class="btn sm" data-act="sel-none">전체 해제</button>
        <span class="muted" data-sel-count>${selCount} / ${d.changed.length} 선택됨</span>
        <div style="flex:1"></div>
        ${s.applying ? `<progress value="${s.progress.done}" max="${s.progress.total}"></progress><span class="muted">${s.progress.done}/${s.progress.total} 배치</span>` : ''}
        <button class="btn primary" data-act="apply" ${s.applying || !selCount ? 'disabled' : ''}>
          ${s.writeOpts.dryRun ? '드라이런 실행' : `선택 ${selCount}건 SharePoint 에 적용`}
        </button>
      </div>
      ${s.lastApply ? renderApplyResult(s.lastApply) : ''}
    </div>

    ${renderCompareCard(s)}`;
}

/** 이번 비교에 실제로 걸린 읽기 시간 (부가 측정치) */
function renderReadLine(r) {
  const both = r.both
    ? ` · <b>Excel 방식 비교</b> — Workbook API ${ms(r.both.wb.ms)} vs 다운로드+SheetJS ${ms(r.both.dl.ms)}
        <span class="badge ok">${r.both.wb.ms <= r.both.dl.ms ? 'Workbook API' : '다운로드+SheetJS'} 우세</span>`
    : '';
  return `<div class="row mt muted" style="font-size:13px">
    읽기 소요 — List ${ms(r.list.ms)} <span class="mono">(${r.list.count}건 · ${esc(r.list.mode)} · ${r.list.pages}페이지)</span>
    · Excel ${esc(r.excel.method)} ${ms(r.excel.ms)} <span class="mono">(${r.excel.rows.length}행)</span>${both}
  </div>`;
}

function renderApplyResult(a) {
  return `
    <div class="alert ${a.failCount ? 'warn' : 'info'} mt">
      <b>${a.dryRun ? '드라이런' : '적용'} 완료</b> — ${a.okCount}건 성공${a.failCount ? `, <b>${a.failCount}건 실패</b>` : ''}
      · ${ms(a.wallMs)} · <b>${n2(a.itemsPerSec)} 건/초</b> · 건당 ${n2(a.msPerItem)} ms
      · 배치 ${a.batchCount}개(크기 ${a.batchSize}, 동시 ${a.concurrency})${a.throttled ? ` · <b>429 스로틀 ${a.throttled}회</b>` : ''}
      ${a.failCount ? `<pre>${esc(a.results.filter((r) => !r.ok).slice(0, 5).map((r) => `${r.key}: HTTP ${r.status} ${r.error}`).join('\n'))}</pre>` : ''}
    </div>`;
}

/**
 * 좌우 분할 비교 그리드.
 * 왼쪽 고정열 = 키(Commission no.), 가운데 = 전(SharePoint 현재값), 오른쪽 = 후(Excel 적용값).
 * 세 영역의 세로 스크롤과 두 패널의 가로 스크롤은 JS 로 동기화된다 (main.js wireCompareSync).
 */
function renderCompareCard(s) {
  const d = s.diff;
  if (!d.changed.length)
    return `<div class="card"><h2>변경 상세</h2><div class="empty">차이가 없습니다. 두 소스가 동일합니다.</div></div>`;

  const shown = d.changed;
  const cols = d.changed[0].cells.map((c) => ({ field: c.field, label: c.label, excelHeader: c.excelHeader }));
  const tableW = cols.length * COL_W;

  const headHtml = (title, cls) => `
    <thead>
      <tr><th class="grp ${cls}" colspan="${cols.length}">${title}</th></tr>
      <tr>${cols.map((c) => `<th title="${esc(c.excelHeader)}">${esc(c.label)}</th>`).join('')}</tr>
    </thead>`;

  const bodyHtml = (side) =>
    shown
      .map((r) => {
        const byField = new Map(r.cells.map((c) => [c.field, c]));
        return `<tr>${cols
          .map((col) => {
            const c = byField.get(col.field);
            const text = side === 'before' ? c?.beforeText : c?.afterText;
            const cls = !c?.changed ? '' : side === 'before' ? 'chg-src' : 'chg-new';
            return `<td class="${cls}" title="${esc(text)}">${esc(text)}</td>`;
          })
          .join('')}</tr>`;
      })
      .join('');

  const keyRows = shown
    .map(
      (r) => `
      <tr>
        <td class="chk"><input type="checkbox" data-act="sel" data-key="${esc(r.key)}" ${s.selected.has(r.key) ? 'checked' : ''} /></td>
        <td class="mono" title="${esc(r.key)}">${esc(r.key)}</td>
        <td class="num"><span class="badge warn">${r.diffCells.length}</span></td>
      </tr>`
    )
    .join('');

  return `
    <div class="card">
      <h2>변경 상세 <span class="sub">왼쪽 = 전(SharePoint 현재값) · 오른쪽 = 후(Excel 적용값) · 빨간색이 변경 예정</span></h2>
      <div class="cmp" id="cmp-grid" style="--col-w:${COL_W}px">
        <div class="cmp-key">
          <table class="grid">
            <colgroup><col style="width:34px" /><col /><col style="width:44px" /></colgroup>
            <thead>
              <tr><th class="grp" colspan="3">&nbsp;</th></tr>
              <tr><th class="chk"></th><th>${esc(CONFIG.keyColumn)}</th><th class="num" title="변경 필드 수">Δ</th></tr>
            </thead>
            <tbody>${keyRows}</tbody>
          </table>
        </div>
        <div class="cmp-pane" data-sync>
          <table class="grid" style="width:${tableW}px">
            <colgroup>${cols.map(() => `<col style="width:${COL_W}px" />`).join('')}</colgroup>
            ${headHtml('전 — SharePoint 현재값', 'before')}
            <tbody>${bodyHtml('before')}</tbody>
          </table>
        </div>
        <div class="cmp-pane" data-sync>
          <table class="grid" style="width:${tableW}px">
            <colgroup>${cols.map(() => `<col style="width:${COL_W}px" />`).join('')}</colgroup>
            ${headHtml('후 — Excel 적용값', 'after')}
            <tbody>${bodyHtml('after')}</tbody>
          </table>
        </div>
      </div>
      <div class="row mt"><span class="muted">변경 행 ${shown.length}건 전체 표시</span></div>
    </div>`;
}

// ------------------------------------------------------------------ ③ 변경 목록

export function renderChanges(s) {
  if (!s.changeLog.length)
    return `<div class="card"><div class="empty">아직 적용 이력이 없습니다. <b>② 비교</b> 에서 변경을 적용하면 여기에 쌓입니다.</div></div>`;

  const entries = s.changeLog
    .map((e, i) => {
      const rows = e.items
        .map(
          (it) => `
        <tr>
          <td class="mono">${esc(it.key)}</td>
          <td class="mono muted">#${esc(it.itemId)}</td>
          <td>${it.cells.map((c) => `<span class="badge">${esc(c.label)}</span>`).join(' ')}</td>
          <td class="before mono" style="font-size:12px">${esc(it.cells.map((c) => c.before).join(' | '))}</td>
          <td class="after mono" style="font-size:12px">${esc(it.cells.map((c) => c.after).join(' | '))}</td>
          <td>${it.ok ? '<span class="badge ok">성공</span>' : `<span class="badge err">HTTP ${esc(it.status)}</span>`}</td>
        </tr>`
        )
        .join('');
      return `
      <details class="diff" ${i === 0 ? 'open' : ''}>
        <summary>
          <span class="key">${esc(new Date(e.at).toLocaleString('ko-KR'))}</span>
          ${e.dryRun ? '<span class="badge warn">드라이런</span>' : '<span class="badge ok">적용</span>'}
          <span class="badge">${e.okCount}건 성공</span>
          ${e.failCount ? `<span class="badge err">${e.failCount}건 실패</span>` : ''}
          <span class="meta">${esc(e.user)} · ${ms(e.wallMs)} · ${n2(e.itemsPerSec)} 건/초</span>
        </summary>
        <div class="diff-body scroll">
          <table>
            <thead><tr><th>키</th><th>Item</th><th>변경 필드</th><th class="col-before">전</th><th class="col-after">후</th><th>결과</th></tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </details>`;
    })
    .join('');

  return `
    <div class="card">
      <h2>적용 이력 <span class="sub">브라우저 로컬에 최근 50회 보관</span></h2>
      <div class="row">
        <button class="btn sm" data-act="export-csv">CSV 내보내기</button>
        <button class="btn sm danger" data-act="clear-log">이력 삭제</button>
      </div>
    </div>
    <div class="card">${entries}</div>`;
}

// ------------------------------------------------------------------ ④ 적용 결과 (변경된 리스트 그리드)

export function renderResult(s) {
  const p = s.postApply;
  if (!p)
    return `<div class="card"><div class="empty">
      아직 적용 결과가 없습니다.<br />
      <span class="muted">② 비교 에서 <b>드라이런을 해제하고</b> 적용하면, SharePoint 를 다시 읽어 실제 반영된 값을 여기에 그리드로 보여줍니다.</span>
    </div></div>`;

  const scope = s.resultScope;
  const shown = scope === 'changed' ? p.rows.filter((r) => r.changed) : p.rows;

  const body = shown
    .map(
      (r) => `
      <tr class="${r.changed ? 'is-changed' : ''}">
        <td class="stick mono" title="${esc(r.key)}">${r.changed ? '<span class="dot"></span>' : ''}${esc(r.key)}</td>
        ${p.cols
          .map((c) => {
            const v = r.values[c.field];
            const mark = r.marks[c.field]; // 'ok' | 'ng' | undefined
            const title = mark === 'ng' ? `기대값: ${r.intended[c.field]}` : v;
            return `<td class="${mark ? 'v-' + mark : ''}" title="${esc(title)}">${esc(v)}</td>`;
          })
          .join('')}
      </tr>`
    )
    .join('');

  return `
    <div class="card">
      <h2>적용 결과 <span class="sub">SharePoint 를 다시 읽은 실제 값 · ${esc(p.at)}</span></h2>
      <div class="stats">
        ${stat('적용 항목', p.appliedCount, 'hi')}
        ${stat('반영 확인', p.okCells, 'good')}
        ${stat('불일치', p.ngCells, p.ngCells ? 'bad' : '')}
        ${stat('재조회 소요', ms(p.readMs))}
        ${stat('전체 리스트', p.rows.length)}
      </div>
      ${
        p.ngCells
          ? alert('warn', `<b>${p.ngCells}개 셀이 기대값과 다릅니다.</b> 열 타입 변환(선택 열의 허용값, 날짜 형식 등)을 확인하세요. 셀에 마우스를 올리면 기대값이 표시됩니다.`)
          : alert('info', '적용한 모든 셀이 SharePoint 에서 기대값 그대로 확인되었습니다.')
      }
      <div class="row mt">
        <label class="field"><span>표시 범위</span>
          <select data-act="result-scope">
            <option value="changed" ${scope === 'changed' ? 'selected' : ''}>변경된 항목만 (${p.rows.filter((r) => r.changed).length})</option>
            <option value="all" ${scope === 'all' ? 'selected' : ''}>전체 리스트 (${p.rows.length})</option>
          </select>
        </label>
        <span class="muted">초록 = 반영 확인 · 빨강 = 기대값 불일치</span>
        <div style="flex:1"></div>
        <button class="btn" data-act="refresh-result" ${s.busy ? 'disabled' : ''}>다시 읽기</button>
      </div>
    </div>

    <div class="card">
      <div class="scroll result-scroll">
        <table class="grid result" style="width:${(p.cols.length + 1) * COL_W}px">
          <colgroup><col style="width:${COL_W}px" />${p.cols.map(() => `<col style="width:${COL_W}px" />`).join('')}</colgroup>
          <thead><tr>
            <th class="stick">${esc(CONFIG.keyColumn)}</th>
            ${p.cols.map((c) => `<th>${esc(c.label)}</th>`).join('')}
          </tr></thead>
          <tbody>${body || `<tr><td colspan="${p.cols.length + 1}" class="empty">표시할 행이 없습니다.</td></tr>`}</tbody>
        </table>
      </div>
      <div class="row mt"><span class="muted">${shown.length}건 전체 표시</span></div>
    </div>`;
}

// ------------------------------------------------------------------ ⑤ 벤치마크

export function renderBench(s) {
  const r = s.readBench;
  const w = s.lastApply;

  const readCard = !r
    ? `<div class="card"><h2>읽기 성능</h2>
        <div class="row"><button class="btn primary" data-act="run-read-bench" ${s.busy || !s.connected ? 'disabled' : ''}>읽기 벤치마크 실행</button>
        <span class="muted">리스트 2가지 · Excel 2가지 방식을 각각 측정합니다.</span></div>
        ${!s.connected ? '<div class="muted mt">먼저 ① 에서 연결하세요.</div>' : ''}</div>`
    : `<div class="card">
        <h2>읽기 성능 <span class="sub">${esc(r.at)}</span></h2>
        <div class="scroll-x"><table>
          <thead><tr><th>소스</th><th>방식</th><th class="num">소요</th><th class="num">건수/행</th><th class="num">건당</th><th class="num">전송량</th><th>비고</th></tr></thead>
          <tbody>
            <tr><td>SharePoint List</td><td>$expand=fields (전체 필드)</td><td class="num">${ms(r.listFull.ms)}</td><td class="num">${r.listFull.count}</td>
                <td class="num">${n2(r.listFull.ms / Math.max(1, r.listFull.count))} ms</td><td class="num">—</td><td class="muted">${r.listFull.pages} 페이지</td></tr>
            <tr><td>SharePoint List</td><td>$select 최적화</td><td class="num">${ms(r.listSelect.ms)}</td><td class="num">${r.listSelect.count}</td>
                <td class="num">${n2(r.listSelect.ms / Math.max(1, r.listSelect.count))} ms</td><td class="num">—</td>
                <td class="${r.listSelect.ms < r.listFull.ms ? 'mono' : 'muted'}">${r.listFull.ms ? ((1 - r.listSelect.ms / r.listFull.ms) * 100).toFixed(0) + '% 단축' : ''}</td></tr>
            <tr><td>Excel</td><td>Workbook API (usedRange, valuesOnly)</td><td class="num">${ms(r.wb.ms)}</td><td class="num">${r.wb.rows}</td>
                <td class="num">${n2(r.wb.ms / Math.max(1, r.wb.rows))} ms</td><td class="num">—</td><td class="muted">세션+시트+범위 3콜</td></tr>
            <tr><td>Excel</td><td>다운로드 + SheetJS</td><td class="num">${ms(r.dl.ms)}</td><td class="num">${r.dl.rows}</td>
                <td class="num">${n2(r.dl.ms / Math.max(1, r.dl.rows))} ms</td><td class="num">${kb(r.dl.bytes)}</td>
                <td class="muted">다운로드 ${ms(r.dl.downloadMs)} + 파싱 ${ms(r.dl.parseMs)}</td></tr>
          </tbody>
        </table></div>
        <div class="row mt"><button class="btn" data-act="run-read-bench" ${s.busy ? 'disabled' : ''}>다시 측정</button>
        <span class="muted">Excel 읽기 승자: <b>${r.wb.ms <= r.dl.ms ? 'Workbook API' : '다운로드 + SheetJS'}</b> (${ms(Math.abs(r.wb.ms - r.dl.ms))} 차이)</span></div>
      </div>`;

  const writeCard = !w
    ? `<div class="card"><h2>쓰기 성능</h2><div class="empty">② 비교 에서 변경을 적용하면 결과가 여기에 표시됩니다.</div></div>`
    : `<div class="card">
        <h2>쓰기 성능 <span class="sub">${w.dryRun ? '드라이런 (네트워크 미포함)' : 'Graph $batch PATCH'}</span></h2>
        <div class="stats">
          ${stat('처리 건수', w.itemCount, 'hi')}
          ${stat('변경 셀', w.cellCount)}
          ${stat('총 소요', ms(w.wallMs), 'hi')}
          ${stat('처리량', n2(w.itemsPerSec) + ' <span style="font-size:12px">건/초</span>', 'good')}
          ${stat('건당', n2(w.msPerItem) + ' <span style="font-size:12px">ms</span>')}
          ${stat('실패', w.failCount, w.failCount ? 'bad' : '')}
          ${stat('429 스로틀', w.throttled, w.throttled ? 'bad' : '')}
        </div>
        <div class="row mt muted">
          배치 ${w.batchCount}개 · 크기 ${w.batchSize} · 동시 ${w.concurrency}
          · 배치 지연 평균 ${ms(w.batchStats.avg)} / p50 ${ms(w.batchStats.p50)} / p95 ${ms(w.batchStats.p95)} / 최대 ${ms(w.batchStats.max)}
        </div>
        <div class="scroll mt" style="max-height:260px"><table>
          <thead><tr><th>#</th><th class="num">건수</th><th class="num">소요</th><th class="num">건당</th></tr></thead>
          <tbody>${w.batchTimings
            .slice()
            .sort((a, b) => a.index - b.index)
            .map((b) => `<tr><td>${b.index}</td><td class="num">${b.size}</td><td class="num">${ms(b.ms)}</td><td class="num">${n2(b.ms / b.size)} ms</td></tr>`)
            .join('')}</tbody>
        </table></div>
      </div>`;

  const calls = s.telemetry.slice(-200).reverse();
  const callCard = `
    <div class="card">
      <h2>원시 호출 로그 <span class="sub">최근 200건 · 클라이언트 측 왕복 시간</span></h2>
      <div class="row"><button class="btn sm" data-act="clear-telemetry">초기화</button></div>
      <div class="scroll mt"><table>
        <thead><tr><th>이름</th><th>메서드</th><th class="num">HTTP</th><th class="num">소요</th><th class="num">재시도</th><th class="num">크기</th></tr></thead>
        <tbody>${
          calls
            .map(
              (c) => `<tr>
              <td class="mono" style="font-size:12px">${esc(c.name)}</td>
              <td class="muted">${esc(c.method)}</td>
              <td class="num">${c.throttled ? `<span class="badge err">${c.status}</span>` : c.status}</td>
              <td class="num">${ms(c.ms)}</td>
              <td class="num">${c.retries || ''}</td>
              <td class="num">${c.bytes ? kb(c.bytes) : ''}</td>
            </tr>`
            )
            .join('') || `<tr><td colspan="6" class="empty">호출 없음</td></tr>`
        }</tbody>
      </table></div>
    </div>`;

  return readCard + writeCard + callCard;
}
