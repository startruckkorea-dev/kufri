import { useEffect, useMemo, useRef, useState } from 'react';
import { CONFIG } from '../lib/config.js';
import { applyChanges } from '../lib/sync.js';
import { ctx } from '../lib/sources.js';

const COL_W = 160;
const ms = (v) => (v == null ? '—' : v >= 1000 ? `${(v / 1000).toFixed(2)} s` : `${Math.round(v)} ms`);

/**
 * 좌우 분할 비교 그리드.
 * 왼쪽 고정열 = Commission no., 가운데 = 전(SharePoint 현재값), 오른쪽 = 후(Excel 적용값).
 *
 * 세로 스크롤은 세 영역이, 가로 스크롤은 두 패널이 함께 움직인다.
 * PoC 는 매 렌더마다 리스너를 다시 붙이기만 했지만 여기서는 useEffect 정리 함수로 해제한다.
 */
function SplitGrid({ cols, rows, selected, onToggle, leftTitle, rightTitle }) {
  const keyRef = useRef(null);
  const leftRef = useRef(null);
  const rightRef = useRef(null);

  useEffect(() => {
    const panes = [leftRef.current, rightRef.current].filter(Boolean);
    const all = [keyRef.current, ...panes].filter(Boolean);
    if (all.length < 2) return undefined;

    let lock = false;
    const bound = all.map((el) => {
      const onScroll = () => {
        if (lock) return;
        lock = true;
        for (const other of all) if (other !== el) other.scrollTop = el.scrollTop;
        if (panes.includes(el)) for (const p of panes) if (p !== el) p.scrollLeft = el.scrollLeft;
        requestAnimationFrame(() => {
          lock = false;
        });
      };
      el.addEventListener('scroll', onScroll);
      return [el, onScroll];
    });

    return () => bound.forEach(([el, fn]) => el.removeEventListener('scroll', fn));
  }, [rows.length, cols.length]);

  const tableW = cols.length * COL_W;
  const colGroup = (
    <colgroup>
      {cols.map((c) => (
        <col key={c.field} style={{ width: COL_W }} />
      ))}
    </colgroup>
  );

  const head = (title, cls) => (
    <thead>
      <tr>
        <th className={`grp ${cls}`} colSpan={cols.length}>
          {title}
        </th>
      </tr>
      <tr>
        {cols.map((c) => (
          <th key={c.field} title={c.title || c.label}>
            {c.label}
          </th>
        ))}
      </tr>
    </thead>
  );

  const body = (side) => (
    <tbody>
      {rows.map((r) => (
        <tr key={r.key}>
          {cols.map((col) => {
            const cell = r.cells[col.field];
            const text = cell ? (side === 'left' ? cell.left : cell.right) : '';
            const cls = !cell?.changed ? '' : side === 'left' ? 'chg-src' : 'chg-new';
            return (
              <td key={col.field} className={cls} title={text}>
                {text}
              </td>
            );
          })}
        </tr>
      ))}
    </tbody>
  );

  return (
    <div className="cmp">
      <div className="cmp-key" ref={keyRef}>
        <table className="grid">
          <colgroup>
            <col style={{ width: 34 }} />
            <col />
            <col style={{ width: 110 }} />
            <col style={{ width: 52 }} />
          </colgroup>
          <thead>
            <tr>
              <th className="grp" colSpan={4}>
                &nbsp;
              </th>
            </tr>
            <tr>
              <th className="chk" />
              <th title={`Excel 헤더: ${CONFIG.mapping.key.excel}`}>{CONFIG.mapping.key.list}</th>
              <th title="행이 원래 있던 리스트. 적용도 이 리스트로 갑니다">리스트</th>
              <th className="num" title="바뀌는 필드 수">
                Δ
              </th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.key}>
                <td className="chk">
                  <input
                    type="checkbox"
                    checked={selected.has(r.key)}
                    onChange={() => onToggle(r.key)}
                    aria-label={`${r.key} 적용 대상 선택`}
                  />
                </td>
                <td className="mono" title={r.key}>
                  {r.key}
                </td>
                <td className="muted" title={r.listTitle}>
                  {r.listTitle}
                </td>
                <td className="num">
                  <span className="badge warn">{r.changedCount}</span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="cmp-pane" ref={leftRef}>
        <table className="grid" style={{ width: tableW }}>
          {colGroup}
          {head(leftTitle, 'before')}
          {body('left')}
        </table>
      </div>

      <div className="cmp-pane" ref={rightRef}>
        <table className="grid" style={{ width: tableW }}>
          {colGroup}
          {head(rightTitle, 'after')}
          {body('right')}
        </table>
      </div>
    </div>
  );
}

const Stat = ({ label, value, cls = '' }) => (
  <div className={`stat ${cls}`}>
    <div className="label">{label}</div>
    <div className="value">{value}</div>
  </div>
);

function ApplyResult({ result }) {
  if (!result) return null;
  const failed = result.results.filter((r) => !r.ok);
  return (
    <div className={`alert ${result.failCount ? 'warn' : 'info'} mt`}>
      <b>{result.dryRun ? '드라이런' : '적용'} 완료</b> — {result.okCount}건 성공
      {result.failCount > 0 ? <b>, {result.failCount}건 실패</b> : null} · {ms(result.wallMs)}
      {result.throttled > 0 ? <b> · 429 스로틀 {result.throttled}회</b> : null}
      {result.dryRun ? null : (
        <div className="muted" style={{ marginTop: 4 }}>
          반영 여부를 확인하려고 비교를 다시 실행했습니다. 아직 남아 있는 행이 있다면 값이 그대로 저장되지 않은
          것입니다 (열 타입·선택 열 허용값 확인).
        </div>
      )}
      {failed.length > 0 ? (
        <pre>{failed.slice(0, 5).map((r) => `${r.key}: HTTP ${r.status} ${r.error ?? ''}`).join('\n')}</pre>
      ) : null}
    </div>
  );
}

export default function Compare({ diff, lastRead, onApplied }) {
  const [selected, setSelected] = useState(() => new Set());
  const [writeOpts, setWriteOpts] = useState({
    batchSize: CONFIG.defaults.batchSize,
    concurrency: CONFIG.defaults.concurrency,
    dryRun: true,
  });
  const [applying, setApplying] = useState(false);
  const [progress, setProgress] = useState({ done: 0, total: 0 });
  const [result, setResult] = useState(null);
  const [writeError, setWriteError] = useState(null);

  // 새 비교 결과가 오면 전체 선택으로 초기화한다
  const changedKeys = useMemo(() => (diff ? diff.changed.map((r) => r.key) : []), [diff]);
  useEffect(() => {
    setSelected(new Set(changedKeys));
    setResult(null);
    setWriteError(null);
  }, [changedKeys]);

  if (!diff) {
    return (
      <div className="card">
        <div className="empty">
          먼저 <b>① 연결 · 매핑</b> 에서 비교를 실행하세요.
        </div>
      </div>
    );
  }

  const u = diff.summary;
  const w = diff.warnings;

  const cols = diff.changed[0]?.cells.map((c) => ({ field: c.field, label: c.label, title: c.excelHeader })) ?? [];
  const rows = diff.changed.map((r) => ({
    key: r.key,
    listTitle: r.listTitle,
    changedCount: r.diffCells.length,
    cells: Object.fromEntries(
      r.cells.map((c) => [c.field, { left: c.beforeText, right: c.afterText, changed: c.changed }])
    ),
  }));

  const targets = diff.changed.filter((r) => selected.has(r.key));
  const cellCount = targets.reduce((a, r) => a + r.diffCells.length, 0);

  const toggle = (key) =>
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });

  async function handleApply() {
    if (targets.length === 0) return;
    if (
      !writeOpts.dryRun &&
      !window.confirm(
        `SharePoint 리스트 '${ctx.listTitle}' 의 ${targets.length}개 항목 ` +
          `${cellCount}개 필드를 실제로 갱신합니다.\n\n` +
          `변경된 필드만 씁니다. 각 항목에는 버전이 하나씩 쌓입니다.\n\n진행할까요?`
      )
    )
      return;

    setApplying(true);
    setWriteError(null);
    setResult(null);
    setProgress({ done: 0, total: Math.ceil(targets.length / writeOpts.batchSize) });

    try {
      const res = await applyChanges(targets, {
        ...writeOpts,
        onProgress: (done, total) => setProgress({ done, total }),
      });
      // 실제로 썼다면 비교를 다시 돌려 반영 여부를 검증한다.
      // 값이 제대로 저장됐다면 해당 행은 더 이상 차이가 없어 목록에서 사라진다.
      // 재비교가 selected/result 를 초기화하므로 결과 배너는 그 뒤에 세팅한다.
      if (!res.dryRun && res.okCount > 0 && onApplied) await onApplied();
      setResult(res);
    } catch (e) {
      console.error(e);
      setWriteError(e);
    } finally {
      setApplying(false);
    }
  }

  return (
    <>
      <div className="card">
        <h2>
          비교 요약 <span className="sub">Excel 기준 · 값이 다른 항목만</span>
        </h2>
        <div className="stats">
          <Stat label="Excel 행" value={u.excelRows} />
          <Stat label="List 항목" value={u.listItems} />
          <Stat label="매칭됨" value={u.matched} />
          <Stat label="변경 행" value={u.changedRows} cls="hi" />
          <Stat label="변경 셀" value={u.changedCells} cls="hi" />
          <Stat label="동일" value={u.unchanged} />
          <Stat label="Excel 에만" value={u.excelOnly} />
          <Stat label="List 에만" value={u.listOnly} />
        </div>

        {lastRead ? (
          <div className="row mt muted" style={{ fontSize: 13 }}>
            읽기 소요 — List {ms(lastRead.list.ms)}{' '}
            <span className="mono">
              ({lastRead.list.count}건
              {lastRead.list.perList?.length > 1
                ? ' = ' + lastRead.list.perList.map((p) => `${p.title} ${p.count}`).join(' + ')
                : ''}{' '}
              · {lastRead.list.mode} · {lastRead.list.pages}페이지)
            </span>{' '}
            · Excel {lastRead.excel.method} {ms(lastRead.excel.ms)}{' '}
            <span className="mono">({lastRead.excel.rows.length}행)</span>
          </div>
        ) : null}

        {w.dupExcelKeys.length > 0 || w.dupListKeys.length > 0 ? (
          <div className="alert warn">
            {w.dupExcelKeys.length > 0 ? (
              <div>
                Excel 에 중복 키 {w.dupExcelKeys.length}건 (첫 행만 사용):{' '}
                <span className="mono">{w.dupExcelKeys.slice(0, 5).join(', ')}</span>
              </div>
            ) : null}
            {w.dupListKeys.length > 0 ? (
              <div>
                List 에 중복 키 {w.dupListKeys.length}건 (리스트 간 중복 포함 · 첫 항목만 대상):{' '}
                <span className="mono">{w.dupListKeys.slice(0, 5).join(', ')}</span>
              </div>
            ) : null}
          </div>
        ) : null}

        {u.excelOnly > 0 || u.listOnly > 0 ? (
          <div className="alert info">
            이 도구는 <b>값이 다른 항목만 갱신</b>합니다. Excel 에만 있는 {u.excelOnly}행(신규 후보)과 List 에만
            있는 {u.listOnly}건(삭제 후보)은 <b>건드리지 않습니다.</b>
          </div>
        ) : null}
      </div>

      <div className="card">
        <h2>
          적용 <span className="sub">변경된 필드만 SharePoint 에 씁니다</span>
        </h2>

        <div className="row">
          <label className="field">
            <span>배치 크기</span>
            <input
              type="number"
              min={1}
              max={20}
              value={writeOpts.batchSize}
              disabled={applying}
              onChange={(e) =>
                setWriteOpts((o) => ({ ...o, batchSize: Math.min(20, Math.max(1, Number(e.target.value) || 20)) }))
              }
              style={{ width: 70 }}
            />
            <span className="muted">(Graph 상한 20)</span>
          </label>
          <label className="field">
            <span>동시 배치 수</span>
            <input
              type="number"
              min={1}
              max={8}
              value={writeOpts.concurrency}
              disabled={applying}
              onChange={(e) =>
                setWriteOpts((o) => ({ ...o, concurrency: Math.min(8, Math.max(1, Number(e.target.value) || 2)) }))
              }
              style={{ width: 70 }}
            />
          </label>
          <label className="field">
            <input
              type="checkbox"
              checked={writeOpts.dryRun}
              disabled={applying}
              onChange={(e) => setWriteOpts((o) => ({ ...o, dryRun: e.target.checked }))}
            />
            <span>드라이런 (실제로 쓰지 않음)</span>
          </label>
        </div>

        <div className="row">
          <span className="muted">
            동시 배치 수는 SharePoint 부하를 좌우합니다. 다른 시스템이 같은 테넌트를 쓰고 있으면 낮게 두세요.
          </span>
        </div>

        <div className="row mt">
          <button className="btn sm" disabled={applying} onClick={() => setSelected(new Set(changedKeys))}>
            전체 선택
          </button>
          <button className="btn sm" disabled={applying} onClick={() => setSelected(new Set())}>
            전체 해제
          </button>
          <span className="muted">
            {targets.length} / {diff.changed.length} 선택됨 · {cellCount}개 필드
          </span>
          <div style={{ flex: 1 }} />
          {applying ? (
            <>
              <progress value={progress.done} max={progress.total || 1} />
              <span className="muted">
                {progress.done}/{progress.total} 배치
              </span>
            </>
          ) : null}
          <button className="btn primary" onClick={handleApply} disabled={applying || targets.length === 0}>
            {writeOpts.dryRun ? '드라이런 실행' : `선택 ${targets.length}건 적용`}
          </button>
        </div>

        {writeError ? (
          <div className="alert err mt">
            <b>쓰기 실패</b> — {writeError.message || String(writeError)}
          </div>
        ) : null}
        <ApplyResult result={result} />
      </div>

      <div className="card">
        <h2>
          변경 상세{' '}
          <span className="sub">
            왼쪽 = 전(SharePoint 현재값) · 오른쪽 = 후(Excel 적용값) · 빨간색이 변경 예정
          </span>
        </h2>
        {rows.length === 0 ? (
          <div className="empty">차이가 없습니다. 두 소스가 동일합니다.</div>
        ) : (
          <>
            <SplitGrid
              cols={cols}
              rows={rows}
              selected={selected}
              onToggle={toggle}
              leftTitle="전 — SharePoint 현재값"
              rightTitle="후 — Excel 적용값"
            />
            <div className="row mt">
              <span className="muted">변경 행 {rows.length}건 전체 표시</span>
            </div>
          </>
        )}
      </div>
    </>
  );
}
