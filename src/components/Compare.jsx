import { useEffect, useRef } from 'react';
import { CONFIG } from '../lib/config.js';

const COL_W = 160;
const ms = (v) => (v == null ? '—' : v >= 1000 ? `${(v / 1000).toFixed(2)} s` : `${Math.round(v)} ms`);

/**
 * 좌우 분할 비교 그리드.
 * 왼쪽 고정열 = Commission no., 가운데 = 전(SharePoint 현재값), 오른쪽 = 후(Excel 적용값).
 *
 * 세로 스크롤은 세 영역이, 가로 스크롤은 두 패널이 함께 움직인다.
 * PoC 는 매 렌더마다 리스너를 다시 붙이기만 했지만 여기서는 useEffect 정리 함수로 해제한다.
 */
function SplitGrid({ cols, rows, leftTitle, rightTitle }) {
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
            <col />
            <col style={{ width: 52 }} />
          </colgroup>
          <thead>
            <tr>
              <th className="grp" colSpan={2}>
                &nbsp;
              </th>
            </tr>
            <tr>
              <th>{CONFIG.keyColumn}</th>
              <th className="num" title="바뀌는 필드 수">
                Δ
              </th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.key}>
                <td className="mono" title={r.key}>
                  {r.key}
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

export default function Compare({ diff, lastRead }) {
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

  const cols =
    diff.changed[0]?.cells.map((c) => ({ field: c.field, label: c.label, title: c.excelHeader })) ?? [];
  const rows = diff.changed.map((r) => ({
    key: r.key,
    changedCount: r.diffCells.length,
    cells: Object.fromEntries(
      r.cells.map((c) => [c.field, { left: c.beforeText, right: c.afterText, changed: c.changed }])
    ),
  }));

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
              ({lastRead.list.count}건 · {lastRead.list.mode} · {lastRead.list.pages}페이지)
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
                List 에 중복 키 {w.dupListKeys.length}건 (첫 항목만 대상):{' '}
                <span className="mono">{w.dupListKeys.slice(0, 5).join(', ')}</span>
              </div>
            ) : null}
          </div>
        ) : null}
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
