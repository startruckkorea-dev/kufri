import { useEffect, useMemo, useRef, useState } from 'react';
import { display, toDate } from '../../lib/diff.js';
import { BENCH_FIELD, WritePanel, isWritable, useWriteBench } from './WriteBench.jsx'; // TEMP(쓰기 속도 측정)

const ms = (v) => (v == null ? '—' : v >= 1000 ? `${(v / 1000).toFixed(2)} s` : `${Math.round(v)} ms`);
const hhmm = (d) => (d ? d.toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' }) : '');

/** 그리드 셀 문자열. 비교 화면과 달리 빈 값은 빈 칸으로 둔다 */
const cellText = (v, col) => (v == null || v === '' ? '' : display(v, col));

/** 타입별 정렬 비교. 빈 값은 항상 뒤로 */
function compare(a, b, col) {
  const ea = a == null || a === '';
  const eb = b == null || b === '';
  if (ea && eb) return 0;
  if (ea) return 1;
  if (eb) return -1;
  if (col.type === 'number') return Number(a) - Number(b);
  if (col.type === 'dateTime') return (toDate(a)?.getTime() ?? 0) - (toDate(b)?.getTime() ?? 0);
  return String(a).localeCompare(String(b), 'ko');
}

/** TEMP: 셀 하나 편집 입력칸. Enter 저장, Esc·포커스 이탈 취소 */
function CellEditor({ initial, onSave, onCancel }) {
  const ref = useRef(null);
  useEffect(() => {
    ref.current?.focus();
    ref.current?.select();
  }, []);
  return (
    <input
      ref={ref}
      type="text"
      defaultValue={initial}
      onKeyDown={(e) => {
        if (e.key === 'Enter') onSave(e.currentTarget.value);
        else if (e.key === 'Escape') onCancel();
      }}
      onBlur={onCancel}
      aria-label="셀 값 입력"
    />
  );
}

/**
 * 전체 데이터 그리드 — 대상 리스트 전체를 Excel 처럼 한 표로 본다.
 * 검색·정렬은 클라이언트 메모리에서만 한다. 사용자 입력은 Graph 요청에 들어가지 않는다 (3-05).
 * TEMP: series 열만 편집 가능 (쓰기 속도 측정). 행 선택(체크박스·Shift 범위)은 일괄 입력용.
 */
export default function ListGrid({ data, busy, onRefresh, onPatched }) {
  const [query, setQuery] = useState('');
  const [sort, setSort] = useState(null); // {field, dir: 1 | -1}
  const [selected, setSelected] = useState(() => new Set()); // TEMP: 행 key
  const [editing, setEditing] = useState(null); // TEMP: 편집 중인 행 key
  const lastClick = useRef(null); // TEMP: Shift 범위 선택 기준 (표시 순서 인덱스)

  const cols = data?.columns ?? [];
  const benchCol = cols.find((c) => c.name === BENCH_FIELD) ?? null; // TEMP
  const keyCol = cols.find((c) => c.isKey) ?? null;
  const bench = useWriteBench({ column: benchCol, keyColumn: keyCol, onPatched }); // TEMP
  const editable = isWritable(benchCol); // TEMP

  const rows = useMemo(() => {
    if (!data) return [];
    let out = data.rows.map((r) => ({
      key: `${r.listId}:${r.id}`,
      id: r.id,
      listId: r.listId,
      listTitle: r.listTitle,
      fields: r.fields,
      cells: cols.map((c) => cellText(r.fields[c.name], c)),
    }));

    const q = query.trim().toLowerCase();
    if (q) {
      out = out.filter(
        (r) => r.cells.some((t) => t.toLowerCase().includes(q)) || r.listTitle.toLowerCase().includes(q)
      );
    }
    if (sort) {
      const col = cols.find((c) => c.name === sort.field);
      if (col) out = [...out].sort((x, y) => sort.dir * compare(x.fields[col.name], y.fields[col.name], col));
    }
    return out;
  }, [data, cols, query, sort]);

  const toggleSort = (field) =>
    setSort((s) => (!s || s.field !== field ? { field, dir: 1 } : s.dir === 1 ? { field, dir: -1 } : null));

  // TEMP: 행 선택. Shift 클릭이면 마지막 클릭 행과의 범위(현재 표시 순서)를 같은 상태로 맞춘다
  const toggleRow = (i, shift) => {
    setSelected((prev) => {
      const next = new Set(prev);
      const from = shift && lastClick.current != null ? Math.min(lastClick.current, i) : i;
      const to = shift && lastClick.current != null ? Math.max(lastClick.current, i) : i;
      const on = !prev.has(rows[i].key);
      for (let k = from; k <= to; k++) {
        if (on) next.add(rows[k].key);
        else next.delete(rows[k].key);
      }
      return next;
    });
    lastClick.current = i;
  };
  const allShown = rows.length > 0 && rows.every((r) => selected.has(r.key));
  const toggleAllShown = () =>
    setSelected((prev) => {
      const next = new Set(prev);
      if (allShown) rows.forEach((r) => next.delete(r.key));
      else rows.forEach((r) => next.add(r.key));
      return next;
    });
  const selectedRows = rows.filter((r) => selected.has(r.key));

  if (!data) {
    return (
      <div className="card">
        <div className="empty">
          {busy ? (
            '리스트를 읽는 중…'
          ) : (
            <>
              데이터를 아직 읽지 않았습니다.{' '}
              <button className="btn sm" onClick={onRefresh}>
                읽기
              </button>
            </>
          )}
        </div>
      </div>
    );
  }

  const perList = data.perList.map((p) => `${p.title} ${p.count}`).join(' + ');
  const colClass = (c, i) =>
    `${i === 0 ? 'stick2' : ''} ${c.type === 'number' ? 'num' : ''} ${c.extra ? 'extra' : ''}`;

  return (
    <div className="card">
      <div className="dg-toolbar">
        <input
          type="text"
          placeholder="검색 (모든 열)"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          style={{ width: 240 }}
          aria-label="검색"
        />
        <span className="muted">
          {query ? `${rows.length} / ${data.count}행` : `${data.count}행`}
          {data.perList.length > 1 ? <span className="mono"> ({perList})</span> : null} · {cols.length}열
          {cols.some((c) => c.extra) ? ` (정의 외 ${cols.filter((c) => c.extra).length}열은 맨 뒤)` : ''}
          {selected.size ? ` · 선택 ${selected.size}행` : ''}
        </span>
        <div style={{ flex: 1 }} />
        <span className="muted" title={`리졸브 ${ms(data.resolveMs)} + 항목 ${ms(data.ms)} (${data.pages}페이지)`}>
          {hhmm(data.readAt)} 읽음 · {ms(data.resolveMs + data.ms)}
          {data.resolveMs > 100 ? ` (리졸브 ${ms(data.resolveMs)} + 항목 ${ms(data.ms)})` : ''}
        </span>
        <button className="btn sm" onClick={onRefresh} disabled={busy || bench.writing}>
          새로고침
        </button>
      </div>

      {data.missing.length ? (
        <div className="alert warn">
          리스트에서 찾지 못한 열 {data.missing.length}개는 표시하지 않습니다. columns.js 의 내부명을 실제 열과
          맞추세요: <span className="mono">{data.missing.join(', ')}</span>
        </div>
      ) : null}

      {/* TEMP(쓰기 속도 측정) */}
      <WritePanel bench={bench} selectedRows={selectedRows} onClearSelection={() => setSelected(new Set())} />

      <div className="dg-wrap">
        <table className="dg">
          <thead>
            <tr>
              <th className="stick1 num">#</th>
              <th className="stickc">
                <input type="checkbox" checked={allShown} onChange={toggleAllShown} aria-label="표시 중인 행 전체 선택" />
              </th>
              {cols.map((c, i) => (
                <th
                  key={c.name}
                  className={colClass(c, i)}
                  onClick={() => toggleSort(c.name)}
                  title={`${c.name} · ${c.type}${c.isKey ? ' · 키' : ''}${c.extra ? ' · 열 정의에 없는 열 (리스트 표시명)' : ''} · 클릭하여 정렬`}
                  aria-sort={sort?.field === c.name ? (sort.dir === 1 ? 'ascending' : 'descending') : 'none'}
                >
                  {c.label || c.displayName || c.name}
                  {sort?.field === c.name ? <span className="muted">{sort.dir === 1 ? ' ▲' : ' ▼'}</span> : null}
                  {editable && c.name === BENCH_FIELD ? <span className="badge warn" style={{ marginLeft: 6 }}>편집</span> : null}
                </th>
              ))}
              <th title="행이 있는 리스트">리스트</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={cols.length + 3} className="empty">
                  {query ? '검색 결과가 없습니다.' : '리스트에 항목이 없습니다.'}
                </td>
              </tr>
            ) : null}
            {rows.map((r, i) => (
              <tr key={r.key} className={selected.has(r.key) ? 'sel' : ''}>
                <td className="stick1 num muted">{i + 1}</td>
                <td className="stickc">
                  <input
                    type="checkbox"
                    checked={selected.has(r.key)}
                    onClick={(e) => toggleRow(i, e.shiftKey)}
                    onChange={() => {}}
                    aria-label={`${r.cells[0]} 선택`}
                  />
                </td>
                {r.cells.map((t, ci) => {
                  const c = cols[ci];
                  const isBench = editable && c.name === BENCH_FIELD; // TEMP
                  if (isBench && editing === r.key) {
                    return (
                      <td key={c.name} className={`${colClass(c, ci)} editing`}>
                        <CellEditor
                          initial={t}
                          onCancel={() => setEditing(null)}
                          onSave={async (v) => {
                            setEditing(null);
                            await bench.write([r], v);
                          }}
                        />
                      </td>
                    );
                  }
                  return (
                    <td
                      key={c.name}
                      className={`${colClass(c, ci)} ${isBench ? 'editable' : ''}`}
                      title={isBench ? '클릭하여 입력' : t}
                      onClick={isBench && !bench.writing ? () => setEditing(r.key) : undefined}
                    >
                      {t}
                    </td>
                  );
                })}
                <td className="muted">{r.listTitle}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
