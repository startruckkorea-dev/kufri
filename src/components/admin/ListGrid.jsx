import { useMemo, useState } from 'react';
import { display, toDate } from '../../lib/diff.js';

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

/**
 * 전체 데이터 그리드 — 대상 리스트 전체를 Excel 처럼 한 표로 본다 (읽기 전용).
 * 검색·정렬은 클라이언트 메모리에서만 한다. 사용자 입력은 Graph 요청에 들어가지 않는다 (3-05).
 */
export default function ListGrid({ data, busy, onRefresh }) {
  const [query, setQuery] = useState('');
  const [sort, setSort] = useState(null); // {field, dir: 1 | -1}

  const rows = useMemo(() => {
    if (!data) return [];
    const cols = data.columns;
    let out = data.rows.map((r) => ({
      key: `${r.listId}:${r.id}`,
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
  }, [data, query, sort]);

  const toggleSort = (field) =>
    setSort((s) => (!s || s.field !== field ? { field, dir: 1 } : s.dir === 1 ? { field, dir: -1 } : null));

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

  const cols = data.columns;
  const perList = data.perList.map((p) => `${p.title} ${p.count}`).join(' + ');

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
        </span>
        <div style={{ flex: 1 }} />
        <span className="muted">
          {hhmm(data.readAt)} 읽음 · {ms(data.ms)}
        </span>
        <button className="btn sm" onClick={onRefresh} disabled={busy}>
          새로고침
        </button>
      </div>

      {data.missing.length ? (
        <div className="alert warn">
          리스트에서 찾지 못한 열 {data.missing.length}개는 표시하지 않습니다. columns.js 의 내부명을 실제 열과
          맞추세요: <span className="mono">{data.missing.join(', ')}</span>
        </div>
      ) : null}

      <div className="dg-wrap">
        <table className="dg">
          <thead>
            <tr>
              <th className="stick1 num">#</th>
              {cols.map((c, i) => (
                <th
                  key={c.name}
                  className={`${i === 0 ? 'stick2' : ''} ${c.type === 'number' ? 'num' : ''} ${c.extra ? 'extra' : ''}`}
                  onClick={() => toggleSort(c.name)}
                  title={`${c.name} · ${c.type}${c.isKey ? ' · 키' : ''}${c.extra ? ' · 열 정의에 없는 열 (리스트 표시명)' : ''} · 클릭하여 정렬`}
                  aria-sort={sort?.field === c.name ? (sort.dir === 1 ? 'ascending' : 'descending') : 'none'}
                >
                  {c.label || c.displayName || c.name}
                  {sort?.field === c.name ? <span className="muted">{sort.dir === 1 ? ' ▲' : ' ▼'}</span> : null}
                </th>
              ))}
              <th title="행이 있는 리스트">리스트</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={cols.length + 2} className="empty">
                  {query ? '검색 결과가 없습니다.' : '리스트에 항목이 없습니다.'}
                </td>
              </tr>
            ) : null}
            {rows.map((r, i) => (
              <tr key={r.key}>
                <td className="stick1 num muted">{i + 1}</td>
                {r.cells.map((t, ci) => (
                  <td key={cols[ci].name} className={`${ci === 0 ? 'stick2' : ''} ${cols[ci].type === 'number' ? 'num' : ''}`} title={t}>
                    {t}
                  </td>
                ))}
                <td className="muted">{r.listTitle}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
