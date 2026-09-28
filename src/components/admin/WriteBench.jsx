// TEMP(쓰기 속도 측정) — series 열 키인 · 선택 행 일괄 채우기/비우기 · 측정 기록.
// 측정이 끝나면 이 파일과 ListGrid 의 "TEMP" 표시 부분을 걷어내거나, eTag 충돌 처리를 붙여 정식 편집으로 키운다.
// 쓰기 경로는 동기화와 같은 applyChanges($batch 20건 · 동시성 N · 변경 필드만)라 측정값이 실제 동기화와 같은 조건이다.
import { useCallback, useState } from 'react';
import { CONFIG } from '../../lib/config.js';
import { applyChanges } from '../../lib/sync.js';

export const BENCH_FIELD = 'series';

const ms = (v) => (v == null ? '—' : v >= 1000 ? `${(v / 1000).toFixed(2)} s` : `${Math.round(v)} ms`);
const hhmmss = (d) => d.toLocaleTimeString('ko-KR', { hour12: false });

/** 쓸 수 있는 열인지. 계산·조회·사용자 열과 읽기 전용은 제외 */
export const isWritable = (col) => !!col && !col.readOnly && !['lookup', 'person', 'calculated'].includes(col.type);

/**
 * @param column     쓰기 대상 열 정의 (그리드 columns 에서 BENCH_FIELD 로 찾은 것)
 * @param keyColumn  실패 표시용 키 열
 * @param onPatched  성공한 항목을 메모리에 반영하는 콜백 [{listId, id, fields}]
 */
export function useWriteBench({ column, keyColumn, onPatched }) {
  const [concurrency, setConcurrency] = useState(CONFIG.defaults.concurrency);
  const [writing, setWriting] = useState(false);
  const [progress, setProgress] = useState({ done: 0, total: 0 });
  const [runs, setRuns] = useState([]); // 최근 측정 기록 (최신이 앞)
  const [error, setError] = useState(null);

  /** rows: 그리드 행 [{id, listId, fields}], value: 문자열 또는 null(비우기) */
  const write = useCallback(
    async (rows, value) => {
      if (!column || !rows.length) return null;
      const v = value == null || String(value).trim() === '' ? null : String(value).trim();
      const changed = rows.map((r) => ({
        key: keyColumn ? String(r.fields[keyColumn.name] ?? r.id) : String(r.id),
        itemId: r.id,
        listId: r.listId,
        diffCells: [{ field: column.name, writeValue: v }],
      }));

      setWriting(true);
      setError(null);
      setProgress({ done: 0, total: Math.ceil(changed.length / CONFIG.defaults.batchSize) });
      try {
        const res = await applyChanges(changed, {
          batchSize: CONFIG.defaults.batchSize,
          concurrency,
          dryRun: false,
          onProgress: (done, total) => setProgress({ done, total }),
        });
        const ok = new Set(res.results.filter((r) => r.ok).map((r) => `${r.listId}:${r.itemId}`));
        onPatched(
          rows.filter((r) => ok.has(`${r.listId}:${r.id}`)).map((r) => ({ listId: r.listId, id: r.id, fields: { [column.name]: v } }))
        );
        setRuns((prev) => [{ at: new Date(), value: v, concurrency, ...res }, ...prev].slice(0, 20));
        return res;
      } catch (e) {
        console.error(e);
        setError(e);
        return null;
      } finally {
        setWriting(false);
      }
    },
    [column, keyColumn, concurrency, onPatched]
  );

  return { column, concurrency, setConcurrency, writing, progress, runs, error, write };
}

/** 일괄 입력 + 동시성 선택 + 측정 기록 표 */
export function WritePanel({ bench, selectedRows, onClearSelection }) {
  const [value, setValue] = useState('');
  const n = selectedRows.length;
  const label = bench.column?.displayName || BENCH_FIELD;

  const fill = async (v) => {
    if (!n) return;
    const what = v == null ? `비웁니다` : `'${v}' 로 씁니다`;
    if (!window.confirm(`선택한 ${n}행의 ${label} 를 ${what}. 동시성 ${bench.concurrency}.\n\n진행할까요?`)) return;
    const res = await bench.write(selectedRows, v);
    if (res && res.failCount === 0) onClearSelection();
  };

  if (!bench.column) {
    return (
      <div className="alert warn">
        TEMP 쓰기 측정: 그리드에 <span className="mono">{BENCH_FIELD}</span> 열이 없어 비활성입니다.
      </div>
    );
  }
  if (!isWritable(bench.column)) {
    return (
      <div className="alert warn">
        TEMP 쓰기 측정: <span className="mono">{BENCH_FIELD}</span> 열이 쓰기 불가 타입({bench.column.type})입니다.
      </div>
    );
  }

  const last = bench.runs[0];

  return (
    <div className="bench">
      <div className="row">
        <span className="badge warn">TEMP · 쓰기 속도 측정</span>
        <span className="muted">{label} 셀을 클릭하면 개별 입력, 행을 체크(Shift 로 범위)하면 일괄 입력</span>
      </div>
      <div className="row mt">
        <input
          type="text"
          placeholder={`${label} 값`}
          value={value}
          onChange={(e) => setValue(e.target.value)}
          disabled={bench.writing}
          style={{ width: 200 }}
          aria-label={`${label} 일괄 입력값`}
        />
        <button className="btn primary sm" disabled={bench.writing || !n || !value.trim()} onClick={() => fill(value)}>
          선택 {n}행에 채우기
        </button>
        <button className="btn danger sm" disabled={bench.writing || !n} onClick={() => fill(null)}>
          선택 {n}행 비우기
        </button>
        <label className="field">
          <span>동시성</span>
          <select value={bench.concurrency} onChange={(e) => setConcurrency(bench, e.target.value)} disabled={bench.writing}>
            {[1, 2, 4, 8].map((c) => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </select>
        </label>
        <span className="muted">배치 {CONFIG.defaults.batchSize}건 고정</span>
        {bench.writing ? (
          <>
            <progress value={bench.progress.done} max={bench.progress.total || 1} />
            <span className="muted">
              {bench.progress.done}/{bench.progress.total} 배치
            </span>
          </>
        ) : null}
      </div>

      {bench.error ? (
        <div className="alert err mt">
          <b>쓰기 실패</b> — {bench.error.message || String(bench.error)}
        </div>
      ) : null}

      {last ? (
        <div className="stats mt">
          <Stat label="건수" value={last.itemCount} />
          <Stat label="총 소요" value={ms(last.wallMs)} cls="hi" />
          <Stat label="건당" value={ms(last.msPerItem)} cls="hi" />
          <Stat label="초당" value={last.itemsPerSec.toFixed(1)} />
          <Stat label="배치" value={`${last.batchCount} × ${last.concurrency}`} />
          <Stat label="429" value={last.throttled} cls={last.throttled ? 'bad' : ''} />
          <Stat label="실패" value={last.failCount} cls={last.failCount ? 'bad' : 'good'} />
        </div>
      ) : null}

      {last?.failCount ? (
        <pre className="alert warn mt" style={{ margin: '8px 0 0' }}>
          {last.results
            .filter((r) => !r.ok)
            .slice(0, 5)
            .map((r) => `${r.key}: HTTP ${r.status} ${r.error ?? ''}`)
            .join('\n')}
        </pre>
      ) : null}

      {bench.runs.length > 1 ? (
        <div className="scroll mt" style={{ maxHeight: 220 }}>
          <table>
            <thead>
              <tr>
                <th>시각</th>
                <th>값</th>
                <th className="num">건수</th>
                <th className="num">동시성</th>
                <th className="num">총 소요</th>
                <th className="num">건당</th>
                <th className="num">초당</th>
                <th className="num">배치 p50 / p95</th>
                <th className="num">429</th>
                <th className="num">실패</th>
              </tr>
            </thead>
            <tbody>
              {bench.runs.map((r, i) => (
                <tr key={i}>
                  <td>{hhmmss(r.at)}</td>
                  <td className="mono">{r.value ?? '(비움)'}</td>
                  <td className="num">{r.itemCount}</td>
                  <td className="num">{r.concurrency}</td>
                  <td className="num">{ms(r.wallMs)}</td>
                  <td className="num">{ms(r.msPerItem)}</td>
                  <td className="num">{r.itemsPerSec.toFixed(1)}</td>
                  <td className="num">
                    {ms(r.batchStats.p50)} / {ms(r.batchStats.p95)}
                  </td>
                  <td className="num">{r.throttled}</td>
                  <td className="num">{r.failCount}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : null}
    </div>
  );
}

const setConcurrency = (bench, v) => bench.setConcurrency(Math.min(8, Math.max(1, Number(v) || 1)));

const Stat = ({ label, value, cls = '' }) => (
  <div className={`stat ${cls}`}>
    <div className="label">{label}</div>
    <div className="value">{value}</div>
  </div>
);
