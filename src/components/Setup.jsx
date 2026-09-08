import { CONFIG } from '../lib/config.js';
import { autoMap } from '../lib/diff.js';

const kb = (b) => (b == null ? '—' : b >= 1048576 ? `${(b / 1048576).toFixed(2)} MB` : `${Math.round(b / 1024)} KB`);

export default function Setup({
  account,
  busy,
  schema,
  sheets,
  sheet,
  excelHeaders,
  keyError,
  mapping,
  mappingReady,
  onConnect,
  onSheetChange,
  onMappingChange,
  onRunDiff,
}) {
  if (!account) {
    return (
      <div className="card">
        <div className="empty">
          상단의 <b>Microsoft 계정으로 로그인</b> 을 눌러 시작하세요.
          <br />
          <span className="muted">로그인한 계정에 부여된 SharePoint 권한 범위에서만 동작합니다.</span>
        </div>
      </div>
    );
  }

  const columns = schema?.columns ?? [];
  const writable = columns.filter((c) => !c.readOnly && !['lookup', 'person', 'calculated'].includes(c.type));
  const pairCols = writable.filter((c) => c.name !== mapping.keyList);
  const pairHeaders = excelHeaders.filter((h) => h !== mapping.keyExcel);
  const keyColLabel = columns.find((c) => c.name === mapping.keyList)?.displayName || mapping.keyList;

  const setPair = (i, patch) =>
    onMappingChange({ ...mapping, pairs: mapping.pairs.map((p, n) => (n === i ? { ...p, ...patch } : p)) });

  return (
    <>
      <div className="card">
        <h2>
          대상 <span className="sub">src/lib/config.js · .env 로 변경 가능</span>
        </h2>
        <dl className="kv">
          <dt>사이트</dt>
          <dd>{CONFIG.hostname + CONFIG.sitePath}</dd>
          <dt>리스트</dt>
          <dd>
            {CONFIG.listName} {schema ? <span className="badge ok">연결됨</span> : null}
          </dd>
          <dt>파일 폴더</dt>
          <dd>Shared Documents/{CONFIG.fileFolder}</dd>
          <dt>파일</dt>
          <dd>
            {schema?.fileItem ? (
              <>
                {schema.fileItem.name} <span className="muted">({kb(schema.fileItem.size)})</span>
              </>
            ) : (
              <>
                {CONFIG.fileBaseName} <span className="muted">(확장자 탐색 예정)</span>
              </>
            )}
          </dd>
        </dl>
        <div className="row mt">
          <button className="btn primary" onClick={() => onConnect(null)} disabled={busy}>
            {schema ? '스키마 다시 읽기' : '연결 및 스키마 로드'}
          </button>
          {schema ? (
            <span className="badge ok">
              리스트 컬럼 {columns.length}개 · Excel 헤더 {excelHeaders.length}개
            </span>
          ) : null}
        </div>
      </div>

      {schema ? (
        <>
          <div className="card">
            <h2>Excel 워크시트</h2>
            <div className="row">
              <label className="field">
                <span>워크시트</span>
                <select value={sheet ?? ''} onChange={(e) => onSheetChange(e.target.value)} disabled={busy}>
                  {sheets.map((n) => (
                    <option key={n} value={n}>
                      {n}
                    </option>
                  ))}
                </select>
              </label>
              <span className="muted">Graph Workbook API 로 읽습니다 (파일 다운로드 없음).</span>
            </div>
          </div>

          <div className="card">
            <h2>
              컬럼 매핑 <span className="sub">Excel 헤더 → SharePoint 리스트 내부 컬럼명</span>
            </h2>

            {keyError ? (
              <>
                <div className="alert warn">{keyError}</div>
                <div className="row">
                  <label className="field">
                    <span>키 (Excel)</span>
                    <select
                      value={mapping.keyExcel}
                      onChange={(e) => onMappingChange({ ...mapping, keyExcel: e.target.value })}
                    >
                      <option value="">— 선택 —</option>
                      {excelHeaders.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label className="field">
                    <span>키 (List)</span>
                    <select
                      value={mapping.keyList}
                      onChange={(e) => onMappingChange({ ...mapping, keyList: e.target.value })}
                    >
                      <option value="">— 선택 —</option>
                      {columns.map((c) => (
                        <option key={c.name} value={c.name}>
                          {(c.displayName || c.name) + ' (' + c.type + ')'}
                        </option>
                      ))}
                    </select>
                  </label>
                </div>
              </>
            ) : (
              <>
                <div className="row">
                  <span className="badge ok">키 자동 매칭됨</span>
                  <span className="muted">기준 키</span> <b className="mono">{CONFIG.keyColumn}</b>
                  <span className="muted">·</span>
                  <span className="muted">Excel</span> <span className="mono">{mapping.keyExcel}</span>
                  <span className="muted">→ List</span> <span className="mono">{keyColLabel}</span>
                  <span className="muted">(내부명 {mapping.keyList})</span>
                </div>
                <div className="row">
                  <span className="muted">키는 매칭 기준이므로 갱신 대상에서 제외됩니다.</span>
                </div>
              </>
            )}

            <div className="scroll mt">
              <table>
                <thead>
                  <tr>
                    <th style={{ width: '34%' }}>Excel 헤더</th>
                    <th style={{ width: 24 }} />
                    <th style={{ width: '34%' }}>List 컬럼</th>
                    <th>타입</th>
                    <th style={{ width: 70 }} />
                  </tr>
                </thead>
                <tbody>
                  {mapping.pairs.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="empty">
                        매핑이 없습니다. 자동 매핑을 누르거나 행을 추가하세요.
                      </td>
                    </tr>
                  ) : null}
                  {mapping.pairs.map((p, i) => (
                    <tr key={i}>
                      <td>
                        <select value={p.excel} onChange={(e) => setPair(i, { excel: e.target.value })}>
                          <option value="">— 선택 —</option>
                          {pairHeaders.map((h) => (
                            <option key={h} value={h}>
                              {h}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td className="muted">→</td>
                      <td>
                        <select value={p.list} onChange={(e) => setPair(i, { list: e.target.value })}>
                          <option value="">— 선택 —</option>
                          {pairCols.map((c) => (
                            <option key={c.name} value={c.name}>
                              {(c.displayName || c.name) + ' (' + c.type + ')'}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td>
                        {p.list ? (
                          <span className="badge">{columns.find((c) => c.name === p.list)?.type ?? ''}</span>
                        ) : null}
                      </td>
                      <td>
                        <button
                          className="btn sm danger"
                          onClick={() =>
                            onMappingChange({ ...mapping, pairs: mapping.pairs.filter((_, n) => n !== i) })
                          }
                        >
                          삭제
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="row mt">
              <button
                className="btn"
                onClick={() => onMappingChange({ ...mapping, pairs: [...mapping.pairs, { excel: '', list: '' }] })}
              >
                행 추가
              </button>
              <button
                className="btn"
                onClick={() => onMappingChange({ ...mapping, pairs: autoMap(excelHeaders, writable, mapping) })}
              >
                이름으로 자동 매핑
              </button>
              <div style={{ flex: 1 }} />
              <button className="btn primary" onClick={onRunDiff} disabled={busy || !mappingReady}>
                데이터 읽기 &amp; 비교 실행
              </button>
            </div>
          </div>
        </>
      ) : null}
    </>
  );
}
