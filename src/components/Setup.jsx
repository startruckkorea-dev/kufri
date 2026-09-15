import { CONFIG } from '../lib/config.js';

const kb = (b) => (b == null ? '—' : b >= 1048576 ? `${(b / 1048576).toFixed(2)} MB` : `${Math.round(b / 1024)} KB`);

const STATUS = {
  ok: { cls: 'ok', text: '확인' },
  missing: { cls: 'err', text: '못 찾음' },
  readonly: { cls: 'warn', text: '쓰기 불가 · 제외' },
};

/**
 * 연결 화면. 대응표(CONFIG.mapping)는 고정이며 여기서는 해석 결과만 보여준다.
 * 사용자가 열을 고르는 UI 는 없다 — 못 찾은 열이 있으면 config.js 를 고쳐야 한다.
 */
export default function Setup({
  account,
  busy,
  schema,
  sheets,
  sheet,
  excelHeaders,
  mapping,
  mappingReport,
  mappingReady,
  onConnect,
  onSheetChange,
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
  const rows = mappingReport?.rows ?? [];

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
            {CONFIG.listNames.join(' + ')}{' '}
            {schema ? <span className="badge ok">{schema.lists.length}개 연결됨</span> : null}
            <div className="muted" style={{ fontFamily: 'inherit' }}>
              행으로 나뉜 리스트를 합쳐 비교하고, 각 행은 원래 있던 리스트에 씁니다.
            </div>
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
                {CONFIG.fileBaseName}* <span className="muted">(접두어 일치 중 최근 수정본)</span>
              </>
            )}
          </dd>
        </dl>
        {schema?.columnMismatch?.length ? (
          <div className="alert warn mt">
            일부 리스트에만 있는 열은 매핑 대상에서 제외했습니다:{' '}
            <span className="mono">{schema.columnMismatch.join(', ')}</span>
          </div>
        ) : null}
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
              열 대응표 <span className="sub">config.js 에 고정 · Excel 헤더 → SharePoint 리스트 열</span>
            </h2>

            {mappingReport && !mappingReport.ok ? (
              <div className="alert err">
                대응표의 열을 찾지 못해 비교를 시작할 수 없습니다. 아래 <b>못 찾음</b> 행의 이름을 실제 열 이름과
                맞춰 config.js 를 고치세요.
              </div>
            ) : null}
            {mappingReport?.readOnly.length ? (
              <div className="alert warn">
                쓸 수 없는 열이라 갱신 대상에서 뺐습니다: <span className="mono">{mappingReport.readOnly.join(', ')}</span>
              </div>
            ) : null}

            <div className="scroll mt">
              <table>
                <thead>
                  <tr>
                    <th style={{ width: '30%' }}>Excel 헤더</th>
                    <th style={{ width: 24 }} />
                    <th style={{ width: '30%' }}>List 열</th>
                    <th>타입</th>
                    <th style={{ width: 130 }}>상태</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, i) => {
                    const st = STATUS[r.status];
                    return (
                      <tr key={i}>
                        <td className="mono">
                          {r.spec.excel}
                          {r.header && r.header !== r.spec.excel ? (
                            <span className="muted"> (실제: {r.header})</span>
                          ) : null}
                        </td>
                        <td className="muted">→</td>
                        <td className="mono">
                          {r.col ? r.col.displayName || r.col.name : r.spec.list}
                          {r.col && r.col.name !== (r.col.displayName || r.col.name) ? (
                            <span className="muted"> (내부명 {r.col.name})</span>
                          ) : null}
                        </td>
                        <td>{r.col ? <span className="badge">{r.col.type}</span> : null}</td>
                        <td>
                          {r.isKey ? <span className="badge">키 · 갱신 안 함</span> : null}{' '}
                          <span className={`badge ${st.cls}`}>{st.text}</span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="row mt">
              <span className="muted">
                갱신 대상 {mapping.pairs.length}열 · 키 <span className="mono">{mapping.keyList || '—'}</span>
              </span>
              <div style={{ flex: 1 }} />
              <button className="btn primary" onClick={() => onRunDiff()} disabled={busy || !mappingReady}>
                데이터 읽기 &amp; 비교 실행
              </button>
            </div>
          </div>
        </>
      ) : null}
    </>
  );
}
