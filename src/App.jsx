import { useCallback, useEffect, useState } from 'react';
import { CONFIG, redirectUri } from './lib/config.js';
import { getAccount, initAuth, login, logout } from './lib/auth.js';
import { ctx, readExcelViaWorkbook, readListItems, resolveAll } from './lib/sources.js';
import { buildDiff, resolveMapping } from './lib/diff.js';
import Setup from './components/Setup.jsx';
import Compare from './components/Compare.jsx';

const emptyMapping = { keyExcel: '', keyList: '', pairs: [] };

export default function App() {
  const [account, setAccount] = useState(null);
  const [booting, setBooting] = useState(true);
  const [busy, setBusy] = useState('');
  const [error, setError] = useState(null);

  const [schema, setSchema] = useState(null); // ctx 스냅샷 (React 재렌더용)
  const [sheets, setSheets] = useState([]);
  const [sheet, setSheet] = useState(null);
  const [excelHeaders, setExcelHeaders] = useState([]);
  const [mapping, setMapping] = useState(emptyMapping); // CONFIG.mapping 을 실제 열 이름으로 해석한 결과
  const [mappingReport, setMappingReport] = useState(null); // resolveMapping() 결과 전체 (화면 표시용)

  const [diff, setDiff] = useState(null);
  const [lastRead, setLastRead] = useState(null);
  const [tab, setTab] = useState('setup');

  useEffect(() => {
    initAuth()
      .then(setAccount)
      .catch((e) => {
        // 원인별로 다른 힌트를 준다. 항상 "리디렉션 URI 확인" 만 띄우면 엉뚱한 곳을 보게 된다.
        const code = e.errorCode || e.name || '';
        const text = `${code} ${e.message}`;
        const hint = /65001|90094|consent/i.test(text)
          ? '동의되지 않은 scope 를 요청했습니다. 앱 등록의 API 권한(관리자 동의)과 config.js 의 scopes 가 문자 단위로 같아야 합니다.'
          : /50011|redirect_uri/i.test(text)
            ? `앱 등록의 SPA 리디렉션 URI 가 '${redirectUri()}' 와 문자 단위로 일치하는지 확인하세요.`
            : '브라우저 콘솔에 Content-Security-Policy 위반 메시지가 있는지도 확인하세요.';
        setError(new Error(`MSAL 초기화 실패 [${code}]: ${e.message} — ${hint}`));
      })
      .finally(() => setBooting(false));
  }, []);

  const run = useCallback(async (label, fn) => {
    setBusy(label);
    setError(null);
    try {
      await fn();
    } catch (e) {
      console.error(e);
      setError(e);
    } finally {
      setBusy('');
    }
  }, []);

  /** 사이트·리스트·파일 리졸브 + Excel 헤더 읽기 + 키 자동 매칭 */
  const connect = useCallback(
    (targetSheet = null) =>
      run('사이트 · 리스트 · 파일 리졸브 중', async () => {
        await resolveAll();
        const wb = await readExcelViaWorkbook(targetSheet);

        setSchema({
          columns: ctx.columns,
          listTitle: ctx.listTitle,
          lists: ctx.lists.map((l) => ({ name: l.name, title: l.title, columnCount: l.columns.length })),
          columnMismatch: ctx.columnMismatch,
          fileItem: ctx.fileItem,
        });
        setSheets(wb.sheets);
        setSheet(wb.sheet);
        setExcelHeaders(wb.headers);

        // 대응표는 고정(CONFIG.mapping). 여기서는 실제 열 이름으로 해석만 하고, 하나라도 못 찾으면 멈춘다.
        const report = resolveMapping(CONFIG.mapping, wb.headers, ctx.columns);
        setMapping(report.mapping);
        setMappingReport(report);
        if (!report.ok) {
          const parts = [];
          if (report.missing.list.length) parts.push(`리스트('${ctx.listTitle}') 열: ${report.missing.list.join(', ')}`);
          if (report.missing.excel.length) parts.push(`Excel 시트('${wb.sheet}') 헤더: ${report.missing.excel.join(', ')}`);
          throw new Error(
            `대응표의 열을 찾지 못했습니다 — ${parts.join(' / ')}. config.js 의 mapping 과 실제 열 이름을 맞추세요.`
          );
        }
      }),
    [run]
  );

  /** 리스트 + Excel 을 읽어 차이를 계산 */
  const runDiff = useCallback(
    (goToCompare = true) =>
      run('리스트 · Excel 읽는 중', async () => {
        const fields = [mapping.keyList, ...mapping.pairs.map((p) => p.list)].filter(Boolean);
        const list = await readListItems(fields);
        const excel = await readExcelViaWorkbook(sheet);
        setDiff(buildDiff(excel, list, mapping));
        setLastRead({ list, excel });
        if (goToCompare) setTab('compare');
      }),
    [mapping, run, sheet]
  );

  const mappingReady = Boolean(mappingReport?.ok && mapping.keyExcel && mapping.keyList && mapping.pairs.length);

  return (
    <>
      <header className="app">
        <h1>
          KUFRI <span className="muted">· Excel → SharePoint List 동기화</span>
        </h1>
        <div className="spacer" />
        <span className="who">{account ? `${account.name || ''} <${account.username}>` : '로그인 필요'}</span>
        {account ? (
          <button className="btn" onClick={() => logout()}>
            로그아웃
          </button>
        ) : (
          <button className="btn" onClick={() => login()} disabled={booting}>
            Microsoft 계정으로 로그인
          </button>
        )}
      </header>

      <nav className="tabs">
        <button aria-selected={tab === 'setup'} onClick={() => setTab('setup')}>
          ① 연결 · 매핑
        </button>
        <button aria-selected={tab === 'compare'} onClick={() => setTab('compare')} disabled={!diff}>
          ② 비교 (전/후)
        </button>
      </nav>

      <main>
        {error && (
          <div className="alert err">
            <b>오류</b> — {error.message || String(error)}
            {error.detail && <pre>{String(error.detail).slice(0, 800)}</pre>}
          </div>
        )}
        {busy && (
          <div className="alert info">
            처리 중… <span className="muted">{busy}</span>
          </div>
        )}

        {tab === 'setup' ? (
          <Setup
            account={account}
            busy={Boolean(busy)}
            schema={schema}
            sheets={sheets}
            sheet={sheet}
            excelHeaders={excelHeaders}
            mapping={mapping}
            mappingReport={mappingReport}
            mappingReady={mappingReady}
            onConnect={connect}
            onSheetChange={(s) => connect(s)}
            onRunDiff={runDiff}
          />
        ) : (
          <Compare diff={diff} lastRead={lastRead} onApplied={() => runDiff(false)} />
        )}
      </main>
    </>
  );
}
