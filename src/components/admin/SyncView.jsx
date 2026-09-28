import Setup from '../Setup.jsx';
import Compare from '../Compare.jsx';

/** Excel → List 동기화. ① 연결·대응표 → ② 비교·적용 두 단계 */
export default function SyncView({ sync, busy }) {
  const { tab, setTab, diff } = sync;
  return (
    <>
      <nav className="tabs">
        <button aria-selected={tab === 'setup'} onClick={() => setTab('setup')}>
          ① 연결 · 대응표
        </button>
        <button aria-selected={tab === 'compare'} onClick={() => setTab('compare')} disabled={!diff}>
          ② 비교 (전/후)
        </button>
      </nav>

      {tab === 'setup' ? (
        <Setup
          account
          busy={busy}
          schema={sync.schema}
          sheets={sync.sheets}
          sheet={sync.sheet}
          excelHeaders={sync.excelHeaders}
          mapping={sync.mapping}
          mappingReport={sync.mappingReport}
          mappingReady={sync.mappingReady}
          onConnect={sync.connect}
          onSheetChange={(s) => sync.connect(s)}
          onRunDiff={sync.runDiff}
        />
      ) : (
        <Compare diff={diff} lastRead={sync.lastRead} onApplied={() => sync.runDiff(false)} />
      )}
    </>
  );
}
