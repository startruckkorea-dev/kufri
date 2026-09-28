import { ADMIN, ADMIN_SUBS, hrefFor } from '../../lib/nav.js';
import ListGrid from './ListGrid.jsx';
import SyncView from './SyncView.jsx';

/** 관리자 탭. 하위 화면: 전체 데이터 그리드 / Excel 동기화 */
export default function Admin({ sub, sync, grid, busy }) {
  return (
    <>
      <nav className="tabs sub" aria-label={ADMIN.title}>
        {ADMIN_SUBS.map((s) => (
          <a key={s.key} href={hrefFor(ADMIN, s.key)} aria-selected={sub === s.key}>
            {s.title}
          </a>
        ))}
      </nav>
      {sub === 'sync' ? (
        <SyncView sync={sync} busy={busy} />
      ) : (
        <ListGrid data={grid.data} busy={busy} onRefresh={() => grid.load(true)} onPatched={grid.patchRows} />
      )}
    </>
  );
}
