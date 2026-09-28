// 전체 데이터 그리드의 데이터 로딩. 열 정의(GRID_COLUMNS)에 있는 열만 $select 로 읽는다 (체크리스트 2-02).
import { useCallback, useState } from 'react';
import { GRID_COLUMNS } from '../lib/columns.js';
import { ctx, readListItems, resolveLists, resolveSite } from '../lib/sources.js';
// CONFIG 는 여기서 쓰지 않는다. 그리드 열은 columns.js 가, 동기화 열은 config.js 의 mapping 이 정한다.
import { findColumn } from '../lib/diff.js';

// Graph /columns 에 hidden=false 로 오는 SharePoint 기본 열. 업무 데이터가 아니므로 그리드에서 뺀다.
// Title 은 리스트 기본 열이지만 이 리스트는 commission_no 를 키로 쓰므로 함께 뺀다.
const SYSTEM_COLUMNS = new Set([
  'ID', 'Title', 'ContentType', 'Attachments', 'Edit', 'DocIcon', 'LinkTitle', 'LinkTitleNoMenu',
  'ItemChildCount', 'FolderChildCount', 'ComplianceAssetId', 'AppAuthor', 'AppEditor',
  'Created', 'Modified', 'Author', 'Editor',
]);
const isSystemColumn = (c) =>
  SYSTEM_COLUMNS.has(c.name) || c.name.startsWith('_') || c.type === 'person' || c.type === 'lookup';

/** @param run  App 의 실행 래퍼 (busy/error 관리) */
export function useListGrid(run) {
  const [data, setData] = useState(null); // {columns, rows, perList, count, ms, readAt, missing}

  const load = useCallback(
    (force = false) =>
      run('리스트 읽는 중', async () => {
        await resolveSite({ force });
        await resolveLists({ force });

        // 열 순서·표시명 = GRID_COLUMNS. 내부명(list)으로 찾고, 없으면 표시명(label)으로 한 번 더.
        // 그래도 없는 열은 건너뛰고 이름을 남겨 화면에서 알린다.
        const columns = [];
        const missing = [];
        const used = new Set();
        for (const s of GRID_COLUMNS) {
          const col = findColumn(ctx.columns, s.list) || findColumn(ctx.columns, s.label);
          if (col && !used.has(col.name)) {
            used.add(col.name);
            columns.push({ ...col, label: s.label, isKey: !!s.key });
          } else if (!col) missing.push(`${s.label} (${s.list})`);
        }
        // 리스트에는 있지만 열 정의에 없는 업무 열은 맨 뒤에 모아 리스트 표시명으로 보여준다.
        // SharePoint 시스템 열(ID·작성자·수정일 등)은 제외한다.
        for (const col of ctx.columns) {
          if (used.has(col.name) || isSystemColumn(col)) continue;
          columns.push({ ...col, label: col.displayName || col.name, extra: true });
        }

        const list = await readListItems(columns.map((c) => c.name));
        setData({
          columns,
          rows: list.items,
          perList: list.perList,
          count: list.count,
          ms: list.ms,
          readAt: new Date(),
          missing,
        });
      }),
    [run]
  );

  return { data, load };
}
