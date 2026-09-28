// Excel → List 동기화 화면의 상태와 흐름. App 에서 분리해 화면을 오가도 비교 결과가 유지되게 한다.
import { useCallback, useState } from 'react';
import { CONFIG } from '../lib/config.js';
import { ctx, readExcelViaWorkbook, readListItems, resolveAll } from '../lib/sources.js';
import { buildDiff, resolveMapping } from '../lib/diff.js';

const emptyMapping = { keyExcel: '', keyList: '', pairs: [] };

/** @param run  App 의 실행 래퍼 (busy/error 관리) */
export function useSyncFlow(run) {
  const [schema, setSchema] = useState(null); // ctx 스냅샷 (React 재렌더용)
  const [sheets, setSheets] = useState([]);
  const [sheet, setSheet] = useState(null);
  const [excelHeaders, setExcelHeaders] = useState([]);
  const [mapping, setMapping] = useState(emptyMapping); // CONFIG.mapping 을 실제 열 이름으로 해석한 결과
  const [mappingReport, setMappingReport] = useState(null); // resolveMapping() 결과 전체 (화면 표시용)
  const [diff, setDiff] = useState(null);
  const [lastRead, setLastRead] = useState(null);
  const [tab, setTab] = useState('setup');

  /** 사이트·리스트·파일 리졸브 + Excel 헤더 읽기 + 대응표 해석. 시트만 바꿀 때는 리졸브를 다시 하지 않는다 */
  const connect = useCallback(
    (targetSheet = null) =>
      run('사이트 · 리스트 · 파일 리졸브 중', async () => {
        await resolveAll({ force: targetSheet == null });
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

  return {
    schema,
    sheets,
    sheet,
    excelHeaders,
    mapping,
    mappingReport,
    mappingReady,
    diff,
    lastRead,
    tab,
    setTab,
    connect,
    runDiff,
  };
}
