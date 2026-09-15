// 데이터 소스 계층 — SharePoint 리스트 읽기 / Excel 파일 읽기(2가지 방식)
import { CONFIG } from './config.js';
import { gfetch, gfetchAll } from './graph.js';

/** 세션 동안 재사용하는 리졸브 결과 */
export const ctx = {
  siteId: null,
  lists: [], // [{id, name, title, columns}] — 행으로 나뉜 리스트들. 순서는 CONFIG.listNames
  listTitle: null, // 표시용. 여러 개면 ' + ' 로 잇는다
  columns: [], // 모든 리스트에 공통인 열 {name, displayName, type, readOnly, dateFormat, choices}
  columnMismatch: [], // 일부 리스트에만 있는 열 이름 (매핑 대상에서 제외)
  driveId: null,
  fileItem: null, // {id, name, size, webUrl}
};

// ---------------------------------------------------------------- 리졸브

export async function resolveAll() {
  const t0 = performance.now();

  const site = await gfetch(`/sites/${CONFIG.hostname}:${CONFIG.sitePath}`, { name: 'resolve.site' });
  ctx.siteId = site.id;

  const lists = await gfetch(`/sites/${ctx.siteId}/lists?$select=id,name,displayName,webUrl`, {
    name: 'resolve.lists',
  });
  const all = lists.value || [];
  const found = CONFIG.listNames.map((want) => all.find((l) => l.name === want || l.displayName === want) || null);
  const missing = CONFIG.listNames.filter((_, i) => !found[i]);
  if (missing.length) {
    throw new Error(
      `리스트 '${missing.join("', '")}' 를 찾지 못했습니다. 사이트에 존재하는 리스트: ` +
        all.map((l) => l.name).join(', ')
    );
  }

  // 리스트마다 열을 읽는다. 순차 호출 — 리스트 수는 적고 부하를 평탄하게 유지한다
  ctx.lists = [];
  for (const l of found) {
    const cols = await gfetch(`/sites/${ctx.siteId}/lists/${l.id}/columns`, { name: `resolve.columns:${l.name}` });
    ctx.lists.push({ id: l.id, name: l.name, title: l.displayName || l.name, columns: mapColumns(cols.value || []) });
  }
  ctx.listTitle = ctx.lists.map((l) => l.title).join(' + ');

  // 매핑 대상은 모든 리스트에 공통인 열만. 일부에만 있는 열은 이름을 남겨 화면에서 알린다.
  const [first, ...rest] = ctx.lists;
  ctx.columns = first.columns.filter((c) => rest.every((l) => l.columns.some((o) => o.name === c.name)));
  const common = new Set(ctx.columns.map((c) => c.name));
  ctx.columnMismatch = [
    ...new Set(ctx.lists.flatMap((l) => l.columns.map((c) => c.name)).filter((n) => !common.has(n))),
  ];

  const drive = await gfetch(`/sites/${ctx.siteId}/drive?$select=id,name,webUrl`, { name: 'resolve.drive' });
  ctx.driveId = drive.id;

  // 확장자를 모르므로 폴더를 나열해 basename 으로 찾는다.
  const folderPath = CONFIG.fileFolder ? `:/${encodeURI(CONFIG.fileFolder)}:` : '';
  const children = await gfetch(
    `/drives/${ctx.driveId}/root${folderPath}/children?$select=id,name,size,webUrl,lastModifiedDateTime`,
    { name: 'resolve.folder' }
  );
  const base = CONFIG.fileBaseName.toLowerCase();
  const files = children.value || [];
  const file =
    files.find((f) => stripExt(f.name).toLowerCase() === base) ||
    // 접두어 일치가 여럿이면 가장 최근에 수정된 파일을 쓴다 (예: SDISP_20260915.xlsx)
    files
      .filter((f) => f.name.toLowerCase().startsWith(base))
      .sort((a, b) => String(b.lastModifiedDateTime).localeCompare(String(a.lastModifiedDateTime)))[0];
  if (!file) {
    throw new Error(
      `'${CONFIG.fileFolder}' 폴더에서 '${CONFIG.fileBaseName}' 파일을 찾지 못했습니다. ` +
        `폴더 내 파일: ${(children.value || []).map((f) => f.name).join(', ')}`
    );
  }
  ctx.fileItem = file;

  return { ...ctx, resolveMs: performance.now() - t0 };
}

function columnType(c) {
  if (c.number) return 'number';
  if (c.currency) return 'number';
  if (c.dateTime) return 'dateTime';
  if (c.boolean) return 'boolean';
  if (c.choice) return 'choice';
  if (c.lookup) return 'lookup';
  if (c.personOrGroup) return 'person';
  if (c.calculated) return 'calculated';
  return 'text';
}

const stripExt = (n) => n.replace(/\.[^.]+$/, '');

/** Graph columnDefinition[] → 화면·비교에서 쓰는 축약형. 숨김 열 제외 */
const mapColumns = (raw) =>
  raw
    .filter((c) => !c.hidden)
    .map((c) => ({
      name: c.name,
      displayName: c.displayName,
      readOnly: !!c.readOnly,
      type: columnType(c),
      dateFormat: c.dateTime?.format || null, // 'dateOnly' | 'dateTime'
      choices: c.choice?.choices || null,
    }));

/** 쓰기 가능한 컬럼만 매핑 대상으로 노출 */
export const writableColumns = () =>
  ctx.columns.filter((c) => !c.readOnly && !['lookup', 'person', 'calculated'].includes(c.type));

// ---------------------------------------------------------------- 리스트 읽기

/**
 * 대상 리스트 전체의 아이템을 합쳐 읽는다. 항목마다 소속 리스트(listId)를 붙여
 * 쓰기 때 원래 리스트로 돌아가게 한다. 리스트는 순차로 읽는다 (부하 평탄화).
 * @param {string[]|null} fieldNames 지정 시 $select 로 필요한 필드만 (읽기 속도 최적화)
 */
export async function readListItems(fieldNames = null) {
  const t0 = performance.now();
  const expand = fieldNames?.length
    ? `fields($select=${[...new Set(fieldNames)].join(',')})`
    : 'fields';

  const items = [];
  const perList = [];
  let pages = 0;
  for (const l of ctx.lists) {
    const url =
      `/sites/${ctx.siteId}/lists/${l.id}/items` +
      `?$expand=${expand}&$select=id,lastModifiedDateTime&$top=${CONFIG.defaults.listPageSize}`;
    const r = await gfetchAll(url, { name: `read.list:${l.name}` });
    for (const it of r.items) {
      items.push({
        id: it.id,
        listId: l.id,
        listTitle: l.title,
        fields: it.fields || {},
        lastModified: it.lastModifiedDateTime,
      });
    }
    perList.push({ title: l.title, count: r.items.length, pages: r.pages });
    pages += r.pages;
  }
  const ms = performance.now() - t0;

  return {
    items,
    ms,
    pages,
    count: items.length,
    perList,
    mode: fieldNames?.length ? 'select 최적화' : '전체 필드',
  };
}

// ---------------------------------------------------------------- Excel 읽기 A: Workbook API

/**
 * Graph Workbook API 로 시트를 읽는다. 파일 다운로드 없음.
 * 세션(persistChanges:false)을 열어 후속 호출 지연을 줄인다.
 */
export async function readExcelViaWorkbook(sheetName = null) {
  const t0 = performance.now();
  const itemBase = `/drives/${ctx.driveId}/items/${ctx.fileItem.id}/workbook`;

  const session = await gfetch(`${itemBase}/createSession`, {
    method: 'POST',
    body: { persistChanges: false },
    name: 'excel.wb.session',
  });
  const sHeaders = { 'workbook-session-id': session.id };

  const sheets = await gfetch(`${itemBase}/worksheets?$select=id,name,position`, {
    headers: sHeaders,
    name: 'excel.wb.sheets',
  });
  const sheetList = (sheets.value || []).sort((a, b) => a.position - b.position);
  const sheet = sheetName ? sheetList.find((s) => s.name === sheetName) : sheetList[0];
  if (!sheet) throw new Error(`워크시트 '${sheetName}' 를 찾을 수 없습니다.`);

  // valuesOnly=true : 서식/수식 payload 제외 → 응답 크기와 서버 처리 시간 감소
  const range = await gfetch(
    `${itemBase}/worksheets/${encodeURIComponent(sheet.id)}/usedRange(valuesOnly=true)?$select=values,address,rowCount,columnCount`,
    { headers: sHeaders, name: 'excel.wb.usedRange' }
  );

  gfetch(`${itemBase}/closeSession`, { method: 'POST', headers: sHeaders, name: 'excel.wb.closeSession' }).catch(
    () => {}
  );

  const grid = range.values || [];
  return {
    ...gridToTable(grid),
    ms: performance.now() - t0,
    method: 'Workbook API',
    sheets: sheetList.map((s) => s.name),
    sheet: sheet.name,
    address: range.address,
  };
}

// Excel 읽기 B(다운로드 + SheetJS)는 실사이트에서 제외했다.
// PoC 의 읽기 속도 비교 전용이었고, npm 의 xlsx 는 취약점 권고가 걸려 있다(체크리스트 9-02).
// 필요해지면 poc-archive 브랜치의 readExcelViaDownload 를 참고할 것.

// ---------------------------------------------------------------- 공통

/** 2차원 배열 → {headers, rows(객체배열)} . 첫 행을 헤더로 본다. */
function gridToTable(grid) {
  if (!grid.length) return { headers: [], rows: [] };
  const headers = (grid[0] || []).map((h, i) => (h == null || h === '' ? `col${i + 1}` : String(h).trim()));
  const rows = [];
  for (let r = 1; r < grid.length; r++) {
    const arr = grid[r] || [];
    if (arr.every((v) => v == null || v === '')) continue;
    const obj = {};
    headers.forEach((h, i) => (obj[h] = arr[i] ?? null));
    obj.__row = r + 1; // 엑셀 실제 행번호 (1-based, 헤더 포함)
    rows.push(obj);
  }
  return { headers, rows };
}
