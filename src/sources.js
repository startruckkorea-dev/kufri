// 데이터 소스 계층 — SharePoint 리스트 읽기 / Excel 파일 읽기(2가지 방식)
import { CONFIG } from './config.js';
import { gfetch, gfetchAll, telemetry } from './graph.js';

/** 세션 동안 재사용하는 리졸브 결과 */
export const ctx = {
  siteId: null,
  listId: null,
  listTitle: null,
  columns: [], // {name, displayName, type, readOnly, dateFormat, choices}
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
  const target = (lists.value || []).find(
    (l) => l.name === CONFIG.listName || l.displayName === CONFIG.listName
  );
  if (!target) {
    throw new Error(
      `리스트 '${CONFIG.listName}' 를 찾지 못했습니다. 사이트에 존재하는 리스트: ` +
        (lists.value || []).map((l) => l.name).join(', ')
    );
  }
  ctx.listId = target.id;
  ctx.listTitle = target.displayName || target.name;

  const cols = await gfetch(`/sites/${ctx.siteId}/lists/${ctx.listId}/columns`, { name: 'resolve.columns' });
  ctx.columns = (cols.value || [])
    .filter((c) => !c.hidden)
    .map((c) => ({
      name: c.name,
      displayName: c.displayName,
      readOnly: !!c.readOnly,
      type: columnType(c),
      dateFormat: c.dateTime?.format || null, // 'dateOnly' | 'dateTime'
      choices: c.choice?.choices || null,
    }));

  const drive = await gfetch(`/sites/${ctx.siteId}/drive?$select=id,name,webUrl`, { name: 'resolve.drive' });
  ctx.driveId = drive.id;

  // 확장자를 모르므로 폴더를 나열해 basename 으로 찾는다.
  const folderPath = CONFIG.fileFolder ? `:/${encodeURI(CONFIG.fileFolder)}:` : '';
  const children = await gfetch(
    `/drives/${ctx.driveId}/root${folderPath}/children?$select=id,name,size,webUrl,lastModifiedDateTime`,
    { name: 'resolve.folder' }
  );
  const base = CONFIG.fileBaseName.toLowerCase();
  const file =
    (children.value || []).find((f) => stripExt(f.name).toLowerCase() === base) ||
    (children.value || []).find((f) => f.name.toLowerCase().startsWith(base));
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

/** 쓰기 가능한 컬럼만 매핑 대상으로 노출 */
export const writableColumns = () =>
  ctx.columns.filter((c) => !c.readOnly && !['lookup', 'person', 'calculated'].includes(c.type));

// ---------------------------------------------------------------- 리스트 읽기

/**
 * 리스트 아이템 전체 읽기.
 * @param {string[]|null} fieldNames 지정 시 $select 로 필요한 필드만 (읽기 속도 최적화)
 */
export async function readListItems(fieldNames = null) {
  const t0 = performance.now();
  const expand = fieldNames?.length
    ? `fields($select=${[...new Set(fieldNames)].join(',')})`
    : 'fields';
  const url =
    `/sites/${ctx.siteId}/lists/${ctx.listId}/items` +
    `?$expand=${expand}&$select=id,lastModifiedDateTime&$top=${CONFIG.defaults.listPageSize}`;

  const { items, pages } = await gfetchAll(url, { name: 'read.list' });
  const ms = performance.now() - t0;

  return {
    items: items.map((it) => ({ id: it.id, fields: it.fields || {}, lastModified: it.lastModifiedDateTime })),
    ms,
    pages,
    count: items.length,
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

// ---------------------------------------------------------------- Excel 읽기 B: 다운로드 + SheetJS

/** 파일을 통째로 받아 브라우저에서 파싱한다. 다운로드/파싱 시간을 분리 계측. */
export async function readExcelViaDownload(sheetName = null) {
  const t0 = performance.now();

  const meta = await gfetch(
    `/drives/${ctx.driveId}/items/${ctx.fileItem.id}?$select=id,name,size,@microsoft.graph.downloadUrl`,
    { name: 'excel.dl.meta' }
  );
  const downloadUrl = meta['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) throw new Error('다운로드 URL을 받지 못했습니다.');

  // downloadUrl 은 사전 인증된 URL — Authorization 헤더를 붙이면 안 된다(CORS preflight 유발).
  const tDl = performance.now();
  const res = await fetch(downloadUrl);
  if (!res.ok) throw new Error(`파일 다운로드 실패 (HTTP ${res.status})`);
  const buf = await res.arrayBuffer();
  const downloadMs = performance.now() - tDl;
  telemetry.add({
    name: 'excel.dl.download',
    method: 'GET',
    url: '(preauth download url)',
    ms: downloadMs,
    status: res.status,
    bytes: buf.byteLength,
  });

  const tParse = performance.now();
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const target = sheetName && wb.SheetNames.includes(sheetName) ? sheetName : wb.SheetNames[0];
  const grid = XLSX.utils.sheet_to_json(wb.Sheets[target], { header: 1, raw: true, defval: null, blankrows: false });
  const parseMs = performance.now() - tParse;
  telemetry.add({ name: 'excel.dl.parse', method: 'CPU', url: '(SheetJS)', ms: parseMs, status: 200 });

  return {
    ...gridToTable(grid),
    ms: performance.now() - t0,
    method: '다운로드 + SheetJS',
    sheets: wb.SheetNames,
    sheet: target,
    bytes: buf.byteLength,
    downloadMs,
    parseMs,
  };
}

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
