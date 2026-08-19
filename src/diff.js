// 비교 엔진 — Excel(원본) 기준으로 SharePoint 리스트와의 차이를 찾는다.
// 동기화 범위: "값이 다른 항목만 갱신" (신규 추가/삭제 없음)
import { ctx } from './sources.js';

const EXCEL_EPOCH = Date.UTC(1899, 11, 30); // 1900 날짜체계 (Excel 의 1900 윤년 버그 보정 포함)

/** Excel 시리얼/문자열/Date → JS Date (실패 시 null) */
export function toDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === 'number') {
    if (v < 1 || v > 100000) return null;
    return new Date(EXCEL_EPOCH + Math.round(v * 86400000));
  }
  const s = String(v).trim();
  const d = new Date(/^\d{4}-\d{2}-\d{2}$/.test(s) ? s + 'T00:00:00Z' : s);
  return isNaN(d) ? null : d;
}

const num = (v) => {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[,\s₩]/g, ''));
  return Number.isFinite(n) ? n : null;
};

const bool = (v) => {
  if (v == null || v === '') return null;
  if (typeof v === 'boolean') return v;
  const s = String(v).trim().toLowerCase();
  if (['true', '1', 'y', 'yes', 'o', '예', '참'].includes(s)) return true;
  if (['false', '0', 'n', 'no', 'x', '아니오', '거짓'].includes(s)) return false;
  return null;
};

/**
 * 비교용 정규화 문자열. 타입별로 표현 차이(1 vs "1", 날짜 포맷 등)를 흡수한다.
 * 정규화 결과가 같으면 "변경 없음"으로 본다.
 */
export function normalize(value, col) {
  if (value == null || value === '') return '';
  switch (col?.type) {
    case 'number': {
      const n = num(value);
      return n == null ? String(value).trim() : String(n);
    }
    case 'dateTime': {
      const d = toDate(value);
      if (!d) return String(value).trim();
      return col.dateFormat === 'dateOnly'
        ? d.toISOString().slice(0, 10)
        : d.toISOString().slice(0, 16); // 분 단위까지만 비교 (초/ms 노이즈 제거)
    }
    case 'boolean': {
      const b = bool(value);
      return b == null ? String(value).trim() : String(b);
    }
    default:
      return String(value).trim().replace(/\s+/g, ' ');
  }
}

/** SharePoint 에 PATCH 할 실제 JSON 값으로 변환 */
export function toFieldValue(value, col) {
  if (value == null || value === '') return null;
  switch (col?.type) {
    case 'number':
      return num(value);
    case 'dateTime': {
      const d = toDate(value);
      return d ? d.toISOString() : null;
    }
    case 'boolean':
      return bool(value);
    default:
      return String(value).trim();
  }
}

/** 화면 표시용 문자열 */
export function display(value, col) {
  if (value == null || value === '') return '(비어 있음)';
  if (col?.type === 'dateTime') {
    const d = toDate(value);
    if (d) return col.dateFormat === 'dateOnly' ? d.toISOString().slice(0, 10) : d.toISOString().slice(0, 16).replace('T', ' ');
  }
  if (col?.type === 'boolean') {
    const b = bool(value);
    if (b != null) return b ? '예' : '아니오';
  }
  return String(value);
}

const colOf = (listField) => ctx.columns.find((c) => c.name === listField) || { name: listField, type: 'text' };

/**
 * @param {object} excel  {headers, rows}
 * @param {object} list   {items:[{id, fields}]}
 * @param {object} mapping {keyExcel, keyList, pairs:[{excel,list}]}
 */
export function buildDiff(excel, list, mapping) {
  const keyCol = colOf(mapping.keyList);
  const pairs = mapping.pairs.filter((p) => p.excel && p.list);

  // 리스트를 키로 인덱싱 (중복 키 감지 포함)
  const index = new Map();
  const dupListKeys = [];
  for (const item of list.items) {
    const k = normalize(item.fields[mapping.keyList], keyCol);
    if (!k) continue;
    if (index.has(k)) dupListKeys.push(k);
    else index.set(k, item);
  }

  const changed = [];
  const unchanged = [];
  const excelOnly = [];
  const dupExcelKeys = [];
  const seenExcel = new Set();
  const matchedIds = new Set();

  for (const row of excel.rows) {
    const k = normalize(row[mapping.keyExcel], keyCol);
    if (!k) continue;
    if (seenExcel.has(k)) dupExcelKeys.push(k);
    seenExcel.add(k);

    const item = index.get(k);
    if (!item) {
      excelOnly.push({ key: k, excelRow: row.__row, row });
      continue;
    }
    matchedIds.add(item.id);

    const cells = [];
    for (const p of pairs) {
      const col = colOf(p.list);
      const before = item.fields[p.list];
      const after = row[p.excel];
      const nb = normalize(before, col);
      const na = normalize(after, col);
      cells.push({
        field: p.list,
        label: col.displayName || p.list,
        excelHeader: p.excel,
        type: col.type,
        before,
        after,
        beforeText: display(before, col),
        afterText: display(after, col),
        changed: nb !== na,
        writeValue: toFieldValue(after, col),
      });
    }

    const diffCells = cells.filter((c) => c.changed);
    const entry = { key: k, itemId: item.id, excelRow: row.__row, cells, diffCells };
    if (diffCells.length) changed.push(entry);
    else unchanged.push(entry);
  }

  const listOnly = list.items.filter((i) => {
    const k = normalize(i.fields[mapping.keyList], keyCol);
    return k && !matchedIds.has(i.id);
  });

  return {
    changed,
    unchanged,
    excelOnly,
    listOnly,
    warnings: {
      dupExcelKeys: [...new Set(dupExcelKeys)],
      dupListKeys: [...new Set(dupListKeys)],
    },
    summary: {
      excelRows: excel.rows.length,
      listItems: list.items.length,
      matched: changed.length + unchanged.length,
      changedRows: changed.length,
      changedCells: changed.reduce((a, c) => a + c.diffCells.length, 0),
      unchanged: unchanged.length,
      excelOnly: excelOnly.length,
      listOnly: listOnly.length,
    },
  };
}

// ---------------------------------------------------------------- 이름 매칭

/** SharePoint 내부 컬럼명의 _x0020_ 형태 이스케이프 복원 ('Commission_x0020_no' → 'Commission no') */
export const decodeSpName = (s) =>
  String(s ?? '').replace(/_x([0-9a-fA-F]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16)));

/** 이름 비교용 정규화 — 대소문자/공백/마침표/언더스코어/괄호 차이를 모두 흡수 */
export const nameKey = (s) =>
  decodeSpName(s)
    .toLowerCase()
    .replace(/[^a-z0-9가-힣]/g, '');

/**
 * 설정된 키 컬럼명(CONFIG.keyColumn)을 양쪽 소스에서 자동으로 찾아낸다.
 * 사용자가 고르지 않는다. 정확 일치 → (후보가 유일할 때만) 부분 일치 순으로 시도.
 */
export function resolveKeyColumns(excelHeaders, columns, keyName) {
  const want = nameKey(keyName);
  const uniqueHit = (arr, getters) => {
    for (const g of getters) {
      const exact = arr.filter((x) => nameKey(g(x)) === want);
      if (exact.length) return exact[0];
    }
    const partial = arr.filter((x) => getters.some((g) => nameKey(g(x)).startsWith(want)));
    return partial.length === 1 ? partial[0] : null;
  };

  const excelHit = uniqueHit(excelHeaders, [(h) => h]);
  const listHit = uniqueHit(columns, [(c) => c.displayName, (c) => c.name]);
  return { keyExcel: excelHit || '', keyList: listHit?.name || '' };
}

/** 헤더 이름으로 Excel↔List 컬럼 자동 매핑 초안 생성. 키 컬럼은 제외한다(갱신 대상 아님). */
export function autoMap(excelHeaders, columns, exclude = {}) {
  const pairs = [];
  for (const h of excelHeaders) {
    if (exclude.keyExcel && nameKey(h) === nameKey(exclude.keyExcel)) continue;
    const hit = columns.find((c) => nameKey(c.displayName) === nameKey(h) || nameKey(c.name) === nameKey(h));
    if (!hit) continue;
    if (exclude.keyList && hit.name === exclude.keyList) continue;
    pairs.push({ excel: h, list: hit.name });
  }
  return pairs;
}
