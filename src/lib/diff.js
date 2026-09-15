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

  // 리스트를 키로 인덱싱 (중복 키 감지 포함). 여러 리스트를 합친 상태이므로
  // 같은 키가 두 리스트에 모두 있으면 여기서 중복으로 잡힌다 — 첫 리스트의 항목만 대상.
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
    // listId: 행이 원래 있던 리스트. 쓰기는 이 리스트로 간다 (행으로 나뉜 리스트 묶음)
    const entry = {
      key: k,
      itemId: item.id,
      listId: item.listId,
      listTitle: item.listTitle,
      excelRow: row.__row,
      cells,
      diffCells,
    };
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

/**
 * 이름 비교용 정규화 — 대소문자/공백/마침표/언더스코어/괄호 차이를 모두 흡수.
 * 글자와 숫자만 남긴다 (유니코드 기준 — 독일어 움라우트, 한글 모두 보존).
 */
export const nameKey = (s) =>
  decodeSpName(s)
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]/gu, '');

/**
 * 키 컬럼명 하나를 양쪽 소스에서 이름으로 찾아낸다. (이름이 같은 소스 쌍용 — 지금의 SDISP 동기화는
 * 이름이 달라 resolveMapping 을 쓴다. 다른 동기화 정의에서 autoMap 과 함께 쓸 수 있게 남겨 둔다)
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

/**
 * 고정 대응표(CONFIG.mapping)를 실제 열 이름으로 해석한다.
 * list 쪽은 내부명/표시명 정확 일치 → 정규화 일치, excel 쪽은 헤더 정확 일치 → 정규화 일치 순.
 * 못 찾은 항목은 missing 에 모은다. 조용히 건너뛰지 않고 호출자가 오류로 멈춘다.
 * 쓸 수 없는 열(읽기 전용·lookup·person·calculated)은 readOnly 에 모으고 갱신 대상에서 뺀다.
 *
 * @returns {{ ok:boolean, mapping:{keyExcel,keyList,pairs}, missing:{list:string[],excel:string[]}, readOnly:string[], rows:Array }}
 */
export function resolveMapping(spec, excelHeaders, columns) {
  const findCol = (want) => {
    const exact = columns.find((c) => c.name === want || c.displayName === want);
    if (exact) return exact;
    const k = nameKey(want);
    return k ? columns.find((c) => nameKey(c.name) === k || nameKey(c.displayName) === k) || null : null;
  };
  const findHeader = (want) => {
    const exact = excelHeaders.find((h) => h === want);
    if (exact) return exact;
    const k = nameKey(want);
    return k ? excelHeaders.find((h) => nameKey(h) === k) || null : null;
  };
  const unwritable = (c) => c.readOnly || ['lookup', 'person', 'calculated'].includes(c.type);

  const missing = { list: [], excel: [] };
  const readOnly = [];
  const rows = []; // 화면 표시용: 대응표 한 줄씩의 해석 결과

  const key = { col: findCol(spec.key.list), header: findHeader(spec.key.excel) };
  if (!key.col) missing.list.push(spec.key.list);
  if (!key.header) missing.excel.push(spec.key.excel);
  rows.push({ isKey: true, spec: spec.key, col: key.col, header: key.header, status: key.col && key.header ? 'ok' : 'missing' });

  const pairs = [];
  for (const p of spec.pairs) {
    const col = findCol(p.list);
    const header = findHeader(p.excel);
    if (!col) missing.list.push(p.list);
    if (!header) missing.excel.push(p.excel);
    let status = 'ok';
    if (!col || !header) status = 'missing';
    else if (unwritable(col)) {
      status = 'readonly';
      readOnly.push(col.displayName || col.name);
    } else pairs.push({ excel: header, list: col.name });
    rows.push({ isKey: false, spec: p, col, header, status });
  }

  return {
    ok: !missing.list.length && !missing.excel.length,
    mapping: { keyExcel: key.header || '', keyList: key.col?.name || '', pairs },
    missing,
    readOnly,
    rows,
  };
}
