// 롤백 — SharePoint 리스트 항목의 버전 기록에서 적용 이전 값을 되찾는다.
//
// 적용 이력(localStorage)에 남은 '전' 값은 화면 표시용 문자열이라 역변환 손실이 있다.
// 버전 기록은 실제 저장됐던 원본 값을 그대로 돌려주므로 이쪽을 1순위로 쓴다.
//   GET  /sites/{site}/lists/{list}/items/{id}/versions?$expand=fields
import { gbatch, chunk, runPool } from './graph.js';
import { ctx } from './sources.js';
import { normalize, display, toFieldValue, nameKey } from './diff.js';

/** 버전 id('3.0','2.0'…) 를 숫자로 — 최신이 앞에 오도록 내림차순 정렬용 */
const verNum = (v) => parseFloat(v.id) || 0;

/** 브라우저 시각(적용 이력)과 SharePoint 서버 시각의 오차 여유 */
const CLOCK_SKEW_MS = 30_000;

/**
 * 항목별 버전 목록을 $batch 로 읽는다. (20건 묶음 × 동시 전송)
 * @returns {{versions: Map<string, Array>, errors: Array, ms: number}}
 */
export async function readItemVersions(itemIds, { concurrency = 4 } = {}) {
  const t0 = performance.now();
  const versions = new Map();
  const errors = [];
  const batches = chunk(itemIds, 20);

  const tasks = batches.map((ids, bi) => async () => {
    const requests = ids.map((id, i) => ({
      id: `${bi}-${i}`,
      method: 'GET',
      url: `/sites/${ctx.siteId}/lists/${ctx.listId}/items/${id}/versions?$expand=fields`,
    }));
    const responses = await gbatch(requests, `restore.versions#${bi}`);
    for (const r of responses) {
      const idx = Number(r.id.split('-')[1]);
      const itemId = ids[idx];
      if (r.status >= 200 && r.status < 300) {
        versions.set(itemId, [...(r.body?.value || [])].sort((a, b) => verNum(b) - verNum(a)));
      } else {
        errors.push({ itemId, status: r.status, error: r.body?.error?.message || JSON.stringify(r.body || {}).slice(0, 200) });
      }
    }
  });

  await runPool(tasks, concurrency);
  return { versions, errors, ms: performance.now() - t0 };
}

/**
 * 복원 계획 수립.
 * 각 항목의 최신 버전 = 현재값, 그 직전 버전 = 되돌릴 값.
 * @param {object} entry     적용 이력 1건 (changeLog)
 * @param {Map} versionsMap  readItemVersions 결과
 * @param {Array} cols       [{field, label, col}] 되돌릴 대상 컬럼 (키/읽기전용 제외)
 * @param {string} appliedBy 적용을 수행한 계정 (외부 수정 감지용)
 */
export function buildRestorePlan(entry, versionsMap, cols, appliedBy) {
  const rows = [];
  const noVersion = [];
  const foreignEdit = [];

  // 적용이 진행된 구간 [start, end]. entry.at 은 적용이 "끝난" 시각이므로
  // 그 구간보다 앞선 버전을 골라야 우리 쓰기가 만든 버전을 되돌림 대상으로 잡지 않는다.
  // 같은 항목에 두 번 이상 적용했더라도 이 방식이면 최초 적용 이전까지 거슬러 올라간다.
  const endAt = Date.parse(entry.at);
  const cutoff = endAt - (entry.wallMs || 0) - CLOCK_SKEW_MS;

  for (const item of entry.items) {
    if (!item.ok) continue; // 애초에 쓰기 실패한 건 되돌릴 것이 없다
    const list = versionsMap.get(item.itemId);

    if (!list || list.length < 2) {
      noVersion.push(item.key);
      continue;
    }

    const current = list[0];
    let ti = list.findIndex((v) => Date.parse(v.lastModifiedDateTime) < cutoff);
    if (ti < 0) ti = 1; // 시각으로 못 고르면 직전 버전으로 물러선다
    const target = list[ti];
    const skippedVersions = ti - 1; // 우리 쓰기 말고 추가로 건너뛴 버전 수

    // 최신 버전을 만든 사람이 적용자와 다르거나, 건너뛴 버전이 더 있으면 외부 수정이 끼어든 것이다.
    const curBy = current.lastModifiedBy?.user?.email || current.lastModifiedBy?.user?.displayName || '';
    const foreign = skippedVersions > 0 || !!(appliedBy && curBy && nameKey(curBy) !== nameKey(appliedBy));
    if (foreign) foreignEdit.push(item.key);

    const cells = cols.map((c) => {
      const cur = current.fields?.[c.field];
      const res = target.fields?.[c.field];
      return {
        field: c.field,
        label: c.label,
        current: cur,
        restore: res,
        currentText: display(cur, c.col),
        restoreText: display(res, c.col),
        changed: normalize(cur, c.col) !== normalize(res, c.col),
        writeValue: toFieldValue(res, c.col),
        // sync.applyChanges / buildPostApply 가 기대하는 필드명 별칭 (복원값이 곧 '적용될 값')
        after: res,
        afterText: display(res, c.col),
      };
    });

    const diffCells = cells.filter((c) => c.changed);
    rows.push({
      key: item.key,
      itemId: item.itemId,
      versionId: target.id,
      versionAt: target.lastModifiedDateTime,
      versionBy: target.lastModifiedBy?.user?.displayName || '',
      currentVersionId: current.id,
      currentAt: current.lastModifiedDateTime,
      currentBy: current.lastModifiedBy?.user?.displayName || '',
      foreign,
      skippedVersions,
      cells,
      diffCells, // sync.applyChanges 가 그대로 소비하는 형태
    });
  }

  return {
    rows,
    warnings: { noVersion, foreignEdit },
    summary: {
      items: rows.length,
      restorable: rows.filter((r) => r.diffCells.length).length,
      alreadySame: rows.filter((r) => !r.diffCells.length).length,
      cells: rows.reduce((a, r) => a + r.diffCells.length, 0),
      skipped: noVersion.length,
    },
  };
}

/**
 * 시점 기준 복원 계획 — 적용 이력이 없어도 쓸 수 있다.
 *
 * 어떤 항목을 건드렸는지 몰라도 된다. 전체 항목의 버전 기록에서 기준 시각 직전 버전을 찾아
 * 현재값과 비교하면, 건드리지 않은 항목은 "이미 동일"로 저절로 걸러진다.
 *
 * @param {Array}  listItems  리스트 전체 항목 (id, fields)
 * @param {Map}    versionsMap readItemVersions 결과
 * @param {Array}  cols       [{field, label, col}] 되돌릴 대상 컬럼
 * @param {number} cutoffMs   기준 시각 (이 시각 이전 상태로 되돌린다)
 * @param {string} keyField   키 컬럼 내부명
 * @param {string} expectedUser 이 계정 외의 수정이 섞였는지 표시하기 위한 기준 계정
 */
export function buildPointInTimePlan(listItems, versionsMap, cols, cutoffMs, keyField, expectedUser) {
  const keyCol = ctx.columns.find((c) => c.name === keyField) || { type: 'text' };
  const rows = [];
  const noVersion = [];
  const allAfterCutoff = [];
  const foreignEdit = [];
  let untouched = 0;

  for (const it of listItems) {
    const key = display(it.fields?.[keyField], keyCol);
    const vlist = versionsMap.get(it.id);

    if (!vlist || !vlist.length) {
      noVersion.push(key);
      continue;
    }

    const ti = vlist.findIndex((v) => Date.parse(v.lastModifiedDateTime) < cutoffMs);
    if (ti < 0) {
      // 기준 시각 이전 버전이 아예 없다 = 그 이후에 처음 만들어진 항목
      allAfterCutoff.push(key);
      continue;
    }
    if (ti === 0) {
      untouched++; // 기준 시각 이후 수정이 없다 = 손댈 것 없음
      continue;
    }

    const current = vlist[0];
    const target = vlist[ti];

    const cells = cols.map((c) => {
      const cur = current.fields?.[c.field];
      const res = target.fields?.[c.field];
      return {
        field: c.field,
        label: c.label,
        current: cur,
        restore: res,
        currentText: display(cur, c.col),
        restoreText: display(res, c.col),
        changed: normalize(cur, c.col) !== normalize(res, c.col),
        writeValue: toFieldValue(res, c.col),
        after: res,
        afterText: display(res, c.col),
      };
    });

    const diffCells = cells.filter((c) => c.changed);
    if (!diffCells.length) {
      untouched++; // 버전은 늘었지만 우리가 되돌릴 컬럼은 그대로
      continue;
    }

    const curBy = current.lastModifiedBy?.user?.email || current.lastModifiedBy?.user?.displayName || '';
    const foreign = !!(expectedUser && curBy && nameKey(curBy) !== nameKey(expectedUser));
    if (foreign) foreignEdit.push(key);

    rows.push({
      key,
      itemId: it.id,
      versionId: target.id,
      versionAt: target.lastModifiedDateTime,
      versionBy: target.lastModifiedBy?.user?.displayName || '',
      currentVersionId: current.id,
      currentAt: current.lastModifiedDateTime,
      currentBy: current.lastModifiedBy?.user?.displayName || '',
      foreign,
      skippedVersions: ti - 1,
      cells,
      diffCells,
    });
  }

  return {
    rows,
    warnings: { noVersion, foreignEdit, allAfterCutoff },
    summary: {
      items: listItems.length,
      restorable: rows.length,
      alreadySame: untouched,
      cells: rows.reduce((a, r) => a + r.diffCells.length, 0),
      skipped: noVersion.length + allAfterCutoff.length,
    },
  };
}

/**
 * 버전 기록을 못 쓸 때의 차선책 — 적용 이력에 남은 '전' 표시 문자열을 되돌린다.
 * 손실: 날짜의 초 단위(분까지만 기록됨). 그 외 텍스트/숫자/선택/예-아니오는 정확히 복원된다.
 */
export function buildRestorePlanFromLog(entry, listItems, cols, keyField) {
  const keyCol = ctx.columns.find((c) => c.name === keyField) || { type: 'text' };
  const byKey = new Map(listItems.map((it) => [normalize(it.fields[keyField], keyCol), it]));
  const labelToCol = new Map(cols.map((c) => [nameKey(c.label), c]));

  const rows = [];
  const missing = [];

  for (const item of entry.items) {
    if (!item.ok) continue;
    const live = byKey.get(item.key);
    if (!live) {
      missing.push(item.key);
      continue;
    }

    const cells = [];
    for (const logged of item.cells) {
      const c = labelToCol.get(nameKey(logged.label));
      if (!c) continue;
      const restoreRaw = parseDisplayText(logged.before, c.col);
      const cur = live.fields?.[c.field];
      cells.push({
        field: c.field,
        label: c.label,
        current: cur,
        restore: restoreRaw,
        currentText: display(cur, c.col),
        restoreText: display(restoreRaw, c.col),
        changed: normalize(cur, c.col) !== normalize(restoreRaw, c.col),
        writeValue: toFieldValue(restoreRaw, c.col),
        after: restoreRaw,
        afterText: display(restoreRaw, c.col),
      });
    }

    const diffCells = cells.filter((x) => x.changed);
    rows.push({ key: item.key, itemId: live.id, fromLog: true, cells, diffCells });
  }

  return {
    rows,
    warnings: { noVersion: [], foreignEdit: [], missing },
    summary: {
      items: rows.length,
      restorable: rows.filter((r) => r.diffCells.length).length,
      alreadySame: rows.filter((r) => !r.diffCells.length).length,
      cells: rows.reduce((a, r) => a + r.diffCells.length, 0),
      skipped: missing.length,
    },
  };
}

/**
 * display() 가 만든 표시 문자열을 원래 값으로 되돌린다.
 * display() 의 날짜는 toISOString() 기반이므로 UTC 로 해석해야 시간대가 밀리지 않는다.
 */
function parseDisplayText(text, col) {
  if (text == null || text === '' || text === '(비어 있음)') return null;
  if (col?.type === 'dateTime') {
    const s = String(text).trim().replace(' ', 'T');
    return /Z$/.test(s) ? s : s + (s.includes('T') ? ':00Z' : 'T00:00:00Z');
  }
  return text;
}
