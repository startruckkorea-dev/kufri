// 쓰기 계층 — Graph $batch(20건 묶음) + 배치 동시 전송, 그리고 그 계측
import { CONFIG } from './config.js';
import { gbatch, runPool, chunk, statsOf } from './graph.js';
import { ctx } from './sources.js';

/**
 * 변경 항목을 SharePoint 리스트에 반영한다.
 * @param {Array} changedRows buildDiff().changed 중 사용자가 선택한 것들
 * @param {object} opts {batchSize, concurrency, dryRun, onProgress}
 */
export async function applyChanges(changedRows, opts = {}) {
  const batchSize = Math.min(20, Math.max(1, opts.batchSize ?? CONFIG.defaults.batchSize));
  const concurrency = Math.max(1, opts.concurrency ?? CONFIG.defaults.concurrency);
  const dryRun = !!opts.dryRun;

  // 1건 = 아이템 1개의 fields PATCH (변경된 필드만 담는다).
  // listId 는 비교 때 기록된 소속 리스트. 한 $batch 안에 서로 다른 리스트의 요청이 섞여도 된다.
  const ops = changedRows.map((row) => {
    if (!row.listId) throw new Error(`'${row.key}' 의 소속 리스트를 알 수 없습니다. 비교를 다시 실행하세요.`);
    return {
      key: row.key,
      itemId: row.itemId,
      listId: row.listId,
      fields: Object.fromEntries(row.diffCells.map((c) => [c.field, c.writeValue])),
      diffCells: row.diffCells,
    };
  });

  const batches = chunk(ops, batchSize);
  const batchTimings = [];
  const results = [];
  let throttled = 0;

  const startedAt = performance.now();

  if (dryRun) {
    return finish({
      dryRun: true,
      ops,
      batches,
      batchSize,
      concurrency,
      startedAt,
      endedAt: performance.now(),
      batchTimings: [],
      results: ops.map((o) => ({ key: o.key, itemId: o.itemId, status: 0, ok: true, dryRun: true })),
      throttled: 0,
    });
  }

  const tasks = batches.map((batchOps, bi) => async () => {
    const requests = batchOps.map((op, i) => ({
      id: `${bi}-${i}`,
      method: 'PATCH',
      url: `/sites/${ctx.siteId}/lists/${op.listId}/items/${op.itemId}/fields`,
      headers: { 'if-match': '*' },
      body: op.fields,
    }));

    const t0 = performance.now();
    const responses = await gbatch(requests, `write.batch#${bi}`);
    const ms = performance.now() - t0;
    batchTimings.push({ index: bi, size: batchOps.length, ms });

    responses.forEach((r) => {
      const [, idx] = r.id.split('-').map(Number);
      const op = batchOps[idx];
      const ok = r.status >= 200 && r.status < 300;
      if (r.status === 429) throttled++;
      results.push({
        key: op.key,
        itemId: op.itemId,
        status: r.status,
        ok,
        fields: op.fields,
        diffCells: op.diffCells,
        error: ok ? null : r.body?.error?.message || JSON.stringify(r.body || {}).slice(0, 300),
      });
    });
    return ms;
  });

  await runPool(tasks, concurrency, (done, total) => opts.onProgress?.(done, total));
  const endedAt = performance.now();

  return finish({ dryRun: false, ops, batches, batchSize, concurrency, startedAt, endedAt, batchTimings, results, throttled });
}

function finish({ dryRun, ops, batches, batchSize, concurrency, startedAt, endedAt, batchTimings, results, throttled }) {
  const wallMs = endedAt - startedAt;
  const okCount = results.filter((r) => r.ok).length;
  const failCount = results.length - okCount;
  const cellCount = ops.reduce((a, o) => a + o.diffCells.length, 0);

  return {
    dryRun,
    itemCount: ops.length,
    cellCount,
    batchCount: batches.length,
    batchSize,
    concurrency,
    wallMs,
    okCount,
    failCount,
    throttled,
    itemsPerSec: wallMs > 0 ? (ops.length / wallMs) * 1000 : 0,
    msPerItem: ops.length ? wallMs / ops.length : 0,
    batchStats: statsOf(batchTimings.map((b) => b.ms)),
    batchTimings,
    results,
  };
}
