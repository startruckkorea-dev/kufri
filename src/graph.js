// Microsoft Graph 호출 계층 — 계측(timing) + 429/5xx 재시도 + $batch
import { CONFIG } from './config.js';
import { getToken } from './auth.js';

/** 모든 호출의 계측 레코드. 벤치마크 화면이 이걸 읽는다. */
export const telemetry = {
  calls: [],
  reset() {
    this.calls.length = 0;
  },
  add(rec) {
    this.calls.push(rec);
    return rec;
  },
  /** name prefix 로 필터한 요약 통계 */
  summary(prefix) {
    const rows = this.calls.filter((c) => !prefix || c.name.startsWith(prefix));
    return statsOf(rows.map((r) => r.ms));
  },
};

export function statsOf(values) {
  if (!values.length) return { n: 0, total: 0, avg: 0, p50: 0, p95: 0, min: 0, max: 0 };
  const s = [...values].sort((a, b) => a - b);
  const at = (p) => s[Math.min(s.length - 1, Math.floor((p / 100) * s.length))];
  const total = s.reduce((a, b) => a + b, 0);
  return { n: s.length, total, avg: total / s.length, p50: at(50), p95: at(95), min: s[0], max: s[s.length - 1] };
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * Graph 단일 호출.
 * @param {string} path  '/sites/...' 또는 절대 URL(nextLink)
 */
export async function gfetch(path, opts = {}) {
  const { method = 'GET', body, headers = {}, name = method + ' ' + path.split('?')[0], raw = false } = opts;
  const url = path.startsWith('http') ? path : CONFIG.graphBase + path;

  let retries = 0;
  let throttleWaitMs = 0;

  while (true) {
    const { token } = await getToken();
    const t0 = performance.now();
    const res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(body ? { 'Content-Type': 'application/json' } : {}),
        ...headers,
      },
      body: body ? JSON.stringify(body) : undefined,
    });
    const ms = performance.now() - t0;

    // 429 / 503 / 504 → Retry-After 존중
    if ((res.status === 429 || res.status === 503 || res.status === 504) && retries < CONFIG.defaults.maxRetry) {
      const wait = (Number(res.headers.get('Retry-After')) || Math.pow(2, retries)) * 1000;
      throttleWaitMs += wait;
      retries++;
      telemetry.add({ name, method, url, ms, status: res.status, retries, throttled: true, throttleWaitMs: wait });
      await sleep(wait);
      continue;
    }

    if (!res.ok) {
      const text = await res.text();
      telemetry.add({ name, method, url, ms, status: res.status, retries, error: text.slice(0, 400) });
      throw new GraphError(`${name} 실패 (HTTP ${res.status})`, res.status, text);
    }

    const rec = telemetry.add({ name, method, url, ms, status: res.status, retries, throttleWaitMs });
    if (raw) return { res, rec };
    const data = res.status === 204 ? null : await res.json();
    return data;
  }
}

export class GraphError extends Error {
  constructor(message, status, detail) {
    super(message);
    this.status = status;
    this.detail = detail;
  }
}

/** @odata.nextLink 를 끝까지 따라가며 value 를 모은다. */
export async function gfetchAll(path, opts = {}) {
  const out = [];
  let next = path;
  let pages = 0;
  while (next) {
    const data = await gfetch(next, { ...opts, name: (opts.name || 'page') + `#${pages}` });
    out.push(...(data.value || []));
    next = data['@odata.nextLink'] || null;
    pages++;
  }
  return { items: out, pages };
}

/**
 * Graph $batch — 최대 20개 요청을 1 HTTP 왕복으로 묶는다.
 * @param {Array<{id,method,url,body,headers}>} requests
 * @returns {Promise<Array<{id,status,body}>>}
 */
export async function gbatch(requests, name = 'batch') {
  if (requests.length > 20) throw new Error('Graph $batch 는 요청 20개가 상한입니다.');
  const payload = {
    requests: requests.map((r) => ({
      id: String(r.id),
      method: r.method,
      url: r.url, // /v1.0 접두어 없이 상대경로
      ...(r.body ? { body: r.body, headers: { 'Content-Type': 'application/json', ...(r.headers || {}) } } : {}),
      ...(r.headers && !r.body ? { headers: r.headers } : {}),
    })),
  };
  const data = await gfetch('/$batch', { method: 'POST', body: payload, name });
  return (data.responses || []).map((r) => ({ id: r.id, status: r.status, body: r.body, headers: r.headers }));
}

/** 동시성 제한 실행기. tasks 는 () => Promise 형태. */
export async function runPool(tasks, concurrency, onProgress) {
  const results = new Array(tasks.length);
  let cursor = 0;
  let done = 0;

  const worker = async () => {
    while (true) {
      const i = cursor++;
      if (i >= tasks.length) return;
      try {
        results[i] = { ok: true, value: await tasks[i]() };
      } catch (e) {
        results[i] = { ok: false, error: e };
      }
      done++;
      onProgress?.(done, tasks.length);
    }
  };

  await Promise.all(Array.from({ length: Math.min(concurrency, tasks.length) }, worker));
  return results;
}

export const chunk = (arr, size) => {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
};
