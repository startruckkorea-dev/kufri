// .env  ->  env.js  생성기
// 브라우저는 .env 를 직접 읽을 수 없으므로, 정적 SPA 에서 쓸 수 있는 형태로 변환한다.
// 생성물 env.js 는 .gitignore 처리되어 있다.
//
//   node scripts/gen-env.mjs
//
// .env 가 없으면 아무 것도 하지 않는다. (src/config.js 의 기본값으로 동작)
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const envPath = resolve(root, '.env');
const outPath = resolve(root, 'env.js');

if (!existsSync(envPath)) {
  console.log('[gen-env] .env 없음 — src/config.js 기본값을 사용합니다. (env.js 생성 안 함)');
  process.exit(0);
}

const env = {};
for (const raw of readFileSync(envPath, 'utf8').split(/\r?\n/)) {
  const line = raw.trim();
  if (!line || line.startsWith('#')) continue;
  const eq = line.indexOf('=');
  if (eq < 0) continue;
  const key = line.slice(0, eq).trim();
  const value = line.slice(eq + 1).trim().replace(/^["']|["']$/g, '');
  if (key.startsWith('KUFRI_') && value) env[key] = value;
}

writeFileSync(
  outPath,
  `// 자동 생성 파일 — 직접 수정하지 마세요. (scripts/gen-env.mjs)\n` +
    `window.__KUFRI_ENV__ = ${JSON.stringify(env, null, 2)};\n`,
  'utf8'
);
console.log(`[gen-env] env.js 생성 완료 — 주입된 키: ${Object.keys(env).join(', ') || '(없음)'}`);
