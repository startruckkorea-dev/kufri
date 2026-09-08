#!/usr/bin/env node
/**
 * PreToolUse 가드 — 커밋은커녕 파일에 써지기 전에 막는다.
 *
 * STK 개발 보안 체크리스트 대응:
 *   5-01 소스코드 내 민감정보 하드코딩 금지
 *   2-05 암호화 키 · 비밀값 안전 관리
 *   7-03 웹 보안 헤더 / XSS 방어  (DOM sink 차단)
 *
 * 예외가 정말 필요하면 해당 줄에 아래 주석을 달고 사유를 남긴다.
 *   // security-ok: DOMPurify 로 정제한 값
 */
let raw = '';
for await (const chunk of process.stdin) raw += chunk;

const allow = (why) => {
  if (why) process.stderr.write(why + '\n');
  process.exit(0);
};

let input;
try {
  input = JSON.parse(raw || '{}');
} catch {
  allow(); // 입력을 못 읽으면 작업을 막지 않는다 (fail-open)
}

const ti = input.tool_input || {};
const filePath = String(ti.file_path || ti.notebook_path || '');
// Write=content, Edit=new_string, NotebookEdit=new_source
const content = String(ti.content ?? ti.new_string ?? ti.new_source ?? '');
if (!content.trim()) allow();

const p = filePath.split(String.fromCharCode(92)).join('/');
// 가드 자신과 체크리스트 문서는 패턴 예시를 담고 있으므로 제외
if (/\.claude\/hooks\//.test(p)) allow();

const ext = (p.match(/\.[a-z0-9]+$/i) || [''])[0].toLowerCase();
const CODE = ['.js', '.mjs', '.cjs', '.ts', '.jsx', '.tsx', '.html', '.htm', '.svelte', '.vue'];
const isCode = CODE.includes(ext);

const SECRET_RULES = [
  { id: '5-01', name: '개인키 블록', re: /-----BEGIN[ A-Z]*PRIVATE KEY-----/ },
  { id: '5-01', name: 'AWS 액세스 키', re: /\bAKIA[0-9A-Z]{16}\b/ },
  { id: '5-01', name: 'JWT 토큰 리터럴', re: /\beyJ[A-Za-z0-9_-]{10,}\.eyJ[A-Za-z0-9_-]{10,}\./ },
  { id: '5-01', name: 'GitHub 토큰', re: /\bgh[pousr]_[A-Za-z0-9]{16,}\b/ },
  { id: '5-01', name: 'Slack 토큰', re: /\bxox[abprs]-[A-Za-z0-9-]{10,}\b/ },
  {
    id: '2-05',
    name: '비밀값 하드코딩',
    re: /\b(client[_-]?secret|password|passwd|secret[_-]?key|api[_-]?key|apikey|access[_-]?token|refresh[_-]?token|private[_-]?key|conn(ection)?[_-]?string)\b\s*[:=]\s*['"`]([^'"`\n]{8,})['"`]/i,
    // 자리표시자 · 환경변수 참조는 통과
    skip: (m) => /^(process\.|import\.meta|\$\{|<|xxx|your|example|dummy|change|placeholder|redacted|\*{3}|test|sample|todo)/i.test(m[3]),
  },
];

const XSS_RULES = [
  { id: '7-03', name: 'innerHTML/outerHTML 대입', re: /\.(inner|outer)HTML\s*(\+?=)[^=]/ },
  { id: '7-03', name: 'insertAdjacentHTML', re: /\.insertAdjacentHTML\s*\(/ },
  { id: '7-03', name: 'document.write', re: /\bdocument\s*\.\s*write(ln)?\s*\(/ },
  { id: '7-03', name: 'eval', re: /(^|[^.\w])eval\s*\(/ },
  { id: '7-03', name: 'new Function', re: /\bnew\s+Function\s*\(/ },
  { id: '7-03', name: 'dangerouslySetInnerHTML', re: /dangerouslySetInnerHTML/ },
  { id: '7-03', name: '문자열 setTimeout/setInterval', re: /\bset(Timeout|Interval)\s*\(\s*['"`]/ },
  { id: '7-03', name: 'javascript: URL', re: /['"`]\s*javascript:/i },
];

const rules = [...SECRET_RULES, ...(isCode ? XSS_RULES : [])];
const findings = [];

content.split(/\r?\n/).forEach((line, i) => {
  if (/security-ok:/.test(line)) return; // 사유를 남긴 예외
  for (const r of rules) {
    const m = line.match(r.re);
    if (!m) continue;
    if (r.skip && r.skip(m)) continue;
    findings.push({ line: i + 1, id: r.id, name: r.name, text: line.trim().slice(0, 120) });
  }
});

if (!findings.length) allow();

const detail = findings
  .map((f) => `  · ${f.line}행 [체크리스트 ${f.id}] ${f.name}\n      ${f.text}`)
  .join('\n');

const reason =
  `STK 개발 보안 체크리스트 위반으로 쓰기를 차단했습니다 — ${p || '(경로 미상)'}\n` +
  detail +
  `\n\n조치:\n` +
  `  · 비밀값(5-01/2-05): 소스에서 제거하고 환경변수나 Secret Manager 로 옮기세요.\n` +
  `    이미 커밋된 적이 있다면 값 자체를 교체해야 합니다.\n` +
  `  · XSS(7-03): textContent / 자동 이스케이프 템플릿을 쓰세요.\n` +
  `    HTML 삽입이 불가피하면 정제(sanitize) 후 해당 줄에 "security-ok: 사유" 주석을 남기세요.`;

process.stdout.write(
  JSON.stringify({
    hookSpecificOutput: {
      hookEventName: 'PreToolUse',
      permissionDecision: 'deny',
      permissionDecisionReason: reason,
    },
  })
);
