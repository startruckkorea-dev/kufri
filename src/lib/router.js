// 해시 라우터. GitHub Pages 는 SPA 폴백이 없어 경로 라우팅 대신 해시를 쓴다 (docs/ui-design.md 5장).
//
// MSAL 도 인증 응답을 URL fragment(#code=…)로 받는다. App 은 initAuth()(handleRedirectPromise) 가
// 끝난 뒤에만 라우트를 그리므로, 여기서 보는 해시는 MSAL 이 소비하고 남은 것이다.
//
// 세그먼트 값은 NAV 레지스트리 조회에만 쓴다. URL·HTML·$filter 어디에도 넣지 않는다.
import { useMemo, useSyncExternalStore } from 'react';
import { ADMIN, ADMIN_SUBS, NAV } from './nav.js';

const NOT_FOUND = Object.freeze({ view: 'notfound', item: null, sub: '' });

/** '#/admin/sync' → { view, item, sub } */
export function parseHash(hash) {
  const segs = String(hash || '')
    .replace(/^#/, '')
    .split('/')
    .filter(Boolean);

  if (segs.length === 0) return { view: NAV[0].key, item: NAV[0], sub: '' }; // '#/' = 첫 메뉴

  const item = NAV.find((n) => n.key === segs[0]);
  if (!item) return NOT_FOUND;

  if (item.key === ADMIN.key) {
    if (segs.length > 2) return NOT_FOUND;
    const sub = segs[1] || '';
    if (!ADMIN_SUBS.some((s) => s.key === sub)) return NOT_FOUND;
    return { view: 'admin', item, sub };
  }

  if (segs.length > 1) return NOT_FOUND;
  return { view: item.key, item, sub: '' };
}

const subscribe = (cb) => {
  window.addEventListener('hashchange', cb);
  return () => window.removeEventListener('hashchange', cb);
};
// getSnapshot 은 같은 상태에서 같은 참조를 돌려줘야 한다 (객체를 만들면 무한 렌더). 문자열을 쓴다.
const getSnapshot = () => location.hash;

export function useRoute() {
  const hash = useSyncExternalStore(subscribe, getSnapshot);
  return useMemo(() => parseHash(hash), [hash]);
}

export const isKnownHash = (hash) => parseHash(hash).view !== 'notfound';
