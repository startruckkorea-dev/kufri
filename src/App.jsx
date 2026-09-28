import { useCallback, useEffect, useRef, useState } from 'react';
import { redirectUri } from './lib/config.js';
import { initAuth, login, logout } from './lib/auth.js';
import { isKnownHash, useRoute } from './lib/router.js';
import { resolveLists } from './lib/sources.js';
import { useSyncFlow } from './hooks/useSyncFlow.js';
import { useListGrid } from './hooks/useListGrid.js';
import Shell from './components/Shell.jsx';
import Placeholder, { LoginCard, NotFound } from './components/Placeholder.jsx';
import Admin from './components/admin/Admin.jsx';

// 로그인 리디렉션 전에 보던 화면. 감사 증적이 아닌 UX 편의이므로 sessionStorage 를 쓴다.
const RETURN_KEY = 'kufri_return_hash';

export default function App() {
  const [account, setAccount] = useState(null);
  const [booting, setBooting] = useState(true);
  const [busy, setBusy] = useState('');
  const [error, setError] = useState(null);
  const route = useRoute();

  useEffect(() => {
    initAuth()
      .then((acc) => {
        setAccount(acc);
        // MSAL 이 fragment(#code=…)를 소비한 뒤다. 등록된 라우트일 때만 이전 화면으로 돌아간다.
        let back = null;
        try {
          back = sessionStorage.getItem(RETURN_KEY);
          sessionStorage.removeItem(RETURN_KEY);
        } catch {
          /* 저장소 접근 불가 시 첫 화면 */
        }
        if (acc && back && isKnownHash(back)) location.hash = back;
      })
      .catch((e) => {
        // 원인별로 다른 힌트를 준다. 항상 "리디렉션 URI 확인" 만 띄우면 엉뚱한 곳을 보게 된다.
        const code = e.errorCode || e.name || '';
        const text = `${code} ${e.message}`;
        const hint = /65001|90094|consent/i.test(text)
          ? '동의되지 않은 scope 를 요청했습니다. 앱 등록의 API 권한(관리자 동의)과 config.js 의 scopes 가 문자 단위로 같아야 합니다.'
          : /50011|redirect_uri/i.test(text)
            ? `앱 등록의 SPA 리디렉션 URI 가 '${redirectUri()}' 와 문자 단위로 일치하는지 확인하세요.`
            : '브라우저 콘솔에 Content-Security-Policy 위반 메시지가 있는지도 확인하세요.';
        setError(new Error(`MSAL 초기화 실패 [${code}]: ${e.message} — ${hint}`));
      })
      .finally(() => setBooting(false));
  }, []);

  const run = useCallback(async (label, fn) => {
    setBusy(label);
    setError(null);
    try {
      await fn();
    } catch (e) {
      console.error(e);
      setError(e);
    } finally {
      setBusy('');
    }
  }, []);

  const sync = useSyncFlow(run);
  const grid = useListGrid(run);

  // 로그인 직후 사이트·리스트·열 리졸브를 미리 해 둔다 (작은 호출 3회, 세션 캐시).
  // 관리자 탭에 들어갈 때 항목 읽기만 남아 첫 화면이 빨라진다. 실패는 조용히 넘기고, 실제 읽기 때 다시 시도되어 오류가 표시된다.
  useEffect(() => {
    if (account) resolveLists().catch(() => {});
  }, [account]);

  const handleLogin = () => {
    try {
      sessionStorage.setItem(RETURN_KEY, location.hash);
    } catch {
      /* 무시 */
    }
    login();
  };

  // 관리자 탭의 전체 데이터는 첫 진입 때 자동으로 읽는다 (세션당 1회, 이후는 새로고침 버튼).
  const requested = useRef(false);
  useEffect(() => {
    if (!account || route.view !== 'admin' || route.sub === 'sync' || grid.data || requested.current) return;
    requested.current = true;
    grid.load();
  }, [account, route, grid]);

  let content;
  if (!account) content = <LoginCard booting={booting} onLogin={handleLogin} />;
  else if (route.view === 'admin') content = <Admin sub={route.sub} sync={sync} grid={grid} busy={Boolean(busy)} />;
  else if (route.view === 'notfound') content = <NotFound />;
  else content = <Placeholder item={route.item} />;

  return (
    <Shell account={account} booting={booting} route={route} onLogin={handleLogin} onLogout={() => logout()}>
      {error && (
        <div className="alert err">
          <b>오류</b> — {error.message || String(error)}
          {error.detail && <pre>{String(error.detail).slice(0, 800)}</pre>}
        </div>
      )}
      {busy && (
        <div className="alert info">
          처리 중… <span className="muted">{busy}</span>
        </div>
      )}
      {booting ? null : content}
    </Shell>
  );
}
