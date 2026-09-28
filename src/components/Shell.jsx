import { NAV, hrefFor } from '../lib/nav.js';

/**
 * 상단바 + 좌측 사이드바 + 콘텐츠. 사이드바 항목은 NAV 레지스트리에서 생성한다.
 * 로그인 전에는 사이드바를 그리지 않는다.
 */
export default function Shell({ account, booting, route, onLogin, onLogout, children }) {
  return (
    <div className="shell">
      <header className="app">
        <a className="brand" href="#/">
          KUFRI
        </a>
        <div className="spacer" />
        <span className="who">{account ? `${account.name || ''} <${account.username}>` : '로그인 필요'}</span>
        {account ? (
          <button className="btn" onClick={onLogout}>
            로그아웃
          </button>
        ) : (
          <button className="btn" onClick={onLogin} disabled={booting}>
            Microsoft 계정으로 로그인
          </button>
        )}
      </header>

      {account ? (
        <nav className="side" aria-label="주 메뉴">
          {NAV.map((item, i) => (
            <a
              key={item.key}
              href={hrefFor(item)}
              className={item.admin ? 'admin' : ''}
              aria-current={route.item?.key === item.key ? 'page' : undefined}
              title={item.desc || item.title}
            >
              <span className="n">{i + 1}</span>
              <span className="t">{item.title}</span>
            </a>
          ))}
        </nav>
      ) : null}

      <main>{children}</main>
    </div>
  );
}
