// 대상 설정.
//
// [체크리스트 2-05] clientId / tenantId 는 SPA 인증 요청 URL 에 그대로 실려 나가는
// 공개 식별자이므로 소스에 둔다. PKCE 공개 클라이언트라 client secret 자체가 없다.
// 그 외 값이 필요해지면 .env 의 VITE_* 로 주입한다(.env 는 .gitignore 처리됨).
const env = import.meta.env;
const pick = (key, fallback) => (env[key] != null && env[key] !== '' ? env[key] : fallback);

export const CONFIG = {
  clientId: pick('VITE_CLIENT_ID', '9b247088-5afb-4622-9c5e-b5f27142761d'),
  tenantId: pick('VITE_TENANT_ID', '19cab1f5-21f4-44df-8ac6-96d6ca595203'),

  hostname: pick('VITE_SP_HOSTNAME', 'startruckkorea.sharepoint.com'),
  sitePath: pick('VITE_SP_SITE_PATH', '/sites/STK-Kufri'),
  listName: pick('VITE_LIST_NAME', 'vehicle_daily'),
  fileFolder: pick('VITE_FILE_FOLDER', 'STK-Kufri_Data'),
  fileBaseName: pick('VITE_FILE_BASENAME', 'afab_test'),

  // 두 소스의 행을 매칭하는 고유 키. 사용자가 고르지 않고 이름으로 자동 매칭한다.
  keyColumn: pick('VITE_KEY_COLUMN', 'Commission no.'),

  graphBase: 'https://graph.microsoft.com/v1.0',
  // [체크리스트 1-03] 최소 권한. Sites.Selected 로 좁힐 수 있는지 검토 중.
  scopes: ['User.Read', 'Sites.ReadWrite.All', 'Files.Read.All'],

  defaults: { batchSize: 20, concurrency: 2, listPageSize: 999, maxRetry: 3 },
};

/** 앱 등록의 SPA 리디렉션 URI 와 문자 단위로 일치해야 한다. */
export function redirectUri() {
  return pick('VITE_REDIRECT_URI', location.origin + '/');
}

export const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: redirectUri(),
    postLogoutRedirectUri: redirectUri(),
    navigateToLoginRequestUrl: false,
  },
  // [체크리스트 1-07] 탭 종료 시 토큰 소멸
  cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
  system: {
    loggerOptions: {
      piiLoggingEnabled: false,
      loggerCallback: (level, message, containsPii) => {
        if (!containsPii && level <= 1) console.warn('[msal]', message);
      },
    },
  },
};

// 매핑 설정만 로컬에 보관한다.
// [체크리스트 4-02] 감사 증적은 localStorage 가 아니라 SharePoint 버전 기록 / 변경 로그 리스트에 남긴다.
const MAP_KEY = 'kufri_mapping_v1';
export const loadMapping = () => {
  try {
    return JSON.parse(localStorage.getItem(MAP_KEY)) || null;
  } catch {
    return null;
  }
};
export const saveMapping = (m) => {
  try {
    localStorage.setItem(MAP_KEY, JSON.stringify(m));
  } catch {
    /* 저장 실패는 기능에 영향 없음 */
  }
};
