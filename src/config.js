// 전역 설정.
// Client ID / Tenant ID 는 SPA 인증 요청 URL에 노출되는 공개 식별자이므로 여기에 직접 둔다.
// (client secret 은 PKCE 공개 클라이언트에서 사용하지 않는다 — 존재하지 않음)
// 커밋하면 안 되는 값이나 환경별 오버라이드는 .env → env.js 경로로 주입된다. (.env.example 참고)

const ENV = (typeof window !== 'undefined' && window.__KUFRI_ENV__) || {};
const pick = (envKey, fallback) => (ENV[envKey] != null && ENV[envKey] !== '' ? ENV[envKey] : fallback);

export const CONFIG = {
  // --- Entra ID 앱 등록 (공개 식별자) ---
  clientId: pick('KUFRI_CLIENT_ID', '9b247088-5afb-4622-9c5e-b5f27142761d'),
  tenantId: pick('KUFRI_TENANT_ID', '19cab1f5-21f4-44df-8ac6-96d6ca595203'),

  // --- SharePoint 대상 ---
  hostname: pick('KUFRI_SP_HOSTNAME', 'startruckkorea.sharepoint.com'),
  sitePath: pick('KUFRI_SP_SITE_PATH', '/sites/STK-Kufri'),
  listName: pick('KUFRI_LIST_NAME', 'vehicle_daily'), // /Lists/vehicle_daily/ 의 URL 세그먼트
  fileFolder: pick('KUFRI_FILE_FOLDER', 'STK-Kufri_Data'), // Shared Documents 하위 폴더
  fileBaseName: pick('KUFRI_FILE_BASENAME', 'afab_test'), // 확장자는 런타임 탐색

  // 두 소스의 행을 매칭하는 고유 키. 사용자가 고르지 않고 이름으로 자동 매칭한다.
  // 표기 차이(대소문자/공백/마침표/언더스코어, SharePoint 의 _x0020_ 이스케이프)는 흡수한다.
  keyColumn: pick('KUFRI_KEY_COLUMN', 'Commission no.'),

  // --- Microsoft Graph ---
  graphBase: 'https://graph.microsoft.com/v1.0',
  scopes: ['User.Read', 'Sites.ReadWrite.All', 'Files.ReadWrite.All'],

  // --- 벤치마크 기본값 ---
  defaults: {
    batchSize: 20, // Graph $batch 최대 20
    concurrency: 4, // 동시에 날리는 배치 수
    listPageSize: 999,
    maxRetry: 3,
  },
};

/** 앱 등록의 SPA 리디렉션 URI 와 문자 단위로 일치해야 한다. */
export function redirectUri() {
  return (
    ENV.KUFRI_REDIRECT_URI ||
    localStorage.getItem('kufri_redirect_uri') ||
    location.origin + location.pathname.replace(/index\.html$/, '')
  );
}

export const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: redirectUri(),
    postLogoutRedirectUri: redirectUri(),
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: 'sessionStorage', // 토큰은 탭 종료 시 소멸
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      piiLoggingEnabled: false,
      loggerCallback: (level, message, containsPii) => {
        if (!containsPii && level <= 1) console.warn('[msal]', message);
      },
    },
  },
};

// --- 매핑 설정 영속화 (사용자별 로컬) ---
const MAP_KEY = 'kufri_mapping_v1';
export const loadMapping = () => {
  try {
    return JSON.parse(localStorage.getItem(MAP_KEY)) || null;
  } catch {
    return null;
  }
};
export const saveMapping = (m) => localStorage.setItem(MAP_KEY, JSON.stringify(m));

// --- 적용 이력 (변경 리스트 화면용) ---
const LOG_KEY = 'kufri_changelog_v1';
export const loadChangeLog = () => {
  try {
    return JSON.parse(localStorage.getItem(LOG_KEY)) || [];
  } catch {
    return [];
  }
};
export const saveChangeLog = (log) => localStorage.setItem(LOG_KEY, JSON.stringify(log.slice(0, 50)));
