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
  // 동기화 대상 리스트. 행으로 나뉜 리스트 여러 개를 하나의 대상으로 다룬다 (열 구성은 같아야 한다).
  // 한 행(키)은 이 중 한 리스트에만 있어야 하며, 쓰기는 그 행이 원래 있던 리스트로 간다.
  listNames: pick('VITE_LIST_NAMES', 'logi_master_1,logi_master_2')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean),
  fileFolder: pick('VITE_FILE_FOLDER', 'STK-Kufri_Data'),
  // 파일명 접두어. 정확히 같은 이름이 없으면 이 접두어로 시작하는 파일 중 가장 최근 수정본을 쓴다.
  fileBaseName: pick('VITE_FILE_BASENAME', 'SDISP'),

  // 두 소스의 열 대응표. 리스트 열 내부명(영문)과 SDISP 파일 헤더(독일어)가 서로 달라
  // 이름 기반 자동 매칭이 불가능하므로 고정한다. 사용자가 화면에서 바꾸지 않는다.
  //   list  : SharePoint 열 내부명 또는 표시명 (정확 일치 → 정규화 일치 순으로 찾는다)
  //   excel : SDISP 파일 헤더
  // key 는 행 매칭 기준이며 갱신하지 않는다. 대응표의 열을 하나라도 못 찾으면 비교를 시작하지 않는다.
  mapping: {
    key: { list: 'commission_no', excel: 'Auftragsnummer' },
    pairs: [
      { list: 'baumuster', excel: 'Baumuster' },
      { list: 'model_in_afab', excel: 'Fahrzeugtyp' },
      { list: 'vessel', excel: 'Schiffsname' },
      { list: 'engine_no', excel: 'Motor-Nr.' },
      { list: 'sub_cat', excel: 'Subkat.' },
      { list: 'color', excel: 'Lack 1' },
      { list: 'vin_no', excel: 'Fahrzeug-Ident-Nr. (FIN)' },
      { list: 'order_date', excel: 'Bestelldatum' },
      { list: 'change', excel: 'Ä' },
      { list: 'changeability_date', excel: 'Änderbarkeitsdatum' },
      { list: 'actualpm', excel: 'LT-Quote' },
      { list: 'disfatch', excel: 'Versand' },
      { list: 'tdd_cal', excel: 'GLT-Err.' },
      { list: 'tdd_actual', excel: 'GLT-Ist' },
      { list: 'shipping', excel: 'Abgangsdatum' },
      { list: 'invoice', excel: 'Rg.-Datum' },
      { list: 'planned_arrival', excel: 'Gepl. Ankunftsdatum' },
    ],
  },

  graphBase: 'https://graph.microsoft.com/v1.0',
  // [체크리스트 1-03] 최소 권한. Sites.Selected 로 좁힐 수 있는지 검토 중.
  //
  // 주의: 여기 적는 scope 는 앱 등록에 "관리자 동의된" 것과 문자 단위로 같아야 한다.
  // Files.ReadWrite.All 에 동의돼 있어도 Files.Read.All 은 별개의 scope 라 동의가 없다.
  // 사용자 동의가 정책으로 막힌 테넌트에서는 AADSTS65001/90094 로 로그인 자체가 실패한다.
  // (실사이트 전환 때 Read.All 로 줄였다가 이 문제로 되돌림. 줄이려면 앱 등록에 먼저 동의를 추가할 것)
  scopes: ['User.Read', 'Sites.ReadWrite.All', 'Files.ReadWrite.All'],

  // concurrency 는 SharePoint 부하 레버. PoC 검증값 4 로 운용한다 (300건 규모에서 RU 한도 내).
  defaults: { batchSize: 20, concurrency: 4, listPageSize: 999, maxRetry: 3 },
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
