import * as msal from '@azure/msal-browser';
// MSAL Browser v3 — Authorization Code Flow + PKCE (SPA 공개 클라이언트)
// PKCE 는 MSAL v3 의 SPA 경로에서 기본이자 강제이며, client secret 을 사용하지 않는다.
import { CONFIG, msalConfig } from './config.js';

let pca = null;
let account = null;

export async function initAuth() {
  pca = new msal.PublicClientApplication(msalConfig);
  await pca.initialize();

  // 로그인 리디렉션 복귀 처리
  const result = await pca.handleRedirectPromise();
  if (result?.account) {
    account = result.account;
  } else {
    account = pca.getAllAccounts()[0] || null;
  }
  if (account) pca.setActiveAccount(account);
  return account;
}

export const getAccount = () => account;

export async function login() {
  await pca.loginRedirect({ scopes: CONFIG.scopes, prompt: 'select_account' });
}

export async function logout() {
  await pca.logoutRedirect({ account });
}

/**
 * 액세스 토큰 획득. 캐시 우선(acquireTokenSilent), 만료/동의필요 시 리디렉션으로 승격.
 * 토큰 획득 소요시간을 함께 반환해 벤치마크에서 순수 API 시간과 분리할 수 있게 한다.
 */
export async function getToken(scopes = CONFIG.scopes) {
  if (!account) throw new Error('로그인이 필요합니다.');
  const t0 = performance.now();
  try {
    const res = await pca.acquireTokenSilent({ scopes, account });
    return { token: res.accessToken, ms: performance.now() - t0, fromCache: res.fromCache !== false };
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      await pca.acquireTokenRedirect({ scopes, account });
      throw new Error('추가 동의가 필요하여 리디렉션합니다.');
    }
    throw e;
  }
}
