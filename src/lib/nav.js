// 좌측 내비게이션 레지스트리. 사이드바·라우트 유효성이 여기서 파생된다.
// key 는 해시 경로 첫 세그먼트(#/logistics). 화면을 추가·이름 변경할 때는 이 파일만 고친다.
export const NAV = [
  { key: 'logistics', title: 'Logistics Status' },
  { key: 'production', title: 'Production Actual' },
  { key: 'shipping', title: 'Shipping Schedule' },
  { key: 'arrival', title: 'Arrival Schedule' },
  { key: 'stock', title: 'Stock Status' },
  { key: 'kufri', title: 'KUFRI' },
  { key: 'report', title: 'REPORT' },
  // 관리자 전용: 전체 데이터 그리드 + Excel 동기화. 접근 제어는 SharePoint 권한이 강제한다(3-03).
  { key: 'admin', title: 'Data Admin', admin: true, desc: '전체 데이터 · Excel 동기화' },
];

export const ADMIN = NAV[NAV.length - 1];

/** 관리자 탭 하위 화면 */
export const ADMIN_SUBS = [
  { key: '', title: '전체 데이터' },
  { key: 'sync', title: 'Excel 동기화' },
];

export const hrefFor = (item, sub = '') => `#/${item.key}${sub ? '/' + sub : ''}`;
