// 화면에 보이는 열 정의 — 전체 데이터 그리드의 열 순서와 표시 이름.
//
//   label : 화면 표시명 (사용자 지정. 리스트의 표시명을 쓰지 않는다)
//   list  : SharePoint 열 내부명. "확인" 은 동기화 대응표(CONFIG.mapping)로 검증된 것,
//           "추정" 은 이름 규칙으로 짐작한 것 — 실제 열은 findColumn 이 정확 일치 → 정규화 일치로 찾고,
//           list 로 못 찾으면 label 로 한 번 더 찾는다. 그래도 없으면 화면에 "찾지 못한 열" 로 나열된다.
//
// 내부명 규칙이 일정하지 않다 (actualpm, originalpm, order_date). 못 찾는 열이 나오면 내부명을 확인해 list 를 고친다.
import { nameKey } from './diff.js';

export const GRID_COLUMNS = [
  { label: 'Current Status', list: 'current_status' }, // 추정
  { label: 'Commission No.', list: 'commission_no', key: true }, // 확인
  { label: 'Model Classification', list: 'model_classification' }, // 추정
  { label: 'Baumuster', list: 'baumuster' }, // 확인
  { label: 'Exception Case', list: 'exception_case' }, // 추정
  { label: 'Model Name Check', list: 'model_name_check' }, // 추정
  { label: 'Model in AFAB', list: 'model_in_afab' }, // 확인
  { label: 'Vessel', list: 'vessel' }, // 확인
  { label: 'Import License', list: 'import_license' }, // 추정
  { label: 'B/L No.', list: 'bl_no' }, // 추정
  { label: 'Engine No.', list: 'engine_no' }, // 확인
  { label: 'Option Remarks or Key Information', list: 'option_remarks' }, // 추정
  { label: 'Tires', list: 'tires' }, // 추정
  { label: 'Days in Stock', list: 'days_in_stock' }, // 추정
  { label: 'SalesMonth', list: 'salesmonth' }, // 추정
  { label: 'Vehicle', list: 'vehicle' }, // 추정
  { label: 'Series', list: 'series' }, // 추정
  { label: 'Model', list: 'model' }, // 추정
  { label: 'Axle', list: 'axle' }, // 추정
  { label: 'Cab', list: 'cab' }, // 추정
  { label: 'MY', list: 'my' }, // 추정
  { label: 'Duty Class', list: 'duty_class' }, // 추정
  { label: 'Sub Cat', list: 'sub_cat' }, // 확인
  { label: 'Color Code', list: 'color' }, // 추정 — 대응표의 color(Lack 1)를 색상 코드로 봄
  { label: 'Color Name', list: 'color_name' }, // 추정
  { label: 'Gen.', list: 'gen' }, // 추정
  { label: 'VIN No.', list: 'vin_no' }, // 확인
  { label: 'Option', list: 'option' }, // 추정
  { label: 'Wheelbase', list: 'wheelbase' }, // 추정
  { label: 'PTO', list: 'pto' }, // 추정
  { label: 'Fleetboard', list: 'fleetboard' }, // 추정
  { label: 'Emission', list: 'emission' }, // 추정
  { label: 'Order date in AFAB', list: 'order_date' }, // 확인
  { label: 'Change', list: 'change' }, // 확인 (표시명은 받은 그대로)
  { label: 'Changeability Date', list: 'changeability_date' }, // 확인
  { label: 'B1 Factory Order', list: 'b1_factory_order' }, // 추정
  { label: 'Original PM', list: 'originalpm' }, // 추정
  { label: 'Actual PM', list: 'actualpm' }, // 확인
  { label: 'TDD Calc.', list: 'tdd_cal' }, // 확인
  { label: 'TDD actual', list: 'tdd_actual' }, // 확인
  { label: 'Dispatch', list: 'dispatch' }, // 확인
  { label: 'Shipping', list: 'shipping' }, // 확인
  { label: 'Invoice', list: 'invoice' }, // 확인
  { label: 'Planned Arrival', list: 'planned_arrival' }, // 확인
  { label: 'ETA (TDD calc.+75)', list: 'eta' }, // 추정
  { label: 'ATA(unipass)', list: 'ata' }, // 추정
  { label: 'Customs Clearance Date', list: 'customs_clearance_date' }, // 추정
];

// diff.js 와 서로 import 한다. 모듈 평가 시점에 nameKey 를 부르면 TDZ 오류가 나므로 첫 호출 때 만든다.
let byKey = null;

/** 리스트 열 내부명 → 화면 표시명. 정의에 없으면 null (호출자가 표시명/내부명으로 대체) */
export function labelFor(listName) {
  if (!byKey) byKey = new Map(GRID_COLUMNS.map((c) => [nameKey(c.list), c.label]));
  return byKey.get(nameKey(listName)) ?? null;
}
