/**
 * 기존 스프레드시트 기반 시스템 설정
 */

const SOURCE_SPREADSHEET_ID = '1C2Rr4oK3y6VKXTv7_R7ciJ6ihcbum_DWBIolUwgJXoQ';

/**
 * 기존 스프레드시트 분석 결과 표시
 */
function showAnalysisDialog() {
  const analysis = analyzeExistingSheet();
  const ui = SpreadsheetApp.getUi();
  ui.alert('스프레드시트 분석 결과', analysis, ui.ButtonSet.OK);
}

/**
 * 메뉴에 분석 도구 추가
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('📋 ASG 직원관리')
    .addItem('🔍 기존 시트 분석', 'showAnalysisDialog')
    .addItem('📥 기존 데이터 가져오기', 'copyExistingData')
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ 출퇴근 관리')
      .addItem('✅ 출근 체크', 'checkIn')
      .addItem('🏠 퇴근 체크', 'checkOut')
      .addItem('📋 출퇴근 현황 보기', 'showAttendanceStatus'))
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 급여 관리')
      .addItem('⚙️ 시급 설정', 'showHourlyWageSettings')
      .addItem('🎯 인센티브 설정', 'showIncentiveSettings')
      .addItem('📊 급여 계산', 'calculateSalary')
      .addItem('💵 급여 명세서 보기', 'showSalarySlip'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📦 플랫폼 정산')
      .addItem('📥 정산 데이터 가져오기', 'importPlatformData')
      .addItem('📊 플랫폼별 통계', 'showPlatformStatistics'))
    .addSeparator()
    .addItem('⚙️ 전체 시스템 초기화', 'initializeSystem')
    .addToUi();
}

/**
 * 전체 시스템 초기화
 */
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '시스템 초기화',
    '기존 스프레드시트의 데이터를 가져와서 새로운 관리 시스템을 구축합니다.\n\n계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // 1. 기존 데이터 복사
  copyExistingData();

  // 2. 필요한 시트 생성
  createManagementSheets();

  // 3. 설정 시트 생성
  createSettingsSheet();

  ui.alert('✅ 시스템 초기화가 완료되었습니다!');
}

/**
 * 관리용 시트 생성
 */
function createManagementSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 출퇴근 기록 시트
  let attendanceSheet = ss.getSheetByName('출퇴근기록');
  if (!attendanceSheet) {
    attendanceSheet = ss.insertSheet('출퇴근기록');
    attendanceSheet.getRange('A1:G1').setValues([[
      '날짜', '이름', '부서', '출근시간', '퇴근시간', '근무시간', '비고'
    ]]);
    attendanceSheet.getRange('A1:G1')
      .setFontWeight('bold')
      .setBackground('#34a853')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    attendanceSheet.setFrozenRows(1);
  }

  // 급여 계산 시트
  let salarySheet = ss.getSheetByName('급여계산');
  if (!salarySheet) {
    salarySheet = ss.insertSheet('급여계산');
    salarySheet.getRange('A1:K1').setValues([[
      '이름', '부서', '근무시간', '시급', '기본급',
      '배민건수', '쿠팡건수', '요기요건수', '땡겨요건수',
      '인센티브합계', '총급여'
    ]]);
    salarySheet.getRange('A1:K1')
      .setFontWeight('bold')
      .setBackground('#fbbc04')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    salarySheet.setFrozenRows(1);
  }

  // 플랫폼 정산 통합 시트
  let platformSheet = ss.getSheetByName('플랫폼정산통합');
  if (!platformSheet) {
    platformSheet = ss.insertSheet('플랫폼정산통합');
    platformSheet.getRange('A1:H1').setValues([[
      '접수날짜', '플랫폼', '사업자번호', '상호명', '타입',
      '담당자', '금액', '비고'
    ]]);
    platformSheet.getRange('A1:H1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    platformSheet.setFrozenRows(1);
  }
}

/**
 * 설정 시트 생성
 */
function createSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let settingsSheet = ss.getSheetByName('설정');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('설정');

    // 시급 설정
    settingsSheet.getRange('A1').setValue('=== 시급 설정 ===').setFontWeight('bold').setFontSize(12);
    settingsSheet.getRange('A2:C2').setValues([['이름', '부서', '시급']]);
    settingsSheet.getRange('A2:C2').setFontWeight('bold').setBackground('#e8f0fe');

    // 인센티브 설정
    settingsSheet.getRange('E1').setValue('=== 인센티브 단가 설정 ===').setFontWeight('bold').setFontSize(12);
    settingsSheet.getRange('E2:F2').setValues([['플랫폼', '건당 인센티브']]);
    settingsSheet.getRange('E2:F2').setFontWeight('bold').setBackground('#fce8e6');
    settingsSheet.getRange('E3:F6').setValues([
      ['배민', 1000],
      ['쿠팡', 1000],
      ['요기요', 1000],
      ['땡겨요', 1000]
    ]);

    settingsSheet.setColumnWidth(1, 150);
    settingsSheet.setColumnWidth(2, 100);
    settingsSheet.setColumnWidth(3, 100);
    settingsSheet.setColumnWidth(5, 150);
    settingsSheet.setColumnWidth(6, 120);
  }
}
