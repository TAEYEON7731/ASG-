/**
 * 직원 관리 시스템 - 메인 파일
 * 스프레드시트가 열릴 때 실행되는 함수
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('📋 직원관리')
    .addItem('➕ 직원 등록', 'showAddEmployeeDialog')
    .addItem('🔍 직원 조회', 'showSearchEmployeeDialog')
    .addSeparator()
    .addItem('📊 통계 보기', 'showStatistics')
    .addSeparator()
    .addItem('⚙️ 초기 설정', 'initializeSheets')
    .addToUi();
}

/**
 * 초기 시트 구조 설정
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 직원 목록 시트 생성
  let employeeSheet = ss.getSheetByName('직원목록');
  if (!employeeSheet) {
    employeeSheet = ss.insertSheet('직원목록');
    employeeSheet.getRange('A1:I1').setValues([[
      '사번', '이름', '부서', '직급', '입사일', '연락처', '이메일', '상태', '등록일'
    ]]);
    employeeSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    employeeSheet.setFrozenRows(1);
  }

  // 근태 관리 시트 생성
  let attendanceSheet = ss.getSheetByName('근태관리');
  if (!attendanceSheet) {
    attendanceSheet = ss.insertSheet('근태관리');
    attendanceSheet.getRange('A1:F1').setValues([[
      '사번', '이름', '날짜', '출근시간', '퇴근시간', '비고'
    ]]);
    attendanceSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
    attendanceSheet.setFrozenRows(1);
  }

  // 급여 관리 시트 생성
  let salarySheet = ss.getSheetByName('급여관리');
  if (!salarySheet) {
    salarySheet = ss.insertSheet('급여관리');
    salarySheet.getRange('A1:F1').setValues([[
      '사번', '이름', '기본급', '수당', '공제', '실수령액'
    ]]);
    salarySheet.getRange('A1:F1').setFontWeight('bold').setBackground('#fbbc04').setFontColor('#ffffff');
    salarySheet.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert('✅ 초기 설정이 완료되었습니다!');
}

/**
 * 사번 자동 생성
 */
function generateEmployeeId() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('직원목록');
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return 'EMP001';
  }

  const lastId = sheet.getRange(lastRow, 1).getValue();
  const num = parseInt(lastId.replace('EMP', '')) + 1;
  return 'EMP' + String(num).padStart(3, '0');
}
