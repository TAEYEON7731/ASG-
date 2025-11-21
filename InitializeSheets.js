/**
 * ASG 직원 관리 시스템 - 시트 초기화 (수정 버전)
 */

function initializeAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    '시스템 초기화',
    '전체 시스템을 초기화하고 새로운 시트를 생성합니다.\n\n계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  // 기존 시트들 제거
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name === 'Sheet1' || name === '시트1') {
      ss.deleteSheet(sheet);
    }
  });

  // 새로운 시트 생성
  create_EmployeeInfoSheet();
  create_AttendanceSheet();
  create_SalarySheet();
  create_AnnualLeaveSheet();
  create_SettingsSheet();
  create_DashboardSheet();  // 대시보드는 마지막에 생성

  // 시트 순서 정렬
  arrangeSheetOrder();

  ui.alert('✅ 시스템 초기화 완료!',
           '모든 시트가 생성되었습니다.\n직원정보 시트에서 직원 정보를 입력해주세요.',
           ui.ButtonSet.OK);
}

/**
 * 1. 직원정보
 */
function create_EmployeeInfoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('직원정보');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('직원정보');
  sheet.setTabColor('#34a853');

  // 헤더
  const headers = [
    '사번', '이름', '부서', '직급', '입사일',
    '연락처', '이메일', '상태', '시급', '급여형태', '비고'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 40);

  // 날짜 형식
  sheet.getRange(2, 5, 100, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 9, 100, 1).setNumberFormat('#,##0"원"');

  // 상태 열에 조건부 서식
  const statusRange = sheet.getRange('H2:H100');
  let rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('재직')
    .setBackground('#d4edda')
    .setFontColor('#155724')
    .setRanges([statusRange])
    .build();

  let rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('퇴사')
    .setBackground('#f8d7da')
    .setFontColor('#721c24')
    .setRanges([statusRange])
    .build();

  sheet.setConditionalFormatRules([rule, rule2]);

  // 열 너비
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 180);
  sheet.setColumnWidth(8, 80);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 200);

  // 데이터 검증
  const deptRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TM팀', '행정팀'], true)
    .build();
  sheet.getRange('C2:C100').setDataValidation(deptRule);

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['재직', '휴직', '퇴사'], true)
    .build();
  sheet.getRange('H2:H100').setDataValidation(statusRule);

  const salaryTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['시급제', '연봉제'], true)
    .build();
  sheet.getRange('J2:J100').setDataValidation(salaryTypeRule);
}

/**
 * 2. 출퇴근기록
 */
function create_AttendanceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('출퇴근기록');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('출퇴근기록');
  sheet.setTabColor('#fbbc04');

  // 헤더
  const headers = [
    '날짜', '요일', '이름', '부서',
    '출근시간', '퇴근시간', '근무시간', '비고'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#fbbc04')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 40);

  // 서식 설정
  sheet.getRange(2, 1, 1000, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 5, 1000, 2).setNumberFormat('hh:mm');
  sheet.getRange(2, 7, 1000, 1).setNumberFormat('0.0"시간"');

  // 조건부 서식 (8시간 이상 근무시 초록색)
  const workHoursRange = sheet.getRange('G2:G1000');
  let rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(8)
    .setBackground('#d4edda')
    .setRanges([workHoursRange])
    .build();

  sheet.setConditionalFormatRules([rule]);

  // 열 너비
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 200);
}

/**
 * 3. 급여계산
 */
function create_SalarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('급여계산');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('급여계산');
  sheet.setTabColor('#ea4335');

  // 상단 정보
  sheet.getRange('A1').setValue('기준 년월:');
  sheet.getRange('B1').setValue(new Date());
  sheet.getRange('B1').setNumberFormat('yyyy-mm');
  sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#fff3cd');

  // 헤더 (플랫폼 인센티브 제거)
  const headers = [
    '이름', '부서', '급여형태', '시급',
    '총근무시간', '기본급', '총급여', '비고'
  ];

  sheet.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#ea4335')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(2);
  sheet.setRowHeight(2, 40);

  // 서식 설정
  sheet.getRange(3, 4, 100, 1).setNumberFormat('#,##0"원"');
  sheet.getRange(3, 5, 100, 1).setNumberFormat('0.0"시간"');
  sheet.getRange(3, 6, 100, 2).setNumberFormat('#,##0"원"');

  // 총급여 열 강조
  sheet.getRange(2, 7, 100, 1).setBackground('#fff3cd');

  // 열 너비
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 200);
}

/**
 * 4. 연차관리 (이미지 기반 재작성 대기)
 */
function create_AnnualLeaveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('연차관리');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('연차관리');
  sheet.setTabColor('#00bcd4');

  // 임시 헤더 (이미지 확인 후 수정 예정)
  const headers = [
    '이름', '입사일', '발생일수', '사용일수', '잔여일수', '비고'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#00bcd4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 40);

  // 열 너비
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 200);
}

/**
 * 5. 설정
 */
function create_SettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('⚙️ 설정');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('⚙️ 설정');
  sheet.setTabColor('#607d8b');

  // 제목
  sheet.getRange('A1').setValue('시스템 설정').setFontSize(16).setFontWeight('bold');
  sheet.setRowHeight(1, 40);

  // 기본 설정
  sheet.getRange('A3').setValue('기본 설정').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A4:B4').setValues([['항목', '값']]).setFontWeight('bold').setBackground('#f8f9fa');

  const basicSettings = [
    ['기본 시급', 13000],
    ['기본 출근시간', '09:00'],
    ['기본 퇴근시간', '18:00'],
    ['정규 근무시간', 8],
    ['주 근무일', 5]
  ];

  sheet.getRange(5, 1, basicSettings.length, 2).setValues(basicSettings);

  // 열 너비
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 120);
}

/**
 * 6. 대시보드 (간소화 버전)
 */
function create_DashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('📊 대시보드');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('📊 대시보드');
  sheet.setTabColor('#4285f4');

  // 제목
  sheet.getRange('A1:F1').merge()
    .setValue('ASG 직원 관리 시스템')
    .setFontSize(24)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  sheet.setRowHeight(1, 60);

  // 현재 날짜
  sheet.getRange('A2:F2').merge()
    .setFormula('="업데이트: " & TEXT(TODAY(), "YYYY년 MM월 DD일")')
    .setHorizontalAlignment('center')
    .setFontSize(11)
    .setFontColor('#666666');

  sheet.setRowHeight(3, 10);

  // 주요 지표 (간단한 카운트만)
  let row = 4;

  // 전체 직원 수
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('📋 전체 직원 수')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#f8f9fa')
    .setVerticalAlignment('middle');

  sheet.getRange(row, 3, 1, 2).merge()
    .setFormula('=COUNTA(직원정보!B2:B100)')
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(row, 50);
  row++;

  // 금일 출근 인원
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('✅ 금일 출근 인원')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#f8f9fa')
    .setVerticalAlignment('middle');

  sheet.getRange(row, 3, 1, 2).merge()
    .setFormula('=COUNTIF(출퇴근기록!A:A, TODAY())')
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(row, 50);
  row++;

  // 이번 달 총 급여
  sheet.getRange(row, 1, 1, 2).merge()
    .setValue('💰 이번 달 총 급여')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#f8f9fa')
    .setVerticalAlignment('middle');

  sheet.getRange(row, 3, 1, 2).merge()
    .setFormula('=SUM(급여계산!G:G)')
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#ffffff')
    .setNumberFormat('#,##0"원"')
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(row, 50);
  row += 2;

  // 빠른 이동 (하이퍼링크)
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('📌 빠른 이동')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f8f9fa');
  row++;

  // 하이퍼링크 버튼 생성
  const links = [
    { name: '👥 직원정보 보기', sheet: '직원정보' },
    { name: '⏰ 출퇴근 기록', sheet: '출퇴근기록' },
    { name: '💵 급여 계산', sheet: '급여계산' },
    { name: '🏖️ 연차 관리', sheet: '연차관리' }
  ];

  links.forEach(link => {
    const cell = sheet.getRange(row, 1, 1, 2).merge();
    cell.setValue(link.name)
      .setFontSize(11)
      .setBackground('#ffffff')
      .setFontColor('#1a73e8')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

    // 하이퍼링크 설정
    const targetSheet = ss.getSheetByName(link.sheet);
    if (targetSheet) {
      const formula = '=HYPERLINK("#gid=' + targetSheet.getSheetId() + '", "' + link.name + '")';
      cell.setFormula(formula);
    }

    row++;
  });

  // 열 너비
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
}

/**
 * 시트 순서 정렬
 */
function arrangeSheetOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const order = [
    '📊 대시보드',
    '직원정보',
    '출퇴근기록',
    '급여계산',
    '연차관리',
    '⚙️ 설정'
  ];

  order.forEach((name, index) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    }
  });

  const dashboard = ss.getSheetByName('📊 대시보드');
  if (dashboard) {
    ss.setActiveSheet(dashboard);
  }
}
