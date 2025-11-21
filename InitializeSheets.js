/**
 * ASG 직원 관리 시스템 - 시트 초기화
 *
 * 회사 정보:
 * - 인원: 8명
 * - 부서: TM팀, 행정팀
 * - 근무시간: 09:00-18:00 (주5일)
 * - 시급: 13,000원 (주휴수당 포함)
 * - 플랫폼: 배민, 쿠팡이츠, 요기요, 땡겨요
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

  // 기존 시트들 제거 (Sheet1 같은 기본 시트만)
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name === 'Sheet1' || name === '시트1') {
      ss.deleteSheet(sheet);
    }
  });

  // 새로운 시트 생성
  create_DashboardSheet();
  create_EmployeeInfoSheet();
  create_AttendanceSheet();
  create_SalarySheet();
  create_PlatformIncentiveSheet();
  create_AnnualLeaveSheet();
  create_SettingsSheet();

  // 시트 순서 정렬
  arrangeSheetOrder();

  ui.alert('✅ 시스템 초기화 완료!',
           '모든 시트가 생성되었습니다.\n각 시트를 확인하고 직원 정보를 입력해주세요.',
           ui.ButtonSet.OK);
}

/**
 * 1. 대시보드 (한눈에 보는 현황)
 */
function create_DashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('📊 대시보드');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('📊 대시보드');

  // 배경색 설정
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

  // 구분선
  sheet.setRowHeight(3, 10);

  // 주요 지표
  const metrics = [
    ['📋 전체 직원 수', '=COUNTA(직원정보!B3:B100)-COUNTIF(직원정보!H3:H100,"퇴사")'],
    ['✅ 금일 출근 인원', '=COUNTIF(출퇴근기록!A3:A100,TODAY())'],
    ['💰 이번 달 총 급여', '=SUM(급여계산!L3:L100)'],
    ['🎯 이번 달 인센티브', '=SUM(급여계산!K3:K100)']
  ];

  let row = 4;
  metrics.forEach((metric, index) => {
    const startRow = row;

    // 레이블
    sheet.getRange(startRow, 1, 1, 2).merge()
      .setValue(metric[0])
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground('#f8f9fa')
      .setVerticalAlignment('middle');

    // 값
    sheet.getRange(startRow, 3, 1, 2).merge()
      .setFormula(metric[1])
      .setFontSize(20)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBackground('#ffffff')
      .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

    sheet.setRowHeight(startRow, 50);
    row++;
  });

  // 구분선
  row++;
  sheet.setRowHeight(row, 10);
  row++;

  // 빠른 링크
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('📌 빠른 이동')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f8f9fa');
  row++;

  const links = [
    ['👥 직원정보 보기', '직원정보'],
    ['⏰ 출퇴근 기록', '출퇴근기록'],
    ['💵 급여 계산', '급여계산'],
    ['🎁 인센티브 정산', '플랫폼인센티브']
  ];

  links.forEach(link => {
    sheet.getRange(row, 1, 1, 2).merge()
      .setValue(link[0])
      .setFontSize(11)
      .setBackground('#ffffff')
      .setFontColor('#1a73e8')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

    // 하이퍼링크는 수동으로 설정 필요 (나중에 사용자가 클릭하면 해당 시트로 이동)
    row++;
  });

  // 열 너비 설정
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
}

/**
 * 2. 직원정보
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

  // 샘플 데이터 (대표 1명 + 직원 예시)
  const sampleData = [
    ['EMP001', '대표', '행정팀', '대표', new Date(2020, 0, 1), '010-0000-0000', 'ceo@asg.com', '재직', 0, '연봉제', ''],
    ['EMP002', '홍길동', 'TM팀', '팀장', new Date(2022, 0, 1), '010-1111-1111', 'hong@asg.com', '재직', 13000, '시급제', ''],
    ['EMP003', '김철수', 'TM팀', '사원', new Date(2023, 5, 1), '010-2222-2222', 'kim@asg.com', '재직', 13000, '시급제', ''],
    ['EMP004', '이영희', '행정팀', '사원', new Date(2023, 8, 1), '010-3333-3333', 'lee@asg.com', '재직', 13000, '시급제', '']
  ];

  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

  // 데이터 영역 서식
  const lastRow = 2 + sampleData.length - 1;
  sheet.getRange(2, 1, sampleData.length, headers.length)
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');

  // 날짜 형식
  sheet.getRange(2, 5, 100, 1).setNumberFormat('yyyy-mm-dd');

  // 시급 형식
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
  sheet.setColumnWidth(1, 80);   // 사번
  sheet.setColumnWidth(2, 100);  // 이름
  sheet.setColumnWidth(3, 100);  // 부서
  sheet.setColumnWidth(4, 100);  // 직급
  sheet.setColumnWidth(5, 120);  // 입사일
  sheet.setColumnWidth(6, 130);  // 연락처
  sheet.setColumnWidth(7, 180);  // 이메일
  sheet.setColumnWidth(8, 80);   // 상태
  sheet.setColumnWidth(9, 100);  // 시급
  sheet.setColumnWidth(10, 100); // 급여형태
  sheet.setColumnWidth(11, 200); // 비고

  // 데이터 검증 (부서)
  const deptRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TM팀', '행정팀'], true)
    .build();
  sheet.getRange('C2:C100').setDataValidation(deptRule);

  // 데이터 검증 (상태)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['재직', '휴직', '퇴사'], true)
    .build();
  sheet.getRange('H2:H100').setDataValidation(statusRule);

  // 데이터 검증 (급여형태)
  const salaryTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['시급제', '연봉제'], true)
    .build();
  sheet.getRange('J2:J100').setDataValidation(salaryTypeRule);
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
    '플랫폼인센티브',
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

  // 대시보드를 활성화
  const dashboard = ss.getSheetByName('📊 대시보드');
  if (dashboard) {
    ss.setActiveSheet(dashboard);
  }
}
