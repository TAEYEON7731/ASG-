/**
 * ASG 직원 관리 시스템 - 메뉴 (수정 버전)
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('📋 ASG 관리')
    .addItem('🏠 대시보드로 이동', 'goToDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ 출퇴근')
      .addItem('✅ 출근 체크', 'checkIn')
      .addItem('🏠 퇴근 체크', 'checkOut')
      .addItem('📋 오늘 출퇴근 현황', 'showTodayAttendance'))
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 급여')
      .addItem('📊 이번 달 급여 계산', 'calculateThisMonthSalary')
      .addItem('💵 급여 명세서 보기', 'showSalarySlip')
      .addItem('📈 급여 통계', 'showSalaryStatistics'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ 시스템')
      .addItem('🔄 시스템 초기화', 'initializeAllSheets')
      .addItem('ℹ️ 사용 가이드', 'showUserGuide')
      .addItem('📞 문의하기', 'showContact'))
    .addToUi();
}

/**
 * 대시보드로 이동
 */
function goToDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('📊 대시보드');

  if (dashboard) {
    ss.setActiveSheet(dashboard);
  } else {
    SpreadsheetApp.getUi().alert('대시보드 시트가 없습니다. 시스템 초기화를 먼저 실행해주세요.');
  }
}

/**
 * 사용 가이드
 */
function showUserGuide() {
  const ui = SpreadsheetApp.getUi();

  const guide = '📖 ASG 직원 관리 시스템 사용 가이드\n\n' +
                '【출퇴근 관리】\n' +
                '1. 출근 시: ASG 관리 > 출퇴근 > 출근 체크\n' +
                '   → 기본 시간(09:00-18:00) 자동 입력\n' +
                '2. 실제 시간이 다른 경우:\n' +
                '   → 출퇴근기록 시트에서 직접 수정\n' +
                '3. 근무시간은 자동으로 계산됩니다\n\n' +
                '【급여 계산】\n' +
                '1. 매월 말일: ASG 관리 > 급여 > 이번 달 급여 계산\n' +
                '2. 출퇴근 기록이 자동 집계됩니다\n' +
                '3. 개인별 명세서 확인 가능\n\n' +
                '【연차 관리】\n' +
                '1. 연차관리 시트에서 사용일수 입력\n' +
                '2. 잔여일수가 자동 계산됩니다\n\n' +
                '【설정】\n' +
                '- 기본 출퇴근 시간은 ⚙️ 설정 시트에서 변경 가능\n' +
                '- 시급도 설정 시트에서 관리';

  ui.alert('사용 가이드', guide, ui.ButtonSet.OK);
}

/**
 * 문의하기
 */
function showContact() {
  const ui = SpreadsheetApp.getUi();

  const contact = '📞 문의하기\n\n' +
                  '시스템 관련 문의사항이 있으시면\n' +
                  '시스템 관리자에게 연락주세요.\n\n' +
                  '이 시스템은 Claude Code로 제작되었습니다.\n' +
                  '추가 기능이 필요하면 언제든 요청해주세요!';

  ui.alert('문의하기', contact, ui.ButtonSet.OK);
}
