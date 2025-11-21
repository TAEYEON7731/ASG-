/**
 * ASG 직원 관리 시스템
 *
 * 스프레드시트가 열릴 때 실행되는 기본 메뉴 설정
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('📋 ASG 관리')
    .addItem('ℹ️ 시스템 정보', 'showSystemInfo')
    .addToUi();
}

/**
 * 시스템 정보 표시
 */
function showSystemInfo() {
  const ui = SpreadsheetApp.getUi();
  const message = 'ASG 직원 관리 시스템\n\n' +
                  '스프레드시트 구조를 작성한 후,\n' +
                  '자동화 기능이 추가될 예정입니다.\n\n' +
                  '현재 상태: 초기화 완료';

  ui.alert('시스템 정보', message, ui.ButtonSet.OK);
}
