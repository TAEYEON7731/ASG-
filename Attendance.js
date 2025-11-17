/**
 * 출퇴근 자동화 시스템
 */

/**
 * 출근 체크
 */
function checkIn() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('출퇴근기록');

  if (!sheet) {
    ui.alert('❌ 출퇴근기록 시트가 없습니다. 시스템 초기화를 먼저 실행해주세요.');
    return;
  }

  // 사용자 정보 입력
  const nameResponse = ui.prompt('출근 체크', '이름을 입력하세요:', ui.ButtonSet.OK_CANCEL);

  if (nameResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const name = nameResponse.getResponseText().trim();
  if (!name) {
    ui.alert('❌ 이름을 입력해주세요.');
    return;
  }

  const now = new Date();
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

  // 오늘 이미 출근했는지 확인
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    const rowName = data[i][1];

    if (rowDate === today && rowName === name) {
      ui.alert('ℹ️ 오늘 이미 출근 체크되었습니다.\n출근시간: ' + data[i][3]);
      return;
    }
  }

  // 부서 정보 가져오기 (기존 직원 목록에서)
  const department = getEmployeeDepartment(name);

  // 출근 기록 추가
  sheet.appendRow([
    now,
    name,
    department,
    timeStr,
    '',  // 퇴근시간 (나중에 입력)
    '',  // 근무시간 (퇴근시 자동 계산)
    ''   // 비고
  ]);

  ui.alert('✅ 출근 체크 완료!\n\n이름: ' + name + '\n부서: ' + department + '\n출근시간: ' + timeStr);
}

/**
 * 퇴근 체크
 */
function checkOut() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('출퇴근기록');

  if (!sheet) {
    ui.alert('❌ 출퇴근기록 시트가 없습니다.');
    return;
  }

  // 사용자 정보 입력
  const nameResponse = ui.prompt('퇴근 체크', '이름을 입력하세요:', ui.ButtonSet.OK_CANCEL);

  if (nameResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const name = nameResponse.getResponseText().trim();
  if (!name) {
    ui.alert('❌ 이름을 입력해주세요.');
    return;
  }

  const now = new Date();
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

  // 오늘 출근 기록 찾기
  const data = sheet.getDataRange().getValues();
  let foundRow = -1;

  for (let i = data.length - 1; i >= 1; i--) {
    const rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    const rowName = data[i][1];

    if (rowDate === today && rowName === name) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow === -1) {
    ui.alert('❌ 오늘 출근 기록이 없습니다. 먼저 출근 체크를 해주세요.');
    return;
  }

  // 이미 퇴근했는지 확인
  const checkOutTime = sheet.getRange(foundRow, 5).getValue();
  if (checkOutTime) {
    ui.alert('ℹ️ 이미 퇴근 체크되었습니다.\n퇴근시간: ' + checkOutTime);
    return;
  }

  // 퇴근 시간 기록
  sheet.getRange(foundRow, 5).setValue(timeStr);

  // 근무 시간 계산
  const checkInTime = sheet.getRange(foundRow, 4).getValue();
  const workHours = calculateWorkHours(checkInTime, timeStr);
  sheet.getRange(foundRow, 6).setValue(workHours);

  ui.alert('✅ 퇴근 체크 완료!\n\n이름: ' + name + '\n퇴근시간: ' + timeStr + '\n근무시간: ' + workHours + '시간');
}

/**
 * 근무 시간 계산
 */
function calculateWorkHours(checkInTime, checkOutTime) {
  const checkIn = new Date('2000-01-01 ' + checkInTime);
  const checkOut = new Date('2000-01-01 ' + checkOutTime);

  const diff = checkOut - checkIn;
  const hours = diff / (1000 * 60 * 60);

  return Math.round(hours * 100) / 100;  // 소수점 2자리
}

/**
 * 직원 부서 정보 가져오기
 */
function getEmployeeDepartment(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 기존 직원 목록 시트들에서 검색
  const sheetNames = ['직원목록', '직원명단', '인원명단'];

  for (let sheetName of sheetNames) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // 이름이 포함된 열 찾기
      for (let j = 0; j < data[i].length; j++) {
        if (data[i][j] === name) {
          // 부서 정보는 보통 이름 다음 열에 있음
          if (j + 1 < data[i].length) {
            return data[i][j + 1] || '미지정';
          }
        }
      }
    }
  }

  return '미지정';
}

/**
 * 출퇴근 현황 보기
 */
function showAttendanceStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('출퇴근기록');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ 출퇴근기록 시트가 없습니다.');
    return;
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const data = sheet.getDataRange().getValues();

  let status = '📊 오늘의 출퇴근 현황 (' + today + ')\n\n';
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';

    if (rowDate === today) {
      count++;
      const name = data[i][1];
      const checkIn = data[i][3];
      const checkOut = data[i][4];
      const workHours = data[i][5];

      status += '👤 ' + name + '\n';
      status += '   출근: ' + checkIn;
      if (checkOut) {
        status += ' | 퇴근: ' + checkOut + ' | 근무: ' + workHours + 'h';
      } else {
        status += ' | 근무중...';
      }
      status += '\n\n';
    }
  }

  if (count === 0) {
    status += '오늘 출근 기록이 없습니다.';
  }

  SpreadsheetApp.getUi().alert(status);
}

/**
 * 월별 근무시간 집계
 */
function getMonthlyWorkHours(name, year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('출퇴근기록');

  if (!sheet) {
    return 0;
  }

  const data = sheet.getDataRange().getValues();
  let totalHours = 0;

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    const date = new Date(data[i][0]);
    const rowYear = date.getFullYear();
    const rowMonth = date.getMonth() + 1;
    const rowName = data[i][1];
    const workHours = parseFloat(data[i][5]) || 0;

    if (rowYear === year && rowMonth === month && rowName === name) {
      totalHours += workHours;
    }
  }

  return Math.round(totalHours * 100) / 100;
}
