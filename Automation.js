/**
 * ASG 직원 관리 시스템 - 자동화 기능 (수정 버전)
 */

/**
 * 출근 체크 (자동 기본값 설정)
 */
function checkIn() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName('출퇴근기록');
  const employeeSheet = ss.getSheetByName('직원정보');
  const settingsSheet = ss.getSheetByName('⚙️ 설정');

  if (!attendanceSheet || !employeeSheet) {
    ui.alert('❌ 오류', '시트가 없습니다. 시스템 초기화를 먼저 실행해주세요.', ui.ButtonSet.OK);
    return;
  }

  // 직원 목록
  const employeeData = employeeSheet.getRange('B2:B100').getValues();
  const employees = employeeData.filter(row => row[0] !== '').map(row => row[0]);

  if (employees.length === 0) {
    ui.alert('❌ 오류', '직원 정보가 없습니다. 직원정보 시트를 먼저 입력해주세요.', ui.ButtonSet.OK);
    return;
  }

  // 이름 선택
  const response = ui.prompt(
    '✅ 출근 체크',
    '이름을 입력하세요:\n\n등록된 직원: ' + employees.join(', '),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const name = response.getResponseText().trim();

  if (!employees.includes(name)) {
    ui.alert('❌ 오류', '등록되지 않은 직원입니다.', ui.ButtonSet.OK);
    return;
  }

  // 오늘 이미 출근했는지 확인
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const data = attendanceSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      const recordDate = new Date(data[i][0]);
      recordDate.setHours(0, 0, 0, 0);

      if (recordDate.getTime() === today.getTime() && data[i][2] === name) {
        ui.alert(
          'ℹ️ 알림',
          name + '님은 이미 출근 체크되었습니다.\n\n' +
          '출근시간: ' + data[i][4] + '\n' +
          '퇴근시간: ' + (data[i][5] || '미체크'),
          ui.ButtonSet.OK
        );
        return;
      }
    }
  }

  // 직원 정보
  const empInfo = getEmployeeInfo(name);

  // 기본 시간 가져오기
  const defaultCheckIn = settingsSheet ? settingsSheet.getRange('B6').getValue() : '09:00';
  const defaultCheckOut = settingsSheet ? settingsSheet.getRange('B7').getValue() : '18:00';

  // 출근 기록 추가 (기본 출퇴근 시간 자동 입력)
  const newRow = attendanceSheet.getLastRow() + 1;
  const now = new Date();

  attendanceSheet.appendRow([
    now,  // 날짜
    '=TEXT(A' + newRow + ',"ddd")',  // 요일
    name,  // 이름
    empInfo.department,  // 부서
    defaultCheckIn,  // 기본 출근시간 자동 입력
    defaultCheckOut,  // 기본 퇴근시간 자동 입력
    '=IF(AND(E' + newRow + '<>"",F' + newRow + '<>""),(F' + newRow + '-E' + newRow + ')*24,"")',  // 근무시간 자동 계산
    ''  // 비고
  ]);

  ui.alert(
    '✅ 출근 완료',
    name + '님 출근이 기록되었습니다.\n\n' +
    '출근시간: ' + defaultCheckIn + ' (기본값)\n' +
    '퇴근시간: ' + defaultCheckOut + ' (기본값)\n' +
    '부서: ' + empInfo.department + '\n\n' +
    '💡 Tip: 실제 시간이 다른 경우 출퇴근기록 시트에서 수정하세요.',
    ui.ButtonSet.OK);
}

/**
 * 퇴근 체크 (기존 기록 업데이트 가능)
 */
function checkOut() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName('출퇴근기록');

  if (!attendanceSheet) {
    ui.alert('❌ 오류', '출퇴근기록 시트가 없습니다.', ui.ButtonSet.OK);
    return;
  }

  // 이름 입력
  const response = ui.prompt('🏠 퇴근 체크', '이름을 입력하세요:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const name = response.getResponseText().trim();

  // 오늘 출근 기록 찾기
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const data = attendanceSheet.getDataRange().getValues();
  let foundRow = -1;

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0]) {
      const recordDate = new Date(data[i][0]);
      recordDate.setHours(0, 0, 0, 0);

      if (recordDate.getTime() === today.getTime() && data[i][2] === name) {
        foundRow = i + 1;
        break;
      }
    }
  }

  if (foundRow === -1) {
    ui.alert('❌ 오류', name + '님의 오늘 출근 기록이 없습니다.\n먼저 출근 체크를 해주세요.', ui.ButtonSet.OK);
    return;
  }

  // 현재 시간
  const now = new Date();
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');

  // 퇴근 시간 업데이트
  attendanceSheet.getRange(foundRow, 6).setValue(timeStr);

  // 근무시간은 수식으로 자동 계산됨
  SpreadsheetApp.flush();  // 계산 강제 실행

  const checkInTime = attendanceSheet.getRange(foundRow, 5).getValue();
  const workHours = attendanceSheet.getRange(foundRow, 7).getValue();

  ui.alert(
    '🏠 퇴근 완료',
    name + '님 퇴근이 기록되었습니다.\n\n' +
    '출근시간: ' + checkInTime + '\n' +
    '퇴근시간: ' + timeStr + '\n' +
    '근무시간: ' + (workHours ? Number(workHours).toFixed(1) + '시간' : '계산 중...'),
    ui.ButtonSet.OK
  );
}

/**
 * 오늘 출퇴근 현황
 */
function showTodayAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName('출퇴근기록');

  if (!attendanceSheet) {
    SpreadsheetApp.getUi().alert('❌ 오류', '출퇴근기록 시트가 없습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const data = attendanceSheet.getDataRange().getValues();
  let status = '📋 오늘 출퇴근 현황 (' + Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd') + ')\n\n';

  let count = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      const recordDate = new Date(data[i][0]);
      recordDate.setHours(0, 0, 0, 0);

      if (recordDate.getTime() === today.getTime()) {
        count++;
        const name = data[i][2];
        const dept = data[i][3];
        const checkIn = data[i][4];
        const checkOut = data[i][5];
        const workHours = data[i][6];

        status += '👤 ' + name + ' (' + dept + ')\n';
        status += '   출근: ' + checkIn + ' | 퇴근: ' + checkOut;

        if (workHours) {
          status += ' | ' + (typeof workHours === 'number' ? workHours.toFixed(1) : workHours) + '시간';
        }

        status += '\n\n';
      }
    }
  }

  if (count === 0) {
    status += '오늘 출근 기록이 없습니다.';
  }

  SpreadsheetApp.getUi().alert('출퇴근 현황', status, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 이번 달 급여 계산 (플랫폼 인센티브 제거)
 */
function calculateThisMonthSalary() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const result = ui.alert(
    '💰 급여 계산',
    '이번 달 급여를 계산하시겠습니까?\n\n' +
    '출퇴근 기록을 기반으로\n' +
    '급여계산 시트가 자동으로 업데이트됩니다.',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  const employeeSheet = ss.getSheetByName('직원정보');
  const salarySheet = ss.getSheetByName('급여계산');

  if (!employeeSheet || !salarySheet) {
    ui.alert('❌ 오류', '필요한 시트가 없습니다.', ui.ButtonSet.OK);
    return;
  }

  // 기존 데이터 클리어
  if (salarySheet.getLastRow() > 2) {
    salarySheet.getRange(3, 1, salarySheet.getLastRow() - 2, salarySheet.getLastColumn()).clearContent();
  }

  // 직원 목록
  const employeeData = employeeSheet.getRange('B2:J100').getValues();
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth() + 1;

  let row = 3;

  employeeData.forEach(emp => {
    if (emp[0] && emp[6] === '재직') {
      const name = emp[0];
      const department = emp[1];
      const salaryType = emp[8] || '시급제';
      const hourlyWage = emp[7] || 13000;

      // 급여계산 시트에 데이터 추가
      salarySheet.getRange(row, 1, 1, 8).setValues([[
        name,
        department,
        salaryType,
        hourlyWage,
        '=SUMIFS(출퇴근기록!G:G, 출퇴근기록!C:C, A' + row + ', 출퇴근기록!A:A, ">="&DATE(' + currentYear + ',' + currentMonth + ',1), 출퇴근기록!A:A, "<"&DATE(' + (currentMonth === 12 ? currentYear + 1 : currentYear) + ',' + (currentMonth === 12 ? 1 : currentMonth + 1) + ',1))',
        '=IF(C' + row + '="시급제", E' + row + '*D' + row + ', 0)',
        '=F' + row + '',  // 총급여 = 기본급
        ''
      ]]);

      row++;
    }
  });

  // 기준 년월 업데이트
  salarySheet.getRange('B1').setValue(Utilities.formatDate(today, 'GMT+9', 'yyyy-MM'));

  ui.alert(
    '✅ 계산 완료',
    '급여 계산이 완료되었습니다!\n\n' +
    '급여계산 시트를 확인해주세요.',
    ui.ButtonSet.OK
  );
}

/**
 * 급여 명세서 보기
 */
function showSalarySlip() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salarySheet = ss.getSheetByName('급여계산');

  if (!salarySheet) {
    ui.alert('❌ 오류', '급여계산 시트가 없습니다.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.prompt('💵 급여 명세서', '이름을 입력하세요:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const name = response.getResponseText().trim();
  const data = salarySheet.getDataRange().getValues();

  for (let i = 2; i < data.length; i++) {
    if (data[i][0] === name) {
      const slip = '💵 급여 명세서\n\n' +
                   '기준월: ' + salarySheet.getRange('B1').getValue() + '\n' +
                   '━━━━━━━━━━━━━━━━━━\n' +
                   '이름: ' + data[i][0] + '\n' +
                   '부서: ' + data[i][1] + '\n' +
                   '급여형태: ' + data[i][2] + '\n\n' +
                   '【기본급】\n' +
                   '시급: ' + Number(data[i][3]).toLocaleString() + '원\n' +
                   '근무시간: ' + (data[i][4] ? Number(data[i][4]).toFixed(1) : '0.0') + '시간\n' +
                   '기본급: ' + Number(data[i][5]).toLocaleString() + '원\n\n' +
                   '━━━━━━━━━━━━━━━━━━\n' +
                   '💰 총 급여: ' + Number(data[i][6]).toLocaleString() + '원';

      ui.alert('급여 명세서', slip, ui.ButtonSet.OK);
      return;
    }
  }

  ui.alert('❌ 오류', name + '님의 급여 정보를 찾을 수 없습니다.\n먼저 급여 계산을 실행해주세요.', ui.ButtonSet.OK);
}

/**
 * 직원 정보 가져오기
 */
function getEmployeeInfo(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeSheet = ss.getSheetByName('직원정보');

  if (!employeeSheet) {
    return { department: '미지정', hourlyWage: 13000 };
  }

  const data = employeeSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === name) {
      return {
        employeeId: data[i][0],
        department: data[i][2] || '미지정',
        position: data[i][3],
        hourlyWage: data[i][8] || 13000,
        salaryType: data[i][9] || '시급제'
      };
    }
  }

  return { department: '미지정', hourlyWage: 13000 };
}

/**
 * 급여 통계
 */
function showSalaryStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salarySheet = ss.getSheetByName('급여계산');

  if (!salarySheet) {
    SpreadsheetApp.getUi().alert('❌ 오류', '급여계산 시트가 없습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const data = salarySheet.getDataRange().getValues();
  let totalSalary = 0;
  let count = 0;

  for (let i = 2; i < data.length; i++) {
    if (data[i][0]) {
      count++;
      totalSalary += Number(data[i][6]) || 0;  // 총급여
    }
  }

  const avgSalary = count > 0 ? totalSalary / count : 0;

  const stats = '📈 급여 통계\n\n' +
                '기준월: ' + salarySheet.getRange('B1').getValue() + '\n' +
                '━━━━━━━━━━━━━━━━━━\n' +
                '대상 인원: ' + count + '명\n\n' +
                '총 급여: ' + totalSalary.toLocaleString() + '원\n' +
                '평균 급여: ' + Math.round(avgSalary).toLocaleString() + '원';

  SpreadsheetApp.getUi().alert('급여 통계', stats, SpreadsheetApp.getUi().ButtonSet.OK);
}
