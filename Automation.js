/**
 * ASG 직원 관리 시스템 - 자동화 기능
 * 출퇴근, 급여계산, 통계
 */

/**
 * 출근 체크
 */
function checkIn() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName('출퇴근기록');
  const employeeSheet = ss.getSheetByName('직원정보');

  if (!attendanceSheet || !employeeSheet) {
    ui.alert('❌ 오류', '시트가 없습니다. 시스템 초기화를 먼저 실행해주세요.', ui.ButtonSet.OK);
    return;
  }

  // 직원 목록 가져오기
  const employeeData = employeeSheet.getRange('B3:B100').getValues();
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
          '출근시간: ' + data[i][4],
          ui.ButtonSet.OK
        );
        return;
      }
    }
  }

  // 부서 정보 가져오기
  const empInfo = getEmployeeInfo(name);

  // 현재 시간
  const now = new Date();
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');

  // 출근 기록 추가
  attendanceSheet.appendRow([
    now,
    '=TEXT(A' + (attendanceSheet.getLastRow() + 1) + ',"ddd")',
    name,
    empInfo.department,
    timeStr,
    '',  // 퇴근시간
    '',  // 근무시간 (나중에 자동 계산)
    '',  // 연장근무
    ''   // 비고
  ]);

  ui.alert(
    '✅ 출근 완료',
    name + '님 출근이 기록되었습니다.\n\n' +
    '출근시간: ' + timeStr + '\n' +
    '부서: ' + empInfo.department,
    ui.ButtonSet.OK
  );
}

/**
 * 퇴근 체크
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

  // 이미 퇴근했는지 확인
  const checkOutTime = attendanceSheet.getRange(foundRow, 6).getValue();
  if (checkOutTime) {
    ui.alert(
      'ℹ️ 알림',
      '이미 퇴근 체크되었습니다.\n\n퇴근시간: ' + checkOutTime,
      ui.ButtonSet.OK
    );
    return;
  }

  // 현재 시간
  const now = new Date();
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');

  // 퇴근 시간 기록
  attendanceSheet.getRange(foundRow, 6).setValue(timeStr);

  // 근무시간 계산 (수식 이미 설정되어 있음)
  const checkInTime = attendanceSheet.getRange(foundRow, 5).getValue();
  const workHours = attendanceSheet.getRange(foundRow, 7).getValue();

  ui.alert(
    '🏠 퇴근 완료',
    name + '님 퇴근이 기록되었습니다.\n\n' +
    '출근시간: ' + checkInTime + '\n' +
    '퇴근시간: ' + timeStr + '\n' +
    '근무시간: ' + (workHours ? workHours.toFixed(1) + '시간' : '계산 중...'),
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
        status += '   출근: ' + checkIn;

        if (checkOut) {
          status += ' | 퇴근: ' + checkOut;
          if (workHours) {
            status += ' | ' + (typeof workHours === 'number' ? workHours.toFixed(1) : workHours) + '시간';
          }
        } else {
          status += ' | 근무중...';
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
 * 이번 달 급여 계산
 */
function calculateThisMonthSalary() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const result = ui.alert(
    '💰 급여 계산',
    '이번 달 급여를 계산하시겠습니까?\n\n' +
    '출퇴근 기록과 플랫폼 인센티브를 기반으로\n' +
    '급여계산 시트가 자동으로 업데이트됩니다.',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  const employeeSheet = ss.getSheetByName('직원정보');
  const salarySheet = ss.getSheetByName('급여계산');
  const platformSheet = ss.getSheetByName('플랫폼인센티브');

  if (!employeeSheet || !salarySheet) {
    ui.alert('❌ 오류', '필요한 시트가 없습니다.', ui.ButtonSet.OK);
    return;
  }

  // 기존 데이터 클리어 (헤더 제외)
  if (salarySheet.getLastRow() > 2) {
    salarySheet.getRange(3, 1, salarySheet.getLastRow() - 2, salarySheet.getLastColumn()).clearContent();
  }

  // 직원 목록
  const employeeData = employeeSheet.getRange('B3:J100').getValues();
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth() + 1;

  let row = 3;

  employeeData.forEach(emp => {
    if (emp[0] && emp[6] === '재직') {  // 이름이 있고 재직 중인 경우
      const name = emp[0];
      const department = emp[1];
      const salaryType = emp[8] || '시급제';
      const hourlyWage = emp[7] || 13000;

      // 플랫폼별 건수 계산
      const platformCounts = getPlatformCountsForEmployee(name, currentYear, currentMonth);

      // 급여계산 시트에 데이터 추가
      salarySheet.getRange(row, 1, 1, 13).setValues([[
        name,
        department,
        salaryType,
        hourlyWage,
        '=SUMIFS(출퇴근기록!G:G, 출퇴근기록!C:C, A' + row + ', 출퇴근기록!A:A, ">="&DATE(' + currentYear + ',' + currentMonth + ',1), 출퇴근기록!A:A, "<"&DATE(' + (currentMonth === 12 ? currentYear + 1 : currentYear) + ',' + (currentMonth === 12 ? 1 : currentMonth + 1) + ',1))',
        '=IF(C' + row + '="시급제", E' + row + '*D' + row + ', 0)',
        platformCounts['배민'],
        platformCounts['쿠팡이츠'],
        platformCounts['요기요'],
        platformCounts['땡겨요'],
        '=G' + row + '*설정!B5+H' + row + '*설정!B6+I' + row + '*설정!B7+J' + row + '*설정!B8',
        '=F' + row + '+K' + row + '',
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
                   '【인센티브】\n' +
                   '배민: ' + data[i][6] + '건\n' +
                   '쿠팡이츠: ' + data[i][7] + '건\n' +
                   '요기요: ' + data[i][8] + '건\n' +
                   '땡겨요: ' + data[i][9] + '건\n' +
                   '인센티브 합계: ' + Number(data[i][10]).toLocaleString() + '원\n\n' +
                   '━━━━━━━━━━━━━━━━━━\n' +
                   '💰 총 급여: ' + Number(data[i][11]).toLocaleString() + '원';

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

  for (let i = 2; i < data.length; i++) {
    if (data[i][1] === name) {  // B열: 이름
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
 * 플랫폼별 건수 집계
 */
function getPlatformCountsForEmployee(name, year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const platformSheet = ss.getSheetByName('플랫폼인센티브');

  const counts = {
    '배민': 0,
    '쿠팡이츠': 0,
    '요기요': 0,
    '땡겨요': 0
  };

  if (!platformSheet) {
    return counts;
  }

  const data = platformSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      const date = new Date(data[i][0]);
      const platform = data[i][1];
      const assignee = data[i][4];

      if (date.getFullYear() === year &&
          date.getMonth() + 1 === month &&
          assignee === name &&
          counts.hasOwnProperty(platform)) {
        counts[platform]++;
      }
    }
  }

  return counts;
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
  let totalIncentive = 0;
  let count = 0;

  for (let i = 2; i < data.length; i++) {
    if (data[i][0]) {
      count++;
      totalSalary += Number(data[i][11]) || 0;
      totalIncentive += Number(data[i][10]) || 0;
    }
  }

  const avgSalary = count > 0 ? totalSalary / count : 0;

  const stats = '📈 급여 통계\n\n' +
                '기준월: ' + salarySheet.getRange('B1').getValue() + '\n' +
                '━━━━━━━━━━━━━━━━━━\n' +
                '대상 인원: ' + count + '명\n\n' +
                '총 급여: ' + totalSalary.toLocaleString() + '원\n' +
                '총 인센티브: ' + totalIncentive.toLocaleString() + '원\n' +
                '평균 급여: ' + Math.round(avgSalary).toLocaleString() + '원';

  SpreadsheetApp.getUi().alert('급여 통계', stats, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 인센티브 통계
 */
function showIncentiveStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const platformSheet = ss.getSheetByName('플랫폼인센티브');

  if (!platformSheet) {
    SpreadsheetApp.getUi().alert('❌ 오류', '플랫폼인센티브 시트가 없습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1;

  const data = platformSheet.getDataRange().getValues();
  const platformStats = {
    '배민': 0,
    '쿠팡이츠': 0,
    '요기요': 0,
    '땡겨요': 0
  };

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      const date = new Date(data[i][0]);
      const platform = data[i][1];

      if (date.getFullYear() === year &&
          date.getMonth() + 1 === month &&
          platformStats.hasOwnProperty(platform)) {
        platformStats[platform]++;
      }
    }
  }

  const total = Object.values(platformStats).reduce((a, b) => a + b, 0);

  const stats = '🎁 플랫폼별 인센티브 통계\n\n' +
                year + '년 ' + month + '월\n' +
                '━━━━━━━━━━━━━━━━━━\n' +
                '배민: ' + platformStats['배민'] + '건\n' +
                '쿠팡이츠: ' + platformStats['쿠팡이츠'] + '건\n' +
                '요기요: ' + platformStats['요기요'] + '건\n' +
                '땡겨요: ' + platformStats['땡겨요'] + '건\n\n' +
                '총 건수: ' + total + '건';

  SpreadsheetApp.getUi().alert('인센티브 통계', stats, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 플랫폼 데이터 입력 안내
 */
function showPlatformDataInput() {
  const ui = SpreadsheetApp.getUi();

  const guide = '📥 플랫폼 데이터 입력 방법\n\n' +
                '1. 플랫폼인센티브 시트로 이동\n' +
                '2. 각 열에 데이터 입력:\n' +
                '   - 날짜\n' +
                '   - 플랫폼 (배민/쿠팡이츠/요기요/땡겨요)\n' +
                '   - 상호명\n' +
                '   - 사업자번호\n' +
                '   - 담당자 (직원 이름)\n' +
                '   - 금액\n\n' +
                '3. 급여 계산 시 자동으로 집계됩니다!\n\n' +
                '💡 Tip: 엑셀에서 복사/붙여넣기 가능합니다.';

  ui.alert('플랫폼 데이터 입력', guide, ui.ButtonSet.OK);

  // 플랫폼인센티브 시트로 이동
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const platformSheet = ss.getSheetByName('플랫폼인센티브');

  if (platformSheet) {
    ss.setActiveSheet(platformSheet);
  }
}
