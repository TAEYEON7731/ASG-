/**
 * 급여 자동 계산 시스템 (시급제 + 인센티브)
 */

/**
 * 시급 설정 다이얼로그
 */
function showHourlyWageSettings() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Noto Sans KR', sans-serif; padding: 20px; }
          table { width: 100%; border-collapse: collapse; }
          th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
          th { background-color: #4285f4; color: white; }
          input { width: 90%; padding: 5px; }
          button { background-color: #34a853; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-top: 10px; }
          button:hover { background-color: #2d8e47; }
        </style>
      </head>
      <body>
        <h2>시급 설정</h2>
        <div id="wageList"></div>
        <button onclick="saveWages()">저장</button>
        <div id="message" style="margin-top: 10px;"></div>

        <script>
          window.onload = function() {
            google.script.run.withSuccessHandler(displayWages).getWageList();
          };

          function displayWages(wages) {
            let html = '<table>';
            html += '<tr><th>이름</th><th>부서</th><th>시급 (원)</th></tr>';

            wages.forEach(function(wage, index) {
              html += '<tr>';
              html += '<td>' + wage.name + '</td>';
              html += '<td>' + wage.department + '</td>';
              html += '<td><input type="number" id="wage_' + index + '" value="' + (wage.hourlyWage || 10000) + '"></td>';
              html += '</tr>';
            });

            html += '</table>';
            document.getElementById('wageList').innerHTML = html;
          }

          function saveWages() {
            const inputs = document.querySelectorAll('input[type="number"]');
            const wages = [];

            inputs.forEach(function(input) {
              wages.push(parseInt(input.value) || 10000);
            });

            google.script.run
              .withSuccessHandler(function() {
                document.getElementById('message').innerHTML = '<span style="color: green;">✅ 저장되었습니다!</span>';
                setTimeout(function() { document.getElementById('message').innerHTML = ''; }, 2000);
              })
              .saveWages(wages);
          }
        </script>
      </body>
    </html>
  `).setWidth(600).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, '시급 설정');
}

/**
 * 시급 목록 가져오기
 */
function getWageList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('설정');

  if (!settingsSheet) {
    return [];
  }

  const data = settingsSheet.getDataRange().getValues();
  const wages = [];

  for (let i = 2; i < data.length; i++) {
    if (data[i][0] && data[i][0] !== '=== 시급 설정 ===') {
      wages.push({
        name: data[i][0],
        department: data[i][1],
        hourlyWage: data[i][2] || 10000
      });
    } else if (data[i][0] === '' && wages.length > 0) {
      break;
    }
  }

  return wages;
}

/**
 * 시급 저장
 */
function saveWages(wages) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('설정');

  if (!settingsSheet) {
    return;
  }

  const wageList = getWageList();

  for (let i = 0; i < wageList.length && i < wages.length; i++) {
    settingsSheet.getRange(i + 3, 3).setValue(wages[i]);
  }
}

/**
 * 인센티브 설정 다이얼로그
 */
function showIncentiveSettings() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Noto Sans KR', sans-serif; padding: 20px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
          th { background-color: #ea4335; color: white; }
          input { width: 90%; padding: 8px; font-size: 14px; }
          button { background-color: #34a853; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-top: 15px; }
          button:hover { background-color: #2d8e47; }
        </style>
      </head>
      <body>
        <h2>플랫폼별 인센티브 단가 설정</h2>
        <table>
          <tr><th>플랫폼</th><th>건당 인센티브 (원)</th></tr>
          <tr><td>배민</td><td><input type="number" id="baemin" value="1000"></td></tr>
          <tr><td>쿠팡</td><td><input type="number" id="coupang" value="1000"></td></tr>
          <tr><td>요기요</td><td><input type="number" id="yogiyo" value="1000"></td></tr>
          <tr><td>땡겨요</td><td><input type="number" id="ddangyo" value="1000"></td></tr>
        </table>
        <button onclick="saveIncentives()">저장</button>
        <div id="message" style="margin-top: 10px;"></div>

        <script>
          window.onload = function() {
            google.script.run.withSuccessHandler(displayIncentives).getIncentiveSettings();
          };

          function displayIncentives(incentives) {
            document.getElementById('baemin').value = incentives['배민'] || 1000;
            document.getElementById('coupang').value = incentives['쿠팡'] || 1000;
            document.getElementById('yogiyo').value = incentives['요기요'] || 1000;
            document.getElementById('ddangyo').value = incentives['땡겨요'] || 1000;
          }

          function saveIncentives() {
            const incentives = {
              '배민': parseInt(document.getElementById('baemin').value) || 1000,
              '쿠팡': parseInt(document.getElementById('coupang').value) || 1000,
              '요기요': parseInt(document.getElementById('yogiyo').value) || 1000,
              '땡겨요': parseInt(document.getElementById('ddangyo').value) || 1000
            };

            google.script.run
              .withSuccessHandler(function() {
                document.getElementById('message').innerHTML = '<span style="color: green;">✅ 저장되었습니다!</span>';
                setTimeout(function() { document.getElementById('message').innerHTML = ''; }, 2000);
              })
              .saveIncentiveSettings(incentives);
          }
        </script>
      </body>
    </html>
  `).setWidth(500).setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, '인센티브 설정');
}

/**
 * 인센티브 설정 가져오기
 */
function getIncentiveSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('설정');

  if (!settingsSheet) {
    return { '배민': 1000, '쿠팡': 1000, '요기요': 1000, '땡겨요': 1000 };
  }

  const data = settingsSheet.getDataRange().getValues();
  const incentives = {};

  for (let i = 0; i < data.length; i++) {
    if (data[i][4]) {  // E열 (플랫폼명)
      const platform = data[i][4];
      const amount = data[i][5] || 1000;
      incentives[platform] = amount;
    }
  }

  return incentives;
}

/**
 * 인센티브 설정 저장
 */
function saveIncentiveSettings(incentives) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('설정');

  if (!settingsSheet) {
    return;
  }

  const platforms = ['배민', '쿠팡', '요기요', '땡겨요'];

  for (let i = 0; i < platforms.length; i++) {
    settingsSheet.getRange(i + 3, 6).setValue(incentives[platforms[i]]);
  }
}

/**
 * 급여 계산 (특정 월)
 */
function calculateSalary() {
  const ui = SpreadsheetApp.getUi();

  const yearResponse = ui.prompt('급여 계산', '연도를 입력하세요 (예: 2024):', ui.ButtonSet.OK_CANCEL);
  if (yearResponse.getSelectedButton() !== ui.Button.OK) return;

  const monthResponse = ui.prompt('급여 계산', '월을 입력하세요 (예: 11):', ui.ButtonSet.OK_CANCEL);
  if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

  const year = parseInt(yearResponse.getResponseText());
  const month = parseInt(monthResponse.getResponseText());

  if (!year || !month || month < 1 || month > 12) {
    ui.alert('❌ 올바른 연도와 월을 입력해주세요.');
    return;
  }

  processMonthlySalary(year, month);
  ui.alert('✅ ' + year + '년 ' + month + '월 급여 계산이 완료되었습니다!');
}

/**
 * 월별 급여 처리
 */
function processMonthlySalary(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salarySheet = ss.getSheetByName('급여계산');

  if (!salarySheet) {
    return;
  }

  // 기존 급여 계산 데이터 삭제 (헤더 제외)
  if (salarySheet.getLastRow() > 1) {
    salarySheet.getRange(2, 1, salarySheet.getLastRow() - 1, salarySheet.getLastColumn()).clearContent();
  }

  const wageList = getWageList();
  const incentives = getIncentiveSettings();

  let row = 2;
  wageList.forEach(function(employee) {
    // 1. 근무시간 가져오기
    const workHours = getMonthlyWorkHours(employee.name, year, month);

    // 2. 기본급 계산 (근무시간 × 시급)
    const baseSalary = workHours * employee.hourlyWage;

    // 3. 플랫폼별 건수 가져오기
    const platformCounts = getPlatformCounts(employee.name, year, month);

    // 4. 인센티브 계산
    const baeminIncentive = platformCounts['배민'] * incentives['배민'];
    const coupangIncentive = platformCounts['쿠팡'] * incentives['쿠팡'];
    const yogiyoIncentive = platformCounts['요기요'] * incentives['요기요'];
    const ddangyoIncentive = platformCounts['땡겨요'] * incentives['땡겨요'];
    const totalIncentive = baeminIncentive + coupangIncentive + yogiyoIncentive + ddangyoIncentive;

    // 5. 총 급여
    const totalSalary = baseSalary + totalIncentive;

    // 6. 급여 시트에 기록
    salarySheet.getRange(row, 1, 1, 11).setValues([[
      employee.name,
      employee.department,
      workHours,
      employee.hourlyWage,
      baseSalary,
      platformCounts['배민'],
      platformCounts['쿠팡'],
      platformCounts['요기요'],
      platformCounts['땡겨요'],
      totalIncentive,
      totalSalary
    ]]);

    row++;
  });

  // 숫자 포맷 적용
  if (row > 2) {
    salarySheet.getRange(2, 3, row - 2, 1).setNumberFormat('#,##0.00');  // 근무시간
    salarySheet.getRange(2, 4, row - 2, 1).setNumberFormat('#,##0');     // 시급
    salarySheet.getRange(2, 5, row - 2, 1).setNumberFormat('#,##0');     // 기본급
    salarySheet.getRange(2, 10, row - 2, 2).setNumberFormat('#,##0');    // 인센티브, 총급여
  }
}

/**
 * 플랫폼별 건수 집계
 */
function getPlatformCounts(name, year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const platformSheet = ss.getSheetByName('플랫폼정산통합');

  const counts = {
    '배민': 0,
    '쿠팡': 0,
    '요기요': 0,
    '땡겨요': 0
  };

  if (!platformSheet) {
    return counts;
  }

  const data = platformSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    const date = new Date(data[i][0]);  // 접수날짜
    const platform = data[i][1];        // 플랫폼
    const assignee = data[i][5];        // 담당자

    if (date.getFullYear() === year &&
        date.getMonth() + 1 === month &&
        assignee === name) {
      if (counts.hasOwnProperty(platform)) {
        counts[platform]++;
      }
    }
  }

  return counts;
}

/**
 * 급여 명세서 보기
 */
function showSalarySlip() {
  const ui = SpreadsheetApp.getUi();

  const nameResponse = ui.prompt('급여 명세서', '이름을 입력하세요:', ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;

  const name = nameResponse.getResponseText().trim();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salarySheet = ss.getSheetByName('급여계산');

  if (!salarySheet) {
    ui.alert('❌ 급여계산 시트가 없습니다.');
    return;
  }

  const data = salarySheet.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      found = true;
      const slip = '💰 급여 명세서\n\n' +
        '이름: ' + data[i][0] + '\n' +
        '부서: ' + data[i][1] + '\n\n' +
        '=== 기본급 ===\n' +
        '근무시간: ' + data[i][2] + ' 시간\n' +
        '시급: ' + data[i][3].toLocaleString() + ' 원\n' +
        '기본급: ' + data[i][4].toLocaleString() + ' 원\n\n' +
        '=== 인센티브 ===\n' +
        '배민: ' + data[i][5] + '건 → ' + (data[i][5] * getIncentiveSettings()['배민']).toLocaleString() + ' 원\n' +
        '쿠팡: ' + data[i][6] + '건 → ' + (data[i][6] * getIncentiveSettings()['쿠팡']).toLocaleString() + ' 원\n' +
        '요기요: ' + data[i][7] + '건 → ' + (data[i][7] * getIncentiveSettings()['요기요']).toLocaleString() + ' 원\n' +
        '땡겨요: ' + data[i][8] + '건 → ' + (data[i][8] * getIncentiveSettings()['땡겨요']).toLocaleString() + ' 원\n' +
        '인센티브 합계: ' + data[i][9].toLocaleString() + ' 원\n\n' +
        '=== 총 급여 ===\n' +
        data[i][10].toLocaleString() + ' 원';

      ui.alert(slip);
      break;
    }
  }

  if (!found) {
    ui.alert('❌ ' + name + '님의 급여 정보를 찾을 수 없습니다.\n먼저 급여 계산을 실행해주세요.');
  }
}
