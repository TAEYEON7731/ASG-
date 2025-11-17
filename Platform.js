/**
 * 플랫폼 정산 데이터 관리 시스템
 */

/**
 * 플랫폼 정산 데이터 가져오기
 */
function importPlatformData() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Noto Sans KR', sans-serif; padding: 20px; }
          .form-group { margin-bottom: 15px; }
          label { display: block; margin-bottom: 5px; font-weight: bold; }
          select, input { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
          textarea { width: 100%; height: 200px; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; font-family: monospace; }
          button { background-color: #4285f4; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; width: 100%; margin-top: 10px; }
          button:hover { background-color: #357ae8; }
          .info { background-color: #e8f0fe; padding: 10px; border-radius: 4px; margin-bottom: 15px; font-size: 13px; }
        </style>
      </head>
      <body>
        <h2>플랫폼 정산 데이터 가져오기</h2>

        <div class="info">
          💡 엑셀에서 데이터를 복사하여 아래에 붙여넣으세요.<br>
          형식: 접수날짜 | 사업자번호 | 상호명 | 타입 | 담당자 (탭으로 구분)
        </div>

        <div class="form-group">
          <label for="platform">플랫폼 선택 *</label>
          <select id="platform">
            <option value="배민">배민</option>
            <option value="쿠팡">쿠팡</option>
            <option value="요기요">요기요</option>
            <option value="땡겨요">땡겨요</option>
          </select>
        </div>

        <div class="form-group">
          <label for="data">데이터 (엑셀에서 복사/붙여넣기)</label>
          <textarea id="data" placeholder="접수날짜	사업자번호	상호명	타입	담당자
2024-11-01	123-45-67890	식당A	일반	홍길동
2024-11-02	098-76-54321	카페B	프리미엄	김철수"></textarea>
        </div>

        <button onclick="importData()">데이터 가져오기</button>
        <div id="message" style="margin-top: 10px;"></div>

        <script>
          function importData() {
            const platform = document.getElementById('platform').value;
            const data = document.getElementById('data').value;

            if (!data.trim()) {
              alert('데이터를 입력해주세요.');
              return;
            }

            const messageDiv = document.getElementById('message');
            messageDiv.innerHTML = '<span style="color: blue;">처리 중...</span>';

            google.script.run
              .withSuccessHandler(function(count) {
                messageDiv.innerHTML = '<span style="color: green;">✅ ' + count + '건의 데이터가 추가되었습니다!</span>';
                document.getElementById('data').value = '';
              })
              .withFailureHandler(function(error) {
                messageDiv.innerHTML = '<span style="color: red;">❌ 오류: ' + error.message + '</span>';
              })
              .processPlatformData(platform, data);
          }
        </script>
      </body>
    </html>
  `).setWidth(600).setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, '플랫폼 정산 데이터 가져오기');
}

/**
 * 플랫폼 데이터 처리
 */
function processPlatformData(platform, dataText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let platformSheet = ss.getSheetByName('플랫폼정산통합');

  if (!platformSheet) {
    // 시트가 없으면 생성
    platformSheet = ss.insertSheet('플랫폼정산통합');
    platformSheet.getRange('A1:H1').setValues([[
      '접수날짜', '플랫폼', '사업자번호', '상호명', '타입', '담당자', '금액', '비고'
    ]]);
    platformSheet.getRange('A1:H1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    platformSheet.setFrozenRows(1);
  }

  const lines = dataText.trim().split('\n');
  let count = 0;

  lines.forEach(function(line, index) {
    // 첫 줄이 헤더인 경우 건너뛰기
    if (index === 0 && (line.includes('접수날짜') || line.includes('날짜'))) {
      return;
    }

    const columns = line.split('\t');

    if (columns.length >= 4) {
      const date = parseDate(columns[0].trim());
      const businessNum = columns[1] ? columns[1].trim() : '';
      const storeName = columns[2] ? columns[2].trim() : '';
      const type = columns[3] ? columns[3].trim() : '';
      const assignee = columns[4] ? columns[4].trim() : '';
      const amount = columns[5] ? parseFloat(columns[5].replace(/[^0-9.-]/g, '')) : 0;
      const memo = columns[6] ? columns[6].trim() : '';

      platformSheet.appendRow([
        date,
        platform,
        businessNum,
        storeName,
        type,
        assignee,
        amount,
        memo
      ]);

      count++;
    }
  });

  return count;
}

/**
 * 날짜 파싱 (여러 형식 지원)
 */
function parseDate(dateStr) {
  if (!dateStr) return new Date();

  // 이미 Date 객체인 경우
  if (dateStr instanceof Date) return dateStr;

  // YYYY-MM-DD 형식
  if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
    return new Date(dateStr);
  }

  // YYYY/MM/DD 형식
  if (dateStr.match(/^\d{4}\/\d{2}\/\d{2}$/)) {
    return new Date(dateStr.replace(/\//g, '-'));
  }

  // MM/DD/YYYY 형식
  if (dateStr.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
    const parts = dateStr.split('/');
    return new Date(parts[2] + '-' + parts[0] + '-' + parts[1]);
  }

  // 기본적으로 Date 생성 시도
  const parsed = new Date(dateStr);
  return isNaN(parsed.getTime()) ? new Date() : parsed;
}

/**
 * 플랫폼별 통계
 */
function showPlatformStatistics() {
  const ui = SpreadsheetApp.getUi();

  const yearResponse = ui.prompt('플랫폼 통계', '연도를 입력하세요 (예: 2024):', ui.ButtonSet.OK_CANCEL);
  if (yearResponse.getSelectedButton() !== ui.Button.OK) return;

  const monthResponse = ui.prompt('플랫폼 통계', '월을 입력하세요 (예: 11):', ui.ButtonSet.OK_CANCEL);
  if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

  const year = parseInt(yearResponse.getResponseText());
  const month = parseInt(monthResponse.getResponseText());

  if (!year || !month || month < 1 || month > 12) {
    ui.alert('❌ 올바른 연도와 월을 입력해주세요.');
    return;
  }

  const stats = calculatePlatformStatistics(year, month);
  displayPlatformStatistics(stats, year, month);
}

/**
 * 플랫폼 통계 계산
 */
function calculatePlatformStatistics(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const platformSheet = ss.getSheetByName('플랫폼정산통합');

  const stats = {
    byPlatform: { '배민': 0, '쿠팡': 0, '요기요': 0, '땡겨요': 0 },
    byEmployee: {},
    total: 0
  };

  if (!platformSheet) {
    return stats;
  }

  const data = platformSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    const date = new Date(data[i][0]);
    const platform = data[i][1];
    const assignee = data[i][5];

    if (date.getFullYear() === year && date.getMonth() + 1 === month) {
      // 플랫폼별 집계
      if (stats.byPlatform.hasOwnProperty(platform)) {
        stats.byPlatform[platform]++;
      }

      // 직원별 집계
      if (assignee) {
        if (!stats.byEmployee[assignee]) {
          stats.byEmployee[assignee] = { '배민': 0, '쿠팡': 0, '요기요': 0, '땡겨요': 0, total: 0 };
        }
        if (stats.byEmployee[assignee].hasOwnProperty(platform)) {
          stats.byEmployee[assignee][platform]++;
        }
        stats.byEmployee[assignee].total++;
      }

      stats.total++;
    }
  }

  return stats;
}

/**
 * 플랫폼 통계 표시
 */
function displayPlatformStatistics(stats, year, month) {
  let message = '📊 플랫폼별 정산 통계 (' + year + '년 ' + month + '월)\n\n';

  message += '=== 플랫폼별 건수 ===\n';
  message += '배민: ' + stats.byPlatform['배민'] + '건\n';
  message += '쿠팡: ' + stats.byPlatform['쿠팡'] + '건\n';
  message += '요기요: ' + stats.byPlatform['요기요'] + '건\n';
  message += '땡겨요: ' + stats.byPlatform['땡겨요'] + '건\n';
  message += '총 건수: ' + stats.total + '건\n\n';

  message += '=== 직원별 처리 건수 ===\n';
  for (let employee in stats.byEmployee) {
    const empStats = stats.byEmployee[employee];
    message += employee + ': ' + empStats.total + '건\n';
    message += '  (배민:' + empStats['배민'] + ', 쿠팡:' + empStats['쿠팡'] +
               ', 요기요:' + empStats['요기요'] + ', 땡겨요:' + empStats['땡겨요'] + ')\n';
  }

  SpreadsheetApp.getUi().alert(message);
}

/**
 * 직원별 월간 플랫폼 건수 가져오기 (다른 파일에서 사용)
 */
function getEmployeePlatformCounts(employeeName, year, month) {
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

    const date = new Date(data[i][0]);
    const platform = data[i][1];
    const assignee = data[i][5];

    if (date.getFullYear() === year &&
        date.getMonth() + 1 === month &&
        assignee === employeeName) {
      if (counts.hasOwnProperty(platform)) {
        counts[platform]++;
      }
    }
  }

  return counts;
}
