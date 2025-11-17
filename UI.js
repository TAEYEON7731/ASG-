/**
 * 사용자 인터페이스 관련 함수
 */

/**
 * 직원 등록 다이얼로그 표시
 */
function showAddEmployeeDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Noto Sans KR', sans-serif;
            padding: 20px;
          }
          .form-group {
            margin-bottom: 15px;
          }
          label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
            color: #333;
          }
          input {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-sizing: border-box;
          }
          button {
            background-color: #4285f4;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            width: 100%;
            margin-top: 10px;
          }
          button:hover {
            background-color: #357ae8;
          }
          .message {
            margin-top: 10px;
            padding: 10px;
            border-radius: 4px;
            display: none;
          }
          .success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
          }
          .error {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
          }
        </style>
      </head>
      <body>
        <h2>직원 등록</h2>
        <form id="employeeForm">
          <div class="form-group">
            <label for="name">이름 *</label>
            <input type="text" id="name" required>
          </div>

          <div class="form-group">
            <label for="department">부서 *</label>
            <input type="text" id="department" required>
          </div>

          <div class="form-group">
            <label for="position">직급 *</label>
            <input type="text" id="position" required>
          </div>

          <div class="form-group">
            <label for="phone">연락처</label>
            <input type="tel" id="phone" placeholder="010-0000-0000">
          </div>

          <div class="form-group">
            <label for="email">이메일</label>
            <input type="email" id="email" placeholder="example@company.com">
          </div>

          <button type="submit">등록</button>
        </form>

        <div id="message" class="message"></div>

        <script>
          document.getElementById('employeeForm').addEventListener('submit', function(e) {
            e.preventDefault();

            const name = document.getElementById('name').value;
            const department = document.getElementById('department').value;
            const position = document.getElementById('position').value;
            const phone = document.getElementById('phone').value;
            const email = document.getElementById('email').value;

            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .addEmployee(name, department, position, phone, email);
          });

          function onSuccess(employeeId) {
            const messageDiv = document.getElementById('message');
            messageDiv.className = 'message success';
            messageDiv.style.display = 'block';
            messageDiv.textContent = '직원이 등록되었습니다! 사번: ' + employeeId;

            document.getElementById('employeeForm').reset();

            setTimeout(() => {
              messageDiv.style.display = 'none';
            }, 3000);
          }

          function onFailure(error) {
            const messageDiv = document.getElementById('message');
            messageDiv.className = 'message error';
            messageDiv.style.display = 'block';
            messageDiv.textContent = '오류: ' + error.message;
          }
        </script>
      </body>
    </html>
  `)
    .setWidth(400)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, '직원 등록');
}

/**
 * 직원 조회 다이얼로그 표시
 */
function showSearchEmployeeDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Noto Sans KR', sans-serif;
            padding: 20px;
          }
          .search-box {
            margin-bottom: 20px;
          }
          input[type="text"] {
            width: 70%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
          }
          button {
            background-color: #4285f4;
            color: white;
            padding: 8px 16px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
          }
          button:hover {
            background-color: #357ae8;
          }
          .results {
            margin-top: 20px;
          }
          .employee-card {
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 15px;
            margin-bottom: 10px;
            background-color: #f9f9f9;
          }
          .employee-card h3 {
            margin-top: 0;
            color: #4285f4;
          }
          .employee-info {
            display: grid;
            grid-template-columns: 100px 1fr;
            gap: 5px;
          }
          .label {
            font-weight: bold;
            color: #666;
          }
        </style>
      </head>
      <body>
        <h2>직원 조회</h2>
        <div class="search-box">
          <input type="text" id="keyword" placeholder="사번 또는 이름을 입력하세요">
          <button onclick="search()">검색</button>
        </div>

        <div id="results" class="results"></div>

        <script>
          function search() {
            const keyword = document.getElementById('keyword').value;
            if (!keyword) {
              alert('검색어를 입력해주세요.');
              return;
            }

            google.script.run
              .withSuccessHandler(displayResults)
              .searchEmployee(keyword);
          }

          function displayResults(employees) {
            const resultsDiv = document.getElementById('results');

            if (employees.length === 0) {
              resultsDiv.innerHTML = '<p>검색 결과가 없습니다.</p>';
              return;
            }

            let html = '<h3>검색 결과 (' + employees.length + '건)</h3>';

            employees.forEach(emp => {
              html += '<div class="employee-card">';
              html += '<h3>' + emp.name + ' (' + emp.employeeId + ')</h3>';
              html += '<div class="employee-info">';
              html += '<div class="label">부서:</div><div>' + emp.department + '</div>';
              html += '<div class="label">직급:</div><div>' + emp.position + '</div>';
              html += '<div class="label">입사일:</div><div>' + new Date(emp.hireDate).toLocaleDateString('ko-KR') + '</div>';
              html += '<div class="label">연락처:</div><div>' + emp.phone + '</div>';
              html += '<div class="label">이메일:</div><div>' + emp.email + '</div>';
              html += '<div class="label">상태:</div><div>' + emp.status + '</div>';
              html += '</div>';
              html += '</div>';
            });

            resultsDiv.innerHTML = html;
          }

          document.getElementById('keyword').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
              search();
            }
          });
        </script>
      </body>
    </html>
  `)
    .setWidth(500)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, '직원 조회');
}

/**
 * 통계 다이얼로그 표시
 */
function showStatistics() {
  const stats = getStatisticsByDepartment();
  let message = '📊 부서별 재직 인원 현황\n\n';

  let total = 0;
  for (const dept in stats) {
    message += dept + ': ' + stats[dept] + '명\n';
    total += stats[dept];
  }

  message += '\n총 인원: ' + total + '명';

  SpreadsheetApp.getUi().alert(message);
}
