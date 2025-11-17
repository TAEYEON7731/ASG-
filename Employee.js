/**
 * 직원 관리 기능
 */

/**
 * 직원 등록
 */
function addEmployee(name, department, position, phone, email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('직원목록');

  if (!sheet) {
    throw new Error('직원목록 시트가 없습니다. 초기 설정을 먼저 실행해주세요.');
  }

  const employeeId = generateEmployeeId();
  const hireDate = new Date();
  const status = '재직';
  const registeredDate = new Date();

  sheet.appendRow([
    employeeId,
    name,
    department,
    position,
    hireDate,
    phone,
    email,
    status,
    registeredDate
  ]);

  return employeeId;
}

/**
 * 직원 조회 (사번 또는 이름으로)
 */
function searchEmployee(keyword) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('직원목록');
  const data = sheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0].toString().includes(keyword) || row[1].includes(keyword)) {
      results.push({
        employeeId: row[0],
        name: row[1],
        department: row[2],
        position: row[3],
        hireDate: row[4],
        phone: row[5],
        email: row[6],
        status: row[7]
      });
    }
  }

  return results;
}

/**
 * 직원 정보 수정
 */
function updateEmployee(employeeId, field, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('직원목록');
  const data = sheet.getDataRange().getValues();

  const fieldMap = {
    '이름': 1,
    '부서': 2,
    '직급': 3,
    '연락처': 5,
    '이메일': 6,
    '상태': 7
  };

  const colIndex = fieldMap[field];
  if (!colIndex) {
    throw new Error('잘못된 필드명입니다.');
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === employeeId) {
      sheet.getRange(i + 1, colIndex + 1).setValue(value);
      return true;
    }
  }

  return false;
}

/**
 * 직원 퇴사 처리
 */
function retireEmployee(employeeId) {
  return updateEmployee(employeeId, '상태', '퇴사');
}

/**
 * 전체 직원 목록 조회
 */
function getAllEmployees() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('직원목록');
  const data = sheet.getDataRange().getValues();
  const employees = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    employees.push({
      employeeId: row[0],
      name: row[1],
      department: row[2],
      position: row[3],
      hireDate: row[4],
      phone: row[5],
      email: row[6],
      status: row[7]
    });
  }

  return employees;
}

/**
 * 부서별 통계
 */
function getStatisticsByDepartment() {
  const employees = getAllEmployees();
  const stats = {};

  employees.forEach(emp => {
    if (emp.status === '재직') {
      if (!stats[emp.department]) {
        stats[emp.department] = 0;
      }
      stats[emp.department]++;
    }
  });

  return stats;
}
