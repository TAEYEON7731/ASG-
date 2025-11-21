/**
 * ASG 직원 관리 시스템 - 시트 초기화 Part 2
 * 출퇴근기록, 급여계산, 플랫폼인센티브, 연차관리, 설정
 */

/**
 * 3. 출퇴근기록
 */
function create_AttendanceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('출퇴근기록');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('출퇴근기록');
  sheet.setTabColor('#fbbc04');

  // 헤더
  const headers = [
    '날짜', '요일', '이름', '부서',
    '출근시간', '퇴근시간', '근무시간',
    '연장근무', '비고'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#fbbc04')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 40);

  // 샘플 데이터 (이번 주)
  const today = new Date();
  const sampleData = [];

  for (let i = 0; i < 5; i++) {
    const date = new Date(today);
    date.setDate(date.getDate() - (4 - i));

    sampleData.push([
      date,
      '=TEXT(A' + (i + 2) + ',"ddd")',  // 요일
      '홍길동',
      'TM팀',
      '09:00',
      '18:00',
      '=IF(AND(E' + (i + 2) + '<>"",F' + (i + 2) + '<>""), (F' + (i + 2) + '-E' + (i + 2) + ')*24, "")',  // 근무시간 자동계산
      '',
      ''
    ]);
  }

  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

  // 서식 설정
  sheet.getRange(2, 1, 100, 1).setNumberFormat('yyyy-mm-dd');  // 날짜
  sheet.getRange(2, 5, 100, 2).setNumberFormat('hh:mm');        // 출퇴근시간
  sheet.getRange(2, 7, 100, 2).setNumberFormat('0.0"시간"');    // 근무시간, 연장근무

  // 데이터 영역 테두리
  sheet.getRange(2, 1, 5, headers.length)
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');

  // 조건부 서식 (8시간 이상 근무시 파란색)
  const workHoursRange = sheet.getRange('G2:G100');
  let rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(8)
    .setBackground('#d4edda')
    .setRanges([workHoursRange])
    .build();

  sheet.setConditionalFormatRules([rule]);

  // 열 너비
  sheet.setColumnWidth(1, 120);  // 날짜
  sheet.setColumnWidth(2, 60);   // 요일
  sheet.setColumnWidth(3, 100);  // 이름
  sheet.setColumnWidth(4, 100);  // 부서
  sheet.setColumnWidth(5, 100);  // 출근시간
  sheet.setColumnWidth(6, 100);  // 퇴근시간
  sheet.setColumnWidth(7, 100);  // 근무시간
  sheet.setColumnWidth(8, 100);  // 연장근무
  sheet.setColumnWidth(9, 200);  // 비고
}

/**
 * 4. 급여계산
 */
function create_SalarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('급여계산');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('급여계산');
  sheet.setTabColor('#ea4335');

  // 상단 정보
  sheet.getRange('A1').setValue('기준 년월:');
  sheet.getRange('B1').setValue('2024-11');
  sheet.getRange('B1').setNumberFormat('yyyy-mm');
  sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#fff3cd');

  // 헤더
  const headers = [
    '이름', '부서', '급여형태', '시급',
    '총근무시간', '기본급',
    '배민', '쿠팡이츠', '요기요', '땡겨요',
    '총인센티브', '총급여', '비고'
  ];

  sheet.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#ea4335')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(2);
  sheet.setRowHeight(2, 40);

  // 샘플 데이터
  const sampleData = [
    [
      '홍길동', 'TM팀', '시급제', 13000,
      '=SUMIF(출퇴근기록!C:C,A3,출퇴근기록!G:G)',  // 총근무시간
      '=E3*D3',  // 기본급
      50, 30, 20, 10,  // 플랫폼 건수
      '=G3*설정!B3+H3*설정!B4+I3*설정!B5+J3*설정!B6',  // 총인센티브
      '=F3+K3',  // 총급여
      ''
    ],
    [
      '김철수', 'TM팀', '시급제', 13000,
      '=SUMIF(출퇴근기록!C:C,A4,출퇴근기록!G:G)',
      '=E4*D4',
      40, 25, 15, 5,
      '=G4*설정!B3+H4*설정!B4+I4*설정!B5+J4*설정!B6',
      '=F4+K4',
      ''
    ]
  ];

  sheet.getRange(3, 1, sampleData.length, headers.length).setValues(sampleData);

  // 서식 설정
  sheet.getRange(3, 4, 100, 1).setNumberFormat('#,##0"원"');     // 시급
  sheet.getRange(3, 5, 100, 1).setNumberFormat('0.0"시간"');     // 총근무시간
  sheet.getRange(3, 6, 100, 1).setNumberFormat('#,##0"원"');     // 기본급
  sheet.getRange(3, 7, 100, 4).setNumberFormat('0"건"');         // 플랫폼 건수
  sheet.getRange(3, 11, 100, 2).setNumberFormat('#,##0"원"');    // 총인센티브, 총급여

  // 데이터 영역 테두리
  sheet.getRange(3, 1, sampleData.length, headers.length)
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');

  // 총급여 열 강조
  sheet.getRange(2, 12, 100, 1).setBackground('#fff3cd');

  // 열 너비
  sheet.setColumnWidth(1, 100);  // 이름
  sheet.setColumnWidth(2, 100);  // 부서
  sheet.setColumnWidth(3, 100);  // 급여형태
  sheet.setColumnWidth(4, 100);  // 시급
  sheet.setColumnWidth(5, 110);  // 총근무시간
  sheet.setColumnWidth(6, 120);  // 기본급
  sheet.setColumnWidth(7, 80);   // 배민
  sheet.setColumnWidth(8, 90);   // 쿠팡이츠
  sheet.setColumnWidth(9, 80);   // 요기요
  sheet.setColumnWidth(10, 80);  // 땡겨요
  sheet.setColumnWidth(11, 120); // 총인센티브
  sheet.setColumnWidth(12, 130); // 총급여
  sheet.setColumnWidth(13, 150); // 비고
}

/**
 * 5. 플랫폼인센티브
 */
function create_PlatformIncentiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('플랫폼인센티브');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('플랫폼인센티브');
  sheet.setTabColor('#9c27b0');

  // 헤더
  const headers = [
    '날짜', '플랫폼', '상호명', '사업자번호',
    '담당자', '금액', '비고'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#9c27b0')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 40);

  // 샘플 데이터
  const sampleData = [
    [new Date(), '배민', '맛있는식당', '123-45-67890', '홍길동', 15000, ''],
    [new Date(), '쿠팡이츠', '행복카페', '234-56-78901', '김철수', 12000, ''],
    [new Date(), '요기요', '든든한집', '345-67-89012', '홍길동', 18000, ''],
    [new Date(), '땡겨요', '신선마트', '456-78-90123', '이영희', 10000, '']
  ];

  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

  // 서식 설정
  sheet.getRange(2, 1, 100, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 6, 100, 1).setNumberFormat('#,##0"원"');

  // 데이터 영역 테두리
  sheet.getRange(2, 1, sampleData.length, headers.length)
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');

  // 플랫폼별 색상 구분 (조건부 서식)
  const platformRange = sheet.getRange('B2:B1000');

  const rules = [];
  const platforms = [
    { name: '배민', color: '#d4edda' },
    { name: '쿠팡이츠', color: '#cfe2ff' },
    { name: '요기요', color: '#fff3cd' },
    { name: '땡겨요', color: '#f8d7da' }
  ];

  platforms.forEach(platform => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(platform.name)
        .setBackground(platform.color)
        .setRanges([platformRange])
        .build()
    );
  });

  sheet.setConditionalFormatRules(rules);

  // 데이터 검증 (플랫폼)
  const platformRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['배민', '쿠팡이츠', '요기요', '땡겨요'], true)
    .build();
  sheet.getRange('B2:B1000').setDataValidation(platformRule);

  // 열 너비
  sheet.setColumnWidth(1, 120);  // 날짜
  sheet.setColumnWidth(2, 100);  // 플랫폼
  sheet.setColumnWidth(3, 150);  // 상호명
  sheet.setColumnWidth(4, 130);  // 사업자번호
  sheet.setColumnWidth(5, 100);  // 담당자
  sheet.setColumnWidth(6, 120);  // 금액
  sheet.setColumnWidth(7, 200);  // 비고
}

/**
 * 6. 연차관리
 */
function create_AnnualLeaveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('연차관리');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('연차관리');
  sheet.setTabColor('#00bcd4');

  // 헤더
  const headers = [
    '이름', '입사일', '근속연수',
    '연차발생일수', '사용일수', '잔여일수', '비고'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#00bcd4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 40);

  // 샘플 데이터
  const sampleData = [
    [
      '홍길동',
      new Date(2022, 0, 1),
      '=DATEDIF(B2,TODAY(),"Y")',  // 근속연수
      '=15+IF(C2>=3,C2-1,0)',       // 연차발생일수 (기본15일+근속가산)
      5,                             // 사용일수
      '=D2-E2',                      // 잔여일수
      ''
    ],
    [
      '김철수',
      new Date(2023, 5, 1),
      '=DATEDIF(B3,TODAY(),"Y")',
      '=15+IF(C3>=3,C3-1,0)',
      2,
      '=D3-E3',
      ''
    ]
  ];

  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

  // 서식 설정
  sheet.getRange(2, 2, 100, 1).setNumberFormat('yyyy-mm-dd');  // 입사일
  sheet.getRange(2, 3, 100, 1).setNumberFormat('0"년"');       // 근속연수
  sheet.getRange(2, 4, 100, 3).setNumberFormat('0"일"');       // 연차 관련

  // 데이터 영역 테두리
  sheet.getRange(2, 1, sampleData.length, headers.length)
    .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');

  // 잔여일수 조건부 서식 (5일 이하면 빨간색)
  const remainRange = sheet.getRange('F2:F100');
  let rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(5)
    .setBackground('#f8d7da')
    .setFontColor('#721c24')
    .setRanges([remainRange])
    .build();

  sheet.setConditionalFormatRules([rule]);

  // 열 너비
  sheet.setColumnWidth(1, 100);  // 이름
  sheet.setColumnWidth(2, 120);  // 입사일
  sheet.setColumnWidth(3, 100);  // 근속연수
  sheet.setColumnWidth(4, 120);  // 연차발생일수
  sheet.setColumnWidth(5, 100);  // 사용일수
  sheet.setColumnWidth(6, 100);  // 잔여일수
  sheet.setColumnWidth(7, 200);  // 비고
}

/**
 * 7. 설정
 */
function create_SettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('⚙️ 설정');

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet('⚙️ 설정');
  sheet.setTabColor('#607d8b');

  // 제목
  sheet.getRange('A1').setValue('시스템 설정').setFontSize(16).setFontWeight('bold');
  sheet.setRowHeight(1, 40);

  // 플랫폼 인센티브 단가
  sheet.getRange('A3').setValue('플랫폼별 인센티브 단가').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A4:B4').setValues([['플랫폼', '건당 금액']]).setFontWeight('bold').setBackground('#f8f9fa');

  const incentiveData = [
    ['배민', 1000],
    ['쿠팡이츠', 1000],
    ['요기요', 1000],
    ['땡겨요', 1000]
  ];

  sheet.getRange(5, 1, incentiveData.length, 2).setValues(incentiveData);
  sheet.getRange(5, 2, incentiveData.length, 1).setNumberFormat('#,##0"원"');

  // 기본 설정
  sheet.getRange('A10').setValue('기본 설정').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A11:B11').setValues([['항목', '값']]).setFontWeight('bold').setBackground('#f8f9fa');

  const basicSettings = [
    ['기본 시급', 13000],
    ['정규 근무시간', 8],
    ['주 근무일', 5]
  ];

  sheet.getRange(12, 1, basicSettings.length, 2).setValues(basicSettings);
  sheet.getRange(12, 2, basicSettings.length, 1).setNumberFormat('#,##0');

  // 열 너비
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 120);
}
