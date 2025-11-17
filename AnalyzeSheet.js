/**
 * 기존 스프레드시트 분석 도구
 */

function analyzeExistingSheet() {
  const targetSpreadsheetId = '1C2Rr4oK3y6VKXTv7_R7ciJ6ihcbum_DWBIolUwgJXoQ';
  const ss = SpreadsheetApp.openById(targetSpreadsheetId);
  const sheets = ss.getSheets();

  let analysis = '=== 스프레드시트 분석 결과 ===\n\n';
  analysis += '스프레드시트 이름: ' + ss.getName() + '\n';
  analysis += '총 시트 개수: ' + sheets.length + '\n\n';

  sheets.forEach(function(sheet, index) {
    analysis += '--- 시트 ' + (index + 1) + ' ---\n';
    analysis += '이름: ' + sheet.getName() + '\n';
    analysis += '행 수: ' + sheet.getLastRow() + '\n';
    analysis += '열 수: ' + sheet.getLastColumn() + '\n';

    if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      analysis += '헤더: ' + headers.join(', ') + '\n';

      // 샘플 데이터 (첫 3행)
      if (sheet.getLastRow() > 1) {
        const sampleRows = Math.min(3, sheet.getLastRow() - 1);
        analysis += '샘플 데이터 (' + sampleRows + '행):\n';
        const sampleData = sheet.getRange(2, 1, sampleRows, sheet.getLastColumn()).getValues();
        sampleData.forEach(function(row, rowIndex) {
          analysis += '  행' + (rowIndex + 2) + ': ' + row.join(' | ') + '\n';
        });
      }
    }
    analysis += '\n';
  });

  Logger.log(analysis);
  return analysis;
}

/**
 * 기존 데이터 복사
 */
function copyExistingData() {
  const sourceSpreadsheetId = '1C2Rr4oK3y6VKXTv7_R7ciJ6ihcbum_DWBIolUwgJXoQ';
  const sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);
  const targetSS = SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheets = sourceSS.getSheets();

  sourceSheets.forEach(function(sourceSheet) {
    const sheetName = sourceSheet.getName();

    // 기존 시트 삭제 (있다면)
    let targetSheet = targetSS.getSheetByName(sheetName);
    if (targetSheet) {
      targetSS.deleteSheet(targetSheet);
    }

    // 시트 복사
    targetSheet = sourceSheet.copyTo(targetSS);
    targetSheet.setName(sheetName);

    Logger.log('복사 완료: ' + sheetName);
  });

  SpreadsheetApp.getUi().alert('✅ 모든 시트가 복사되었습니다!');
}
