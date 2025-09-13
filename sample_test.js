/**
 * Google Sheets 시작 시 실행되어 커스텀 메뉴를 생성하는 함수
 * 도구 메뉴에 "중복값 합침" 기능을 추가
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('중복제거')
    .addItem('중복값 합침', 'mergeDuplicateValues')
    .addToUi();
}

/**
 * A열 기준으로 중복된 행들의 숫자 데이터를 합치고 중복 행을 삭제하는 메인 함수
 * - 6행에서 "예약리스트" 헤더 위치를 동적으로 찾아 기준 범위 설정
 * - D열부터 예약리스트 전 열까지의 숫자 데이터를 합계 처리
 * - A열에 동일한 값을 가진 행들 중 첫 번째 행에 데이터를 합치고 나머지 행 삭제
 */
function mergeDuplicateValues() {
  try {
    // 스프레드시트 및 기본 정보 가져오기
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // 상수 정의
    const HEADER_ROW = 6; // 헤더가 있는 행 번호
    const DATA_START_ROW = 7; // 데이터 시작 행 번호
    const START_COL = 4; // D열 (데이터 범위 시작)
    const RESERVATION_HEADER = "예약리스트"; // 찾을 헤더명
    
    // 데이터 유효성 검사
    if (lastRow < DATA_START_ROW) {
      SpreadsheetApp.getUi().alert('데이터가 충분하지 않습니다. 최소 7행 이상의 데이터가 필요합니다.');
      return;
    }
    
    // 6행에서 "예약리스트" 헤더 위치 찾기
    const headerRow = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
    let reservationListCol = -1;
    
    for (let i = 0; i < headerRow.length; i++) {
      if (headerRow[i] === RESERVATION_HEADER) {
        reservationListCol = i + 1; // 1-based index로 변환
        break;
      }
    }
    
    // 예약리스트 헤더 존재 여부 확인
    if (reservationListCol === -1) {
      SpreadsheetApp.getUi().alert(`"${RESERVATION_HEADER}" 헤더를 ${HEADER_ROW}행에서 찾을 수 없습니다.\n헤더가 정확히 입력되어 있는지 확인해주세요.`);
      return;
    }
    
    // 기준 범위 계산: D열부터 예약리스트 바로 전 열까지
    const endCol = reservationListCol - 1;
    
    if (START_COL >= endCol) {
      SpreadsheetApp.getUi().alert('기준 범위가 올바르지 않습니다.\nD열부터 예약리스트 전 열까지의 범위를 확인해주세요.');
      return;
    }
    
    // 데이터 범위 가져오기
    const aColumnData = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1).getValues();
    const dataRange = sheet.getRange(DATA_START_ROW, START_COL, lastRow - HEADER_ROW, endCol - START_COL + 1);
    const numericData = dataRange.getValues();
    
    // 중복값 처리를 위한 변수 초기화
    const processedValues = {};
    const rowsToDelete = [];
    
    // A열의 각 값에 대해 중복 처리 수행
    for (let i = 0; i < aColumnData.length; i++) {
      const currentValue = aColumnData[i][0];
      
      // 빈 값이나 null 값은 건너뛰기
      if (currentValue === '' || currentValue === null) {
        continue;
      }
      
      // 이미 처리된 값이면 건너뛰기
      if (processedValues[currentValue]) {
        continue;
      }
      
      // 동일한 값을 가진 모든 행의 인덱스 찾기
      const duplicateIndices = [];
      for (let j = 0; j < aColumnData.length; j++) {
        if (aColumnData[j][0] === currentValue) {
          duplicateIndices.push(j);
        }
      }
      
      // 중복이 있는 경우만 처리 (2개 이상의 행이 같은 값을 가질 때)
      if (duplicateIndices.length > 1) {
        // 첫 번째 행을 기준으로 값들을 합치기
        const firstRowIndex = duplicateIndices[0];
        const summedRow = [];
        
        // 각 열에 대해 숫자값 합계 계산
        for (let col = 0; col < numericData[firstRowIndex].length; col++) {
          let sum = 0;
          let hasValidValue = false;
          
          // 중복된 모든 행의 해당 열 값을 합계
          for (let k = 0; k < duplicateIndices.length; k++) {
            const rowIndex = duplicateIndices[k];
            const cellValue = numericData[rowIndex][col];
            
            // 숫자로 변환 가능한 값만 더하기 (빈 문자열, null, NaN 제외)
            if (cellValue !== '' && cellValue !== null && !isNaN(cellValue)) {
              const numValue = parseFloat(cellValue);
              if (numValue !== 0) {
                hasValidValue = true;
              }
              sum += numValue;
            }
          }
          
          // 유효한 값이 있거나 합계가 0이 아닌 경우에만 값을 설정, 그렇지 않으면 빈 셀로 유지
          summedRow.push(hasValidValue || sum !== 0 ? sum : '');
        }
        
        // 합친 값을 첫 번째 행에 설정
        sheet.getRange(firstRowIndex + DATA_START_ROW, START_COL, 1, summedRow.length).setValues([summedRow]);
        
        // 첫 번째 행 이후의 중복 행들을 삭제 목록에 추가
        for (let k = 1; k < duplicateIndices.length; k++) {
          rowsToDelete.push(duplicateIndices[k] + DATA_START_ROW); // 실제 시트 행 번호로 변환
        }
      }
      
      // 처리된 값으로 표시
      processedValues[currentValue] = true;
    }
    
    // 행 삭제 (뒤에서부터 삭제하여 인덱스 변경 방지)
    rowsToDelete.sort((a, b) => b - a);
    
    for (let i = 0; i < rowsToDelete.length; i++) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    
    // 작업 완료 메시지 생성
    const message = `중복값 합치기가 완료되었습니다.
- 기준 범위: D열 ~ ${String.fromCharCode(64 + endCol)}열
- 삭제된 행: ${rowsToDelete.length}개
- 예약리스트 헤더 위치: ${String.fromCharCode(64 + reservationListCol)}${HEADER_ROW}`;
    
    SpreadsheetApp.getUi().alert(message);
    
  } catch (error) {
    // 예외 처리: 오류 발생 시 사용자에게 알림
    SpreadsheetApp.getUi().alert(`오류가 발생했습니다: ${error.message}`);
    console.error('mergeDuplicateValues 함수 실행 중 오류:', error);
  }
}