/**
 * 스프레드시트 노란색 칠하기 및 출고완료 자동화 시스템
 *
 * 주요 기능:
 * 1. C열 체크박스 체크 시 해당 행의 D7:AH198 숫자 셀들을 노란색으로 칠하기
 * 2. 노란색 셀의 행(A열) = BB열, 열(6행) = BC열 일치 시 BE열에 "출고완료" 입력
 * 3. 수동 노란색 칠하기도 같은 로직으로 동작
 */

// 상수 정의
const CONFIG = {
  CHECK_COLUMN: 3,           // C열 (체크박스)
  DATA_START_COLUMN: 4,      // D열 시작
  DATA_END_COLUMN: 34,       // AH열 (4 + 30 = 34)
  DATA_START_ROW: 7,         // 7행부터
  DATA_END_ROW: 198,         // 198행까지
  HEADER_ROW: 6,             // 헤더 행
  LOOKUP_START_ROW: 7,       // BB/BC/BD 데이터 시작 행
  AK_COLUMN: 37,             // AK열 (메모 트리거)
  BB_COLUMN: 54,             // BB열 (주문자 데이터)
  BC_COLUMN: 55,             // BC열 (제품 데이터)
  BD_COLUMN: 56,             // BD열 (수량 데이터)
  BE_COLUMN: 57,             // BE열 (결과)
  YELLOW_COLOR: '#FFFF00'    // 노란색 코드
};

/**
 * 스프레드시트 시작 시 실행되는 함수
 * 커스텀 메뉴 생성 및 권한 설정
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('노란색 칠하기')
    .addItem('수동 출고완료 확인', 'checkAllYellowCells')
    .addItem('노란색 모두 제거', 'clearAllYellowHighlights')
    .addItem('트리거 생성', 'createTriggers')
    .addItem('트리거 삭제', 'deleteAllTriggers')
    .addItem('트리거 상태 확인', 'showTriggerStatus')
    .addToUi();
}

/**
 * 셀 편집 시 실행되는 트리거 함수 (경량화)
 * C열 체크박스 체크 시 노란색 칠하기 수행
 * AK열 메모 입력 시 해당 행의 노란색 셀 감지하여 출고완료 처리
 */
function onEdit(e) {
  // 빠른 처리를 위한 간단한 중복 방지
  const lock = LockService.getScriptLock();
  try {
    // 1초만 대기 - 실패시 즉시 건너뜀
    lock.waitLock(1000);
  } catch (lockError) {
    return; // 로그 없이 조용히 건너뜀
  }

  try {
    if (!e) return;

    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const row = range.getRow();
    const column = range.getColumn();

    // C열 체크박스 체크 감지
    if (column === CONFIG.CHECK_COLUMN &&
        row >= CONFIG.DATA_START_ROW &&
        row <= CONFIG.DATA_END_ROW) {

      const isChecked = range.getValue();

      if (isChecked) {
        // 체크된 경우 즉시 노란색 칠하기 (지연 없음)
        highlightYellowCellsFast(sheet, row);
      } else {
        // 체크 해제된 경우 노란색 제거
        clearYellowHighlightFast(sheet, row);
      }
    }

    // AK열 메모 입력 감지 (수동 노란색 칠하기 트리거)
    if (column === CONFIG.AK_COLUMN &&
        row >= CONFIG.DATA_START_ROW &&
        row <= CONFIG.DATA_END_ROW) {

      // 해당 행의 노란색 셀 검사 및 출고완료 처리 (즉시 실행)
      checkRowForYellowCellsFast(sheet, row);
    }

  } catch (error) {
    // 에러 로그도 최소화
    console.error('onEdit 오류:', error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 특정 행의 숫자 셀들을 노란색으로 칠하기 (안정성 강화 버전)
 */
function highlightYellowCellsStable(sheet, row) {
  try {
    console.log(`🎨 안정적 노란색 칠하기 시작: 행 ${row}`);

    // 처리 전 상태 확인
    Utilities.sleep(100);

    // D열부터 AH열까지의 범위
    const range = sheet.getRange(row, CONFIG.DATA_START_COLUMN, 1,
                                CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1);
    const values = range.getValues()[0];

    let processedCells = 0;
    const cellsToHighlight = [];

    // 먼저 노란색 칠할 셀들을 식별
    for (let i = 0; i < values.length; i++) {
      const cellValue = values[i];
      if (typeof cellValue === 'number' && cellValue > 0) {
        cellsToHighlight.push({
          col: CONFIG.DATA_START_COLUMN + i,
          value: cellValue
        });
      }
    }

    console.log(`📊 노란색 칠할 셀 개수: ${cellsToHighlight.length}`);

    // 배치로 노란색 칠하기 (한 번에 처리하여 안정성 확보)
    if (cellsToHighlight.length > 0) {
      for (let cell of cellsToHighlight) {
        try {
          const cellRange = sheet.getRange(row, cell.col);
          cellRange.setBackground(CONFIG.YELLOW_COLOR);
          processedCells++;
          console.log(`✅ 노란색 칠하기: 행 ${row}, 열 ${cell.col}, 값: ${cell.value}`);

          // 각 셀 처리 후 짧은 지연 (안정성)
          Utilities.sleep(50);
        } catch (cellError) {
          console.error(`❌ 셀 칠하기 실패 (행${row}, 열${cell.col}):`, cellError);
        }
      }

      // 모든 칠하기 완료 후 처리 지연
      Utilities.sleep(300);

      console.log(`✅ 노란색 칠하기 완료: ${processedCells}/${cellsToHighlight.length}개 처리됨`);

      // 노란색 칠하기 후 출고완료 확인
      checkShipmentCompletion(sheet);
    } else {
      console.log(`⚠️ 행 ${row}에 칠할 숫자 셀이 없습니다.`);
    }

  } catch (error) {
    console.error('❌ 안정적 노란색 칠하기 오류:', error);
  }
}

/**
 * 특정 행의 숫자 셀들을 노란색으로 칠하기 (초고속 버전)
 */
function highlightYellowCellsFast(sheet, row) {
  try {
    const range = sheet.getRange(row, CONFIG.DATA_START_COLUMN, 1,
                                CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1);
    const values = range.getValues()[0];
    const backgrounds = [];

    // 배치 처리를 위한 배열 생성
    for (let i = 0; i < values.length; i++) {
      if (typeof values[i] === 'number' && values[i] > 0) {
        backgrounds.push(CONFIG.YELLOW_COLOR);
      } else {
        backgrounds.push(null);
      }
    }

    // 한 번에 배경색 설정 (가장 빠른 방법)
    range.setBackgrounds([backgrounds]);

    // 출고완료 확인 (전체 스캔 대신 해당 행만)
    checkRowForShipmentFast(sheet, row);

  } catch (error) {
    console.error('빠른 노란색 칠하기 실패:', error.message);
  }
}

/**
 * 특정 행의 노란색 하이라이트 제거 + 출고완료 삭제 (초고속 버전)
 */
function clearYellowHighlightFast(sheet, row) {
  try {
    // 1. 노란색 제거
    const range = sheet.getRange(row, CONFIG.DATA_START_COLUMN, 1,
                                CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1);
    range.setBackground(null);

    // 2. 해당 행의 출고완료도 제거
    clearShipmentForRow(sheet, row);

  } catch (error) {
    console.error('빠른 노란색 제거 실패:', error.message);
  }
}

/**
 * 특정 행에 해당하는 출고완료 상태 정확히 제거
 */
function clearShipmentForRow(sheet, targetRow) {
  try {
    console.log(`🗑️ 출고완료 삭제 시작: 행 ${targetRow}`);

    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.LOOKUP_START_ROW) return;

    // 해당 행의 데이터와 헤더 정보 가져오기
    const rowRange = sheet.getRange(targetRow, CONFIG.DATA_START_COLUMN, 1,
                                   CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1);
    const values = rowRange.getValues()[0];

    const headerData = sheet.getRange(CONFIG.HEADER_ROW, CONFIG.DATA_START_COLUMN, 1,
                                     CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1).getValues()[0];

    const orderName = sheet.getRange(targetRow, 1).getValue();

    // BB/BC/BD/BE 데이터 가져오기
    const lookupData = sheet.getRange(CONFIG.LOOKUP_START_ROW, CONFIG.BB_COLUMN,
                                     lastRow - CONFIG.LOOKUP_START_ROW + 1, 4).getValues();

    console.log(`🔍 삭제 대상 주문자: "${orderName}"`);

    let deletedCount = 0;

    // 해당 행의 숫자가 있는 셀들을 기준으로 정확한 매칭 삭제
    for (let i = 0; i < values.length; i++) {
      const cellValue = values[i];

      if (typeof cellValue === 'number' && cellValue > 0) {
        const productName = headerData[i];

        console.log(`📋 삭제할 항목: "${orderName}" + "${productName}" + ${cellValue}`);

        // BB/BC/BD와 정확한 매칭 찾기
        for (let k = 0; k < lookupData.length; k++) {
          const bbValue = lookupData[k][0]; // BB열 주문자
          const bcValue = lookupData[k][1]; // BC열 제품
          const bdValue = lookupData[k][2]; // BD열 수량

          if (bbValue === orderName && bcValue === productName && bdValue === cellValue) {
            const matchRow = CONFIG.LOOKUP_START_ROW + k;
            sheet.getRange(matchRow, CONFIG.BE_COLUMN).setValue('');
            deletedCount++;
            console.log(`✅ 출고완료 삭제: BE${matchRow} (${orderName}+${productName}+${cellValue})`);
          }
        }
      }
    }

    console.log(`📊 출고완료 삭제 완료: ${deletedCount}개 항목 삭제됨`);

  } catch (error) {
    console.error('❌ 출고완료 삭제 실패:', error);
  }
}

/**
 * 특정 행의 노란색 셀 검사 (초고속 버전)
 */
function checkRowForYellowCellsFast(sheet, targetRow) {
  try {
    const rowRange = sheet.getRange(targetRow, CONFIG.DATA_START_COLUMN, 1,
                                   CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1);
    const backgrounds = rowRange.getBackgrounds()[0];
    const values = rowRange.getValues()[0];

    // 헤더와 BB/BC/BD 데이터 한 번에 가져오기
    const headerData = sheet.getRange(CONFIG.HEADER_ROW, CONFIG.DATA_START_COLUMN, 1,
                                     CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1).getValues()[0];

    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.LOOKUP_START_ROW) return;

    const lookupData = sheet.getRange(CONFIG.LOOKUP_START_ROW, CONFIG.BB_COLUMN,
                                     lastRow - CONFIG.LOOKUP_START_ROW + 1, 4).getValues();

    const orderName = sheet.getRange(targetRow, 1).getValue();

    // 노란색 셀 찾기 및 매칭 (로그 최소화)
    for (let j = 0; j < backgrounds.length; j++) {
      if (isYellowColor(backgrounds[j])) {
        const productName = headerData[j];
        const yellowValue = values[j];

        // 빠른 매칭
        for (let k = 0; k < lookupData.length; k++) {
          if (lookupData[k][0] === orderName &&
              lookupData[k][1] === productName &&
              lookupData[k][2] === yellowValue) {

            const matchRow = CONFIG.LOOKUP_START_ROW + k;
            sheet.getRange(matchRow, CONFIG.BE_COLUMN).setValue('출고완료');
            return; // 첫 매칭에서 즉시 종료
          }
        }
      }
    }
  } catch (error) {
    console.error('빠른 행 검사 실패:', error.message);
  }
}

/**
 * 특정 행의 출고완료 확인 (초고속 버전)
 */
function checkRowForShipmentFast(sheet, targetRow) {
  // AK열 트리거와 동일한 로직 사용
  checkRowForYellowCellsFast(sheet, targetRow);
}

/**
 * 특정 행의 숫자 셀들을 노란색으로 칠하기 (기존 버전)
 */
function highlightYellowCells(sheet, row) {
  // 빠른 버전 호출
  highlightYellowCellsFast(sheet, row);
}

/**
 * 특정 행의 노란색 하이라이트 제거 (기존 버전)
 */
function clearYellowHighlight(sheet, row) {
  // 빠른 버전 호출
  clearYellowHighlightFast(sheet, row);
}

/**
 * 특정 행의 노란색 셀 검사 및 출고완료 처리
 * AK열 메모 입력 시 호출됨
 */
function checkRowForYellowCells(sheet, targetRow) {
  try {
    console.log(`🔍 행 ${targetRow}의 노란색 셀 검사 시작`);

    // 해당 행의 데이터 영역(D~AH) 배경색 가져오기
    const rowRange = sheet.getRange(targetRow, CONFIG.DATA_START_COLUMN,
                                   1, CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1);
    const backgrounds = rowRange.getBackgrounds()[0];
    const values = rowRange.getValues()[0];

    // 6행 헤더 데이터 가져오기
    const headerRange = sheet.getRange(CONFIG.HEADER_ROW, CONFIG.DATA_START_COLUMN,
                                      1, CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1);
    const headerData = headerRange.getValues()[0];

    // BB/BC/BD/BE 데이터 가져오기
    const lastRow = sheet.getLastRow();
    const lookupRange = sheet.getRange(CONFIG.LOOKUP_START_ROW, CONFIG.BB_COLUMN,
                                      lastRow - CONFIG.LOOKUP_START_ROW + 1, 4);
    const lookupData = lookupRange.getValues();

    let yellowCellFound = false;

    // 해당 행에서 노란색 셀 찾기
    for (let j = 0; j < backgrounds.length; j++) {
      if (isYellowColor(backgrounds[j])) {
        const yellowCol = CONFIG.DATA_START_COLUMN + j;
        const yellowValue = values[j];

        yellowCellFound = true;

        // 노란색 셀 정보 추출
        const orderName = sheet.getRange(targetRow, 1).getValue(); // A열 주문자
        const productName = headerData[j]; // 6행 제품명

        console.log(`🟡 행 ${targetRow}에서 노란색 셀 발견:`);
        console.log(`   위치: 열${yellowCol}`);
        console.log(`   주문자(A${targetRow}): "${orderName}"`);
        console.log(`   제품(${String.fromCharCode(67 + j)}6): "${productName}"`);
        console.log(`   수량: ${yellowValue}`);

        // BB/BC/BD 데이터와 매칭
        for (let k = 0; k < lookupData.length; k++) {
          const bbValue = lookupData[k][0]; // BB열 주문자
          const bcValue = lookupData[k][1]; // BC열 제품
          const bdValue = lookupData[k][2]; // BD열 수량

          if (bbValue === orderName && bcValue === productName && bdValue === yellowValue) {
            const matchRow = CONFIG.LOOKUP_START_ROW + k;
            console.log(`✅ 매칭 성공! BB/BC/BD 행 ${matchRow}에 출고완료 처리`);

            // BE열에 "출고완료" 입력
            const beCell = sheet.getRange(matchRow, CONFIG.BE_COLUMN);
            beCell.setValue('출고완료');
            console.log(`✅ BE${matchRow}에 "출고완료" 입력됨`);
            break;
          }
        }
      }
    }

    if (!yellowCellFound) {
      console.log(`⚠️ 행 ${targetRow}에서 노란색 셀을 찾을 수 없습니다.`);
    }

  } catch (error) {
    console.error('❌ 특정 행 노란색 셀 검사 오류:', error);
  }
}

/**
 * 수동 노란색 칠하기 감지 함수
 */
function detectManualYellowHighlight(e) {
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const column = range.getColumn();

  // 데이터 영역 내에서 배경색 변경 감지
  if (row >= CONFIG.DATA_START_ROW && row <= CONFIG.DATA_END_ROW &&
      column >= CONFIG.DATA_START_COLUMN && column <= CONFIG.DATA_END_COLUMN) {

    // 약간의 지연을 두고 출고완료 확인 (배경색 설정이 완료된 후)
    Utilities.sleep(100);
    checkShipmentCompletion(sheet);
  }
}

/**
 * 노란색 셀 감지 및 출고완료 처리 메인 로직 (안정성 강화)
 * 새로운 로직: 노란색 셀의 정보를 BB/BC/BD 데이터와 매칭
 */
function checkShipmentCompletion(sheet) {
  // 처리 중복 방지
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(3000);
  } catch (lockError) {
    console.log(`⚠️ 출고완료 처리 Lock 실패: ${lockError.message}`);
    return;
  }

  try {
    console.log('=== 안정적 출고완료 확인 시작 ===');

    // 처리 안정성을 위한 지연
    Utilities.sleep(200);

    // BB/BC/BD/BE 열의 데이터 가져오기 (7행부터 마지막까지)
    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.LOOKUP_START_ROW) {
      console.log('⚠️ BB/BC/BD 데이터가 없습니다.');
      return;
    }

    const lookupRange = sheet.getRange(CONFIG.LOOKUP_START_ROW, CONFIG.BB_COLUMN,
                                      lastRow - CONFIG.LOOKUP_START_ROW + 1, 4);
    const lookupData = lookupRange.getValues();

    console.log(`📊 BB/BC/BD/BE 데이터 범위: BB${CONFIG.LOOKUP_START_ROW}:BE${lastRow} (${lookupData.length}행)`);

    // 6행 헤더 데이터 가져오기 (D열부터 AH열까지)
    const headerRange = sheet.getRange(CONFIG.HEADER_ROW, CONFIG.DATA_START_COLUMN,
                                      1, CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1);
    const headerData = headerRange.getValues()[0];

    console.log('📋 제품 헤더:', headerData.slice(0, 5) + '...');

    // 데이터 영역에서 노란색 셀들 검색
    const dataRange = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.DATA_START_COLUMN,
                                   CONFIG.DATA_END_ROW - CONFIG.DATA_START_ROW + 1,
                                   CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1);

    const backgrounds = dataRange.getBackgrounds();
    const values = dataRange.getValues();
    let yellowCellCount = 0;
    let matchedCount = 0;

    // 처리할 노란색 셀들 먼저 수집
    const yellowCells = [];

    for (let i = 0; i < backgrounds.length; i++) {
      for (let j = 0; j < backgrounds[i].length; j++) {
        const backgroundColor = backgrounds[i][j];

        if (isYellowColor(backgroundColor)) {
          const yellowRow = CONFIG.DATA_START_ROW + i;
          const yellowCol = CONFIG.DATA_START_COLUMN + j;
          const yellowValue = values[i][j];

          yellowCells.push({
            row: yellowRow,
            col: yellowCol,
            value: yellowValue,
            headerIndex: j
          });
        }
      }
    }

    console.log(`🔍 발견된 노란색 셀 개수: ${yellowCells.length}`);

    // 각 노란색 셀에 대해 안정적으로 처리
    for (let cell of yellowCells) {
      try {
        yellowCellCount++;

        // 노란색 셀의 정보 추출
        const orderName = sheet.getRange(cell.row, 1).getValue(); // A열 주문자
        const productName = headerData[cell.headerIndex]; // 6행 제품명

        console.log(`🟡 노란색 셀 ${yellowCellCount}:`);
        console.log(`   위치: 행${cell.row}, 열${cell.col}`);
        console.log(`   주문자(A${cell.row}): "${orderName}"`);
        console.log(`   제품: "${productName}"`);
        console.log(`   수량: ${cell.value}`);

        // BB/BC/BD 데이터와 매칭 확인 (안전성 체크 추가)
        for (let k = 0; k < lookupData.length; k++) {
          const bbValue = lookupData[k][0]; // BB열 주문자
          const bcValue = lookupData[k][1]; // BC열 제품
          const bdValue = lookupData[k][2]; // BD열 수량

          // null/undefined 체크 추가
          if (bbValue && bcValue && bdValue &&
              bbValue === orderName && bcValue === productName && bdValue === cell.value) {

            const matchRow = CONFIG.LOOKUP_START_ROW + k;
            console.log(`✅ 매칭 성공! 행 ${matchRow}:`);
            console.log(`   BB${matchRow}: "${bbValue}" === "${orderName}"`);
            console.log(`   BC${matchRow}: "${bcValue}" === "${productName}"`);
            console.log(`   BD${matchRow}: ${bdValue} === ${cell.value}`);

            try {
              // BE열에 "출고완료" 입력 (안전성 확보)
              const beCell = sheet.getRange(matchRow, CONFIG.BE_COLUMN);
              beCell.setValue('출고완료');
              matchedCount++;
              console.log(`✅ BE${matchRow}에 "출고완료" 입력됨`);

              // 처리 간 짧은 지연
              Utilities.sleep(100);
            } catch (beError) {
              console.error(`❌ BE${matchRow} 입력 실패:`, beError);
            }
            break;
          }
        }

      } catch (cellError) {
        console.error(`❌ 노란색 셀 처리 실패 (행${cell.row}, 열${cell.col}):`, cellError);
      }
    }

    console.log(`📊 처리 결과: 노란색 셀 ${yellowCellCount}개, 출고완료 ${matchedCount}개`);
    console.log('=== 출고완료 확인 종료 ===\n');

  } catch (error) {
    console.error('❌ 출고완료 확인 오류:', error);
    SpreadsheetApp.getUi().alert('오류', '출고완료 확인 중 오류가 발생했습니다: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 색상이 노란색인지 확인하는 함수
 */
function isYellowColor(color) {
  if (!color) return false;

  const yellowVariants = [
    '#FFFF00',  // 순수 노란색
    '#FFFFFF00', // 알파값 포함
    '#ffff00',  // 소문자
    '#FF0',     // 축약형
    'yellow'    // 색상명
  ];

  const normalizedColor = color.toString().toUpperCase();
  return yellowVariants.some(yellow =>
    normalizedColor === yellow.toUpperCase() ||
    normalizedColor.includes('FFFF00')
  );
}

/**
 * 모든 노란색 셀 확인 (수동 실행용)
 */
function checkAllYellowCells() {
  const sheet = SpreadsheetApp.getActiveSheet();
  console.log('모든 노란색 셀 확인 시작');

  checkShipmentCompletion(sheet);

  SpreadsheetApp.getUi().alert('노란색 셀 확인이 완료되었습니다.');
}

/**
 * 모든 노란색 하이라이트 제거
 */
function clearAllYellowHighlights() {
  const sheet = SpreadsheetApp.getActiveSheet();

  try {
    const range = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.DATA_START_COLUMN,
                               CONFIG.DATA_END_ROW - CONFIG.DATA_START_ROW + 1,
                               CONFIG.DATA_END_COLUMN - CONFIG.DATA_START_COLUMN + 1);

    // 모든 배경색 제거
    range.setBackground(null);

    // 결과 셀도 초기화
    const resultCell = sheet.getRange(CONFIG.HEADER_ROW, CONFIG.RESULT_COLUMN);
    resultCell.setValue('');

    console.log('모든 노란색 하이라이트 제거 완료');
    SpreadsheetApp.getUi().alert('모든 노란색 하이라이트가 제거되었습니다.');

  } catch (error) {
    console.error('노란색 제거 오류:', error);
    SpreadsheetApp.getUi().alert('오류가 발생했습니다: ' + error.message);
  }
}

/**
 * 필요한 트리거들 생성
 */
function createTriggers() {
  // 기존 트리거 삭제
  deleteAllTriggers();

  // onEdit 트리거 생성 (체크박스용)
  ScriptApp.newTrigger('onEdit')
    .onEdit()
    .create();

  // onChange 트리거 생성 (시트 구조 변경용)
  ScriptApp.newTrigger('onChange')
    .onChange()
    .create();

  // 시간 기반 트리거 생성 (1분마다 수동 노란색 감지)
  ScriptApp.newTrigger('checkAllYellowCells')
    .timeBased()
    .everyMinutes(1)
    .create();

  SpreadsheetApp.getUi().alert('트리거가 성공적으로 생성되었습니다!\n\n• 체크박스: 즉시 반영\n• 수동 노란색 칠하기: 1분마다 자동 확인');
  console.log('모든 트리거 생성 완료 (onEdit + onChange + 1분 타이머)');
}

/**
 * 모든 트리거 삭제
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  console.log(`${triggers.length}개의 트리거가 삭제되었습니다.`);
}

/**
 * 트리거 상태 확인
 */
function showTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  let message = `현재 설정된 트리거: ${triggers.length}개\n\n`;

  triggers.forEach((trigger, index) => {
    message += `${index + 1}. ${trigger.getHandlerFunction()} (${trigger.getEventType()})\n`;
  });

  if (triggers.length === 0) {
    message += '설정된 트리거가 없습니다.\n"트리거 생성" 메뉴를 실행해주세요.';
  }

  SpreadsheetApp.getUi().alert(message);
}

/**
 * 시트 변경 감지 트리거 (배경색 변경 등)
 */
function onChange(e) {
  console.log('📝 시트 변경 감지 - 수동 노란색 칠하기 체크:', e);

  // 약간의 지연을 두고 출고완료 상태 확인 (배경색 변경 완료 대기)
  Utilities.sleep(300);

  const sheet = SpreadsheetApp.getActiveSheet();
  console.log('🔄 수동 변경으로 인한 출고완료 상태 재확인 시작');
  checkShipmentCompletion(sheet);
}

/**
 * 권한 초기화 함수 (최초 1회 실행)
 */
function initializePermissions() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();

    // 권한 확인을 위한 기본 작업
    const testRange = sheet.getRange(1, 1);
    testRange.getValue();

    SpreadsheetApp.getUi().alert(
      '권한 초기화 완료',
      '노란색 칠하기 시스템이 준비되었습니다!\n\n다음 단계:\n1. "트리거 생성" 메뉴 실행\n2. C열에 체크박스 추가\n3. BB열, BC열에 목표 위치 설정',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    console.error('권한 초기화 오류:', error);
    SpreadsheetApp.getUi().alert('권한 설정 중 오류가 발생했습니다: ' + error.message);
  }
}