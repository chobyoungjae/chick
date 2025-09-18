/**
 * ========================================
 * 이 파일은 현재 사용하지 않습니다.
 * 새로운 노란색칠하기.js 파일을 사용해주세요.
 * ========================================
 */

/*
function processCheckout() {
  // 중복 실행 방지
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(1000); // 1초 대기
  } catch (e) {
    console.log("다른 프로세스가 실행 중입니다. 건너뜀.");
    return;
  }

  // 스프레드시트 가져오기
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 현재 활성 시트명 확인
  const activeSheet = spreadsheet.getActiveSheet();
  console.log("현재 활성 시트:", activeSheet.getName());

  // 모든 시트명 확인
  const sheets = spreadsheet.getSheets();
  console.log("전체 시트 목록:");
  sheets.forEach((sheet) => console.log("- " + sheet.getName()));

  const frontSheet = spreadsheet.getSheetByName("프론트앤드");
  const backupSheet = spreadsheet.getSheetByName("일별 발주량 백업본");

  if (!frontSheet) {
    console.log("프론트앤드 시트를 찾을 수 없습니다!");
    SpreadsheetApp.getUi().alert(
      "오류",
      "프론트앤드 시트를 찾을 수 없습니다.\n시트명을 확인해주세요.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  if (!backupSheet) {
    console.log("일별 발주량 백업본 시트를 찾을 수 없습니다!");
    SpreadsheetApp.getUi().alert(
      "오류",
      "일별 발주량 백업본 시트를 찾을 수 없습니다.\n시트명을 확인해주세요.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // 프론트앤드 시트에서 필요한 데이터 가져오기
  const startDate = frontSheet.getRange("A1").getValue(); // 시작일
  const endDate = frontSheet.getRange("A2").getValue(); // 종료일

  console.log("=== 날짜 디버그 ===");
  console.log("A1 원본 값:", startDate);
  console.log("A2 원본 값:", endDate);
  console.log("A1 타입:", typeof startDate);
  console.log("A2 타입:", typeof endDate);
  console.log("A1이 Date 객체인가:", startDate instanceof Date);
  console.log("A2이 Date 객체인가:", endDate instanceof Date);
  console.log("A1 Boolean 체크:", !!startDate);
  console.log("A2 Boolean 체크:", !!endDate);

  // 날짜 조건을 더 관대하게 체크
  const hasStartDate =
    startDate && (startDate instanceof Date || typeof startDate === "string");
  let hasEndDate =
    endDate && (endDate instanceof Date || typeof endDate === "string");

  console.log("hasStartDate:", hasStartDate);
  console.log("hasEndDate:", hasEndDate);

  // A2가 비어있으면 A1의 날짜를 사용
  if (hasStartDate && !hasEndDate) {
    console.log("A2가 비어있어서 A1의 날짜를 사용합니다.");
    frontSheet.getRange("A2").setValue(startDate);
    hasEndDate = true;
  }

  // 날짜가 없거나 다른 경우 함수 종료
  if (!hasStartDate || !hasEndDate) {
    console.log("A1 또는 A2에 날짜가 없습니다.");
    SpreadsheetApp.getUi().alert(
      "오류",
      "A1(시작일) 또는 A2(종료일)에 날짜가 입력되지 않았습니다.\n날짜를 입력해주세요.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // 날짜를 Date 객체로 변환
  const startDateObj =
    startDate instanceof Date ? startDate : new Date(startDate);
  const endDateObj = endDate instanceof Date ? endDate : new Date(endDate);

  console.log("변환된 시작일:", startDateObj);
  console.log("변환된 종료일:", endDateObj);

  // A1과 B1이 같은 날짜인지 확인 (날짜만 비교)
  const startDateStr = startDateObj.toDateString();
  const endDateStr = endDateObj.toDateString();

  console.log("비교용 시작일:", startDateStr);
  console.log("비교용 종료일:", endDateStr);

  if (startDateStr !== endDateStr) {
    console.log("A1과 A2가 다른 날짜입니다.");
    console.log("A1 날짜:", startDateStr);
    console.log("A2 날짜:", endDateStr);

    // 사용자에게 알림창 표시
    SpreadsheetApp.getUi().alert(
      "날짜 불일치",
      `시작일과 종료일이 다릅니다.\n\n시작일(A1): ${startDateStr}\n종료일(A2): ${endDateStr}\n\n시작일과 종료일을 같게 맞춰주세요.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    console.log("함수를 실행하지 않습니다.");
    return;
  }

  console.log("A1과 A2가 같은 날짜입니다. 함수를 실행합니다:", startDateStr);

  // 8행 헤더명 가져오기 (D:AH 열)
  const headers = frontSheet.getRange("8:8").getValues()[0];
  const productHeaders = headers.slice(3, 34); // D열부터 AH열까지

  // 9행부터 마지막 행까지 데이터 가져오기
  const lastRow = frontSheet.getLastRow();
  const checkoutRange = frontSheet.getRange("C9:C" + lastRow); // C열 출고완료 상태
  const orderRange = frontSheet.getRange("A9:A" + lastRow); // A열 주문자
  const quantityRange = frontSheet.getRange("D9:AH" + lastRow); // D:AH열 수량

  const checkoutStatus = checkoutRange.getValues();
  const orders = orderRange.getValues();
  const quantities = quantityRange.getValues();

  // 일별 발주량 백업본 시트의 데이터 가져오기
  const backupLastRow = backupSheet.getLastRow();
  const backupData = backupSheet.getRange("A1:F" + backupLastRow).getValues();

  let processedCount = 0;

  // 모든 행 처리 (출고완료와 미완료 모두)
  for (let i = 0; i < checkoutStatus.length; i++) {
    const orderName = orders[i][0]; // 주문자명
    const quantityRow = quantities[i]; // 해당 행의 D:AH 수량 데이터
    const status = checkoutStatus[i][0]; // 현재 상태

    console.log(`행 ${i + 9}: 주문자 = ${orderName}, 상태 = ${status}`);

    // 수량이 있는 열들 찾기
    for (let j = 0; j < quantityRow.length; j++) {
      if (quantityRow[j] && quantityRow[j] > 0) {
        // 수량이 있는 경우
        const productName = productHeaders[j]; // 해당 제품명

        console.log(`  제품: ${productName}, 수량: ${quantityRow[j]}`);

        if (status === "출고완료") {
          // 출고완료인 경우 백업시트에 "출고완료" 입력
          const updated = updateBackupSheet(
            backupSheet,
            backupData,
            orderName,
            productName,
            startDate,
            endDate,
            "출고완료"
          );
          if (updated) processedCount++;
        } else if (status === "미완료" || status === "" || !status) {
          // 미완료인 경우 백업시트에서 "출고완료" 삭제
          const updated = updateBackupSheet(
            backupSheet,
            backupData,
            orderName,
            productName,
            startDate,
            endDate,
            ""
          );
          if (updated) processedCount++;
        }
      }
    }
  }

  // 완료 메시지
  console.log(
    `출고완료 처리가 완료되었습니다! 업데이트된 행 수: ${processedCount}`
  );

  // 락 해제
  lock.releaseLock();
}

function updateBackupSheet(
  sheet,
  data,
  orderName,
  productName,
  startDate,
  endDate,
  setValue
) {
  console.log(`백업 시트에서 찾는 조건:`);
  console.log(`  주문자: "${orderName}"`);
  console.log(`  제품: "${productName}"`);
  console.log(`  시작일: ${startDate}`);
  console.log(`  종료일: ${endDate}`);
  console.log(`  설정값: "${setValue}"`);

  let updated = false;

  for (let i = 1; i < data.length; i++) {
    // 1행은 헤더이므로 2행부터 시작
    const row = data[i];
    const backupOrderName = row[1]; // B열: 주문자
    const backupProductName = row[2]; // C열: 제품
    const backupDate = row[4]; // E열: 날짜

    // 주문자와 제품이 일치하는 행만 자세히 로그
    if (backupOrderName === orderName && backupProductName === productName) {
      console.log(`  후보 행 ${i + 1}:`);
      console.log(
        `    주문자: "${backupOrderName}" === "${orderName}" = ${
          backupOrderName === orderName
        }`
      );
      console.log(
        `    제품: "${backupProductName}" === "${productName}" = ${
          backupProductName === productName
        }`
      );
      console.log(`    날짜: ${backupDate}`);
      console.log(`    날짜 >= 시작일: ${backupDate >= startDate}`);
      console.log(`    날짜 <= 종료일: ${backupDate <= endDate}`);

      // 조건 확인
      if (
        backupOrderName === orderName &&
        backupProductName === productName &&
        backupDate >= startDate &&
        backupDate <= endDate
      ) {
        if (setValue === "출고완료") {
          console.log(`    ✅ 조건 일치! 행 ${i + 1}에 출고완료 입력`);
        } else {
          console.log(`    ✅ 조건 일치! 행 ${i + 1}에서 출고완료 삭제`);
        }

        // F열에 값 입력 (실제 행 번호는 i+1)
        sheet.getRange(i + 1, 6).setValue(setValue);
        updated = true;
      } else {
        console.log(`    ❌ 조건 불일치`);
      }
    }
  }

  if (!updated) {
    console.log("  조건에 맞는 행을 찾지 못했습니다.");
  }

  return updated;
}

// 1. onEdit 트리거용 함수 (셀 편집 시)
/*
function onEdit(e) {
  // 셀 편집 시 processCheckout 실행
  console.log("셀이 편집되었습니다:", e);
  processCheckout();
}

// 2. onChange 트리거용 함수 (시트 구조 변경 시)
function onChange(e) {
  // 체크박스 변경도 감지할 수 있음
  console.log("시트가 변경되었습니다:", e);
}

// 2. 수동 트리거 생성 함수들
function createOnEditTrigger() {
  ScriptApp.newTrigger("onEdit").onEdit().create();

  SpreadsheetApp.getUi().alert("onEdit 트리거가 생성되었습니다!");
}

function createOnChangeTrigger() {
  ScriptApp.newTrigger("onChange").onChange().create();

  SpreadsheetApp.getUi().alert("onChange 트리거가 생성되었습니다!");
}

// 3. 시간 기반 트리거 (1분마다 체크박스 상태 확인)
function createTimeTrigger() {
  ScriptApp.newTrigger("checkAllCheckboxes")
    .timeBased()
    .everyMinutes(1)
    .create();

  SpreadsheetApp.getUi().alert("1분마다 실행되는 트리거가 생성되었습니다!");
}

function checkAllCheckboxes() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const frontSheet = spreadsheet.getSheetByName("프론트앤드");

  // 이전 상태와 비교해서 새로 체크된 것만 처리
  processCheckout();
}

// 4. 모든 트리거 삭제
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  SpreadsheetApp.getUi().alert("모든 트리거가 삭제되었습니다!");
}

// 5. 현재 트리거 상태 확인
function showTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let message = `현재 설정된 트리거: ${triggers.length}개\n\n`;

  triggers.forEach((trigger, index) => {
    message += `${
      index + 1
    }. ${trigger.getHandlerFunction()} (${trigger.getEventType()})\n`;
  });

  SpreadsheetApp.getUi().alert(message);
}

// 권한 설정용 초기화 함수 (최초 1회 실행 필요)
function initializePermissions() {
  try {
    // 스프레드시트 접근 권한 확인
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();

    // UI 알림 권한 확인
    SpreadsheetApp.getUi().alert(
      "권한 설정이 완료되었습니다!\n이제 체크박스를 사용할 수 있습니다."
    );

    console.log("권한 초기화 완료");
  } catch (error) {
    console.error("권한 설정 중 오류:", error);
  }
}

// 수동 실행용 함수
function manualRun() {
  processCheckout();
}

// 트리거 상태 확인 함수
function checkTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  console.log("현재 설정된 트리거 수:", triggers.length);

  triggers.forEach((trigger, index) => {
    console.log(
      `트리거 ${index + 1}:`,
      trigger.getHandlerFunction(),
      trigger.getEventType()
    );
  });
}
*/
