function combineRowsByColumnA() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return; // 데이터 없으면 종료

  // A열 기준 오름차순 정렬
  data.sort(function(a, b) {
    return a[0] > b[0] ? 1 : (a[0] < b[0] ? -1 : 0);
  });

  var header = data.shift(); // 첫 줄은 헤더
  var mergedData = [];
  var currentKey = data[0][0];
  var currentRow = [currentKey];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[0] === currentKey) {
      currentRow = currentRow.concat(row.slice(1));
    } else {
      mergedData.push(currentRow);
      currentKey = row[0];
      currentRow = [currentKey].concat(row.slice(1));
    }
  }
  mergedData.push(currentRow);

  // 모든 행의 길이 맞추기
  var maxLen = mergedData.reduce(function(max, row) {
    return Math.max(max, row.length);
  }, 0);
  mergedData = mergedData.map(function(row) {
    while (row.length < maxLen) row.push("");
    return row;
  });

  // 결과 쓰기
  var output = [header];
  mergedData.forEach(function(row) {
    output.push(row);
  });

  sheet.clearContents();
  sheet.getRange(1, 1, output.length, maxLen).setValues(output);
}
