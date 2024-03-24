function myFunction() {
  Logger.log("Hello Mayoneeeez");
}

function scheduleLine() {

  // 現在見ているスプレッドシートを取得する
  const currentSheet = SpreadsheetApp.getActiveSheet();
  // シート名を取得する
  const currentSheetName = currentSheet.getName();

  //選択されているセルの行・列の位置を取得
  //TODO trycatchで選択されていない場合exception出したい
  const selectedCell = currentSheet.getActiveCell();
  const selectedRow = selectedCell.getRow();
  const selectedColumn = selectedCell.getColumn();

  //選択した予定開始日/終了日を取得
  const scheduledStartDate = currentSheet.getRange('L' + selectedRow).getValue();
  const scheduledEndDate = currentSheet.getRange('M' + selectedRow).getValue();

  // シートのデータ範囲を取得
  const dataRange = currentSheet.getDataRange();
  const values = dataRange.getValues()[4]; //日付行で値が存在する箇所の値を配列として取得

  // タイムゾーンを取得
  const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  // 条件に合致するセルの位置を保持するための配列
  const matchingCells = [];

  for(let col = 0; col < values.length; col++){
    const cellValue = values[col];
    Logger.log(cellValue.getDate());
    Logger.log(scheduledStartDate.getDate());
    Logger.log(cellValue == scheduledStartDate);
    // Logger.log(Utilities.formatDate(cellValue, timeZone, 'yyyy-MM-dd'));
    // Logger.log(Utilities.formatDate(scheduledStartDate, timeZone, 'yyyy-MM-dd'));

    if(cellValue == scheduledStartDate){
      matchingCells.push({ row: selectedRow, col: col + 1 });
    }

    if(cellValue == scheduledEndDate){
      matchingCells.push({ row: selectedRow, col: col + 1 });
    }
  }


  //デバッグ用
  // Browser.msgBox("セルの選択位置","行："+selectedRow+ "、列："+selectedColumn, Browser.Buttons.OK);

  Logger.log(currentSheet);
  Logger.log(currentSheetName);
  Logger.log(scheduledStartDate);
  Logger.log(scheduledEndDate);
  Logger.log(dataRange);
  Logger.log(selectedRow);
  Logger.log(selectedColumn);
  Logger.log(values);
  Logger.log(matchingCells);
}