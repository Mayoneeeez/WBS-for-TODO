function myFunction() {
  Logger.log("Hello Mayoneeeez");
}

//TODO
//選択せずに予定線入力を押したときtrycatchで選択されていない場合exception出したい
//成功メッセージを出したい
//共通要素はどこかで管理したい
function scheduleLine() {

  // 現在見ているスプレッドシートを取得する
  const currentSheet = SpreadsheetApp.getActiveSheet();
  // シート名を取得する
  const currentSheetName = currentSheet.getName();

  //選択されているセルの行・列の位置を取得
  const selectedCell = currentSheet.getActiveCell();
  const selectedRow = selectedCell.getRow();
  const selectedColumn = selectedCell.getColumn();

  Logger.log("--------------------");
  const selectedCells = currentSheet.getActiveRangeList();
  for(let i = 0; i < selectedCells.length; i++){
    Logger.log(selectedCells[i].getRow());
    Logger.log(selectedCells[i].getColumn());
    Logger.log(selectedCells[i].getLastRow());
    Logger.log(selectedCells[i].getLastColumn());
  }
  Logger.log("--------------------");

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
    // Logger.log(cellValue);
    // Logger.log(scheduledStartDate);
    // Logger.log(typeof cellValue);
    // Logger.log(typeof scheduledStartDate);
    // Logger.log(String(cellValue) == scheduledStartDate);
    // Logger.log(Utilities.formatDate(cellValue, timeZone, 'yyyy-MM-dd'));
    // Logger.log(Utilities.formatDate(scheduledStartDate, timeZone, 'yyyy-MM-dd'));

    if(String(cellValue) == scheduledStartDate){
      matchingCells.push({ row: selectedRow, col: col + 1 });
    }

    if(String(cellValue) == scheduledEndDate){
      matchingCells.push({ row: selectedRow, col: col + 1 });
    }
  }

  for(let col = matchingCells[0].col; col < matchingCells[1].col; col++){
  const selectedRow = selectedCell.getRow();
    currentSheet.getRange(selectedRow, col).setBackground('Blue');
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
  Logger.log(matchingCells[0].row);
}



function actualLine() {

  // 現在見ているスプレッドシートを取得する
  const currentSheet = SpreadsheetApp.getActiveSheet();
  // シート名を取得する
  const currentSheetName = currentSheet.getName();

  //選択されているセルの行・列の位置を取得
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

    if(String(cellValue) == scheduledStartDate){
      matchingCells.push({ row: selectedRow, col: col + 1 });
    }

    if(String(cellValue) == scheduledEndDate){
      matchingCells.push({ row: selectedRow, col: col + 1 });
    }
  }

  for(let col = matchingCells[0].col; col < matchingCells[1].col; col++){
  const selectedRow = selectedCell.getRow();
    currentSheet.getRange(selectedRow + 1, col).setBackground('Orange');
  }
}