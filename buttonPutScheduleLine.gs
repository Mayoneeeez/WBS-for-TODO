//概要
//予定線の挿入ボタン機能

const scheduleLineColor = "#a0d8ef";
const scheduleLineResultMessage = "予定線挿入に成功しました。"

//TODO
//共通要素はどこかで管理したい
function putScheduleLine() {

 try{
  // 現在見ているスプレッドシートを取得する
  const currentSheet = SpreadsheetApp.getActiveSheet();
  // シート名を取得する
  const currentSheetName = currentSheet.getName();

  //選択されたセルの行・列情報を配列として取得
  const selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
  const rangeList = selection.getActiveRangeList();
  const ranges = rangeList.getRanges();
  let rowList = [];
  let columnList = [];
  for (let i = 0; i < ranges.length; i++) {
    let range = ranges[i];
    let numRows = range.getNumRows();
    let numCols = range.getNumColumns();

    for (let row = 1; row <= numRows; row++) {
      for (let col = 1; col <= numCols; col++) {
        let cell = range.getCell(row, col);
        let rowIndex = cell.getRow();
        let columnIndex = cell.getColumn();
        rowList.push(rowIndex);
        columnList.push(columnIndex);
      }
    }
  }
  
  //エラーハンドリング
  //タスクの行範囲外判定
  let target = [1, 2, 3, 4, 5]; 
  if(rowList.some(value => target.includes(value))) {
    throw new Error("予定線を挿入したいタスク番号を1つ以上選択してからボタンを押してください。")
  }

  //タスクの列範囲外判定
  target = [1];
  if((columnList.filter(value => !target.includes(value))).length > 0){
    throw new Error("予定線を挿入したいタスク番号を1つ以上選択してからボタンを押してください。")
  }


  for(let i = 0; i < rowList.length; i++){
    let selectedRow = rowList[i];
    let selectedCol = columnList[i];
    //予定線行の1行下の行に実績線を入れる私用のため、2重にひかれないように調整
    if(selectedRow % 2 == 1){
      continue;
    }

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

    for(let col = matchingCells[0].col; col <= matchingCells[1].col; col++){
      //選択した予定線行に色を付ける
      currentSheet.getRange(selectedRow, col).setBackground(scheduleLineColor);
    }
  }
 }catch(e){
  Browser.msgBox(e.message);
 }
 
 Browser.msgBox(scheduleLineResultMessage);

}