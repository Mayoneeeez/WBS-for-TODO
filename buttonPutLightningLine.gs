//概要
//イナズマ線の挿入ボタン機能

const lightningLineColor = "#ffff00";
const lightningLineResultMessage001 = "イナズマ線挿入に成功しました。"
const lightningLineErrorMessage001 = "イナズマ線を挿入する起点となる日を選択してからボタンを押してください。";
const lightningLineErrorMessage002 = "イナズマ線を挿入する起点となる日は1日のみにしてください。";

async function putLightningLine() {

  let message;//処理完了時メッセ―ジ

  try {
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
    let target_row = [1, 2, 3, 4, 5];
    if (rowList.some(value => target_row.includes(value))) {
      throw new Error(lightningLineResultMessage001)
    }

    //タスクの列範囲外判定
    let target_col = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];
    if (columnList.some(value => target_col.includes(value))) {
      throw new Error(lightningLineResultMessage001)
    }

    //イナズマ線を挿入したい一列のみを指定してない判定
    for (let i = 1; i < columnList.length; i++) {
      if (columnList[0] !== columnList[i]) {
        throw new Error(lightningLineResultMessage002)
      }
    }


    //選択日付をイナズマ色で塗る
    currentSheet.getRange(5, columnList[0]).setBackground(lightningLineColor);

    for (let i = 0; i < rowList.length; i++) {

      let selectedRow = rowList[i];
      let selectedCol = columnList[i];

      const scheduledStartDate = currentSheet.getRange('L' + selectedRow).getValue();//選択した予定開始日
      const scheduledEndDate = currentSheet.getRange('M' + selectedRow).getValue();//選択した終了日を取得
      const progressRate = currentSheet.getRange(selectedRow, 11).getValue() || 0.01;//進捗率(入力されていない場合1%とすることで予定開始日にイナズマ線を入力する)
      const dateDiff = getDateDifferenceInDays(scheduledStartDate, scheduledEndDate);//予定開始日・終了日の差分

      //予定線行の1行下の行に実績線を入れる私用のため、2重にひかれないように調整
      if (selectedRow % 2 == 1) {
        continue;
      }

      //予定日が入力されていない場合その行の処理をしない
      if (scheduledStartDate == null) {
        continue;
      }


      // シートのデータ範囲を取得
      const dataRange = currentSheet.getDataRange();
      const values = dataRange.getValues()[4]; //日付行で値が存在する箇所の値を配列として取得

      //予定開始日に対応するセル取得
      for (let col = 0; col < values.length; col++) {
        const cellValue = values[col];

        if (String(cellValue) == scheduledStartDate) {
          //進捗率に応じてイナズマ線を挿入
          if (progressRate == 1 && col < selectedCol) {
            currentSheet.getRange(selectedRow, selectedCol).setBackground(lightningLineColor);
          } else {
            currentSheet.getRange(selectedRow, col + Math.ceil(progressRate * dateDiff) + 1).setBackground(lightningLineColor);
            Logger.log(col)
            Logger.log(Math.ceil(progressRate * dateDiff))
          }
          break;
        }
      }


    }

    // 処理成功メッセージ
    message = lightningLineResultMessage001;

  } catch (e) {
    // エラーメッセージ
    message = e.message;
  }

  // メッセージを表示
  await showMessage(message);

}

// メッセージを表示する関数
function showMessage(message) {
  SpreadsheetApp.flush(); // 変更を即座に反映するために追加
  Browser.msgBox(message);
}

//日付の差分を取得する関数
function getDateDifferenceInDays(dateString1, dateString2) {
  // 日付文字列からDateオブジェクトを作成
  var date1 = new Date(Date.parse(dateString1));
  var date2 = new Date(Date.parse(dateString2));

  // 日数の差分を計算
  var oneDay = 24 * 60 * 60 * 1000; // 1日のミリ秒数
  var diffInDays = Math.round(Math.abs((date2.getTime() - date1.getTime()) / oneDay));

  return diffInDays;
}