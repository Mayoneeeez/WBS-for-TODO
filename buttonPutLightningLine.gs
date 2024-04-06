function putLightningLine() {

  // 現在見ているスプレッドシートを取得する
  const currentSheet = SpreadsheetApp.getActiveSheet();

  // 円の画像を取得
  var imageUrl = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAOEAAADhCAMAAAAJbSJIAAAAA1BMVEX/AAAZ4gk3AAAASElEQVR4nO3BgQAAAADDoPlTX+AIVQEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADwDcaiAAFXD1ujAAAAAElFTkSuQmCC'; // 画像のURLを指定してください

  // 画像をシートに挿入
  var a = currentSheet.insertImage(imageUrl, 10, 10).setHeight(10).setWidth(10);
  a; // 1行目、1列目に挿入されます
}
