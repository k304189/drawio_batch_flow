// draw.ioに読み込ませるCSVファイルを作成する関数
function outputDrawioCsvConfig() {
  const csvArray = [...getConfigArray(), ...getDataArray()];
  
  const html = HtmlService.createTemplateFromFile("csvConfig");
  html.csvConfig = csvArray.join("\n");
  SpreadsheetApp.getUi().showModalDialog(html.evaluate(), "Draw.io CSV");
}

// 出力するCSVファイルのデータ部分を出力する関数
function getDataArray() {
  // シートからデータを取得する
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const batchDataSheet = spreadSheet.getSheetByName("シート1");

  const batchDataArray = batchDataSheet.getDataRange().getValues();

  // CSVに出力するように加工する
  const outputCsvDataArray = batchDataArray.map((row, index) => {
    let batchFrom = "batchFrom";
    // 先頭行以外は、batchFrom1～3の3列を1列に集約する
    if(index !== 0) {
      // 実行時間がDATEオブジェクトなので、文字列に変換する
      if(row[3]) {
        row[3] = Utilities.formatDate(row[3], 'Asia/Tokyo', 'HH:mm');
      }
      // 1列に複数の値を設定する場合、ダブルクォートで囲み、カンマで区切る必要がある
      batchFrom = `"${row.slice(-3).join(",")}"`
    } 
    return [...row.slice(0, -3), batchFrom].join(",");
  });
  
  return outputCsvDataArray;
}

// 出力するCSVファイルの設定部分を出力する関数
function getConfigArray() {
  const labelConfigArray = [
    "No.%no%：%name%",
    "<b>担当者</b>：%person%",
    "<b>実行サイクル</b>：%cycle%",
    "<b>開始時間</b>：%time%",
  ];

  const drawioCsvConfigArray = [
    `# label: ${labelConfigArray.join('<br>')}`,
    '# styles: { \\',
    '#   "daily1" : "fillColor=#048FFD ;strokeColor=#FFFFFF;html=1;", \\',
    '#   "weekly1" : "fillColor=#76C2FD ;strokeColor=#FFFFFF;html=1;", \\',
    '#   "other1" : "fillColor=#DBEEFD;strokeColor=#FFFFFF;html=1;", \\',
    '#   "daily2" : "fillColor=#93FD11;strokeColor=#FFFFFF;html=1;", \\',
    '#   "weekly2" : "fillColor=#C1FD77;strokeColor=#FFFFFF;html=1;", \\',
    '#   "other2" : "fillColor=#EEFDDC;strokeColor=#FFFFFF;html=1;" \\',
    '# }',
    '# stylename: class',
    '# connect: {"from": "batchFrom", "to": "name", "invert": "true", "style":"curved=0;endArrow=blockThin;"}',
    '# width: auto',
    '# height: auto',
    '# padding: 5',
    '# ignore: batchFrom',
    '# nodespacing: 60',
    '# levelspacing: 60',
    '# edgespacing: 40',
    '# layout: verticalflow',
    '## **********************************************************',
    '## CSV Data',
    '## **********************************************************'
  ];

  return drawioCsvConfigArray;
}
