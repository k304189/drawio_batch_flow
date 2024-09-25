const COLOR_MAP = {
  "other-dataset" : {
    "BASE TABLE": "#B2D8FF",
    "EXTERNAL": "#B2FFB2",
    "VIEW": "#B2FFFF",
    "FUNCTION": "#FFFFB2",
    "TABLE FUNCTION": "#FFB2B2",
    "PROCEDURE": "#FFB2FF",
  },
  "my-dataset" : {
    "BASE TABLE": "#7FBFFF",
    "EXTERNAL": "#7FFF7F",
    "VIEW": "#7FFFFF",
    "FUNCTION": "#FFFF7F",
    "TABLE FUNCTION": "#FF7F7F",
    "PROCEDURE": "#FF7FFF",
  }
};

const SHAPE_MAP = {
  "BASE TABLE": "cylinder",
  "EXTERNAL": "cylinder",
  "VIEW": "document",
  "FUNCTION": "process",
  "TABLE FUNCTION": "process",
  "PROCEDURE": "process",
}

// draw.ioに読み込ませるCSVファイルを作成する関数
function outputFlowDiagramCsv() {
  const csvArray = [...getFlowDiagramCsvConfigArray(), ...getFlowDiagramCsvDataArray()];
  
  const html = HtmlService.createTemplateFromFile("csvConfig");
  html.csvConfig = csvArray.join("\n");
  SpreadsheetApp.getUi().showModalDialog(html.evaluate(), "Draw.io CSV");
}

// 出力するCSVファイルのデータ部分を出力する関数
function getFlowDiagramCsvDataArray() {
  // シートからデータを取得する
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const batchDataSheet = spreadSheet.getSheetByName("シート2");
  const dataset = batchDataSheet.getRange(2, 2).getValue();

  const batchDataArray = getDatasetRelatedData(dataset);

  // CSVに出力するように加工する
  let classConfigArray = [];
  const outputCsvDataArray = batchDataArray.map((row) => {
    classConfigArray = [];

    if(row[1] !== "my-project") {
      classConfigArray.push("other-project");
      classConfigArray.push("", "", "");
    } else {
      const objectType = row[0];
      let useColorMap;
      classConfigArray.push("my-project");
      classConfigArray.push(SHAPE_MAP[objectType]);

      if(dataset === row[2]) {
        useColorMap = COLOR_MAP["my-dataset"];
      } else {
        useColorMap = COLOR_MAP["other-dataset"];
      }
      classConfigArray.push(useColorMap[objectType]);
      classConfigArray.push("#FFFFFF");
    }
    return [...row, ...classConfigArray];
  });
  
  return outputCsvDataArray;
}

function getDatasetRelatedData(dataset) {
  // シートからデータを取得する
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const batchDataSheet = spreadSheet.getSheetByName("シート2");
  const startCell = batchDataSheet.getRange(10, 1);

  const batchDataArray = batchDataSheet.getRange(
    11,
    1,
    startCell.getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() - 9,
    startCell.getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn()
  ).getValues();


  const targetDataArray = batchDataArray;

  return targetDataArray;
}

// 出力するCSVファイルの設定部分を出力する関数
function getFlowDiagramCsvConfigArray() {
  const labelConfigArray = [
    "%object_name%"
  ];

  const configArray = [
    `# label: ${labelConfigArray.join('<br>')}`,
    '# styles: { \\',
    '#   "other-project" : "shape=cloud;fillColor=#A9A9A9;strokeColor=#FFFFFF;html=1;", \\',
    '#   "my-project" : "shape=%shape%;fillColor=%fill_color%;strokeColor=%stroke_color%;html=1;" \\',
    '# }',
    '# stylename: class',
    '# connect: {"from": "refference_objects", "to": "object_full_name", "invert": "true", "style":"curved=0;endArrow=blockThin;"}',
    '# connect: {"from": "data_creation_objects", "to": "object_full_name", "invert": "true", "style":"curved=0;endArrow=blockThin;dashed=1;"}',
    '# width: 100',
    '# height: 100',
    '# padding: 5',
    '# ignore: class,refference_objects,object_full_name,data_creation_objects,shape,fill_color,stroke_color',
    '# nodespacing: 60',
    '# levelspacing: 60',
    '# edgespacing: 40',
    '# layout: verticalflow',
    '## **********************************************************',
    '## CSV Data',
    '## **********************************************************',
    'object_type,project_id,dataset_id,object_name,object_full_name,refference_objects,data_creation_objects,class,shape,fill_color,stroke_color',
  ];

  return configArray;
}
