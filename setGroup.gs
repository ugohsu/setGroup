/*
HTML 表示
*/

function doGet() {
    const htmlOutput = HtmlService.createTemplateFromFile('index').evaluate();
    htmlOutput
    // スマホ対応
        .addMetaTag('viewport', 'width=device=width, initial-scale=1')
        .setTitle('グループ編成');
    return htmlOutput;
}

/*
スプレッドシートの読み込み
*/

// URL を受け取ると、そのスプレッドシートの名前一覧を返す関数
function replSheetNames(URL) {
  let SS = SpreadsheetApp.openByUrl(URL);
  let sheetNames = SS.getSheets().map(x => x.getSheetName());
  return sheetNames;
}

/*
データの読み込み
*/

function inputData(URL, dataSheet) {
  let SS = SpreadsheetApp.openByUrl(URL);
  let rslt = SS.getSheetByName(dataSheet).getDataRange().getValues();
  convertToString(rslt);
  return rslt;
}

function convertToString(array) {
  for (let i = 0; i < array.length; i++) {
    for (let j = 0; j < array[i].length; j++){
      array[i][j] = String(array[i][j]);
    }
  }
}

/*
データの書き込み
*/

function writeSpreadSheet(URL, data) {
  // 読み込み
  let SS = SpreadsheetApp.openByUrl(URL);

  // シート名の設定
  let sheetNames = SS.getSheets().map(x => x.getSheetName());
  let name = chgName(sheetNames, "group");

  // 新しいシートの挿入
  SS.insertSheet(name);

  // 新しいシートにデータを書き込み
  let body = [data.names].concat(data.body);
  let wSheet = SS.getSheetByName(name);
  wSheet.getRange(1, 1, body.length, body[0].length).setValues(body);

  // メッセージ
  return name + "に書き込みました。";
}

// name が重複している場合に末尾に再帰的に .x を付ける
function chgName(list, name) {
  if (list.indexOf(name) != -1) {
    name = name + ".x";
    return chgName(list, name);
  } else {
    return name;
  }
}
