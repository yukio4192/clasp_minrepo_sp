function main() {
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  // インポート先のURL一覧を取得
  // TODO: rangeの指定を引数で出来るようにする
  let importData = getImportData("url_source", "A19:D21");
  // シートをコピーして追加し、URL情報を挿入する
  importData.forEach((index) => {
    // 新しいシートを作成
    let newSheet = createAndRenameSheet('am/d' ,index.date);
    // URL情報を挿入
    setB2Value(newSheet, index.url);
  });
}

/**
 * 指定されたシート名と範囲からURLと日付のデータを取得する
 *
 * @param {string} sheetName シート名
 * @param {string} range 範囲
 * @returns {object[]} データの配列
 */
function getImportData(sheetName, range) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var values = sheet.getRange(range).getValues();

  var data = [];
  for (var i = 0; i < values.length; i++) {
    data.push({
      date: values[i][0], // 日付列のインデックスは0と仮定
      url: values[i][3], // URL列のインデックスは3と仮定
    });
  }
  return data;
}

/**
 * 指定されたシート名を元に、新しいシートを作成して名前を変更する
 *
 * @param {string} sourceSheetName コピー元のシート名
 * @param {string} newSheetName 新しいシート名
 * @returns {SpreadsheetApp.Sheet} 作成された新しいシート
 */
function createAndRenameSheet(sourceSheetName, newSheetName) {
  // アクティブなスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // コピー元のシートを取得
  var sourceSheet = spreadsheet.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    throw new Error('コピー元のシート "' + sourceSheetName + '" が見つかりません。');
  }

  // シートをコピーしてリネーム
  var newSheet = sourceSheet.copyTo(spreadsheet);
  newSheet.setName(newSheetName);

  return newSheet;
}

/**
* 指定されたシートのB2セルに文字列を入力する
*
* @param {SpreadsheetApp.Sheet} sheet 操作対象のシート
* @param {string} url B2セルに入力する元URL
*/
function setB2Value(sheet, url) {
  // B2セルを取得
  var range = sheet.getRange("B2");
  // 共通のクエリパラメータを末尾に追加
  var queryString = "?kishu=all&sort=num";
  // B2セルに文字列を設定
  range.setValue(url + queryString);
}