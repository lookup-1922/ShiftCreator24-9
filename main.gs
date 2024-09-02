function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("Scripts");
  //アイテムを追加
  menu.addItem("ファイル作成", "createFile");
  menu.addItem("シフト枠を作成", "createShiftSlots")
  menu.addItem("フォーム作成", "createForm");
  menu.addItem("回答を紐づける", "linkInput");
  menu.addItem("シフト作成", "createShift");
  //スプレッドシートに反映
  menu.addToUi();
}

function initializeSettings() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // シートが存在しない場合は新しく作成し、存在する場合はそのシートを取得
  let nameListSheet = ss.getSheetByName('name_list') || ss.insertSheet('name_list');
  let shiftConfigSheet = ss.getSheetByName('shift_config') || ss.insertSheet('shift_config');
  let shiftOutputSheet = ss.getSheetByName('shift_output') || ss.insertSheet('shift_output');
  
  // シートの内容をクリア
  nameListSheet.clear();
  shiftConfigSheet.clear();
  shiftOutputSheet.clear();

  // name_list シートの1行目を設定
  nameListSheet.getRange('A1').setValue('ID');
  nameListSheet.getRange('B1').setValue('名前');
  nameListSheet.getRange('C1').setValue('対象');
  nameListSheet.getRange('D1').setValue('回数');

  // shift_config シートの1行目を設定
  shiftConfigSheet.getRange('A1').setValue('日付');
  shiftConfigSheet.getRange('B1').setValue('時間');
  shiftConfigSheet.getRange('C1').setValue('役割1-人数');
  shiftConfigSheet.getRange('D1').setValue('役割2-人数');
  shiftConfigSheet.getRange('E1').setValue('役割3-人数');

  // shift_output シートの1行目を設定
  shiftOutputSheet.getRange('A1').setValue('日付');
  shiftOutputSheet.getRange('B1').setValue('時間');
  shiftOutputSheet.getRange('C1').setValue('役割1');
  shiftOutputSheet.getRange('D1').setValue('役割2');
  shiftOutputSheet.getRange('E1').setValue('役割3');

  // 成功の通知
  ui.alert('初期設定が完了しました。シートが作成されました。');
}
