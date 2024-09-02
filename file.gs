function selectFiles() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = DriveApp.getFileById(ss.getId()).getParents().next();
  
  let files = folder.getFiles();
  let file_list = [];
  while (files.hasNext()) {
    let file = files.next();
    let file_id = file.getId();
    let file_name = file.getName();
    file_list.push({ id: file_id, name: file_name });
  }

  // ファイルリストを番号付きで作成
  let fileListText = "操作するファイルを番号で指定してください:\n";
  file_list.forEach((file, index) => {
    fileListText += `${index + 1}. ${file.name}\n`;
  });

  // ダイアログでファイル番号を入力してもらう
  let response = ui.prompt(fileListText + "\n番号を入力してください:", ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    let selectedNumber = parseInt(response.getResponseText(), 10);
    
    // 入力が有効な番号かどうかを確認
    if (selectedNumber > 0 && selectedNumber <= file_list.length) {
      let selectedFile = file_list[selectedNumber - 1];
      Logger.log("Selected File ID: " + selectedFile.id);
      ui.alert("選択されたファイル:", `ファイル名: ${selectedFile.name}\nファイルID: ${selectedFile.id}`, ui.ButtonSet.OK);
      return selectedFile.id;
    } else {
      ui.alert("無効な番号が指定されました。ファイル選択がキャンセルされました。");
    }
  } else {
    ui.alert("ファイル選択がキャンセルされました。");
  }
}

function createFile() {
  const ui = SpreadsheetApp.getUi(); // UIを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // アクティブなスプレッドシートを取得
  const folder = DriveApp.getFileById(ss.getId()).getParents().next(); // 現在のフォルダを取得

  // ダイアログでファイル名を入力してもらう
  const response = ui.prompt("新しいファイル名を入力してください:");
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const newFileName = response.getResponseText();
    
    if (newFileName) {
      // アクティブなスプレッドシートをコピーして新しいファイルを作成
      const file = DriveApp.getFileById(ss.getId());
      const newFile = file.makeCopy(newFileName, folder);
      
      // コピーが成功したことを通知
      ui.alert("ファイルが正常に作成されました!", `ファイル "${newFileName}" が同じフォルダ内に作成されました。`, ui.ButtonSet.OK);
    } else {
      ui.alert("ファイル名が入力されていません。ファイルは作成されませんでした。");
    }
  } else {
    ui.alert("ファイル作成がキャンセルされました。");
  }
}
