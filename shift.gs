// 未作成の関数
function createShiftSlots() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(selectFiles());
  const folder = DriveApp.getFileById(ss.getId()).getParents().next();
  const name_list = ss.getSheetByName("name_list");
  const shift_config = ss.getSheetByName("shift_config");
  const shift_output = ss.getSheetByName("shift_output");

}

function createForm() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(selectFiles());
  const shift_config = ss.getSheetByName("shift_config");
  const shift_output = ss.getSheetByName("shift_output");

  // シートからデータを一括で取得
  const shiftConfigData = shift_config.getDataRange().getValues(); // 全データ取得
  const firstRow = shiftConfigData[0]; // 1行目（ヘッダー）
  const shiftConfigRows = shiftConfigData.slice(1); // データ行

  let roles = [];
  let shift_slots = [];

  // shift_configの1行目をshift_outputの1行目にコピー
  shift_output.getRange(1, 1, 1, firstRow.length).setValues([firstRow]);

  // rolesにはshift_configのC1, D1, E1, ...の文字列を格納
  for (let col = 2; col < firstRow.length; col++) {
    const cellValue = firstRow[col];
    if (cellValue) {
      roles.push(cellValue);
    }
  }

  // shift_slotsにはA2+B2, A3+B3, A4+B4の文字列を格納
  shiftConfigRows.forEach(row => {
    const aValue = row[0]; // A列の値
    const bValue = row[1]; // B列の値
    if (aValue && bValue) {
      shift_slots.push(aValue + bValue);
    }
  });

  const response = ui.prompt("フォーム名を入力してください:");

  if (response.getSelectedButton() == ui.Button.OK) {
    const newForm = FormApp.create(response.getResponseText());

    newForm.addTextItem()
      .setTitle("出席番号（半角数字のみ）")
      .setRequired(true);

    newForm.addTextItem()
      .setTitle("名前")
      .setRequired(true);

    newForm.addCheckboxItem()
      .setTitle("役割を選んでください。")
      .setChoiceValues(roles)
      .showOtherOption(false)
      .setRequired(true);

    newForm.addCheckboxItem()
      .setTitle("行けないものをすべて選択してください。")
      .setChoiceValues(shift_slots)
      .showOtherOption(false)
      .setRequired(false);

    newForm.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    ui.alert(newForm.shortenFormUrl(newForm.getPublishedUrl()));
  } else {
    ui.alert("フォーム作成がキャンセルされました。");
  }
}

function linkInput() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(selectFiles());
  const name_list = ss.getSheetByName("name_list");
  const form_response = ss.getSheetByName("フォームの回答 1");

  // name_listとform_responseのデータをまとめて取得
  const nameListData = name_list.getDataRange().getValues();
  const formResponseData = form_response.getDataRange().getValues();

  // IDごとに最新の回答を取得するためのマップを作成
  let latestResponses = {};

  for (let i = 1; i < formResponseData.length; i++) { // 1行目はヘッダー行なのでスキップ
    const timestamp = formResponseData[i][0]; // タイムスタンプ (A列)
    const id = formResponseData[i][1]; // フォーム回答のID (B列)

    if (!latestResponses[id] || latestResponses[id].timestamp < timestamp) {
      latestResponses[id] = {
        timestamp: timestamp,
        data: formResponseData[i].slice(1) // タイムスタンプ以降のデータをコピー
      };
    }
  }

  // name_listのIDと一致する最新のフォーム回答をコピー
  let updated = false;
  for (let j = 1; j < nameListData.length; j++) { // 1行目はヘッダー行なのでスキップ
    const nameListId = nameListData[j][0]; // name_listのID (A列)

    if (latestResponses[nameListId]) {
      // E列以降に最新のフォーム回答をコピー
      name_list.getRange(j + 1, 5, 1, latestResponses[nameListId].data.length).setValues([latestResponses[nameListId].data]);
      updated = true;
    }
  }

  if (updated) {
    ui.alert('最新のフォーム回答がname_listに関連付けられました。');
  } else {
    ui.alert('一致するIDが見つかりませんでした。');
  }

}

function createShift() {
  /*
  name_listのC列の対象かどうかはfalseになっているかどうか。
  空白でもtrueでもなんでもいい。
  */

  const ss = SpreadsheetApp.openById(selectFiles());
  const shift_output = ss.getSheetByName("shift_output");

  const shiftOutputData = shift_output.getRange("A2").getValue();

  if (shiftOutputData) {
    changeShift(ss);
  } else {
    createNewShift(ss);
  }

}

function createNewShift(ss) {
  const name_list = ss.getSheetByName("name_list");
  const shift_config = ss.getSheetByName("shift_config");
  const shift_output = ss.getSheetByName("shift_output");

  // データを一括で取得
  const nameListData = name_list.getDataRange().getValues();
  const shiftConfigData = shift_config.getDataRange().getValues();
  const shiftOutputData = shift_output.getDataRange().getValues();

  // shiftOutputData のサイズを調整
  if (shiftOutputData.length < shiftConfigData.length) {
    shiftOutputData.length = shiftConfigData.length;
  }
  shiftConfigData.forEach((row, index) => {
    if (shiftOutputData[index] === undefined) {
      shiftOutputData[index] = [];
    }
    while (shiftOutputData[index].length < row.length) {
      shiftOutputData[index].push(null);
    }
  });

  // 日付と時間を shiftOutputData にコピー
  for (let i = 1; i < shiftConfigData.length; i++) {
    if (i < shiftOutputData.length) { // 範囲チェック
      shiftOutputData[i][0] = shiftConfigData[i][0]; // 日付
      shiftOutputData[i][1] = shiftConfigData[i][1]; // 時間
    }
  }

  // iは行を表す。
  for (let i = 1; i < shiftConfigData.length; i++) {
    let shift_slots = (shiftConfigData[i][0] + shiftConfigData[i][1]);
    // jは列を表す。
    for (let j = 2; j < shiftConfigData[0].length; j++) {
      // シフト割り当ての前に、D列の回数が少ない順に挿入ソート
      for (let m = 2; m < nameListData.length; m++) {  // 2行目から開始
        let currentRow = nameListData[m];
        let n = m - 1;

        // D列の値を比較してソート
        while (n >= 1 && nameListData[n][3] > currentRow[3]) {  // n >= 1 で1行目を無視
          nameListData[n + 1] = nameListData[n];
          n--;
        }
        nameListData[n + 1] = currentRow;
      }

      let role_need = shiftConfigData[0][j];

      let box = [];
      let usedNames = [];

      for (let k = 1; k <= shiftConfigData[i][j]; k++) {
        for (let l = 1; l < nameListData.length; l++) {
          let name = nameListData[l][1];
          let role = nameListData[l][6].split(',');
          let busy_time = nameListData[l][7].split(',');

          if (role.includes(role_need)
            && !busy_time.includes(shift_slots)
            && !usedNames.includes(name)) {
            box.push(name);
            usedNames.push(name); // 名前を追加済みとして記録
            nameListData[l][3] = nameListData[l][3] + 1;
            break;
          } else if (l === nameListData.length - 1) {
            box.push("null")
            break;
          }
        }
        shiftOutputData[i][j] = box.join(',');
      }
    }
  }

  // 元の並び順に戻すために、A列（インデックス 0）を使用して挿入ソート
  for (let m = 2; m < nameListData.length; m++) {  // 2行目から開始
    let currentRow = nameListData[m];
    let n = m - 1;

    // A列のIDを比較して元の順序に戻す
    while (n >= 1 && nameListData[n][0] > currentRow[0]) {  // n >= 1 で1行目を無視
      nameListData[n + 1] = nameListData[n];
      n--;
    }
    nameListData[n + 1] = currentRow;
  }

  name_list.getRange(1, 1, nameListData.length, nameListData[0].length).setValues(nameListData);
  shift_output.getRange(1, 1, shiftOutputData.length, shiftOutputData[0].length).setValues(shiftOutputData);
}

// 未作成の関数
function changeShift(ss) {
  createNewShift(ss);
  //条件を満たしていない人を探してそこだけ変える
}
