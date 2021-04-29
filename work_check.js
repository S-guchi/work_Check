const ssIdObj = SpreadsheetApp.openById("スプレッドシートのId");
const chkSsObj = ssIdObj.getSheetByName("作業チェックシート");
const tejunSsObj = ssIdObj.getSheetByName("作業手順書");
const historySsObj = ssIdObj.getSheetByName("作業履歴");

function onChange(e) {
  const sheetName = e.source.getActiveSheet().getName();
  switch (sheetName) {
    case '作業チェックシート':
      WorkCheck(e)
      break;
  }
}


class ChkWorkProp {
  constructor() {
    this.startRow = 10;//開始行
    this.endRow = 110;//最終行
    this.numRows = 101;//行数
  }
}

const WorkCheck = e => {
  const cellPosition = e.source.getActiveRange().getA1Notation();
  if (cellPosition === 'K6') {
    const boolean = ChkTool();
    if (boolean) {
      //手順書を読み込み
      TejunLoad(e.source.getActiveRange().getValue())
    } else {
      Browser.msgBox("使用中です");
    }
  }
}

const TejunLoad = tejunName => {
  class ChkWorkPropLoad extends ChkWorkProp {
    constructor() {
      super();
      this.tejunValues = tejunSsObj.getRange(4, 1, tejunSsObj.getLastRow(), 4).getValues()
      this.list = this.MakeListAry();
    }
    // 手順書リスト作成処理
    MakeListAry() {
      const ary = []
      this.tejunValues.forEach(function (value) {
        if (value[1].indexOf(tejunName) !== -1) {
          ary.push([value[2], value[3]]);
        }
      })
      return ary
    }
  }
  const p = new ChkWorkPropLoad();
  // 行の非表示を解除
  chkSsObj.showRows(p.startRow, p.numRows)
  // F列とG列のチェックボックスを空にする
  chkSsObj.getRange(`F${p.startRow}:G${p.endRow}`).setValue("False")
  // 作業手順から抜き出したデータを張り付ける
  chkSsObj.getRange(10, 4, p.list.length, 2).setValues(p.list)
  // 余分な行を非表示にする
  chkSsObj.hideRows(p.startRow + p.list.length, p.numRows - p.list.length)
  // 非表示行のチェックボックスをTrueにする。
  chkSsObj.getRange(`F${p.startRow + p.list.length}:G${p.endRow}`).setValue("True")
  // クリア処理のために値を確保する
  chkSsObj.getRange("K2").setValue(p.list.length)
}

const ChkBox = () => {
  const p = new ChkWorkProp
  const status = chkSsObj.getRange(`F${p.startRow}:G${chkSsObj.getRange("K2").getValue() + p.startRow - 1}`).getValues().flat();
  // チェックボックスのチェック
  if (status.indexOf(false) === -1) {
    ResultPaste();
  } else {
    Browser.msgBox("NG", "チェックされていない項目があります。確認してください。", Browser.Buttons.OK);
  }
}
// チェックボックスにチェックが入っていたらfalseを返す
const ChkTool = () => {
  const p = new ChkWorkProp
  const status = chkSsObj.getRange(`F${p.startRow}:G${chkSsObj.getRange("K2").getValue() + p.startRow - 1}`).getValues().flat();
  if (status.indexOf(true) > -1) {
    return false;
  } else {
    return true;
  }
}
//作業結果貼り付け
const ResultPaste = () => {
  const lastRow = historySsObj.getLastRow();
  const pasteList = [[
    historySsObj.getRange("J2").getValue() + 1,//履歴の行数
    chkSsObj.getRange("k3").getValue(),        //管理番号
    new Date(),                                //作業日時
    chkSsObj.getRange("k7").getValue(),        //開始時間
    chkSsObj.getRange("k115").getValue(),      //終了時間
    chkSsObj.getRange("k4").getValue(),        //作業実施社名
    chkSsObj.getRange("k5").getValue(),        //作業確認者名
    chkSsObj.getRange("I6").getValue(),        //手順署名
  ]]
  if (Browser.msgBox("OK", "チェックOK。履歴に登録します。", Browser.Buttons.OK_CANCEL) == "ok") {
    historySsObj.getRange(lastRow + 1, 1, 1, 8).setValues(pasteList)
    chkclear();
  }
}
// チェックボックスクリア
const chkclear = () => {
  const p = new ChkWorkProp();
  chkSsObj.getRange(`F${p.startRow}:G${p.endRow}`).setValue("False")
  chkSsObj.getRange(`F${chkSsObj.getRange("K2").getValue() + p.startRow}:G${p.endRow}`).setValue("True")
}
//開始時刻ボタン
const getTime1 = () => chkSsObj.getRange("K7").setValue(`${new Date().getHours()}:${new Date().getMinutes()}`);
//終了時刻ボタン
const getTime2 = () => chkSsObj.getRange("K115").setValue(`${new Date().getHours()}:${new Date().getMinutes()}`);
