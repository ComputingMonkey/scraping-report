const sheet = SpreadsheetApp.getActiveSheet();
var kanekosanID = '<@UTYEPRTN2>';
var kanekosanID = '<@UL0TX44FP>';//+0あとでmyIDに変更
//incoming webhookのURL
const postUrl = 'https://hooks.slack.com/services/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';
const username = '完了報告bot';  // 通知時に表示されるユーザー名
const icon = ':bear:';  // 通知時に表示されるアイコン

function test(){
  SpreadsheetApp.getActiveSheet().getRange(1,1).setValue('=reportDone()');
}
function reportDone(message) {
  //slackに送る
  var jsonData =
  {
     "username" : username,
     "icon_emoji": icon,
     "text" : message
  };
  var payload = JSON.stringify(jsonData);
  var options =
  {
    "method" : "post",
    "contentType" : "application/json",
    "payload" : payload
  };
  UrlFetchApp.fetch(postUrl, options);
  return '報告完了';
}

function putMsg(){
  //入力するメッセージボックス
  var result = Browser.inputBox("完了報告プログラム", "完了したURLを入力して下さい\n※リストにあるURLを入力して下さい", Browser.Buttons.OK_CANCEL);
  if (result != "cancel"){
    //完了した企業とワーカーをURLから検索
    const doneRow = getKeyRow(sheet,result,4+0);
    const company = sheet.getRange(doneRow,3+0).getValue();
    const worker = sheet.getRange(doneRow,1+0).getValue();
    //完了日時をセット
    sheet.getRange(doneRow, 6+0).setValue(getNow());
    //メッセージを作成し、slackに送信
    const message = worker + 'が' + company + 'のスクレイピングを終了しました\n' + kanekosanID + 'さんは確認をお願い致します';
    reportDone(message);
    //完了した企業をslackに報告
    sheet.getRange(doneRow,5+0).setValue(reportDone(message))
    Browser.msgBox(worker + 'が' + company + 'について完了報告しました');
  }else{
    Browser.msgBox('報告がキャンセルされました');
  }  
}

//#シート、キーワード、キーワードを探したい列の番号:キーワードの行;n列目の特定のキーワードがある行を取得
function getKeyRow(sheet,key,keyCol) {
  //n列目の配列を取得
  var array2 = sheet.getRange(1,keyCol,sheet.getLastRow()).getValues();
  var array = Array.prototype.concat.apply([],array2);
  //キーワードがある列を検索
  var key_row = array.indexOf(key) + 1;
  //列の番号を返す
  return key_row;
}
//現在日時
function getNow() {
  var d = new Date();
  var y = d.getFullYear();
  var mon = d.getMonth() + 1;
  var d2 = d.getDate();
  var h = d.getHours();
  var min = d.getMinutes();
  var s = d.getSeconds();
  var now = y+"/"+mon+"/"+d2+" "+h+":"+min+":"+s;
  return now;
}