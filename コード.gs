const MY_BOOK = SpreadsheetApp.getActiveSpreadsheet();
const TEMPLATE_SHEET_NAME = "【文面FM】";
const MANUAL_SHEET_NAME = "つかいかた";

function onOpen() {
  var myMenu = [
    { name: '1つのシートに対して実行', functionName: 'allSendBySheetName' },
    { name: '全てのシートに対して実行', functionName: 'allSendAllSheet' },
    //null,
    //{name:'送信確認',functionName:'isSentCheck2'},
  ];
  MY_BOOK.addMenu('一斉送信', myMenu);
}

function allSendAllSheet() {
  const ui = SpreadsheetApp.getUi();
  res = ui.alert("送信前確認", "すべてのシートに記載されているメールアドレスすべてにメールを一斉送信します。よろしいですか？", ui.ButtonSet.OK_CANCEL);

  if (res == ui.Button.OK) {
    const sheets = MY_BOOK.getSheets();

    sheets.forEach(sheet => {
      const sheetName = sheet.getName()

      if ([TEMPLATE_SHEET_NAME, MANUAL_SHEET_NAME].includes(sheetName)) return;

      allSend(sheet);
    });

    ui.alert("一斉送信終了");
  }

}

function allSendBySheetName() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt("一斉送信したいシートの名前を記入", ui.ButtonSet.OK_CANCEL).getResponseText();
  const mySheet = MY_BOOK.getSheetByName(res);

  if (!mySheet) ui.alert(`"${res}"シートは存在しません`);

  allSend(mySheet);
  ui.alert("一斉送信終了");
}

function allSend(mySheet) {
  const dataSheet = MY_BOOK.getSheetByName(TEMPLATE_SHEET_NAME);

  let txtData = dataSheet.getDataRange().getValues();
  //配列txtDataに、【文面FM】シートの全てのデータを格納
  let txt = txtData[6][1];
  //txtに文面のFMが入った
  let regex = /\\\d+\\/g;
  let result = txt.match(regex);
  let replaceWord = /\\/g;
  let cnt_replace = result.length;
  let replace_c = [];
  for (let i = 1; i <= cnt_replace; i++) {
    replace_c.push(Number(result[i - 1].replace(replaceWord, "")))
  }
  let dataRange = mySheet.getDataRange();
  let myTxt;
  let myCol;
  let txtArr = []
  let myData = dataRange.getValues();
  mySheet.getRange(2, 2, myData.length - 1, 2).removeCheckboxes();
  mySheet.getRange(2, 2, myData.length - 1, 2).insertCheckboxes();
  myData = dataRange.getValues();
  for (content of myData) {
    if (content[3] != "" && content[0] == true) {
      content[1] = true;
      myTxt = txt;
      for (idx in result) {
        myCol = replace_c[idx] + 5;
        //ここ列指定
        /*　//ここのくだりは空欄時エラーを発生させる。
        if(content[myCol] == ""){
          content[1] = false;
          break;
        }
        */
        myTxt = myTxt.replace(result[idx], content[myCol]);
      }
      txtArr.push(myTxt);
    } else {
      txtArr.push(null)
    }
  }
  let options = {};
  if (txtData[1][1] != "") {
    options['name'] = txtData[1][1];
  }
  let attachmentsHTML;
  if (txtData[8][1] != "") {
    attachmentsHTML = makeAttachmentsHTML(txtData);
  }
  let address;
  let subject = txtData[4][1];
  let body;
  for (idx in myData) {
    content = myData[idx];
    if (content[3] != "" && content[0] == true && content[1] == true) {
      address = content[3];
      var mycc = content[4];
      if (mycc != "") {//ccあり
        options['cc'] = mycc;
      } else {         //ccなし
        delete options['cc'];
      }
      var mybcc = content[5];
      if (mybcc != "") {//bccあり
        options['bcc'] = mybcc;
      } else {          //bccなし
        delete options['bcc'];

      }
      body = txtArr[idx];
      if (attachmentsHTML) {
        options['htmlBody'] = body.replaceAll("\n", "<br>") + '<br>' + attachmentsHTML;
      }
      if (txtData[2][1] == true) {//プレビューのみオン
        GmailApp.createDraft(address, subject, body, options);
        content[2] = false;
      } else {
        GmailApp.sendEmail(address, subject, body, options);
      }
    }
  }
  if (txtData[2][1] != true) {
    Utilities.sleep(10 * 1000);

    myData = isSentCheck2(myData);
  }
  dataRange.setValues(myData);

  Logger.log(MailApp.getRemainingDailyQuota());
}

function makeAttachmentsHTML(txtData) {
  let last = txtData.length;
  let myHTML = ""
  for (var i = 8; i < last; i++) {
    var fileID = txtData[i][1];
    if (fileID != "") {
      var myFile = DriveApp.getFileById(txtData[i][1]);
      var fileURL = myFile.getUrl();
      var fileName = myFile.getName();
      var fileType = myFile.getMimeType();
      myHTML +=
        '<br><div dir="ltr">' +
        '<div contenteditable="false" class="gmail_chip gmail_drive_chip" style="width:396px;height:18px;max-height:18px;background-color:rgb(245,245,245);padding:5px;font-family:arial;font-weight:bold;font-size:13px;border:1px solid rgb(221,221,221);line-height:1">' +
        '<a href="' + fileURL + '" target="_blank" style="display:inline-block;max-width:366px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-decoration-line:none;padding:1px 0px;border:none" aria-label="' + fileName + '">' +
        '<img style="vertical-align: bottom; border: none;" src="https://drive-thirdparty.googleusercontent.com/16/type/'
        + fileType + '">' +
        '<span dir="ltr" style="vertical-align:bottom;text-decoration:none">' +
        fileName +
        '</span>' +
        '</a>' +
        '<img src="//ssl.gstatic.com/ui/v1/icons/common/x_8px.png" style="opacity: 0.55; cursor: pointer; float: right; position: relative; top: -1px; display: none;">' +
        '</div>' +
        '</div>';
    }
  }
  return myHTML;
}

function isSentCheck2(myData) {
  if (!myData) {//単独実行対応
    const mySheet = MY_BOOK.getSheetByName("送信画面");
    myData = mySheet.getDataRange().getValues();
  }
  let checkArr = [];
  for (i in myData) {
    if (myData[i][0] == true && myData[i][1] == true) {
      checkArr.push([i, null, myData[i][3]]);
    }
  }
  /*配列checkArrに、送信にチェックが入っており、かつ内容に不備がない行の情報を取得
  **各要素の詳細↓
  [  0  , 1  ,  2  ]
  [行番号,null,メアド]
          ↑ True or Falseを格納するため、一旦nullを格納
  */
  if (checkArr.length > 0) {
    checkArr = checkArr.reverse();
    //配列の並び順を逆順にした
    let threads = GmailApp.search('in:sent', 0, checkArr.length);
    if (threads) {
      for (content of checkArr) {
        for (i in threads) {
          var messages = threads[i].getMessages();
          var address = messages[0].getTo();
          if (address == content[2]) {
            content[1] = true;
            if (messages.length > 1) {
              var message = messages[1];
              var sender = message.getFrom();
              if (sender.includes('mailer-daemon@googlemail.com')) {
                var body = message.getPlainBody();
                if ((body.includes('アドレス不明') &&
                  body.includes('配信されませんでした')) ||
                  body.includes('メールはブロックされました。')) {
                  content[1] = false;
                }
              }
            }
            threads.splice(i, 1);
            break;
          }
        }
      }
    }
    for (i in checkArr) {
      myData[checkArr[i][0]][2] = checkArr[i][1]
    }
  }
  return myData;
}

function myFunc() {
  let threads = GmailApp.search('in:sent', 0, 3);
  for (i in threads) {
    var messages = threads[i].getMessages();
    for (j in messages) {
      var message = messages[j];
      console.log(message.getTo());
    }
  }
}

function myFunc2() {
  let myArr = [];
  Logger.log(myArr.length);
}


