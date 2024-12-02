//-----------------------------------------------------------------------------------------------------------------------------------------------
//グローバル変数の置き場
const sheetId ="" //適宜自分のシートIDに変更
 //アクティブなスレッドシートの取得
 let ss = SpreadsheetApp.getActiveSpreadsheet();
 let sheet = SpreadsheetApp.getActiveSheet();
 let templateSheet = ss.getSheetByName('template');
 let templateSheet2  =ss.getSheetByName('template2')
 //今日の日付
let today = Utilities.formatDate(new Date(),'JST','yyyy-MM');

 //LINE developerで登録したチャネルアクセストークン
 let ACCESS_TOKEN='';
 //LINEへ応答メッセージを送るAPI`
 let LINE_ENDPOINT = "https://api.line.me/v2/bot/message/reply";
 //自分のユーザーID
 let userId = "";

//---------------------------------------------------------------------------------------------------------------------------------------------- 

//----------------------------------------------------------------------------------------------------------------------------------------------
//LINEからPOSTリクエストが渡されたときに実行される処理
function doPost(e) {
  
  let json = JSON.parse(e.postData.contents);
  //LINE側へ応答するためのトークンを作成(LINEからのリクエストに入っているので、それを取得する)
  let reply_token = json.events[0]?.replyToken; //リプライトークン
  let messageText = json.events[0]?.message?.text; //メッセージ内容
  if(!reply_token || !messageText){
    return ConferenceDataService.createTextOutput(JSON.stringify({status:"no data"}));
  }

  let logsheet = SpreadsheetApp.openById(sheetId).getSheetByName('log');

  if(json.events){
    json.events.forEach(event =>{
      if(event.type === "message"){
        let userID = event.source.userId;
        let message = event.message.text;
        let timestamp = new Date(event.timestamp);

        //スプレッドシートに記録
        if(logsheet){
        logsheet.appendRow([userID,message,timestamp]);
        }

      }
    });
  }
  //リファクタリング
  const commands ={
    "登録":(params) => handlemasmainbot(reply_token,params),
    "削除":(params) => handlemasdelete(reply_token,params), 
    "新規ローン削除":(params) => handledeleteEntry3(reply_token,params),
    "新規ローン登録":(params) => handlenewLoan(reply_token,params),
    "ローン削除":(params) => handledeleteEntry2(reply_token,params),
    "ローン":(params) => handleregistrationLoan(reply_token,params),
    "予算":(prams) => handleregistrationBuget(reply_token,prams),
    "収支確認":() => checkBuget(reply_token),
    "支出のグラフ":() => chartBudget(reply_token),
    "履歴":() => checkLog(reply_token),
    "収支一覧":() => list(reply_token),
    "コマンドの説明":() => explaianComment(reply_token),
    "コマンド一覧":() => explaianComment(reply_token),
  };

  const command = Object.keys(commands).find((cmd) => messageText.startsWith(cmd));
  if(command){
    const params = messageText
    .replace(/\r\n|\r/g,"\n") //改行コードを統一
    .split("\n") //改行で分割
    .slice(1) //コマンド部分を除外
    .map((s) => s.trim());//各要素の空白を除外

    if(typeof commands[command] === "function"){
      commands[command](params); //コマンドの実行
    }else {
    console.error(`Error: Command "${command}" is not a valid function.`);
    }
     }else{
      reply(reply_token,"対応していないメッセージです\対応できるコマンドは\n削除\n登録\n複数削除\n複数登録\n収支一覧\nローン\nローン削除\n新規ローン登録\n新規ローン削除\n履歴\n収支確認\n支出のグラフ\nコマンドの説明(コマンド一覧)")
     }

  //デバック用
  return ContentService.createTextOutput(JSON.stringify({"status":"success"})).setMimeType(ContentService.MimeType.JSON);

  }
//----------------------------------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------------------------------
//呼び出し関数
function handleProcess(reply_token, params, FunctionName, actionName) {
  try {
    // 引数のチェック
    if (params.length < 2) {
      reply(reply_token, `入力形式が正しくありません。\n${actionName}\nカテゴリ\n金額\nの形式で入力してください`);
      return;
    }

    const category = [];
    const amount = [];

    // カテゴリと金額を分割
    for (let i = 0; i < params.length; i++) {
      if (i % 2 === 0) {
        category.push(params[i]);
      } else {
        amount.push(Number(params[i].trim()));
      }
    }

    // カテゴリと金額の数が一致する場合
    if (category.length === amount.length) {
    let allSuccess = [];

    for (let i = 0; i < category.length; i++) {
    const result = FunctionName(category[i], amount[i]);

    if (result) {
      allSuccess.push(true);
    } else {
      allSuccess.push(false);
    }
    }

    // すべてがtrueかを確認
    let allTrue = allSuccess.every((success) => success);

    if (allTrue) { 
      reply(reply_token, `すべての項目を正常に${actionName}しました`);
    } else {
      reply(reply_token, `一部の項目で${actionName}に失敗しました。入力内容を確認してください`);
    }

} else {
  reply(reply_token, `カテゴリーの数と金額の数が一致しませんでした。入力内容が正しいか確認してください`);
}

  } catch (error) {
    console.error(`Error occurred during ${actionName} process:`, error);
    reply(reply_token, `エラーが発生しました。管理者にお問い合わせください`);
  }
}


//--------------------------------------------------------------------------------------------------------------------------------------------


//----------------------------------------------------------------------------------------------------------------------------------------------
//handle関数
function handlemasdelete(reply_token,params){   //登録データの削除
  handleProcess(reply_token,params,deleteEntry,"削除");
  //集計 
  setValue();
}

function handledeleteEntry2(reply_token,params){   //ローンデータの削除
  handleProcess(reply_token,params,deleteEntry2,"ローンデータ削除");
  setLoanValue();
  answerLoan();
  setValue();
}

function handledeleteEntry3(reply_token,params){  //ローンの削除
  handleProcess(reply_token,params,deleteEntry3,"新規ローン削除");
}

function handlenewLoan(reply_token,params){  //新規ローンの登録
  handleProcess(reply_token,params,newLoan,"新規ローン登録");
}

function handleregistrationLoan(reply_token,params){ //ローン
  handleProcess(reply_token,params,registrationLoan,"ローンデータ");
  //ローンデータの計算
  setLoanValue();
  //ローン計算の結果を出力
  answerLoan();
  //ローンデータから集計の更新
  setValue();
}

//通常処理(登録)
function handlemasmainbot(reply_token,params){
  handleProcess(reply_token,params,mainbot,"支出の登録");
  setValue();
}

function handleregistrationBuget(reply_token,params){
  handleProcess(reply_token,params,registrationBuget,"予算の登録");
}
//-----------------------------------------------------------------------------------------------------------------

//-----------------------------------------------------------------------------------------------------------------
//返答のメイン処理　

function explaianComment(reply_token){
    let message = "マンドの説明一覧\nすべてのコマンドは\n[コマンド名\n項目\n金額]\nでコマンドを実行できます\n削除:登録した収入・支出の削除を行います\n\n収支一覧:収入・支出の一覧を表示します\n\nローン登録:新規ローン項目の登録を行います\n\nローン登録削除:登録した新規ローン項目の削除を行います\n\nローン:ローンの支払った金額の登録を行います\n\nローン削除:登録したローン金額の削除を行います\n\n支出のグラフ:支出のグラフを表示します\n\n履歴:登録した収支の一覧を参照します\n\n予算:今月の予算を作成します\n\n収支確認:現在の収支を報告します";
    reply(reply_token,message);
}


//指定されたカテゴリーと金額を確かめて削除
function deleteEntry(category,amount){
  let spreadSheet = SpreadsheetApp.openById(sheetId);
  let today = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM内訳');
  let targetSheet = spreadSheet.getSheetByName(today);

  if(!targetSheet) return false;

  let data = targetSheet.getDataRange().getValues(); //全データを取得
  for(let i = data.length-1; i>0; i--){
    if(data[i][1] === category && data[i][2] == amount){
      targetSheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

//指定されたカテゴリーと金額を確かめてローンデータ削除
function deleteEntry2(category,amount){
  let spreadSheet = SpreadsheetApp.openById(sheetId);
  let targetSheet = spreadSheet.getSheetByName("Loan");

  if(!targetSheet) return false;

  let data = targetSheet.getDataRange().getValues(); //全データを取得
  for(let i=1; i<data.length; i++){
    if(data[i][1] === category && data[i][2] == amount){
      targetSheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

//新規ローン削除
function deleteEntry3(category,amount){
  let spreadSheet = SpreadsheetApp.openById(sheetId);
  let targetSheet = spreadSheet.getSheetByName("Loan2");

  if(!targetSheet) return false;

  let data = targetSheet.getDataRange().getValues(); //全データを取得
  for(let i=1; i<data.length; i++){
    if(data[i][0] === category && data[i][1] == amount){
      targetSheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

//ローン登録
function registrationLoan(category,amount){
  
  // 対象のスプレッドシートを取得
  let spreadSheet = SpreadsheetApp.openById(sheetId);
  let  todayDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  // Laonシートを取得
  let targetSheet = spreadSheet.getSheetByName("Loan");

  if (targetSheet){

    // 最終行を取得
    var lastRow = targetSheet.getLastRow();
    var targetRow = lastRow + 1;

    // 書き込み
    targetSheet.getRange(targetRow, 1).setValue(todayDate);  // 日付
    targetSheet.getRange(targetRow, 2).setValue(category);  // 費目
    targetSheet.getRange(targetRow, 3).setValue(amount);  // 金額
    return ture;
  }
  else{
    // 対象のシートが見つからない場合
    return false;
  }
}


//ローン残高の表示処理
function answerLoan(reply_token){
  // 対象のスプレッドシートを取得
  let spreadSheet = SpreadsheetApp.openById(sheetId);

  // Laonシートを取得
  let targetSheet = spreadSheet.getSheetByName("Loan2");
  let lastRow = targetSheet.getLastRow(); 
  let rawcategories = targetSheet.getRange(2,1,lastRow-1,1).getValues().flat(); //項目の一覧を取得
  let targetCells = targetSheet.getRange(2,3,lastRow-1,1).getValues().flat(); //ローン残高を取得 
  
  targetCells = targetCells.map(cell =>{
    if(cell === 0){
      return "支払いが完了しています";
    }
    return cell;
  });

  //メッセージのための配列を生成
  let newMessage = [];

  for(let i=0; i<rawcategories.length; i++){
    newMessage.push(rawcategories[i]+":"+targetCells[i]);
  }

  let message = newMessage.join('\n');

  reply(reply_token,message);
}

//新規ローンの登録
function newLoan(category,amount){
    // 対象のスプレッドシートを取得
  let spreadSheet = SpreadsheetApp.openById(sheetId);

  // Load2のシートを取得
  let targetSheet = spreadSheet.getSheetByName("Loan2");

  if (targetSheet){

    // 最終行を取得
    var lastRow = targetSheet.getLastRow();
    var targetRow = lastRow + 1;

    // 書き込み
    targetSheet.getRange(targetRow, 1).setValue(category);  // 費目
    targetSheet.getRange(targetRow, 2).setValue(amount);  // 金額
    return true;
  }
  else{
    // 対象のシートが見つからない場合
    return false;
  }
}

function checkBuget(reply_token){
  let spreadSheet = SpreadsheetApp.openById(sheetId);


  let today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');

  //チェックする項目のあるファイルを読み取る
  let targetSheet = spreadSheet.getSheetByName(today);

  //シートが見つからないとき
  if(!targetSheet){
    return "本日用のシートが見つかりません";
  }

  let lastRow = targetSheet.getLastRow();
  let rawcategories = targetSheet.getRange(7,1,lastRow,1).getValues().flat(); //項目の一覧を取得
  let categories = rawcategories.filter((item)=>item && item !== "合計" && item !=="項目" && item !== "支出");
  let rawBugets = targetSheet.getRange(7,3,lastRow,1).getValues().flat();//予算のお金を取得
  let bugets =  rawBugets.filter((item)=> item !== "予算" && item !== "");

  let rawamounts = targetSheet.getRange(7,2,lastRow,1).getValues().flat();//実績のお金を取得
  let amounts = rawamounts.filter((item)=> item !== "実績" && item !== "");

  let finals = targetSheet.getRange(7,4,lastRow,1).getValues().flat();//差額のお金を取得
  let final = finals.filter((item)=> item !== "差額" && item !== "");

   let messageArray = [];

    for(let i=0; i<categories.length-1; i++){
      messageArray.push(`${categories[i]} : ${bugets[i]} - ${amounts[i]} = ${final[i]}`);
    }
  
     let message = messageArray.join("\n");

     let postmessage = "項目：予算-実績=差額\n"+ message; 

    reply(reply_token,postmessage);
  
}

function list(reply_token){
   // 対象のスプレッドシートを取得
  let spreadSheet = SpreadsheetApp.openById(sheetId);


  let today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');

  //チェックする項目のあるファイルを読み取る
  let targetSheet = spreadSheet.getSheetByName(today);

  //シートが見つからないとき
  if(!targetSheet){
    return "本日用のシートが見つかりません";
  }

  let lastRow = targetSheet.getLastRow();
  let rawcategories = targetSheet.getRange(7,1,lastRow,1).getValues().flat(); //項目の一覧を取得
  let categories = rawcategories.filter((item)=>item && item !== "合計" && item !=="項目");
  let message = categories.join("\n");
  
  reply(reply_token,message);
}

//収支の画像を送る
function pushImage(userId,src,srcPreview){
  let url = "https://api.line.me/v2/bot/message/push";
  let headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    "Authorization": "Bearer " + ACCESS_TOKEN,
  };
  let postData = {
    "to":userId,
    "messages":[
      {
      "type":"image",
      "originalContentUrl": src, //画像のURL
      "previewImageUrl":srcPreview, //プレビュー画像のURL
      }
    ]
  };

  let options = {
    "method" : "post",
    "headers" : headers,
    "payload":JSON.stringify(postData)
  };

   try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log("Image sent successfully: " + response.getContentText());
  } catch (error) {
    Logger.log("Failed to send image: " + error.message);
  }

}

//登録確認用関数
function checkLog(reply_token){
  // 対象のスプレッドシートを取得
  let spreadSheet = SpreadsheetApp.openById(sheetId);

  // 本日の日付を取得
  let today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM内訳');
  let targetSheet = spreadSheet.getSheetByName(today);

    // 最終行を取得
    var lastRow = targetSheet.getLastRow();

    // 書き込み
    let categories= targetSheet.getRange(2,2,lastRow,1).getValues().flat();  // 費目
    let amounts =targetSheet.getRange(2,3,lastRow,1).getValues().flat();  // 金額
    
    let messageArray = [];

    for(let i=0; i<categories.length-1; i++){
      messageArray.push(`${categories[i]}:${amounts[i]}`);
    }
  
     let message = messageArray.join("\n");

    reply(reply_token,message);
}

//----------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------
//サブ処理関数



//スプレッドシートへの書き込み
function mainbot(category,amount){

  //LINEから受け取ったメッセージの内容が形式通りかチェックする
  let validate = isValid(category,amount);
  if(validate != "OK"){
    return false;
  }

   // 対象のスプレッドシートを取得
  let spreadSheet = SpreadsheetApp.openById(sheetId);

  // 本日の日付を取得
  let today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM内訳');
  let targetSheet = spreadSheet.getSheetByName(today);
  let  todayDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  if (targetSheet){

    // 最終行を取得
    var lastRow = targetSheet.getLastRow();
    var targetRow = lastRow + 1;

    // 書き込み
    targetSheet.getRange(targetRow, 1).setValue(todayDate);  // 日付
    targetSheet.getRange(targetRow, 2).setValue(category);  // 費目
    targetSheet.getRange(targetRow, 3).setValue(amount);  // 金額
    return true;
  }
  else{
    // 対象のシートが見つからない場合
    return false;
  }

}

//予算の登録
function registrationBuget(category,amount) {

  //LINEから受け取ったメッセージの内容が形式通りかチェックする
  let validate = isValid(category,amount);
  if(validate != "OK"){
    return false;
  }

  const ss = SpreadsheetApp.openById(sheetId);


  //代入先の定数定義
  let mainSheet = ss.getSheetByName(today);
  let lastRow = mainSheet.getLastRow();
  let rawcategories = mainSheet.getRange(7,1,lastRow,1).getValues().flat(); //項目の一覧を取得
  let targetCell = mainSheet.getRange(7,3,lastRow,1).getValues().flat(); //書き込み対象の範囲を取得

  let targetLine = mainSheet.getRange(7,3,lastRow,1); //出力範囲の指定
  
  for(let i=0; i<rawcategories.length; i++){
      if(category == rawcategories[i]){
        targetCell[i] = amount;
    }
  } 

  //更新された値をスプレッドシートに書き込み
  targetLine.setValues(targetCell.map((value) => [value]));

  return true;
}

//円グラフの作成
function chartBudget(reply_token){
  let targetSheet = SpreadsheetApp.openById(sheetId).getSheetByName(today);
  if(!targetSheet){
    reply(reply_token,"対象のシートが見つかりません");
    return;
  }

  let range= targetSheet.getRange("A13:B24");
  let chart = targetSheet.newChart()
  .addRange(range)
  .setChartType(Charts.ChartType.PIE)
  .setPosition(2,6,0,0)
  .setOption('pieSliceText',"value-and-percentage")
  .setOption('legend',{position:"right"})
  .setOption('title','今月の収支')
  .build();            

  //シートをチャートを挿入
  targetSheet.insertChart(chart);

  //グラフ画像を取得
  let blog = chart.getBlob();


  //Google ドライブに保存　
  let folderId = "1Uug3qxUXljjjD13tUSLkAbLRkHtZ_3fg"; //googleドライブの一時フォルダのID
  // let graphImg = chart.getAs('image/png');
  let folder = DriveApp.getFolderById(folderId);
  let file = folder.createFile(blog);
  file.setName(today + "_chart.png");

  //公開設定する
  file.setSharing(DriveApp.Access.ANYONE,DriveApp.Permission.EDIT);
  let imageUrl = file.getDownloadUrl();
  //画像の送信
  pushImage(userId,imageUrl,imageUrl)


  // //一時ファイルを削除する
  file.setTrashed(true);

  reply(reply_token,"支出のグラフを送信しました");
}




//自動集計システム
function setValue() {
  const ss = SpreadsheetApp.openById(sheetId);

  let sub = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM内訳');
  let sheet = ss.getSheetByName(sub);

  if(!sheet) console.error("指定されたシートが見つかりません: " + sub);
  
  //行の末尾まで値を取得
  let lastCol = sheet.getLastRow();
  let values = sheet.getRange(2,2,lastCol,2).getValues();


  //代入先の定数定義
  let mainSheet = ss.getSheetByName(today);
  let lastRow = mainSheet.getLastRow();
  let rawcategories = mainSheet.getRange(7,1,lastRow,1).getValues().flat(); //項目の一覧を取得
  let targetCell = mainSheet.getRange(7,2,lastRow,1).getValues().flat();
  
  //数値が入っているセルを0に初期化
  targetCell = targetCell.map(value => (isFinite(value) && value !== "" ? 0: value));

  //出力範囲設定　
  let targetRange = mainSheet.getRange(7,2,lastRow,1);
 
   for (let i =0; i< values.length; i++){
    for(let j=0; j< rawcategories.length; j++){


        if(values[i][0] == rawcategories[j]){
        targetCell[j] += values[i][1];
      }
    }
   }  

  targetRange.setValues(targetCell.map((value) => [value]));
}



function setLoanValue(){
   const ss = SpreadsheetApp.openById(sheetId);

  let sheet = ss.getSheetByName("Loan");
  if(!sheet) console.error("指定されたシートが見つかりません: " + sub);

  //行の末尾まで値を取得
  let lastCol = sheet.getLastRow();
  let values = sheet.getRange(2,2,lastCol,2).getValues();

  //代入先の定数定義
  let mainSheet = ss.getSheetByName("Loan2");
  let lastRow = mainSheet.getLastRow(); 
  let rawcategories = mainSheet.getRange(2,1,lastRow-1,1).getValues().flat(); //項目の一覧を取得
  let targetCell = mainSheet.getRange(2,3,lastRow-1,1).getValues().flat(); //ローン残高を取得
  let loan = mainSheet.getRange(2,2,lastRow-1,1).getValues().flat();//ローン残高の元本を取得

  //出力範囲設定　
  let targetRange = mainSheet.getRange(2,3,lastRow-1,1);
 
   for (let i =0; i< values.length; i++){
    for(let j=0; j< rawcategories.length; j++){

        if(values[i][0] == rawcategories[j]){
        loan[j] -= values[i][1];
        targetCell[j] = loan[j];
      }
    }
   }

  targetRange.setValues(targetCell.map((value) => [value]));

  //ローンを収支のグラフにも反映させるために計算する
  
  let sum = values.reduce((total,row) =>{
    return total + row[1];
  },0);

  registData("ローン",sum);
}

//通常メッセージ対応　
function reply(reply_token,message){
  UrlFetchApp.fetch(LINE_ENDPOINT,{
    "headers":{
      "Content-Type":"application/json; charset=UTF-8",
      "Authorization":"Bearer " + ACCESS_TOKEN
    },
    "method":"post",
    "payload":JSON.stringify({
    "replyToken":reply_token,
    "messages":[{
      "type":"text",
      "text":message,
    }]
    })
  });

  return ContentService.createTextOutput(JSON.stringify({"content":"post ok"})).setMimeType(ContentService.MimeType.JSON);
}
//-------------------------------------------------------------------------------------------------------------------------

//------------------------------------------------------------------------------------------------------------------------
//LINE返答以外の処理

//毎日の処理
function remaind(){
  let message = "今日の収支報告を記入してください";

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast',{
    method:'post',
    headers:{
      'Content-Type':'application/json',
      'Authorization':'Bearer ' + ACCESS_TOKEN, 
    },
    "payload":JSON.stringify({
      "messages":[
        {
          type:'text',
          text:message
        },
      ]
    }),
  });
}

//月末の処理(締め切りを伝える)
function remaind2(){
  let message = "本日月末です！\n23時までに収支を確定させてください";

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast',{
    method:'post',
    headers:{
      'Content-Type':'application/json',
      'Authorization':'Bearer ' + ACCESS_TOKEN, 
    },
    "payload":JSON.stringify({
      "messages":[
        {
          type:'text',
          text:message
        },
      ]
    }),
  });
}

//月末の処理(収支の発表)
function notify(){
  let getSpredSheet = SpreadsheetApp.openById(sheetId);
  let sheetName = Utilities.formatDate(new Date(), 'JST' , 'yyyy-MM');
  let targetSheet = getSpredSheet.getSheetByName(sheetName);

  //収支合計の取得　
  let values = targetSheet.getRange("B3").getValue();
  let message = `今月の収支は${values}円でした`;

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast',{
    method:'post',
    headers:{
      'Content-Type':'application/json',
      'Authorization':'Bearer ' + ACCESS_TOKEN, 
    },
    "payload":JSON.stringify({
      "messages":[
        {
          type:'text',
          text:message
        },
      ]
    }),
  });
}

//毎月のシート作成
function createSheeet(){
  //シートが作成できなかったら報告する
  try{
    //通常処理:テンプレートのを指定
    ss.insertSheet(getName(),0,{template:templateSheet});
    ss.insertSheet(getName2(),-1,{template:templateSheet2});
  } catch(e){
    console.log('今月のシート作成が完了していません');
  }
}

//シートの名前の生成
function getName(){
  return(today);
}

function getName2(){
  let todayName = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM内訳');
  return(todayName);
}
//--------------------------------------------------------------------------------------------------------------------

// // メッセージ内容が正しいか確かめる
function isValid(category,amount){


  // 対象のスプレッドシートを取得
  let spreadSheet = SpreadsheetApp.openById(sheetId);


  let today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');

  //チェックする項目のあるファイルを読み取る
  let targetSheet = spreadSheet.getSheetByName(today);

  //シートが見つからないとき
  if(!targetSheet){
    return false;
  }

  let lastRow = targetSheet.getLastRow();
  let rawcategories = targetSheet.getRange(7,1,lastRow,1).getValues().flat(); //項目の一覧を取得
  let categories = rawcategories.filter((item)=>item && item !== "合計" && item !=="項目");

  if(!categories.includes(category)){
    return false;
  }
  if(isNaN(amount) || amount <=0 ){
    return "金額を正しく入力してください";
  }
  
  return "OK";

}