
/**
 * @fileOverview Google Hangouts Chatのマニュアル検索botです。
 * @author hinohara
 * 本コードはUnderscore for GASライブラリを使用しています
 * プロジェクトキー：M3i7wmUA_5n0NSEaa6NnNqOBao7QLBR4j
 */

/**
 * ユーザの発言内容を取得
 *
 * @param {Object} event チャットイベントのオブジェクト
 */
function onMessage(event) {
  try{
    Logger.log("発言を受け取りました。");
    var message = getManualURL(event.message.text);;
    Logger.log(message);
    return { "text": message};
  }catch(e){
    Logger.log("error:onMessage");
  }
}

/**
 * スプレッドシート取得、発言内容を検索
 *
 * @param {string} target 取得した発言内容
 * @return {string} message 一致したtargetに対応する情報
 */
function getManualURL(target) {
  try{
    var _ = Underscore.load();
//    var target = "#0079";

    var url = "https://docs.google.com/spreadsheets/d/14gF_mGaMyxqxowmRM9LTySC4jmbeDqpjdXLIOX6CLxY/edit?usp=sharing";
    var spreadsheet = SpreadsheetApp.openByUrl(url).getActiveSheet();

    // (2行,1列)から109行まで1列分を取得する ID,タイトルとURL,キーワード
    var fileIdValues = spreadsheet.getRange(2,1,109,1).getValues();
    var fileInfoValues = spreadsheet.getRange(2,2,109,2).getValues();
    var fileKeyWordsValues = spreadsheet.getRange(2,4,109,1).getValues();
    Logger.log("{終了}:スプレッドシート取得処理");
    
    // メッセージ作成用配列
    var resultMessage = [];
    var resultStrMsg = "";
    var tagsAr = [];
    var namesAr = [];

    // ●全件表示の場合
    if(target === "全件"){
      // 検索ワードが全件だったら全件出力する
      resultMessage.push("登録されている全件を表示します。");
      for(var l =0; l<109; l++){
        resultMessage.push(fileIdValues[l]);
        resultMessage.push(fileInfoValues[l]);
      }
      resultStrMsg = resultMessage.join("\n").toString();
    
    // ●ID検索の場合
    }else if(target.lastIndexOf("#", 0) === 0){
      // 検索ワードの先頭が#だったらID検索に入る
      var idResultsInfo =[];
      for(var o=0; o<109; o++){
        if(_.indexOf(fileIdValues[o], target) !== -1){
          resultMessage.push(fileInfoValues[o][0]);
          resultMessage.push(fileInfoValues[o][1]);
      }
    }
      resultStrMsg = resultMessage.join("\n").toString();
      
    // ●キーワード、ファイル名検索の場合
    }else{
      // ファイル名だけ配列に再格納
      for(var n=0; n<109; n++){
        namesAr[n] = fileInfoValues[n][0];
      }
      // 部分一致したインデックスと対応する情報を結果用の配列に格納
      var fileNameResultsId = [];
      var fileNameResultsInfo =[];
      for(var i=0; i<109; i++){
        if(_.indexOf(namesAr[i], target) !== -1){
          fileNameResultsId.push(fileIdValues[i][0]);
          fileNameResultsInfo.push(fileInfoValues[i][0]);
          fileNameResultsInfo.push(fileInfoValues[i][1]);
        }
      }
      
      // キーワードを配列に分割
      for(var m=0; m<109; m++){
        tagsAr[m] = fileKeyWordsValues[m][0].split(",");
      }
      // 部分一致したインデックスと対応する情報を結果用の配列に格納
      var keyWordsResultsId = [];
      var keyWordsResultsInfo =[];
      
      // AND検索用文字列分割
      var targetItems = [];
      targetItems = String(target).split(' ');
      
      // 検索文字列が複数あった場合
      if(targetItems.length > 1) {
        var countMax = 0;
        var countList = [];
        
        for(var j=0; j<109; j++){
          var count = 0;
          // keyword配列の中に検索文字が含まれていた場合
          for(var l=0; l<targetItems.length; l++){
            if(_.indexOf(tagsAr[j], targetItems[l]) !== -1){
              count+=1;
              // 最大ヒット数を格納
              if(count > countMax)
                countMax = count;
              }
            }
          countList[j] = count;
        }
        for(var m=0; m<109; m++){
          // 最大ヒット数に該当したItemを格納
          if(countList[m] === countMax){
            keyWordsResultsId.push(fileIdValues[m][0]);
            keyWordsResultsInfo.push(fileInfoValues[m][0]);
            keyWordsResultsInfo.push(fileInfoValues[m][1]);
          }
        }
      } else {
        // 部分一致
        for(var j=0; j<109; j++){
          if(_.indexOf(tagsAr[j], target) !== -1){
            keyWordsResultsId.push(fileIdValues[j][0]);
            keyWordsResultsInfo.push(fileInfoValues[j][0]);
            keyWordsResultsInfo.push(fileInfoValues[j][1]);
          }
        }
      }
      // 検索結果結合と重複削除処理
      var idResults = _.union(fileNameResultsId, keyWordsResultsId);     
      var keywordsResults = _.union(fileNameResultsInfo, keyWordsResultsInfo);
    
      // 検索結果が10件以上の場合は件数、名前、IDを返す
      var h = 0;
      if(idResults.length > 10){
        resultMessage.push("ヒット件数が10件を超えています。「#ID」指定で絞り込みが可能です。");
        resultMessage.push("検索ワード:"+ target +" ヒット件数:"+ idResults.length +"件");
        for(var k =0; k<10; k++){
          resultMessage.push(idResults[k]);
          resultMessage.push(keywordsResults[h]);
          h = h+2;
        }// 各要素で改行させる文字列に変換
        resultStrMsg = resultMessage.join("\n").toString();
        
        // 検索結果が10件未満の場合は名前、URLを返す
      }else if(idResults.length > 0){
        resultStrMsg = keywordsResults.join("\n").toString();
        
        // 検索結果が0件の場合
      }else{
        resultStrMsg = "該当項目がありません。";
      }
    }
    return resultStrMsg;
  }catch(e){
    Logger.log("error:getManualURL"); 
  }
};
