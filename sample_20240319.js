// メニュー作成：ファイルOPEN時 
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('マクロ');
  menu.addItem('Yahoo!ニュースを取得', 'get_data1');
  menu.addToUi();
}


// ニュースを取得してシートを更新する
function get_data1(){

  var date = new Date();
  var dateStamp = Utilities.formatDate(date,"JST","yyyy-MM-dd-HH:mm:ss");

  var url = "https://news.yahoo.co.jp/topics/top-picks";
  var html = UrlFetchApp.fetch(url).getContentText();

  // ライブラリ：Paser
  // <li class="newsFeed_item ～ </li> に挟まれたテキストを（複数）抽出
  var contentsList = Parser.data(html).from('<li class="newsFeed_item').to('</li>').iterate();

  var result = []; // 結果格納用
  
  for(var s in contentsList){
    var content = contentsList[s];
    var id = "", title = "", link = "", imgLink = ""; 
    var newsMatch = content.match(/content_id:([a-z0-9].+?);/);
    if(newsMatch){
      id = newsMatch[1];
    }
    var titleMatch = content.match(/<div\sclass\="newsFeed_item_title">(.+?)<\/div>/);
    if(titleMatch){
      title = titleMatch[1];
    }
    var linkMatch = content.match(/<a\sclass\="newsFeed_item_link"\shref\="(.+?)"/);
    if(linkMatch){
      link = linkMatch[1];
    }
    var imgMatch = content.match(/<img\sloading\="lazy"\ssrc\="(.+?)"/);
    if(imgMatch){
      imgLink = imgMatch[1];
    }

    result.push([id,title,link,imgLink,dateStamp]);
  }

  // このスプレッドシート
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // シート名 news_list を取得
  try{
    var sh = ss.getSheetByName("news_list");

    // 過去に取得した表がある場合は取得して結合
    if(sh.getLastRow() > 1){ // ヘッダーを除く2行目以降
      var list = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();

      result = result.concat(list);
    }

  }catch(e){
    // news_listが無い場合は作成
    Logger.log(e.message);
    var sh = ss.insertSheet();
    sh.setName("news_list");
  }

  // 時刻とIDの降順でソート
  result.sort(function(a,b){
    if(a[4]>b[4]) return -1;
    if(a[4]<b[4]) return 1;
    if(a[0]>b[0]) return -1;
    if(a[0]<b[0]) return 1;
    return 0;
  });

  // 空の表を作成、ヘッダー設定
  var output = [["ID","TITLE","LINK","IMG_LINK","TIME_STAMP"]];

  // ID重複を除く ID同じ場合は新しいほうが格納される
  for(var a in result){

    // outputの最終行と重複していなければ格納
    if(output[output.length-1][0] !== result[a][0]){
      output.push(result[a]);
    }
    
  }

  // シートの値をクリア
  sh.clearContents();
  // 反映
  sh.getRange(1,1,output.length,output[0].length).setValues(output);

}

