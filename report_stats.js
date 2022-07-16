var access_token = 'O3hFzlYEB9QhQS0DrdP10Pq1xaHmKEatBR9kX9/C7tVLwXYze5ZbUC50r6zrlXo3twju9semK+Bl8yxHmbvBhBxdH2lyfOv+O0/0xQEdljYFNCAzUOh+pQtkgm9lhha0S6cjddtmaG3eyJQRewuoPAdB04t89/1O/w1cDnyilFU=';

function getSpreadSheet() {
  var sheet = SpreadsheetApp.openById('1jCi7Pt5aPLLZARUl1K7V_9TJKNUM3wSmgpuT8F4UfTo').getSheetByName ('pine_planning');
  //var range = sheet.getRange().getValues();
  //console.log(range)
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2,1,lastRow -1,4).getValues();
  //console.log(range);
  var latestData = range[lastRow-2];
  var prevData = range[lastRow -3];
  var weekData = range[lastRow -8];
  //console.log(latestData)
  //console.log(prevData)
  var img_url  = charData()
  assembleData(latestData,prevData,weekData,img_url)
}

function assembleData(data_list,prev_data_list,week_data_list,img_u){
  //daily
  // flwr = data_list[1] - prev_data_list[1]
  // flw = data_list[2] - prev_data_list[2]
  // pst = data_list[3] - prev_data_list[3]
  //weekly
  flwr_w = data_list[1] - week_data_list[1]
  flw_w = data_list[2] - week_data_list[2]
  pst_w = data_list[3] - week_data_list[3]

  text = '■週間レポートです。\n'
  // text = text + 'Follwer :'+ data_list[1]+ ' 前日比:'+flwr + '\n';
  // text = text + 'Follow :'+data_list[2]+ '　前日比:'+flw + '\n';
  // text = text + 'Posts'+data_list[3]+ '　前日比:'+pst + '\n';
  text = text + 'Follwer :'+ data_list[1]+ ' 週間比:'+flwr_w + '\n';
  text = text + 'Follow :'+data_list[2]+ '　週間比:'+flw_w + '\n';
  text = text + 'Posts'+data_list[3]+ '　週間比:'+pst_w + '\n';
  text = text + '■コメント\n'
  if(pst_w <= 1){
    var com = commentGenerator("NOTGOOD")
    text = text + '１週間投稿が増えていないようです。\n' + com
  }else if (pst_w <= 2){
    var com = commentGenerator("GOOD")
    text = text + '１週間で２投稿以上増えています。\n' + com
  }else if (pst_w <= 3){
    var com = commentGenerator("VERYGOOD")
    text = text + '３投稿以上増えています。\n' + com
  }
  //console.log(prev_data_list)
  console.log(text)
  sendReport(text,img_u)
}

function sendReport(txt,img) {
  var to = 'C085ca04eef519354cedda16372aa616f'
  var url = "https://api.line.me/v2/bot/message/push";
  var headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    'Authorization': 'Bearer ' + access_token,
  };

    var postData = {
    "to" : to,
    "messages" : [
      {
        'type':'text',
        'text':txt,
      },
      {
        'type':'image',
        'originalContentUrl':img,
        'previewImageUrl':img,
      }
    ]
  };

  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };
  return UrlFetchApp.fetch(url, options);
};

function commentGenerator(flg_act){
  var flg = flg_act
  var comments_normal = ["もう少し工夫をしましょう","自発的に改善すると楽しい作業になりますよ","先週は忙しかったですか？","見直しも大事です","フォロワーとの交流はできていますか？","伝えたい相手を意識した投稿ができていますか？","要改善！"]
  var comments_good = ["改善の兆しが見えてきました！","この調子で進めましょう！","無理せず、ただ淡々と継続しましょう！"]
  var comments_verygood = ["めっちゃ頑張りましたね！","調子が取り戻せてきましたね！","よく頑張りました！今週もお疲れ様です！"]
  var comment = ""
  if(flg=="NOTGOOD"){
    var num_result = randomInt(comments_normal) - 1
    comment = comments_normal[num_result]
  } else if(flg == "GOOD"){
    var num_result = randomInt(comments_good) - 1
    comment = comments_good[num_result]
  } else if(flg == "VERYGOOD"){
    var num_result = randomInt(comments_verygood) - 1
    comment = comments_verygood[num_result]
  }
  return comment
}

function randomInt(comment_list){
  var num_target = comment_list.length
  var num = Math.random()
  var num_c = Math.floor(num * num_target)+1
  console.log(num_c)
  return num_c
}

function charData(){
  //シート名をして指定してシートを取得します。今回の場合は「graph」シート
  var mySheet = SpreadsheetApp.openById('1jCi7Pt5aPLLZARUl1K7V_9TJKNUM3wSmgpuT8F4UfTo').getSheetByName("Graph");
  var today_f = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'YYYY-MM-dd');
  //getChartsメソッドでシート内のチャートを取得します。配列で取得されます。
  var charts  = mySheet.getCharts();
  
  //charts配列に格納されたデータから0番目のグラフを画像として取得
  var imageBlob = charts[0].getBlob().getAs('image/png').setName("chart_"+today_f+".png");
  //フォルダIDを指定して、フォルダを取得
  var folder = DriveApp.getFolderById('1mYR-oBwIyjhggFXbzMf_xUzKX7RZty82');
  
  //フォルダにcreateFileメソッドを実行して、ファイルを作成
  var file = folder.createFile(imageBlob);
    // 公開設定する
  file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT)
  return file.getDownloadUrl()
  //DriveApp.getFolderById(folderId).removeFile(file)
}
