function myFunction() {
  var ids = []; //formの結果を表示したSpreadSheetのIDを格納
  const num = ids.length; //シートの枚数
  
  var doc = DocumentApp.openById( ); //書き込みたいdocumentのIDをいれる
  
  var body = doc.getBody();
    
  var header = body.appendParagraph("3S 授業アンケート");
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  for(var i = 0;i < num;i++){
    var spreadSheet = SpreadsheetApp.openById(ids[i]);
    var sheet = spreadSheet.getActiveSheet();
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
  
    var subject = sheet.getSheetValues(1, 1, 15, 15);
    
    if(i == 1)Logger.log(subject);
    
    var paragraph;
    
    for(var j = 1;j < 15;j++){
      if(j%2){
        // Append a section header paragraph.
        paragraph = body.appendParagraph(subject[0][j]).setHeading(DocumentApp.ParagraphHeading.HEADING2);
        
        for(var p = 1;p < lastCol;p++){
          if(subject[p][j]){
            paragraph = body.appendParagraph(subject[p][j]).setHeading(DocumentApp.ParagraphHeading.HEADING4);
          }
        }
      }else{
        var scores = Charts.newDataTable();
        
        scores.addColumn(Charts.ColumnType.NUMBER, "score");
        scores.addColumn(Charts.ColumnType.NUMBER, "number");
        var score = [];
        var flag = 0;
        for(var p = 0;p < 6;p++){
          score.push(0);
        }
        
        
        for(var p = 1;p < lastCol;p++){
          if(subject[p][j]){
            score[subject[p][j]]++;
            flag = 1;
          }
        }
        for(var q = 1;q < score.length;q++){
          //if(score[q]){
            var row = [];
            row.push(Number(q));
            row.push(Number(score[q]));
            scores.addRow(row);
          //}
        }
        if(flag){
          var data = scores.build();
        
        
          paragraph = body.appendParagraph("\n").setHeading(DocumentApp.ParagraphHeading.HEADING4);
        
          var childLastIndex = doc.getBody().getNumChildren()-1;
          // 現在のドキュメントに存在する、一番最後の子要素の情報を取得します
          var lastChild = doc.getBody().getChild(childLastIndex); 
          // カーソルとして設定するポジション情報を、子要素の情報をもとに保持します。
          //var position = doc.newPosition(lastChild, 1);
          var position = doc.newPosition(paragraph.getChild(0), 0);
          // 現在のドキュメントのカーソル位置を、最後の子要素の位置に設定します。
          doc.setCursor(position);
        
       
          var image = makeChart(data);
          var cursor = doc.getCursor();
          if (cursor) {
            cursor.insertInlineImage(image);
          } else {
            body.insertImage(0, image);
          }

        }else{
          paragraph = body.appendParagraph("なし").setHeading(DocumentApp.ParagraphHeading.HEADING4);
        }
        

      }
    }

    
  }
  
  saveAsPdf(doc);
 
}

function makeChart(data) {
  var chart = Charts.newBarChart()
       .setDataTable(data)
       .setRange(1, 5)
       .setOption("legend","none")
       .setOption("vAxis", {minValue:0,maxValue:5,gridlined:{count:1},title:"受けてよかったか"})
       .setOption("hAxis", {minValue:0,maxValue:7,gridlined:{count:1},title:"人数"})
       .setOption("lineWidth",4)
       //.setColors(["#7FFFD4","#BA55D3"])
       .build();
  
  var image = chart.getBlob();
  
  return image;
 
 // saveGraph(chart);
}

function saveGraph(chart){
  var folderId =  ;  //保存したいドライブのフォルダのIDをいれる
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'YYYY-MM-dd');
  try {
    var graphImg = chart.getBlob(); // グラフを画像に変換
    var folder = DriveApp.getFolderById(folderId);
    folder.createFile(graphImg).setName(today);
  } catch (e) {
    Logger.log(e);
  }
}

function saveAsPdf(doc){
  var docId =  ;    //書き込みたいdocumentのIDをいれる
  var folderId =  ; //保存したいドライブのフォルダのIDをいれる
  
  var file = DriveApp.getFileById(docId);
  
  var newfile = file.getAs(MimeType.PDF);
  
  try {
    var folder = DriveApp.getFolderById(folderId);
    folder.createFile(newfile);
  } catch (e) {
    Logger.log(e);
  }
  
  
  
}
