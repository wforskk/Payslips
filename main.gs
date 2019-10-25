function myFunction(){
  //出勤情報
  var workDayFile = SpreadsheetApp.openById('hogehoge');
  var workDaySheet = workDayFile.getSheetByName('シート名');
  var workDayListCount = workDaySheet.getLastRow();
  //算出元データの取得
  var workDayData = workDaySheet.getRange(1, 1, workDayListCount, 7).getValues();
  
  //経費情報
  var costsFile = SpreadsheetApp.openById('hogehoge');
  var costsSheet = costsFile.getSheetByName('シート名');
  var costsListCount = costsSheet.getLastRow(); 
  //算出元データの取得
  var costsData = costsSheet.getRange(1, 1, costsListCount, 7).getValues();

  //給料明細一覧情報(出力先シート)
  var payslipFile = SpreadsheetApp.openById('hogehoge');
  var payslipSheet = payslipFile.getSheetByName('シート名');
  var payslipDaySheet = payslipFile.getSheetByName('シート名');
  //プラス１は値がない時、ヘッダーを削除させないための対策
  var payslipListCount = payslipSheet.getLastRow()+ 1;
  
  //当月取得
  var excutionFunction = new Date().getDate() + 1;
  
  //if(１ === 1){
  
  //一覧シートの初期化
  payslipSheet.getRange("A2:E"+payslipListCount).clear();
  payslipListCount = 2;
  
  //出勤報告から明細を算出する
  for(var i=1; i < workDayData.length; i++){ 
    //当月取得
    var nowMonth = Utilities.formatDate(new Date(), "JST", "MM");
    //出勤報告の遅延分を回収するための値取得
    var workDayDelay = workDayData[i][5] + workDayData[i][6];
    //出勤月取得
    var workDayMonth = workDayData[i][2];
    var checkWorkDayMonth = Utilities.formatDate(workDayMonth, "JST", "MM");
    //手動用
    var testMonth = payslipDaySheet.getRange("A1").getValue();
    nowMonth = Utilities.formatDate(testMonth, "JST", "MM");
    if(nowMonth == checkWorkDayMonth || workDayDelay == "遅延"){
    
      //日付
      var valueInput = workDayData[i][2];
      var valueOutput = payslipFile.getRange("A"+payslipListCount).setValue(valueInput);
      //支給名
      var valueInput = workDayData[i][1];
      var valueOutput = payslipFile.getRange("B" + payslipListCount).setValue(valueInput + "講師業務");
      //料金
      var valueOutput = payslipFile.getRange("C" + payslipListCount).setValue("3000");
      //名前
      var valueInput = workDayData[i][3];
      var valueOutput = payslipFile.getRange("D" + payslipListCount).setValue(valueInput);
      
      //出力先シートのレコードを次にする
      payslipListCount = payslipListCount + 1;
      
      //対応区分に対応済を設定する
      //workDayData[i][6] = "対応済";
      var count = i + 1;
      workDaySheet.getRange("G" + count).setValue("対応済");
    }
  }
  
  //経費申請から明細を算出する
  for(var i=1; i < costsData.length; i++){ 
    //経費申請月取得
    var costsMonth = costsData[i][2];
    var checkCostsMonth = Utilities.formatDate(costsMonth, "JST", "MM");
    if(nowMonth == checkCostsMonth){
      //日付
      var valueInput = costsData[i][2];
      var valueOutput = payslipFile.getRange("A" + payslipListCount).setValue(valueInput);
      //料金名
      var valueInput = costsData[i][4];
      //支給名
      var valueInput2 = costsData[i][1];
      var valueOutput = payslipFile.getRange("B" + payslipListCount).setValue(valueInput2 + "_" + valueInput);
      //料金
      var valueInput = costsData[i][5];
      var valueOutput = payslipFile.getRange("C" + payslipListCount).setValue(valueInput);
      //名前
      var valueInput = costsData[i][3];
      var valueOutput = payslipFile.getRange("D" + payslipListCount).setValue(valueInput);
      
      //出力先シートのレコードを次にする
      payslipListCount = payslipListCount + 1;
    }
  }
  //明細データ取得(low,col,low,col)→二次元配列
  var doubleCheck = [];
  var payslipData = payslipSheet.getRange(1, 1, payslipListCount-1, 4).getValues();
  var checkCount = 0;
  for(var i=1; i<payslipData.length; i++){
    //重複チェック用に出力した値を取得
    //当月取得
    var forDateToString = Utilities.formatDate(payslipData[i][0], "JST", "yyyy/MM/dd");
    var val = forDateToString.toString() + "," + payslipData[i][1] + "," + payslipData[i][2] + "," + payslipData[i][3];
    doubleCheck[checkCount] = val;
    checkCount += 1; 
  }
  //実際の重複チェック
  for(var i=0; i<doubleCheck.length; i++){
    for(var j=0; j<doubleCheck.length; j++){
      if(doubleCheck[i] == doubleCheck[j] && i != j){
        //+2は明細データ（payslipData）がヘッダー合わせてx行、チェック用明細データ（doubleCheck）はx-1行
        //チェック用明細データの値を2行目に出力するには、ヘッダーとゼロ始まりの文をプラス
        payslipFile.getRange("E" + (i+2)).setValue("重複エラー内容を確認してください。");
      }
    }   
  }
}
// スプレッドシートのメニューからPDF作成用の関数を実行出来るように、「スクリプト」というメニューを追加。
function onOpen() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var entries = [       
        {
            name : "一覧出力",
            functionName : "myFunction"
        }
        ];
    sheet.addMenu("スクリプト", entries);
};