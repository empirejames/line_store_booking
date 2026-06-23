function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var sheet = ss.getSheets()[0]; 

    var data = JSON.parse(e.postData.contents);
    
    // 擷取前端輸入的資料 (與最新 Google Sheet 欄位對應)
    var dateStr = data.date || Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd");
    var yesterdayRemain = Number(data.yesterdayRemain) || 0; // 昨日剩
    var addedTender = Number(data.addedTender) || 0;         // 新增嫩
    var chickenUsed = Number(data.chickenUsed) || 0;         // 今日用
    var roastCount = Number(data.roastCount) || 0;           // 烤(午/晚)
    var limited = Number(data.limited) || 0;                 // 限定
    var riceAmount = Number(data.riceAmount) || 0;           // 飯量(鍋)
    
    var revenueLunch = Number(data.revenueLunch) || 0;       // 業績(午)
    var revenueDinner = Number(data.revenueDinner) || 0;     // 業績(晚)
    var expenses = Number(data.expenses) || 0;               // 支出(午/晚)
    var remittance = Number(data.remittance) || 0;           // 匯款業績

    // 🚀 執行公式計算
    var totalRevenue = revenueLunch + revenueDinner;
    var estimatedRevenue = (chickenUsed * 2 * 130) + (roastCount * 140);
    var difference = estimatedRevenue - totalRevenue;

    // 將資料寫入試算表最新一列 (必須嚴格對應試算表由左至右的順序)
    // 日期 | 昨日剩 | 新增嫩 | 今日用 | 烤(午/晚) | 限定 | 飯量(鍋) | 預估業績 | 業績(午) | 業績(晚) | 支出(午/晚) | 匯款業績 | 總業績 | 差異值
    var rowData = [
      dateStr,           // 1. 日期
      yesterdayRemain,   // 2. 昨日剩
      addedTender,       // 3. 新增嫩
      chickenUsed,       // 4. 今日用
      roastCount,        // 5. 烤(午/晚)
      limited,           // 6. 限定
      riceAmount,        // 7. 飯量(鍋)
      estimatedRevenue,  // 8. 預估業績 (系統計算)
      revenueLunch,      // 9. 業績(午)
      revenueDinner,     // 10. 業績(晚)
      expenses,          // 11. 支出(午/晚)
      remittance,        // 12. 匯款業績
      totalRevenue,      // 13. 總業績 (系統計算)
      difference         // 14. 差異值 (系統計算)
    ];
    
    sheet.appendRow(rowData);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "報表已成功送出並記錄！",
      summary: "總業績: $" + totalRevenue + "\n差異值: $" + difference
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doOptions(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.JSON).setHeaders(headers);
}
