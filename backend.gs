function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var data = JSON.parse(e.postData.contents);
    
    // 基本資料
    var shift = data.shift; // 'lunch' 或 'dinner'
    var dateStr = data.date || Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd");
    
    // 根據日期自動產生/尋找月份工作表 (格式: 2026年6月)
    var d = new Date(dateStr);
    var sheetName = d.getFullYear() + "年" + (d.getMonth() + 1) + "月";
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      // 找不到就建立新的工作表
      sheet = ss.insertSheet(sheetName);
      // 自動補上表頭
      var headers = ["日期", "昨日剩", "新增嫩", "今日用", "烤(午/晚)", "嫩雞結餘", "飯量(鍋)", "限定", "業績(午)", "業績(晚)", "支出(午/晚)", "匯款業績", "總業績", "差異值"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
      sheet.setFrozenRows(1);
    }

    // 取得目前工作表內的所有資料，用來尋找今天那一行
    var allValues = sheet.getDataRange().getValues();

    // 🌟 新增邏輯：如果是來拿初始資料的
    if (data.action === 'getInitData') {
      var yesterdayRemain = "";
      for (var j = allValues.length - 1; j >= 1; j--) {
        var cellDate = allValues[j][0];
        var formatted = "";
        if (cellDate instanceof Date) {
          formatted = Utilities.formatDate(cellDate, "GMT+8", "yyyy/MM/dd");
        } else {
          var d = new Date(cellDate);
          if (!isNaN(d.getTime())) {
            formatted = Utilities.formatDate(d, "GMT+8", "yyyy/MM/dd");
          } else {
            formatted = String(cellDate).trim();
          }
        }
        var inputDateStr = Utilities.formatDate(new Date(dateStr), "GMT+8", "yyyy/MM/dd");
        if (formatted && formatted !== inputDateStr) {
          // 找到非今天的最新一筆 (也就是昨天)
          var prevRemain = Number(allValues[j][1]) || 0; // 昨日剩
          var prevAdded  = Number(allValues[j][2]) || 0; // 新增嫩
          var prevUsed   = Number(allValues[j][3]) || 0; // 今日用
          yesterdayRemain = prevRemain + prevAdded - prevUsed;
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        yesterdayRemain: yesterdayRemain
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var targetRowIndex = -1;
    var existingData = [];

    // 從最後一列往上找，找尋日期相符的列 (因為通常是新增在最下面)
    for (var i = allValues.length - 1; i >= 1; i--) { 
      var cellDate = allValues[i][0];
      var cellDateStr = "";
      if (cellDate instanceof Date) {
        cellDateStr = Utilities.formatDate(cellDate, "GMT+8", "yyyy/MM/dd");
      } else {
        var d = new Date(cellDate);
        if (!isNaN(d.getTime())) {
          cellDateStr = Utilities.formatDate(d, "GMT+8", "yyyy/MM/dd");
        } else {
          cellDateStr = String(cellDate).trim();
        }
      }

      var inputDateStr = Utilities.formatDate(new Date(dateStr), "GMT+8", "yyyy/MM/dd");

      if (cellDateStr === inputDateStr) {
        targetRowIndex = i + 1; // Apps Script 的列號從 1 開始
        existingData = allValues[i];
        break;
      }
    }

    // 若找不到今日資料，預設一個長度為 14 的空陣列
    if (targetRowIndex === -1) {
      existingData = new Array(14).fill(""); 
      existingData[0] = dateStr;
    }

    // 準備提取與計算變數 (轉為數字，若無則補 0)
    function parseNum(val) { return Number(val) || 0; }

    var summaryMsg = "";

    if (shift === 'lunch') {
      // ==== 午班結帳邏輯 ====
      var revLunch = parseNum(data.revenueLunch);
      var roastLunch = parseNum(data.roastLunch);
      var expLunch = parseNum(data.expensesLunch);
      
      existingData[8] = revLunch;    // 第 9 欄: 業績(午)
      existingData[4] = roastLunch;  // 第 5 欄: 烤(午/晚) -> 暫存午餐數量
      existingData[10] = expLunch;   // 第 11 欄: 支出(午/晚) -> 暫存午餐數量
      
      // 午班尚未算完總業績，暫時以午班業績當作總業績
      existingData[12] = revLunch;   // 第 13 欄: 總業績
      
      summaryMsg = "午餐業績已記錄：$" + revLunch;

    } else if (shift === 'dinner') {
      // ==== 晚班結帳邏輯 ====
      var revDinner = parseNum(data.revenueDinner);
      var roastDinner = parseNum(data.roastDinner);
      var expDinner = parseNum(data.expensesDinner);
      
      // 讀取已經存在的「午班」資料來加總
      var currentRoastLunch = parseNum(existingData[4]);
      var currentExpLunch = parseNum(existingData[10]);
      var currentRevLunch = parseNum(existingData[8]); 

      // 更新盤點與獨立欄位
      existingData[9] = revDinner;                               // 業績(晚)
      existingData[1] = parseNum(data.yesterdayRemain) || "";    // 昨日剩
      existingData[2] = parseNum(data.addedTender) || "";        // 新增嫩
      var chickenUsed = parseNum(data.chickenUsed);
      existingData[3] = chickenUsed;                             // 今日用(嫩雞)
      existingData[5] = parseNum(data.limited) || "";            // 限定
      existingData[6] = parseNum(data.riceAmount) || "";         // 飯量(鍋)

      // 🔄 加總共用欄位 (午 + 晚)
      var totalRoast = currentRoastLunch + roastDinner;
      var totalExpenses = currentExpLunch + expDinner;
      var totalRevenue = currentRevLunch + revDinner;

      // 💰 自動計算匯款業績 (總業績 - 支出)
      var remittance = totalRevenue - totalExpenses;

      existingData[4] = totalRoast;     // 烤(午/晚)
      existingData[10] = totalExpenses; // 支出(午/晚)
      existingData[11] = remittance;    // 匯款業績 (自動計算)
      existingData[12] = totalRevenue;  // 總業績

      // 執行核心公式：預估業績與差異值
      var estimatedRevenue = (chickenUsed * 2 * 130) + (totalRoast * 140);
      var difference = totalRevenue - estimatedRevenue;

      existingData[7] = estimatedRevenue; // 預估業績
      existingData[13] = difference;      // 差異值

      summaryMsg = "全日結算完成！\n預估業績: $" + estimatedRevenue + "\n實際總業績: $" + totalRevenue + "\n差異值: $" + difference;
    }

    // 寫回 Google Sheets
    if (targetRowIndex === -1) {
      // 找不到今日，代表午班沒人填，晚班直接新增一行
      sheet.appendRow(existingData);
    } else {
      // 找到今日，更新同一行！(這就是動態更新的核心)
      sheet.getRange(targetRowIndex, 1, 1, existingData.length).setValues([existingData]);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "報表已成功送出並記錄！",
      summary: summaryMsg
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "系統發生錯誤: " + error.toString()
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

// 🌟 新增 doGet 函式：供 GitHub Actions 讀取當月平均業績
function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === 'getAverage') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var d = new Date();
      // 在台灣時區執行，也可以使用 Utilities.formatDate(new Date(), "GMT+8", "yyyy年M月")，但這裡簡單計算即可
      // 注意：GAS 時區預設可能不同，最好強制轉為 GMT+8
      var formattedDate = Utilities.formatDate(new Date(), "GMT+8", "yyyy/M");
      var parts = formattedDate.split('/');
      var sheetName = parts[0] + "年" + parts[1] + "月";
      var sheet = ss.getSheetByName(sheetName);
      
      var average = 0;
      var total = 0;
      var daysCount = 0;

      if (sheet) {
        var data = sheet.getDataRange().getValues();
        // 假設第一列是表頭，從第二列開始
        for (var i = 1; i < data.length; i++) {
          var dailyTotal = Number(data[i][12]); // 第13欄: 總業績
          if (!isNaN(dailyTotal) && dailyTotal > 0) {
            total += dailyTotal;
            daysCount++;
          }
        }
        if (daysCount > 0) {
          average = Math.round(total / daysCount);
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        month: sheetName,
        total: total,
        days: daysCount,
        average: average
      })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}
