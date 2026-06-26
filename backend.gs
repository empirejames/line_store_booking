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
      var headers = ["日期", "昨日剩", "新增嫩", "今日用", "烤(午/晚)", "限定(份)", "飯量(鍋)", "預估業績", "業績(午)", "業績(晚)", "支出(午/晚)", "匯款業績", "總業績", "差異值"];
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

    // 若找不到今日資料，預設一個長度為 19 的空陣列 (新增備註欄)
    if (targetRowIndex === -1) {
      existingData = new Array(19).fill(""); 
      existingData[0] = dateStr;
    }

    // 準備提取與計算變數 (轉為數字，若無則補 0)
    function parseNum(val) { return Number(val) || 0; }

    var summaryMsg = "";

    if (shift === 'lunch') {
      // ==== 午班結帳邏輯 ====
      var revLunch = parseNum(data.revenueLunch);
      var tenderLunch = parseNum(data.tenderLunch);
      var roastLunch = parseNum(data.roastLunch);
      var expLunch = parseNum(data.expensesLunch);
      
      var estRevLunch = (tenderLunch * 2 * 130) + (roastLunch * 140);
      var diffLunch = revLunch - estRevLunch;

      existingData[4] = tenderLunch;   // 嫩(午)
      existingData[6] = roastLunch;    // 烤(午)
      existingData[10] = estRevLunch;  // 預估業績(午)
      existingData[11] = revLunch;     // 業績(午)
      existingData[12] = diffLunch;    // 差異值(午)
      existingData[14] = expLunch;     // 支出(午/晚) -> 暫存午餐支出
      existingData[16] = revLunch;     // 總業績 -> 暫存午餐業績
      
      summaryMsg = "午餐結算完成！\n預估: $" + estRevLunch + "\n實際: $" + revLunch + "\n差異: $" + diffLunch;

    } else if (shift === 'dinner') {
      // ==== 晚班結帳邏輯 ====
      var revDinner = parseNum(data.revenueDinner);
      var roastDinner = parseNum(data.roastDinner);
      var tenderDinner = parseNum(data.tenderDinner);
      var expDinner = parseNum(data.expensesDinner);
      
      // 讀取已經存在的「午班」資料來加總與計算
      var tenderLunch = parseNum(existingData[4]);
      var roastLunch = parseNum(existingData[6]);
      var expLunch = parseNum(existingData[14]);
      var revLunch = parseNum(existingData[11]); 

      var totalTenderUsed = tenderLunch + tenderDinner;
      var totalRoast = roastLunch + roastDinner;
      var totalExpenses = expLunch + expDinner;
      var totalRevenue = revLunch + revDinner;
      var remittance = totalRevenue - totalExpenses;

      // 更新獨立欄位
      existingData[1] = parseNum(data.yesterdayRemain) || "";    // 昨日剩
      existingData[2] = parseNum(data.addedTender) || "";        // 新增嫩
      existingData[3] = totalTenderUsed;                         // 今日用(午+晚)
      existingData[5] = tenderDinner;                            // 嫩雞(晚)
      existingData[7] = roastDinner;                             // 烤(晚)
      existingData[8] = parseNum(data.limited) || "";            // 限定
      existingData[9] = parseNum(data.riceAmount) || "";         // 飯量(鍋)
      existingData[13] = revDinner;                              // 業績(晚)
      
      // 🔄 更新共用加總欄位
      existingData[14] = totalExpenses; // 支出(午/晚)
      existingData[15] = remittance;    // 匯款業績
      existingData[16] = totalRevenue;  // 總業績

      // 全日預估業績與差異值
      var estimatedRevenueTotal = (totalTenderUsed * 2 * 130) + (totalRoast * 140);
      var differenceTotal = totalRevenue - estimatedRevenueTotal;

      existingData[17] = differenceTotal; // 差異值 (全日)

      summaryMsg = "全日結算完成！\n總業績: $" + totalRevenue + "\n全日差異: $" + differenceTotal + "\n午班差異: $" + parseNum(existingData[12]);
    }

    // 處理備註欄位 (如果有填寫的話)
    if (data.notes && data.notes.trim() !== "") {
      var shiftName = shift === 'lunch' ? "[午班]" : "[晚班]";
      var newNote = shiftName + " " + data.notes.trim();
      if (existingData[18] && existingData[18].trim() !== "") {
        // 若已經有備註，則換行附加
        // 避免重複寫入同一班的相同備註 (簡單防呆)
        if (!existingData[18].includes(newNote)) {
            existingData[18] = existingData[18] + "\n" + newNote;
        }
      } else {
        existingData[18] = newNote;
      }
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
      var chartLabels = [];
      var chartData = [];
      
      var todayDateStr = "今日";
      var todayLunchDiff = 0;
      var todayTotalDiff = 0;

      if (sheet) {
        var data = sheet.getDataRange().getValues();
        // 假設第一列是表頭，從第二列開始
        for (var i = 1; i < data.length; i++) {
          var dateValue = new Date(data[i][0]); // 取得日期
          var dayOfWeek = dateValue.getDay();   // 0=日, 1=一, 2=二, 3=三, 4=四, 5=五, 6=六
          
          var dailyTotal = Number(data[i][16]); // 第17欄: 總業績 (index 16)
          
          // 圖表資料：每天都加入（包含假日，看整體趨勢）
          if (!isNaN(dailyTotal) && dailyTotal > 0) {
             var weekdays = ["日", "一", "二", "三", "四", "五", "六"];
             var dayStr = (dateValue.getMonth() + 1) + "/" + dateValue.getDate() + "(" + weekdays[dayOfWeek] + ")";
             chartLabels.push(dayStr);
             chartData.push(dailyTotal);
          }
          
          // 平均業績：只統計週一(1) 到 週五(5)
          if (dayOfWeek >= 1 && dayOfWeek <= 5) {
            if (!isNaN(dailyTotal) && dailyTotal > 0) {
              total += dailyTotal;
              daysCount++;
            }
          }
        }
        if (daysCount > 0) {
          average = Math.round(total / daysCount);
        }
        
        // 取得最後一天 (通常為今日) 的差異值
        if (data.length > 1) {
            var lastRow = data[data.length - 1];
            var lastDate = new Date(lastRow[0]);
            var mm = lastDate.getMonth() + 1;
            var dd = lastDate.getDate();
            var weekdays = ["日", "一", "二", "三", "四", "五", "六"];
            todayDateStr = mm + "/" + dd + "(" + weekdays[lastDate.getDay()] + ")";
            
            todayLunchDiff = Number(lastRow[12]) || 0;
            todayTotalDiff = Number(lastRow[17]) || 0;
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        month: sheetName,
        total: total,
        days: daysCount,
        average: average,
        chartLabels: chartLabels,
        chartData: chartData,
        todayDateStr: todayDateStr,
        todayLunchDiff: todayLunchDiff,
        todayTotalDiff: todayTotalDiff
      })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}
