// ========================================
// 🤖 LINE 對話機器人設定
// 請將您的 LINE Channel Access Token 貼在下方引號內
// ========================================
var LINE_CHANNEL_ACCESS_TOKEN = 'JN5CDum6yi1xrpO30WCVd7mL3BFrnZxOg4liMhENE7RF8WR1cdZqWhiki7vZ6wbYeURE4tCEUAQyg73uYN1e6fuJCqRXSssv57TXWg5Uoiv2qEHxNqiLpwCwqPE80pT4ixkXza4HoE5ABlpGYYUKgAdB04t89/1O/w1cDnyilFU=';

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var data = JSON.parse(e.postData.contents);
    
    // =============================================
    // 🤖 LINE Webhook 攔截器：判斷是否為 LINE 對話事件
    // =============================================
    if (data.events) {
      return handleLineWebhook(data);
    }
    
    // =============================================
    // 📝 以下為原本的 LIFF 填單邏輯
    // =============================================
    
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
      // 自動補上表頭 (新版 19 欄)
      var headers = [
        "日期", "昨日剩", "新增嫩", "今日用", "嫩(午)", "嫩雞(晚)", "烤(午)", "烤(晚)", "限定(份)", "飯量(鍋)", 
        "預估業績(午)", "業績(午)", "差異值(午)", "業績(晚)", "支出(午/晚)", "匯款業績", "總業績", "差異值", "備註"
      ];
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

// =============================================
// 🤖 LINE 對話機器人：Webhook 處理函式
// =============================================

function handleLineWebhook(data) {
  var events = data.events;
  
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    
    // 只處理文字訊息
    if (event.type !== 'message' || event.message.type !== 'text') continue;
    
    var userText = event.message.text.trim();
    var replyToken = event.replyToken;
    var replyMsg = "";
    
    // 判斷是否在詢問銷量 (彈性多種問法)
    var salesKeywords = ["平均", "銷售", "銷量", "賣多少", "賣出", "多少份", "業績"];
    var isAskingSales = salesKeywords.some(function(kw) { return userText.includes(kw); });
    
    // 🔍 關鍵字辨識
    if (userText === "今日業績" || userText === "今日") {
      replyMsg = getTodayReport();
    } else if (userText === "本月總額" || userText === "本月") {
      replyMsg = getMonthReport();
    } else if (userText.startsWith("查詢")) {
      // 擷取商品名稱，例如「查詢貢丸湯」或「查詢 貢丸湯」
      var productName = userText.replace("查詢", "").trim();
      if (productName !== "") {
        replyMsg = queryProductSales(productName);
      } else {
        replyMsg = "⚠️ 請輸入要查詢的商品名稱，例如：「查詢貢丸湯」";
      }
    } else if (userText === "指令" || userText === "功能" || userText === "help") {
      replyMsg = "🤖 虛擬會計指令清單：\n\n"
        + "📌 「今日業績」— 查看今天的營收與差異值\n"
        + "📌 「本月總額」— 查看本月累積營收與日均業績\n"
        + "📌 「查詢 [商品名]」— 跨表查詢商品平均日銷量（如：查詢貢丸湯）\n"
        + "📌 「指令」— 顯示此說明";
    } else if (isAskingSales) {
      // 如果句子包含問銷量的關鍵字，進入智慧比對模式
      replyMsg = smartQueryProductSales(userText);
    }
    
    // 如果有匹配到指令，就回覆
    if (replyMsg !== "") {
      replyToLine(replyToken, replyMsg);
    }
  }
  
  // LINE Webhook 必須回傳 200 OK
  return ContentService.createTextOutput(JSON.stringify({status: "ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================
// 📊 查詢「今日業績」
// =============================================
function getTodayReport() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();
    var formattedDate = Utilities.formatDate(now, "GMT+8", "yyyy/M");
    var parts = formattedDate.split('/');
    var sheetName = parts[0] + "年" + parts[1] + "月";
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) return "⚠️ 找不到本月工作表「" + sheetName + "」，可能今天尚未填過任何資料。";
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return "⚠️ 本月工作表目前還沒有任何資料。";
    
    // 取最後一筆資料（通常為今天或最近一天）
    var lastRow = data[data.length - 1];
    var lastDate = new Date(lastRow[0]);
    var weekdays = ["日", "一", "二", "三", "四", "五", "六"];
    var mm = lastDate.getMonth() + 1;
    var dd = lastDate.getDate();
    var dayName = weekdays[lastDate.getDay()];
    
    var revLunch    = Number(lastRow[11]) || 0;  // 業績(午)
    var revDinner   = Number(lastRow[13]) || 0;  // 業績(晚)
    var totalRev    = Number(lastRow[16]) || 0;  // 總業績
    var diffLunch   = Number(lastRow[12]) || 0;  // 差異值(午)
    var diffTotal   = Number(lastRow[17]) || 0;  // 差異值
    var expenses    = Number(lastRow[14]) || 0;  // 支出
    var remittance  = Number(lastRow[15]) || 0;  // 匯款業績
    var notes       = lastRow[18] || "";         // 備註
    
    var msg = "📊 " + mm + "/" + dd + "(" + dayName + ") 即時戰報\n"
      + "━━━━━━━━━━━━\n"
      + "🌤️ 午班業績：$" + revLunch.toLocaleString() + "\n"
      + "🌙 晚班業績：$" + revDinner.toLocaleString() + "\n"
      + "💰 全日總額：$" + totalRev.toLocaleString() + "\n"
      + "━━━━━━━━━━━━\n"
      + "📈 午班差異：$" + diffLunch.toLocaleString() + "\n"
      + "📈 全日差異：$" + diffTotal.toLocaleString() + "\n"
      + "💸 總支出：$" + expenses.toLocaleString() + "\n"
      + "🏦 匯款業績：$" + remittance.toLocaleString();
    
    if (notes) {
      msg += "\n━━━━━━━━━━━━\n📝 備註：" + notes;
    }
    
    return msg;
    
  } catch (err) {
    return "❌ 查詢失敗：" + err.toString();
  }
}

// =============================================
// 📊 查詢「本月總額」
// =============================================
function getMonthReport() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();
    var formattedDate = Utilities.formatDate(now, "GMT+8", "yyyy/M");
    var parts = formattedDate.split('/');
    var sheetName = parts[0] + "年" + parts[1] + "月";
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) return "⚠️ 找不到本月工作表「" + sheetName + "」。";
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return "⚠️ 本月工作表目前還沒有任何資料。";
    
    var totalRevenue = 0;
    var totalExpenses = 0;
    var totalDiffLunch = 0;
    var totalDiffAll = 0;
    var daysCount = 0;
    var bestDay = "";
    var bestDayRev = 0;
    
    for (var i = 1; i < data.length; i++) {
      var dailyTotal = Number(data[i][16]) || 0;
      var dailyExp   = Number(data[i][14]) || 0;
      var dailyDiffL = Number(data[i][12]) || 0;
      var dailyDiffA = Number(data[i][17]) || 0;
      
      if (dailyTotal > 0) {
        totalRevenue += dailyTotal;
        totalExpenses += dailyExp;
        totalDiffLunch += dailyDiffL;
        totalDiffAll += dailyDiffA;
        daysCount++;
        
        if (dailyTotal > bestDayRev) {
          bestDayRev = dailyTotal;
          var bd = new Date(data[i][0]);
          var weekdays = ["日", "一", "二", "三", "四", "五", "六"];
          bestDay = (bd.getMonth() + 1) + "/" + bd.getDate() + "(" + weekdays[bd.getDay()] + ")";
        }
      }
    }
    
    var average = daysCount > 0 ? Math.round(totalRevenue / daysCount) : 0;
    var netRevenue = totalRevenue - totalExpenses;
    
    var msg = "📊 " + sheetName + " 月報總覽\n"
      + "━━━━━━━━━━━━\n"
      + "💰 累積總營收：$" + totalRevenue.toLocaleString() + "\n"
      + "💸 累積總支出：$" + totalExpenses.toLocaleString() + "\n"
      + "🏦 淨營收：$" + netRevenue.toLocaleString() + "\n"
      + "━━━━━━━━━━━━\n"
      + "📅 已營業天數：" + daysCount + " 天\n"
      + "📈 日均業績：$" + average.toLocaleString() + "\n"
      + "━━━━━━━━━━━━\n"
      + "📊 累計午班差異：$" + totalDiffLunch.toLocaleString() + "\n"
      + "📊 累計全日差異：$" + totalDiffAll.toLocaleString() + "\n"
      + "━━━━━━━━━━━━\n"
      + "🏆 最佳單日：" + bestDay + " $" + bestDayRev.toLocaleString();
    
    return msg;
    
  } catch (err) {
    return "❌ 查詢失敗：" + err.toString();
  }
}

// =============================================
// 📤 透過 LINE Reply API 回覆訊息
// =============================================
function replyToLine(replyToken, message) {
  var url = "https://api.line.me/v2/bot/message/reply";
  var payload = {
    replyToken: replyToken,
    messages: [{
      type: "text",
      text: message
    }]
  };
  
  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

// =============================================
// 🧪 測試函式：在 Apps Script 編輯器直接執行
// 用來驗證 Token 是否正確、能否發送訊息
// =============================================
function testBot() {
  var url = "https://api.line.me/v2/bot/message/broadcast";
  var payload = {
    messages: [{
      type: "text",
      text: "🤖 測試成功！您的虛擬會計機器人已經上線了！\n\n試試輸入以下指令：\n📌 今日業績\n📌 本月總額\n📌 指令"
    }]
  };
  
  var response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  Logger.log("HTTP Status: " + response.getResponseCode());
  Logger.log("Response: " + response.getContentText());
}

// =============================================
// 📦 查詢「商品平均日銷量」(跨表查詢)
// =============================================
function queryProductSales(productName) {
  try {
    // 外部商品銷售紀錄表的 ID
    var extSheetId = "1bvRiNRZYrhG4u4zf-T3dSJO17OpkhXVNueN2YaQtfqY";
    var extSs = SpreadsheetApp.openById(extSheetId);
    var sheet = extSs.getSheetByName("五月");
    
    if (!sheet) {
      return "⚠️ 找不到名為「五月」的工作表，請確認表格結構。";
    }
    
    var data = sheet.getDataRange().getValues();
    var totalSales = 0;
    var found = false;
    
    // 假設第一列是標題，從第二列開始搜尋 (index 1)
    // A欄是商品名稱 (index 0)
    // G欄是銷售數量 (index 6)
    for (var i = 1; i < data.length; i++) {
      var rowProductName = String(data[i][0]).trim();
      var rowSales = Number(data[i][6]);
      
      // 使用 includes 進行模糊搜尋，只要包含關鍵字就算
      if (rowProductName !== "" && rowProductName.includes(productName)) {
        if (!isNaN(rowSales) && rowSales > 0) {
          totalSales += rowSales;
        }
        found = true;
      }
    }
    
    if (!found) {
      return "⚠️ 找不到與「" + productName + "」相關的商品紀錄。";
    }
    
    // 平均日銷算法：總銷量 / 23天
    var avgSales = totalSales / 23;
    // 取到小數點第一位
    avgSales = Math.round(avgSales * 10) / 10;
    
    var msg = "📦 查詢商品：" + productName + "\n"
            + "━━━━━━━━━━━━\n"
            + "📊 總銷售量：" + totalSales + " 份\n"
            + "📅 計算天數：23 天\n"
            + "📈 平均日銷量：" + avgSales + " 份/天";
            
    return msg;
    
  } catch (err) {
    return "❌ 查詢失敗，可能是尚未授權存取該外部表格，或是網址有誤。\n詳細錯誤：" + err.toString();
  }
}

// =============================================
// 🧠 智慧查詢「商品平均日銷量」(反向比對句意)
// =============================================
function smartQueryProductSales(sentence) {
  try {
    // 外部商品銷售紀錄表的 ID
    var extSheetId = "1bvRiNRZYrhG4u4zf-T3dSJO17OpkhXVNueN2YaQtfqY";
    var extSs = SpreadsheetApp.openById(extSheetId);
    var sheet = extSs.getSheetByName("五月");
    
    if (!sheet) {
      return "⚠️ 找不到名為「五月」的工作表，請確認表格結構。";
    }
    
    var data = sheet.getDataRange().getValues();
    var matchedProductName = "";
    var totalSales = 0;
    
    // 假設第一列是標題，從第二列開始搜尋 (index 1)
    for (var i = 1; i < data.length; i++) {
      var rowProductName = String(data[i][0]).trim();
      var rowSales = Number(data[i][6]);
      
      if (rowProductName !== "") {
        // 反向比對：如果整句話包含了這個商品名稱
        if (sentence.includes(rowProductName)) {
          matchedProductName = rowProductName; // 紀錄匹配到的商品
          if (!isNaN(rowSales) && rowSales > 0) {
            totalSales += rowSales; // 一併加總
          }
        }
      }
    }
    
    // 如果掃完了整個表格，還是沒發現有哪個商品被提到
    if (matchedProductName === "") {
      return "🤔 抱歉，我知道您想問銷量，但我無法在這句話中辨識出您想查詢哪一個「特定的商品名稱」喔！\n請試著輸入明確的商品名，例如：請問嫩雞賣多少？";
    }
    
    // 平均日銷算法：總銷量 / 23天
    var avgSales = totalSales / 23;
    // 取到小數點第一位
    avgSales = Math.round(avgSales * 10) / 10;
    
    var msg = "🧠 智慧辨識商品：" + matchedProductName + "\n"
            + "━━━━━━━━━━━━\n"
            + "📊 總銷售量：" + totalSales + " 份\n"
            + "📅 計算天數：23 天\n"
            + "📈 平均日銷量：" + avgSales + " 份/天";
            
    return msg;
    
  } catch (err) {
    return "❌ 查詢失敗，可能是尚未授權存取該外部表格，或是網址有誤。\n詳細錯誤：" + err.toString();
  }
}
