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
    } else if (action === 'getAllProductSales') {
      var targetMonthStr = e.parameter.month; // e.g. "六月"
      var monthMapping = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二"];
      var targetMonthStr = e.parameter.month; // e.g. "六月"
      if (!targetMonthStr) {
        var m = new Date().getMonth();
        targetMonthStr = monthMapping[m] + "月";
      }
      
      // 計算上個月的名稱
      var targetMonthName = targetMonthStr.replace("月", "");
      var mIndex = monthMapping.indexOf(targetMonthName);
      var prevMonthStr = "";
      if (mIndex > 0) {
        prevMonthStr = monthMapping[mIndex - 1] + "月";
      } else if (mIndex === 0) {
        prevMonthStr = "十二月";
      }
      
      var extSheetId = "1bvRiNRZYrhG4u4zf-T3dSJO17OpkhXVNueN2YaQtfqY";
      var extSs = SpreadsheetApp.openById(extSheetId);
      
      // 抓取當月
      var sheet = extSs.getSheetByName(targetMonthStr);
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "error",
          message: "找不到名為「" + targetMonthStr + "」的工作表"
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // 抓取上個月 (如果不存在就不強求)
      var prevSheet = extSs.getSheetByName(prevMonthStr);
      var prevSalesDict = {};
      if (prevSheet) {
        var prevData = prevSheet.getDataRange().getValues();
        for (var k = 1; k < prevData.length; k++) {
          var pName = String(prevData[k][0]).trim();
          var pSales = Number(prevData[k][6]);
          if (pName !== "" && !isNaN(pSales) && pSales > 0) {
            if (!prevSalesDict[pName]) prevSalesDict[pName] = 0;
            prevSalesDict[pName] += pSales;
          }
        }
      }
      
      var data = sheet.getDataRange().getValues();
      var salesDict = {};
      
      // A欄(0)商品名, G欄(6)銷售數量
      for (var i = 1; i < data.length; i++) {
        var productName = String(data[i][0]).trim();
        var sales = Number(data[i][6]);
        
        if (productName !== "" && !isNaN(sales) && sales > 0) {
          if (!salesDict[productName]) {
            salesDict[productName] = 0;
          }
          salesDict[productName] += sales;
        }
      }
      
      var result = [];
      for (var p in salesDict) {
        result.push({ 
          name: p, 
          sales: salesDict[p],
          prevSales: prevSalesDict[p] !== undefined ? prevSalesDict[p] : null
        });
      }
      
      // 依照銷量由高到低排序
      result.sort(function(a, b) { return b.sales - a.sales; });
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        month: targetMonthStr,
        prevMonth: prevMonthStr,
        data: result
      })).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'getTimeDistribution') {
      var timeSheetId = "1AL0rR0w1xObsWaN-iIKV0tpDHR5iihujiue-dtFSsuE";
      var timeSs = SpreadsheetApp.openById(timeSheetId);
      var sheet = timeSs.getSheets()[0]; // 讀取第一個工作表
      
      var data = sheet.getDataRange().getValues();
      var timeData = [];
      
      // 從第二列開始讀取 (略過標題列)
      for (var i = 1; i < data.length; i++) {
        var timeStr = String(data[i][0]).trim(); // A欄: 時段
        var transactions = Number(data[i][4]);   // E欄: 交易(索引為4)
        
        if (timeStr !== "" && !isNaN(transactions)) {
          timeData.push({
            time: timeStr,
            transactions: transactions
          });
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        data: timeData
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
    var salesKeywords = ["平均", "銷售", "銷量", "賣多少", "賣出", "多少份", "業績", "查詢"];
    var isAskingSales = salesKeywords.some(function(kw) { return userText.includes(kw); });
    
    // 🔍 關鍵字辨識 (使用彈性比對)
    if (userText === "今日業績" || userText === "今日" || (userText.includes("今日") && (userText.includes("業績") || userText.includes("營收")))) {
      replyMsg = getTodayReport();
    } else if (userText === "本月總額" || userText === "本月" || ((userText.includes("本月") || userText.includes("這個月")) && (userText.includes("業績") || userText.includes("營收") || userText.includes("總額")))) {
      replyMsg = getMonthReport();
    } else if (userText.includes("圖表")) {
      replyMsg = "📈 您的專屬業績圖表已準備好：\n\n🔗 點擊下方連結查看：\nhttps://empirejames.github.io/line_store_booking/chart.html";
    } else if (userText.includes("報表")) {
      replyMsg = "📊 您的專屬商品銷售報表已準備好：\n\n🔗 點擊下方連結查看排行榜：\nhttps://empirejames.github.io/line_store_booking/report.html";
    } else if (userText.includes("時間") || userText.includes("時段") || userText.includes("熱度")) {
      replyMsg = "🔥 您的專屬「時間點熱度分析圖表」已準備好：\n\n🔗 點擊下方連結查看各時段客流量：\nhttps://empirejames.github.io/line_store_booking/time_chart.html";
    } else if (userText.toLowerCase().includes("excel")) {
      replyMsg = "🔗 您的 Excel 營收記帳表連結如下：\nhttps://docs.google.com/spreadsheets/d/1Yw47QEBNeIO1IjeItZ6d0CmJBdnKGeGBzOTUHBUJEPA/edit?gid=1596698359#gid=1596698359";
    } else if (userText === "指令" || userText === "功能" || userText === "help") {
      replyMsg = getHelpFlexMessage();
    } else if (userText === "如何查詢商品?") {
      replyMsg = "💡 查詢商品說明：\n"
               + "請輸入「月份 + 商品名」，例如：「六月好吃嫩雞飯」\n\n"
               + "📌 支援的商品名稱列表：\n"
               + "好吃嫩雞飯\n"
               + "好吃烤雞飯\n"
               + "雙蛋嫩雞飯\n"
               + "雙倍嫩雞飯\n"
               + "貢丸湯\n"
               + "嫩烤雙拼飯\n"
               + "嫩雞扒蛋飯\n"
               + "蘿蔔湯\n"
               + "荷包蛋\n"
               + "燙青菜\n"
               + "雞油飯\n"
               + "嫩雞一份\n"
               + "貢丸一顆\n"
               + "烤雞一份\n"
               + "白飯一碗\n"
               + "菜飯";
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
function replyToLine(replyToken, messagePayload) {
  var url = "https://api.line.me/v2/bot/message/reply";
  
  var messages = [];
  if (typeof messagePayload === 'string') {
    messages = [{ type: "text", text: messagePayload }];
  } else if (Array.isArray(messagePayload)) {
    messages = messagePayload;
  } else if (typeof messagePayload === 'object') {
    messages = [messagePayload];
  }
  
  var payload = {
    replyToken: replyToken,
    messages: messages
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
// 🧠 智慧查詢「商品平均日銷量」(反向比對句意與動態月份)
// =============================================
function smartQueryProductSales(sentence) {
  try {
    // 1. 判斷使用者問的是哪個月？
    var monthMapping = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二"];
    var targetMonthNum = new Date().getMonth() + 1; // 預設為當前月份
    
    // 檢查字串中是否包含「X月」
    for (var m = 1; m <= 12; m++) {
      if (sentence.includes(m + "月") || sentence.includes("0" + m + "月") || sentence.includes(monthMapping[m-1] + "月")) {
        targetMonthNum = m;
        break;
      }
    }
    
    var extSheetName = monthMapping[targetMonthNum - 1] + "月"; // e.g. "六月"
    
    // 2. 獲取該月份的營業天數（從目前的記帳本獲取，扣除禮拜六）
    var internalSs = SpreadsheetApp.getActiveSpreadsheet();
    var currentYear = new Date().getFullYear();
    var internalSheetName = currentYear + "年" + targetMonthNum + "月";
    var internalSheet = internalSs.getSheetByName(internalSheetName);
    var operatingDays = 0;
    
    if (internalSheet) {
      var internalData = internalSheet.getDataRange().getValues();
      // 假設第一列為標題，從第二列開始
      for (var r = 1; r < internalData.length; r++) {
        var rowDate = internalData[r][0]; // A 欄是日期
        if (rowDate instanceof Date) {
          // getDay() 回傳 0-6，0 是週日，6 是週六
          if (rowDate.getDay() !== 6) {
            operatingDays++;
          }
        }
      }
    }
    
    // 預防萬一天數為 0 導致除以 0 錯誤
    if (operatingDays === 0) {
      operatingDays = 1; 
    }

    // 3. 連線外部表格獲取商品銷量
    var extSheetId = "1bvRiNRZYrhG4u4zf-T3dSJO17OpkhXVNueN2YaQtfqY";
    var extSs = SpreadsheetApp.openById(extSheetId);
    var sheet = extSs.getSheetByName(extSheetName);
    
    if (!sheet) {
      return "⚠️ 找不到名為「" + extSheetName + "」的銷售工作表。";
    }
    
    var data = sheet.getDataRange().getValues();
    var matchedProductName = "";
    var totalSales = 0;
    
    // 移除不必要的關鍵字，讓商品配對更精準
    var cleanSentence = sentence.replace("查詢", "").replace(targetMonthNum + "月", "").replace(extSheetName, "");

    var maxLcsLen = 0;

    // 第一階段：掃描所有商品，找出跟使用者的句子「重疊字數最多」的最佳匹配商品
    for (var i = 1; i < data.length; i++) {
      var rowProductName = String(data[i][0]).trim();
      
      if (rowProductName !== "" && rowProductName.length >= 2) {
        var lcs = "";
        // 找出 cleanSentence 和 rowProductName 的最長共同子字串 (Longest Common Substring)
        for (var start = 0; start < cleanSentence.length; start++) {
          for (var end = start + 1; end <= cleanSentence.length; end++) {
            var sub = cleanSentence.substring(start, end);
            if (rowProductName.includes(sub) && sub.length > lcs.length) {
              lcs = sub;
            }
          }
        }
        
        // 只要有 2 個字以上吻合，且比目前找到的更長，就認定它是最可能的商品
        if (lcs.length >= 2 && lcs.length > maxLcsLen) {
          maxLcsLen = lcs.length;
          matchedProductName = rowProductName;
        }
      }
    }
    
    // 第二階段：如果有找到匹配的商品，就把該商品的所有銷量加總
    if (matchedProductName !== "") {
      for (var j = 1; j < data.length; j++) {
        if (String(data[j][0]).trim() === matchedProductName) {
          var rowSales = Number(data[j][6]);
          if (!isNaN(rowSales) && rowSales > 0) {
            totalSales += rowSales; 
          }
        }
      }
    } else {
      return "🤔 抱歉，我無法在「" + extSheetName + "」的表格中辨識出您想查詢的「商品名稱」喔！\n表格內的商品可能有特定全名（例如：好吃嫩雞飯），請確定您的句子有包含商品的部分名稱。";
    }
    
    var avgSales = totalSales / operatingDays;
    avgSales = Math.round(avgSales * 10) / 10;
    
    var msg = "📦 查詢商品：" + matchedProductName + " (" + extSheetName + ")\n"
            + "━━━━━━━━━━━━\n"
            + "📊 總銷售量：" + totalSales + " 份\n"
            + "📅 營業天數：" + operatingDays + " 天 (已扣除週六)\n"
            + "📈 平均日銷量：" + avgSales + " 份/天";
            
    return msg;
    
  } catch (err) {
    return "❌ 查詢失敗，詳細錯誤：" + err.toString();
  }
}

// =============================================
// 🎴 產生精美的 Flex Message (Carousel) 指令選單
// =============================================
function getHelpFlexMessage() {
  return [{
    "type": "flex",
    "altText": "虛擬會計指令選單",
    "contents": {
      "type": "carousel",
      "contents": [
        {
          "type": "bubble",
          "header": { "type": "box", "layout": "vertical", "contents": [{ "type": "text", "text": "💰 營收速報", "weight": "bold", "color": "#1DB446", "size": "sm" }] },
          "hero": { "type": "image", "url": "https://images.unsplash.com/photo-1554224155-8d04cb21cd6c?w=400", "size": "full", "aspectRatio": "20:13", "aspectMode": "cover" },
          "body": { "type": "box", "layout": "vertical", "contents": [
            { "type": "text", "text": "營收管理", "weight": "bold", "size": "xl" },
            { "type": "text", "text": "快速查看當日與當月業績", "color": "#aaaaaa", "size": "xs", "wrap": true }
          ] },
          "footer": { "type": "box", "layout": "vertical", "spacing": "sm", "contents": [
            { "type": "button", "style": "primary", "height": "sm", "action": { "type": "message", "label": "今日業績", "text": "今日業績" } },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "本月總額", "text": "本月總額" } }
          ] }
        },
        {
          "type": "bubble",
          "header": { "type": "box", "layout": "vertical", "contents": [{ "type": "text", "text": "📊 數據圖表", "weight": "bold", "color": "#1DB446", "size": "sm" }] },
          "hero": { "type": "image", "url": "https://images.unsplash.com/photo-1551288049-bebda4e38f71?w=400", "size": "full", "aspectRatio": "20:13", "aspectMode": "cover" },
          "body": { "type": "box", "layout": "vertical", "contents": [
            { "type": "text", "text": "視覺化分析", "weight": "bold", "size": "xl" },
            { "type": "text", "text": "掌握銷售排行與客流熱度", "color": "#aaaaaa", "size": "xs", "wrap": true }
          ] },
          "footer": { "type": "box", "layout": "vertical", "spacing": "sm", "contents": [
            { "type": "button", "style": "primary", "height": "sm", "action": { "type": "message", "label": "查看業績圖表", "text": "圖表" } },
            { "type": "button", "style": "primary", "height": "sm", "color": "#1e293b", "action": { "type": "message", "label": "商品銷售報表", "text": "報表" } },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "時間熱度分析", "text": "時間熱度" } }
          ] }
        },
        {
          "type": "bubble",
          "header": { "type": "box", "layout": "vertical", "contents": [{ "type": "text", "text": "🛠️ 進階工具", "weight": "bold", "color": "#1DB446", "size": "sm" }] },
          "hero": { "type": "image", "url": "https://images.unsplash.com/photo-1432888498266-38ffec3eaf0a?w=400", "size": "full", "aspectRatio": "20:13", "aspectMode": "cover" },
          "body": { "type": "box", "layout": "vertical", "contents": [
            { "type": "text", "text": "快速連結與查詢", "weight": "bold", "size": "xl" },
            { "type": "text", "text": "下載表單或查詢單一商品", "color": "#aaaaaa", "size": "xs", "wrap": true }
          ] },
          "footer": { "type": "box", "layout": "vertical", "spacing": "sm", "contents": [
            { "type": "button", "style": "primary", "height": "sm", "action": { "type": "message", "label": "取得 Excel 連結", "text": "excel" } },
            { "type": "button", "style": "secondary", "height": "sm", "action": { "type": "message", "label": "如何查詢商品?", "text": "如何查詢商品?" } }
          ] }
        }
      ]
    }
  }];
}
