const GAS_URL = process.env.GAS_URL;
const LINE_TOKEN = process.env.LINE_TOKEN;
const LINE_USER_ID = process.env.LINE_USER_ID;

if (!GAS_URL || !LINE_TOKEN || !LINE_USER_ID) {
  console.error("Missing environment variables.");
  process.exit(1);
}

async function run() {
  try {
    console.log("Fetching data from GAS...");
    
    // 1. 取得每日平均業績
    // 使用 fetch 可以自動跟隨 Google Apps Script 的 302 重新導向
    const res = await fetch(`${GAS_URL}?action=getAverage`);
    const dataText = await res.text();
    
    let result;
    try {
      result = JSON.parse(dataText);
    } catch (err) {
      console.error("Failed to parse GAS response. Raw response:", dataText.substring(0, 200));
      process.exit(1);
    }

    if (result.status !== 'success') {
      console.error("Error from GAS:", result.message);
      process.exit(1);
    }
    
    console.log("Data fetched successfully. Preparing Flex Message...");

    // 2. 建立精美的 Flex Message 結構
    const flexMessage = {
      messages: [
        {
          type: "flex",
          altText: "📊 每日業績戰報",
          contents: {
            type: "bubble",
            size: "mega",
            header: {
              type: "box",
              layout: "vertical",
              paddingAll: "20px",
              paddingTop: "24px",
              paddingBottom: "24px",
              background: {
                type: "linearGradient",
                angle: "45deg",
                startColor: "#1A2980",
                endColor: "#26D0CE"
              },
              contents: [
                {
                  type: "text",
                  text: "📊 當月平日業績戰報",
                  color: "#ffffff",
                  size: "xl",
                  weight: "bold",
                  align: "center"
                },
                {
                  type: "text",
                  text: `${result.month} (週一至週五)`,
                  color: "#ffffffcc",
                  size: "sm",
                  align: "center",
                  margin: "md"
                }
              ]
            },
            body: {
              type: "box",
              layout: "vertical",
              paddingAll: "24px",
              contents: [
                {
                  type: "text",
                  text: "🔥 目前每日平均業績",
                  color: "#888888",
                  size: "sm",
                  weight: "bold",
                  align: "center"
                },
                {
                  type: "text",
                  text: `$${result.average.toLocaleString()}`,
                  size: "4xl",
                  color: "#FF3B30",
                  weight: "bold",
                  align: "center",
                  margin: "md"
                },
                {
                  type: "separator",
                  margin: "xxl",
                  color: "#eeeeee"
                },
                {
                  type: "box",
                  layout: "vertical",
                  margin: "xxl",
                  spacing: "md",
                  contents: [
                    {
                      type: "box",
                      layout: "horizontal",
                      contents: [
                        {
                          type: "text",
                          text: "累積營業總額",
                          size: "sm",
                          color: "#555555"
                        },
                        {
                          type: "text",
                          text: `$${result.total.toLocaleString()}`,
                          size: "sm",
                          color: "#111111",
                          align: "end",
                          weight: "bold"
                        }
                      ]
                    },
                    {
                      type: "box",
                      layout: "horizontal",
                      contents: [
                        {
                          type: "text",
                          text: "已記錄天數",
                          size: "sm",
                          color: "#555555"
                        },
                        {
                          type: "text",
                          text: `${result.days} 天`,
                          size: "sm",
                          color: "#111111",
                          align: "end",
                          weight: "bold"
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            footer: {
              type: "box",
              layout: "vertical",
              paddingAll: "16px",
              backgroundColor: "#fafafa",
              contents: [
                {
                  type: "button",
                  style: "primary",
                  color: "#00B900",
                  height: "sm",
                  action: {
                    type: "uri",
                    label: "📈 查看詳細營收圖表",
                    uri: "https://liff.line.me/2010481539-ovIpcDDA?page=chart&v=" + new Date().getTime()
                  }
                },
                {
                  type: "text",
                  text: "LINE Bookkeeping System",
                  color: "#cccccc",
                  size: "xs",
                  align: "center",
                  margin: "md"
                }
              ]
            }
          }
        }
      ]
    };

    console.log("Sending Flex Message to LINE...");

    // 3. 發送至 LINE
    const lineRes = await fetch('https://api.line.me/v2/bot/message/broadcast', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_TOKEN}`
      },
      body: JSON.stringify(flexMessage)
    });

    const lineData = await lineRes.text();
    if (lineRes.ok) {
      console.log("Successfully sent Flex Message!");
    } else {
      console.error("LINE API Error:", lineRes.status, lineData);
      process.exit(1);
    }
  } catch (err) {
    console.error("Script execution failed:", err);
    process.exit(1);
  }
}

run();
