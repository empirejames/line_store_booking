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
      to: LINE_USER_ID,
      messages: [
        {
          type: "flex",
          altText: "📊 本月營業額戰報",
          contents: {
            type: "bubble",
            body: {
              type: "box",
              layout: "vertical",
              contents: [
                {
                  type: "text",
                  text: "📊 本月營業額戰報",
                  weight: "bold",
                  color: "#1DB446",
                  size: "xl"
                },
                {
                  type: "text",
                  text: result.month,
                  size: "sm",
                  color: "#aaaaaa",
                  margin: "sm"
                },
                {
                  type: "separator",
                  margin: "xxl"
                },
                {
                  type: "box",
                  layout: "vertical",
                  margin: "xxl",
                  spacing: "sm",
                  contents: [
                    {
                      type: "box",
                      layout: "horizontal",
                      contents: [
                        {
                          type: "text",
                          text: "累積總額",
                          size: "sm",
                          color: "#555555",
                          flex: 0
                        },
                        {
                          type: "text",
                          text: `$${result.total.toLocaleString()}`,
                          size: "sm",
                          color: "#111111",
                          align: "end"
                        }
                      ]
                    },
                    {
                      type: "box",
                      layout: "horizontal",
                      contents: [
                        {
                          type: "text",
                          text: "營業天數",
                          size: "sm",
                          color: "#555555",
                          flex: 0
                        },
                        {
                          type: "text",
                          text: `${result.days} 天`,
                          size: "sm",
                          color: "#111111",
                          align: "end"
                        }
                      ]
                    }
                  ]
                },
                {
                  type: "separator",
                  margin: "xxl"
                },
                {
                  type: "box",
                  layout: "horizontal",
                  margin: "md",
                  contents: [
                    {
                      type: "text",
                      text: "🔥 每日平均",
                      size: "md",
                      color: "#ff5551",
                      weight: "bold"
                    },
                    {
                      type: "text",
                      text: `$${result.average.toLocaleString()}`,
                      size: "xl",
                      color: "#ff5551",
                      align: "end",
                      weight: "bold"
                    }
                  ]
                }
              ]
            }
          }
        }
      ]
    };

    console.log("Sending Flex Message to LINE...");

    // 3. 發送至 LINE
    const lineRes = await fetch('https://api.line.me/v2/bot/message/push', {
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
