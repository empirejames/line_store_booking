const https = require('https');

const GAS_URL = process.env.GAS_URL;
const LINE_TOKEN = process.env.LINE_TOKEN;
const LINE_USER_ID = process.env.LINE_USER_ID;

if (!GAS_URL || !LINE_TOKEN || !LINE_USER_ID) {
  console.error("Missing environment variables.");
  process.exit(1);
}

// 1. 取得每日平均業績
https.get(`${GAS_URL}?action=getAverage`, (res) => {
  let data = '';

  res.on('data', (chunk) => {
    data += chunk;
  });

  res.on('end', () => {
    try {
      const result = JSON.parse(data);
      if (result.status !== 'success') {
        console.error("Error from GAS:", result.message);
        process.exit(1);
      }
      
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

      // 3. 發送至 LINE
      const postData = JSON.stringify(flexMessage);
      const options = {
        hostname: 'api.line.me',
        path: '/v2/bot/message/push',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${LINE_TOKEN}`,
          'Content-Length': Buffer.byteLength(postData)
        }
      };

      const req = https.request(options, (lineRes) => {
        let lineData = '';
        lineRes.on('data', (chunk) => { lineData += chunk; });
        lineRes.on('end', () => {
          if (lineRes.statusCode === 200) {
            console.log("Successfully sent Flex Message!");
          } else {
            console.error("LINE API Error:", lineRes.statusCode, lineData);
            process.exit(1);
          }
        });
      });

      req.on('error', (e) => {
        console.error("Request to LINE API failed:", e);
        process.exit(1);
      });

      req.write(postData);
      req.end();

    } catch (e) {
      console.error("Failed to parse GAS response:", e);
      process.exit(1);
    }
  });
}).on('error', (e) => {
  console.error("Request to GAS failed:", e);
  process.exit(1);
});
