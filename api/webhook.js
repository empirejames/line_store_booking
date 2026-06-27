export default async function handler(req, res) {
  // 只處理 POST 請求 (LINE Webhook)
  if (req.method !== 'POST') {
    return res.status(200).json({ status: 'ok', message: 'Webhook proxy is running!' });
  }

  const body = req.body;
  const gasUrl = process.env.GAS_URL;

  if (!gasUrl) {
    console.error('GAS_URL environment variable is not set');
    return res.status(200).json({ status: 'error', message: 'GAS_URL not configured' });
  }

  try {
    // 第一步：POST 到 GAS，手動處理 302 重定向
    let response = await fetch(gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
      redirect: 'manual' // 不自動跟隨，手動處理
    });

    // 第二步：如果收到 302/301，手動跟隨重定向並保持 POST 方法
    let maxRedirects = 5;
    while (
      (response.status === 301 || response.status === 302 || response.status === 307 || response.status === 308) 
      && maxRedirects > 0
    ) {
      const location = response.headers.get('location');
      if (!location) break;

      response = await fetch(location, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
        redirect: 'manual'
      });
      maxRedirects--;
    }

    console.log('GAS response status:', response.status);
  } catch (error) {
    console.error('Proxy forwarding error:', error);
  }

  // 無論如何都回傳 200 給 LINE
  return res.status(200).json({ status: 'ok' });
}
