const fetch = require('node-fetch');

module.exports = async (req, res) => {
  // 1. 认证逻辑：校验访问密码
  const { password } = req.method === 'GET' ? req.query : req.body;
  const expectedPassword = process.env.PASSWORD;
  if (password !== expectedPassword) {
    return res.status(401).json({
      error: '认证失败，请提供有效密码或联系管理员获取访问权限。'
    });
  }

  // 2. 提取并校验请求参数
  const params = req.method === 'GET' ? req.query : req.body;
  const { refresh_token, client_id, email } = params;

  if (!refresh_token || !client_id || !email) {
    return res.status(400).json({
      error: '缺少必要参数：refresh_token、client_id 或 email'
    });
  }

  // 3. 核心逻辑：刷新令牌
  async function refreshMicrosoftToken() {
    const tokenEndpoint = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';
    try {
      const response = await fetch(tokenEndpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: client_id,
          grant_type: 'refresh_token',
          refresh_token: refresh_token,
          scope: 'offline_access https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send'
        }).toString(),
        timeout: 15000 // 超时保护
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`微软API请求失败：状态码${response.status}，响应：${errorText}`);
      }

      const data = await response.json();
      if (!data.access_token || !data.refresh_token) {
        throw new Error('微软未返回有效令牌对');
      }

      return {
        access_token: data.access_token,
        new_refresh_token: data.refresh_token,
        scope: data.scope,
        expires_in: data.expires_in
      };
    } catch (error) {
      throw new Error(`令牌刷新失败：${error.message}`);
    }
  }

  try {
    const tokenData = await refreshMicrosoftToken();
    return res.status(200).json({
      message: '令牌刷新成功',
      data: {
        email,
        access_token: tokenData.access_token,
        refresh_token: tokenData.new_refresh_token,
        scope: tokenData.scope.split(' '),
        expires_in: tokenData.expires_in,
        timestamp: new Date().toISOString()
      }
    });
  } catch (error) {
    console.error('全局错误捕获：', error);
    return res.status(500).json({
      error: '服务器处理失败',
      details: error.message
    });
  }
};
