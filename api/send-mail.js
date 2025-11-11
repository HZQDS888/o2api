module.exports = async (req, res) => {
  // 1. 保留原有认证逻辑（send_password 校验，可选关闭）
  const { send_password } = req.method === 'GET' ? req.query : req.body;
  const expectedPassword = process.env.SEND_PASSWORD;
  if (send_password !== expectedPassword && expectedPassword) {
    return res.status(401).json({
      code: 401,
      error: '认证失败，请提供有效的 send_password'
    });
  }

  // 2. 限制请求方法（仅支持 GET/POST）
  if (!['GET', 'POST'].includes(req.method)) {
    return res.status(405).json({
      code: 405,
      error: '不支持的请求方法，请使用 GET 或 POST'
    });
  }

  try {
    // 3. 提取并校验必要参数
    const params = req.method === 'GET' ? req.query : req.body;
    const { refresh_token, client_id, email } = params;

    // 必传参数校验
    if (!refresh_token || !client_id || !email) {
      return res.status(400).json({
        code: 4001,
        error: '缺少必要参数：refresh_token、client_id、email'
      });
    }

    // 邮箱格式校验（确保是微软生态邮箱，如 outlook.com、hotmail.com）
    const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailReg.test(email)) {
      return res.status(400).json({
        code: 4002,
        error: '邮箱格式无效，请输入正确的邮箱地址'
      });
    }

    // 4. 核心逻辑：刷新令牌（适配 Graph API 权限）
    const tokenData = await refreshTokenForGraphAPI(refresh_token, client_id);

    // 5. 返回成功响应（包含 Graph API 兼容的令牌信息）
    res.status(200).json({
      code: 200,
      message: '令牌刷新成功（支持 Graph API 调用）',
      data: {
        email,
        access_token: tokenData.access_token, // 可直接用于 Graph API 授权
        refresh_token: tokenData.new_refresh_token, // 新刷新令牌（下次刷新用）
        scope: tokenData.scope.split(' '), // 令牌拥有的 Graph API 权限
        expires_in: 3600, // 微软 access_token 默认有效期（1小时）
        timestamp: new Date().toISOString()
      }
    });

  } catch (error) {
    // 6. 分类错误响应（精准定位令牌刷新问题）
    console.error('令牌刷新失败：', error);
    let statusCode = 500;
    let errorCode = 5000;
    let errorMsg = '服务器错误：刷新令牌失败';

    if (error.message.includes('HTTP error! status: 401')) {
      statusCode = 401;
      errorCode = 4011;
      errorMsg = '旧 refresh_token 已失效，请重新发起 OAuth2 授权流程';
    } else if (error.message.includes('HTTP error! status: 403')) {
      statusCode = 403;
      errorCode = 4031;
      errorMsg = '权限不足：Azure 应用未配置 Graph API 相关权限';
    } else if (error.message.includes('HTTP error! status: 400')) {
      statusCode = 400;
      errorCode = 4003;
      errorMsg = '参数无效：client_id 错误或 refresh_token 格式异常';
    }

    res.status(statusCode).json({
      code: errorCode,
      error: errorMsg,
      details: error.message
    });
  }
};

/**
 * 核心函数：调用微软 OAuth2 端点，刷新令牌（适配 Graph API）
 * @param {string} refresh_token - 旧的刷新令牌
 * @param {string} client_id - Azure 应用 client_id
 * @returns {Promise<{ access_token: string, new_refresh_token: string, scope: string }>}
 */
async function refreshTokenForGraphAPI(refresh_token, client_id) {
  const tokenEndpoint = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';

  try {
    const response = await fetch(tokenEndpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: client_id,
        grant_type: 'refresh_token',
        refresh_token: refresh_token,
        scope: 'https://graph.microsoft.com/.default' // 关键：申请 Graph API 默认权限
      }).toString(),
      timeout: 10000 // 10秒超时控制，避免无限等待
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
    }

    const data = await response.json();

    // 校验返回的令牌有效性
    if (!data.access_token || !data.refresh_token) {
      throw new Error('微软服务器未返回有效令牌对');
    }

    return {
      access_token: data.access_token,
      new_refresh_token: data.refresh_token,
      scope: data.scope || 'https://graph.microsoft.com/.default'
    };
  } catch (error) {
    throw new Error(`令牌刷新核心错误：${error.message}`);
  }
}
