module.exports = async (req, res) => {
  // 主服务必须配置的中间件（即使只用GET，也建议配置）
  // app.use(express.json());
  // app.use(express.urlencoded({ extended: true }));

  // 1. 认证逻辑（send_password校验）
  const { send_password } = req.method === 'GET' ? req.query : req.body;
  const expectedPassword = process.env.SEND_PASSWORD;
  if (send_password !== expectedPassword && expectedPassword) {
    return res.status(401).json({
      code: 401,
      error: '认证失败，请提供有效的 send_password'
    });
  }

  // 2. 仅支持GET/POST
  if (!['GET', 'POST'].includes(req.method)) {
    return res.status(405).json({
      code: 405,
      error: '不支持的请求方法，请使用 GET 或 POST'
    });
  }

  try {
    // 3. 提取参数（兜底空字符串，避免undefined）
    const params = req.method === 'GET' ? req.query : req.body;
    const { refresh_token = '', client_id = '', email = '' } = params;

    // 必传参数校验
    if (!refresh_token.trim() || !client_id.trim() || !email.trim()) {
      return res.status(400).json({
        code: 4001,
        error: '缺少必要参数：refresh_token、client_id、email'
      });
    }

    // 邮箱格式校验
    const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailReg.test(email)) {
      return res.status(400).json({
        code: 4002,
        error: '邮箱格式无效，请输入正确的邮箱地址'
      });
    }

    // 4. GET请求URL长度友好提示（不强制拦截，仅提示）
    if (req.method === 'GET') {
      const requestUrl = req.originalUrl || req.url || '';
      if (requestUrl.length > 2000) {
        console.warn('GET请求URL过长，可能导致失败');
        // 不拦截，让用户尝试，失败后提示用POST
      }
    }

    // 5. 核心：刷新令牌
    const tokenData = await refreshTokenForGraphAPI(refresh_token, client_id);

    // 6. 成功响应
    res.status(200).json({
      code: 200,
      message: '令牌刷新成功（支持 Graph API 调用）',
      data: {
        email,
        access_token: tokenData.access_token,
        refresh_token: tokenData.new_refresh_token,
        scope: (tokenData.scope || 'https://graph.microsoft.com/.default').split(' '),
        expires_in: 3600,
        timestamp: new Date().toISOString()
      }
    });

  } catch (error) {
    // 7. 错误处理（区分GET URL过长的情况）
    console.error('令牌刷新失败：', error);
    let statusCode = 500;
    let errorCode = 5000;
    let errorMsg = '服务器错误：刷新令牌失败';
    let tip = '';

    if (error.message.includes('HTTP error! status: 401')) {
      statusCode = 401;
      errorCode = 4011;
      errorMsg = '旧 refresh_token 已失效';
      tip = '请重新获取新的 refresh_token';
    } else if (error.message.includes('HTTP error! status: 403')) {
      statusCode = 403;
      errorCode = 4031;
      errorMsg = '权限不足';
      tip = 'Azure 应用需配置 offline_access 和 Graph API 权限';
    } else if (req.method === 'GET' && (error.message.includes('414') || error.message.includes('too long'))) {
      statusCode = 400;
      errorCode = 4004;
      errorMsg = 'GET请求URL过长导致失败';
      tip = '请改用POST请求（参数放JSON体中，无长度限制）';
    } else if (error.message.includes('Cannot read properties of undefined')) {
      statusCode = 500;
      errorCode = 5001;
      errorMsg = '服务器参数解析错误';
      tip = '已修复，请重新尝试';
    }

    res.status(statusCode).json({
      code: errorCode,
      error: errorMsg,
      tip: tip,
      details: error.message
    });
  }
};

// 核心刷新函数（修复所有undefined场景）
async function refreshTokenForGraphAPI(refresh_token, client_id) {
  const tokenEndpoint = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';

  // 超时控制
  const fetchWithTimeout = (url, options, timeout = 10000) => {
    return Promise.race([
      fetch(url, options),
      new Promise((_, reject) => setTimeout(() => reject(new Error('fetch failed: timeout')), timeout))
    ]);
  };

  try {
    if (!refresh_token.trim() || !client_id.trim()) {
      throw new Error('refresh_token 或 client_id 为空');
    }

    const response = await fetchWithTimeout(tokenEndpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: client_id,
        grant_type: 'refresh_token',
        refresh_token: refresh_token,
        scope: 'https://graph.microsoft.com/.default'
      }).toString()
    });

    if (!response.ok) {
      const errorText = await response.text().catch(() => '无详细错误');
      throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
    }

    const data = await response.json().catch(() => ({ access_token: '', refresh_token: '' }));
    if (!data.access_token || !data.refresh_token) {
      throw new Error('微软服务器未返回有效令牌对');
    }

    return {
      access_token: data.access_token,
      new_refresh_token: data.refresh_token,
      scope: data.scope || 'https://graph.microsoft.com/.default'
    };
  } catch (error) {
    throw new Error(`核心错误：${error.message}`);
  }
}
