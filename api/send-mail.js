module.exports = async (req, res) => {
  // 主服务必须配置中间件（否则参数解析失败）
  // app.use(express.json());
  // app.use(express.urlencoded({ extended: true }));

  // 1. 认证逻辑
  const { send_password } = req.method === 'GET' ? req.query : req.body;
  const expectedPassword = process.env.SEND_PASSWORD;
  if (send_password !== expectedPassword && expectedPassword) {
    return res.status(401).json({
      code: 401,
      error: '认证失败，请提供有效的 send_password'
    });
  }

  // 2. 限制请求方法
  if (!['GET', 'POST'].includes(req.method)) {
    return res.status(405).json({
      code: 405,
      error: '不支持的请求方法，请使用 GET 或 POST'
    });
  }

  try {
    // 3. 提取并校验参数
    const params = req.method === 'GET' ? req.query : req.body;
    const { refresh_token = '', client_id = '', email = '' } = params;

    if (!refresh_token.trim() || !client_id.trim() || !email.trim()) {
      return res.status(400).json({
        code: 4001,
        error: '缺少必要参数：refresh_token、client_id、email'
      });
    }

    const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailReg.test(email)) {
      return res.status(400).json({
        code: 4002,
        error: '邮箱格式无效，请输入正确的邮箱地址'
      });
    }

    // 4. 核心：刷新令牌（增加微软返回数据打印）
    const tokenData = await refreshTokenForGraphAPI(refresh_token, client_id);

    // 5. 成功响应
    res.status(200).json({
      code: 200,
      message: '令牌刷新成功（支持 Graph API 调用）',
      data: {
        email,
        access_token: tokenData.access_token,
        refresh_token: tokenData.new_refresh_token,
        scope: tokenData.scope.split(' '),
        expires_in: 3600,
        timestamp: new Date().toISOString()
      }
    });

  } catch (error) {
    // 6. 精准解析微软返回的错误
    console.error('令牌刷新失败：', error);
    let statusCode = 500;
    let errorCode = 5000;
    let errorMsg = '服务器错误：刷新令牌失败';
    let tip = '';

    // 解析微软返回的错误（如 invalid_grant、invalid_client 等）
    if (error.message.includes('HTTP error! status: 400')) {
      // 提取微软返回的 error 字段
      const microsoftError = error.message.match(/response: ({.*})/);
      if (microsoftError && microsoftError[1]) {
        try {
          const errData = JSON.parse(microsoftError[1]);
          switch (errData.error) {
            case 'invalid_grant':
              statusCode = 401;
              errorCode = 4011;
              errorMsg = 'refresh_token 无效或已过期';
              tip = '请重新发起 OAuth2 授权流程，获取新的 refresh_token';
              break;
            case 'invalid_client':
              statusCode = 400;
              errorCode = 4003;
              errorMsg = 'client_id 无效或 Azure 应用配置错误';
              tip = '检查 Azure 应用的 client_id 是否正确，且已启用"允许公共客户端流"';
              break;
            case 'invalid_scope':
              statusCode = 400;
              errorCode = 4005;
              errorMsg = '权限范围无效';
              tip = 'Azure 应用未配置对应的 Graph API 权限';
              break;
            default:
              errorMsg = `微软返回错误：${errData.error_description || errData.error}`;
              tip = '参考 Azure 应用配置指南检查权限和认证设置';
          }
        } catch (e) {
          errorMsg = '参数无效：微软拒绝请求';
          tip = '检查 client_id 和 refresh_token 是否匹配';
        }
      }
    } else if (error.message.includes('微软服务器未返回有效令牌对')) {
      statusCode = 400;
      errorCode = 4004;
      errorMsg = '微软服务器未返回令牌';
      tip = '可能是 Azure 应用未配置 offline_access 权限，或 refresh_token 已失效';
    } else if (error.message.includes('fetch failed')) {
      statusCode = 504;
      errorCode = 5041;
      errorMsg = '请求超时';
      tip = '服务器网络无法访问微软令牌端点，检查防火墙设置';
    }

    res.status(statusCode).json({
      code: errorCode,
      error: errorMsg,
      tip: tip,
      details: error.message
    });
  }
};

// 核心函数：增加微软返回数据完整解析
async function refreshTokenForGraphAPI(refresh_token, client_id) {
  const tokenEndpoint = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';

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
        // 优化：明确指定权限，而非默认 .default
        scope: 'offline_access https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send'
      }).toString()
    });

    // 打印微软返回的完整响应（方便排查）
    const responseText = await response.text();
    console.log('微软令牌端点返回：', responseText);

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}, response: ${responseText}`);
    }

    const data = JSON.parse(responseText);

    // 校验令牌是否存在
    if (!data.access_token || !data.refresh_token) {
      throw new Error('微软服务器未返回有效令牌对');
    }

    return {
      access_token: data.access_token,
      new_refresh_token: data.refresh_token,
      scope: data.scope || 'offline_access https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send'
    };
  } catch (error) {
    throw new Error(`核心错误：${error.message}`);
  }
}
