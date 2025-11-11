module.exports = async (req, res) => {
  // 重要提醒：主服务需配置中间件（否则POST请求体无法解析）
  // app.use(express.json());
  // app.use(express.urlencoded({ extended: true }));

  // 1. 保留原有认证逻辑（send_password 校验，可选关闭）
  const { send_password } = req.method === 'GET' ? req.query : req.body;
  const expectedPassword = process.env.SEND_PASSWORD;
  if (send_password !== expectedPassword && expectedPassword) {
    return res.status(401).json({
      code: 401,
      error: '认证失败，请提供有效的 send_password',
      tip: '若未配置 SEND_PASSWORD 环境变量，可删除该校验逻辑'
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

    // 邮箱格式校验
    const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailReg.test(email)) {
      return res.status(400).json({
        code: 4002,
        error: '邮箱格式无效，请输入正确的邮箱地址'
      });
    }

    // 关键优化：GET请求参数长度校验（避免URL超限）
    if (req.method === 'GET') {
      const urlLength = req.originalUrl.length;
      if (urlLength > 2000) { // 多数浏览器/服务器限制GET URL长度在2KB左右
        return res.status(400).json({
          code: 4004,
          error: 'GET请求URL过长（refresh_token太长），请改用POST请求',
          tip: 'POST请求无长度限制，参数放在JSON请求体中'
        });
      }
    }

    // 4. 核心逻辑：刷新令牌（适配 Graph API 权限，增加超时和错误捕获）
    const tokenData = await refreshTokenForGraphAPI(refresh_token, client_id);

    // 5. 返回成功响应
    res.status(200).json({
      code: 200,
      message: '令牌刷新成功（支持 Graph API 调用）',
      data: {
        email,
        access_token: tokenData.access_token,
        refresh_token: tokenData.new_refresh_token, // 保存新令牌用于下次刷新
        scope: tokenData.scope.split(' '),
        expires_in: 3600, // 1小时有效期
        timestamp: new Date().toISOString()
      }
    });

  } catch (error) {
    // 6. 优化错误信息，精准指引
    console.error('令牌刷新失败：', error);
    let statusCode = 500;
    let errorCode = 5000;
    let errorMsg = '服务器错误：刷新令牌失败';
    let tip = '';

    if (error.message.includes('HTTP error! status: 401')) {
      statusCode = 401;
      errorCode = 4011;
      errorMsg = '旧 refresh_token 已失效';
      tip = '请重新发起 OAuth2 授权流程获取新的 refresh_token';
    } else if (error.message.includes('HTTP error! status: 403')) {
      statusCode = 403;
      errorCode = 4031;
      errorMsg = '权限不足';
      tip = 'Azure 应用需配置 offline_access 和 Graph API 相关权限（如 Mail.Read）';
    } else if (error.message.includes('HTTP error! status: 400')) {
      statusCode = 400;
      errorCode = 4003;
      errorMsg = '参数无效';
      tip = '检查 client_id 是否正确，或 refresh_token 格式是否异常';
    } else if (error.message.includes('fetch failed')) {
      statusCode = 504;
      errorCode = 5041;
      errorMsg = '请求超时';
      tip = '检查服务器网络是否能访问 https://login.microsoftonline.com';
    }

    res.status(statusCode).json({
      code: errorCode,
      error: errorMsg,
      tip: tip,
      details: error.message
    });
  }
};

/**
 * 核心函数：调用微软 OAuth2 端点，刷新令牌（优化超时和错误处理）
 */
async function refreshTokenForGraphAPI(refresh_token, client_id) {
  const tokenEndpoint = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';

  // 超时控制函数（避免无限等待）
  const fetchWithTimeout = (url, options, timeout = 10000) => {
    return Promise.race([
      fetch(url, options),
      new Promise((_, reject) => 
        setTimeout(() => reject(new Error('fetch failed: timeout')), timeout)
      )
    ]);
  };

  try {
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
      const errorText = await response.text().catch(() => '无详细错误信息');
      throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
    }

    const data = await response.json();

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
