const fetch = require('node-fetch'); // 确保安装依赖：npm install node-fetch

module.exports = async (req, res) => {
  // 1. 保留原认证逻辑：校验访问密码
  const { send_password } = req.method === 'GET' ? req.query : req.body;
  const expectedPassword = process.env.SEND_PASSWORD;
  if (send_password !== expectedPassword && expectedPassword) {
    return res.status(401).json({
      code: 401,
      error: '认证失败，请提供有效凭证或联系管理员。',
      error_type: 'auth_failed',
      tip: '检查SEND_PASSWORD环境变量是否配置正确'
    });
  }

  // 2. 保留原请求方法限制
  if (!['GET', 'POST'].includes(req.method)) {
    return res.status(405).json({
      code: 405,
      error: '不支持的请求方法，请使用 GET 或 POST',
      error_type: 'method_not_allowed',
      tip: ''
    });
  }

  try {
    // 3. 提取并校验参数（移除邮件相关参数，保留令牌必要参数）
    const params = req.method === 'GET' ? req.query : req.body;
    const { refresh_token, client_id, email } = params;

    // 校验令牌刷新必需参数
    if (!refresh_token || !client_id || !email) {
      return res.status(400).json({
        code: 400,
        error: '缺少必要参数',
        error_type: 'missing_params',
        tip: '请确保携带refresh_token（旧令牌）、client_id（Azure应用ID）、email（关联邮箱）'
      });
    }

    // 4. 核心逻辑：刷新微软令牌（替换原邮件发送逻辑）
    const tokenData = await refreshMicrosoftToken(refresh_token, client_id);

    // 5. 成功响应（返回新令牌信息）
    return res.status(200).json({
      code: 200,
      message: '令牌刷新成功',
      error_type: null,
      data: {
        email: email,
        access_token: tokenData.access_token,
        refresh_token: tokenData.new_refresh_token, // 新的refresh_token（下次刷新需使用）
        scope: tokenData.scope.split(' '),
        expires_in: tokenData.expires_in || 3600,
        timestamp: new Date().toISOString(),
        next_refresh_time: new Date(Date.now() + 3500 * 1000).toISOString() // 推荐下次刷新时间
      },
      tip: '请妥善保存新的refresh_token，原令牌将逐步失效'
    });

  } catch (error) {
    // 6. 精准错误分类（与之前优化逻辑一致）
    console.error('令牌刷新失败:', error.message);
    let statusCode = 500;
    let errorType = 'unknown_error';
    let tip = '请查看日志获取详细信息';

    // 解析微软API返回的错误
    if (error.message.includes('HTTP error! status: 400')) {
      const microsoftError = error.message.match(/response: ({.*})/);
      if (microsoftError && microsoftError[1]) {
        try {
          const errData = JSON.parse(microsoftError[1]);
          switch (errData.error) {
            case 'invalid_grant':
              statusCode = 401;
              errorType = 'invalid_token';
              tip = 'refresh_token无效或已过期，请重新发起OAuth2授权流程';
              break;
            case 'invalid_client':
              statusCode = 400;
              errorType = 'invalid_client_id';
              tip = 'client_id无效或Azure应用未启用"允许公共客户端流"';
              break;
            case 'invalid_scope':
              statusCode = 400;
              errorType = 'invalid_scope';
              tip = 'Azure应用需配置offline_access及Graph API相关权限';
              break;
            default:
              errorType = 'microsoft_api_error';
              tip = `微软返回错误：${errData.error_description || errData.error}`;
          }
        } catch (e) {
          errorType = 'invalid_request';
          tip = '参数格式错误或微软响应解析失败';
        }
      }
    } else if (error.message.includes('timeout')) {
      statusCode = 504;
      errorType = 'request_timeout';
      tip = '请求微软服务器超时，检查网络连通性';
    } else if (error.message.includes('ENOTFOUND')) {
      statusCode = 503;
      errorType = 'network_error';
      tip = '无法连接到微软令牌服务器，检查防火墙配置';
    }

    // 错误响应格式
    return res.status(statusCode).json({
      code: statusCode,
      error: error.message,
      error_type: errorType,
      tip: tip,
      details: process.env.NODE_ENV === 'development' ? error.stack : '生产环境隐藏详细堆栈'
    });
  }
};

/**
 * 核心函数：调用微软API刷新令牌
 * @param {string} refresh_token - 旧的refresh_token
 * @param {string} client_id - Azure应用client_id
 * @returns {Promise<Object>} 包含新令牌的对象
 */
async function refreshMicrosoftToken(refresh_token, client_id) {
  const tokenEndpoint = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';
  try {
    const response = await fetch(tokenEndpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: client_id,
        grant_type: 'refresh_token',
        refresh_token: refresh_token,
        // 配置令牌权限（适配邮件/IMAP等场景）
        scope: 'offline_access https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send https://outlook.office365.com/IMAP.AccessAsUser.All'
      }).toString(),
      timeout: 15000 // 15秒超时保护
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
      scope: data.scope || 'offline_access https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send',
      expires_in: data.expires_in
    };
  } catch (error) {
    throw new Error(`令牌刷新核心错误：${error.message}`);
  }
}
