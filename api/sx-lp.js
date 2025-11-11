const { fetchWithTimeout, validateParams, escapeJson } = require('./utils'); // 假设工具函数抽离到utils.js中，若在同一文件则直接使用

// ===================== 全局配置 =====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email'], // 必要参数
  SUPPORTED_METHODS: ['GET', 'POST'], // 支持的请求方法
  REQUEST_TIMEOUT: 10000, // 请求超时10秒
};

// ===================== 核心函数 =====================
/**
 * 刷新微软账户的令牌（获取新的access_token和refresh_token）
 * @param {string} refresh_token - 旧的刷新令牌
 * @param {string} client_id - Azure应用客户端ID
 * @returns {Promise<{ access_token: string, refresh_token: string }>} 新的令牌对
 */
async function refreshMicrosoftToken(refresh_token, client_id) {
  try {
    const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'scope': 'https://graph.microsoft.com/.default' // 申请默认权限
      }).toString()
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP错误！状态码：${response.status}，响应：${errorText}`);
    }

    const data = await response.json();
    if (!data.access_token || !data.refresh_token) {
      throw new Error('未获取到有效令牌对');
    }
    return {
      access_token: data.access_token,
      refresh_token: data.refresh_token
    };
  } catch (error) {
    throw new Error(`刷新令牌失败：${error.message}`);
  }
}

// ===================== 主入口函数 =====================
module.exports = async (req, res) => {
  try {
    // 1. 限制请求方法
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({
        code: 405,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    // 2. 可选密码认证（与邮件服务保持一致）
    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        code: 4010,
        error: '认证失败，请提供有效的验证密码'
      });
    }

    // 3. 提取并校验必要参数
    const params = isGet ? req.query : req.body;
    const { refresh_token, client_id, email } = params;
    const missingParams = CONFIG.REQUIRED_PARAMS.filter(key => !params[key]);

    if (missingParams.length > 0) {
      return res.status(400).json({
        code: 4001,
        error: `缺少必要参数：${missingParams.join('、')}`
      });
    }

    // 4. 校验参数格式
    const paramError = validateParams({ email, refresh_token, client_id });
    if (paramError) {
      return res.status(400).json({
        code: 4002,
        error: paramError.message
      });
    }

    // 5. 发起令牌刷新请求
    const tokenPair = await refreshMicrosoftToken(refresh_token, client_id);

    // 6. 返回成功响应
    return res.status(200).json({
      code: 200,
      message: '令牌刷新成功',
      data: {
        email,
        access_token: tokenPair.access_token,
        refresh_token: tokenPair.refresh_token,
        timestamp: new Date().toISOString()
      }
    });

  } catch (error) {
    // 统一错误分类响应
    let statusCode = 500;
    let errorCode = 5000;

    if (error.message.includes('HTTP错误！状态码：401')) {
      statusCode = 401;
      errorCode = 4011;
      error.message = '旧refresh_token已失效，请重新获取授权';
    } else if (error.message.includes('请求超时')) {
      statusCode = 504;
      errorCode = 5041;
    }

    res.status(statusCode).json({
      code: errorCode,
      error: `服务器错误：${error.message}`
    });
  }
};
