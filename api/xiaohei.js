// ===================== 全局配置与工具函数 =====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  REQUEST_TIMEOUT: 10000, // 请求超时10秒
  SUPPORTED_METHODS: ['GET', 'POST'], // 支持的请求方法
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email'], // 仅保留令牌刷新必要参数
  TOKEN_EXPIRE_BUFFER: 300, // 令牌过期缓冲时间（5分钟）
  REFRESH_RETRY_COUNT: 1 // 令牌刷新重试次数
};

// 令牌缓存（内存级，重启后失效）
const tokenCache = new Map();

// 请求超时封装
async function fetchWithTimeout(url, options = {}, timeout = CONFIG.REQUEST_TIMEOUT) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeout);
  try {
    const response = await fetch(url, { ...options, signal: controller.signal });
    clearTimeout(timeoutId);
    return response;
  } catch (error) {
    clearTimeout(timeoutId);
    throw new Error(error.name === "AbortError" ? "请求超时（超过10秒）" : error.message);
  }
}

// 参数校验（仅保留令牌相关校验）
function validateParams(params) {
  const { email } = params;
  const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailReg.test(email)) return new Error("邮箱格式无效，请输入正确的邮箱地址");
  if (params.refresh_token?.length < 50) return new Error("refresh_token格式无效");
  if (params.client_id?.length < 10) return new Error("client_id格式无效");
  return null;
}

// ===================== 令牌刷新核心函数 =====================
/**
 * 统一刷新令牌（支持重试）
 * @param {string} refresh_token - 刷新令牌
 * @param {string} client_id - 客户端ID
 * @param {string} email - 关联邮箱（作为缓存key）
 * @returns {object} { access_token, expires_in, expireTime }
 */
async function refreshAccessToken(refresh_token, client_id, email, retryCount = 0) {
  try {
    const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id,
        grant_type: 'refresh_token',
        refresh_token,
        scope: 'https://graph.microsoft.com/.default' // 保持原授权范围
      }).toString()
    });

    if (!response.ok) {
      const errorText = await response.text();
      // 400/401 直接判定refresh_token失效
      if ([400, 401].includes(response.status)) {
        throw new Error(`refresh_token已失效，请重新获取授权：${errorText}`);
      }
      throw new Error(`HTTP错误！状态码：${response.status}，响应：${errorText}`);
    }

    const data = await response.json();
    if (!data.access_token) throw new Error("刷新令牌失败，未返回有效access_token");

    // 计算过期时间（当前时间 + 有效期 - 缓冲时间）
    const expireTime = Date.now() + (data.expires_in * 1000) - CONFIG.TOKEN_EXPIRE_BUFFER * 1000;
    // 缓存令牌（按邮箱区分，避免多用户冲突）
    tokenCache.set(email, {
      access_token: data.access_token,
      expireTime,
      refresh_token,
      client_id
    });

    console.log(`【令牌刷新成功】邮箱：${email}，有效期至：${new Date(expireTime).toLocaleString()}`);
    return {
      access_token: data.access_token,
      expires_in: data.expires_in, // 原始有效期（秒）
      expireTime, // 实际过期时间戳（毫秒）
      expireTimeStr: new Date(expireTime).toLocaleString() // 格式化过期时间
    };
  } catch (error) {
    // 重试逻辑
    if (retryCount < CONFIG.REFRESH_RETRY_COUNT) {
      console.log(`【令牌刷新失败，重试${retryCount + 1}次】`, error.message);
      return refreshAccessToken(refresh_token, client_id, email, retryCount + 1);
    }
    throw new Error(`令牌刷新失败（已重试${CONFIG.REFRESH_RETRY_COUNT}次）：${error.message}`);
  }
}

/**
 * 获取有效令牌（优先用缓存，过期自动刷新）
 * @param {string} refresh_token - 刷新令牌
 * @param {string} client_id - 客户端ID
 * @param {string} email - 关联邮箱
 * @returns {object} 有效令牌信息
 */
async function getValidAccessToken(refresh_token, client_id, email) {
  const cachedToken = tokenCache.get(email);

  // 缓存存在且未过期，直接返回
  if (cachedToken && Date.now() < cachedToken.expireTime) {
    return {
      access_token: cachedToken.access_token,
      expires_in: Math.ceil((cachedToken.expireTime - Date.now()) / 1000), // 剩余有效期（秒）
      expireTime: cachedToken.expireTime,
      expireTimeStr: new Date(cachedToken.expireTime).toLocaleString(),
      fromCache: true // 标记来自缓存
    };
  }

  // 缓存过期或不存在，重新刷新
  const freshToken = await refreshAccessToken(refresh_token, client_id, email);
  return { ...freshToken, fromCache: false };
}

// ===================== 主入口函数（仅处理令牌刷新请求） =====================
module.exports = async (req, res) => {
  try {
    // 1. 限制请求方法
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({
        code: 405,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    // 2. 密码认证（保留原有逻辑，可根据需求删除）
    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        code: 4010,
        error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
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
    const paramError = validateParams(params);
    if (paramError) {
      return res.status(400).json({
        code: 4002,
        error: paramError.message
      });
    }

    // 5. 核心：获取有效令牌（自动刷新）
    const tokenInfo = await getValidAccessToken(refresh_token, client_id, email);

    // 6. 返回结果
    return res.status(200).json({
      code: 200,
      message: tokenInfo.fromCache ? '从缓存获取有效令牌' : '令牌刷新成功',
      data: tokenInfo
    });

  } catch (error) {
    // 统一错误分类响应
    let statusCode = 500;
    let errorCode = 5000;

    if (error.message.includes('refresh_token已失效')) {
      statusCode = 401;
      errorCode = 4013;
    } else if (error.message.includes('HTTP错误！状态码：401')) {
      statusCode = 401;
      errorCode = 4011;
      error.message = '认证失效，请检查refresh_token';
    } else if (error.message.includes('请求超时')) {
      statusCode = 504;
      errorCode = 5041;
    }

    res.status(statusCode).json({
      code: errorCode,
      error: `令牌刷新失败：${error.message}`
    });
  }
};
