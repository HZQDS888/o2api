// 无需 Redis 依赖！删除 ioredis，纯内存缓存（Serverless 无状态适配）
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  REQUEST_TIMEOUT: 10000,
  SUPPORTED_METHODS: ['GET', 'POST'],
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email'],
  TOKEN_EXPIRE_BUFFER: 300, // access_token 缓冲时间（5分钟）
  REFRESH_RETRY_COUNT: 1, // 重试次数
  // 预警配置
  EXPIRE_WARNING_THRESHOLD: 3600, // access_token 预警（1小时）
  REFRESH_TOKEN_WARNING_THRESHOLD: 7 * 24 * 3600, // refresh_token 预警（7天）
  ENABLE_WARNING_LOG: true,
  REFRESH_TOKEN_EXPIRE_DAYS: 90, // refresh_token 90天有效期
};

// ===================== 工具函数（适配 Serverless）=====================
// 请求超时封装（避免无响应）
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

// 参数校验
function validateParams(params) {
  const { email } = params;
  const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailReg.test(email)) return new Error("邮箱格式无效，请输入正确的邮箱地址");
  if (params.refresh_token?.length < 50) return new Error("refresh_token格式无效");
  if (params.client_id?.length < 10) return new Error("client_id格式无效");
  return null;
}

// 内存缓存（Serverless 无状态适配：单实例有效，冷启动后重置，不影响核心功能）
const memoryCache = new Map();
const cacheTool = {
  getKey: (email) => `token:${email}`,
  set(email, tokenData) {
    const key = this.getKey(email);
    memoryCache.set(key, tokenData);
    return true;
  },
  get(email) {
    const key = this.getKey(email);
    return memoryCache.get(key) || null;
  },
  delete(email) {
    const key = this.getKey(email);
    return memoryCache.delete(key);
  }
};

// ===================== 核心令牌逻辑（保留90天续期）=====================
/**
 * 刷新令牌（滚动续期，适配 Serverless）
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
        scope: 'https://graph.microsoft.com/.default offline_access' // 必须包含
      }).toString()
    });

    if (!response.ok) {
      const errorText = await response.text();
      if ([400, 401].includes(response.status)) {
        throw new Error(`refresh_token已失效，请重新授权（需含offline_access权限）：${errorText}`);
      }
      throw new Error(`HTTP错误！状态码：${response.status}，响应：${errorText}`);
    }

    const data = await response.json();
    if (!data.access_token) throw new Error("未返回有效access_token");

    // 滚动续期：获取新的refresh_token
    const newRefreshToken = data.refresh_token || refresh_token;
    // 计算有效期
    const refreshTokenExpireTime = Date.now() + (CONFIG.REFRESH_TOKEN_EXPIRE_DAYS * 24 * 60 * 60 * 1000);
    const accessTokenExpireTime = Date.now() + (data.expires_in * 1000) - CONFIG.TOKEN_EXPIRE_BUFFER * 1000;

    const tokenData = {
      access_token: data.access_token,
      accessTokenExpireTime,
      refresh_token: newRefreshToken,
      refreshTokenExpireTime,
      lastRefreshTime: Date.now()
    };

    // 缓存（Serverless 单实例有效，不影响续期）
    cacheTool.set(email, tokenData);
    console.log(`【刷新成功】邮箱：${email} | access_token 1小时 | refresh_token 90天`);

    return {
      access_token: data.access_token,
      accessTokenExpireTime,
      accessTokenExpireTimeStr: new Date(accessTokenExpireTime).toLocaleString(),
      refresh_token: newRefreshToken,
      refreshTokenExpireTime,
      refreshTokenExpireTimeStr: new Date(refreshTokenExpireTime).toLocaleString(),
      refreshTokenRemainingTime: Math.ceil((refreshTokenExpireTime - Date.now()) / 3600000),
    };
  } catch (error) {
    if (retryCount < CONFIG.REFRESH_RETRY_COUNT) {
      console.log(`【重试刷新】第${retryCount + 1}次`, error.message);
      return refreshAccessToken(refresh_token, client_id, email, retryCount + 1);
    }
    throw new Error(`刷新失败（重试${CONFIG.REFRESH_RETRY_COUNT}次）：${error.message}`);
  }
}

/**
 * 获取有效令牌（含预警）
 */
async function getValidAccessToken(refresh_token, client_id, email) {
  const cachedToken = cacheTool.get(email);

  // 缓存有效则返回
  if (cachedToken && Date.now() < cachedToken.accessTokenExpireTime) {
    const accessTokenRemaining = Math.ceil((cachedToken.accessTokenExpireTime - Date.now()) / 3600000);
    const refreshTokenRemaining = Math.ceil((cachedToken.refreshTokenExpireTime - Date.now()) / 3600000);
    
    const tokenInfo = {
      access_token: cachedToken.access_token,
      accessTokenExpireTimeStr: new Date(cachedToken.accessTokenExpireTime).toLocaleString(),
      accessTokenRemainingHours: accessTokenRemaining,
      refresh_token: cachedToken.refresh_token,
      refreshTokenExpireTimeStr: new Date(cachedToken.refreshTokenExpireTime).toLocaleString(),
      refreshTokenRemainingDays: Math.ceil(refreshTokenRemaining / 24),
      fromCache: true,
      warnings: []
    };

    // 预警逻辑
    if (accessTokenRemaining < CONFIG.EXPIRE_WARNING_THRESHOLD / 3600) {
      tokenInfo.warnings.push(`access_token剩余${accessTokenRemaining}小时`);
    }
    if (refreshTokenRemaining < CONFIG.REFRESH_TOKEN_WARNING_THRESHOLD / 3600) {
      tokenInfo.warnings.push(`refresh_token剩余${Math.ceil(refreshTokenRemaining / 24)}天，需重新授权`);
    }

    if (tokenInfo.warnings.length > 0 && CONFIG.ENABLE_WARNING_LOG) {
      console.warn(`【预警】邮箱：${email} | ${tokenInfo.warnings.join(' | ')}`);
    }

    return tokenInfo;
  }

  // 缓存失效，重新刷新
  const freshToken = await refreshAccessToken(refresh_token, client_id, email);
  return {
    ...freshToken,
    fromCache: false,
    warnings: []
  };
}

// ===================== 主函数（Vercel Serverless 适配）=====================
module.exports = async (req, res) => {
  // 强制设置响应头（避免跨域/编码问题）
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Access-Control-Allow-Origin', '*'); // 按需调整跨域配置
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');

  try {
    // 处理 OPTIONS 预检请求
    if (req.method === 'OPTIONS') {
      return res.status(200).end();
    }

    // 1. 限制请求方法
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({
        code: 405,
        error: `不支持的方法，请用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    // 2. 提取参数（适配 GET/POST）
    const params = req.method === 'GET' ? req.query : req.body;
    const { refresh_token, client_id, email, password } = params;

    // 3. 密码认证（可选）
    const expectedPassword = process.env.PASSWORD;
    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        code: 4010,
        error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
      });
    }

    // 4. 校验必要参数
    const missingParams = CONFIG.REQUIRED_PARAMS.filter(key => !params[key]);
    if (missingParams.length > 0) {
      return res.status(400).json({
        code: 4001,
        error: `缺少参数：${missingParams.join('、')}`
      });
    }

    // 5. 校验参数格式
    const paramError = validateParams(params);
    if (paramError) {
      return res.status(400).json({
        code: 4002,
        error: paramError.message
      });
    }

    // 6. 核心逻辑：获取有效令牌
    const tokenInfo = await getValidAccessToken(refresh_token, client_id, email);

    // 7. 返回结果
    return res.status(200).json({
      code: 200,
      message: tokenInfo.fromCache ? '从缓存获取令牌' : '令牌刷新成功（90天续期）',
      hasWarnings: tokenInfo.warnings.length > 0,
      warnings: tokenInfo.warnings,
      data: {
        access_token: tokenInfo.access_token,
        accessTokenExpireTimeStr: tokenInfo.accessTokenExpireTimeStr,
        accessTokenRemainingHours: tokenInfo.accessTokenRemainingHours,
        refresh_token: tokenInfo.refresh_token, // 新的refresh_token（90天）
        refreshTokenExpireTimeStr: tokenInfo.refreshTokenExpireTimeStr,
        refreshTokenRemainingDays: tokenInfo.refreshTokenRemainingDays,
        fromCache: tokenInfo.fromCache
      }
    });

  } catch (error) {
    // 统一错误处理（避免未捕获Promise导致崩溃）
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

    console.error(`【崩溃错误】`, error);
    return res.status(statusCode).json({
      code: errorCode,
      error: `令牌刷新失败：${error.message}`
    });
  }
};
