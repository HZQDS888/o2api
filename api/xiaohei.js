// 需先安装依赖：npm install ioredis
const Redis = require('ioredis');

// ===================== 全局配置（新增refresh_token续期相关）=====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  REQUEST_TIMEOUT: 10000,
  SUPPORTED_METHODS: ['GET', 'POST'],
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email'],
  TOKEN_EXPIRE_BUFFER: 300, // access_token过期缓冲（5分钟）
  REFRESH_RETRY_COUNT: 1, // 刷新重试次数
  // Redis配置
  REDIS_CONFIG: {
    host: process.env.REDIS_HOST || 'localhost',
    port: process.env.REDIS_PORT || 6379,
    password: process.env.REDIS_PASSWORD || '',
    db: process.env.REDIS_DB || 0,
    connectTimeout: 5000,
  },
  // 预警配置
  EXPIRE_WARNING_THRESHOLD: 3600, // access_token预警阈值（1小时）
  REFRESH_TOKEN_WARNING_THRESHOLD: 7 * 24 * 3600, // refresh_token预警阈值（7天）
  ENABLE_WARNING_LOG: true,
  // refresh_token配置（固定90天有效期）
  REFRESH_TOKEN_EXPIRE_DAYS: 90, // Microsoft默认refresh_token有效期
};

// ===================== 工具函数 + Redis初始化 =====================
let redisClient;
let useMemoryCache = false;

// 初始化Redis连接
async function initRedis() {
  try {
    redisClient = new Redis(CONFIG.REDIS_CONFIG);
    await redisClient.ping();
    console.log('【Redis连接成功】已启用Redis持久化存储令牌');
    useMemoryCache = false;
  } catch (error) {
    console.error('【Redis连接失败】降级为内存缓存：', error.message);
    useMemoryCache = true;
    global.memoryTokenCache = new Map();
  }
}
initRedis().catch(err => console.error('Redis初始化异常：', err));

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

// 参数校验
function validateParams(params) {
  const { email } = params;
  const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailReg.test(email)) return new Error("邮箱格式无效，请输入正确的邮箱地址");
  if (params.refresh_token?.length < 50) return new Error("refresh_token格式无效");
  if (params.client_id?.length < 10) return new Error("client_id格式无效");
  return null;
}

// 缓存操作工具
const cacheTool = {
  getKey: (email) => `token:${email}`,
  async set(email, tokenData) {
    const key = this.getKey(email);
    if (useMemoryCache) {
      global.memoryTokenCache.set(key, tokenData);
      return true;
    }
    // Redis存储：access_token过期时间 + refresh_token 90天有效期（取较短者）
    const ttl = Math.min(
      Math.ceil((tokenData.expireTime - Date.now()) / 1000),
      CONFIG.REFRESH_TOKEN_EXPIRE_DAYS * 24 * 3600
    );
    await redisClient.set(key, JSON.stringify(tokenData), 'EX', ttl);
    return true;
  },
  async get(email) {
    const key = this.getKey(email);
    if (useMemoryCache) {
      return global.memoryTokenCache.get(key) || null;
    }
    const data = await redisClient.get(key);
    return data ? JSON.parse(data) : null;
  },
  async delete(email) {
    const key = this.getKey(email);
    if (useMemoryCache) {
      return global.memoryTokenCache.delete(key);
    }
    return await redisClient.del(key) > 0;
  }
};

// ===================== 核心功能（refresh_token 90天滚动续期）=====================
/**
 * 刷新令牌（支持refresh_token滚动续期，有效期90天）
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
        // 必须包含offline_access，否则无法获取可续期的refresh_token
        scope: 'https://graph.microsoft.com/.default offline_access'
      }).toString()
    });

    if (!response.ok) {
      const errorText = await response.text();
      if ([400, 401].includes(response.status)) {
        throw new Error(`refresh_token已失效，请重新获取授权（需包含offline_access权限）：${errorText}`);
      }
      throw new Error(`HTTP错误！状态码：${response.status}，响应：${errorText}`);
    }

    const data = await response.json();
    if (!data.access_token) throw new Error("刷新令牌失败，未返回有效access_token");

    // 关键：获取Microsoft返回的新refresh_token（滚动续期，重新计算90天）
    const newRefreshToken = data.refresh_token || refresh_token; // 兼容部分场景未返回新token的情况
    // 计算refresh_token过期时间（90天）
    const refreshTokenExpireTime = Date.now() + (CONFIG.REFRESH_TOKEN_EXPIRE_DAYS * 24 * 60 * 60 * 1000);
    // 计算access_token过期时间（1小时 - 缓冲时间）
    const accessTokenExpireTime = Date.now() + (data.expires_in * 1000) - CONFIG.TOKEN_EXPIRE_BUFFER * 1000;

    // 存储的令牌数据（包含新的refresh_token和其90天有效期）
    const tokenData = {
      access_token: data.access_token,
      accessTokenExpireTime,
      refresh_token: newRefreshToken, // 存储新的refresh_token（滚动续期）
      refreshTokenExpireTime,
      client_id,
      // 记录刷新时间，便于追溯
      lastRefreshTime: Date.now()
    };

    // 保存到缓存（覆盖旧的refresh_token）
    await cacheTool.set(email, tokenData);
    console.log(`【令牌刷新成功】邮箱：${email} | access_token有效期1小时 | refresh_token有效期90天（滚动续期）`);

    return {
      access_token: data.access_token,
      accessTokenExpireTime,
      accessTokenExpireTimeStr: new Date(accessTokenExpireTime).toLocaleString(),
      refresh_token: newRefreshToken, // 返回新的refresh_token，建议用户备份
      refreshTokenExpireTime,
      refreshTokenExpireTimeStr: new Date(refreshTokenExpireTime).toLocaleString(),
      refreshTokenRemainingTime: Math.ceil((refreshTokenExpireTime - Date.now()) / 3600000), // 剩余小时数
    };
  } catch (error) {
    if (retryCount < CONFIG.REFRESH_RETRY_COUNT) {
      console.log(`【令牌刷新失败，重试${retryCount + 1}次】`, error.message);
      return refreshAccessToken(refresh_token, client_id, email, retryCount + 1);
    }
    throw new Error(`令牌刷新失败（已重试${CONFIG.REFRESH_RETRY_COUNT}次）：${error.message}`);
  }
}

/**
 * 获取有效令牌（含access_token自动刷新 + refresh_token过期预警）
 */
async function getValidAccessToken(refresh_token, client_id, email) {
  // 1. 从缓存获取令牌
  const cachedToken = await cacheTool.get(email);

  // 2. 缓存有效：返回并添加预警
  if (cachedToken && Date.now() < cachedToken.accessTokenExpireTime) {
    const accessTokenRemaining = Math.ceil((cachedToken.accessTokenExpireTime - Date.now()) / 3600000);
    const refreshTokenRemaining = Math.ceil((cachedToken.refreshTokenExpireTime - Date.now()) / 3600000);
    
    const tokenInfo = {
      access_token: cachedToken.access_token,
      accessTokenExpireTime: cachedToken.accessTokenExpireTime,
      accessTokenExpireTimeStr: new Date(cachedToken.accessTokenExpireTime).toLocaleString(),
      accessTokenRemaining, // access_token剩余小时数
      refresh_token: cachedToken.refresh_token,
      refreshTokenExpireTime: cachedToken.refreshTokenExpireTime,
      refreshTokenExpireTimeStr: new Date(cachedToken.refreshTokenExpireTime).toLocaleString(),
      refreshTokenRemainingTime: refreshTokenRemaining, // refresh_token剩余小时数
      fromCache: true,
      warnings: []
    };

    // 添加预警（access_token快过期）
    if (accessTokenRemaining < CONFIG.EXPIRE_WARNING_THRESHOLD / 3600) {
      tokenInfo.warnings.push(`access_token即将过期，剩余${accessTokenRemaining}小时`);
    }

    // 添加预警（refresh_token快过期，提前7天提醒）
    if (refreshTokenRemaining < CONFIG.REFRESH_TOKEN_WARNING_THRESHOLD / 3600) {
      tokenInfo.warnings.push(`refresh_token即将过期，剩余${Math.ceil(refreshTokenRemaining / 24)}天，请尽快重新授权`);
    }

    // 打印预警日志
    if (tokenInfo.warnings.length > 0 && CONFIG.ENABLE_WARNING_LOG) {
      console.warn(`【令牌预警】邮箱：${email} | ${tokenInfo.warnings.join(' | ')}`);
    }

    return tokenInfo;
  }

  // 3. 缓存失效：重新刷新（返回新的refresh_token）
  const freshToken = await refreshAccessToken(refresh_token, client_id, email);
  return {
    ...freshToken,
    fromCache: false,
    warnings: []
  };
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

    // 2. 密码认证（可按需删除）
    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        code: 4010,
        error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
      });
    }

    // 3. 校验必要参数
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

    // 5. 核心：获取有效令牌（含90天续期+双预警）
    const tokenInfo = await getValidAccessToken(refresh_token, client_id, email);

    // 6. 返回结果（含新的refresh_token和有效期信息）
    return res.status(200).json({
      code: 200,
      message: tokenInfo.fromCache ? '从缓存获取有效令牌' : '令牌刷新成功（refresh_token已滚动续期）',
      hasWarnings: tokenInfo.warnings.length > 0,
      warnings: tokenInfo.warnings,
      data: {
        access_token: tokenInfo.access_token,
        accessTokenExpireTimeStr: tokenInfo.accessTokenExpireTimeStr,
        accessTokenRemainingHours: tokenInfo.accessTokenRemaining,
        refresh_token: tokenInfo.refresh_token, // 返回新的refresh_token，建议备份
        refreshTokenExpireTimeStr: tokenInfo.refreshTokenExpireTimeStr,
        refreshTokenRemainingDays: Math.ceil(tokenInfo.refreshTokenRemainingTime / 24),
        fromCache: tokenInfo.fromCache
      }
    });

  } catch (error) {
    // 统一错误处理
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
    } else if (error.message.includes('Redis')) {
      errorCode = 5004;
      error.message = `Redis操作失败：${error.message}`;
    }

    res.status(statusCode).json({
      code: errorCode,
      error: `令牌刷新失败：${error.message}`
    });
  }
};
