// ===================== 微软令牌刷新API（单文件版）=====================
// 依赖说明：需先执行 npm install express axios cors 安装依赖
// 使用方式：node server.js （或 PASSWORD=你的密码 node server.js 启用认证）
// 接口地址：GET /api/refresh-token?refresh_token=XXX&client_id=XXX&password=XXX（可选）
// ======================================================================

const express = require('express');
const axios = require('axios');
const cors = require('cors');

// ===================== 全局配置（参考你的代码风格：抽离常量+中文提示）=====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
  REQUEST_TIMEOUT: 15000, // 请求超时15秒
  SUPPORTED_METHODS: ['GET'], // 仅支持GET请求
  REQUIRED_PARAMS: ['refresh_token', 'client_id'], // 必要参数
  ERROR_CODES: {
    MISSING_PARAMS: 4001,
    INVALID_PARAM_FORMAT: 4002,
    UNSUPPORTED_METHOD: 4051,
    AUTH_FAILED: 4011,
    TOKEN_REFRESH_FAILED: 5001,
    REQUEST_TIMEOUT: 5041,
    SERVER_ERROR: 5000
  }
};

// ===================== 工具函数（参考你的优化：超时+校验）=====================
// 请求超时封装
async function fetchWithTimeout(url, options = {}, timeout = CONFIG.REQUEST_TIMEOUT) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeout);
  try {
    const response = await axios({
      ...options,
      url,
      signal: controller.signal
    });
    clearTimeout(timeoutId);
    return response;
  } catch (error) {
    clearTimeout(timeoutId);
    if (error.name === 'AbortError' || error.code === 'ECONNABORTED') {
      throw new Error(`请求超时（超过${timeout/1000}秒）`);
    }
    throw new Error(error.message);
  }
}

// 参数格式校验
function validateParams(params) {
  const { refresh_token, client_id } = params;
  if (refresh_token?.length < 50) return new Error("refresh_token格式无效（长度过短）");
  if (client_id?.length < 10) return new Error("client_id格式无效（长度过短）");
  return null;
}

// ===================== 核心业务函数（令牌刷新）=====================
async function refreshMicrosoftToken(refreshToken, clientId) {
  try {
    const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      data: new URLSearchParams({
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
        client_id: clientId
      }).toString()
    });

    if (response.status === 200 && response.data.refresh_token) {
      console.log(`[成功] 令牌刷新成功 - client_id: ${clientId.substring(0, 10)}...`);
      return {
        success: true,
        new_refresh_token: response.data.refresh_token,
        access_token: response.data.access_token,
        expires_in: response.data.expires_in,
        token_type: response.data.token_type
      };
    } else {
      throw new Error(`响应异常：${JSON.stringify(response.data)}`);
    }
  } catch (error) {
    let errorMsg = "令牌刷新失败";
    if (error.message.includes('invalid_grant')) errorMsg = "无效的refresh_token或client_id";
    else if (error.message.includes('unauthorized_client')) errorMsg = "client_id未授权或无效";
    else if (error.message.includes('rate_limited')) errorMsg = "请求频率限制，请稍后再试";
    else errorMsg += `：${error.message}`;
    
    console.error(`[失败] ${errorMsg} - client_id: ${clientId.substring(0, 10)}...`);
    throw new Error(errorMsg);
  }
}

// ===================== 主服务入口（参考你的统一响应格式）=====================
const app = express();
const port = process.env.PORT || 3000;

// 中间件
app.use(cors());
app.use(express.json());

// 健康检查接口
app.get('/health', (req, res) => {
  res.status(200).json({
    code: 200,
    status: 'ok',
    message: '令牌刷新API服务正常运行'
  });
});

// 核心刷新接口（GET请求）
app.get('/api/refresh-token', async (req, res) => {
  try {
    // 1. 限制请求方法
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({
        code: CONFIG.ERROR_CODES.UNSUPPORTED_METHOD,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('')}`
      });
    }

    // 2. 密码认证（可选：通过环境变量启用）
    const { password } = req.query;
    const expectedPassword = process.env.PASSWORD;
    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        code: CONFIG.ERROR_CODES.AUTH_FAILED,
        error: '认证失败，请联系小黑-QQ:113575320 购买权限再使用'
      });
    }

    // 3. 校验必要参数
    const missingParams = CONFIG.REQUIRED_PARAMS.filter(key => !req.query[key]);
    if (missingParams.length > 0) {
      return res.status(400).json({
        code: CONFIG.ERROR_CODES.MISSING_PARAMS,
        error: `缺少必要参数：${missingParams.join('、')}`
      });
    }

    // 4. 校验参数格式
    const paramError = validateParams(req.query);
    if (paramError) {
      return res.status(400).json({
        code: CONFIG.ERROR_CODES.INVALID_PARAM_FORMAT,
        error: paramError.message
      });
    }

    // 5. 执行刷新并响应
    const { refresh_token, client_id } = req.query;
    const result = await refreshMicrosoftToken(refresh_token, client_id);
    
    res.status(200).json({
      code: 200,
      message: '令牌刷新成功',
      data: result
    });

  } catch (error) {
    // 统一错误响应
    let statusCode = 500;
    let errorCode = CONFIG.ERROR_CODES.SERVER_ERROR;

    if (error.message.includes('请求超时')) {
      statusCode = 504;
      errorCode = CONFIG.ERROR_CODES.REQUEST_TIMEOUT;
    } else if (error.message.includes('无效的') || error.message.includes('未授权')) {
      statusCode = 400;
      errorCode = CONFIG.ERROR_CODES.INVALID_PARAM_FORMAT;
    }

    res.status(statusCode).json({
      code: errorCode,
      error: error.message,
      data: null
    });
  }
});

// 启动服务
app.listen(port, () => {
  console.log(`[服务启动成功] 微软令牌刷新API运行在 http://localhost:${port}`);
  console.log(`[使用说明] GET接口：/api/refresh-token?refresh_token=你的旧token&client_id=你的clientId`);
  process.env.PASSWORD && console.log(`[安全提示] 已启用密码认证，请求需携带参数：password=${process.env.PASSWORD}`);
});

module.exports = app; // 支持部署平台识别
