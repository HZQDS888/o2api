// 先安装依赖：npm install axios （比fetch更稳定，容错性强）
const axios = require('axios');

module.exports = async (req, res) => {
  // 主服务必须配置中间件（否则参数解析失败）
  // app.use(express.json());
  // app.use(express.urlencoded({ extended: true }));

  // 1. 保留原认证逻辑
  const { send_password } = req.method === 'GET' ? req.query : req.body;
  const expectedPassword = process.env.SEND_PASSWORD;
  if (send_password !== expectedPassword && expectedPassword) {
    return res.status(401).json({
      code: 401,
      error: '认证失败，请提供有效的 send_password',
      error_type: 'auth_failed',
      tip: ''
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
    // 3. 保留原参数提取+校验，优化容错
    const params = req.method === 'GET' ? req.query : req.body;
    const { refresh_token = '', client_id = '', email = '' } = params;

    if (!refresh_token.trim() || !client_id.trim() || !email.trim()) {
      return res.status(400).json({
        code: 4001,
        error: '缺少必要参数：refresh_token、client_id、email',
        error_type: 'missing_params',
        tip: '请检查参数是否完整，格式为：email=xxx&client_id=xxx&refresh_token=xxx'
      });
    }

    const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailReg.test(email)) {
      return res.status(400).json({
        code: 4002,
        error: '邮箱格式无效，请输入正确的邮箱地址',
        error_type: 'invalid_email',
        tip: '支持outlook.com、hotmail.com等微软生态邮箱'
      });
    }

    // 4. 核心：刷新令牌（替换为axios，优化错误捕获）
    const tokenData = await refreshTokenForGraphAPI(refresh_token, client_id);

    // 5. 保留原成功响应格式，补充细节
    res.status(200).json({
      code: 200,
      message: '令牌刷新成功（支持 Graph API 调用）',
      error_type: null,
      data: {
        email,
        access_token: tokenData.access_token,
        refresh_token: tokenData.new_refresh_token,
        scope: tokenData.scope.split(' '),
        expires_in: tokenData.expires_in || 3600, // 兼容微软返回字段
        timestamp: new Date().toISOString(),
        next_refresh_time: new Date(Date.now() + 3500 * 1000).toISOString() // 提示下次刷新时间
      },
      tip: '请保存新的refresh_token，下次刷新需使用该令牌'
    });

  } catch (error) {
    // 6. 核心优化：对齐Python工具的精准错误分类
    console.error(`[${new Date().toISOString()}] 令牌刷新失败：`, error.message);
    let statusCode = 500;
    let errorCode = 5000;
    let errorMsg = '服务器错误：刷新令牌失败';
    let errorType = 'unknown_error';
    let tip = '请查看日志获取详细信息';

    // 解析微软返回的错误（核心修复：细化错误类型）
    if (error.message.includes('HTTP error! status: 400')) {
      const microsoftError = error.message.match(/response: ({.*})/);
      if (microsoftError && microsoftError[1]) {
        try {
          const errData = JSON.parse(microsoftError[1]);
          const errDesc = errData.error_description || '';

          // 对齐Python工具的错误分类
          if (errData.error === 'invalid_grant' || errDesc.includes('refresh_token') || errDesc.includes('expired')) {
            statusCode = 401;
            errorCode = 4011;
            errorMsg = 'refresh_token 无效或已过期';
            errorType = 'invalid_token';
            tip = '请重新发起 OAuth2 授权流程，获取新的 refresh_token';
          } else if (errData.error === 'invalid_client' || errDesc.includes('application with identifier')) {
            statusCode = 400;
            errorCode = 4003;
            errorMsg = 'client_id 无效或 Azure 应用配置错误';
            errorType = 'invalid_client_id';
            tip = '检查 Azure 应用的client_id是否正确，且已启用"允许公共客户端流"';
          } else if (errData.error === 'invalid_scope' || errDesc.includes('scope')) {
            statusCode = 400;
            errorCode = 4005;
            errorMsg = '权限范围无效';
            errorType = 'invalid_scope';
            tip = 'Azure 应用需配置 offline_access + Graph API 权限（如Mail.Read）';
          } else if (errDesc.includes('account security interrupt') || errDesc.includes('compromised')) {
            statusCode = 403;
            errorCode = 4031;
            errorMsg = '账号存在安全风险，需手动验证';
            errorType = 'risk_account';
            tip = '请用该邮箱登录outlook.com完成安全验证后重试';
          } else if (errDesc.includes('service abuse mode')) {
            statusCode = 403;
            errorCode = 4032;
            errorMsg = '账号因滥用被锁定';
            errorType = 'locked_account';
            tip = '联系微软支持解锁，或更换账号';
          } else if (errDesc.includes('rate limit')) {
            statusCode = 429;
            errorCode = 4291;
            errorMsg = '请求过于频繁，已被限流';
            errorType = 'rate_limited';
            tip = '15分钟后再试，或降低请求频率';
          } else {
            errorMsg = `微软返回错误：${errData.error_description || errData.error}`;
            errorType = 'microsoft_api_error';
          }
        } catch (e) {
          errorMsg = '参数无效：微软拒绝请求';
          errorType = 'invalid_request';
          tip = '检查client_id和refresh_token是否匹配，且无特殊字符';
        }
      }
    } else if (error.message.includes('微软服务器未返回有效令牌对')) {
      statusCode = 400;
      errorCode = 4004;
      errorMsg = '微软服务器未返回有效令牌';
      errorType = 'no_token_returned';
      tip = '确认Azure应用已配置offline_access权限，且refresh_token未失效';
    } else if (error.message.includes('timeout') || error.message.includes('ETIMEDOUT')) {
      statusCode = 504;
      errorCode = 5041;
      errorMsg = '请求超时';
      errorType = 'request_timeout';
      tip = '服务器网络无法访问微软端点，检查防火墙/代理';
    } else if (error.message.includes('ENOTFOUND') || error.message.includes('getaddrinfo')) {
      statusCode = 503;
      errorCode = 5031;
      errorMsg = '网络连接失败';
      errorType = 'network_error';
      tip = '检查服务器网络是否能访问 https://login.microsoftonline.com';
    }

    // 最终错误响应（保留原格式，新增error_type字段）
    res.status(statusCode).json({
      code: errorCode,
      error: errorMsg,
      error_type: errorType,
      tip: tip,
      details: process.env.NODE_ENV === 'development' ? error.message : '生产环境隐藏详细错误'
    });
  }
};

// 核心函数：替换fetch为axios，优化容错和错误传递
async function refreshTokenForGraphAPI(refresh_token, client_id) {
  const tokenEndpoint = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';

  try {
    if (!refresh_token.trim() || !client_id.trim()) {
      throw new Error('refresh_token 或 client_id 为空');
    }

    // 用axios替换fetch（更稳定，错误处理更完善）
    const response = await axios.post(
      tokenEndpoint,
      new URLSearchParams({
        client_id: client_id,
        grant_type: 'refresh_token',
        refresh_token: refresh_token,
        scope: 'offline_access https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send'
      }),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        timeout: 15000, // 延长超时时间到15秒
        validateStatus: null // 不自动抛出HTTP错误，手动处理
      }
    );

    // 打印响应（方便排查，生产环境可注释）
    console.log(`[${new Date().toISOString()}] 微软令牌端点返回：`, response.status, response.data);

    // 处理微软返回的非200状态码
    if (response.status !== 200) {
      throw new Error(`HTTP error! status: ${response.status}, response: ${JSON.stringify(response.data)}`);
    }

    const data = response.data;

    // 校验令牌有效性（容错处理）
    if (!data.access_token || !data.refresh_token) {
      throw new Error('微软服务器未返回有效令牌对');
    }

    return {
      access_token: data.access_token,
      new_refresh_token: data.refresh_token,
      scope: data.scope || 'offline_access https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send',
      expires_in: data.expires_in // 保留微软返回的过期时间
    };
  } catch (error) {
    // 传递原始错误信息，方便上层分类
    if (error.code === 'ECONNABORTED') {
      throw new Error('fetch failed: timeout');
    } else if (error.code === 'ENOTFOUND') {
      throw new Error('fetch failed: network error');
    } else {
      throw new Error(`核心错误：${error.message}`);
    }
  }
}
