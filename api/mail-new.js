const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// ===================== 全局配置（统一环境变量注入）=====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  GRAPH_API_BASE_URL: 'https://graph.microsoft.com/v1.0/me/mailFolders',
  IMAP_CONFIG: {
    host: 'outlook.office365.com',
    port: 993,
    tls: true,
    // 生产环境建议移除 rejectUnauthorized: false，确保证书校验
    tlsOptions: { rejectUnauthorized: false },
    connTimeout: 10000,
    authTimeout: 10000
  },
  MAILBOX_MAP: {
    '收件箱': 'inbox',
    'inbox': 'inbox',
    '已发送': 'sentitems',
    'sentitems': 'sentitems',
    '草稿': 'draft',
    'drafts': 'draft',
    '删除邮件': 'deleteditems',
    'deleteditems': 'deleteditems',
    '垃圾邮件': 'junkemail',
    'junk': 'junkemail'
  },
  REQUEST_TIMEOUT: process.env.REQUEST_TIMEOUT || 10000,
  SUPPORTED_METHODS: ['GET', 'POST'],
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email', 'mailbox'],
  PASSWORD: process.env.PASSWORD || '' // 统一配置管理
};

// ===================== 工具函数 =====================
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

// HTML特殊字符转义（防XSS）
function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// JSON响应特殊字符转义
function escapeJson(str) {
  if (!str) return str;
  return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n');
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

// 公共OAuth Token请求函数（提取重复逻辑）
async function requestOAuthToken(refresh_token, client_id, scope = '') {
  const body = new URLSearchParams({
    client_id,
    grant_type: 'refresh_token',
    refresh_token
  });
  if (scope) body.append('scope', scope);

  const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`状态码：${response.status}，响应：${errorText}`);
  }

  const responseText = await response.text();
  return JSON.parse(responseText);
}

// ===================== 核心业务函数 =====================
// 生成邮件HTML
function generateEmailHtml(emailData) {
  const { send, subject, text, html: emailHtml, date } = emailData;
  const escapedText = escapeHtml(text || '');
  const escapedHtml = emailHtml || `<p>${escapedText.replace(/\n/g, '<br>')}</p>`;

  return `
    <!DOCTYPE html>
    <html lang="zh-CN">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>${escapeHtml(subject || '无主题邮件')}</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; line-height: 1.6; margin: 0; padding: 20px; background: #f5f5f5; }
          .email-container { max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
          .email-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid #eee; }
          .email-title { margin: 0 0 15px; color: #2d3748; }
          .email-meta { color: #4a5568; font-size: 0.9em; }
          .email-meta span { display: block; margin-bottom: 5px; }
          .email-content { color: #1a202c; }
        </style>
      </head>
      <body>
        <div class="email-container">
          <div class="email-header">
            <h1 class="email-title">${escapeHtml(subject || '无主题')}</h1>
            <div class="email-meta">
              <span><strong>发件人：</strong>${escapeHtml(send || '未知发件人')}</span>
              <span><strong>发送日期：</strong>${new Date(date).toLocaleString() || '未知日期'}</span>
            </div>
          </div>
          <div class="email-content">
            ${escapedHtml}
          </div>
        </div>
      </body>
    </html>
  `;
}

// 获取access_token
async function getAccessToken(refresh_token, client_id) {
  try {
    const data = await requestOAuthToken(refresh_token, client_id);
    return data.access_token;
  } catch (error) {
    throw new Error(`获取access_token失败：${error.message}`);
  }
}

// 生成IMAP认证字符串
const generateAuthString = (user, accessToken) => {
  const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
  return Buffer.from(authString).toString('base64');
};

// 检查Graph API权限
async function checkGraphPermission(refresh_token, client_id) {
  try {
    const data = await requestOAuthToken(refresh_token, client_id, 'https://graph.microsoft.com/.default');
    const scopes = data.scope?.split(' ') || [];
    // 精准判断是否包含目标权限
    const hasMailPermission = scopes.includes('https://graph.microsoft.com/Mail.ReadWrite');
    return {
      access_token: data.access_token,
      status: hasMailPermission
    };
  } catch (error) {
    console.error('Graph API权限检查失败：', error);
    return { access_token: '', status: false };
  }
}

// 获取文件夹最新1封邮件（Graph API）
async function getLatestEmailByGraph(access_token, mailbox) {
  if (!access_token) throw new Error("access_token不存在");

  // 强制只取最新1封，按接收时间降序排序
  const url = `${CONFIG.GRAPH_API_BASE_URL}/${mailbox}/messages?$top=1&$orderby=receivedDateTime desc`;
  const response = await fetchWithTimeout(url, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      "Authorization": `Bearer ${access_token}`
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`状态码${response.status}，响应：${errorText}`);
  }

  const responseData = await response.json();
  const emails = responseData.value || [];
  if (emails.length === 0) return null;

  // 格式化单封邮件数据
  const item = emails[0];
  return {
    send: item['from']?.['emailAddress']?.['address'] || '未知发件人',
    subject: item['subject'] || '无主题',
    text: item['bodyPreview'] || '',
    html: item['body']?.['content'] || '',
    date: item['createdDateTime'] || new Date().toISOString(),
  };
}

// 获取文件夹最新1封邮件（IMAP）
async function getLatestEmailByImap(email, access_token, mailbox) {
  return new Promise((resolve, reject) => {
    const authString = generateAuthString(email, access_token);
    const imap = new Imap({ ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString });

    imap.once("ready", async () => {
      try {
        // 打开目标文件夹
        await new Promise((resolveOpen, rejectOpen) => {
          imap.openBox(mailbox, true, (err, box) => err ? rejectOpen(err) : resolveOpen(box));
        });

        // 搜索所有邮件，取最后1封（最新）
        const results = await new Promise((resolveSearch, rejectSearch) => {
          imap.search(["ALL"], (err, results) => err ? rejectSearch(err) : resolveSearch(results));
        });

        if (results.length === 0) {
          imap.end();
          return resolve(null);
        }

        // 只获取最新1封邮件
        const latestMailId = results.slice(-1);
        const f = imap.fetch(latestMailId, { bodies: "" });

        f.on("message", async (msg) => {
          try {
            const stream = await new Promise((resolveStream) => msg.on("body", resolveStream));
            const mail = await simpleParser(stream);

            const formattedData = {
              send: escapeJson(mail.from?.text || '未知发件人'),
              subject: escapeJson(mail.subject || '无主题'),
              text: escapeJson(mail.text || ''),
              html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
              date: mail.date || new Date().toLocaleString()
            };
            resolve([formattedData]); // 统一返回数组格式
          } catch (err) {
            reject(new Error(`解析邮件失败：${err.message}`));
          }
        });

        f.once("end", () => imap.end());
      } catch (err) {
        imap.end();
        reject(new Error(`IMAP操作失败：${err.message}`));
      }
    });

    imap.once('error', (err) => {
      reject(new Error(`IMAP连接失败：${err.message}`));
    });

    imap.connect();
  });
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

    // 2. 密码认证
    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    if (CONFIG.PASSWORD && password !== CONFIG.PASSWORD) {
      return res.status(401).json({
        code: 4010,
        error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
      });
    }

    // 3. 提取并校验参数
    const params = isGet ? req.query : req.body;
    let { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;
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

    // 5. 文件夹名称标准化（支持中文/英文输入）
    const normalizedMailbox = CONFIG.MAILBOX_MAP[mailbox.toLowerCase()];
    if (!normalizedMailbox) {
      const supportedMailboxes = Object.keys(CONFIG.MAILBOX_MAP).filter(key => !/[a-z]/.test(key)).join('、');
      return res.status(400).json({
        code: 4003,
        error: `不支持的文件夹名称：${mailbox}，支持的中文文件夹：${supportedMailboxes}`
      });
    }
    mailbox = normalizedMailbox;

    // 6. 优先使用Graph API（取最新1封）
    console.log(`【开始】检查Graph API权限，目标文件夹：${mailbox}`);
    const graphResult = await checkGraphPermission(refresh_token, client_id);

    if (graphResult.status) {
      console.log(`【成功】Graph API权限通过，获取${mailbox}最新1封邮件`);
      const emailData = await getLatestEmailByGraph(graphResult.access_token, mailbox);

      if (!emailData) {
        const mailboxCN = Object.keys(CONFIG.MAILBOX_MAP).find(key => CONFIG.MAILBOX_MAP[key] === mailbox);
        return res.status(200).json({
          code: 2001,
          message: `当前“${mailboxCN}”文件夹无邮件`,
          data: null
        });
      }

      // 统一返回数组格式
      const responseData = [emailData];
      if (response_type === 'html') {
        const html = generateEmailHtml(emailData);
        return res.status(200).send(html);
      }
      return res.status(200).json({ code: 200, message: '邮件获取成功', data: responseData });
    }

    // 7. 降级使用IMAP（取最新1封）
    console.log(`【降级】Graph API权限不足，使用IMAP获取${mailbox}最新1封邮件`);
    const accessToken = await getAccessToken(refresh_token, client_id);
    const imapEmailData = await getLatestEmailByImap(email, accessToken, mailbox);

    if (!imapEmailData) {
      const mailboxCN = Object.keys(CONFIG.MAILBOX_MAP).find(key => CONFIG.MAILBOX_MAP[key] === mailbox);
      return res.status(200).json({
        code: 2001,
        message: `当前“${mailboxCN}”文件夹无邮件`,
        data: null
      });
    }

    if (response_type === 'html') {
      const html = generateEmailHtml(imapEmailData[0]);
      return res.status(200).send(html);
    }
    return res.status(200).json({ code: 200, message: '邮件获取成功', data: imapEmailData });

  } catch (error) {
    // 统一错误处理
    let statusCode = 500;
    let errorCode = 5000;

    if (error.message.includes('状态码：401')) {
      statusCode = 401;
      errorCode = 4011;
      error.message = '认证失效，请刷新refresh_token';
    } else if (error.message.includes('状态码：403')) {
      statusCode = 403;
      errorCode = 4031;
      error.message = '权限不足，需开启Mail.ReadWrite权限';
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
