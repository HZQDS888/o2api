const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// ===================== 全局配置与工具函数（优化点：抽离常量+统一工具）=====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  GRAPH_API_BASE_URL: 'https://graph.microsoft.com/v1.0/me/mailFolders',
  IMAP_CONFIG: {
    host: 'outlook.office365.com',
    port: 993,
    tls: true,
    tlsOptions: { rejectUnauthorized: false },
    connTimeout: 10000, // 连接超时10秒
    authTimeout: 10000 // 认证超时10秒
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
  REQUEST_TIMEOUT: 10000, // 请求超时10秒
  SUPPORTED_METHODS: ['GET', 'POST'], // 支持的请求方法
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email', 'mailbox'] // 必要参数
};

// 请求超时封装（优化点：避免无限等待）
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

// HTML特殊字符转义（防XSS，优化点：强化安全性）
function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// JSON响应特殊字符转义（优化点：避免前端解析异常）
function escapeJson(str) {
  if (!str) return str;
  return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n');
}

// 参数校验（优化点：细粒度校验，提前拦截无效请求）
function validateParams(params) {
  const { email } = params;
  // 邮箱格式校验
  const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailReg.test(email)) return new Error("邮箱格式无效，请输入正确的邮箱地址");
  // Token和ClientID长度校验（简单防错）
  if (params.refresh_token?.length < 50) return new Error("refresh_token格式无效");
  if (params.client_id?.length < 10) return new Error("client_id格式无效");
  return null;
}

// ===================== 核心业务函数 =====================
// 生成邮件HTML（优化点：中文文案优化）
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
          .email-text { white-space: pre-line; }
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

// 获取access_token（优化点：替换为超时请求，中文错误）
async function get_access_token(refresh_token, client_id) {
  try {
    const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token
      }).toString()
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP错误！状态码：${response.status}，响应：${errorText}`);
    }

    const responseText = await response.text();
    const data = JSON.parse(responseText);
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

// 检查Graph API权限（优化点：超时请求，中文日志）
async function graph_api(refresh_token, client_id) {
  try {
    const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'scope': 'https://graph.microsoft.com/.default'
      }).toString()
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Graph API请求失败：状态码${response.status}，响应：${errorText}`);
    }

    const responseText = await response.text();
    const data = JSON.parse(responseText);
    const hasMailPermission = data.scope?.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1;

    return {
      access_token: data.access_token,
      status: hasMailPermission
    };
  } catch (error) {
    console.error('Graph API权限检查失败：', error);
    return { access_token: '', status: false };
  }
}

// 获取邮件（优化点：超时请求，中文错误，空数据处理）
async function get_emails(access_token, mailbox, returnRaw = false) {
  if (!access_token) {
    throw new Error("access_token不存在");
  }

  try {
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
      throw new Error(`获取邮件失败：状态码${response.status}，响应：${errorText}`);
    }

    const responseData = await response.json();
    const emails = responseData.value || [];

    if (emails.length === 0) return null;

    const response_emails = emails.map(item => ({
      send: item['from']?.['emailAddress']?.['address'] || '未知发件人',
      subject: item['subject'] || '无主题',
      text: item['bodyPreview'] || '',
      html: item['body']?.['content'] || '',
      date: item['createdDateTime'] || new Date().toISOString(),
    }));

    return returnRaw ? response_emails[0] : response_emails;
  } catch (error) {
    console.error('获取邮件异常：', error);
    throw new Error(`邮件获取异常：${error.message}`);
  }
}

// ===================== 主入口函数（优化点：全流程中文提示+错误分类）=====================
module.exports = async (req, res) => {
  try {
    // 1. 限制请求方法
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({
        code: 405,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    // 2. 获取参数并验证密码
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

    // 5. 优先尝试Graph API
    console.log("【开始】检查Graph API权限");
    const graph_api_result = await graph_api(refresh_token, client_id);

    if (graph_api_result.status) {
      console.log("【成功】Graph API权限通过，使用Graph API获取邮件");
      // 中文文件夹映射
      const normalizedMailbox = CONFIG.MAILBOX_MAP[mailbox.toLowerCase()];
      if (!normalizedMailbox) {
        const supportedMailboxes = Object.keys(CONFIG.MAILBOX_MAP).filter(key => !/[a-z]/.test(key)).join('、');
        return res.status(400).json({
          code: 4003,
          error: `不支持的文件夹名称：${mailbox}，支持的中文文件夹：${supportedMailboxes}`
        });
      }
      mailbox = normalizedMailbox;

      // 获取邮件并处理响应
      const emailData = await get_emails(graph_api_result.access_token, mailbox, true);
      if (!emailData) {
        const mailboxCN = Object.keys(CONFIG.MAILBOX_MAP).find(key => CONFIG.MAILBOX_MAP[key] === mailbox);
        return res.status(200).json({
          code: 2001,
          message: `当前“${mailboxCN}”文件夹无邮件`,
          data: null
        });
      }

      if (response_type === 'html') {
        const htmlResponse = generateEmailHtml(emailData);
        return res.status(200).send(htmlResponse);
      } else {
        return res.status(200).json({
          code: 200,
          message: '邮件获取成功',
          data: [emailData]
        });
      }
    }

    // 6. 降级使用IMAP协议
    console.log("【降级】Graph API权限不足或失败，使用IMAP协议获取邮件");
    const access_token = await get_access_token(refresh_token, client_id);
    const authString = generateAuthString(email, access_token);
    const imap = new Imap({ ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString });

    // IMAP连接逻辑（优化点：async/await统一异步写法，中文错误）
    imap.once("ready", async () => {
      try {
        // 打开邮箱文件夹
        await new Promise((resolve, reject) => {
          imap.openBox(mailbox, true, (err, box) => err ? reject(err) : resolve(box));
        });

        // 搜索最新邮件
        const results = await new Promise((resolve, reject) => {
          imap.search(["ALL"], (err, results) => err ? reject(err) : resolve(results));
        });

        if (results.length === 0) {
          imap.end();
          return res.status(200).json({
            code: 2001,
            message: `当前“${mailbox}”文件夹无邮件`,
            data: null
          });
        }

        // 获取并解析最新邮件（优化点：async/await替换回调）
        const latestMail = results.slice(-1);
        const f = imap.fetch(latestMail, { bodies: "" });

        f.on("message", async (msg) => {
          try {
            const stream = await new Promise((resolve) => msg.on("body", resolve));
            const mail = await simpleParser(stream);

            const responseData = {
              send: escapeJson(mail.from?.text || '未知发件人'),
              subject: escapeJson(mail.subject || '无主题'),
              text: escapeJson(mail.text || ''),
              html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
              date: mail.date || new Date().toLocaleString()
            };

            if (response_type === 'html') {
              const htmlResponse = generateEmailHtml(responseData);
              res.status(200).send(htmlResponse);
            } else {
              res.status(200).json({
                code: 200,
                message: '邮件获取成功',
                data: responseData
              });
            }
          } catch (err) {
            console.error('解析邮件失败：', err);
            res.status(500).json({
              code: 5003,
              error: `解析邮件失败：${err.message}`
            });
          }
        });

        f.once("end", () => imap.end());
      } catch (err) {
        imap.end();
        console.error('IMAP获取邮件失败：', err);
        res.status(500).json({
          code: 5002,
          error: `IMAP操作失败：${err.message}`
        });
      }
    });

    // IMAP错误处理
    imap.once('error', (err) => {
      console.error('IMAP连接错误：', err);
      res.status(500).json({
        code: 5001,
        error: `IMAP连接失败：${err.message}`
      });
    });

    imap.connect();

  } catch (error) {
    // 统一错误分类响应
    let statusCode = 500;
    let errorCode = 5000;

    if (error.message.includes('HTTP错误！状态码：401')) {
      statusCode = 401;
      errorCode = 4011;
      error.message = '认证失效，请刷新refresh_token';
    } else if (error.message.includes('HTTP错误！状态码：403')) {
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
