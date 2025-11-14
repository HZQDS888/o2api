const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// ===================== 全局配置与工具函数（仅新增日期对比工具，其他不变）=====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  GRAPH_API_BASE_URL: 'https://graph.microsoft.com/v1.0/me/mailFolders',
  IMAP_CONFIG: {
    host: 'outlook.office365.com',
    port: 993,
    tls: true,
    tlsOptions: { rejectUnauthorized: false }, // 恢复原有配置
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
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email', 'mailbox'], // 必要参数
  TARGET_FOLDERS: ['inbox', 'junkemail'] // 新增：固定检查收件箱+垃圾箱
};

// 请求超时封装（不变）
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

// HTML特殊字符转义（防XSS，不变）
function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// JSON响应特殊字符转义（不变）
function escapeJson(str) {
  if (!str) return str;
  return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n');
}

// 新增：对比两封邮件，返回最新的一封（核心工具函数）
function getLatestEmail(email1, email2) {
  if (!email1) return email2;
  if (!email2) return email1;
  // 统一转换为时间戳对比（兼容所有日期格式）
  const time1 = new Date(email1.date).getTime() || 0;
  const time2 = new Date(email2.date).getTime() || 0;
  return time1 > time2 ? email1 : email2;
}

// 参数校验（不变，仅保留原有逻辑）
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

// ===================== 核心业务函数（仅修改取件逻辑，其他不变）=====================
// 生成邮件HTML（不变）
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

// 获取access_token（不变）
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

// 生成IMAP认证字符串（不变）
const generateAuthString = (user, accessToken) => {
  const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
  return Buffer.from(authString).toString('base64');
};

// 检查Graph API权限（不变）
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

// 原有单个文件夹取件（不变，供内部调用）
async function get_single_folder_email(access_token, mailbox) {
  try {
    const url = `${CONFIG.GRAPH_API_BASE_URL}/${mailbox}/messages?$top=1&$orderby=receivedDateTime desc`;
    const response = await fetchWithTimeout(url, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        "Authorization": `Bearer ${access_token}`
      },
    });

    if (!response.ok) return null;
    const responseData = await response.json();
    const email = responseData.value?.[0];
    if (!email) return null;

    // 统一格式化邮件数据
    return {
      send: email['from']?.['emailAddress']?.['address'] || '未知发件人',
      subject: email['subject'] || '无主题',
      text: email['bodyPreview'] || '',
      html: email['body']?.['content'] || '',
      date: email['createdDateTime'] || new Date().toISOString(),
    };
  } catch (error) {
    console.error(`获取${mailbox}邮件失败：`, error);
    return null;
  }
}

// 新增：Graph API双文件夹取最新邮件（核心修改）
async function get_dual_folder_latest_email_graph(access_token) {
  // 并行获取两个文件夹的邮件
  const [inboxEmail, junkEmail] = await Promise.all([
    get_single_folder_email(access_token, 'inbox'),
    get_single_folder_email(access_token, 'junkemail')
  ]);
  // 对比返回最新的一封
  return getLatestEmail(inboxEmail, junkEmail);
}

// 新增：IMAP双文件夹取最新邮件（核心修改）
async function get_dual_folder_latest_email_imap(imapConfig) {
  const imap = new Imap(imapConfig);
  let inboxEmail = null;
  let junkEmail = null;

  // Promise封装IMAP操作
  const fetchEmails = new Promise((resolve, reject) => {
    imap.once('ready', async () => {
      try {
        // 1. 获取收件箱最新邮件
        await new Promise((res, rej) => {
          imap.openBox('inbox', true, (err) => err ? rej(err) : res());
        });
        const inboxResults = await new Promise((res, rej) => {
          imap.search(["ALL"], (err, resArr) => err ? rej(err) : res(resArr));
        });
        if (inboxResults.length > 0) {
          const latestInbox = inboxResults.slice(-1);
          const f1 = imap.fetch(latestInbox, { bodies: "" });
          await new Promise((res) => {
            f1.on('message', async (msg) => {
              const stream = await new Promise((r) => msg.on("body", r));
              const mail = await simpleParser(stream);
              inboxEmail = {
                send: escapeJson(mail.from?.text || '未知发件人'),
                subject: escapeJson(mail.subject || '无主题'),
                text: escapeJson(mail.text || ''),
                html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
                date: mail.date || new Date().toISOString()
              };
              res();
            });
          });
        }

        // 2. 获取垃圾箱最新邮件
        await new Promise((res, rej) => {
          imap.openBox('junkemail', true, (err) => err ? rej(err) : res());
        });
        const junkResults = await new Promise((res, rej) => {
          imap.search(["ALL"], (err, resArr) => err ? rej(err) : res(resArr));
        });
        if (junkResults.length > 0) {
          const latestJunk = junkResults.slice(-1);
          const f2 = imap.fetch(latestJunk, { bodies: "" });
          await new Promise((res) => {
            f2.on('message', async (msg) => {
              const stream = await new Promise((r) => msg.on("body", r));
              const mail = await simpleParser(stream);
              junkEmail = {
                send: escapeJson(mail.from?.text || '未知发件人'),
                subject: escapeJson(mail.subject || '无主题'),
                text: escapeJson(mail.text || ''),
                html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
                date: mail.date || new Date().toISOString()
              };
              res();
            });
          });
        }

        imap.end();
        // 对比返回最新的一封
        resolve(getLatestEmail(inboxEmail, junkEmail));
      } catch (err) {
        imap.end();
        reject(err);
      }
    });

    imap.once('error', (err) => reject(err));
    imap.connect();
  });

  return fetchEmails;
}

// ===================== 主入口函数（仅修改取件逻辑调用，其他不变）=====================
module.exports = async (req, res) => {
  try {
    // 1. 限制请求方法（不变）
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({
        code: 405,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    // 2. 获取参数并验证密码（不变，恢复原有明文对比）
    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        code: 4010,
        error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
      });
    }

    // 3. 提取并校验必要参数（不变）
    const params = isGet ? req.query : req.body;
    let { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;
    const missingParams = CONFIG.REQUIRED_PARAMS.filter(key => !params[key]);

    if (missingParams.length > 0) {
      return res.status(400).json({
        code: 4001,
        error: `缺少必要参数：${missingParams.join('、')}`
      });
    }

    // 4. 校验参数格式（不变）
    const paramError = validateParams(params);
    if (paramError) {
      return res.status(400).json({
        code: 4002,
        error: paramError.message
      });
    }

    // 5. 优先尝试Graph API（修改：调用双文件夹取件）
    console.log("【开始】检查Graph API权限");
    const graph_api_result = await graph_api(refresh_token, client_id);

    if (graph_api_result.status) {
      console.log("【成功】Graph API权限通过，获取收件箱+垃圾箱最新邮件");
      // 直接调用双文件夹取件逻辑，忽略原mailbox参数
      const latestEmail = await get_dual_folder_latest_email_graph(graph_api_result.access_token);

      if (!latestEmail) {
        return res.status(200).json({
          code: 2001,
          message: "收件箱和垃圾箱均无邮件",
          data: null
        });
      }

      // 处理响应（不变）
      if (response_type === 'html') {
        const htmlResponse = generateEmailHtml(latestEmail);
        return res.status(200).send(htmlResponse);
      } else {
        return res.status(200).json({
          code: 200,
          message: '邮件获取成功',
          data: [latestEmail]
        });
      }
    }

    // 6. 降级使用IMAP协议（修改：调用双文件夹取件）
    console.log("【降级】Graph API权限不足，使用IMAP获取收件箱+垃圾箱最新邮件");
    const access_token = await get_access_token(refresh_token, client_id);
    const authString = generateAuthString(email, access_token);
    const imapConfig = { ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString };

    // 调用IMAP双文件夹取件逻辑
    const latestEmailImap = await get_dual_folder_latest_email_imap(imapConfig);

    if (!latestEmailImap) {
      return res.status(200).json({
        code: 2001,
        message: "收件箱和垃圾箱均无邮件",
        data: null
      });
    }

    // 处理响应（不变）
    if (response_type === 'html') {
      const htmlResponse = generateEmailHtml(latestEmailImap);
      return res.status(200).send(htmlResponse);
    } else {
      return res.status(200).json({
        code: 200,
        message: '邮件获取成功',
        data: [latestEmailImap]
      });
    }

  } catch (error) {
    // 统一错误分类响应（不变）
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
