const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// ===================== 全局配置与工具函数（仅新增文件夹配置，其他不变）=====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  GRAPH_API_BASE_URL: 'https://graph.microsoft.com/v1.0/me/mailFolders',
  IMAP_CONFIG: {
    host: 'outlook.office365.com',
    port: 993,
    tls: true,
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
  REQUEST_TIMEOUT: 10000,
  SUPPORTED_METHODS: ['GET', 'POST'],
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email', 'mailbox'],
  // 修复：区分Graph API和IMAP的文件夹名称（兼容不同账户）
  TARGET_FOLDERS: {
    graph: ['inbox', 'junkemail'], // Graph API标准名称
    imap: ['INBOX', 'Junk'] // IMAP标准名称（Outlook垃圾箱默认叫Junk）
  }
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

// HTML特殊字符转义（不变）
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

// 对比两封邮件，返回最新的一封（不变）
function getLatestEmail(email1, email2) {
  if (!email1) return email2;
  if (!email2) return email1;
  const time1 = new Date(email1.date).getTime() || 0;
  const time2 = new Date(email2.date).getTime() || 0;
  return time1 > time2 ? email1 : email2;
}

// 参数校验（不变）
function validateParams(params) {
  const { email } = params;
  const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailReg.test(email)) return new Error("邮箱格式无效，请输入正确的邮箱地址");
  if (params.refresh_token?.length < 50) return new Error("refresh_token格式无效");
  if (params.client_id?.length < 10) return new Error("client_id格式无效");
  return null;
}

// ===================== 核心业务函数（修复文件夹兼容性，其他不变）=====================
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

// 原有单个文件夹取件（修复：文件夹不存在时返回null，不中断流程）
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

    // 修复：如果文件夹不存在（404），返回null
    if (!response.ok) {
      console.warn(`文件夹${mailbox}访问失败，状态码：${response.status}`);
      return null;
    }

    const responseData = await response.json();
    const email = responseData.value?.[0];
    if (!email) return null;

    return {
      send: email['from']?.['emailAddress']?.['address'] || '未知发件人',
      subject: email['subject'] || '无主题',
      text: email['bodyPreview'] || '',
      html: email['body']?.['content'] || '',
      date: email['createdDateTime'] || new Date().toISOString(),
    };
  } catch (error) {
    console.error(`获取${mailbox}邮件失败：`, error);
    return null; // 修复：任何错误都返回null，避免影响另一文件夹
  }
}

// Graph API双文件夹取最新邮件（不变，依赖上面的修复）
async function get_dual_folder_latest_email_graph(access_token) {
  const [inboxEmail, junkEmail] = await Promise.all([
    get_single_folder_email(access_token, CONFIG.TARGET_FOLDERS.graph[0]),
    get_single_folder_email(access_token, CONFIG.TARGET_FOLDERS.graph[1])
  ]);
  return getLatestEmail(inboxEmail, junkEmail);
}

// IMAP双文件夹取最新邮件（修复：用IMAP标准文件夹名称Junk，失败时跳过）
async function get_dual_folder_latest_email_imap(imapConfig) {
  const imap = new Imap(imapConfig);
  let inboxEmail = null;
  let junkEmail = null;

  const fetchEmails = new Promise((resolve, reject) => {
    imap.once('ready', async () => {
      try {
        // 1. 获取收件箱最新邮件（IMAP标准名称INBOX）
        try {
          await new Promise((res, rej) => {
            imap.openBox(CONFIG.TARGET_FOLDERS.imap[0], true, (err) => err ? rej(err) : res());
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
        } catch (err) {
          console.error('IMAP获取收件箱邮件失败：', err);
        }

        // 2. 获取垃圾箱最新邮件（修复：用IMAP标准名称Junk，失败时跳过）
        try {
          await new Promise((res, rej) => {
            imap.openBox(CONFIG.TARGET_FOLDERS.imap[1], true, (err) => err ? rej(err) : res());
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
        } catch (err) {
          console.error('IMAP获取垃圾箱邮件失败：', err);
          // 垃圾箱获取失败，不中断，继续返回收件箱邮件
        }

        imap.end();
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

// ===================== 主入口函数（不变）=====================
module.exports = async (req, res) => {
  try {
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({
        code: 405,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        code: 4010,
        error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
      });
    }

    const params = isGet ? req.query : req.body;
    let { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;
    const missingParams = CONFIG.REQUIRED_PARAMS.filter(key => !params[key]);

    if (missingParams.length > 0) {
      return res.status(400).json({
        code: 4001,
        error: `缺少必要参数：${missingParams.join('、')}`
      });
    }

    const paramError = validateParams(params);
    if (paramError) {
      return res.status(400).json({
        code: 4002,
        error: paramError.message
      });
    }

    console.log("【开始】检查Graph API权限");
    const graph_api_result = await graph_api(refresh_token, client_id);

    if (graph_api_result.status) {
      console.log("【成功】Graph API权限通过，获取收件箱+垃圾箱最新邮件");
      const latestEmail = await get_dual_folder_latest_email_graph(graph_api_result.access_token);

      if (!latestEmail) {
        return res.status(200).json({
          code: 2001,
          message: "收件箱和垃圾箱均无邮件",
          data: null
        });
      }

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

    console.log("【降级】Graph API权限不足，使用IMAP获取收件箱+垃圾箱最新邮件");
    const access_token = await get_access_token(refresh_token, client_id);
    const authString = generateAuthString(email, access_token);
    const imapConfig = { ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString };

    const latestEmailImap = await get_dual_folder_latest_email_imap(imapConfig);

    if (!latestEmailImap) {
      return res.status(200).json({
        code: 2001,
        message: "收件箱和垃圾箱均无邮件",
        data: null
      });
    }

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
