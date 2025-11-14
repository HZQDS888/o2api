const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;
const bcrypt = require('bcrypt'); // 新增：密码哈希对比（需安装：npm i bcrypt）

// ===================== 全局配置与工具函数 =====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  GRAPH_API_BASE_URL: 'https://graph.microsoft.com/v1.0/me/mailFolders',
  IMAP_CONFIG: {
    host: 'outlook.office365.com',
    port: 993,
    tls: true,
    // 移除危险配置：生产环境必须校验TLS证书
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
  RESPONSE_TYPES: ['json', 'html'], // 支持的响应格式
  TARGET_FOLDERS: ['inbox', 'junkemail'] // 固定检查：收件箱+垃圾箱
};

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

// 对比两封邮件，返回最新的一封（核心工具函数）
function getLatestEmail(email1, email2) {
  if (!email1) return email2;
  if (!email2) return email1;

  // 统一转换为时间戳（兼容ISO字符串、Date对象、普通字符串）
  const time1 = new Date(email1.date).getTime() || 0;
  const time2 = new Date(email2.date).getTime() || 0;
  return time1 > time2 ? email1 : email2;
}

// 参数校验（增强版）
function validateParams(params) {
  const { email, response_type = 'json' } = params;

  // 邮箱格式校验
  const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailReg.test(email)) return new Error("邮箱格式无效，请输入正确的邮箱地址");

  // Token和ClientID长度校验
  if (params.refresh_token?.length < 50) return new Error("refresh_token格式无效");
  if (params.client_id?.length < 10) return new Error("client_id格式无效");

  // 响应格式校验
  if (!CONFIG.RESPONSE_TYPES.includes(response_type)) {
    return new Error(`响应格式仅支持${CONFIG.RESPONSE_TYPES.join('或')}`);
  }

  return null;
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

// 公共Token获取函数（复用逻辑）
async function fetchToken(refresh_token, client_id, scope = '') {
  const body = new URLSearchParams({
    'client_id': client_id,
    'grant_type': 'refresh_token',
    'refresh_token': refresh_token
  });
  if (scope) body.append('scope', scope);

  const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`HTTP错误！状态码：${response.status}，响应：${errorText}`);
  }

  const responseText = await response.text();
  return JSON.parse(responseText);
}

// 获取access_token
async function getAccessToken(refresh_token, client_id) {
  try {
    const data = await fetchToken(refresh_token, client_id);
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
    const data = await fetchToken(refresh_token, client_id, 'https://graph.microsoft.com/.default');
    const hasMailPermission = data.scope?.includes('https://graph.microsoft.com/Mail.ReadWrite') || false;
    return {
      access_token: data.access_token,
      status: hasMailPermission
    };
  } catch (error) {
    console.error('Graph API权限检查失败：', error);
    return { access_token: '', status: false };
  }
}

// Graph API：并行获取双文件夹最新邮件（核心修改）
async function getDualFolderLatestEmailGraph(access_token) {
  // 并行请求两个文件夹的最新邮件
  const folderPromises = CONFIG.TARGET_FOLDERS.map(async (folder) => {
    try {
      const url = `${CONFIG.GRAPH_API_BASE_URL}/${folder}/messages?$top=1&$orderby=receivedDateTime desc`;
      const response = await fetchWithTimeout(url, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          "Authorization": `Bearer ${access_token}`
        }
      });

      if (!response.ok) return null;
      const data = await response.json();
      const emailItem = data.value?.[0];
      if (!emailItem) return null;

      // 统一格式化邮件数据
      return {
        send: emailItem['from']?.['emailAddress']?.['address'] || '未知发件人',
        subject: emailItem['subject'] || '无主题',
        text: emailItem['bodyPreview'] || '',
        html: emailItem['body']?.['content'] || '',
        date: emailItem['createdDateTime'] || new Date().toISOString()
      };
    } catch (err) {
      console.error(`Graph API获取${folder}邮件失败：`, err);
      return null; // 单个文件夹失败不影响整体
    }
  });

  // 筛选有效邮件并返回最新
  const folderEmails = await Promise.all(folderPromises);
  const validEmails = folderEmails.filter(Boolean);
  if (validEmails.length === 0) return null;
  return validEmails.reduce((prev, curr) => getLatestEmail(prev, curr), validEmails[0]);
}

// IMAP：并行获取双文件夹最新邮件（核心修改）
async function getDualFolderLatestEmailImap(imapConfig) {
  const imap = new Imap(imapConfig);
  let folderEmails = [];

  // Promise封装IMAP操作，避免回调混乱
  const fetchEmails = new Promise((resolve, reject) => {
    imap.once('ready', async () => {
      try {
        // 遍历两个目标文件夹
        for (const folder of CONFIG.TARGET_FOLDERS) {
          try {
            // 打开文件夹（只读模式）
            await new Promise((res, rej) => {
              imap.openBox(folder, false, (err) => err ? rej(err) : res());
            });

            // 搜索该文件夹所有邮件，取最新一封
            const searchResults = await new Promise((res, rej) => {
              imap.search(["ALL"], (err, results) => err ? rej(err) : res(results));
            });

            if (searchResults.length === 0) continue;

            // 获取并解析最新邮件
            const latestMailId = searchResults.slice(-1);
            const fetchStream = imap.fetch(latestMailId, { bodies: "" });

            await new Promise((res) => {
              fetchStream.on('message', async (msg) => {
                const bodyStream = await new Promise((r) => msg.on("body", r));
                const mail = await simpleParser(bodyStream);
                // 格式化数据
                folderEmails.push({
                  send: escapeJson(mail.from?.text || '未知发件人'),
                  subject: escapeJson(mail.subject || '无主题'),
                  text: escapeJson(mail.text || ''),
                  html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
                  date: mail.date || new Date().toISOString()
                });
                res();
              });
            });
          } catch (err) {
            console.error(`IMAP获取${folder}邮件失败：`, err);
            continue; // 单个文件夹失败跳过
          }
        }
        imap.end();
        resolve(folderEmails);
      } catch (err) {
        imap.end();
        reject(err);
      }
    });

    // IMAP错误处理
    imap.once('error', (err) => reject(err));
    imap.once('end', () => console.log('IMAP连接已关闭'));

    imap.connect();
  });

  // 执行获取并筛选最新邮件
  folderEmails = await fetchEmails;
  const validEmails = folderEmails.filter(Boolean);
  if (validEmails.length === 0) return null;
  return validEmails.reduce((prev, curr) => getLatestEmail(prev, curr), validEmails[0]);
}

// ===================== 主入口函数 =====================
module.exports = async (req, res) => {
  let isResponded = false; // 防止重复响应的标志

  // 响应工具函数（统一处理）
  const sendResponse = (statusCode, data) => {
    if (isResponded) return;
    isResponded = true;
    res.status(statusCode).json(data);
  };

  const sendHtmlResponse = (statusCode, html) => {
    if (isResponded) return;
    isResponded = true;
    res.status(statusCode).send(html);
  };

  try {
    // 1. 限制请求方法
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return sendResponse(405, {
        code: 405,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    // 2. 获取请求参数
    const isGet = req.method === 'GET';
    const params = isGet ? req.query : req.body;
    const { password, refresh_token, client_id, email, mailbox, response_type = 'json' } = params;

    // 3. 密码校验（哈希对比，更安全）
    const expectedPassword = process.env.PASSWORD;
    if (expectedPassword) {
      const isPasswordValid = await bcrypt.compare(password, expectedPassword);
      if (!isPasswordValid) {
        return sendResponse(401, {
          code: 4010,
          error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
        });
      }
    }

    // 4. 检查必要参数
    const missingParams = CONFIG.REQUIRED_PARAMS.filter(key => !params[key]);
    if (missingParams.length > 0) {
      return sendResponse(400, {
        code: 4001,
        error: `缺少必要参数：${missingParams.join('、')}`
      });
    }

    // 5. 校验参数格式
    const paramError = validateParams(params);
    if (paramError) {
      return sendResponse(400, {
        code: 4002,
        error: paramError.message
      });
    }

    // 6. 优先使用Graph API
    console.log("【开始】检查Graph API权限");
    const graphResult = await checkGraphPermission(refresh_token, client_id);

    if (graphResult.status && graphResult.access_token) {
      console.log("【成功】Graph API权限通过，获取双文件夹最新邮件");
      const latestEmail = await getDualFolderLatestEmailGraph(graphResult.access_token);

      if (!latestEmail) {
        return sendResponse(200, {
          code: 2001,
          message: "收件箱和垃圾箱均无邮件",
          data: null
        });
      }

      // 处理响应格式
      if (response_type === 'html') {
        const html = generateEmailHtml(latestEmail);
        return sendHtmlResponse(200, html);
      } else {
        return sendResponse(200, {
          code: 200,
          message: '邮件获取成功',
          data: [latestEmail]
        });
      }
    }

    // 7. 降级使用IMAP协议
    console.log("【降级】Graph API权限不足，使用IMAP获取双文件夹最新邮件");
    const accessToken = await getAccessToken(refresh_token, client_id);
    const authString = generateAuthString(email, accessToken);
    const imapConfig = { ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString };

    // 获取IMAP渠道的最新邮件
    const latestEmailImap = await getDualFolderLatestEmailImap(imapConfig);

    if (!latestEmailImap) {
      return sendResponse(200, {
        code: 2001,
        message: "收件箱和垃圾箱均无邮件",
        data: null
      });
    }

    // 处理响应格式
    if (response_type === 'html') {
      const html = generateEmailHtml(latestEmailImap);
      return sendHtmlResponse(200, html);
    } else {
      return sendResponse(200, {
        code: 200,
        message: '邮件获取成功',
        data: [latestEmailImap]
      });
    }

  } catch (error) {
    // 统一错误分类处理
    let statusCode = 500;
    let errorCode = 5000;
    let errorMsg = error.message;

    if (errorMsg.includes('HTTP错误！状态码：401')) {
      statusCode = 401;
      errorCode = 4011;
      errorMsg = '认证失效，请刷新refresh_token';
    } else if (errorMsg.includes('HTTP错误！状态码：403')) {
      statusCode = 403;
      errorCode = 4031;
      errorMsg = '权限不足，需开启Mail.ReadWrite权限';
    } else if (errorMsg.includes('请求超时')) {
      statusCode = 504;
      errorCode = 5041;
    }

    return sendResponse(statusCode, {
      code: errorCode,
      error: `服务器错误：${errorMsg}`
    });
  }
};
