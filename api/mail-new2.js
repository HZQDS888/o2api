
const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;
const crypto = require('crypto'); // 新增：用于密码哈希对比

// ===================== 全局配置（整合所有优化点）=====================
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
  // 中文→英文文件夹映射（精简冗余）
  MAILBOX_MAP: {
    '收件箱': 'inbox',
    '已发送': 'sentitems',
    '草稿': 'draft',
    '删除邮件': 'deleteditems',
    '垃圾邮件': 'junkemail'
  },
  // 英文→中文文件夹映射（用于返回结果标注）
  MAILBOX_CN_MAP: {
    'inbox': '收件箱',
    'junkemail': '垃圾箱',
    'sentitems': '已发送',
    'draft': '草稿',
    'deleteditems': '删除邮件'
  },
  // 强制扫描的文件夹组（收件箱+垃圾箱，取全局最新）
  MANDATORY_SEARCH_FOLDERS: ['inbox', 'junkemail'],
  REQUEST_TIMEOUT: 10000,
  SUPPORTED_METHODS: ['GET', 'POST'],
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email', 'mailbox'],
  // 安全配置：参数长度限制
  PARAM_LIMITS: {
    refresh_token: 500,
    client_id: 100,
    email: 100
  }
};

// ===================== 工具函数（整合优化）=====================
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

// HTML特殊字符转义+危险标签过滤（防XSS）
function escapeHtml(str) {
  if (!str) return '';
  // 先过滤危险标签
  let safeStr = str
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<iframe[\s\S]*?<\/iframe>/gi, '')
    .replace(/on\w+="[^"]*"/gi, '')
    .replace(/on\w+='[^']*'/gi, '');
  // 转义特殊字符
  return safeStr
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

// 参数校验（增强版：格式+长度+必填）
function validateParams(params) {
  const { email, refresh_token, client_id } = params;
  // 邮箱格式校验
  const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailReg.test(email)) return new Error("邮箱格式无效，请输入正确的邮箱地址");
  // 长度校验
  if (refresh_token?.length > CONFIG.PARAM_LIMITS.refresh_token) return new Error(`refresh_token长度不能超过${CONFIG.PARAM_LIMITS.refresh_token}字符`);
  if (client_id?.length > CONFIG.PARAM_LIMITS.client_id) return new Error(`client_id长度不能超过${CONFIG.PARAM_LIMITS.client_id}字符`);
  if (email.length > CONFIG.PARAM_LIMITS.email) return new Error(`邮箱长度不能超过${CONFIG.PARAM_LIMITS.email}字符`);
  // 非空校验（补充）
  if (!refresh_token) return new Error("refresh_token不能为空");
  if (!client_id) return new Error("client_id不能为空");
  return null;
}

// 密码哈希对比（增强安全性）
function verifyPassword(inputPassword, expectedPassword) {
  if (!expectedPassword) return true; // 未设置密码时跳过校验
  const inputHash = crypto.createHash('md5').update(inputPassword).digest('hex');
  const expectedHash = crypto.createHash('md5').update(expectedPassword).digest('hex');
  return inputHash === expectedHash;
}

// ===================== 核心业务函数（整合全局最新邮件逻辑）=====================
// 生成邮件HTML
function generateEmailHtml(emailData) {
  const { send, subject, text, html: emailHtml, date, fromFolder } = emailData;
  const escapedText = escapeHtml(text || '');
  const escapedHtml = emailHtml || `<p>${escapedText.replace(/\n/g, '<br>')}</p>`;
  const folderText = fromFolder ? `（来自${fromFolder}）` : '';

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
          .folder-tag { display: inline-block; padding: 2px 8px; background: #e8f4f8; color: #4299e1; border-radius: 4px; font-size: 0.8em; margin-left: 8px; }
        </style>
      </head>
      <body>
        <div class="email-container">
          <div class="email-header">
            <h1 class="email-title">${escapeHtml(subject || '无主题')}${fromFolder ? `<span class="folder-tag">${fromFolder}</span>` : ''}</h1>
            <div class="email-meta">
              <span><strong>发件人：</strong>${escapeHtml(send || '未知发件人')}</span>
              <span><strong>发送日期：</strong>${date ? new Date(date).toLocaleString() : '未知日期'}</span>
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

// 获取access_token（合并重复请求逻辑）
async function getAccessToken(refresh_token, client_id, scope = '') {
  try {
    const bodyParams = {
      'client_id': client_id,
      'grant_type': 'refresh_token',
      'refresh_token': refresh_token
    };
    if (scope) bodyParams.scope = scope;

    const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams(bodyParams).toString()
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP错误！状态码：${response.status}，响应：${errorText}`);
    }

    const responseText = await response.text();
    // JSON解析容错
    let data;
    try {
      data = JSON.parse(responseText);
    } catch (parseErr) {
      throw new Error(`Token响应解析失败：${parseErr.message}`);
    }
    return data;
  } catch (error) {
    throw new Error(`获取access_token失败：${error.message}`);
  }
}

// 检查Graph API权限
async function checkGraphPermission(refresh_token, client_id) {
  try {
    const data = await getAccessToken(refresh_token, client_id, 'https://graph.microsoft.com/.default');
    const hasMailPermission = data.scope?.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1;
    return {
      access_token: data.access_token,
      status: hasMailPermission,
      expires_in: data.expires_in || 3600 // Token有效期（默认1小时）
    };
  } catch (error) {
    console.error('Graph API权限检查失败：', error);
    return { access_token: '', status: false, expires_in: 0 };
  }
}

// 单个文件夹获取最新邮件（Graph API）
async function getLatestEmailFromFolder(access_token, folder) {
  try {
    const url = `${CONFIG.GRAPH_API_BASE_URL}/${folder}/messages?$top=1&$orderby=receivedDateTime desc&$select=from,subject,bodyPreview,body,createdDateTime`;
    const response = await fetchWithTimeout(url, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        "Authorization": `Bearer ${access_token}`
      },
    });

    if (!response.ok) return null;

    const responseData = await response.json();
    const emails = responseData.value || [];
    if (emails.length === 0) return null;

    const email = emails[0];
    return {
      send: email['from']?.['emailAddress']?.['address'] || '未知发件人',
      subject: email['subject'] || '无主题',
      text: email['bodyPreview'] || '',
      html: email['body']?.['content'] || '',
      date: email['createdDateTime'] || new Date().toISOString(),
      fromFolder: CONFIG.MAILBOX_CN_MAP[folder] || folder,
      timestamp: new Date(email['createdDateTime']).getTime()
    };
  } catch (error) {
    console.error(`Graph API查询文件夹${folder}失败：`, error);
    return null;
  }
}

// 全局最新邮件查询（Graph API：扫描收件箱+垃圾箱）
async function getGraphGlobalLatestEmail(access_token) {
  let allEmails = [];

  for (const folder of CONFIG.MANDATORY_SEARCH_FOLDERS) {
    const email = await getLatestEmailFromFolder(access_token, folder);
    if (email) allEmails.push(email);
  }

  if (allEmails.length === 0) return null;
  // 按时间戳排序取最新
  allEmails.sort((a, b) => b.timestamp - a.timestamp);
  delete allEmails[0].timestamp;
  return allEmails[0];
}

// 单个文件夹获取最新邮件（IMAP）
async function getLatestEmailFromImapFolder(imap, folder) {
  try {
    // 打开文件夹
    const box = await new Promise((resolve, reject) => {
      imap.openBox(folder, true, (err, box) => err ? reject(err) : resolve(box));
    });

    // 搜索最新邮件
    const results = await new Promise((resolve, reject) => {
      imap.search(["ALL"], (err, results) => err ? reject(err) : resolve(results));
    });

    if (results.length === 0) return null;

    // 解析最新邮件
    const latestMailId = results.slice(-1);
    const f = imap.fetch(latestMailId, { bodies: "" });

    return new Promise((resolve) => {
      f.on("message", async (msg) => {
        const stream = await new Promise((resolve) => msg.on("body", resolve));
        const mail = await simpleParser(stream);
        resolve({
          send: mail.from?.text || '未知发件人',
          subject: mail.subject || '无主题',
          text: mail.text || '',
          html: mail.html || '',
          date: mail.date || new Date(),
          fromFolder: CONFIG.MAILBOX_CN_MAP[folder] || folder,
          timestamp: new Date(mail.date).getTime() || 0
        });
      });
      f.once("end", () => resolve(null));
    });
  } catch (error) {
    console.error(`IMAP查询文件夹${folder}失败：`, error);
    return null;
  }
}

// 全局最新邮件查询（IMAP：扫描收件箱+垃圾箱）
async function getImapGlobalLatestEmail(imap) {
  let allEmails = [];

  for (const folder of CONFIG.MANDATORY_SEARCH_FOLDERS) {
    const email = await getLatestEmailFromImapFolder(imap, folder);
    if (email) allEmails.push(email);
  }

  if (allEmails.length === 0) return null;
  allEmails.sort((a, b) => b.timestamp - a.timestamp);
  delete allEmails[0].timestamp;
  return allEmails[0];
}

// 生成IMAP认证字符串
const generateAuthString = (user, accessToken) => {
  const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
  return Buffer.from(authString).toString('base64');
};

// ===================== 主入口函数（完整流程）=====================
module.exports = async (req, res) => {
  // 全局超时控制（防止请求挂起）
  const globalTimeout = setTimeout(() => {
    res.status(504).json({
      code: 5041,
      error: '请求超时：全局处理超过15秒'
    });
  }, 15000);

  try {
    // 1. 限制请求方法
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      clearTimeout(globalTimeout);
      return res.status(405).json({
        code: 405,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    // 2. 提取参数
    const isGet = req.method === 'GET';
    const params = isGet ? req.query : req.body;
    const { password, refresh_token, client_id, email, mailbox, response_type = 'json' } = params;

    // 3. 密码校验
    const expectedPassword = process.env.PASSWORD;
    if (!verifyPassword(password, expectedPassword)) {
      clearTimeout(globalTimeout);
      return res.status(401).json({
        code: 4010,
        error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
      });
    }

    // 4. 必要参数校验
    const missingParams = CONFIG.REQUIRED_PARAMS.filter(key => !params[key]);
    if (missingParams.length > 0) {
      clearTimeout(globalTimeout);
      return res.status(400).json({
        code: 4001,
        error: `缺少必要参数：${missingParams.join('、')}`
      });
    }

    // 5. 参数格式校验
    const paramError = validateParams(params);
    if (paramError) {
      clearTimeout(globalTimeout);
      return res.status(400).json({
        code: 4002,
        error: paramError.message
      });
    }

    // 6. 文件夹名称标准化
    const normalizedMailbox = CONFIG.MAILBOX_MAP[mailbox.toLowerCase()] || mailbox.toLowerCase();
    const supportedMailboxes = Object.keys(CONFIG.MAILBOX_MAP).join('、');
    if (!CONFIG.MAILBOX_CN_MAP[normalizedMailbox] && !CONFIG.MANDATORY_SEARCH_FOLDERS.includes(normalizedMailbox)) {
      clearTimeout(globalTimeout);
      return res.status(400).json({
        code: 4003,
        error: `不支持的文件夹名称：${mailbox}，支持的中文文件夹：${supportedMailboxes}`
      });
    }

    // 7. 优先使用Graph API
    console.log("【开始】检查Graph API权限");
    const graphResult = await checkGraphPermission(refresh_token, client_id);

    if (graphResult.status && graphResult.access_token) {
      console.log("【成功】Graph API权限通过");
      let emailData = null;
      const isMandatoryFolder = CONFIG.MANDATORY_SEARCH_FOLDERS.includes(normalizedMailbox);

      if (isMandatoryFolder) {
        // 扫描收件箱+垃圾箱，取全局最新
        emailData = await getGraphGlobalLatestEmail(graphResult.access_token);
      } else {
        // 仅查询指定文件夹
        const folderEmail = await getLatestEmailFromFolder(graphResult.access_token, normalizedMailbox);
        if (folderEmail) delete folderEmail.timestamp;
        emailData = folderEmail;
      }

      // 处理响应
      if (!emailData) {
        const folderDesc = isMandatoryFolder ? '收件箱和垃圾箱' : CONFIG.MAILBOX_CN_MAP[normalizedMailbox];
        clearTimeout(globalTimeout);
        return res.status(200).json({
          code: 2001,
          message: `当前${folderDesc}均无邮件`,
          data: null
        });
      }

      clearTimeout(globalTimeout);
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

    // 8. 降级使用IMAP协议
    console.log("【降级】Graph API权限不足或失败，使用IMAP协议");
    const tokenData = await getAccessToken(refresh_token, client_id);
    const accessToken = tokenData.access_token;
    if (!accessToken) {
      clearTimeout(globalTimeout);
      return res.status(401).json({
        code: 4011,
        error: '认证失效，请刷新refresh_token'
      });
    }

    const authString = generateAuthString(email, accessToken);
    const imap = new Imap({ ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString });
    let imapClosed = false;

    // IMAP连接就绪
    imap.once("ready", async () => {
      try {
        let emailData = null;
        const isMandatoryFolder = CONFIG.MANDATORY_SEARCH_FOLDERS.includes(normalizedMailbox);

        if (isMandatoryFolder) {
          // 扫描收件箱+垃圾箱，取全局最新
          emailData = await getImapGlobalLatestEmail(imap);
        } else {
          // 仅查询指定文件夹
          const folderEmail = await getLatestEmailFromImapFolder(imap, normalizedMailbox);
          if (folderEmail) delete folderEmail.timestamp;
          emailData = folderEmail;
        }

        // 关闭IMAP连接
        imapClosed = true;
        imap.destroy();
        clearTimeout(globalTimeout);

        if (!emailData) {
          const folderDesc = isMandatoryFolder ? '收件箱和垃圾箱' : CONFIG.MAILBOX_CN_MAP[normalizedMailbox];
          return res.status(200).json({
            code: 2001,
            message: `当前${folderDesc}均无邮件`,
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
            data: emailData
          });
        }
      } catch (err) {
        imapClosed = true;
        imap.destroy();
        clearTimeout(globalTimeout);
        console.error('IMAP操作失败：', err);
        return res.status(500).json({
          code: 5002,
          error: `IMAP操作失败：${err.message}`
        });
      }
    });

    // IMAP错误处理
    imap.once('error', (err) => {
      if (!imapClosed) {
        imap.destroy();
        clearTimeout(globalTimeout);
      }
      console.error('IMAP连接错误：', err);
      return res.status(500).json({
        code: 5001,
        error: `IMAP连接失败：${err.message}`
      });
    });

    // IMAP关闭事件
    imap.once('end', () => {
      console.log('IMAP连接已关闭');
    });

    imap.connect();

  } catch (error) {
    clearTimeout(globalTimeout);
    // 错误分类处理
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

    return res.status(statusCode).json({
      code: errorCode,
      error: `服务器错误：${errorMsg}`
    });
  }
};
