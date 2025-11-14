
const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// ===================== 全局配置与工具函数 =====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  GRAPH_API_BASE_URL: 'https://graph.microsoft.com/v1.0/me/mailFolders',
  IMAP_CONFIG: {
    host: 'outlook.office365.com',
    port: 993,
    tls: true,
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

// 获取access_token（通用token请求）
async function requestToken(refresh_token, client_id, scope = '') {
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
    return JSON.parse(responseText);
  } catch (error) {
    throw new Error(`获取token失败：${error.message}`);
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
    const tokenData = await requestToken(refresh_token, client_id, 'https://graph.microsoft.com/.default');
    const hasMailPermission = tokenData.scope?.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1;
    return {
      access_token: tokenData.access_token,
      status: hasMailPermission
    };
  } catch (error) {
    console.error('Graph API权限检查失败：', error);
    return { access_token: '', status: false };
  }
}

// 单文件夹拉取邮件（Graph API）
async function getSingleBoxEmails(access_token, mailbox, returnRaw = false) {
  if (!access_token) throw new Error("access_token不存在");

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
      throw new Error(`状态码${response.status}，响应：${errorText}`);
    }

    const responseData = await response.json();
    const emails = responseData.value || [];
    if (emails.length === 0) return null;

    const formattedEmails = emails.map(item => ({
      send: item['from']?.['emailAddress']?.['address'] || '未知发件人',
      subject: item['subject'] || '无主题',
      text: item['bodyPreview'] || '',
      html: item['body']?.['content'] || '',
      date: item['createdDateTime'] || new Date().toISOString(),
      mailbox: mailbox
    }));

    return returnRaw ? formattedEmails[0] : formattedEmails;
  } catch (error) {
    console.error(`拉取${mailbox}文件夹邮件异常：`, error);
    throw new Error(`邮件获取异常：${error.message}`);
  }
}

// ===================== 新增：收件箱与垃圾箱对比核心函数 =====================
// Graph API版：对比收件箱和垃圾箱，返回最新邮件
async function getLatestEmailBetweenInboxAndJunk(access_token) {
  const targetBoxes = ['inbox', 'junkemail']; // 固定对比文件夹
  const emailList = [];

  for (const box of targetBoxes) {
    try {
      const email = await getSingleBoxEmails(access_token, box, true);
      if (email) emailList.push(email);
    } catch (err) {
      console.error(`拉取${box}文件夹失败（不影响整体流程）：`, err);
    }
  }

  if (emailList.length === 0) return null;
  // 按时间戳降序排序，取最新一封
  return emailList.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime())[0];
}

// IMAP版：对比收件箱和垃圾箱，返回最新邮件
async function getLatestEmailByImap(imapConfig) {
  const imap = new Imap(imapConfig);
  const targetBoxes = ['inbox', 'junkemail'];
  const emailList = [];

  // IMAP连接封装
  const connectImap = () => new Promise((resolve, reject) => {
    imap.once('ready', resolve);
    imap.once('error', reject);
    imap.connect();
  });

  // 单文件夹拉取（IMAP）
  const fetchBoxLatestEmail = (boxName) => new Promise((resolve) => {
    imap.openBox(boxName, true, (err, box) => {
      if (err) {
        console.error(`IMAP打开${boxName}失败：`, err);
        return resolve(null);
      }

      imap.search(["ALL"], (err, results) => {
        if (err || results.length === 0) return resolve(null);

        const latestMail = results.slice(-1);
        const f = imap.fetch(latestMail, { bodies: "" });
        f.on("message", (msg) => {
          msg.on("body", async (stream) => {
            const mail = await simpleParser(stream);
            resolve({
              send: escapeJson(mail.from?.text || '未知发件人'),
              subject: escapeJson(mail.subject || '无主题'),
              text: escapeJson(mail.text || ''),
              html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
              date: mail.date || new Date().toLocaleString(),
              mailbox: boxName
            });
          });
        });
      });
    });
  });

  try {
    await connectImap();
    for (const box of targetBoxes) {
      const email = await fetchBoxLatestEmail(box);
      if (email) emailList.push(email);
    }
  } catch (err) {
    console.error('IMAP双文件夹拉取失败：', err);
  } finally {
    imap.end();
  }

  if (emailList.length === 0) return null;
  return emailList.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime())[0];
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

    // 2. 密码校验（补充环境变量校验）
    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (!expectedPassword) {
      return res.status(500).json({ code: 5004, error: '服务器未配置访问密码' });
    }
    if (password !== expectedPassword) {
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

    const paramError = validateParams(params);
    if (paramError) {
      return res.status(400).json({ code: 4002, error: paramError.message });
    }

    // 4. 中文文件夹标准化
    const requestBox = CONFIG.MAILBOX_MAP[mailbox.toLowerCase()];
    if (!requestBox) {
      const supportedMailboxes = Object.keys(CONFIG.MAILBOX_MAP).filter(key => !/[a-z]/.test(key)).join('、');
      return res.status(400).json({ code: 4003, error: `不支持的文件夹名称：${mailbox}，支持的中文文件夹：${supportedMailboxes}` });
    }

    // 5. 优先使用Graph API
    console.log("【开始】检查Graph API权限");
    const graphResult = await checkGraphPermission(refresh_token, client_id);

    if (graphResult.status) {
      console.log("【成功】Graph API权限通过，处理邮件拉取");
      let emailData;

      // 判断是否需要对比（仅收件箱/垃圾箱）
      const needCompare = ['inbox', 'junkemail'].includes(requestBox);
      if (needCompare) {
        emailData = await getLatestEmailBetweenInboxAndJunk(graphResult.access_token);
      } else {
        emailData = await getSingleBoxEmails(graphResult.access_token, requestBox, true);
      }

      // 响应处理
      if (!emailData) {
        const msg = needCompare ? '收件箱和垃圾箱均无邮件' : `当前“${mailbox}”文件夹无邮件`;
        return res.status(200).json({ code: 2001, message: msg, data: null });
      }

      const mailboxCN = Object.keys(CONFIG.MAILBOX_MAP).find(key => CONFIG.MAILBOX_MAP[key] === emailData.mailbox);
      if (response_type === 'html') {
        const htmlResponse = generateEmailHtml(emailData);
        return res.status(200).send(htmlResponse);
      } else {
        return res.status(200).json({
          code: 200,
          message: needCompare ? `最新邮件来自“${mailboxCN}”` : '邮件获取成功',
          data: [emailData] // 统一数组格式
        });
      }
    }

    // 6. 降级使用IMAP协议
    console.log("【降级】Graph API权限不足，使用IMAP协议");
    const tokenData = await requestToken(refresh_token, client_id);
    const authString = generateAuthString(email, tokenData.access_token);
    const imapConfig = { ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString };

    let emailData;
    if (['inbox', 'junkemail'].includes(requestBox)) {
      // 对比收件箱和垃圾箱
      emailData = await getLatestEmailByImap(imapConfig);
    } else {
      // 其他文件夹单文件夹拉取
      emailData = await new Promise((resolve, reject) => {
        const imap = new Imap(imapConfig);
        imap.once('ready', async () => {
          try {
            await new Promise((res, rej) => imap.openBox(requestBox, true, err => err ? rej(err) : res()));
            const results = await new Promise((res, rej) => imap.search(["ALL"], (err, res) => err ? rej(err) : res(res)));
            if (results.length === 0) {
              imap.end();
              return resolve(null);
            }

            const f = imap.fetch(results.slice(-1), { bodies: "" });
            f.on("message", async (msg) => {
              const stream = await new Promise(res => msg.on("body", res));
              const mail = await simpleParser(stream);
              resolve({
                send: escapeJson(mail.from?.text || '未知发件人'),
                subject: escapeJson(mail.subject || '无主题'),
                text: escapeJson(mail.text || ''),
                html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
                date: mail.date || new Date().toLocaleString(),
                mailbox: requestBox
              });
            });
            f.once("end", () => imap.end());
          } catch (err) {
            imap.end();
            reject(err);
          }
        });

        imap.once('error', (err) => reject(new Error(`IMAP连接失败：${err.message}`)));
        imap.connect();
      });
    }

    // IMAP响应处理
    if (!emailData) {
      const msg = ['inbox', 'junkemail'].includes(requestBox) ? '收件箱和垃圾箱均无邮件' : `当前“${mailbox}”文件夹无邮件`;
      return res.status(200).json({ code: 2001, message: msg, data: null });
    }

    const mailboxCN = Object.keys(CONFIG.MAILBOX_MAP).find(key => CONFIG.MAILBOX_MAP[key] === emailData.mailbox);
    if (response_type === 'html') {
      const htmlResponse = generateEmailHtml(emailData);
      res.status(200).send(htmlResponse);
    } else {
      res.status(200).json({
        code: 200,
        message: ['inbox', 'junkemail'].includes(requestBox) ? `最新邮件来自“${mailboxCN}”` : '邮件获取成功',
        data: [emailData]
      });
    }

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
