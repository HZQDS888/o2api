const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;
const DOMPurify = require('dompurify'); // 新增：强化HTML过滤

// ===================== 全局配置与工具函数 =====================
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
  // 固定查询的两个文件夹（收件箱+垃圾箱）
  TARGET_FOLDERS: {
    graph: ['inbox', 'junkemail'], // Graph API文件夹名
    imap: ['INBOX', 'JUNK']        // IMAP文件夹名
  },
  REQUEST_TIMEOUT: 10000,
  SUPPORTED_METHODS: ['GET', 'POST'],
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email'] // 移除mailbox的必填要求
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

// HTML特殊字符转义+危险标签过滤
function escapeHtml(str) {
  if (!str) return '';
  return DOMPurify.sanitize(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// 参数校验（移除mailbox校验，保留原有逻辑）
function validateParams(params) {
  const { email } = params;
  const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailReg.test(email)) return new Error("邮箱格式无效，请输入正确的邮箱地址");
  if (params.refresh_token?.length < 50) return new Error("refresh_token格式无效");
  if (params.client_id?.length < 10) return new Error("client_id格式无效");
  return null;
}

// ===================== 核心业务函数（重点修改） =====================
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

// 获取access_token
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

// 检查Graph API权限
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

// Graph API：查询收件箱+垃圾箱，返回最新1封
async function get_latest_email_from_both_folders(access_token) {
  if (!access_token) {
    throw new Error("access_token不存在");
  }

  try {
    // 并行查询两个文件夹
    const folderPromises = CONFIG.TARGET_FOLDERS.graph.map(folder => {
      const url = `${CONFIG.GRAPH_API_BASE_URL}/${folder}/messages?$top=5&$orderby=receivedDateTime desc`;
      return fetchWithTimeout(url, {
        method: 'GET',
        headers: { 'Content-Type': 'application/json', "Authorization": `Bearer ${access_token}` }
      });
    });

    const responses = await Promise.allSettled(folderPromises);
    const allEmails = [];

    // 解析所有成功返回的邮件
    for (const res of responses) {
      if (res.status === 'fulfilled' && res.value.ok) {
        const data = await res.value.json();
        allEmails.push(...(data.value || []));
      }
    }

    if (allEmails.length === 0) return null;

    // 按时间排序，取最新1封
    allEmails.sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
    const latestEmail = allEmails[0];

    return {
      send: latestEmail.from?.emailAddress?.address || '未知发件人',
      subject: latestEmail.subject || '无主题',
      text: latestEmail.bodyPreview || '',
      html: latestEmail.body?.content || '',
      date: latestEmail.receivedDateTime || new Date().toISOString(),
      folder: CONFIG.TARGET_FOLDERS.graph.find(f => f === latestEmail.parentFolderId?.split('/').pop()) || '未知文件夹'
    };
  } catch (error) {
    console.error('多文件夹邮件获取失败：', error);
    throw new Error(`邮件获取异常：${error.message}`);
  }
}

// IMAP：查询收件箱+垃圾箱，返回最新1封
async function get_latest_email_from_both_folders_imap(access_token, email) {
  const authString = generateAuthString(email, access_token);
  const imap = new Imap({ ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString });

  return new Promise((resolve, reject) => {
    imap.once("ready", async () => {
      try {
        const allEmails = [];

        // 逐个查询文件夹（IMAP不支持并行打开多个文件夹）
        for (const folder of CONFIG.TARGET_FOLDERS.imap) {
          await new Promise((openResolve, openReject) => {
            imap.openBox(folder, true, async (err, box) => {
              if (err) {
                console.error(`打开文件夹${folder}失败：`, err);
                return openResolve();
              }

              // 搜索该文件夹所有邮件，按时间倒序
              const results = await new Promise((searchResolve, searchReject) => {
                imap.search(['ALL'], (err, res) => err ? searchReject(err) : searchResolve(res));
              });

              if (results.length === 0) {
                return openResolve();
              }

              // 取该文件夹最新1封
              const latestMailId = results.slice(-1)[0];
              const f = imap.fetch(latestMailId, { bodies: "" });

              f.once("message", async (msg) => {
                const stream = await new Promise(resolve => msg.on("body", resolve));
                const mail = await simpleParser(stream);

                allEmails.push({
                  send: mail.from?.text || '未知发件人',
                  subject: mail.subject || '无主题',
                  text: mail.text || '',
                  html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
                  date: mail.date || new Date(),
                  folder: folder
                });
              });

              f.once("end", () => openResolve());
              f.on("error", (err) => {
                console.error(`解析文件夹${folder}邮件失败：`, err);
                openResolve();
              });
            });
          });
        }

        imap.end();

        if (allEmails.length === 0) {
          resolve(null);
          return;
        }

        // 按时间排序，取最新1封
        allEmails.sort((a, b) => new Date(b.date) - new Date(a.date));
        resolve(allEmails[0]);
      } catch (err) {
        imap.end();
        reject(err);
      }
    });

    imap.once('error', (err) => {
      imap.end();
      reject(new Error(`IMAP连接失败：${err.message}`));
    });

    imap.connect();
  });
}

// ===================== 主入口函数（完全保持原有请求方式） =====================
module.exports = async (req, res) => {
  try {
    // 1. 限制请求方法
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({
        code: 405,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    // 2. 密码验证（保持原有逻辑）
    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        code: 4010,
        error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
      });
    }

    // 3. 提取参数（忽略mailbox，内部固定查询两个文件夹）
    const params = isGet ? req.query : req.body;
    let { refresh_token, client_id, email, response_type = 'json' } = params;
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
      console.log("【成功】Graph API权限通过，查询收件箱+垃圾箱最新邮件");
      const latestEmail = await get_latest_email_from_both_folders(graph_api_result.access_token);

      if (!latestEmail) {
        return res.status(200).json({
          code: 2001,
          message: '收件箱和垃圾箱中均无邮件',
          data: null
        });
      }

      if (response_type === 'html') {
        const htmlResponse = generateEmailHtml(latestEmail);
        return res.status(200).send(htmlResponse);
      } else {
        return res.status(200).json({
          code: 200,
          message: `邮件获取成功（来自${latestEmail.folder === 'inbox' ? '收件箱' : '垃圾箱'}）`,
          data: [latestEmail]
        });
      }
    }

    // 6. 降级使用IMAP协议
    console.log("【降级】Graph API权限不足，使用IMAP查询收件箱+垃圾箱最新邮件");
    const access_token = await get_access_token(refresh_token, client_id);
    const latestEmail = await get_latest_email_from_both_folders_imap(access_token, email);

    if (!latestEmail) {
      return res.status(200).json({
        code: 2001,
        message: '收件箱和垃圾箱中均无邮件',
        data: null
      });
    }

    if (response_type === 'html') {
      const htmlResponse = generateEmailHtml(latestEmail);
      return res.status(200).send(htmlResponse);
    } else {
      return res.status(200).json({
        code: 200,
        message: `邮件获取成功（来自${latestEmail.folder === 'INBOX' ? '收件箱' : '垃圾箱'}）`,
        data: [latestEmail]
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
