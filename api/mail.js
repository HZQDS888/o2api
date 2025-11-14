const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// ===================== 全局配置与工具函数 =====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  GRAPH_API_BASE_URL: 'https://graph.microsoft.com/v1.0/me',
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
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email']
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

// HTML特殊字符转义
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

// ===================== 核心业务函数 =====================
// 生成邮件HTML
function generateEmailHtml(emailData) {
  const { send, subject, text, html: emailHtml, date, folder } = emailData;
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
              <span><strong>所在文件夹：</strong>${escapeHtml(folder || '未知文件夹')}</span>
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

    const data = await response.json();
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

    const data = await response.json();
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

// 获取全局最新邮件（Graph API）
async function getGlobalLatestEmail(access_token) {
  if (!access_token) {
    throw new Error("access_token不存在");
  }

  try {
    const url = `${CONFIG.GRAPH_API_BASE_URL}/messages?$top=1&$orderby=receivedDateTime desc&$select=from,subject,bodyPreview,body,createdDateTime,parentFolderId`;
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

    const latestEmail = emails[0];
    const folderId = latestEmail.parentFolderId;
    let folderName = '未知文件夹';
    try {
      const folderRes = await fetchWithTimeout(`${CONFIG.GRAPH_API_BASE_URL}/mailFolders/${folderId}`, {
        headers: { "Authorization": `Bearer ${access_token}` }
      });
      if (folderRes.ok) {
        const folderData = await folderRes.json();
        const folderAlias = folderData.displayName?.toLowerCase() || '';
        folderName = Object.keys(CONFIG.MAILBOX_MAP).find(key => !/[a-z]/.test(key) && CONFIG.MAILBOX_MAP[key] === folderAlias) || folderData.displayName || folderName;
      }
    } catch (e) {
      console.warn('获取文件夹名称失败：', e);
    }

    return {
      send: latestEmail['from']?.['emailAddress']?.['address'] || '未知发件人',
      subject: latestEmail['subject'] || '无主题',
      text: latestEmail['bodyPreview'] || '',
      html: latestEmail['body']?.['content'] || '',
      date: latestEmail['createdDateTime'] || new Date().toISOString(),
      folder: folderName
    };
  } catch (error) {
    console.error('获取全局最新邮件异常：', error);
    throw new Error(`邮件获取异常：${error.message}`);
  }
}

// IMAP：获取单个文件夹的最新邮件
async function getFolderLatestEmail(imap, folderPath) {
  return new Promise((resolve) => { // 直接resolve null，避免reject中断整个流程
    imap.openBox(folderPath, true, (err, box) => {
      if (err) {
        console.warn(`无法打开文件夹 ${folderPath}:`, err.message);
        return resolve(null);
      }

      imap.search([['UID', '*']], (err, results) => {
        if (err || results.length === 0) {
          return resolve(null);
        }

        const latestUid = results[results.length - 1];
        const f = imap.fetch(latestUid, { bodies: "" });

        f.on("message", async (msg) => {
          try {
            const stream = await new Promise(resolve => msg.on("body", resolve));
            const mail = await simpleParser(stream);
            
            // 查找文件夹的中文名称
            let folderCN = '未知文件夹';
            const folderName = folderPath.split('/').pop().toLowerCase();
            for (const [cnName, enName] of Object.entries(CONFIG.MAILBOX_MAP)) {
              if (enName.toLowerCase() === folderName) {
                folderCN = cnName;
                break;
              }
            }

            resolve({
              send: mail.from?.text || '未知发件人',
              subject: mail.subject || '无主题',
              text: mail.text || '',
              html: mail.html || '',
              date: new Date(mail.date || Date.now()).toISOString(),
              folder: folderCN
            });
          } catch (e) {
            console.error(`解析文件夹 ${folderPath} 中的邮件失败:`, e);
            resolve(null);
          }
        });

        f.on("error", (err) => {
          console.error(`获取文件夹 ${folderPath} 中的邮件失败:`, err);
          resolve(null);
        });
      });
    });
  });
}

// 递归遍历getBoxes返回的嵌套对象，收集所有文件夹路径
function collectAllFolderPaths(boxes, parentPath = '') {
  let paths = [];
  for (const [name, box] of Object.entries(boxes)) {
    const fullPath = parentPath ? `${parentPath}/${name}` : name;
    paths.push(fullPath);
    if (box.children) {
      paths = paths.concat(collectAllFolderPaths(box.children, fullPath));
    }
  }
  return paths;
}

// IMAP：获取全局最新邮件（遍历所有有效文件夹）
async function getImapGlobalLatestEmail(imap) {
  return new Promise((resolve, reject) => {
    // 使用 getBoxes() 获取所有文件夹
    imap.getBoxes((err, boxes) => {
      if (err) {
        return reject(new Error(`获取文件夹列表失败：${err.message}`));
      }
      
      // 递归收集所有文件夹路径
      const allFolderPaths = collectAllFolderPaths(boxes);
      
      // 筛选出我们感兴趣的文件夹
      const validFolderNames = new Set(Object.values(CONFIG.MAILBOX_MAP).map(name => name.toLowerCase()));
      const validFolderPaths = allFolderPaths.filter(path => {
        const folderName = path.split('/').pop().toLowerCase();
        return validFolderNames.has(folderName);
      });

      if (validFolderPaths.length === 0) {
        return resolve(null);
      }

      // 并行获取每个文件夹的最新邮件
      Promise.all(validFolderPaths.map(folder => getFolderLatestEmail(imap, folder)))
        .then(folderEmails => {
          const validEmails = folderEmails.filter(Boolean);
          if (validEmails.length === 0) return resolve(null);

          const globalLatest = validEmails.sort((a, b) => new Date(b.date) - new Date(a.date))[0];
          resolve(globalLatest);
        })
        .catch(err => {
          console.error('获取文件夹邮件时发生错误：', err);
          resolve(null);
        });
    });
  });
}


// ===================== 主入口函数 =====================
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
    const { refresh_token, client_id, email, response_type = 'json' } = params;
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

    if (params.mailbox) {
      console.warn(`已忽略mailbox参数：${params.mailbox}，将返回全局最新邮件`);
    }

    console.log("【开始】检查Graph API权限");
    const graph_api_result = await graph_api(refresh_token, client_id);

    if (graph_api_result.status) {
      console.log("【成功】Graph API权限通过，获取全局最新邮件");
      const emailData = await getGlobalLatestEmail(graph_api_result.access_token);
      if (!emailData) {
        return res.status(200).json({ code: 2001, message: "邮箱无任何邮件", data: null });
      }
      if (response_type === 'html') {
        return res.status(200).send(generateEmailHtml(emailData));
      } else {
        return res.status(200).json({ code: 200, message: '全局最新邮件获取成功', data: emailData });
      }
    }

    console.log("【降级】Graph API权限不足或失败，使用IMAP获取全局最新邮件");
    const access_token = await get_access_token(refresh_token, client_id);
    const authString = generateAuthString(email, access_token);
    const imap = new Imap({ ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString });

    imap.once("ready", async () => {
      try {
        const emailData = await getImapGlobalLatestEmail(imap);
        imap.end(); // 在获取到结果后立即关闭连接

        if (!emailData) {
          return res.status(200).json({ code: 2001, message: "邮箱无任何邮件", data: null });
        }
        if (response_type === 'html') {
          res.status(200).send(generateEmailHtml(emailData));
        } else {
          res.status(200).json({ code: 200, message: '全局最新邮件获取成功', data: emailData });
        }
      } catch (err) {
        imap.end();
        console.error('IMAP获取全局最新邮件失败：', err);
        res.status(500).json({ code: 5002, error: `IMAP操作失败：${err.message}` });
      }
    });

    imap.once('error', (err) => {
      console.error('IMAP连接错误：', err);
      res.status(500).json({ code: 5001, error: `IMAP连接失败：${err.message}` });
    });
    
    // 处理imap关闭事件
    imap.once('end', () => {
        console.log('IMAP连接已关闭');
    });

    imap.connect();

  } catch (error) {
    let statusCode = 500;
    let errorCode = 5000;
    if (error.message.includes('HTTP错误！状态码：401')) {
      statusCode = 401; errorCode = 4011; error.message = '认证失效，请刷新refresh_token';
    } else if (error.message.includes('HTTP错误！状态码：403')) {
      statusCode = 403; errorCode = 4031; error.message = '权限不足，需开启Mail.ReadWrite权限';
    } else if (error.message.includes('请求超时')) {
      statusCode = 504; errorCode = 5041;
    }
    res.status(statusCode).json({ code: errorCode, error: `服务器错误：${error.message}` });
  }
};
