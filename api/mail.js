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
    connTimeout: 20000,
    authTimeout: 20000,
    debug: console.log // 开启IMAP调试日志，查看底层通信
  },
  // 垃圾箱所有常见别名（覆盖Outlook/Gmail/国内邮箱）
  JUNK_FOLDERS: [
    'junkemail', 'Junk Email', '垃圾邮件', 
    'junk', '[Gmail]/垃圾邮件', 'Spam', '垃圾'
  ],
  REQUEST_TIMEOUT: 20000,
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
    throw new Error(error.name === "AbortError" ? "请求超时（超过20秒）" : error.message);
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

// 获取access_token（强制请求邮件权限）
async function get_access_token(refresh_token, client_id) {
  try {
    const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'scope': 'https://graph.microsoft.com/Mail.ReadWrite offline_access'
      }).toString()
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP错误！状态码：${response.status}，响应：${errorText}`);
    }

    const data = await response.json();
    console.log('access_token权限范围：', data.scope);
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

// Graph API：优先查询垃圾箱最新邮件
async function getJunkLatestEmail(access_token) {
  try {
    // 先获取所有文件夹，找到垃圾箱ID
    const foldersRes = await fetchWithTimeout(`${CONFIG.GRAPH_API_BASE_URL}/mailFolders`, {
      headers: { "Authorization": `Bearer ${access_token}` }
    });
    if (!foldersRes.ok) throw new Error(`获取文件夹失败：${foldersRes.status}`);

    const foldersData = await foldersRes.json();
    const junkFolder = foldersData.value.find(folder => 
      CONFIG.JUNK_FOLDERS.some(alias => 
        folder.displayName.toLowerCase().includes(alias.toLowerCase())
      )
    );

    if (!junkFolder) {
      console.warn('未找到垃圾箱文件夹，将查询全局邮件');
      return null;
    }

    console.log(`找到垃圾箱：${junkFolder.displayName}，文件夹ID：${junkFolder.id}`);
    // 查询垃圾箱最新10封邮件（确保不遗漏）
    const junkEmailsRes = await fetchWithTimeout(
      `${CONFIG.GRAPH_API_BASE_URL}/mailFolders/${junkFolder.id}/messages?$top=10&$orderby=receivedDateTime desc`,
      { headers: { "Authorization": `Bearer ${access_token}` } }
    );
    if (!junkEmailsRes.ok) throw new Error(`查询垃圾箱邮件失败：${junkEmailsRes.status}`);

    const junkEmailsData = await junkEmailsRes.json();
    const emails = junkEmailsData.value || [];
    console.log(`垃圾箱邮件总数：${emails.length}`);

    if (emails.length === 0) return null;
    // 取最新一封
    const latestEmail = emails.sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime))[0];
    return {
      send: latestEmail['from']?.['emailAddress']?.['address'] || latestEmail['from']?.['emailAddress']?.['name'] || '未知发件人',
      subject: latestEmail['subject'] || '无主题',
      text: latestEmail['bodyPreview'] || '',
      html: latestEmail['body']?.['content'] || '',
      date: latestEmail['receivedDateTime'] || latestEmail['createdDateTime'] || new Date().toISOString(),
      folder: junkFolder.displayName
    };
  } catch (error) {
    console.error('Graph API查询垃圾箱失败：', error);
    return null;
  }
}

// 获取全局最新邮件（Graph API）
async function getGlobalLatestEmail(access_token) {
  // 优先查询垃圾箱
  const junkLatestEmail = await getJunkLatestEmail(access_token);
  if (junkLatestEmail) return junkLatestEmail;

  // 垃圾箱无邮件，查询全局
  try {
    const url = `${CONFIG.GRAPH_API_BASE_URL}/messages?$top=10&$orderby=receivedDateTime desc`;
    const response = await fetchWithTimeout(url, {
      headers: { "Authorization": `Bearer ${access_token}` }
    });
    if (!response.ok) throw new Error(`全局查询失败：${response.status}`);

    const responseData = await response.json();
    const emails = responseData.value || [];
    if (emails.length === 0) return null;

    const latestEmail = emails.sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime))[0];
    return {
      send: latestEmail['from']?.['emailAddress']?.['address'] || '未知发件人',
      subject: latestEmail['subject'] || '无主题',
      text: latestEmail['bodyPreview'] || '',
      html: latestEmail['body']?.['content'] || '',
      date: latestEmail['receivedDateTime'] || new Date().toISOString(),
      folder: '全局最新'
    };
  } catch (error) {
    console.error('全局邮件查询失败：', error);
    throw new Error(`邮件获取异常：${error.message}`);
  }
}

// IMAP：单独查询垃圾箱最新邮件
async function getImapJunkLatestEmail(imap) {
  return new Promise((resolve) => {
    // 遍历所有垃圾箱别名，尝试打开
    const tryOpenJunkFolder = async (folderAliases, index = 0) => {
      if (index >= folderAliases.length) {
        console.warn('所有垃圾箱别名都无法打开，将遍历所有文件夹');
        return resolve(null);
      }

      const folderPath = folderAliases[index];
      try {
        // 只读模式打开垃圾箱
        const box = await new Promise((resolveBox, rejectBox) => {
          imap.openBox(folderPath, false, (err, box) => err ? rejectBox(err) : resolveBox(box));
        });

        console.log(`成功打开垃圾箱：${folderPath}，邮件总数：${box.messages.total}`);
        if (box.messages.total === 0) return tryOpenJunkFolder(folderAliases, index + 1);

        // 搜索垃圾箱所有邮件，取最新
        const results = await new Promise((resolveSearch, rejectSearch) => {
          imap.search(['ALL'], (err, res) => err ? rejectSearch(err) : resolveSearch(res));
        });
        if (results.length === 0) return tryOpenJunkFolder(folderAliases, index + 1);

        const latestUid = results[results.length - 1];
        const f = imap.fetch(latestUid, { bodies: '' });

        f.on("message", async (msg) => {
          try {
            const stream = await new Promise(resolveStream => msg.on("body", resolveStream));
            const mail = await simpleParser(stream);
            resolve({
              send: mail.from?.text || mail.from?.value?.[0]?.address || '未知发件人',
              subject: mail.subject || '无主题',
              text: mail.text || '',
              html: mail.html || '',
              date: new Date(mail.date || Date.now()).toISOString(),
              folder: folderPath
            });
          } catch (e) {
            console.error(`解析垃圾箱邮件失败：`, e);
            tryOpenJunkFolder(folderAliases, index + 1);
          }
        });

        f.on("error", (err) => {
          console.error(`获取垃圾箱邮件失败：`, err);
          tryOpenJunkFolder(folderAliases, index + 1);
        });
      } catch (err) {
        console.warn(`无法打开垃圾箱 ${folderPath}：`, err.message);
        tryOpenJunkFolder(folderAliases, index + 1);
      }
    };

    // 开始尝试所有垃圾箱别名
    tryOpenJunkFolder(CONFIG.JUNK_FOLDERS);
  });
}

// IMAP：获取全局最新邮件（优先垃圾箱）
async function getImapGlobalLatestEmail(imap) {
  // 优先查询垃圾箱
  const junkEmail = await getImapJunkLatestEmail(imap);
  if (junkEmail) return junkEmail;

  // 垃圾箱无邮件，遍历所有文件夹
  return new Promise((resolve, reject) => {
    imap.getBoxes((err, boxes) => {
      if (err) return reject(new Error(`获取文件夹列表失败：${err.message}`));

      const collectPaths = (boxes, parent = '') => {
        let paths = [];
        for (const [name, box] of Object.entries(boxes)) {
          const path = parent ? `${parent}/${name}` : name;
          paths.push(path);
          if (box.children) paths = paths.concat(collectPaths(box.children, path));
        }
        return paths;
      };

      const allPaths = collectPaths(boxes);
      console.log(`遍历所有文件夹（共${allPaths.length}个）`);

      Promise.all(allPaths.map(path => new Promise((res) => {
        imap.openBox(path, false, (err, box) => {
          if (err) return res(null);
          imap.search(['ALL'], (err, results) => {
            if (err || results.length === 0) return res(null);
            const latestUid = results[results.length - 1];
            imap.fetch(latestUid, { bodies: '' }).on("message", async (msg) => {
              const stream = await new Promise(rs => msg.on("body", rs));
              const mail = await simpleParser(stream);
              res({ ...mail, folder: path, date: new Date(mail.date).toISOString() });
            });
          });
        });
      })))
        .then(emails => {
          const validEmails = emails.filter(Boolean);
          if (validEmails.length === 0) return resolve(null);
          const latest = validEmails.sort((a, b) => new Date(b.date) - new Date(a.date))[0];
          resolve({
            send: latest.from?.text || '未知发件人',
            subject: latest.subject || '无主题',
            text: latest.text || '',
            html: latest.html || '',
            date: latest.date,
            folder: latest.folder
          });
        })
        .catch(err => resolve(null));
    });
  });
}

// 检查Graph API权限
async function graph_api(refresh_token, client_id) {
  try {
    const token = await get_access_token(refresh_token, client_id);
    const res = await fetchWithTimeout(`${CONFIG.GRAPH_API_BASE_URL}/mailFolders`, {
      headers: { "Authorization": `Bearer ${token}` }
    });
    return { access_token: token, status: res.ok };
  } catch (error) {
    console.error('Graph API权限检查失败：', error);
    return { access_token: '', status: false };
  }
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

    console.log("【开始】检查Graph API权限");
    const graph_api_result = await graph_api(refresh_token, client_id);

    if (graph_api_result.status) {
      console.log("【成功】Graph API权限通过，优先查询垃圾箱");
      const emailData = await getGlobalLatestEmail(graph_api_result.access_token);
      if (!emailData) {
        return res.status(200).json({ code: 2001, message: "邮箱无任何邮件", data: null });
      }
      if (response_type === 'html') {
        return res.status(200).send(generateEmailHtml(emailData));
      } else {
        return res.status(200).json({ code: 200, message: '垃圾箱最新邮件获取成功', data: emailData });
      }
    }

    console.log("【降级】Graph API权限不足，使用IMAP优先查询垃圾箱");
    const access_token = await get_access_token(refresh_token, client_id);
    const authString = generateAuthString(email, access_token);
    const imap = new Imap({ ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString });

    imap.once("ready", async () => {
      try {
        const emailData = await getImapGlobalLatestEmail(imap);
        imap.end();

        if (!emailData) {
          return res.status(200).json({ code: 2001, message: "邮箱无任何邮件", data: null });
        }
        if (response_type === 'html') {
          res.status(200).send(generateEmailHtml(emailData));
        } else {
          res.status(200).json({ code: 200, message: '垃圾箱最新邮件获取成功', data: emailData });
        }
      } catch (err) {
        imap.end();
        console.error('IMAP操作失败：', err);
        res.status(500).json({ code: 5002, error: `IMAP操作失败：${err.message}` });
      }
    });

    imap.once('error', (err) => {
      console.error('IMAP连接错误：', err);
      res.status(500).json({ code: 5001, error: `IMAP连接失败：${err.message}` });
    });

    imap.once('end', () => console.log('IMAP连接已关闭'));
    imap.connect();

  } catch (error) {
    let statusCode = 500;
    let errorCode = 5000;
    if (error.message.includes('401')) {
      statusCode = 401; errorCode = 4011; error.message = '认证失效，请刷新refresh_token';
    } else if (error.message.includes('403')) {
      statusCode = 403; errorCode = 4031; error.message = '权限不足，需开启Mail.Read权限';
    } else if (error.message.includes('请求超时')) {
      statusCode = 504; errorCode = 5041;
    }
    res.status(statusCode).json({ code: errorCode, error: `服务器错误：${error.message}` });
  }
};
