const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// ===================== 全局配置（只改2处！）=====================
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
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email', 'mailbox', 'code'],
  REQUIRE_CODE: true,
  MANAGE_PASSWORD: 'admin123', // 👉 改成你的管理密码（比如myadmin888）
  MANAGE_TRIGGER: 'manage-page', // 触发管理页面的参数（不用改）
  // 卡密直接存在内存中（无需文件，适配只读系统）
  CODE_LIST: [
    // 初始测试卡密（可直接用，也能通过管理页面修改/新增）
    { code: "XIAOHEI001", remaining: 50, total: 100, expiresAt: "2025-12-31T00:00:00.000Z" },
    { code: "XIAOHEI002", remaining: 30, total: 50, expiresAt: "2025-12-31T00:00:00.000Z" }
  ]
};

// ===================== 工具函数（不用改）=====================
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

function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function escapeJson(str) {
  if (!str) return str;
  return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n');
}

function validateParams(params) {
  const { email } = params;
  const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailReg.test(email)) return new Error("邮箱格式无效，请输入正确的邮箱地址");
  if (params.refresh_token?.length < 50) return new Error("refresh_token格式无效");
  if (params.client_id?.length < 10) return new Error("client_id格式无效");
  return null;
}

// ===================== 卡密核心功能（内存操作，无文件读写）=====================
async function verifyAndDeductCode(code) {
  if (!code) return null;
  const codeObj = CONFIG.CODE_LIST.find(c => c.code === code);

  if (!codeObj) return null;
  const now = new Date();
  if (codeObj.expiresAt && new Date(codeObj.expiresAt) < now) return null;
  if (codeObj.remaining <= 0) return null;

  // 直接修改内存中的次数
  codeObj.remaining -= 1;
  console.log(`卡密 ${code} 调用成功，剩余次数：${codeObj.remaining}`);
  return codeObj;
}

async function addNewCode(code, times = 100, days = 365) {
  if (CONFIG.CODE_LIST.find(c => c.code === code)) return { success: false, msg: '卡密已存在！' };
  
  const now = new Date();
  const expiresAt = new Date(now);
  expiresAt.setDate(expiresAt.getDate() + days);
  
  // 新增卡密到内存
  CONFIG.CODE_LIST.push({
    code,
    remaining: times,
    total: times,
    expiresAt: expiresAt.toISOString()
  });
  return { success: true, msg: `新增卡密【${code}】成功！` };
}

async function updateCodeRemaining(code, new_times) {
  const codeObj = CONFIG.CODE_LIST.find(c => c.code === code);
  if (!codeObj) return { success: false, msg: '卡密不存在！' };
  
  // 修改内存中的次数
  codeObj.remaining = new_times;
  return { success: true, msg: `卡密【${code}】次数已改为${new_times}！` };
}

async function queryAllCodes() {
  // 直接从内存读取卡密列表
  return CONFIG.CODE_LIST.map(item => ({
    code: item.code,
    remaining: item.remaining,
    total: item.total,
    expiresAt: new Date(item.expiresAt).toLocaleDateString()
  }));
}

async function disableCode(code) {
  return updateCodeRemaining(code, 0);
}

// ===================== 可视化管理页面（不用改，自动生效）=====================
function getManagePageHtml(result = '', codes = []) {
  const codeListHtml = codes.map(item => `
    <tr>
      <td>${item.code}</td>
      <td>${item.remaining}</td>
      <td>${item.total}</td>
      <td>${item.expiresAt}</td>
    </tr>
  `).join('');

  return `
  <!DOCTYPE html>
  <html lang="zh-CN">
  <head>
    <meta charset="UTF-8">
    <title>卡密管理后台</title>
    <style>
      body { font-family: Arial, sans-serif; max-width: 1200px; margin: 0 auto; padding: 20px; background: #f5f5f5; }
      .container { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-bottom: 20px; }
      h1, h2 { color: #2d3748; text-align: center; }
      .form-group { margin: 15px 0; }
      label { display: inline-block; width: 120px; font-weight: bold; }
      input { padding: 8px; width: 300px; border: 1px solid #ddd; border-radius: 4px; }
      button { padding: 10px 20px; background: #4299e1; color: white; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px; }
      button:hover { background: #3182ce; }
      .result { margin: 20px 0; padding: 15px; border-radius: 4px; background: #e8f4f8; color: #2d3748; }
      table { width: 100%; border-collapse: collapse; margin-top: 20px; }
      th, td { padding: 12px; text-align: center; border: 1px solid #ddd; }
      th { background: #f8f9fa; }
    </style>
  </head>
  <body>
    <h1>卡密管理后台</h1>

    <!-- 操作结果提示 -->
    <div class="result">${result}</div>

    <!-- 新增卡密 -->
    <div class="container">
      <h2>1. 新增卡密</h2>
      <form method="GET">
        <input type="hidden" name="manage" value="add">
        <div class="form-group">
          <label>卡密：</label>
          <input type="text" name="code" required placeholder="比如VIP888">
        </div>
        <div class="form-group">
          <label>初始次数：</label>
          <input type="number" name="times" value="100" min="1">
        </div>
        <div class="form-group">
          <label>有效期（天）：</label>
          <input type="number" name="days" value="365" min="1">
        </div>
        <div class="form-group">
          <label>管理密码：</label>
          <input type="password" name="admin_pwd" required placeholder="输入你的管理密码">
          <button type="submit">新增</button>
        </div>
      </form>
    </div>

    <!-- 修改卡密次数 -->
    <div class="container">
      <h2>2. 修改卡密次数</h2>
      <form method="GET">
        <input type="hidden" name="manage" value="update">
        <div class="form-group">
          <label>卡密：</label>
          <input type="text" name="code" required placeholder="要修改的卡密">
        </div>
        <div class="form-group">
          <label>新剩余次数：</label>
          <input type="number" name="new_times" required min="0" placeholder="0=禁用">
        </div>
        <div class="form-group">
          <label>管理密码：</label>
          <input type="password" name="admin_pwd" required placeholder="输入你的管理密码">
          <button type="submit">修改</button>
        </div>
      </form>
    </div>

    <!-- 禁用卡密 -->
    <div class="container">
      <h2>3. 禁用卡密</h2>
      <form method="GET">
        <input type="hidden" name="manage" value="disable">
        <div class="form-group">
          <label>卡密：</label>
          <input type="text" name="code" required placeholder="要禁用的卡密">
        </div>
        <div class="form-group">
          <label>管理密码：</label>
          <input type="password" name="admin_pwd" required placeholder="输入你的管理密码">
          <button type="submit">禁用</button>
        </div>
      </form>
    </div>

    <!-- 查看所有卡密 -->
    <div class="container">
      <h2>4. 所有卡密列表</h2>
      <form method="GET">
        <input type="hidden" name="manage" value="query">
        <div class="form-group">
          <label>管理密码：</label>
          <input type="password" name="admin_pwd" required placeholder="输入你的管理密码">
          <button type="submit">查询</button>
        </div>
      </form>
      ${codes.length > 0 ? `
        <table>
          <tr>
            <th>卡密</th>
            <th>剩余次数</th>
            <th>总次数</th>
            <th>有效期至</th>
          </tr>
          ${codeListHtml}
        </table>
      ` : '<p style="text-align:center; margin-top:20px;">点击查询查看所有卡密</p>'}
    </div>
  </body>
  </html>
  `;
}

// ===================== 核心业务函数（不用改）=====================
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

async function fetchOAuthToken(refresh_token, client_id, scope = '') {
  const bodyParams = {
    client_id,
    grant_type: 'refresh_token',
    refresh_token
  };
  if (scope) bodyParams.scope = scope;

  const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(bodyParams).toString()
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`OAuth Token请求失败：状态码${response.status}，响应：${errorText}`);
  }

  return response.json();
}

async function get_access_token(refresh_token, client_id) {
  try {
    const data = await fetchOAuthToken(refresh_token, client_id);
    return data.access_token;
  } catch (error) {
    throw new Error(`获取access_token失败：${error.message}`);
  }
}

const generateAuthString = (user, accessToken) => {
  const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
  return Buffer.from(authString).toString('base64');
};

async function graph_api(refresh_token, client_id) {
  try {
    const data = await fetchOAuthToken(refresh_token, client_id, 'https://graph.microsoft.com/.default');
    const hasMailPermission = data.scope?.includes('https://graph.microsoft.com/Mail.ReadWrite');
    return {
      access_token: data.access_token,
      status: hasMailPermission
    };
  } catch (error) {
    console.error('Graph API权限检查失败：', error);
    return { access_token: '', status: false };
  }
}

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

// ===================== 主入口（双触发管理页面，无文件读写）=====================
module.exports = async (req, res) => {
  try {
    // 👉 两种访问方式（任选一种，必打开管理页面）
    // 方式1：参数触发（推荐）：https://xiaoheifk.cn/api/xiaohei?manage-page=1
    // 方式2：路径触发（备用）：https://xiaoheifk.cn/api/xiaohei/manage-codes
    const isManagePage = req.path === '/manage-codes' || req.query[CONFIG.MANAGE_TRIGGER] === '1';
    
    if (isManagePage) {
      const { manage, admin_pwd } = req.query;
      let result = '请执行对应操作';
      let codes = [];

      // 有操作时验证密码并执行
      if (manage && admin_pwd) {
        if (admin_pwd !== CONFIG.MANAGE_PASSWORD) {
          result = '❌ 管理密码错误！';
        } else {
          switch (manage) {
            case 'add':
              const { code, times, days } = req.query;
              if (!code) result = '❌ 缺少卡密参数！';
              else {
                const addRes = await addNewCode(code, Number(times) || 100, Number(days) || 365);
                result = addRes.success ? `✅ ${addRes.msg}` : `❌ ${addRes.msg}`;
              }
              break;
            case 'update':
              const { code: updateCode, new_times } = req.query;
              if (!updateCode || new_times === undefined) result = '❌ 缺少卡密或新次数！';
              else {
                const updateRes = await updateCodeRemaining(updateCode, Number(new_times));
                result = updateRes.success ? `✅ ${updateRes.msg}` : `❌ ${updateRes.msg}`;
              }
              break;
            case 'disable':
              const { code: disableCode } = req.query;
              if (!disableCode) result = '❌ 缺少卡密参数！';
              else {
                const disableRes = await disableCode(disableCode);
                result = disableRes.success ? `✅ 卡密【${disableCode}】已禁用！` : `❌ ${disableRes.msg}`;
              }
              break;
            case 'query':
              codes = await queryAllCodes();
              result = `✅ 共查询到${codes.length}个卡密`;
              break;
            default:
              result = '❌ 无效操作！';
          }
        }
      }

      // 返回管理页面
      res.status(200).send(getManagePageHtml(result, codes));
      return;
    }

    // 👇 正常API调用逻辑（别人调用时）
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({
        code: 405,
        error: `不支持的请求方法，请使用${CONFIG.SUPPORTED_METHODS.join('或')}`
      });
    }

    // 卡密验证（必须带有效卡密）
    const isGet = req.method === 'GET';
    const params = isGet ? req.query : req.body;
    const { code } = params;
    const codeInfo = await verifyAndDeductCode(code);
    if (!codeInfo) {
      return res.status(401).json({
        code: 4012,
        error: '卡密无效、已过期或次数已耗尽！'
      });
    }

    // 密码验证
    const { password } = params;
    const expectedPassword = process.env.PASSWORD;
    if (password !== expectedPassword && expectedPassword) {
      return res.status(401).json({
        code: 4010,
        error: '认证失败 请联系小黑-QQ:113575320 购买权限再使用'
      });
    }

    // 校验必要参数
    const { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;
    const missingParams = CONFIG.REQUIRED_PARAMS.filter(key => !params[key]);
    if (missingParams.length > 0) {
      return res.status(400).json({
        code: 4001,
        error: `缺少必要参数：${missingParams.join('、')}`
      });
    }

    // 校验参数格式
    const paramError = validateParams(params);
    if (paramError) {
      return res.status(400).json({
        code: 4002,
        error: paramError.message
      });
    }

    // Graph API逻辑
    console.log("【开始】检查Graph API权限");
    const graph_api_result = await graph_api(refresh_token, client_id);
    if (graph_api_result.status) {
      console.log("【成功】Graph API权限通过");
      const normalizedMailbox = CONFIG.MAILBOX_MAP[mailbox.toLowerCase()];
      if (!normalizedMailbox) {
        const supportedMailboxes = Object.keys(CONFIG.MAILBOX_MAP).filter(key => !/[a-z]/.test(key)).join('、');
        return res.status(400).json({
          code: 4003,
          error: `不支持的文件夹名称：${mailbox}，支持的中文文件夹：${supportedMailboxes}`
        });
      }
      mailbox = normalizedMailbox;

      const emailData = await get_emails(graph_api_result.access_token, mailbox, true);
      if (!emailData) {
        const mailboxCN = Object.keys(CONFIG.MAILBOX_MAP).find(key => CONFIG.MAILBOX_MAP[key] === mailbox);
        return res.status(200).json({
          code: 2001,
          message: `当前“${mailboxCN}”文件夹无邮件`,
          data: null,
          remainingCalls: codeInfo.remaining
        });
      }

      if (response_type === 'html') {
        res.status(200).send(generateEmailHtml(emailData));
      } else {
        res.status(200).json({
          code: 200,
          message: '邮件获取成功',
          data: [emailData],
          remainingCalls: codeInfo.remaining
        });
      }
      return;
    }

    // 降级IMAP逻辑
    console.log("【降级】使用IMAP协议");
    const access_token = await get_access_token(refresh_token, client_id);
    const authString = generateAuthString(email, access_token);
    const imap = new Imap({ ...CONFIG.IMAP_CONFIG, user: email, xoauth2: authString });

    imap.once("ready", async () => {
      try {
        await new Promise((resolve, reject) => {
          imap.openBox(mailbox, true, (err, box) => err ? reject(err) : resolve(box));
        });

        const results = await new Promise((resolve, reject) => {
          imap.search(["ALL"], (err, results) => err ? reject(err) : resolve(results));
        });

        if (results.length === 0) {
          imap.end();
          return res.status(200).json({
            code: 2001,
            message: `当前“${mailbox}”文件夹无邮件`,
            data: null,
            remainingCalls: codeInfo.remaining
          });
        }

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
              res.status(200).send(generateEmailHtml(responseData));
            } else {
              res.status(200).json({
                code: 200,
                message: '邮件获取成功',
                data: responseData,
                remainingCalls: codeInfo.remaining
              });
            }
          } catch (err) {
            console.error('解析邮件失败：', err);
            res.status(500).json({
              code: 5003,
              error: `解析邮件失败：${err.message}`,
              remainingCalls: codeInfo.remaining
            });
          }
        });

        f.once("end", () => imap.end());
      } catch (err) {
        imap.end();
        console.error('IMAP操作失败：', err);
        res.status(500).json({
          code: 5002,
          error: `IMAP操作失败：${err.message}`,
          remainingCalls: codeInfo.remaining
        });
      }
    });

    imap.once('error', (err) => {
      console.error('IMAP连接错误：', err);
      res.status(500).json({
        code: 5001,
        error: `IMAP连接失败：${err.message}`
      });
    });

    imap.connect();

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
