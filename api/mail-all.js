const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// 工具函数：生成邮件列表HTML（带跳转链接）
function generateEmailListHtml(emails, mailbox, baseUrl) {
    return `
        <!DOCTYPE html>
        <html lang="zh-CN">
        <head>
            <meta charset="UTF-8">
            <title>邮件列表 - ${mailbox}</title>
            <style>
                * { margin: 0; padding: 0; box-sizing: border-box; }
                body { font-family: "Microsoft YaHei", Arial, sans-serif; padding: 20px; max-width: 1200px; margin: 0 auto; }
                h1 { color: #2c3e50; margin: 20px 0; padding-bottom: 10px; border-bottom: 2px solid #3498db; }
                .mailbox-info { color: #7f8c8d; margin-bottom: 20px; }
                table { width: 100%; border-collapse: collapse; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
                th { background-color: #3498db; color: white; padding: 12px 15px; text-align: left; }
                td { padding: 12px 15px; border-bottom: 1px solid #ecf0f1; }
                tr:hover { background-color: #f8f9fa; cursor: pointer; }
                .subject { font-weight: 500; color: #2c3e50; }
                .from { color: #34495e; }
                .date { color: #7f8c8d; font-size: 0.9em; }
                .preview { color: #666; font-size: 0.9em; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 300px; }
            </style>
        </head>
        <body>
            <h1>邮件列表</h1>
            <div class="mailbox-info">当前文件夹：${mailbox} | 共 ${emails.length} 封邮件</div>
            <table>
                <tr>
                    <th>发件人</th>
                    <th>主题</th>
                    <th>日期</th>
                    <th>预览</th>
                </tr>
                ${emails.map(email => `
                    <tr onclick="window.location='${baseUrl}&message_id=${email.id}'">
                        <td class="from">${email.send || '未知发件人'}</td>
                        <td class="subject">${email.subject || '无主题'}</td>
                        <td class="date">${email.date ? new Date(email.date).toLocaleString() : '未知时间'}</td>
                        <td class="preview">${email.text ? email.text.substring(0, 100) + (email.text.length > 100 ? '...' : '') : '无内容'}</td>
                    </tr>
                `).join('')}
            </table>
        </body>
        </html>
    `;
}

// 工具函数：生成邮件详情HTML（仿示例样式，支持图片、链接、交互按钮）
function generateEmailDetailHtml(email, baseUrl) {
    // 修复邮件内图片链接（确保内嵌图片能加载）
    let content = email.html || email.text || '无内容';
    // 补全相对路径图片的域名（适配Outlook等邮件服务器）
    content = content.replace(/src="\/\//g, 'src="https://');
    content = content.replace(/src="\/(?!\/)/g, 'src="https://outlook.office365.com/');

    return `
        <!DOCTYPE html>
        <html lang="zh-CN">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>邮件详情 - ${email.subject}</title>
            <style>
                body { 
                    font-family: "Microsoft YaHei", Arial, sans-serif; 
                    padding: 0; 
                    margin: 0; 
                    background: #f5f5f5; 
                }
                .email-container {
                    max-width: 800px;
                    margin: 0 auto;
                    background: #fff;
                    box-shadow: 0 0 10px rgba(0,0,0,0.1);
                }
                .email-header {
                    background: #2c3e50;
                    color: #fff;
                    padding: 20px;
                }
                .email-title {
                    font-size: 1.5em;
                    margin-bottom: 10px;
                }
                .email-meta {
                    font-size: 0.9em;
                    color: #ddd;
                }
                .email-meta span {
                    margin-right: 20px;
                }
                .email-content {
                    padding: 20px;
                    line-height: 1.8;
                }
                .email-content img {
                    max-width: 100%;
                    height: auto;
                    margin: 10px 0;
                }
                .email-content a {
                    color: #3498db;
                    text-decoration: underline;
                }
                .close-btn {
                    display: block;
                    width: 100%;
                    padding: 15px;
                    background: #3498db;
                    color: #fff;
                    border: none;
                    font-size: 1em;
                    cursor: pointer;
                }
                .close-btn:hover {
                    background: #2980b9;
                }
            </style>
        </head>
        <body>
            <div class="email-container">
                <div class="email-header">
                    <div class="email-title">${email.subject || '无主题'}</div>
                    <div class="email-meta">
                        <span>发件人: ${email.send || '未知'}</span>
                        <span>日期: ${email.date ? new Date(email.date).toLocaleString() : '未知'}</span>
                    </div>
                </div>
                <div class="email-content">
                    ${content}
                </div>
                <button class="close-btn" onclick="window.location='${baseUrl}'">关闭</button>
            </div>
        </body>
        </html>
    `;
}

// 核心函数：获取邮件列表（Graph API，含邮件ID）
async function get_emails(access_token, mailbox) {
    if (!access_token) return [];
    try {
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${mailbox}/messages?$top=100&$select=id,from,subject,bodyPreview,createdDateTime,body`, {
            method: 'GET',
            headers: { "Authorization": `Bearer ${access_token}` }
        });
        if (!response.ok) return [];
        const responseData = await response.json();
        return responseData.value.map(item => ({
            id: item.id, // 邮件唯一ID，用于详情页跳转
            send: item.from?.emailAddress?.address,
            subject: item.subject,
            text: item.bodyPreview,
            html: item.body?.content,
            date: item.createdDateTime,
        }));
    } catch (error) {
        console.error('获取邮件列表失败：', error);
        return [];
    }
}

// 核心函数：获取单封邮件详情（Graph API）
async function get_email_detail(access_token, message_id) {
    if (!access_token || !message_id) return null;
    try {
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${message_id}?$select=id,from,subject,body,createdDateTime`, {
            method: 'GET',
            headers: { "Authorization": `Bearer ${access_token}` }
        });
        if (!response.ok) return null;
        const data = await response.json();
        return {
            id: data.id,
            send: data.from?.emailAddress?.address,
            subject: data.subject,
            text: data.body?.content, // 纯文本内容
            html: data.body?.content, // HTML内容（含图片、链接）
            date: data.createdDateTime,
        };
    } catch (error) {
        console.error('获取邮件详情失败：', error);
        return null;
    }
}

// 原有辅助函数（保持不变）
async function get_access_token(refresh_token, client_id) {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
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
        throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
    }

    const responseText = await response.text();
    try {
        const data = JSON.parse(responseText);
        return data.access_token;
    } catch (parseError) {
        throw new Error(`Failed to parse JSON: ${parseError.message}, response: ${responseText}`);
    }
}

const generateAuthString = (user, accessToken) => {
    const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
    return Buffer.from(authString).toString('base64');
}

async function graph_api(refresh_token, client_id) {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
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
        throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
    }

    const responseText = await response.text();
    try {
        const data = JSON.parse(responseText);
        return {
            access_token: data.access_token,
            status: data.scope.includes('https://graph.microsoft.com/Mail.ReadWrite')
        };
    } catch (parseError) {
        throw new Error(`Failed to parse JSON: ${parseError.message}, response: ${responseText}`);
    }
}

// 主入口函数
module.exports = async (req, res) => {
    // 1. 验证密码
    const { password } = req.method === 'GET' ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;
    if (password !== expectedPassword && expectedPassword) {
        return res.status(401).json({ error: '认证失败，请提供正确密码' });
    }

    // 2. 获取所有参数
    const params = req.method === 'GET' ? req.query : req.body;
    const { 
        refresh_token, 
        client_id, 
        email, 
        mailbox,
        response_type = 'json',
        message_id, // 单封邮件ID（用于详情页）
    } = params;

    // 3. 检查必要参数
    if (!refresh_token || !client_id || !email || !mailbox) {
        return res.status(400).json({ 
            error: '缺少必要参数：refresh_token、client_id、email 或 mailbox' 
        });
    }

    try {
        // 4. 处理 Graph API 权限验证
        const graph_api_result = await graph_api(refresh_token, client_id);
        if (!graph_api_result.status) {
            // 若Graph API权限不足，此处可 fallback 到 IMAP（需补充IMAP逻辑）
            return res.status(403).json({ error: '无邮件读取权限，请检查应用权限配置' });
        }

        // 5. 适配 mailbox 命名（Graph API 规范）
        let graphMailbox = mailbox.toLowerCase();
        if (!['inbox', 'junkemail'].includes(graphMailbox)) {
            graphMailbox = 'inbox'; // 默认为收件箱
        }

        // 6. 区分“列表页”和“详情页”逻辑
        if (message_id) {
            // 详情页：获取单封邮件并渲染HTML
            const emailDetail = await get_email_detail(graph_api_result.access_token, message_id);
            if (!emailDetail) {
                return res.status(404).json({ error: '邮件不存在或已被删除' });
            }

            // 生成详情页HTML（含返回列表的“关闭”按钮）
            const baseUrl = `${req.protocol}://${req.get('host')}${req.path}?${new URLSearchParams({
                refresh_token,
                client_id,
                email,
                mailbox,
                password,
                response_type: 'html'
            }).toString()}`;
            const html = generateEmailDetailHtml(emailDetail, baseUrl);
            return res.status(200).send(html);
        } else {
            // 列表页：获取邮件列表并渲染HTML
            const emails = await get_emails(graph_api_result.access_token, graphMailbox);
            
            // 生成列表页HTML（每封邮件可点击跳转详情页）
            const baseUrl = `${req.protocol}://${req.get('host')}${req.path}?${new URLSearchParams({
                refresh_token,
                client_id,
                email,
                mailbox,
                password,
                response_type: 'html'
            }).toString()}`;
            const html = generateEmailListHtml(emails, graphMailbox, baseUrl);
            return res.status(200).send(html);
        }

    } catch (error) {
        console.error('系统错误：', error);
        res.status(500).json({ error: error.message });
    }
};
