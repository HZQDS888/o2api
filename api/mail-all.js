const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// 工具函数：将邮件列表转换为 HTML
function generateEmailHtml(emails, mailbox) {
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
                tr:hover { background-color: #f8f9fa; }
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
                    <tr>
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

async function get_emails(access_token, mailbox) {
    if (!access_token) {
        console.log("Failed to obtain access token'");
        return [];
    }

    try {
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${mailbox}/messages?$top=10000`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                "Authorization": `Bearer ${access_token}`
            },
        });

        if (!response.ok) {
            const errorText = await response.text();
            return [];
        }

        const responseData = await response.json();
        return responseData.value.map(item => ({
            send: item['from']['emailAddress']['address'],
            subject: item['subject'],
            text: item['bodyPreview'],
            html: item['body']['content'],
            date: item['createdDateTime'],
        }));

    } catch (error) {
        console.error('Error fetching emails:', error);
        return [];
    }
}

module.exports = async (req, res) => {
    // 验证密码
    const { password } = req.method === 'GET' ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;
    if (password !== expectedPassword && expectedPassword) {
        return res.status(401).json({
            error: 'Authentication failed. Please provide valid credentials.'
        });
    }

    // 获取所有参数（包含 response_type，默认 json）
    const params = req.method === 'GET' ? req.query : req.body;
    const { 
        refresh_token, 
        client_id, 
        email, 
        mailbox,
        response_type = 'json'  // 默认为 JSON 格式
    } = params;

    // 检查必要参数
    if (!refresh_token || !client_id || !email || !mailbox) {
        return res.status(400).json({ 
            error: 'Missing required parameters: refresh_token, client_id, email, or mailbox' 
        });
    }

    try {
        // 处理 Graph API 分支
        const graph_api_result = await graph_api(refresh_token, client_id);
        if (graph_api_result.status) {
            // 适配 mailbox 命名（Graph API 规范）
            let graphMailbox = mailbox.toLowerCase();
            if (!['inbox', 'junkemail'].includes(graphMailbox)) {
                graphMailbox = 'inbox'; // 默认为收件箱
            }

            const emails = await get_emails(graph_api_result.access_token, graphMailbox);
            
            // 根据 response_type 返回对应格式
            if (response_type === 'html') {
                const html = generateEmailHtml(emails, graphMailbox);
                return res.status(200).send(html); // 返回 HTML
            } else {
                return res.status(200).json(emails); // 返回 JSON（默认）
            }
        }

        // 处理 IMAP 分支
        const access_token = await get_access_token(refresh_token, client_id);
        const authString = generateAuthString(email, access_token);

        const imap = new Imap({
            user: email,
            xoauth2: authString,
            host: 'outlook.office365.com',
            port: 993,
            tls: true,
            tlsOptions: { rejectUnauthorized: false }
        });

        const emailList = [];
        imap.once("ready", async () => {
            try {
                // 打开指定邮箱
                await new Promise((resolve, reject) => {
                    imap.openBox(mailbox, true, (err, box) => {
                        if (err) return reject(err);
                        resolve(box);
                    });
                });

                // 搜索邮件
                const results = await new Promise((resolve, reject) => {
                    imap.search(["ALL"], (err, results) => {
                        if (err) return reject(err);
                        resolve(results);
                    });
                });

                // 获取邮件内容
                const f = imap.fetch(results, { bodies: "" });
                f.on("message", (msg, seqno) => {
                    msg.on("body", (stream, info) => {
                        simpleParser(stream, (err, mail) => {
                            if (err) throw err;
                            emailList.push({
                                send: mail.from.text,
                                subject: mail.subject,
                                text: mail.text,
                                html: mail.html,
                                date: mail.date,
                            });
                        });
                    });
                });

                f.once("end", () => imap.end());
            } catch (err) {
                imap.end();
                return res.status(500).json({ error: err.message });
            }
        });

        imap.once('error', (err) => {
            console.error('IMAP error:', err);
            return res.status(500).json({ error: err.message });
        });

        imap.once('end', () => {
            // 根据 response_type 返回对应格式
            if (response_type === 'html') {
                const html = generateEmailHtml(emailList, mailbox);
                res.status(200).send(html); // 返回 HTML
            } else {
                res.status(200).json(emailList); // 返回 JSON（默认）
            }
            console.log('IMAP connection ended');
        });

        imap.connect();

    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ error: error.message }); // 错误响应始终返回 JSON
    }
};
