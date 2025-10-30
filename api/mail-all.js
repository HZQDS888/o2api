const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// 工具函数：将邮件列表转换为HTML
function generateEmailHtml(emails, mailbox) {
    return `
        <!DOCTYPE html>
        <html>
        <head>
            <title>Emails - ${mailbox}</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                h1 { color: #333; }
                table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
                th { background-color: #f5f5f5; }
                tr:hover { background-color: #f9f9f9; }
                .subject { font-weight: 500; }
                .date { color: #666; font-size: 0.9em; }
            </style>
        </head>
        <body>
            <h1>Emails in ${mailbox} (${emails.length} items)</h1>
            <table>
                <tr>
                    <th>From</th>
                    <th>Subject</th>
                    <th>Date</th>
                    <th>Preview</th>
                </tr>
                ${emails.map(email => `
                    <tr>
                        <td>${email.send || 'N/A'}</td>
                        <td class="subject">${email.subject || 'No subject'}</td>
                        <td class="date">${email.date ? new Date(email.date).toLocaleString() : 'N/A'}</td>
                        <td>${email.text ? email.text.substring(0, 100) + '...' : 'No preview'}</td>
                    </tr>
                `).join('')}
            </table>
        </body>
        </html>
    `;
}

async function get_access_token(refresh_token, client_id) {
    // （原逻辑不变）
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
    // （原逻辑不变）
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
    // （原逻辑不变，返回原始邮件数组，不在此转换格式）
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
    // 1. 验证密码
    const { password } = req.method === 'GET' ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;
    if (password !== expectedPassword && expectedPassword) {
        return res.status(401).json({
            error: 'Authentication failed. Please provide valid credentials.'
        });
    }

    // 2. 获取所有参数（包括 response_type，默认 json）
    const params = req.method === 'GET' ? req.query : req.body;
    const { 
        refresh_token, 
        client_id, 
        email, 
        mailbox,
        response_type = 'json'  // 默认返回 JSON
    } = params;

    // 3. 检查必要参数
    if (!refresh_token || !client_id || !email || !mailbox) {
        return res.status(400).json({ 
            error: 'Missing required parameters: refresh_token, client_id, email, or mailbox' 
        });
    }

    try {
        // 4. 处理 Graph API 分支
        const graph_api_result = await graph_api(refresh_token, client_id);
        if (graph_api_result.status) {
            // 适配 mailbox 命名（Graph API 与 IMAP 可能不同）
            let graphMailbox = mailbox.toLowerCase();
            if (!['inbox', 'junkemail'].includes(graphMailbox)) {
                graphMailbox = 'inbox';
            }

            const emails = await get_emails(graph_api_result.access_token, graphMailbox);
            
            // 根据 response_type 返回对应格式
            if (response_type === 'html') {
                const html = generateEmailHtml(emails, graphMailbox);
                return res.status(200).send(html);
            } else {
                return res.status(200).json(emails);
            }
        }

        // 5. 处理 IMAP 分支
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
                res.status(200).send(html);
            } else {
                res.status(200).json(emailList);
            }
            console.log('IMAP connection ended');
        });

        imap.connect();

    } catch (error) {
        console.error('Error:', error);
        // 错误响应也可根据格式调整（这里简化为始终返回JSON错误）
        res.status(500).json({ error: error.message });
    }
};
