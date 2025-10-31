const axios = require('axios');

// 原始输入参数
const input = 'DannyFlores5923@outlook.com----tzgjjas945806----dbc8e03a-b00c-46bd-ae65-b683e7707cb0----M.C560_SN1.0.U.-Cl!MaFF53ULSHroBeqY91tHBsG5aQBzFVxmHMg5X*SUHOzYYn*ONRXTDGwm8L7GW4BG3BPxLfi69coDKeY75UpyCRc4lqPAhwi43QVI2wOhOoFrbPv3j4C79SSIHugPrsNdMRQxMvjEWwMIEHn6i5Lvgdv9j5vl8znqmbiWK3MtDg7cFlYM!LEfRHvQ2oMYmgFdGWpsJG4RyJfPQwJ!8dtRfC7M3dUOOKAY8QHUOcqICIRqZKml5kfJhEAv1aPuSICe8i8xgGUjb!OmjtpBKoL*MkwBWwfnKk14ERxpgEm9Abeo28EbFGgYFDg9BmQ*!N15rvrj22I63KQVJv7kK8s8sZBMKB**is7MAhZLMhyQ6nQLCWAiLnZu4L3wRv*7ltxL2UNzzXG1Cljk3reByPfg$';
const [email, password, clientId, oldRefreshToken] = input.split('----');

// 验证参数完整性
if (!email || !password || !clientId || !oldRefreshToken) {
    throw new Error('参数格式错误，必须包含：邮箱----密码----clientId----refreshToken');
}

// 刷新令牌并按指定格式输出
async function refreshAndFormatOutput() {
    try {
        const response = await axios.post(
            'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
            new URLSearchParams({
                client_id: clientId,
                grant_type: 'refresh_token',
                refresh_token: oldRefreshToken,
                scope: 'https://graph.microsoft.com/.default'
            }),
            {
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                timeout: 15000
            }
        );

        const newRefreshToken = response.data.refresh_token;
        if (!newRefreshToken) {
            throw new Error('未获取到新的refresh_token');
        }

        // 按指定格式拼接输出：邮箱----密码----clientId----新refreshToken
        console.log(`${email}----${password}----${clientId}----${newRefreshToken}`);
        return `${email}----${password}----${clientId}----${newRefreshToken}`;

    } catch (error) {
        const errorMsg = error.response 
            ? `刷新失败 [${error.response.status}]: ${JSON.stringify(error.response.data)}`
            : `刷新失败: ${error.message}`;
        console.error(errorMsg);
        process.exit(1);
    }
}

// 执行
refreshAndFormatOutput();
