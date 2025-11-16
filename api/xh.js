/**
 * 使用 refresh_token 刷新 access_token
 * @param {string} refresh_token - 旧的 refresh_token
 * @param {string} client_id - 应用的 client_id
 * @param {string} email - 用户邮箱，用于缓存键
 * @param {number} retryCount - 当前重试次数
 * @returns {Promise<Object>} - 包含新令牌信息的对象
 */
async function refreshAccessToken(refresh_token, client_id, email, retryCount = 0) {
  try {
    const response = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id,
        grant_type: 'refresh_token',
        refresh_token,
        // 必须包含 offline_access 才能获取新的 refresh_token
        scope: 'https://graph.microsoft.com/.default offline_access' 
      }).toString()
    });

    if (!response.ok) {
      let errorMessage = 'Unknown error';
      try {
        const errorData = await response.json();
        errorMessage = errorData.error_description || errorData.error;
      } catch (e) {
        errorMessage = await response.text();
      }

      const status = response.status;
      if (status === 401) {
        throw new Error(`[401] refresh_token已失效，请重新授权: ${errorMessage}`);
      } else if (status === 400) {
        throw new Error(`[400] 刷新请求参数错误: ${errorMessage}`);
      } else {
        throw new Error(`[${status}] HTTP错误: ${errorMessage}`);
      }
    }

    const data = await response.json();
    if (!data.access_token) {
      throw new Error("刷新令牌失败：未返回有效的 access_token");
    }

    // 关键：获取并更新新的 refresh_token
    const newRefreshToken = data.refresh_token || refresh_token;
    const isTokenUpdated = data.refresh_token !== undefined;

    // 计算 access_token 过期时间
    const expireTime = Date.now() + (data.expires_in * 1000) - CONFIG.TOKEN_EXPIRE_BUFFER * 1000;
    
    const tokenData = {
      access_token: data.access_token,
      expires_in: data.expires_in,
      expireTime,
      refresh_token: newRefreshToken,
      client_id,
      // 使用配置的 refresh_token 有效期，而不是硬编码
      refreshTokenExpireTime: Date.now() + (CONFIG.REFRESH_TOKEN_VALIDITY_DAYS * 24 * 60 * 60 * 1000)
    };

    // 存储到缓存，覆盖旧的记录
    await cacheTool.set(email, tokenData);
    
    console.log(
      `【令牌刷新成功】邮箱：${email}, ` +
      `access_token有效期: ${Math.floor(data.expires_in / 60)}分钟, ` +
      `refresh_token${isTokenUpdated ? '已更新' : '未更新'}`
    );

    return tokenData;
  } catch (error) {
    // 仅对非致命性错误进行重试
    if (retryCount < CONFIG.REFRESH_RETRY_COUNT &&
        !error.message.includes('refresh_token已失效') &&
        !error.message.includes('刷新请求参数错误')) {
      
      const delay = 1000 * Math.pow(2, retryCount); // 指数退避策略
      console.log(`【令牌刷新失败，将在 ${delay}ms 后进行第 ${retryCount + 1} 次重试】`, error.message);
      
      await new Promise(resolve => setTimeout(resolve, delay));
      return refreshAccessToken(refresh_token, client_id, email, retryCount + 1);
    }

    // 重试次数用尽或遇到致命错误，抛出最终错误
    const finalError = new Error(`令牌刷新失败（已重试 ${retryCount} 次）: ${error.message}`);
    console.error(finalError.message);
    throw finalError;
  }
}

