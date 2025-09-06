const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
require('isomorphic-fetch');

let graphClient = null;

/**
 * Microsoft Graph API 클라이언트를 초기화하고 반환합니다.
 * Azure Functions의 환경 변수를 사용하여 인증합니다.
 */
function getAuthenticatedClient() {
    if (graphClient) {
        return graphClient;
    }
    
    // Azure AD 앱 등록 정보 (Azure Functions 환경 변수에 설정 필요)
    const tenantId = process.env.MS_TENANT_ID;
    const clientId = process.env.MS_CLIENT_ID;
    const clientSecret = process.env.MS_CLIENT_SECRET;

    if (!tenantId || !clientId || !clientSecret) {
        throw new Error("Missing Microsoft Graph authentication environment variables.");
    }
    
    // 앱 권한을 사용한 인증
    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    
    // Microsoft Graph 클라이언트 초기화
    graphClient = Client.initWithMiddleware({
        authProvider: {
            getAccessToken: async () => {
                const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");
                return tokenResponse.token;
            },
        },
    });

    return graphClient;
}

module.exports = { getAuthenticatedClient };
