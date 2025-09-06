const { getAuthenticatedClient } = require('../helpers/graphClient');

// OneDrive에서 데이터를 가져올 파일 경로
// "내파일 > 바탕 화면 > 가드신호 > 가드신호데이터정리.xlsm"
const DATA_FILE_PATH = '/drive/root:/바탕 화면/가드신호/가드신호데이터정리.xlsm:';

// 데이터를 읽어올 시트 이름
const WORKER_SHEET = '근무자';
const SITE_SHEET = '근무처';
const MANAGER_SHEET = '현장담당자';

/**
 * Excel 시트에서 한 열의 데이터를 배열로 읽어오는 함수
 * @param {object} graphClient - Microsoft Graph API 클라이언트
 * @param {string} sheetName - 읽어올 시트 이름
 * @returns {Promise<string[]>} - 데이터 배열
 */
async function getColumnData(graphClient, sheetName) {
    try {
        const url = `${DATA_FILE_PATH}/workbook/worksheets('${sheetName}')/usedRange(valuesOnly=true)`;
        const response = await graphClient.api(url).get();
        
        if (response && response.values) {
            // 첫 번째 열(A열)의 데이터만 추출하여 배열로 만듭니다.
            // filter(Boolean)으로 빈 값은 제거합니다.
            return response.values.map(row => row[0]).filter(Boolean);
        }
        return [];
    } catch (error) {
        console.error(`Error reading sheet ${sheetName}:`, error);
        throw new Error(`Failed to read data from ${sheetName}`);
    }
}

module.exports = async function (context, req) {
    context.log('get-dropdown-data function processed a request.');

    try {
        const graphClient = getAuthenticatedClient();

        // 각 시트의 데이터를 병렬로 가져옵니다.
        const [workers, sites, managers] = await Promise.all([
            getColumnData(graphClient, WORKER_SHEET),
            getColumnData(graphClient, SITE_SHEET),
            getColumnData(graphClient, MANAGER_SHEET),
        ]);

        context.res = {
            status: 200,
            body: {
                workers,
                sites,
                managers,
            },
            headers: { 'Content-Type': 'application/json' }
        };
    } catch (error) {
        context.log.error(error);
        context.res = {
            status: 500,
            body: { error: 'Failed to retrieve data from OneDrive.', details: error.message },
            headers: { 'Content-Type': 'application/json' }
        };
    }
};
