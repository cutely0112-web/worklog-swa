const { getAuthenticatedClient } = require('../helpers/graphClient');
const exceljs = require('exceljs');
const stream = require('stream');

// --- 설정값 (필요 시 OneDrive 경로에 맞게 수정) ---
const DATA_FILE_PATH = '/drive/root:/바탕 화면/가드신호/가드신호데이터정리.xlsm:';
const TEMPLATE_FILE_PATH = '/drive/root:/바탕 화면/가드신호/근무일지/근무일지.xlsx';
const PDF_FOLDER_PATH = '/drive/root:/바탕 화면/가드신호/근무일지/pdf';
const LOG_SHEET_NAME = '데이터'; // 데이터가 기록될 시트 이름

module.exports = async function (context, req) {
    context.log('submit-worklog function processed a request.');

    const data = req.body;

    // 필수 데이터 확인
    if (!data || !data.worker || !data.site || !data.date || !data.signature) {
        context.res = { status: 400, body: { error: "Missing required data." } };
        return;
    }

    try {
        const graphClient = getAuthenticatedClient();

        // 1. 데이터 정리 파일(.xlsm)에 제출 내용 한 줄 추가
        await appendDataToLog(graphClient, data);
        
        // 2. 근무일지 템플릿 파일(.xlsx)을 기반으로 PDF 생성
        const pdfPath = await createWorklogPdf(graphClient, data, context);

        context.res = {
            status: 200,
            body: { message: "Successfully submitted and PDF created.", pdfPath: pdfPath }
        };

    } catch (error) {
        context.log.error('Submission failed:', error);
        context.res = {
            status: 500,
            body: { error: 'An error occurred during submission.', details: error.message }
        };
    }
};

/**
 * 데이터 로그 시트에 한 줄을 추가하는 함수
 */
async function appendDataToLog(graphClient, data) {
    const url = `${DATA_FILE_PATH}/workbook/worksheets('${LOG_SHEET_NAME}')/usedRange/add`;
    
    // Excel 시트에 기록될 순서대로 배열 생성
    const rowData = [[
        data.date, data.site, data.worker, data.manager,
        data.timeDay, data.timeNight, data.signCar, data.note
    ]];

    await graphClient.api(url).post({
        index: null, // 마지막 행에 추가
        values: rowData,
    });
}

/**
 * 템플릿 파일에 데이터를 채우고 PDF로 변환하여 저장하는 함수
 */
async function createWorklogPdf(graphClient, data, context) {
    // a. 템플릿 파일 다운로드
    const templateFileStream = await graphClient.api(TEMPLATE_FILE_PATH + '/content').getStream();
    const workbook = new exceljs.Workbook();
    await workbook.xlsx.read(templateFileStream);

    const worksheet = workbook.worksheets[0]; // 첫 번째 시트 사용

    // b. 셀에 데이터 채우기 (셀 주소는 근무일지.xlsx 양식에 맞게 수정 필요)
    worksheet.getCell('C4').value = data.site;      // 근무지
    worksheet.getCell('C5').value = new Date(data.date); // 근무일자
    worksheet.getCell('C6').value = `주간 : ${data.timeDay || ''}\n야간 : ${data.timeNight || ''}`;
    worksheet.getCell('C7').value = data.signCar;   // 싸인카 사용
    worksheet.getCell('C8').value = data.worker;    // 근무자명
    worksheet.getCell('C9').value = data.manager;   // 현장담당자
    worksheet.getCell('C11').value = data.note;     // 비고
    
    // 서명 이미지 삽입
    const signatureBase64 = data.signature.split(';base64,').pop();
    const imageBuffer = Buffer.from(signatureBase64, 'base64');
    const imageId = workbook.addImage({
        buffer: imageBuffer,
        extension: 'png',
    });
    // 서명 이미지 위치 및 크기 조절 (tl: top-left, br: bottom-right)
    // E9 셀에 맞게 조절 (양식에 따라 col, row 인덱스 및 offset 수정 필요)
    worksheet.addImage(imageId, {
        tl: { col: 4.05, row: 8.05 }, // E9 셀 시작점 (0-based index)
        br: { col: 5.8, row: 9.8 }   // F10 셀 끝점 근처
    });


    // c. 수정된 Excel 파일을 메모리 버퍼로 변환
    const modifiedXlsxBuffer = await workbook.xlsx.writeBuffer();

    // d. 수정된 Excel 파일을 임시로 OneDrive에 업로드
    const tempFileName = `temp_${Date.now()}.xlsx`;
    const tempFileUrl = `/drive/root:/temp/${tempFileName}:/content`;
    const uploadResponse = await graphClient.api(tempFileUrl).put(modifiedXlsxBuffer);
    const tempFileId = uploadResponse.id;

    // e. 업로드된 임시 파일을 PDF로 변환
    const pdfFileName = `[근무일지] ${data.site}_${data.worker}_${data.date}.pdf`;
    const convertUrl = `/drive/items/${tempFileId}/content?format=pdf`;
    
    // 변환에 시간이 걸릴 수 있으므로 폴링 로직 추가
    let pdfContent;
    for (let i = 0; i < 5; i++) { // 최대 5번 시도
        try {
            const pdfResponse = await graphClient.api(convertUrl).get();
            if(pdfResponse){
                pdfContent = pdfResponse;
                break;
            }
        } catch(e) {
            context.log(`PDF conversion attempt ${i+1} failed, retrying...`);
            await new Promise(resolve => setTimeout(resolve, 2000)); // 2초 대기
        }
    }
    if (!pdfContent) {
        throw new Error('Failed to convert file to PDF after multiple attempts.');
    }
    
    // f. 생성된 PDF를 최종 폴더에 업로드
    const finalPdfUrl = `${PDF_FOLDER_PATH}/${pdfFileName}:/content`;
    const finalUploadResponse = await graphClient.api(finalPdfUrl).put(pdfContent);

    // g. 임시 파일 삭제
    await graphClient.api(`/drive/items/${tempFileId}`).delete();

    return finalUploadResponse.webUrl;
}
