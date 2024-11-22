/*
 * Version: 1.0.2
 * Update: ダウンロードAPIのレスポンス処理を修正
 */

// フォーマット変換処理
async function convertFormat(workbookId, format, shouldCombine = false) {
    try {
        // 統合処理が必要な場合
        if (shouldCombine) {
            await gatheringSheets(workbookId, worksheetContents.map(ws => ws.sheetId));
        }

        // フォーマットに応じた出力処理
        let result;
        switch (format) {
            case 'pdf':
                result = await ZS.exportToPDF(workbookId);
                break;
            case 'excel':
            case 'excel_combined':
                result = await ZS.exportToExcel(workbookId);
                break;
            case 'zohosheet':
            case 'zohosheet_combined':
                return `https://sheet.zoho.jp/sheet/open/${workbookId}`;
            default:
                throw new Error(`Unsupported format: ${format}`);
        }

        if (!result || !result.download_url) {
            throw new Error('Failed to get download URL');
        }

        return result.download_url;
    } catch (error) {
        console.error('Format conversion failed:', error);
        throw error;
    }
}

// CRMレコードへのファイル添付
async function attachToRecord(recordId, fileUrl, fileName) {
    try {
        const response = await ZOHO.CRM.API.attachFile({
            Entity: ENTITY,
            RecordID: recordId,
            File: {
                Name: fileName,
                URL: fileUrl
            }
        });

        if (!response.data || response.data.length === 0) {
            throw new Error('File attachment failed');
        }

        return response.data[0];
    } catch (error) {
        console.error('File attachment failed:', error);
        throw error;
    }
}

// メイン処理関数
async function processTemplate(templateSettings, recordIds, templateUrl) {
    const results = {
        attachments: [],
        downloads: []
    };

    try {
        // 各レコードに対して処理を実行
        for (const recordId of recordIds) {
            // レコード情報の取得
            const record = await ZOHO.CRM.API.getRecord({
                Entity: ENTITY,
                RecordID: recordId
            });

            if (!record.data || record.data.length === 0) {
                throw new Error(`Record not found: ${recordId}`);
            }

            const recordData = record.data[0];
            const fileName = `${recordData.Name}_${recordData.Deals.name}`;

            // テンプレートからワークブックを作成
            const createResult = await createSheetFromTemplate(fileName, templateUrl);
            if (!createResult.details || !createResult.details.statusMessage) {
                throw new Error('Failed to create workbook from template');
            }

            const workbookId = createResult.details.statusMessage.resource_id;

            // データに添付する場合の処理
            if (templateSettings.attachToRecord) {
                const fileUrl = await convertFormat(workbookId, templateSettings.attachFormat);
                const attachment = await attachToRecord(recordId, fileUrl, `${fileName}.${getFileExtension(templateSettings.attachFormat)}`);
                results.attachments.push({
                    recordId,
                    fileName,
                    fileUrl,
                    attachment
                });
            }

            // ダウンロード用の処理
            if (templateSettings.download) {
                const fileUrl = await convertFormat(
                    workbookId, 
                    templateSettings.downloadFormat, 
                    templateSettings.downloadFormat.includes('combined')
                );
                results.downloads.push({
                    fileName: `${fileName}.${getFileExtension(templateSettings.downloadFormat)}`,
                    fileUrl,
                    format: templateSettings.downloadFormat
                });
            }

            // プログレスバーを更新
            progressNext();
        }

        return results;
    } catch (error) {
        console.error('Template processing failed:', error);
        throw error;
    }
}

// ファイル拡張子の取得
function getFileExtension(format) {
    switch (format) {
        case 'pdf':
            return 'pdf';
        case 'excel':
        case 'excel_combined':
            return 'xlsx';
        case 'zohosheet':
        case 'zohosheet_combined':
            return 'zohosheet';
        default:
            return '';
    }
}

// エラーメッセージの表示
function showError(message) {
    const errorMessage = document.getElementById('errorMessage');
    errorMessage.textContent = message;
    errorMessage.classList.add('show');
}

// エラーメッセージの非表示
function hideError() {
    const errorMessage = document.getElementById('errorMessage');
    errorMessage.classList.remove('show');
    errorMessage.textContent = '';
}
