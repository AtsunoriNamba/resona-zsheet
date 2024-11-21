/*
 * Version: 1.3.7
 * Update: 設定項目のレイアウトを横並びに変更
 */

let ENTITY
let ENTITY_IDS
let FIELDS = []
let RELATED_LISTS = []
let RECORD
let MODULE
let Z
let ZS
let TEMPLATE_SELECTOR
let DBG = false
let workbook

let zSheetTemplate
let zSingleTemplateSheetId
let zMultiTemplateSheetId
let zSingleTemplateContents
let zMultiTemplateContents
let zCurrentTemplateContents
let worksheetContents = []

let SETTINGS = {
	"productionOrgId": '',
	"SheetTemplateUrl": [
        {
            "name":'',
            "url":'',
            "attachToRecord": false,
            "attachFormat": 'pdf',
            "download": false,
            "downloadFormat": 'pdf',
        }
    ]
}

let gather
let widgetData
let WidgetKey

// 設定UIの初期化と制御
function initializeSettingsUI() {
    const settingsBtn = document.getElementById('settingsBtn');
    const settingsPanel = document.getElementById('settingsPanel');
    const operationUI = document.getElementById('operation-ui');
    const progress = document.getElementById('progress');
    
    const templateList = document.querySelector('.template-list');
    const templateAdd = document.querySelector('.template-add');

    // 表示領域を固定サイズに設定
    ZOHO.CRM.UI.Resize({
        height: "400",
        width: "700"
    });

    // テンプレート一覧の生成
    function renderTemplateList() {
        templateList.innerHTML = '';
        SETTINGS.SheetTemplateUrl.forEach((template, index) => {
            const item = document.createElement('div');
            item.className = 'template-item';
            item.id = `template-item-${index}`;
            item.innerHTML = `
            <div class="itemsettings">
                <div class="itemsettings-template">
                    <input type="text" placeholder="テンプレート名" value="${template.name || ''}" data-index="${index}" data-field="name">
                    <input type="url" placeholder="URL" value="${template.url || ''}" data-index="${index}" data-field="url">
                </div>
                <div class="itemsettings-postprocesses">
                    <div class="postprocess-label">作成後の処理</div>
                    <div class="postprocess-fields">
                        <div class="setting-field">
                            <div class="form-check">
                                <input type="checkbox" class="form-check-input" id="attachToRecord-${index}" ${template.attachToRecord ? 'checked' : ''}>
                                <label class="form-check-label" for="attachToRecord-${index}">データに添付</label>
                            </div>
                            <select class="form-select form-select-sm" id="attachFormat-${index}" style="display:${template.attachToRecord ? 'block' : 'none'};">
                                <option value="pdf" ${template.attachFormat === 'pdf' ? 'selected' : ''}>PDF</option>
                                <option value="excel" ${template.attachFormat === 'excel' ? 'selected' : ''}>Excel</option>
                                <option value="zohosheet" ${template.attachFormat === 'zohosheet' ? 'selected' : ''}>ZohoSheet</option>
                            </select>
                        </div>

                        <div class="setting-field">
                            <div class="form-check">
                                <input type="checkbox" class="form-check-input" id="download-${index}" ${template.download ? 'checked' : ''}>
                                <label class="form-check-label" for="download-${index}">ダウンロード</label>
                            </div>
                            <select class="form-select form-select-sm" id="downloadFormat-${index}" style="display:${template.download ? 'block' : 'none'};">
                                <option value="pdf" ${template.downloadFormat === 'pdf' ? 'selected' : ''}>PDF</option>
                                <option value="excel" ${template.downloadFormat === 'excel' ? 'selected' : ''}>Excel</option>
                                <option value="excel_combined" ${template.downloadFormat === 'excel_combined' ? 'selected' : ''}>Excel（1ブックにまとめる）</option>
                                <option value="zohosheet" ${template.downloadFormat === 'zohosheet' ? 'selected' : ''}>ZohoSheet</option>
                                <option value="zohosheet_combined" ${template.downloadFormat === 'zohosheet_combined' ? 'selected' : ''}>ZohoSheet（1ブックにまとめる）</option>
                            </select>
                        </div>
                    </div>
                </div>
            </div>
            <span class="material-icons template-remove" data-index="${index}">remove_circle_outline</span>
            `;
            templateList.appendChild(item);

            // 入力値の変更を監視
            item.querySelectorAll('input[type="text"], input[type="url"]').forEach(input => {
                input.addEventListener('change', (e) => {
                    const index = parseInt(e.target.dataset.index);
                    const field = e.target.dataset.field;
                    SETTINGS.SheetTemplateUrl[index][field] = e.target.value;
                    saveWidgetSettings(WidgetKey);
                });
            });

            // チェックボックスの変更を監視
            const attachToRecord = document.getElementById(`attachToRecord-${index}`);
            const attachFormat = document.getElementById(`attachFormat-${index}`);
            const download = document.getElementById(`download-${index}`);
            const downloadFormat = document.getElementById(`downloadFormat-${index}`);

            // データに添付の制御
            attachToRecord.addEventListener('change', (e) => {
                SETTINGS.SheetTemplateUrl[index].attachToRecord = e.target.checked;
                attachFormat.style.display = e.target.checked ? 'block' : 'none';
                saveWidgetSettings(WidgetKey);
            });

            // 添付形式の制御
            attachFormat.addEventListener('change', (e) => {
                SETTINGS.SheetTemplateUrl[index].attachFormat = e.target.value;
                saveWidgetSettings(WidgetKey);
            });

            // ダウンロードの制御
            download.addEventListener('change', (e) => {
                SETTINGS.SheetTemplateUrl[index].download = e.target.checked;
                downloadFormat.style.display = e.target.checked ? 'block' : 'none';
                saveWidgetSettings(WidgetKey);
            });

            // ダウンロード形式の制御
            downloadFormat.addEventListener('change', (e) => {
                SETTINGS.SheetTemplateUrl[index].downloadFormat = e.target.value;
                saveWidgetSettings(WidgetKey);
            });

            // 削除ボタンの処理
            item.querySelector('.template-remove').addEventListener('click', (e) => {
                const index = parseInt(e.target.dataset.index);
                SETTINGS.SheetTemplateUrl.splice(index, 1);
                if (SETTINGS.SheetTemplateUrl.length === 0) {
                    SETTINGS.SheetTemplateUrl.push({
                        "name":'',
                        "url":'',
                        "attachToRecord": false,
                        "attachFormat": 'pdf',
                        "download": false,
                        "downloadFormat": 'pdf',
                    });
                }
                saveWidgetSettings(WidgetKey);
                renderTemplateList();
            });
        });
    }

    // テンプレート追加ボタンの処理
    templateAdd.addEventListener('click', () => {
        SETTINGS.SheetTemplateUrl.push({
            "name":'',
            "url":'',
            "attachToRecord": false,
            "attachFormat": 'pdf',
            "download": false,
            "downloadFormat": 'pdf',
        });
        saveWidgetSettings(WidgetKey);
        renderTemplateList();
    });

    // 設定パネルの開閉制御
    settingsBtn.addEventListener('click', () => {
        const isSettingsVisible = settingsPanel.style.display === 'block';
        if (isSettingsVisible) {
            settingsPanel.style.display = 'none';
            operationUI.style.display = 'flex';
        } else {
            settingsPanel.style.display = 'block';
            operationUI.style.display = 'none';
        }
    });
    renderTemplateList();
}

async function loadWidgetSettings(key){
	let settingData = await getOrgVariable(`widget_${key}`)
	if(!settingData){
		const orgInfo = await ZOHO.CRM.CONFIG.getOrgInfo()
		SETTINGS.productionOrgId = orgInfo.org[0].zgid
		await saveWidgetSettings(key)
		return
	}else{
		SETTINGS = JSON.parse(settingData.value)
        if (!SETTINGS.SheetTemplateUrl) {
            SETTINGS.SheetTemplateUrl = [{name:'', url:''}];
        }
        // 古い設定の移行
        SETTINGS.SheetTemplateUrl.forEach(template => {
            if (template.createAsSheets !== undefined) {
                delete template.createAsSheets;
            }
            if (template.attachFormat === 'xlsx') {
                template.attachFormat = 'excel';
            }
            if (template.downloadFormat === 'xlsx') {
                template.downloadFormat = 'excel';
            }
            if (template.attachFormat === 'url') {
                template.attachFormat = 'zohosheet';
            }
            if (template.downloadFormat === 'url') {
                template.downloadFormat = 'zohosheet';
            }
            // 添付形式のcombinedオプションを通常のフォーマットに変換
            if (template.attachFormat === 'excel_combined') {
                template.attachFormat = 'excel';
            }
            if (template.attachFormat === 'zohosheet_combined') {
                template.attachFormat = 'zohosheet';
            }
        });
        initializeSettingsUI();
	}
}

async function saveWidgetSettings(key){
	let settingData = await getOrgVariable(`widget_${key}`)
	if(!settingData){
		await createOrgVariables(`widget_${key}`)
		settingData = await getOrgVariable(`widget_${key}`)
	}
	await updateOrgVariavbles(`widget_${key}`, SETTINGS)
}

async function getOrgVariable(key){
	let result = await ZOHO.CRM.CONNECTION.invoke("zohooauth", {
		"url":`${ApiDomain}/crm/v7/settings/variables`,
		"method" : "GET",
	})
	let variables = result.details.statusMessage.variables.find((v) => v.name == key)
	return variables
}

async function createOrgVariables(key){
	let result = await ZOHO.CRM.CONNECTION.invoke("zohooauth", {
		"url":`${ApiDomain}/crm/v7/settings/variables`,
		"method" : "POST",
		"param_type" : 2,
		"headers" : {
			"Content-Type" : "application/json"
		},
		"parameters" : {
			"variables":[
				{
					"variable_group":{
						"name":"General"
					},
					"name":key,
					"api_name":key,
					"value":JSON.stringify(SETTINGS),
					"type":"textarea"
				}
			]
		}
	})
	return result
}

async function updateOrgVariavbles(key,val){
	let variable = await getOrgVariable(key)
	let result = await ZOHO.CRM.CONNECTION.invoke("zohooauth", {
		"url":`${ApiDomain}/crm/v7/settings/variables`,
		"method" : "PUT",
		"param_type" : 2,
		"headers" : {
			"Content-Type" : "application/json"
		},
		"parameters" : {
			"variables":[
				{
					"id":variable.id,
					"value":JSON.stringify(SETTINGS),
				}
			]
		}
	})
	return result
}

// プログレスバーの制御用変数
let currentProgress = 0
let totalRecords = 0

function initProgress(total) {
	currentProgress = 0
	totalRecords = total
	const progressBar = document.getElementById('progressBar')
	progressBar.style.width = '0%'
	progressBar.setAttribute('aria-valuenow', '0')
}

function progressNext() {
	currentProgress++
	const percentage = (currentProgress / totalRecords) * 100
	const progressBar = document.getElementById('progressBar')
	progressBar.style.width = percentage + '%'
	progressBar.setAttribute('aria-valuenow', percentage)
}

const PRDUCTION_ORGID = "90001619930"
let TEMPLATE_CRMVAR
let ENVIROMENT = "production"
let ApiDomain = "https://www.zohoapis.jp"
let fileNameAddition = ""

let WORKING_BOOK_ID

let GatherSalesNumbers = ""

let API_COUNT = {}

let IP
