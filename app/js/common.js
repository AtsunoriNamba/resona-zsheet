/*
 * Version: 1.3.1
 * Update: SheetTemplateUrlの初期化処理を追加
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
	"attachToRecord": false,
	"attachFormat": 'pdf',
	"download": false,
	"downloadFormat": 'pdf',
	"createAsSheets": false,
	"SheetTemplateUrl": [{name:'', url:''}]
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
    
    const attachToRecord = document.getElementById('attachToRecord');
    const attachFormat = document.getElementById('attachFormat');
    const download = document.getElementById('download');
    const downloadFormat = document.getElementById('downloadFormat');
    const createAsSheetsGroup = document.getElementById('createAsSheetsGroup');
    const createAsSheets = document.getElementById('createAsSheets');
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
        if (!Array.isArray(SETTINGS.SheetTemplateUrl)) {
            SETTINGS.SheetTemplateUrl = [{name:'', url:''}];
        }
        SETTINGS.SheetTemplateUrl.forEach((template, index) => {
            const item = document.createElement('div');
            item.className = 'template-item';
            item.innerHTML = `
                <input type="text" placeholder="テンプレート名" value="${template.name || ''}" data-index="${index}" data-field="name">
                <input type="url" placeholder="URL" value="${template.url || ''}" data-index="${index}" data-field="url">
                <span class="material-icons template-remove" data-index="${index}">remove_circle_outline</span>
            `;
            templateList.appendChild(item);

            // 入力値の変更を監視
            item.querySelectorAll('input').forEach(input => {
                input.addEventListener('change', (e) => {
                    const index = parseInt(e.target.dataset.index);
                    const field = e.target.dataset.field;
                    SETTINGS.SheetTemplateUrl[index][field] = e.target.value;
                    saveWidgetSettings(WidgetKey);
                });
            });

            // 削除ボタンの処理
            item.querySelector('.template-remove').addEventListener('click', (e) => {
                const index = parseInt(e.target.dataset.index);
                SETTINGS.SheetTemplateUrl.splice(index, 1);
                if (SETTINGS.SheetTemplateUrl.length === 0) {
                    SETTINGS.SheetTemplateUrl.push({name:'', url:''});
                }
                saveWidgetSettings(WidgetKey);
                renderTemplateList();
            });
        });
    }

    // テンプレート追加ボタンの処理
    templateAdd.addEventListener('click', () => {
        SETTINGS.SheetTemplateUrl.push({name:'', url:''});
        saveWidgetSettings(WidgetKey);
        renderTemplateList();
    });

    // 設定値をUIに反映
    function updateUI() {
        attachToRecord.checked = SETTINGS.attachToRecord;
        attachFormat.value = SETTINGS.attachFormat;
        download.checked = SETTINGS.download;
        downloadFormat.value = SETTINGS.downloadFormat;
        createAsSheets.checked = SETTINGS.createAsSheets;

        // 添付形式の表示制御
        attachFormat.style.display = SETTINGS.attachToRecord ? 'block' : 'none';

        // ダウンロード形式の表示制御
        downloadFormat.style.display = SETTINGS.download ? 'block' : 'none';

        // シートとして作成の表示制御
        const showCreateAsSheets = SETTINGS.download && SETTINGS.downloadFormat !== 'pdf';
        createAsSheetsGroup.style.display = showCreateAsSheets ? 'block' : 'none';
        if (!showCreateAsSheets) {
            createAsSheets.checked = false;
            SETTINGS.createAsSheets = false;
        }

        // テンプレート一覧の更新
        renderTemplateList();
    }

    // 設定パネルの開閉制御
    settingsBtn.addEventListener('click', () => {
        const isSettingsVisible = settingsPanel.style.display === 'block';
        if (isSettingsVisible) {
            settingsPanel.style.display = 'none';
            operationUI.style.display = 'flex';
            progress.style.display = 'block';
        } else {
            settingsPanel.style.display = 'block';
            operationUI.style.display = 'none';
            progress.style.display = 'none';
        }
    });

    // 設定値の変更を監視
    attachToRecord.addEventListener('change', (e) => {
        SETTINGS.attachToRecord = e.target.checked;
        updateUI();
        saveWidgetSettings(WidgetKey);
    });

    attachFormat.addEventListener('change', (e) => {
        SETTINGS.attachFormat = e.target.value;
        updateUI();
        saveWidgetSettings(WidgetKey);
    });

    download.addEventListener('change', (e) => {
        SETTINGS.download = e.target.checked;
        updateUI();
        saveWidgetSettings(WidgetKey);
    });

    downloadFormat.addEventListener('change', (e) => {
        SETTINGS.downloadFormat = e.target.value;
        updateUI();
        saveWidgetSettings(WidgetKey);
    });

    createAsSheets.addEventListener('change', (e) => {
        SETTINGS.createAsSheets = e.target.checked;
        saveWidgetSettings(WidgetKey);
    });

    updateUI();
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
