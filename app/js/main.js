window.onload = async function(){
	ZOHO.embeddedApp.on("PageLoad", async function(data){
		debugger
		WidgetKey = `${data.Entity}_${data.ButtonPosition}`
		const orgInfo = await ZOHO.CRM.CONFIG.getOrgInfo()
		const orgId = orgInfo.org[0].zgid
		ApiDomain = orgId == PRDUCTION_ORGID ? "https://www.zohoapis.jp" : "https://crmsandbox.zoho.jp"

		fileNameAddition = orgId == PRDUCTION_ORGID ? "" : "（テスト）"
		widgetData = data

		await loadWidgetSettings(WidgetKey)
		TEMPLATE_CRMVAR = 
		await waitFor("#loadCheck")

		let templateUrl

		// debugger
		document.querySelector("#operation-ui").style.display = "block"


		templateSelectoerSetup(SETTINGS.SheetTemplateUrl, document.querySelector("#invTemplateSelect"))


		function templateSelectoerSetup(v, elm){
			debugger
			let variable = v
			let zTemplateName
			let templateOptionHtml = ""
			for(let entry of v){
				zTemplateName = entry.name
				zSheetTemplate = entry.url
				templateOptionHtml += `<option value="${zSheetTemplate}">${zTemplateName}</option>`
			}
			elm.innerHTML = templateOptionHtml
		}
		
		ZOHO.CRM.UI.Resize({height:"200", width:"540"})

		ENTITY = data.Entity
		ENTITY_IDS = data.EntityId

		FIELDS[ENTITY] = await Z.getFields(ENTITY)
		
		//ENTITYのモジュール情報を取得
		let modules = await Z.getAllModules()
		for(let idx in modules){
			if(modules[idx].api_name == ENTITY){
				MODULE = modules[idx]
				break
			}
		}
		
		initOperationUI()

	})

	ZOHO.embeddedApp.init();
	// loadHtml("https://air.southernwave.net/ZohoWidgetDevelopment/RecordToExcel/widget.html")

	function addWorkbookLink(workbookName, workbookUrl) {
		const linksContainer = document.getElementById('workbookLinks')
		const link = document.createElement('a')
		link.href = workbookUrl
		link.target = '_blank'
		link.className = 'workbook-link'
		link.textContent = workbookName
		linksContainer.appendChild(link)

		// リンクが追加されるたびにUIのサイズを調整
		const currentHeight = 200 + linksContainer.offsetHeight
		ZOHO.CRM.UI.Resize({height: currentHeight.toString(), width:"540"})
	}

	function completeProgress() {
		const progressBar = document.getElementById('progressBar')
		// アニメーションクラスを削除
		progressBar.classList.remove('progress-bar-animated')
		progressBar.classList.remove('progress-bar-striped')
		// プログレスバーを完了状態に
		progressBar.style.width = '100%'
		progressBar.setAttribute('aria-valuenow', '100')
		
		// 作成ボタンのスピナーを停止し、無効化状態を維持
		const generateBtnInProgress = document.getElementById('generateBtnInProgress')
		generateBtnInProgress.classList.remove('spinner-border')
		generateBtnInProgress.classList.remove('spinner-border-sm')
		generateBtnInProgress.innerHTML = '完了'
	}

	async function createZohoSheetDocuments(data){
		// プログレスバーの初期化
		initProgress(data.EntityId.length)



		document.getElementById("generateBtnText").style.display = "none"
		document.getElementById("generateBtnInProgress").style.display = "block"
		document.getElementById("generateBtn").setAttribute("disabled", true)

		//Sheetテンプレートから新規作成
		// debugger
		for(let idx in data.EntityId){
			let recordData = await Z.getRecord(ENTITY, data.EntityId[idx])
			if(IP == "61.200.96.103"){ fileNameAddition += "TEST" }
			let WorkbookName = `${recordData.Name}_${recordData.field5.name}`

			let createBookRes = await createSheetFromTemplate(WorkbookName, zSheetTemplate)
			createdWorkbookId = createBookRes.details.statusMessage.resource_id
			workbookUrl = createBookRes.details.statusMessage.workbook_url

			WORKING_BOOK_ID = createdWorkbookId

			//SheetをWorkDriveに移動
			// if(IP != "61.200.96.103"){
				// レプロ
				// let mvRes = await moveFile(createdWorkbookId, directoryId)
			// }
			
			let worksheets = await ZS.getWorksheetList(createdWorkbookId)
			//debugger
			await generateSheet(WORKING_BOOK_ID, ENTITY, [data.EntityId[idx]])
			
			// リンクを追加
			addWorkbookLink(WorkbookName, workbookUrl)
		}
	
		// 処理完了時の表示更新
		completeProgress()
		document.getElementById("closeBtnArea").style.display = "flex"
		document.getElementById("closeBtn").addEventListener("click", function() {
			ZOHO.CRM.UI.Popup.close()
		})
	}

	async function getFieldInfo(module_apiName){
		if(typeof FIELDS[module_apiName] === "undefined"){
			FIELDS[module_apiName] = (await ZOHO.CRM.META.getFields({Entity:module_apiName})).fields
			return FIELDS[module_apiName]
		}else{
			return FIELDS[module_apiName]
		}
	}
	async function getRelatedListInfo(module_apiName){
		if(typeof RELATED_LISTS[module_apiName] === "undefined"){
			RELATED_LISTS[module_apiName] = (await ZOHO.CRM.META.getRelatedList({Entity:module_apiName})).related_lists
			return RELATED_LISTS[module_apiName]
		}else{
			return RELATED_LISTS[module_apiName]
		}
	}


	function waitFor(selector) {
		return new Promise(function (res, rej) {
			waitForElementToDisplay(selector, 200);
			function waitForElementToDisplay(selector, time) {
				if (document.querySelector(selector) != null) {
					res(document.querySelector(selector));
				}
				else {
					setTimeout(function () {
						waitForElementToDisplay(selector, time);
					}, time);
				}
			}
		});
	}


	function initOperationUI(){
		let genBtn = document.getElementById("generateBtn")
		genBtn.addEventListener("click", function(){
			zSheetTemplate = document.getElementById("invTemplateSelect").value
			createZohoSheetDocuments(widgetData)
		})
	}

}
