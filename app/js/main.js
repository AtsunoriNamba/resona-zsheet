window.onload = async function(){
	ZOHO.embeddedApp.on("PageLoad", async function(data){
		// debugger
		WidgetKey = `${data.Entity}_${data.ButtonPosition}`
		const orgInfo = await ZOHO.CRM.CONFIG.getOrgInfo()
		const orgId = orgInfo.org[0].zgid
		ApiDomain = orgId == PRDUCTION_ORGID ? "https://www.zohoapis.jp" : "https://crmsandbox.zoho.jp"

		fileNameAddition = orgId == PRDUCTION_ORGID ? "" : "（テスト）"
		widgetData = data

		await loadWidgetSettings(WidgetKey)
		TEMPLATE_CRMVAR = 
		await waitFor("#loadCheck")

		// debugger
		document.querySelector("#operation-ui").style.display = "block"

		templateSelectoerSetup(SETTINGS.SheetTemplateUrl, document.querySelector("#invTemplateSelect"))

		function templateSelectoerSetup(v, elm){
			let templateCheckboxHtml = ""
			for(let entry of v){
				const id = `template-${v.indexOf(entry)}`
				templateCheckboxHtml += `
					<div class="template-checkbox-item">
						<div class="form-check">
							<input class="form-check-input" type="checkbox" 
								value="${entry.url}" 
								data-index="${v.indexOf(entry)}"
								id="${id}">
							<label class="form-check-label" for="${id}">
								${entry.name}
							</label>
						</div>
					</div>`
			}
			elm.innerHTML = templateCheckboxHtml
		}
		
		ZOHO.CRM.UI.Resize({height:"300", width:"900"})

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

	ZOHO.embeddedApp.init()

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

	async function createZohoSheetDocuments(data) {
		try {
			// UIの更新
			document.getElementById("generateBtnText").style.display = "none"
			document.getElementById("generateBtnInProgress").style.display = "block"
			document.getElementById("generateBtn").setAttribute("disabled", true)

			// 選択されたテンプレートの設定を取得
			const templateContainer = document.getElementById("invTemplateSelect")
			const selectedTemplates = Array.from(templateContainer.querySelectorAll('input[type="checkbox"]:checked'))
				.map(checkbox => ({
					index: checkbox.dataset.index,
					url: checkbox.value
				}))

			if (selectedTemplates.length === 0) {
				throw new Error('テンプレートが選択されていません')
			}

			// プログレスバーの初期化（テンプレート数 × レコード数）
			initProgress(data.EntityId.length * selectedTemplates.length)

			let allResults = { downloads: [] }

			// 各テンプレートを順次処理
			for (const template of selectedTemplates) {
				const templateSettings = SETTINGS.SheetTemplateUrl[template.index]
				const results = await processTemplate(templateSettings, data.EntityId, template.url)
				allResults.downloads = allResults.downloads.concat(results.downloads)
			}

			// 結果の表示
			if (allResults.downloads.length > 0) {
				allResults.downloads.forEach(download => {
					addWorkbookLink(download.fileName, download.fileUrl)
				})
			}

			// 処理完了時の表示更新
			completeProgress()
			document.getElementById("closeBtnArea").style.display = "flex"
			document.getElementById("closeBtn").addEventListener("click", function() {
				ZOHO.CRM.UI.Popup.close()
			})

		} catch (error) {
			console.error('Document creation failed:', error)
			
			// エラー表示
			const generateBtn = document.getElementById("generateBtn")
			const generateBtnText = document.getElementById("generateBtnText")
			const generateBtnInProgress = document.getElementById("generateBtnInProgress")
			
			generateBtnText.style.display = "block"
			generateBtnText.textContent = "エラーが発生しました"
			generateBtnText.style.color = "red"
			generateBtnInProgress.style.display = "none"
			generateBtn.classList.remove("btn-primary")
			generateBtn.classList.add("btn-danger")

			// エラーメッセージの表示
			const errorMessage = document.getElementById("errorMessage")
			errorMessage.textContent = error.message
			errorMessage.style.display = "block"
		}
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
			createZohoSheetDocuments(widgetData)
		})
	}
}
