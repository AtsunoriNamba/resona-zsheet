/*
 * Version: 1.3.2
 * Update: ダウンロードAPIのパラメータ形式を修正
 */
//レコード情報キャッシュ用オブジェクト
Z = {
	records:{},
	fields:{},
	module:{},
	search:{},
	apiCount:{},
	searchRecord:async function(entity, criteria){
		if(Z.search[entity]){
			if(Z.search[entity][criteria]){
				return Z.search[entity][criteria]
			}
		}else{
			Z.search[entity] = {}
		}
		apiCounter("ZOHO.CRM.API.searchRecords")
		// console.log(`ZOHO.CRM.API.searchRecords({Entity:${entity}, Criteria:${criteria}}})`)
		let res = await ZOHO.CRM.API.searchRecord({Entity:entity, Type:"criteria", Query:criteria})
		Z.search[entity][criteria] = res.data

		for(let record of res.data){
			if(!Z.records[entity]){ Z.records[entity] = {} }
			if(!Z.records[entity][record.id]){ Z.records[entity][record.id] = {} }
			Z.records[entity][record.id].data = record
		}
		return res.data
	},
	getRecord:async function(entity, id){
		if(Z.records[entity]){
			if(Z.records[entity][id]){
				if(Z.records[entity][id].data){ return Z.records[entity][id].data }
			}else{
				Z.records[entity][id] = {}
			}
		}else{
			Z.records[entity] = {}
			Z.records[entity][id] = {}
		}
		apiCounter("ZOHO.CRM.API.getRecord")
		//console.log(`ZOHO.CRM.API.getRecord({Entity:${entity}, RecordID:${id}}})`)
		let res = await ZOHO.CRM.API.getRecord({Entity:entity, RecordID:id})
		let record = res.data[0]
		Z.records[entity][id].data = record
		return record
	},
	getRelatedRecords:async function(entity, id, related){	
		if(Z.records[entity]){
			if(Z.records[entity][id]){
				if(Z.records[entity][id].related){
					if(Z.records[entity][id].related[related]){ return Z.records[entity][id].related[related] }
				}else{
					Z.records[entity][id].related = {}
				}
			}else{
				Z.records[entity][id] = {}
				Z.records[entity][id].related = {}
			}
		}else{
			Z.records[entity] = {}
			Z.records[entity][id] = {}
			Z.records[entity][id].related = {}
		}
		apiCounter("ZOHO.CRM.API.getRelatedRecords")
		//console.log(`ZOHO.CRM.API.getRelatedRecords({Entity:${entity}, RecordID:${id}}}),RelatedList:${related}`)
		let res = await ZOHO.CRM.API.getRelatedRecords({Entity:entity, RecordID:id, RelatedList:related})
		let records = res.data
		Z.records[entity][id].related[related] = records
		return records
	},
	getSubform:async function(entity, id, subform){
		if(Z.fields[`${entity}__subform__${subform}`]){
			return Z.fields[`${entity}__subform__${subform}`]
		}
		apiCounter("ZOHO.CRM.API.getFields")
		//console.log(`ZOHO.CRM.META.getFields({Entity:${entity}, RelatedList:${subform}})`)
		let res = await ZOHO.CRM.META.getFields({Entity:subform})
	},

	getAllRecords:async function(entity){
		apiCounter("ZOHO.CRM.API.getAllRecords")
		//console.log(`ZOHO.CRM.API.getAllRecords({Entity:${entity}})`)
		let module = await Z.getModule(entity)
		let res = await ZOHO.CRM.API.getAllRecords({Entity:entity})
		let records = res.data
		for( record of records){
			if(!Z.records[entity]){ Z.records[entity] = {} }
			if(!Z.records[entity][record.id]){ Z.records[entity][record.id] = {} }
			Z.records[entity][record.id].data = record
		}
		return records
	},
	getFields: async function(entity){
		if(Z.fields[entity]){ return Z.fields[entity] }
		apiCounter("ZOHO.CRM.API.getFields")
		//console.log(`ZOHO.CRM.META.getFields({Entity:${entity}})`)
		let res = await ZOHO.CRM.META.getFields({Entity:entity})
		let fields = res.fields
		Z.fields[entity] = fields
		return fields
	},
	getRelatedList: async function(entity){
		if(Z.fields[entity]?.related_lists){ return Z.fields[entity].related_lists}
		apiCounter("ZOHO.CRM.API.getRelatedList")
		//console.log(`ZOHO.CRM.META.getRelatedList({Entity:${entity}})`)
		let res = await ZOHO.CRM.META.getRelatedList({Entity:entity})
		Z.fields[entity].related_lists = res.related_lists
		return res.related_lists
	},
	getModule: async function(entity){
		if(Z.module[entity]){ return Z.module[entity] }
		apiCounter("ZOHO.CRM.API.getModules")
		//console.log(`ZOHO.CRM.META.getModules()`)
		let res = await ZOHO.CRM.META.getModules()
		let modules = res.modules
		for( module of modules){
			if(module.api_name == entity){
				Z.module[entity] = module
				return module
			}
		}
	},
	getAllModules: async function(){
		if(Z.modules){ return Z.modules }
		apiCounter("ZOHO.CRM.API.getModules")
		//console.log(`ZOHO.CRM.META.getModules()`)
		let res = await ZOHO.CRM.META.getModules()
		let modules = res.modules
		Z.modules = modules
		return modules
	}
}

ZS = {
	copySheetApiCount:0,
	sheetNames:[],
	sheetContents:[],
	timer : new Worker("timer.js"),
	zsApi:async function(url, method, param_type, parameters){
		return new Promise(async function(resolve, reject){
			try{
				ZOHO.CRM.CONNECTION.invoke("zohooauth",{
					"url": url,
					"method" : method,
					"param_type" : param_type,
					"parameters" : parameters
				}).then(function(result){
					console.log(parameters)
					console.log(result)
					if(!result.details || result.details.statusMessage.error_code){
						if(result.details.statusMessage.error_code == "2950"){
							alert("APIの呼び出し回数が上限に達しました。処理件数を減らして再度実行してください。")
							ZOHO.CRM.UI.Popup.close()
							return
						}else{
							debugger
							alert("エラーが発生しました。再度実行してください。")
							ZOHO.CRM.UI.Popup.close()
							return
						}
					}else{
						resolve( result.details.statusMessage )
					}
				}).catch(function(error){
					console.log(error)
					alert("APIの呼び出し回数が上限に達しました。処理件数を減らして再度実行してください。")
					reject(error)
				})
			} catch(error){
				alert("エラーが発生しました。再度実行してください。" + JSON.stringify(error))
				ZOHO.CRM.UI.Popup.close()
				return
			}
		})
	},
	getSheetContents: async function(wbid,wsid,force=false){
		if(ZS.sheetContents[wbid]?.[wsid] && force != true){ return ZS.sheetContents[wbid][wsid] }
		if(!ZS.sheetContents[wbid]){ ZS.sheetContents[wbid] = {} }
		if(!ZS.sheetContents[wbid][wsid]){ ZS.sheetContents[wbid][wsid] = {} }

		apiCounter("worksheet.content.get")
		result = await ZS.zsApi(
			`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
			{
				method:"worksheet.content.get",
				worksheet_id:wsid,
				start_row:1,
				start_column:1,
				visible_rows_only:false,
				visible_columns_only:false
			}
		)
		let contents = result.range_details
		let maxCol = result.used_column
		let maxRow = result.used_row

		//コンテンツのない行を補完
		for(let r=0; r<maxRow; r++){
			let contentRow = contents.find( (row) => row.row_index == r+1 )
			if(!contentRow){
				contentRow = { row_index:r+1, row_details:[] }
				contents.push(contentRow)
			}
		}
		// contentsをrowIndexでソート
		contents.sort( (a,b) => a.row_index - b.row_index )

		//コンテンツのない列を補完
		for(let r=0; r<maxRow; r++){
			let contentRow = contents.find( (row) => row.row_index == r+1 )
			for(let c=0; c<maxCol; c++){
				let contentCol = contentRow.row_details.find( (col) => col.column_index == c+1 )
				if(!contentCol){
					contentCol = { column_index:c+1, content:"" }
					contentRow.row_details.push(contentCol)
				}
			}
			// row_detailsをcolumnIndexでソート
			contentRow.row_details.sort( (a,b) => a.column_index - b.column_index )
		}

		ZS.sheetContents[wbid][wsid] = contents
		return contents
	},
	getWorksheetList: async function(wbid, force=false){
		if(ZS.sheetNames[wbid] && force != true){ return ZS.sheetNames[wbid] }
		if(!ZS.sheetNames[wbid]){ ZS.sheetNames[wbid] = {} }
		apiCounter("worksheet.list")
		let res = await ZS.zsApi(
			`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
			{ method:"worksheet.list" }
		)

		let worksheets = res.worksheet_names
		ZS.sheetNames[wbid] = worksheets
		return worksheets
	},
	deleteRows: async function(wbid,wsid,rows){
		apiCounter("worksheet.rows.delete")
		let result = await ZS.zsApi(
			`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
			{
				method:"worksheet.rows.delete",
				worksheet_id:wsid,
				row_index_array:rows
			}
		)
		return result
	},
	updateSheetViaCsv: async function(wbid,wsid,csv){
		apiCounter("worksheet.csvdata.set")
		let result = await ZS.zsApi(
			`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
			{
				method:"worksheet.csvdata.set",
				worksheet_id:wsid,
				row:1,
				column:1,
				ignore_empty:true,
				data:csv
			}
		)
		console.log("### worksheet.csvdata.set result ###")
		console.log(result)
		return result
	},
	copySheet: async function(wbid,origWsid,newWsName){
		if(ZS.copySheetApiCount == 30){
			apiCounter("workbook.copy")
			let result = await ZS.zsApi(
				`https://sheet.zoho.jp/api/v2/copy`,"POST",1,
				{
					method:"workbook.copy",
					resource_id:wbid,
				}
			)
			WORKING_BOOK_ID = result.resource_id
			ZS.sheetContents[WORKING_BOOK_ID] = ZS.sheetContents[wbid]
			wbid = WORKING_BOOK_ID
			ZS.copySheetApiCount = 0
			await deleteFile(wbid)
		}
		apiCounter("worksheet.copy")
		let result = await ZS.zsApi(
			`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
			{
				method:"worksheet.copy",
				worksheet_id:origWsid,
				new_worksheet_name:newWsName
			}
		)
		let newWs = result.worksheet_names.find( (ws) => ws.worksheet_name == result.new_worksheet_name )
		ZS.sheetContents[wbid][newWs.worksheet_id] = ZS.sheetContents[wbid][origWsid]
		ZS.copySheetApiCount++
		return result
	},
	deleteSheet: async function(wbid,wsid){
		ZS.sheetContents[wbid][wsid] = null
		apiCounter("worksheet.delete")
		let result = await ZS.zsApi(
			`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
			{
				method:"worksheet.delete",
				worksheet_id:wsid
			}
		)
		return result
	},

	// PDF出力 (APIドキュメント参照: https://www.zoho.com/sheet/help/api/v2/#WORKBOOK-Download-workbook)
	exportToPDF: async function(wbid) {
		apiCounter("workbook.download.pdf")

		let result = await ZOHO.CRM.CONNECTION.invoke("zohooauth",{
			"url": `https://sheet.zoho.jp/api/v2/download/${wbid}?method=workbook.download`,
			"method" : "POST",
			"param_type" : 1,
			"parameters" :{
				// method:"workbook.download",
				format:"pdf",
				page_settings:{
					print_type:"WORKBOOK",
					scale:3,
					add_gridlines:false,
					add_button:false,
					add_image:true,
					align_vcenter:false,
					align_hcenter:true,
					margin_left:0,
					margin_right:0,
					margin_top:0.25,
					margin_bottom:0.25,
				}
			},
			"RESPONSE_TYPE":"stream"
		})
		return result
	},
	// Excel出力
	exportToExcel: async function(wbid) {
		apiCounter("workbook.download.xlsx")
		const result = await ZS.zsApi(
			`https://sheet.zoho.jp/api/v2/download/${wbid}?method=workbook.download`,
			"POST",
			2,
			`format=xlsx`
		)
		return result
	}
}

function apiCounter(api){
	if(API_COUNT[api]){
		API_COUNT[api]++
	}else{
		API_COUNT[api] = 1
	}
	console.log(`## API ## ${api} : ${API_COUNT[api]}`)
}

async function createSheetFromTemplate(workbookName, templateUrl){
    let templateId = '';
    if (templateUrl.includes('/')) {
        let parts = templateUrl.split('/');
        let lastPart = parts.pop();

        if (lastPart.includes('?')) {
            templateId = lastPart.split('?')[0];
        } else {
            templateId = lastPart;
        }
    }

	let res = await ZOHO.CRM.CONNECTION.invoke("zohooauth",{
		"url": "https://sheet.zoho.jp/api/v2/createfromtemplate",
		"method" : "POST",
		"param_type" : 1,
		"parameters" : {
			workbook_name:workbookName,
			resource_id:templateId,
			method:"workbook.createfromtemplate"
		}
	})
	return res
}
