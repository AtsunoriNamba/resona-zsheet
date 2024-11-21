async function generateSheet(workbookid, moduleApiName, recordId = [], forceGather = false){
	let maxCol, maxRow
	let currentWsName, currentWsId
	let contents

	let gatherBaseSheetId
	let gatherBaseRecordId

	let sheetInfos = []

	let worksheets = await ZS.getWorksheetList(WORKING_BOOK_ID)
	worksheets = await ZS.getWorksheetList(WORKING_BOOK_ID, true)

	zSingleTemplateSheetId = worksheets[0].worksheet_id
	zSingleTemplateContents = await ZS.getSheetContents(WORKING_BOOK_ID, zSingleTemplateSheetId, true)

	for( rId of recordId ){
		let result
		let record = await Z.getRecord(moduleApiName, rId)
		let newSheetName

		newSheetName = `${record.Name}-${record.Deals.name}`

		//テンプレートにするシートの選択
		let templateSheetId

		templateSheetId = zSingleTemplateSheetId
		zCurrentTemplateContents = zSingleTemplateContents

		//ワークシートをコピー
		result = await ZS.copySheet(WORKING_BOOK_ID, templateSheetId, newSheetName)
		currentWsName = newSheetName
		currentWsId = result.worksheet_names.find( (ws) => ws.worksheet_name == currentWsName ).worksheet_id
		sheetInfos.push({sheetId: currentWsId, sheetName: currentWsName, recordId: rId})

		if(!ZS.sheetContents[WORKING_BOOK_ID]){ZS.sheetContents[WORKING_BOOK_ID] = {}}
		if(!ZS.sheetContents[WORKING_BOOK_ID][currentWsId]){ZS.sheetContents[WORKING_BOOK_ID][currentWsId] = {}}

		// 置換前のテンプレート内容を保存
		let originalContents = await ZS.getSheetContents(WORKING_BOOK_ID, currentWsId)
		
		let replacedContents = await replaceSheetVariables(WORKING_BOOK_ID, currentWsId, templateSheetId, moduleApiName, rId)

		let sheetContents = replacedContentsToSheetContents(replacedContents)
		ZS.sheetContents[WORKING_BOOK_ID][currentWsId] = sheetContents
		worksheetContents.push({sheetId: currentWsId, contents:contents})
		
		debugger
		console.log(JSON.stringify(originalContents.slice(80)))
		console.log(JSON.stringify(replacedContents.slice(80)))

		// 置換前後の内容を比較して行を削除
		await clearingRows(WORKING_BOOK_ID, currentWsId, originalContents)
		
		// 各レコードの処理完了時にプログレスバーを更新
		progressNext()
	}
	
	//雛形シートを削除
	await ZS.deleteSheet(WORKING_BOOK_ID, zSingleTemplateSheetId)
}

async function gatheringSheets(workbookid, baseSheetId, gatherSheetIds = []){
	let salesNos = []
	for(let r of gatherSheetIds){
		let record = await Z.getRecord("Deals", r.recordId)
		salesNos.push(record.Sales_No)
	}
	let gatherSheetName = salesNos.join(",")
	if(gatherSheetName.length > 31){
		gatherSheetName = gatherSheetName.substring(0,29) + "…"
	}
	//ワークシートをコピー
	let result = await ZS.copySheet( workbookid, baseSheetId, gatherSheetName )

	let currentWsName = result.new_worksheet_name
	let currentWsId = result.worksheet_names.find( (ws) => ws.worksheet_name == currentWsName ).worksheet_id

	let gatherSheetContents = await ZS.getSheetContents(workbookid, currentWsId)

	//合算シートのコンテンツを１列目以外を空にする
	for(let row of gatherSheetContents){
		for(let col of row.row_details){
			if(col.column_index == 1){ continue }
			col.content = ""
		}
	}

	//生成済みのシートのコンテンツを取得
	let gatheredRow = []
	for(let sheetIdInfo of gatherSheetIds){
		let sheetId = sheetIdInfo.sheetId
		if(sheetId == zSingleTemplateSheetId){ continue }
		if(sheetId == zMultiTemplateSheetId){ continue }
		if(sheetId == currentWsId ){ continue }

		let wsContents = ZS.sheetContents[workbookid][sheetId]

		//各シートの行をループ
		for(let contentsRowIdx in wsContents){
			//１列目が空ならスキップ
			if(wsContents[contentsRowIdx].row_details[0].content == ""){ continue }

			//空でなければ、繰り返しキーを取得
			let key = wsContents[contentsRowIdx].row_details[0].content

			//行のコンテンツが空ならスキップ
			let isEmptyRow = true
			for(let isEmptyCol of wsContents[contentsRowIdx].row_details){
				let c = isEmptyCol.content
				if(
					c != "" &&
					c != key &&
					c != "☆" &&
					c != "※" &&
					c != '" "&CHAR(10)&" "' &&
					!c.match(/[ ]+/)
				){
					isEmptyRow = false
					break
				}
			}
			if(isEmptyRow){ continue }

			//合算シートを行ごとにループ
			for(let gatherRow of gatherSheetContents){
				//繰り返しキーが合致したら
				if(gatherRow.row_details[0].content == key){
					//転記済みの行ならスキップ
					if(gatheredRow.find( r => r == gatherRow.row_index)){ continue }
					//合算シートに行をコピー
					gatherRow.row_details = wsContents[contentsRowIdx].row_details
					//行番号を転記済みとして記録
					gatheredRow.push(gatherRow.row_index)
					break
				}
			}
		}

		await ZS.deleteSheet(workbookid, sheetId)
	}

	//特例措置：請求Noは決め打ちで記入
	// gatherSheetContents.find( row => row.row_index == 2 ).row_details[12].content = "No. " + gatherSheetName.replace(/,/g, "/")
	// console.log(gatherSheetContents)
	//contentsからcsvデータを生成。ダブルクオーテーションと改行コードをエスケープする
	let csvData = gatherSheetContents.map( (row) => row.row_details.map( (col) => `"${col.content.replace(/"/g, '""').replace(/\n/g, '\r')}"` ).join(",") ).join("\n")

	//csvデータでシートを更新
	await ZS.updateSheetViaCsv(workbookid, currentWsId, csvData)
	zCurrentTemplateContents = zMultiTemplateContents
	ZS.sheetContents[workbookid][currentWsId] = replacedContentsToSheetContents(gatherSheetContents)
	return gatherSheetContents
}

/**
 * シート内の不要な行を削除する関数 v4.3
 * @param {string} workbookId - 対象のワークブックID
 * @param {string} sheetId - 対象のシートID
 * @param {Array} originalContents - 置換前のシートコンテンツ
 */
async function clearingRows(workbookId, sheetId, originalContents) {
    // 現在のシートの内容を取得
    let currentContents = await ZS.getSheetContents(workbookId, sheetId)
    
    // 無視する内容のリスト
    const ignoredContents = new Set([
        "", " ", "  ", "☆", "※", '" "&CHAR(10)&" "',
        "     万円", "  "
    ])

    /**
     * コンテンツが有効かどうかを判定する関数
     * @param {string} content - 判定対象のコンテンツ
     * @returns {boolean} - 有効なコンテンツかどうか
     */
    const isValidContent = (content) => {
        if (!content) return false
        if (ignoredContents.has(content)) return false
        if (content.match(/^[ ]+$/)) return false
        return true
    }

    /**
     * 行がリピートキーを持っているかを判定する関数
     * @param {Object} row - 行オブジェクト
     * @returns {boolean} - リピートキーを持っているかどうか
     */
    const hasRepeatKey = (row) => {
        const firstCol = row.row_details[0]
        return isValidContent(firstCol.content)
    }

    /**
     * 行のコンテンツが変更されているかを判定する関数
     * @param {Object} currentRow - 現在の行
     * @param {Object} originalRow - 元の行
     * @returns {boolean} - コンテンツが変更されているかどうか
     */
    const hasContentChanged = (currentRow, originalRow) => {
        return currentRow.row_details.some((col, index) => {
            if (index === 0) return false // キー列は除外
            const currentContent = col.content
            const originalContent = originalRow.row_details[index].content
            return currentContent !== originalContent
        })
    }

    // 1. 各行を個別に判定
    const rowsToDelete = []
    for (let rowIndex = 0; rowIndex < currentContents.length; rowIndex++) {
        const currentRow = currentContents[rowIndex]
        const originalRow = originalContents[rowIndex]

        // リピートキーがない行はスキップ（削除対象外）
        if (!hasRepeatKey(currentRow)) {
            continue
        }

        // リピートキーがあり、コンテンツが変更されていない行は削除対象
        if (!hasContentChanged(currentRow, originalRow)) {
            rowsToDelete.push(rowIndex)
        }
    }

    // 2. 連続する行をまとめる
    const deleteRanges = []
    if (rowsToDelete.length > 0) {
        // 降順にソート
        rowsToDelete.sort((a, b) => b - a)

        let currentRange = {
            start_row: rowsToDelete[0],
            end_row: rowsToDelete[0]
        }

        for (let i = 1; i < rowsToDelete.length; i++) {
            const currentRow = rowsToDelete[i]
            
            // 連続する行の場合
            if (currentRow === currentRange.start_row - 1) {
                currentRange.start_row = currentRow
            } else {
                // 連続しない場合は新しい範囲を開始
                deleteRanges.push({...currentRange})
                currentRange = {
                    start_row: currentRow,
                    end_row: currentRow
                }
            }
        }
        // 最後の範囲を追加
        deleteRanges.push(currentRange)
    }

    // デバッグ情報の出力
    console.log('RowsToDelete:', rowsToDelete)
    console.log('DeleteRanges:', deleteRanges)

    // 3. 削除の実行
    if (deleteRanges.length > 0) {
        const result = await ZS.deleteRows(workbookId, sheetId, deleteRanges)
        console.log('Delete Result:', result)
    }
}

function replacedContentsToSheetContents(replacedContents){
	let sheetContents = replacedContents.map(function(row){
		let cols = row.row_details.map(function(col){
			if(col.content == ""){
				let templateRow = zCurrentTemplateContents.find( r => r.row_index == row.row_index )
				let templateCol = templateRow.row_details.find( c => c.column_index == col.column_index )
				let templateContent = templateCol.content
				return {column_index:col.column_index, content:templateContent}
			}else{
				let replacedContent = col.content
				return {column_index:col.column_index, content:replacedContent}
			}
		})
		return {row_index:row.row_index, row_details:cols}
	})
	return sheetContents
}
