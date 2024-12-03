async function generateSheet(workbookid, moduleApiName, recordId = [], forceGather = false){
    let maxCol, maxRow;
    let currentWsName, currentWsId;
    let contents;

    let gatherBaseSheetId;
    let gatherBaseRecordId;

    let sheetInfos = [];

    let worksheets = await ZS.getWorksheetList(WORKING_BOOK_ID, true);

    zSingleTemplateSheetId = worksheets[0].worksheet_id;
    zSingleTemplateContents = await ZS.getSheetContents(WORKING_BOOK_ID, zSingleTemplateSheetId, true);

    const delayInterval = 2000; // 60 * 1000 / 30 = 2000 ms
    for (let i = 0; i < recordId.length; i++) {
        let rId = recordId[i];
        let result;
        let record = await Z.getRecord(moduleApiName, rId);
        let newSheetName;

        newSheetName = `${record.Name}`;

        // Template selection
        let templateSheetId;

        templateSheetId = zSingleTemplateSheetId;
        zCurrentTemplateContents = zSingleTemplateContents;

        // Copy worksheet
        result = await ZS.copySheet(WORKING_BOOK_ID, templateSheetId, newSheetName);
        currentWsName = newSheetName;
        currentWsId = result.worksheet_names.find(ws => ws.worksheet_name == currentWsName).worksheet_id;
        sheetInfos.push({ sheetId: currentWsId, sheetName: currentWsName, recordId: rId });

        if (!ZS.sheetContents[WORKING_BOOK_ID]) { ZS.sheetContents[WORKING_BOOK_ID] = {}; }
        if (!ZS.sheetContents[WORKING_BOOK_ID][currentWsId]) { ZS.sheetContents[WORKING_BOOK_ID][currentWsId] = {}; }

        // Save original template content
        let originalContents = await ZS.getSheetContents(WORKING_BOOK_ID, currentWsId);

        let replacedContents = await replaceSheetVariables(WORKING_BOOK_ID, currentWsId, templateSheetId, moduleApiName, rId);

        let sheetContents = replacedContentsToSheetContents(replacedContents);
        ZS.sheetContents[WORKING_BOOK_ID][currentWsId] = sheetContents;
        worksheetContents.push({ sheetId: currentWsId, contents: contents });

        // Compare and clear rows
        await clearingRows(WORKING_BOOK_ID, currentWsId, originalContents);

        // Introduce delay if more than 30 records
        if (recordId.length > 30 && (i + 1) % 30 === 0) {
            await delay(delayInterval);
        }
    }
    
    // Delete template sheet
    await ZS.deleteSheet(WORKING_BOOK_ID, zSingleTemplateSheetId);


	async function delay(ms) {
		return new Promise(resolve => setTimeout(resolve, ms));
	}
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
        for (let range of deleteRanges) {
            const result = await ZS.deleteRows(workbookId, sheetId, [range])
            console.log('Delete Range:', range)
        }
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
