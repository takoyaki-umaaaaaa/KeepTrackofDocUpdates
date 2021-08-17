# 出力先 Excel fileの操作 class
# Requires -Version 5.0

class XlOutput
{
	# Property定義
	hidden [string]$FileName = ""
	hidden [object]$xlInstance = $null
	hidden [object]$outputWkbook = $null
	hidden [string]$URLstr = $null			# file list 取得元のURL。$SpecHistURLarray配列の中の1要素。現在処理中のURL。
	hidden [string[]]$fileList = @()		# $URLstrから取得した file一覧
	hidden [string]$CategoryStr = $null		# categoryごとに output workbook の sheetに分けられる。categoryは XlInputを Auto filter した key

	# constractor
	XlOutput()
	{
	}


	# Excel bookを保存
	[void] save ()
	{
		if( $null -eq $this.outputWkbook )
		{
			errlog "Excel bookを開く前に保存しようとした"
			return
		}

		$this.outputWkbook.Save()
	}


	# Excel bookの終了
	[void] Quit()
	{
		if( $null -eq $this.outputWkbook ){ return }

		[System.Runtime.InteropServices.Marshal]::ReleaseComObject( $this.outputWkbook )
		$this.outputWkbook = $null

		$this.xlInstance.Quit()
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject( $this.xlInstance )
		$this.xlInstance = $null
	}


	# 出力先 Excel file を開き、出力先 sheetの存在を確認する。sheetがなければ作成する。
	[boolean] StartExcel ([string[]]$categoryList, [string]$fileName )
	{
		try {
			# ファイル名を覚えておく
			$this.FileName = $fileName

			# 出力先 Excel起動
			$this.xlInstance = New-Object -comobject Excel.Application
			$this.xlInstance.Visible = $false
			$this.xlInstance.DisplayAlerts = $false
			$this.outputWkbook = $this.xlInstance.Workbooks.Open($fileName)
		}
		catch {
			errlog "Excel file ($fileName) が開けません"
			return $false
		}


		try {
			# 出力先 Excel workbookに 「category」名の sheetがなければ作成する
			$empty_Var = [System.Type]::Missing
			[object]$outSheets = $this.outputWkbook.WorkSheets
			[string[]]$outSheetNameArray = @()
			[string[]]$outSheetNameArray += $outSheets | ForEach-Object { $_.Name }
			$categoryList | ForEach-Object {
				if( $outSheetNameArray -notcontains $_ ){  # -neの左側が配列の場合、全要素をチェックする
					Write-Host "出力先ファイルにシート[$_]がないため、作成します"
					$xlsheet = $outSheets.Add($empty_Var, $outSheets.item($outSheets.Count))
					$xlsheet.Name = $_
				}
			}
		}
		catch {
			errlog "Excel file の操作に失敗"
			return $false
		}

		return $true
	}



	# 引数URLにアクセスし、仕様書ファイルリストを取得
	[string[]] getFilelistFromURL ( [string]$url )
	{
		$this.URLstr = $url			# 出力ファイルに書き込むため、global scope変数に処理中URLを入れる(使うのは子関数以下なので$urlも参照できるけど)
		$this.fileList = @()

		try {
			# svn serverにファイル一覧を要求
			# ※ このsvnは tortoiseSVNのコマンドライン版を想定
			# 　 通常の svn で同じ動作になるかは未確認
			$this.fileList = Write-Output "R" | svn list $this.URLstr
		}
		catch {
			# tortoiseSVN Ver 1.14.1 ではコマンド失敗でExceptionを設定しているのでここに来る
			errlog "svnでの list取得失敗 - 対象URL：$($this.URLstr)"
			$Error[0] | Select-Object -Property * | errlog
		}

		$this.fileList = Get-Content "$PSScriptRoot\resource\test.txt"	# 動作確認用(仕様書リスト用)

		return $this.fileList
	}


	# 引数filelist情報を、出力先 Excel file に書き込み
	[void] writeListToExcelFile( [string[]]$strList, [string]$cate )
	{
		$this.CategoryStr = $cate		# 処理中の categoryを覚えておく

		$strList | ForEach-Object {
			Write-Host "`nFile名：$_ を処理します"
			$this.writeStringToExcelFile( $_, $cate )
		}
	}


	# 引数の文字列を出力先Excelシートに書き込む
	hidden [void] writeStringToExcelFile ( [string]$outName, [string]$cate )
	{
		# 出力先の Table名を作成 (Table名は [カテゴリ名_仕様書名] としている)
		# まずはファイル名から「仕様書名」を切り出す
		if( $outName -match '(Ver|Rev)' ){
			$specDocName = $outName -replace '(_*Ver|_*Rev).*$', $null		# "Ver", "Rev" 以降の文字列を削除し、仕様書別に区分できる形にする
		}
		else {
			$specDocName = [System.IO.Path]::GetFileNameWithoutExtension( $outName )
		}
		$TableNameToEdit = $this.CategoryStr + "_" + $specDocName		# 処理中のカテゴリ名 + 仕様書名 で 目的の書き込み先Table名を作成
		$TableNameToEdit = $TableNameToEdit -replace '[- ]', '_'		# Table名に空白、ハイフンは使えないようなので
		debugOut -g1 "テーブル名：[" -w1 "$TableNameToEdit" -g2 "] を探します"

		# 更新先Tableを探す
		$sheet = $this.outputWkbook.WorkSheets($this.CategoryStr)
		$searchTableResult = $false
		foreach( $listObj in $sheet.ListObjects ){
			debugOut -g1 "現Table名：[" -w1 "$($listObj.DisplayName())" -g2 "]"

			if ( $TableNameToEdit -eq $listObj.DisplayName() ) {
				Write-Host "更新対象tableが見つかりました($TableNameToEdit)"
				$searchTableResult = $true
				break		# 見つかったので抜ける
			}
		}

		if ( $searchTableResult ) {
			debugOut -w1 "$TableNameToEdit" -g2 "Tableに１行追加します"

			# Table の最後尾 cell位置を探す
			$outRangeTable = $sheet.ListObjects($TableNameToEdit).Range
			$startRow = $outRangeTable.item(1).Row + $outRangeTable.Rows.Count
			$startCol = $outRangeTable.item(1).Column

			# Tableに1行追加
			$this.addRowToExcelTbl( $outName, $outRangeTable, $startRow, $startCol, [ref]$sheet )
		}
		else {
			Write-Host "更新対象テーブルが見つかりません。新規作成します。"

			# 空いている箇所に Tableを新規作成
			[object]$tableListObj = $this.createExcelTbl( $TableNameToEdit, [ref]$sheet )

			$row = $tableListObj.Range.item(1).Row + 1
			$col = $tableListObj.Range.item(1).Column
			debugOut -g1 "新規作成テーブルに書き込み 行：$row, 列：$col"

			# 作成した Table に1件書き込み
			$this.addRowToExcelTbl( $outName, $tableListObj.Range, $row, $col, [ref]$sheet )
		}
	}


	# Excel sheetに仕様書一覧Tableを作成
	hidden [object]createExcelTbl( [string]$tableName, [ref]$wksheet )
	{

		[void]$wksheet.value.Activate()  # おまじない(糞)

		# Tableを作成する範囲を決める
		[int]$newTblRow = 0
		[int]$newTblCol = 0
		[int]$newCells = $wksheet.value.UsedRange.Cells.Count
		if( $newCells -le 1 ){
			# Tableが全くない sheetの場合は初期位置は決め打ちする
			$newTblRow = 4
			$newTblCol = 2
		}
		else {
			$newTblRow = $wksheet.value.UsedRange.item(1).Row
			$newTblCol = $wksheet.value.UsedRange.Columns($wksheet.value.UsedRange.Columns.Count).Column + 2
		}
		[string]$left_top = $wksheet.value.cells($newTblRow, $newTblCol).Address($false, $false)
		[string]$right_bottom = $wksheet.value.cells($newTblRow, $newTblCol + 3).Address($false, $false)

		[object]$newTblRange = $wksheet.value.Range($left_top, $right_bottom)
		[void]$newTblRange.Select()

		# Tableのヘッダ文字列を書き込み
		$wksheet.value.cells($newTblRow, $newTblCol + 0).Value = "更新日時"
		$wksheet.value.cells($newTblRow, $newTblCol + 1).Value = "仕様書名"
		$wksheet.value.cells($newTblRow, $newTblCol + 2).Value = "Ver"
		$wksheet.value.cells($newTblRow, $newTblCol + 3).Value = "URL"

		$wksheet.value.cells($newTblRow, $newTblCol + 0).ColumnWidth = 12.46
		$wksheet.value.cells($newTblRow, $newTblCol + 1).ColumnWidth = 24.88
		$wksheet.value.cells($newTblRow, $newTblCol + 2).ColumnWidth = 8.04
		$wksheet.value.cells($newTblRow, $newTblCol + 3).ColumnWidth = 32.33

		# Tableを作成
		$xlTableValue = 1	# [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange
		$guessYes = 1		# [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
		$outListObjTable = $wksheet.value.ListObjects.Add( $xlTableValue, # 追加対象の種別(範囲 == Table)
			$newTblRange,   # Table追加なら、Tableにする領域
			$false,		 # 外部データソースとリンクするか
			$guessYes )	 # 領域にヘッダを含んでいるか
		$outListObjTable.Name = $tableName
		$outListObjTable.TableStyle = "TableStyleMedium2"

		Write-Host "シート名：$($wksheet.value.Name) に新規Table [$tableName] を追加しました。"
		return $outListObjTable	# 出力
	}


	# Excel sheetの仕様書一覧Tableに1行追加
	[void] addRowToExcelTbl( [string]$documentName, [object]$tableRange, [object]$row, [object]$col, [ref]$wksheet )
	{
		# ファイル名から Verを取得
		if( $documentName -match '(Ver|Rev)' ){
			[string]$ver = $documentName -replace '^.*?(?=(ver|rev))', $null
			if( $ver -ne $null ){
				[int]$extIdx = $ver.LastIndexOf('.')  # 見つからなかったら -1
				if ( $extIdx -gt 0 ) {
					# idx 0 での match は対象外にしたい
					$ver = $ver.Substring( 0, $extIdx )
				}
			}
		}
		else {
			$ver = ""		# ファイル名に Versionが無い場合、Versionについては何も出力しない
		}

		debugOut -g1 "1件追加：行=$row, 列=$col, 名称=$documentName, Ver=$ver"

		# Tableの範囲直下に書き込むと、勝手にTable領域を拡張してくれる
		# 列は左から 更新日時, ファイル名, Ver, URL
		$wksheet.value.cells($row, $col + 0).Value = Get-Date -Format "yyyy/MM/dd"
		$wksheet.value.cells($row, $col + 1).Value = $documentName
		$wksheet.value.cells($row, $col + 2).Value = $ver
		$wksheet.value.cells($row, $col + 3).Value = $this.URLstr

		$empty_Var = [System.Type]::Missing  # 呼び出し先引数のdefault値になる。https://docs.microsoft.com/ja-jp/dotnet/api/system.type.missing?view=net-5.0
		[void]$tableRange.Sort(
			$tableRange.item(1), # Sort key 1
			2,		  # 降順
			$empty_Var, # Sort key 2
			$empty_Var, # ピボットテーブルの場合にのみ指定する箇所
			$empty_Var, # Key 2 の並べ替え順序
			$empty_Var, # Sort key 3
			$empty_Var, # Key 3 の並べ替え順序
			1 )		 # 先頭行にヘッダーを含む

		Write-Host "シート名：$($wksheet.value.Name) Tableに $documentName を追加しました。"
	}
}
