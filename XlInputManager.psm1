# Memo
# Property に勝手に作られる setter/getter の override方法 (constractor で Add-Member -Value -SecondValue を使う)
# getter / setter 定義 (https://stackoverflow.com/questions/40977472/is-it-possible-to-override-the-getter-setter-functions-in-a-powershell-5-class)
# Value:getter, SecondValue:setter になっている

# Requires -Version 5.0


# 入力元 Excel Tableの列指定用列挙子
# ※ Table列指定用定義のため、値が1から始まっています
enum XlInputTableRow {
	DATETIME	= 1
	MAILTITLE	= 2
	CATEGORY	= 3
	URL			= 4
	URLCHECK	= 5
	MAILID		= 6
}

enum XlInputCheckURL {
	checking	= 0
	checked		= 1
	failed		= 2
}


# XlInput classと呼び出し側でやりとりするデータ型の定義
class XlInputTableData:ICloneable
{
	# Table 1行分の情報
	[string]$Datetime = ""
	[string]$MailTitle = ""
	[string]$Category =""
	[string]$URL = ""
	[string]$URLcheck = ""
	[string]$MailID = ""

	# 付属情報
	hidden [int]$absoluteRowPosition = 0

	[object]Clone()
	{
		return $this.MemberwiseClone()
	}
}



# 入力元 Excel fileの操作 class
class XlInput
{
	# Property定義
	hidden [string]$FileName = ""
	hidden [string]$SheetName = ""
	hidden [string]$TableName = ""
	hidden [object]$xlInstance = $null
	hidden [object]$inputWkbook = $null
	hidden [object]$inputWksheet = $null		# Read対象の sheetは１つなので握ったままでよい

	# constractor
	XlInput ()
	{
	}


	# Excel bookを保存
	[void] save ()
	{
		if( $null -eq $this.inputWkbook )
		{
			errlog "Excel bookを開く前に保存しようとした"
			return
		}

		$this.inputWkbook.Save()
	}


	# Excel bookの終了
	[void] Quit()
	{
		if( $null -eq $this.inputWkbook ){ return }

		[System.Runtime.InteropServices.Marshal]::ReleaseComObject( $this.inputWkbook )
		$this.inputWkbook = $null

		$this.xlInstance.Quit()
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject( $this.xlInstance )
		$this.xlInstance = $null
	}


	# 入力元 Excel file を開き、情報取得元 Tableの存在を確認する
	[boolean] startExcel ([string]$tableName, [string]$sheetName, [string]$fileName)
	{
		try {
			# 名称を覚えておく
			$this.FileName = $fileName
			$this.SheetName = $sheetName
			$this.TableName = $tableName

			# 入力元Excel起動 (別プロセスで起動される。通常起動で後から起動しても同じプロセスにはならない)
			$this.xlInstance = New-Object -comobject Excel.Application
			$this.xlInstance.Visible = $false
			$this.xlInstance.DisplayAlerts = $false		# 警告dialog や 上書き保存dialogなど、script入力できない表示を抑制する設定
			$this.inputWkbook = $this.xlInstance.Workbooks.Open($this.FileName)
		}
		catch {
			errlog "Excel file ($($this.FileName)) が開けません"
			return $false
		}


		try {
			# 入力元Excel file内に、指定名称のシートが存在するか確認
			[string[]]$tableNameArr = @()
			$tableNameArr += $this.inputWkbook.WorkSheets | ForEach-Object { $_.Name }
			if( $tableNameArr -notcontains $this.SheetName ){
				errlog "入力元Excel file ($($this.FileName)) に Sheet ($($this.SheetName)) がありません"
				return $false
			}

			# 指定名称のシート内に、指定名称の Tableが登録されているか確認
			$this.inputWksheet = $this.inputWkbook.WorkSheets($sheetName)
			[string[]]$tableNameArr = @()
			$tableNameArr += $this.inputWksheet.ListObjects | ForEach-Object { $_.Name }
			if( $tableNameArr -notcontains $tableName ){
				errlog "入力元Excel file ($this.FileName) の Sheet ($sheetName) に Table ($tableName) がありません"
				return $false
			}
		}
		catch {
			errlog "Excel file の操作に失敗"
			return $false
		}

		return $true
	}


	# Filter に使う keyを Table内の列から取得
	[string[]] getFilterKey()
	{
		try {
			[object]$tblRange = $this.inputWksheet.ListObjects($this.TableName).DataBodyRange
			$left_top = $tblRange.item([XlInputTableRow]::CATEGORY).Address()
			$right_bottom = $tblRange.item([XlInputTableRow]::CATEGORY + ($tblRange.Rows.Count - 1) * $tblRange.Columns.Count).Address()
			$categoryList = $this.inputWksheet.Range($left_top, $right_bottom) | ForEach-Object { $_.Value() }
			$categoryUniqueList = $categoryList | Sort-Object -Unique
			
			# 一意なkey配列(string[])
			return $categoryUniqueList
		}
		catch {
			errlog "Filterに使う keyの取得に失敗"
			return $null
		}
	}


	# 入力元の仕様書更新履歴Tableから、引数keyで絞り込んだ結果の仕様書格納先URLリストを取得
	[XlInputTableData[]]readFilterData( [string]$filterkey )		# filterkey == category
	{

		Write-Host "$filterkey に対して情報更新します。"

		# 列あたり複数条件(OR)の場合、@("条件1", "条件2") で指定可能
		[int] $flValue = 7			# [Microsoft.Office.Interop.Excel.XlAutoFilterOperator]::xlFilterValues
		[void]$this.inputWksheet.ListObjects($this.TableName).Range.AutoFilter([XlInputTableRow]::CATEGORY, $filterkey, $flValue )

		# get count of visible rows
		# Auto fileter適用後の、表示中itemのみを取得
		[int] $spcellType = 12		# [Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeVisible
		$visibleRange = $this.inputWksheet.ListObjects($this.TableName).DataBodyRange.SpecialCells($spcellType)

		# --- 参考情報 -------------------------------------
		# Auto fileter適用後の行数を取得する
		# 選択中の全item数 ÷ 列数 で行数取得
		# (SpecialCells.Rows.Countは機能しない。複数領域選択状態のため。分割された最初の領域の行数が返されるっぽい)
		$itemCount = $visibleRange.CountLarge
		$colsCount = $visibleRange.Columns.CountLarge		# 行単位で選択されるため列は分断されない。なので Countが使える
		$rowsCount = $itemCount / $colsCount

		debugOut -g1 "絞り込み結果 Cell数 = " -w2 "$itemCount"
		debugOut -g1 "絞り込み結果 列数   = " -w2 "$colsCount"
		debugOut -g1 "絞り込み結果 行数   = " -w2 "$rowsCount"

		# 先頭行と最終行の行番号を取得する(Excel画面の左にある行番号)
		$startRow = $visibleRange.item(1).Row
		# 非表示行を含めて計算してしまうため、テーブル終端には届かない
		$endRow = $visibleRange.End(-4121).Row
		debugOut -g1 "Table Start Row = " -w2 "$startRow"
		debugOut -g1 "Table End   Row = " -w2 "$endRow"
		# -------------------------------------------------


		# 絞り込んだ結果を一行ずつ配列に格納(古い順に並んでいる)
		[XlInputTableData[]]$outlist = @()
		[XlInputTableData]$oneLineData = [XlInputTableData]::new()

		$visibleRange.Rows | ForEach-Object {

			$oneLineData.Datetime	= $_.Cells([XlInputTableRow]::DATETIME).Value()
			$oneLineData.MailTitle	= $_.Cells([XlInputTableRow]::MAILTITLE).Value()
			$oneLineData.Category	= $_.Cells([XlInputTableRow]::CATEGORY).Value()
			$oneLineData.URL		= $_.Cells([XlInputTableRow]::URL).Value()
			$oneLineData.URLcheck	= $_.Cells([XlInputTableRow]::URLCHECK).Value()
			$oneLineData.MailID		= $_.Cells([XlInputTableRow]::MAILID).Value()
			$oneLineData.absoluteRowPosition = $_.Cells(0).Row + 1		# Table開始行はヘッダ部のため、+1している

			if( $oneLineData.URLcheck -ne [string][XlInputCheckURL]::checked ){
				debugOut -g1 "情報取得先：" -w2 "$($oneLineData.URL)"
				$outlist += $oneLineData.Clone()	# 仕様書取得先URL一覧の最後尾に1件追加
			}
			else {
				debugOut -y1 "更新済み情報のためスキップ" -g2 " URL：$($oneLineData.URL)"
			}
		}

		# Excelの Table情報を配列として返す
		return $outlist
	}


	# 入力元 Excel sheetに URLチェック情報書き込み
	[void]writeURLCheckStatus( [string]$strData, [XlInputTableData]$rowInfo )
	{
		[int]$tableTopCol = $this.inputWksheet.ListObjects($this.TableName).Range.item(0).Column
		$absoluteRow = $rowInfo.absoluteRowPosition -  1
		$absoluteCol = $tableTopCol + [XlInputTableRow]::URLCHECK

		$this.inputWksheet.cells( $absoluteRow, $absoluteCol ).Value = $strData
	}


	# Autofilterを解除する
	[void]RemoveAutofilter()
	{
		if( $null -eq $this.inputWksheet.AutoFilter ){ return }

		$this.inputWksheet.AutoFilter.ShowAllData()
	}
}
