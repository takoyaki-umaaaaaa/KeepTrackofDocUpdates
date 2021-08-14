# Memo
# Property �ɏ���ɍ���� setter/getter �� override���@ (constractor �� Add-Member -Value -SecondValue ���g��)
# getter / setter ��` (https://stackoverflow.com/questions/40977472/is-it-possible-to-override-the-getter-setter-functions-in-a-powershell-5-class)
# Value:getter, SecondValue:setter �ɂȂ��Ă���

# Requires -Version 5.0


# ���͌� Excel Table�̗�w��p�񋓎q
# �� Table��w��p��`�̂��߁A�l��1����n�܂��Ă��܂�
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


# XlInput class�ƌĂяo�����ł��Ƃ肷��f�[�^�^�̒�`
class XlInputTableData:ICloneable
{
	# Table 1�s���̏��
	[string]$Datetime = ""
	[string]$MailTitle = ""
	[string]$Category =""
	[string]$URL = ""
	[string]$URLcheck = ""
	[string]$MailID = ""

	# �t�����
	hidden [int]$absoluteRowPosition = 0

	[object]Clone()
	{
		return $this.MemberwiseClone()
	}
}



# ���͌� Excel file�̑��� class
class XlInput
{
	# Property��`
	hidden [string]$FileName = ""
	hidden [string]$SheetName = ""
	hidden [string]$TableName = ""
	hidden [object]$xlInstance = $null
	hidden [object]$inputWkbook = $null
	hidden [object]$inputWksheet = $null		# Read�Ώۂ� sheet�͂P�Ȃ̂ň������܂܂ł悢

	# constractor
	XlInput ()
	{
	}


	# Excel book��ۑ�
	[void] save ()
	{
		if( $null -eq $this.inputWkbook )
		{
			errlog "Excel book���J���O�ɕۑ����悤�Ƃ���"
			return
		}

		$this.inputWkbook.Save()
	}


	# Excel book�̏I��
	[void] Quit()
	{
		if( $null -eq $this.inputWkbook ){ return }

		[System.Runtime.InteropServices.Marshal]::ReleaseComObject( $this.inputWkbook )
		$this.inputWkbook = $null

		$this.xlInstance.Quit()
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject( $this.xlInstance )
		$this.xlInstance = $null
	}


	# ���͌� Excel file ���J���A���擾�� Table�̑��݂��m�F����
	[boolean] startExcel ([string]$tableName, [string]$sheetName, [string]$fileName)
	{
		try {
			# ���̂��o���Ă���
			$this.FileName = $fileName
			$this.SheetName = $sheetName
			$this.TableName = $tableName

			# ���͌�Excel�N�� (�ʃv���Z�X�ŋN�������B�ʏ�N���Ōォ��N�����Ă������v���Z�X�ɂ͂Ȃ�Ȃ�)
			$this.xlInstance = New-Object -comobject Excel.Application
			$this.xlInstance.Visible = $false
			$this.xlInstance.DisplayAlerts = $false		# �x��dialog �� �㏑���ۑ�dialog�ȂǁAscript���͂ł��Ȃ��\����}������ݒ�
			$this.inputWkbook = $this.xlInstance.Workbooks.Open($this.FileName)
		}
		catch {
			errlog "Excel file ($($this.FileName)) ���J���܂���"
			return $false
		}


		try {
			# ���͌�Excel file���ɁA�w�薼�̂̃V�[�g�����݂��邩�m�F
			[string[]]$tableNameArr = @()
			$tableNameArr += $this.inputWkbook.WorkSheets | ForEach-Object { $_.Name }
			if( $tableNameArr -notcontains $this.SheetName ){
				errlog "���͌�Excel file ($($this.FileName)) �� Sheet ($($this.SheetName)) ������܂���"
				return $false
			}

			# �w�薼�̂̃V�[�g���ɁA�w�薼�̂� Table���o�^����Ă��邩�m�F
			$this.inputWksheet = $this.inputWkbook.WorkSheets($sheetName)
			[string[]]$tableNameArr = @()
			$tableNameArr += $this.inputWksheet.ListObjects | ForEach-Object { $_.Name }
			if( $tableNameArr -notcontains $tableName ){
				errlog "���͌�Excel file ($this.FileName) �� Sheet ($sheetName) �� Table ($tableName) ������܂���"
				return $false
			}
		}
		catch {
			errlog "Excel file �̑���Ɏ��s"
			return $false
		}

		return $true
	}


	# Filter �Ɏg�� key�� Table���̗񂩂�擾
	[string[]] getFilterKey()
	{
		try {
			[object]$tblRange = $this.inputWksheet.ListObjects($this.TableName).DataBodyRange
			$left_top = $tblRange.item([XlInputTableRow]::CATEGORY).Address()
			$right_bottom = $tblRange.item([XlInputTableRow]::CATEGORY + ($tblRange.Rows.Count - 1) * $tblRange.Columns.Count).Address()
			$categoryList = $this.inputWksheet.Range($left_top, $right_bottom) | ForEach-Object { $_.Value() }
			$categoryUniqueList = $categoryList | Sort-Object -Unique
			
			# ��ӂ�key�z��(string[])
			return $categoryUniqueList
		}
		catch {
			errlog "Filter�Ɏg�� key�̎擾�Ɏ��s"
			return $null
		}
	}


	# ���͌��̎d�l���X�V����Table����A����key�ōi�荞�񂾌��ʂ̎d�l���i�[��URL���X�g���擾
	[XlInputTableData[]]readFilterData( [string]$filterkey )		# filterkey == category
	{

		Write-Host "$filterkey �ɑ΂��ď��X�V���܂��B"

		# �񂠂��蕡������(OR)�̏ꍇ�A@("����1", "����2") �Ŏw��\
		[int] $flValue = 7			# [Microsoft.Office.Interop.Excel.XlAutoFilterOperator]::xlFilterValues
		[void]$this.inputWksheet.ListObjects($this.TableName).Range.AutoFilter([XlInputTableRow]::CATEGORY, $filterkey, $flValue )

		# get count of visible rows
		# Auto fileter�K�p��́A�\����item�݂̂��擾
		[int] $spcellType = 12		# [Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeVisible
		$visibleRange = $this.inputWksheet.ListObjects($this.TableName).DataBodyRange.SpecialCells($spcellType)

		# --- �Q�l��� -------------------------------------
		# Auto fileter�K�p��̍s�����擾����
		# �I�𒆂̑Sitem�� �� �� �ōs���擾
		# (SpecialCells.Rows.Count�͋@�\���Ȃ��B�����̈�I����Ԃ̂��߁B�������ꂽ�ŏ��̗̈�̍s�����Ԃ������ۂ�)
		$itemCount = $visibleRange.CountLarge
		$colsCount = $visibleRange.Columns.CountLarge		# �s�P�ʂőI������邽�ߗ�͕��f����Ȃ��B�Ȃ̂� Count���g����
		$rowsCount = $itemCount / $colsCount

		debugOut -g1 "�i�荞�݌��� Cell�� = " -w2 "$itemCount"
		debugOut -g1 "�i�荞�݌��� ��   = " -w2 "$colsCount"
		debugOut -g1 "�i�荞�݌��� �s��   = " -w2 "$rowsCount"

		# �擪�s�ƍŏI�s�̍s�ԍ����擾����(Excel��ʂ̍��ɂ���s�ԍ�)
		$startRow = $visibleRange.item(1).Row
		# ��\���s���܂߂Čv�Z���Ă��܂����߁A�e�[�u���I�[�ɂ͓͂��Ȃ�
		$endRow = $visibleRange.End(-4121).Row
		debugOut -g1 "Table Start Row = " -w2 "$startRow"
		debugOut -g1 "Table End   Row = " -w2 "$endRow"
		# -------------------------------------------------


		# �i�荞�񂾌��ʂ���s���z��Ɋi�[(�Â����ɕ���ł���)
		[XlInputTableData[]]$outlist = @()
		[XlInputTableData]$oneLineData = [XlInputTableData]::new()

		$visibleRange.Rows | ForEach-Object {

			$oneLineData.Datetime	= $_.Cells([XlInputTableRow]::DATETIME).Value()
			$oneLineData.MailTitle	= $_.Cells([XlInputTableRow]::MAILTITLE).Value()
			$oneLineData.Category	= $_.Cells([XlInputTableRow]::CATEGORY).Value()
			$oneLineData.URL		= $_.Cells([XlInputTableRow]::URL).Value()
			$oneLineData.URLcheck	= $_.Cells([XlInputTableRow]::URLCHECK).Value()
			$oneLineData.MailID		= $_.Cells([XlInputTableRow]::MAILID).Value()
			$oneLineData.absoluteRowPosition = $_.Cells(0).Row + 1		# Table�J�n�s�̓w�b�_���̂��߁A+1���Ă���

			if( $oneLineData.URLcheck -ne [string][XlInputCheckURL]::checked ){
				debugOut -g1 "���擾��F" -w2 "$($oneLineData.URL)"
				$outlist += $oneLineData.Clone()	# �d�l���擾��URL�ꗗ�̍Ō����1���ǉ�
			}
			else {
				debugOut -y1 "�X�V�ςݏ��̂��߃X�L�b�v" -g2 " URL�F$($oneLineData.URL)"
			}
		}

		# Excel�� Table����z��Ƃ��ĕԂ�
		return $outlist
	}


	# ���͌� Excel sheet�� URL�`�F�b�N��񏑂�����
	[void]writeURLCheckStatus( [string]$strData, [XlInputTableData]$rowInfo )
	{
		[int]$tableTopCol = $this.inputWksheet.ListObjects($this.TableName).Range.item(0).Column
		$absoluteRow = $rowInfo.absoluteRowPosition -  1
		$absoluteCol = $tableTopCol + [XlInputTableRow]::URLCHECK

		$this.inputWksheet.cells( $absoluteRow, $absoluteCol ).Value = $strData
	}


	# Autofilter����������
	[void]RemoveAutofilter()
	{
		if( $null -eq $this.inputWksheet.AutoFilter ){ return }

		$this.inputWksheet.AutoFilter.ShowAllData()
	}
}
