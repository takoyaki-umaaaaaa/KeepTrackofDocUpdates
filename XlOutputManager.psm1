# �o�͐� Excel file�̑��� class
# Requires -Version 5.0

class XlOutput
{
	# Property��`
	hidden [string]$FileName = ""
	hidden [object]$xlInstance = $null
	hidden [object]$outputWkbook = $null
	hidden [string]$URLstr = $null			# file list �擾����URL�B$SpecHistURLarray�z��̒���1�v�f�B���ݏ�������URL�B
	hidden [string[]]$fileList = @()		# $URLstr����擾���� file�ꗗ
	hidden [string]$CategoryStr = $null		# category���Ƃ� output workbook �� sheet�ɕ�������Bcategory�� XlInput�� Auto filter ���� key

	# constractor
	XlOutput()
	{
	}


	# Excel book��ۑ�
	[void] save ()
	{
		if( $null -eq $this.outputWkbook )
		{
			errlog "Excel book���J���O�ɕۑ����悤�Ƃ���"
			return
		}

		$this.outputWkbook.Save()
	}


	# Excel book�̏I��
	[void] Quit()
	{
		if( $null -eq $this.outputWkbook ){ return }

		[System.Runtime.InteropServices.Marshal]::ReleaseComObject( $this.outputWkbook )
		$this.outputWkbook = $null

		$this.xlInstance.Quit()
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject( $this.xlInstance )
		$this.xlInstance = $null
	}


	# �o�͐� Excel file ���J���A�o�͐� sheet�̑��݂��m�F����Bsheet���Ȃ���΍쐬����B
	[boolean] StartExcel ([string[]]$categoryList, [string]$fileName )
	{
		try {
			# �t�@�C�������o���Ă���
			$this.FileName = $fileName

			# �o�͐� Excel�N��
			$this.xlInstance = New-Object -comobject Excel.Application
			$this.xlInstance.Visible = $false
			$this.xlInstance.DisplayAlerts = $false
			$this.outputWkbook = $this.xlInstance.Workbooks.Open($fileName)
		}
		catch {
			errlog "Excel file ($fileName) ���J���܂���"
			return $false
		}


		try {
			# �o�͐� Excel workbook�� �ucategory�v���� sheet���Ȃ���΍쐬����
			$empty_Var = [System.Type]::Missing
			[object]$outSheets = $this.outputWkbook.WorkSheets
			[string[]]$outSheetNameArray = @()
			[string[]]$outSheetNameArray += $outSheets | ForEach-Object { $_.Name }
			$categoryList | ForEach-Object {
				if( $outSheetNameArray -notcontains $_ ){  # -ne�̍������z��̏ꍇ�A�S�v�f���`�F�b�N����
					Write-Host "�o�͐�t�@�C���ɃV�[�g[$_]���Ȃ����߁A�쐬���܂�"
					$xlsheet = $outSheets.Add($empty_Var, $outSheets.item($outSheets.Count))
					$xlsheet.Name = $_
				}
			}
		}
		catch {
			errlog "Excel file �̑���Ɏ��s"
			return $false
		}

		return $true
	}



	# ����URL�ɃA�N�Z�X���A�d�l���t�@�C�����X�g���擾
	[string[]] getFilelistFromURL ( [string]$url )
	{
		$this.URLstr = $url			# �o�̓t�@�C���ɏ������ނ��߁Aglobal scope�ϐ��ɏ�����URL������(�g���͎̂q�֐��ȉ��Ȃ̂�$url���Q�Ƃł��邯��)
		$this.fileList = @()

		try {
			# svn server�Ƀt�@�C���ꗗ��v��
			# �� ����svn�� tortoiseSVN�̃R�}���h���C���ł�z��
			# �@ �ʏ�� svn �œ�������ɂȂ邩�͖��m�F
			$this.fileList = Write-Output "R" | svn list $this.URLstr
		}
		catch {
			# tortoiseSVN Ver 1.14.1 �ł̓R�}���h���s��Exception��ݒ肵�Ă���̂ł����ɗ���
			errlog "svn�ł� list�擾���s - �Ώ�URL�F$($this.URLstr)"
			$Error[0] | Select-Object -Property * | errlog
		}

		$this.fileList = Get-Content "$PSScriptRoot\resource\test.txt"	# ����m�F�p(�d�l�����X�g�p)

		return $this.fileList
	}


	# ����filelist�����A�o�͐� Excel file �ɏ�������
	[void] writeListToExcelFile( [string[]]$strList, [string]$cate )
	{
		$this.CategoryStr = $cate		# �������� category���o���Ă���

		$strList | ForEach-Object {
			Write-Host "`nFile���F$_ ���������܂�"
			$this.writeStringToExcelFile( $_, $cate )
		}
	}


	# �����̕�������o�͐�Excel�V�[�g�ɏ�������
	hidden [void] writeStringToExcelFile ( [string]$outName, [string]$cate )
	{
		# �o�͐�� Table�����쐬 (Table���� [�J�e�S����_�d�l����] �Ƃ��Ă���)
		# �܂��̓t�@�C��������u�d�l�����v��؂�o��
		if( $outName -match '(Ver|Rev)' ){
			$specDocName = $outName -replace '(_*Ver|_*Rev).*$', $null		# "Ver", "Rev" �ȍ~�̕�������폜���A�d�l���ʂɋ敪�ł���`�ɂ���
		}
		else {
			$specDocName = [System.IO.Path]::GetFileNameWithoutExtension( $outName )
		}
		$TableNameToEdit = $this.CategoryStr + "_" + $specDocName		# �������̃J�e�S���� + �d�l���� �� �ړI�̏������ݐ�Table�����쐬
		$TableNameToEdit = $TableNameToEdit -replace '[- ]', '_'		# Table���ɋ󔒁A�n�C�t���͎g���Ȃ��悤�Ȃ̂�
		debugOut -g1 "�e�[�u�����F[" -w1 "$TableNameToEdit" -g2 "] ��T���܂�"

		# �X�V��Table��T��
		$sheet = $this.outputWkbook.WorkSheets($this.CategoryStr)
		$searchTableResult = $false
		foreach( $listObj in $sheet.ListObjects ){
			debugOut -g1 "��Table���F[" -w1 "$($listObj.DisplayName())" -g2 "]"

			if ( $TableNameToEdit -eq $listObj.DisplayName() ) {
				Write-Host "�X�V�Ώ�table��������܂���($TableNameToEdit)"
				$searchTableResult = $true
				break		# ���������̂Ŕ�����
			}
		}

		if ( $searchTableResult ) {
			debugOut -w1 "$TableNameToEdit" -g2 "Table�ɂP�s�ǉ����܂�"

			# Table �̍Ō�� cell�ʒu��T��
			$outRangeTable = $sheet.ListObjects($TableNameToEdit).Range
			$startRow = $outRangeTable.item(1).Row + $outRangeTable.Rows.Count
			$startCol = $outRangeTable.item(1).Column

			# Table��1�s�ǉ�
			$this.addRowToExcelTbl( $outName, $outRangeTable, $startRow, $startCol, [ref]$sheet )
		}
		else {
			Write-Host "�X�V�Ώۃe�[�u����������܂���B�V�K�쐬���܂��B"

			# �󂢂Ă���ӏ��� Table��V�K�쐬
			[object]$tableListObj = $this.createExcelTbl( $TableNameToEdit, [ref]$sheet )

			$row = $tableListObj.Range.item(1).Row + 1
			$col = $tableListObj.Range.item(1).Column
			debugOut -g1 "�V�K�쐬�e�[�u���ɏ������� �s�F$row, ��F$col"

			# �쐬���� Table ��1����������
			$this.addRowToExcelTbl( $outName, $tableListObj.Range, $row, $col, [ref]$sheet )
		}
	}


	# Excel sheet�Ɏd�l���ꗗTable���쐬
	hidden [object]createExcelTbl( [string]$tableName, [ref]$wksheet )
	{

		[void]$wksheet.value.Activate()  # ���܂��Ȃ�(��)

		# Table���쐬����͈͂����߂�
		[int]$newTblRow = 0
		[int]$newTblCol = 0
		[int]$newCells = $wksheet.value.UsedRange.Cells.Count
		if( $newCells -le 1 ){
			# Table���S���Ȃ� sheet�̏ꍇ�͏����ʒu�͌��ߑł�����
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

		# Table�̃w�b�_���������������
		$wksheet.value.cells($newTblRow, $newTblCol + 0).Value = "�X�V����"
		$wksheet.value.cells($newTblRow, $newTblCol + 1).Value = "�d�l����"
		$wksheet.value.cells($newTblRow, $newTblCol + 2).Value = "Ver"
		$wksheet.value.cells($newTblRow, $newTblCol + 3).Value = "URL"

		$wksheet.value.cells($newTblRow, $newTblCol + 0).ColumnWidth = 12.46
		$wksheet.value.cells($newTblRow, $newTblCol + 1).ColumnWidth = 24.88
		$wksheet.value.cells($newTblRow, $newTblCol + 2).ColumnWidth = 8.04
		$wksheet.value.cells($newTblRow, $newTblCol + 3).ColumnWidth = 32.33

		# Table���쐬
		$xlTableValue = 1	# [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange
		$guessYes = 1		# [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
		$outListObjTable = $wksheet.value.ListObjects.Add( $xlTableValue, # �ǉ��Ώۂ̎��(�͈� == Table)
			$newTblRange,   # Table�ǉ��Ȃ�ATable�ɂ���̈�
			$false,		 # �O���f�[�^�\�[�X�ƃ����N���邩
			$guessYes )	 # �̈�Ƀw�b�_���܂�ł��邩
		$outListObjTable.Name = $tableName
		$outListObjTable.TableStyle = "TableStyleMedium2"

		Write-Host "�V�[�g���F$($wksheet.value.Name) �ɐV�KTable [$tableName] ��ǉ����܂����B"
		return $outListObjTable	# �o��
	}


	# Excel sheet�̎d�l���ꗗTable��1�s�ǉ�
	[void] addRowToExcelTbl( [string]$documentName, [object]$tableRange, [object]$row, [object]$col, [ref]$wksheet )
	{
		# �t�@�C�������� Ver���擾
		if( $documentName -match '(Ver|Rev)' ){
			[string]$ver = $documentName -replace '^.*?(?=(ver|rev))', $null
			if( $ver -ne $null ){
				[int]$extIdx = $ver.LastIndexOf('.')  # ������Ȃ������� -1
				if ( $extIdx -gt 0 ) {
					# idx 0 �ł� match �͑ΏۊO�ɂ�����
					$ver = $ver.Substring( 0, $extIdx )
				}
			}
		}
		else {
			$ver = ""		# �t�@�C������ Version�������ꍇ�AVersion�ɂ��Ă͉����o�͂��Ȃ�
		}

		debugOut -g1 "1���ǉ��F�s=$row, ��=$col, ����=$documentName, Ver=$ver"

		# Table�͈̔͒����ɏ������ނƁA�����Table�̈���g�����Ă����
		# ��͍����� �X�V����, �t�@�C����, Ver, URL
		$wksheet.value.cells($row, $col + 0).Value = Get-Date -Format "yyyy/MM/dd"
		$wksheet.value.cells($row, $col + 1).Value = $documentName
		$wksheet.value.cells($row, $col + 2).Value = $ver
		$wksheet.value.cells($row, $col + 3).Value = $this.URLstr

		$empty_Var = [System.Type]::Missing  # �Ăяo���������default�l�ɂȂ�Bhttps://docs.microsoft.com/ja-jp/dotnet/api/system.type.missing?view=net-5.0
		[void]$tableRange.Sort(
			$tableRange.item(1), # Sort key 1
			2,		  # �~��
			$empty_Var, # Sort key 2
			$empty_Var, # �s�{�b�g�e�[�u���̏ꍇ�ɂ̂ݎw�肷��ӏ�
			$empty_Var, # Key 2 �̕��בւ�����
			$empty_Var, # Sort key 3
			$empty_Var, # Key 3 �̕��בւ�����
			1 )		 # �擪�s�Ƀw�b�_�[���܂�

		Write-Host "�V�[�g���F$($wksheet.value.Name) Table�� $documentName ��ǉ����܂����B"
	}
}
