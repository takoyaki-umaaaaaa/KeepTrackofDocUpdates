# Update the specification documents list.
# �d�l���ꗗ���X�V����
#
# ����F
# 1. �d�l���X�V�������܂Ƃ߂�Excel file�����ʂ��Ƃɗ���(�ۑ���URL)��ǂݏo��
# 2. SVN��URL���� file list ���擾����
# 3. �ŐV�d�l���ꗗ���܂Ƃ߂�Excel file�ɏ�������

# Requires -Version 5.0
Using module .\XlInputManager.psm1
Using module .\XlOutputManager.psm1
Import-Module -Name $PSScriptRoot\OutputMessage.psm1

# �ݒ�t�@�C���ǂݍ��� (Script�Ɠ����f�B���N�g��)
if( -not (Test-Path "$PSScriptRoot\Settings.json") ) { Write-Host -ForegroundColor Red "�ݒ�t�@�C��(Settings.json)������܂���"; exit -1}
$SettingsJson = (Get-Content "$PSScriptRoot\Settings.json" -Encoding UTF8 -Raw | ConvertFrom-Json)

# ���� Excel�t�@�C��
[string]$xlInputFileName = $SettingsJson.InputExcelFile
[string]$xlInputSheetName = $SettingsJson.InputExcelSheet		# �X�V������񂪋L�ڂ��ꂽ sheet��
[string]$xlInputTableName = $SettingsJson.InputExcelTable		# �X�V������񂪋L�ڂ��ꂽ Table��
# �o�� Excel�t�@�C��
[string]$xlOutputFileName = $SettingsJson.OutputExcelFile

if( -not (Test-Path -LiteralPath $xlInputFileName) ) { Write-Host -ForegroundColor Red "Settings.json�ɋL�ڂ̓ǂݍ��݃t�@�C��($xlInputFileName)��������܂���B�p�X���������Ă��������B"; exit -1}
if( -not (Test-Path -LiteralPath $xlOutputFileName) ){ Write-Host -ForegroundColor Red "Settings.json�ɋL�ڂ̏������݃t�@�C��($xlOutputFileName)��������܂���B�p�X���������Ă��������B"; exit -1}



# ���ݒ�
Set-StrictMode -Version 3.0
$ErrorActionPreference = "stop"						# �G���[�����������ꍇ�̓X�N���v�g�̎��s���~
$PSDefaultParameterValues['out-file:width'] = 2000	# Script���s����1�s������2000�����ݒ�
debugOut_setMode -outputStatus						# �f�o�b�O�o�͐ݒ�(�璷�o�̗͂L��)
errlog_setOutFile ($PSScriptRoot + "\KeepTrackofDocUpdates.log")   # Log �o�͐�

[XlInput]$excelInput = [XlInput]::new()
[XlOutput]$excelOutput = [XlOutput]::new()

# ���͌��� Excel ���N�����A�f�[�^�̓ǂݍ��ݏ���������(sheet, Table object�擾)
[boolean]$result = $excelInput.StartExcel( $xlInputTableName, $xlInputSheetName, $xlInputFileName )
if( -not $result ){ exit -1 }

# Auto filter�Ɏg�� key�z�����͌� Excel �� Table����擾
[string[]]$categoryKeyList = $excelInput.getFilterKey()
if( $null -eq $categoryKeyList ){
	errlog "$xlInputFileName ���� category �̎擾�Ɏ��s"
	exit -1
}


# �o�͐� Excel���N�����A�f�[�^�̏������ݏ���������(�V�Kcategory������ꍇ�� �V�Kcategory����sheet���쐬)
$result = $excelOutput.StartExcel( $categoryKeyList, $xlOutputFileName )
if( -not $result ){ exit -1 }


# category�P�ʂŁu�d�l���ꗗ���擾�v�A�u�o�͐� Excel�ւ̏������݁v���J��Ԃ�
foreach ($cate in $categoryKeyList) {
	try{
		# �d�l���X�V����Table���Acategory�ōi�荞��file�i�[��URL���X�g���擾
		[XlInputTableData[]]$listdata = @()
		$listdata = $excelInput.readFilterData( $cate )

		# �쐬����URL���X�g1�s���ƂɎd�l�����X�g���擾���A�ŐV�d�l���ꗗ�ɏ�������
		$listdata | ForEach-Object {
			Write-Host "`n------------"
			Write-Host "$($_.URL) ���������܂�"
			$excelInput.writeURLCheckStatus( [string][XlInputCheckURL]::checking, $_ ); $excelInput.Save()


			# ���͌� file ������肵�� URL�ɃA�N�Z�X���A�d�l�����X�g���擾
			[string[]]$filenameList = $excelOutput.getFilelistFromURL( $($_.URL) ) 

			if( ($filenameList | Measure-Object).count -gt 0 ){ # 0���̏ꍇ�� $filenameList�� $null �����邽�߁A���̂܂܂ł̓��\�b�h���g���Ȃ����߂̉����
				$excelInput.writeURLCheckStatus( [string][XlInputCheckURL]::checked, $_ ); $excelInput.Save()

				# �d�l�����X�g���o�͐� Excel file �ɏ�������
				$excelOutput.writeListToExcelFile( $filenameList, $cate )
			}
			else {
				$excelInput.writeURLCheckStatus( [string][XlInputCheckURL]::failed, $_ ); $excelInput.Save()

				# �G���[�ɂ�郊�X�g�擾���s��A���������t�@�C����1�����Ȃ��ꍇ�͏����𔲂��Ď���URL��
				errlog "file list �� 0���̂��ߏo�͏������X�L�b�v - $cate :�Ώ�URL :  $($_.URL)"
				# continue����� category�� foreach()�ɖ߂��Ă��܂�(ForEach-Object�͓r����~�ł��Ȃ�)�̂� else�߂ŏ������X�L�b�v
			}
		}

		# �������
		$listdata.Clear()
	}
	catch {
		$Error[0] | Select-Object -Property * | errlog
	}
}


$excelInput.RemoveAutofilter()		# Auto filter����(�����\�ȏꍇ�̂�)
$excelInput.Save()
$excelInput.Quit()
$excelOutput.Save()
$excelOutput.Quit()

exit 0
