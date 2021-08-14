# Update the specification documents list.
# 仕様書一覧を更新する
#
# 動作：
# 1. 仕様書更新履歴をまとめたExcel fileから種別ごとに履歴(保存先URL)を読み出す
# 2. SVNでURLから file list を取得する
# 3. 最新仕様書一覧をまとめたExcel fileに書き込む

# Requires -Version 5.0
Using module .\XlInputManager.psm1
Using module .\XlOutputManager.psm1
Import-Module -Name $PSScriptRoot\OutputMessage.psm1

# 設定ファイル読み込み (Scriptと同じディレクトリ)
if( -not (Test-Path "$PSScriptRoot\Settings.json") ) { Write-Host -ForegroundColor Red "設定ファイル(Settings.json)がありません"; exit -1}
$SettingsJson = (Get-Content "$PSScriptRoot\Settings.json" -Encoding UTF8 -Raw | ConvertFrom-Json)

# 入力 Excelファイル
[string]$xlInputFileName = $SettingsJson.InputExcelFile
[string]$xlInputSheetName = $SettingsJson.InputExcelSheet		# 更新履歴情報が記載された sheet名
[string]$xlInputTableName = $SettingsJson.InputExcelTable		# 更新履歴情報が記載された Table名
# 出力 Excelファイル
[string]$xlOutputFileName = $SettingsJson.OutputExcelFile

if( -not (Test-Path -LiteralPath $xlInputFileName) ) { Write-Host -ForegroundColor Red "Settings.jsonに記載の読み込みファイル($xlInputFileName)が見つかりません。パスを見直してください。"; exit -1}
if( -not (Test-Path -LiteralPath $xlOutputFileName) ){ Write-Host -ForegroundColor Red "Settings.jsonに記載の書き込みファイル($xlOutputFileName)が見つかりません。パスを見直してください。"; exit -1}



# 環境設定
Set-StrictMode -Version 3.0
$ErrorActionPreference = "stop"						# エラーが発生した場合はスクリプトの実行を停止
$PSDefaultParameterValues['out-file:width'] = 2000	# Script実行中は1行あたり2000文字設定
debugOut_setMode -outputStatus						# デバッグ出力設定(冗長出力の有無)
errlog_setOutFile ($PSScriptRoot + "\KeepTrackofDocUpdates.log")   # Log 出力先

[XlInput]$excelInput = [XlInput]::new()
[XlOutput]$excelOutput = [XlOutput]::new()

# 入力元の Excel を起動し、データの読み込み準備をする(sheet, Table object取得)
[boolean]$result = $excelInput.StartExcel( $xlInputTableName, $xlInputSheetName, $xlInputFileName )
if( -not $result ){ exit -1 }

# Auto filterに使う key配列を入力元 Excel の Tableから取得
[string[]]$categoryKeyList = $excelInput.getFilterKey()
if( $null -eq $categoryKeyList ){
	errlog "$xlInputFileName から category の取得に失敗"
	exit -1
}


# 出力先 Excelを起動し、データの書き込み準備をする(新規categoryがある場合は 新規category名のsheetを作成)
$result = $excelOutput.StartExcel( $categoryKeyList, $xlOutputFileName )
if( -not $result ){ exit -1 }


# category単位で「仕様書一覧を取得」、「出力先 Excelへの書き込み」を繰り返す
foreach ($cate in $categoryKeyList) {
	try{
		# 仕様書更新履歴Tableより、categoryで絞り込んだfile格納先URLリストを取得
		[XlInputTableData[]]$listdata = @()
		$listdata = $excelInput.readFilterData( $cate )

		# 作成したURLリスト1行ごとに仕様書リストを取得し、最新仕様書一覧に書き込み
		$listdata | ForEach-Object {
			Write-Host "`n------------"
			Write-Host "$($_.URL) を処理します"
			$excelInput.writeURLCheckStatus( [string][XlInputCheckURL]::checking, $_ ); $excelInput.Save()


			# 入力元 file から入手した URLにアクセスし、仕様書リストを取得
			[string[]]$filenameList = $excelOutput.getFilelistFromURL( $($_.URL) ) 

			if( ($filenameList | Measure-Object).count -gt 0 ){ # 0件の場合は $filenameListに $null が入るため、そのままではメソッドが使えないための回避策
				$excelInput.writeURLCheckStatus( [string][XlInputCheckURL]::checked, $_ ); $excelInput.Save()

				# 仕様書リストを出力先 Excel file に書き込み
				$excelOutput.writeListToExcelFile( $filenameList, $cate )
			}
			else {
				$excelInput.writeURLCheckStatus( [string][XlInputCheckURL]::failed, $_ ); $excelInput.Save()

				# エラーによるリスト取得失敗や、そもそもファイルが1件もない場合は処理を抜けて次のURLへ
				errlog "file list が 0件のため出力処理をスキップ - $cate :対象URL :  $($_.URL)"
				# continueすると categoryの foreach()に戻ってしまう(ForEach-Objectは途中停止できない)ので else節で処理をスキップ
			}
		}

		# 解放処理
		$listdata.Clear()
	}
	catch {
		$Error[0] | Select-Object -Property * | errlog
	}
}


$excelInput.RemoveAutofilter()		# Auto filter解除(解除可能な場合のみ)
$excelInput.Save()
$excelInput.Quit()
$excelOutput.Save()
$excelOutput.Quit()

exit 0
