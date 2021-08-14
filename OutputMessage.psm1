[switch]$script:debugOutput = $true	  # デバッグ出力の有効/無効(debugOut_setModeで指定)
[string]$script:LogFilename = ""				# Log 出力先(errlog_setOutFileで指定)



function debugOut_setMode( [switch]$outputStatus ) {
	$script:debugOutput = $outputStatus
}

function debugOut ($g1 = "", $w1 = "", $y1 = "", $dy1 = "", $r1 = "", $dr1 = "", $g2 = "", $w2 = "", $y2 = "", $dy2 = "", $r2 = "", $dr2 = "") {
	# Write-Hostの情報
	# https://docs.microsoft.com/ja-jp/powershell/module/microsoft.powershell.utility/write-host

	if ( $script:debugOutput -eq $false ) {
		return
	}
	Write-Host -ForegroundColor Gray $g1 -NoNewline; `
	Write-Host -ForegroundColor White $w1 -NoNewline; `
	Write-Host -ForegroundColor Yellow $y1 -NoNewline; `
	Write-Host -ForegroundColor DarkYellow $dy1 -NoNewline; `
	Write-Host -ForegroundColor Red $r1 -NoNewline; `
	Write-Host -ForegroundColor DarkRed $dr1 -NoNewline; `
	Write-Host -ForegroundColor Gray $g2 -NoNewline; `
	Write-Host -ForegroundColor White $w2 -NoNewline; `
	Write-Host -ForegroundColor Yellow $y2 -NoNewline; `
	Write-Host -ForegroundColor DarkYellow $dy2 -NoNewline; `
	Write-Host -ForegroundColor Red $r2 -NoNewline; `
	Write-Host -ForegroundColor DarkRed $dr2
}


function errlog_setOutFile( [string]$fullpath ) {
	$script:LogFilename = ""

	# 存在するフォルダのみ受け付ける
	$folder = [System.IO.Path]::GetDirectoryName( $fullpath )
	if( -not (Test-Path $folder) ){
		Write-Host -ForegroundColor Yellow "ログ出力先フォルダが存在しません"
		return
	}

	$script:LogFilename = $fullpath
	
	[string]$ymdhms = Get-Date -Format "yyyy/MM/dd HH:mm"
	Out-File -InputObject "`n`n-----------------------------" -Encoding oem -Append -FilePath $script:LogFilename
	Out-File -InputObject "$ymdhms : Start logging" -Encoding oem -Append -FilePath $script:LogFilename
}

function errlog {
	param (
		[Parameter( ValueFromPipeline = $true, Mandatory = $true )]   # pipeでも使える, 必須param
		[string]$strErr
	)
	process {
		if( [string]::IsNullOrEmpty($strErr) ){
			Write-Host -ForegroundColor Yellow "ログの出力先が設定されていません"
			return
		}
		
		[string]$scriptname = [System.IO.Path]::GetFileName( $MyInvocation.ScriptName )
		Write-Host -ForegroundColor Red $strErr
		Out-File -InputObject "$scriptname : Line $($MyInvocation.ScriptLineNumber) : $strErr" -Encoding oem -Append -FilePath $script:LogFilename
	}
}
