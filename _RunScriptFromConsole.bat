powershell -ExecutionPolicy RemoteSigned -File %1


rem Scriptの実行を許可する(管理者権限が必要)
rem powershell -ExecutionPolicy RemoteSigned

rem Current userのみ、scriptの実行を許可する(管理者権限不要)
rem Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
