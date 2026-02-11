@echo off
setlocal

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$port=4173; $cons=Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue; if($cons){$pids=$cons|Select-Object -ExpandProperty OwningProcess -Unique; foreach($procId in $pids){try{Stop-Process -Id $procId -Force -ErrorAction Stop; Write-Host ('Stopped PID ' + $procId)}catch{}}} else {Write-Host 'GUI server is not running.'}"

endlocal
exit /b 0
