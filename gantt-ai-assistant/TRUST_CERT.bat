@echo off
echo ============================================================
echo  STEP 1 of 2: Trust the localhost certificate
echo  (This requires Administrator privileges)
echo ============================================================
echo.

set "CERT=c:\Users\hp\Desktop\aaaadaatassets\project\HEART-FAILURE-DEATH-PREDICTOR\gantt-ai-assistant\certs\localhost.crt"

echo Importing certificate into Windows Trusted Root store...
powershell -Command "Start-Process powershell -Verb RunAs -ArgumentList 'Import-Certificate -FilePath \"%CERT%\" -CertStoreLocation Cert:\LocalMachine\Root; Write-Host Certificate trusted successfully! -ForegroundColor Green; pause'"

echo.
echo Done! Now double-click START_ALL.bat to launch the app.
pause
