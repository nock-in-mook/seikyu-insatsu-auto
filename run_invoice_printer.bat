@echo off
chcp 65001 >nul
set PYTHONUTF8=1

set LOGFILE=%~dp0run_log.txt
echo ===== %DATE% %TIME% ===== >> "%LOGFILE%"
echo CWD: %CD% >> "%LOGFILE%"
echo PATH: %PATH% >> "%LOGFILE%"

:: Google Drive マウント確認（最大60秒待機）
set RETRY=0
:WAIT_GDRIVE
if exist "G:\マイドライブ\_claude-sync\shared-env" goto GDRIVE_OK
set /a RETRY+=1
if %RETRY% GTR 12 (
    echo ERROR: Google Drive not mounted after 60s >> "%LOGFILE%"
    exit /b 1
)
echo Waiting for Google Drive... %RETRY% >> "%LOGFILE%"
timeout /t 5 /nobreak >nul
goto WAIT_GDRIVE
:GDRIVE_OK

echo Google Drive OK >> "%LOGFILE%"
py -3.14 "%~dp0load_env_and_run.py" %* >> "%LOGFILE%" 2>&1
set EXITCODE=%ERRORLEVEL%
echo EXIT CODE: %EXITCODE% >> "%LOGFILE%"
exit /b %EXITCODE%
