@echo off
chcp 65001 >nul
echo.
echo  Syncing remote Claude projects via scp...
echo.

set REMOTE_USER=admin
set REMOTE_HOST=100.96.48.30
set REMOTE_PATH=/c/Users/admin/.claude/projects
set LOCAL_TEMP=C:\Users\cent\AppData\Local\Temp\claude_remote_projects

if not exist "%LOCAL_TEMP%" mkdir "%LOCAL_TEMP%"

scp -r -o StrictHostKeyChecking=no %REMOTE_USER%@%REMOTE_HOST%:%REMOTE_PATH%/. "%LOCAL_TEMP%"
if errorlevel 1 (
    echo  [warn] 远程同步失败，仅统计本地数据
)

echo.
echo  Running Claude Token Stats...
echo.
python "%~dp0claude_token_stats.py"
echo.
pause
