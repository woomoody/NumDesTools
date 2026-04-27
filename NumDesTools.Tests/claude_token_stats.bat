@echo off
chcp 65001 >nul
echo.
echo  Running Claude Token Stats...
echo.
python "%~dp0claude_token_stats.py"
echo.
pause
