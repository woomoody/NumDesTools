@echo off
chcp 65001 >nul
echo.
echo  Running Claude Token Stats...
echo.
set DATE_ARG=%1
if "%DATE_ARG%"=="" set DATE_ARG=today
python "%~dp0claude_token_stats.py" --date "%DATE_ARG%"
echo.
pause
