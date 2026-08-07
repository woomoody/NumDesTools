@echo off
chcp 65001 >nul
echo.
echo  Running Claude Token Stats...
echo.
set DATE_ARG=%1
if "%DATE_ARG%"=="" set DATE_ARG=today
python "%~dp0claude_token_stats.py" --date "%DATE_ARG%" %2
echo.
echo  Syncing history snapshot to git...
pushd "%~dp0"
git add token_stats_history.json 2>nul
git commit -m "chore(token-stats): update history snapshot %DATE_ARG%" 2>nul
git push < nul 2>nul
popd
echo  done (commit/push skipped if no changes)
echo.
pause
