@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo [缺陷扫描] 开始...
dotnet run --project NumDesTools.Scanner.csproj -- --bugs --release
pause
