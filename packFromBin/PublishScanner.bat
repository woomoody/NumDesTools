@echo off
setlocal
set "PACK_DIR=%~dp0"
set "PROJ=%PACK_DIR%..\NumDesTools.Scanner\NumDesTools.Scanner.csproj"
set "OUT=%PACK_DIR%scanner_publish\"

echo [Scanner] 清理旧包...
if exist "%OUT%" rd /s /q "%OUT%"
mkdir "%OUT%"

echo [Scanner] 打单文件独立包（win-x64）...
dotnet publish "%PROJ%" -c Release -r win-x64 ^
  /p:PublishSingleFile=true ^
  /p:SelfContained=true ^
  /p:IncludeNativeLibrariesForSelfExtract=true ^
  /p:PublishReadyToRun=false ^
  -o "%OUT%"

if %ERRORLEVEL% neq 0 (
    echo [Scanner] 打包失败！
    pause
    exit /b 1
)

echo [Scanner] 完成：%OUT%NumDesTools.Scanner.exe
pause
endlocal
