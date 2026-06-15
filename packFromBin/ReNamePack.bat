@echo off
setlocal

set "PACK_DIR=%~dp0"
for %%I in ("%PACK_DIR%.") do set "PROJ_ROOT=%%~dpI"

set "SRC64=%PROJ_ROOT%NumDesTools\bin\Release\net9.0-windows\publish\NumDesTools-AddIn64-packed.xll"
set "XLL64=%PACK_DIR%NumDesToolsPack64.xll"

echo Source: %SRC64%

if exist "%SRC64%" (
    if exist "%XLL64%" del /f /q "%XLL64%"
    copy "%SRC64%" "%XLL64%"
    echo [64bit] %XLL64% generated.
) else (
    echo [64bit] Source not found.
)

endlocal
