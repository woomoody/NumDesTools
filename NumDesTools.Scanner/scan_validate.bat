@echo off
chcp 65001 > nul
cd /d "%~dp0"

:: 用法:
::   scan_validate.bat                          全量校验所有表
::   scan_validate.bat --pick                   列出最近活动，交互式选择后单活动校验
::   scan_validate.bat --pick --md              交互式选择 + 输出 MD 报告
::   scan_validate.bat --md                     全量校验 + 输出 MD 报告（默认路径 Config 目录）
::   scan_validate.bat --md D:\report.md        全量校验 + 输出到指定路径
::   scan_validate.bat --errors-only            只显示错误（不显示警告）
::   scan_validate.bat --activity 74005         直接指定 activityID 校验
::   scan_validate.bat --activity 74005 --md    单活动校验 + 输出 MD

echo [配置自检] 开始...

:: 双击直接运行时默认全量校验 + 生成并打开 MD 报告
:: 命令行带参数时使用传入参数（如 --pick --md / --activity 74005）
if "%~1"=="" (
    dotnet run --project NumDesTools.Scanner.csproj -- --validate --md
) else (
    dotnet run --project NumDesTools.Scanner.csproj -- --validate %*
)

if exist "C:\Users\cent\Documents\NumDesTools\Config\validate_latest.md" (
    start "" "C:\Users\cent\Documents\NumDesTools\Config\validate_latest.md"
)
pause
