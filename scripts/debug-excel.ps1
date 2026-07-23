# NumDesTools Debug Launcher
# 用法: keyr F5  或  opencode command: debug-excel
param([switch]$Release)

$ErrorActionPreference = "Stop"
$root = "C:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools"
$csproj = "$root\NumDesTools\NumDesTools.csproj"
$config = if ($Release) { "Release" } else { "Debug" }
$packScript = "$root\packFromBin\ReNamePack.bat"
$excel = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$xll = "$root\NumDesTools\bin\$config\net9.0-windows\NumDesTools-AddIn64.xll"

Write-Host "[1/3] dotnet build $config ..." -ForegroundColor Cyan
dotnet build $csproj -c $config
if ($LASTEXITCODE -ne 0) { throw "构建失败" }

Write-Host "[2/3] pack XLL ..." -ForegroundColor Cyan
& $packScript

Write-Host "[3/3] 启动 Excel + XLL ..." -ForegroundColor Cyan
Write-Host "  XLL: $xll" -ForegroundColor DarkGray
Write-Host "  如需断点调试，在代码中加 Debugger.Launch() 后重新 build" -ForegroundColor Yellow

Start-Process -FilePath $excel -ArgumentList "`"$xll`""
Write-Host "Excel 已启动。Attach debugger: VS -> 调试 -> 附加到进程 -> EXCEL.EXE" -ForegroundColor Green