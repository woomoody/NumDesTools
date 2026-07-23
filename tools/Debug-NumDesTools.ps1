#!/usr/bin/env pwsh
#Requires -Version 7.2
<#
.SYNOPSIS
    一键 Debug 启动 NumDesTools Excel-DNA 插件

.DESCRIPTION
    1. Build Debug 配置
    2. 启动 Excel 并加载生成的 XLL
    3. 输出 Excel PID，方便手动附加调试器

.NOTES
    Windows Terminal Command Palette 调用示例：
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools\tools\Debug-NumDesTools.ps1"
#>

[CmdletBinding()]
param(
    [string]$Configuration = "Debug",
    [switch]$AttachVsCode,
    [switch]$NoBuild
)

$ErrorActionPreference = "Stop"

# 项目路径
$RepoRoot = Split-Path -Parent $PSScriptRoot
$Project = Join-Path $RepoRoot "NumDesTools\NumDesTools.csproj"
$PackScript = Join-Path $RepoRoot "packFromBin\ReNamePack.bat"
$ExcelPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"

# 兼容其他 Office 安装位置
if (-not (Test-Path $ExcelPath)) {
    $ExcelPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
}
if (-not (Test-Path $ExcelPath)) {
    Write-Error "找不到 EXCEL.EXE，请检查 Office 安装路径"
}

# 构建
if (-not $NoBuild) {
    Write-Host "[1/3] Building $Configuration ..." -ForegroundColor Cyan
    dotnet build `"$Project`" -c $Configuration
    if ($LASTEXITCODE -ne 0) { throw "Build failed" }

    # Release 模式下需要重命名打包
    if ($Configuration -eq "Release") {
        & $PackScript
        if ($LASTEXITCODE -ne 0) { throw "Pack script failed" }
    }
}

# 查找 XLL
$XllDir = Join-Path $RepoRoot "NumDesTools\bin\$Configuration\net9.0-windows"
$XllPath = Get-ChildItem -Path $XllDir -Filter "NumDesTools-AddIn64*.xll" | Select-Object -First 1 -ExpandProperty FullName

if (-not $XllPath) {
    throw "找不到 XLL 文件，请确认已 Build 成功: $XllDir"
}

Write-Host "[2/3] XLL: $XllPath" -ForegroundColor Cyan

# 启动 Excel 并加载 XLL
Write-Host "[3/3] Launching Excel with XLL ..." -ForegroundColor Cyan
$proc = Start-Process -FilePath $ExcelPath -ArgumentList `"$XllPath`" -PassThru

Write-Host "Excel PID: $($proc.Id)" -ForegroundColor Green
Write-Host "Excel Path: $ExcelPath" -ForegroundColor Green

# 可选：打开 VS Code 并提示附加
if ($AttachVsCode) {
    Start-Process -FilePath "code" -ArgumentList `"$RepoRoot`" -WindowStyle Hidden
    Write-Host "VS Code 已打开。按 F5 或 Ctrl+Shift+D -> 'Run and Debug' -> '.NET Core Launch (console)'" -ForegroundColor Yellow
}

Write-Host "`n提示：可在 VS Code 中按 Ctrl+Shift+D 选择 '.NET Core Launch (console)' 附加调试。" -ForegroundColor Gray
