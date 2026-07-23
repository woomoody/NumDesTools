#!/usr/bin/env pwsh
#Requires -Version 7.2
<#
.SYNOPSIS
    在 Windows Terminal 命令面板注册 "Debug NumDesTools"
#>

$ErrorActionPreference = "Stop"

$ActionDebug = [PSCustomObject]@{
    command = [PSCustomObject]@{
        action       = "newTab"
        commandline  = 'pwsh -NoProfile -ExecutionPolicy Bypass -File "C:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools\tools\Debug-NumDesTools.ps1"'
        startingDirectory = "C:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools"
    }
    name = "Debug NumDesTools"
    icon = "🚀"
}

$ActionRelease = [PSCustomObject]@{
    command = [PSCustomObject]@{
        action       = "newTab"
        commandline  = 'pwsh -NoProfile -ExecutionPolicy Bypass -File "C:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools\tools\Debug-NumDesTools.ps1" -Configuration Release'
        startingDirectory = "C:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools"
    }
    name = "Release NumDesTools"
    icon = "📦"
}

# 可能的 settings.json 路径
$paths = @(
    "$env:LOCALAPPDATA\Microsoft\Windows Terminal\settings.json",
    "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json",
    "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\settings.json"
)

$settingsPath = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $settingsPath) {
    Write-Error "找不到 Windows Terminal 的 settings.json。手动检查这些路径：`n$($paths -join "`n")"
}

Write-Host "找到: $settingsPath" -ForegroundColor Cyan

$json = Get-Content $settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json

# 确保 actions 数组存在
if (-not $json.actions) {
    $json | Add-Member -MemberType NoteProperty -Name actions -Value @()
}

# 移除已存在的同名 action，再添加新的
foreach ($act in @($ActionDebug, $ActionRelease)) {
    $existing = $json.actions | Where-Object { $_.name -eq $act.name }
    if ($existing) {
        $json.actions = $json.actions | Where-Object { $_.name -ne $act.name }
    }
    $json.actions += $act
}

# 保存
$json | ConvertTo-Json -Depth 10 | Set-Content $settingsPath -Encoding UTF8
Write-Host "已注册。保存后完全关闭并重启 Windows Terminal，然后 Ctrl+Shift+P 搜 'Debug NumDesTools'" -ForegroundColor Green
