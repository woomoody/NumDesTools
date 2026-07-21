<#
.SYNOPSIS
    把 NumDesTools xlsx 冲突解决工具装成 git 的 merge driver（*.xlsx/*.xlsm，见仓库根 .gitattributes）。
    装完之后，任何 GUI（SmartGit/TortoiseGit/GitExtensions/…）或纯命令行只要走 git merge/rebase，
    xlsx 冲突都会自动弹出这个工具——不需要针对每个 GUI 单独配置外部工具。

.DESCRIPTION
    1. Release build Scanner + Rust conflict-tui
    2. 拷到 %LOCALAPPDATA%\NumDesTools\ConflictTool\（工程目录之外，克隆/删仓库不影响这份安装）
    3. git config --global 注册 merge.numdes.driver 指向这份安装

    只需要在每台要用这个工具的机器上跑一次（仓库里的 .gitattributes 已经声明 xlsx 用 numdes driver，
    但 driver 具体指向哪个本机路径没法提交，必须本机注册）。
#>

$ErrorActionPreference = "Stop"
$repoRoot = (Resolve-Path "$PSScriptRoot\..").Path
$installDir = Join-Path $env:LOCALAPPDATA "NumDesTools\ConflictTool"

Write-Host "== 1/3 Release build NumDesTools.Scanner ==" -ForegroundColor Cyan
dotnet build "$repoRoot\NumDesTools.Scanner\NumDesTools.Scanner.csproj" -c Release
if ($LASTEXITCODE -ne 0) { throw "Scanner build 失败" }

Write-Host "== 2/3 拷贝到 $installDir ==" -ForegroundColor Cyan
$buildOut = "$repoRoot\NumDesTools.Scanner\bin\Release\net9.0-windows"
if (-not (Test-Path $buildOut)) { throw "找不到 build 产物：$buildOut" }
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
Copy-Item "$buildOut\*" $installDir -Recurse -Force

$rustExe = "$repoRoot\tools\conflict-tui\target\release\conflict-tui.exe"
if (Test-Path $rustExe) {
    Copy-Item $rustExe $installDir -Force
    Write-Host "  已带上 Rust TUI（有则用，没有自动回退 Spectre 版）"
} else {
    Write-Host "  未找到 conflict-tui.exe（cargo build --release 过一遍可以带上 Rust TUI，跳过不影响可用性）" -ForegroundColor Yellow
}

Write-Host "== 3/3 注册 git merge driver ==" -ForegroundColor Cyan
$scannerExe = Join-Path $installDir "NumDesTools.Scanner.exe"
git config --global merge.numdes.name "NumDesTools xlsx 冲突解决"
git config --global merge.numdes.driver "`"$scannerExe`" --conflict %A %B %O --no-add"

Write-Host ""
Write-Host "完成。任意仓库里 *.xlsx/*.xlsm 有 merge=numdes 声明（本仓库 .gitattributes 已加），" -ForegroundColor Green
Write-Host "以后 git merge / git rebase / 任何 GUI 走到 xlsx 冲突都会自动弹这个工具。" -ForegroundColor Green
