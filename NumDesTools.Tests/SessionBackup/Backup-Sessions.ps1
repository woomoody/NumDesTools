<#
.SYNOPSIS
    会话目录定时备份入口。
    被 SessionStart hook 或 Windows 计划任务调用。
    日志写到 Documents\tmp\session-backup.log
#>
param(
    # Claude Code 项目会话根（其下每个子目录是一个项目）
    [string] $ProjectsRoot = 'C:\Users\cent\.claude\projects',
    # 备份工作仓位置（独立目录，不污染源）
    [string] $RepoPath = 'C:\Users\cent\Documents\claude-sessions-backup'
)

$ErrorActionPreference = 'Continue'
$logDir = Join-Path $env:USERPROFILE 'Documents\tmp'
if (-not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$logFile = Join-Path $logDir 'session-backup.log'

# dot-source 核心函数
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Invoke-SessionBackup.ps1')

function Write-BackupLog {
    param([string] $Msg)
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Msg
    Add-Content -LiteralPath $logFile -Value $line -Encoding UTF8
}

try {
    if (-not (Test-Path -LiteralPath $ProjectsRoot)) {
        Write-BackupLog "SKIP: projects root not found $ProjectsRoot"
        return
    }
    # 遍历每个项目子目录，分别备份到 sessions/<项目名>/，避免多项目 memory 同名覆盖
    $projectDirs = Get-ChildItem -LiteralPath $ProjectsRoot -Directory -ErrorAction SilentlyContinue
    if (-not $projectDirs) {
        Write-BackupLog "NO-PROJECTS: nothing under $ProjectsRoot"
        return
    }
    $committed = $false
    foreach ($proj in $projectDirs) {
        $r = Invoke-SessionBackup -RepoPath $RepoPath -SourcePath $proj.FullName -StagingSubdir "sessions\$($proj.Name)"
        if ($r.Committed) {
            $committed = $true
            Write-BackupLog "COMMIT proj=$($proj.Name) hash=$($r.CommitHash) msg=$($r.Message)"
        }
        elseif ($r.Changed) {
            Write-BackupLog "CHANGED proj=$($proj.Name) not committed: $($r.Message)"
        }
        else {
            Write-BackupLog "NO-CHANGE proj=$($proj.Name): $($r.Message)"
        }
    }
    if (-not $committed) {
        Write-BackupLog "NO-CHANGE-ALL: no project had changes"
    }
}
catch {
    Write-BackupLog "ERROR: $($_.Exception.Message)"
}
