<#
.SYNOPSIS
    会话目录定时备份入口。
    被 SessionStart hook 或 Windows 计划任务调用。
    日志写到 Documents\tmp\session-backup.log
#>
param(
    # 要备份的会话目录（Claude Code 项目会话根）
    [string] $SourcePath = 'C:\Users\cent\.claude\projects\c--Pro-ExcelToolsAlbum-ExcelDna-Pro-NumDesTools',
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
    if (-not (Test-Path -LiteralPath $SourcePath)) {
        Write-BackupLog "SKIP: source not found $SourcePath"
        return
    }
    $r = Invoke-SessionBackup -RepoPath $RepoPath -SourcePath $SourcePath
    if ($r.Committed) {
        Write-BackupLog "COMMIT hash=$($r.CommitHash) msg=$($r.Message)"
    }
    elseif ($r.Changed) {
        Write-BackupLog "CHANGED but not committed: $($r.Message)"
    }
    else {
        Write-BackupLog "NO-CHANGE: $($r.Message)"
    }
}
catch {
    Write-BackupLog "ERROR: $($_.Exception.Message)"
}
