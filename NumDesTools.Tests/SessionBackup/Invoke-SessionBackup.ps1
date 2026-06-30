<#
.SYNOPSIS
    会话目录 git 备份的核心逻辑函数。
    负责：初始化工作仓（若不存在）→ 复制源文件 → git add -A → 判空 → 提交。
    返回 PSCustomObject: Changed/Committed/CommitHash/Message
#>
function Invoke-SessionBackup {
    param(
        [Parameter(Mandatory = $true)] [string] $RepoPath,
        [Parameter(Mandatory = $true)] [string] $SourcePath,
        [string] $StagingSubdir = 'sessions'
    )
    $result = [pscustomobject]@{
        Changed     = $false
        Committed   = $false
        CommitHash  = ''
        Message     = ''
    }

    if (-not (Test-Path -LiteralPath $SourcePath)) {
        $result.Message = "source not found: $SourcePath"
        return $result
    }

    # 初始化工作仓（首次运行）
    $gitDir = $RepoPath + '\.git'
    if (-not (Test-Path -LiteralPath $gitDir)) {
        New-Item -ItemType Directory -Path $RepoPath -Force | Out-Null
        git -C $RepoPath init --quiet
        git -C $RepoPath config user.email 'backup@local'
        git -C $RepoPath config user.name 'session-backup'
    }

    # 把源目录内容镜像复制到工作仓\<StagingSubdir>
    $stagingDir = $RepoPath.TrimEnd('\') + '\' + $StagingSubdir
    if (-not (Test-Path -LiteralPath $stagingDir)) {
        New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null
    }

    # 用 .NET 的 FileSystemInfo 直接调 robocopy.exe，避免 PowerShell 参数解析干扰
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'robocopy.exe'
    $psi.Arguments = '"' + $SourcePath + '" "' + $stagingDir + '" /MIR /NFL /NDL /NJH /NJS /NC /NS /NP'
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $p = [System.Diagnostics.Process]::Start($psi)
    $p.WaitForExit()
    $global:LASTEXITCODE = 0

    git -C $RepoPath add -A 2>&1 | Out-Null

    $status = git -C $RepoPath status --porcelain
    if ([string]::IsNullOrWhiteSpace($status)) {
        $result.Message = 'no changes'
        return $result
    }

    $result.Changed = $true
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    git -C $RepoPath commit -m "backup @ $ts" --quiet
    $hash = git -C $RepoPath rev-parse HEAD
    $result.Committed = $true
    $result.CommitHash = $hash
    $result.Message = "backup @ $ts"
    return $result
}
