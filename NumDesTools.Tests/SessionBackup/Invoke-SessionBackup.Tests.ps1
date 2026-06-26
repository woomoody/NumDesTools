# Pester 3.4.0 兼容测试
# 运行: powershell -NoProfile -Command "Invoke-Pester -Script path\Invoke-SessionBackup.Tests.ps1 -EnableExit"

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = Join-Path $here 'Invoke-SessionBackup.ps1'
. $sut

function New-TestFixture {
    $tempRoot = Join-Path $env:TEMP "session-backup-test-$(Get-Random)"
    $sourcePath = Join-Path $tempRoot 'source'
    $repoPath = Join-Path $tempRoot 'repo'
    New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
    return [pscustomobject]@{
        TempRoot   = $tempRoot
        SourcePath = $sourcePath
        RepoPath   = $repoPath
    }
}

Describe 'Invoke-SessionBackup' {
    Context 'First run' {
        It 'inits repo and commits all, returns Committed=true with non-empty hash' {
            $f = New-TestFixture
            try {
                'hello world' | Set-Content -LiteralPath (Join-Path $f.SourcePath 'a.txt')
                $r = Invoke-SessionBackup -RepoPath $f.RepoPath -SourcePath $f.SourcePath

                $r.Changed      | Should Be $true
                $r.Committed    | Should Be $true
                $r.CommitHash   | Should Not BeNullOrEmpty
                (Test-Path (Join-Path $f.RepoPath '.git')) | Should Be $true
            }
            finally {
                Remove-Item -Recurse -Force $f.TempRoot -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'No changes' {
        It 'second run with unchanged source returns Changed=false and no new commit' {
            $f = New-TestFixture
            try {
                'hello' | Set-Content -LiteralPath (Join-Path $f.SourcePath 'a.txt')
                $first = Invoke-SessionBackup -RepoPath $f.RepoPath -SourcePath $f.SourcePath
                $second = Invoke-SessionBackup -RepoPath $f.RepoPath -SourcePath $f.SourcePath

                $second.Changed   | Should Be $false
                $second.Committed | Should Be $false
                $headAfterSecond = git -C $f.RepoPath rev-parse HEAD
                $headAfterSecond | Should Be $first.CommitHash
            }
            finally {
                Remove-Item -Recurse -Force $f.TempRoot -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'Incremental commit' {
        It 'changed source returns Changed=true, Committed=true, and a new hash different from first' {
            $f = New-TestFixture
            try {
                'v1' | Set-Content -LiteralPath (Join-Path $f.SourcePath 'a.txt')
                $first = Invoke-SessionBackup -RepoPath $f.RepoPath -SourcePath $f.SourcePath

                'v2' | Set-Content -LiteralPath (Join-Path $f.SourcePath 'a.txt')
                $second = Invoke-SessionBackup -RepoPath $f.RepoPath -SourcePath $f.SourcePath

                $second.Changed   | Should Be $true
                $second.Committed | Should Be $true
                $second.CommitHash | Should Not BeNullOrEmpty
                $second.CommitHash | Should Not Be $first.CommitHash
            }
            finally {
                Remove-Item -Recurse -Force $f.TempRoot -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'New file picked up' {
        It 'newly added source file is included in the next backup commit' {
            $f = New-TestFixture
            try {
                'a1' | Set-Content -LiteralPath (Join-Path $f.SourcePath 'a.txt')
                Invoke-SessionBackup -RepoPath $f.RepoPath -SourcePath $f.SourcePath | Out-Null

                'b1' | Set-Content -LiteralPath (Join-Path $f.SourcePath 'b.txt')
                $r = Invoke-SessionBackup -RepoPath $f.RepoPath -SourcePath $f.SourcePath

                $r.Committed | Should Be $true
                $stagedB = Join-Path $f.RepoPath 'sessions\b.txt'
                (Test-Path -LiteralPath $stagedB) | Should Be $true
                'b1' | Should Be (Get-Content -LiteralPath $stagedB -Raw).Trim()
            }
            finally {
                Remove-Item -Recurse -Force $f.TempRoot -ErrorAction SilentlyContinue
            }
        }
    }
}
