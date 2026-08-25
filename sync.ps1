param(
    [string]$SourceFile
)

$ErrorActionPreference = "Stop"

Write-Host "============================================"
Write-Host "  数据上传 - 亚马逊门锁售后分析系统"
Write-Host "============================================"
Write-Host ""

if (-not $SourceFile) {
    # Priority: project data folder (non-encrypted) > desktop (may be 绿盾-encrypted)
    $ProjectData = Join-Path $PSScriptRoot "data\after-sale-data-compact.json"
    $DesktopJson = Join-Path $env:USERPROFILE "Desktop\after-sale-data-compact.json"
    if (Test-Path -LiteralPath $ProjectData) {
        $SourceFile = $ProjectData
        Write-Host "[提示] 未拖入文件，已自动定位项目 data 目录文件"
    } elseif (Test-Path -LiteralPath $DesktopJson) {
        $SourceFile = $DesktopJson
        Write-Host "[提示] 未拖入文件，已自动定位桌面文件"
    } else {
        Write-Host "用法：把 after-sale-data-compact.json 拖到 上传数据.bat 上"
        Write-Host "      或直接双击运行（脚本会自动查找项目 data 目录或桌面文件）"
        Write-Host ""
        Read-Host "按回车退出"
        exit 1
    }
}

if (-not (Test-Path -LiteralPath $SourceFile)) {
    Write-Host "[X] 文件不存在：$SourceFile"
    Read-Host "按回车退出"
    exit 1
}

# 检测绿盾加密
$IsEncrypted = $false
try {
    $FirstLine = Get-Content -LiteralPath $SourceFile -TotalCount 1 -ErrorAction Stop
    if ($FirstLine -notmatch '^\s*[\{\[]') { $IsEncrypted = $true }
} catch {
    $IsEncrypted = $true
}

if ($IsEncrypted) {
    Write-Host "[检测] 文件被绿盾加密: $(Split-Path $SourceFile -Leaf)"
    # Try fallback: project data folder
    $ProjectData = Join-Path $PSScriptRoot "data\after-sale-data-compact.json"
    if (Test-Path -LiteralPath $ProjectData) {
        $FallbackLine = Get-Content -LiteralPath $ProjectData -TotalCount 1 -ErrorAction SilentlyContinue
        if ($FallbackLine -match '^\s*[\{\[]') {
            Write-Host "[自动] 改用项目 data 目录中的未加密版本"
            $SourceFile = $ProjectData
            $IsEncrypted = $false
        } else {
            Write-Host "[X] 项目 data 目录中的文件也被加密，请先用记事本另存解密的文件"
            Write-Host "    路径: $ProjectData"
            Read-Host "按回车退出"
            exit 1
        }
    } else {
        Write-Host "[X] 项目 data 目录中未找到文件"
        Write-Host "    请运行 convert_doorlock_to_compact.py 生成数据，或双击 BAT 不拖文件"
        Read-Host "按回车退出"
        exit 1
    }
}

$Repo = "git@github.com:zhongshanms/doorlock-after-sale-analysis.git"
$Cache = Join-Path $PSScriptRoot ".sync_cache"
$Branch = "main"

# 宝塔同步已移除（2026-08-03）：改为本地 git push origin 仅推 GitHub

Write-Host "源文件：$(Split-Path $SourceFile -Leaf)"
Write-Host ""

# ── 查找 Git ──
$Git = @(
    "$env:USERPROFILE\.workbuddy\vendor\PortableGit\mingw64\bin\git.exe",
    "$env:USERPROFILE\.workbuddy\vendor\PortableGit\cmd\git.exe",
    "$env:USERPROFILE\scoop\shims\git.exe",
    "$env:ProgramData\scoop\shims\git.exe",
    "C:\Program Files\Git\bin\git.exe",
    "C:\Program Files\Git\cmd\git.exe",
    "C:\Program Files (x86)\Git\bin\git.exe",
    "C:\Program Files (x86)\Git\cmd\git.exe",
    "$env:LOCALAPPDATA\Programs\Git\bin\git.exe",
    "$env:LOCALAPPDATA\Programs\Git\cmd\git.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $Git) {
    # 通配扫描 WorkBuddy binaries\PortableGit\versions\*\mingw64\bin\git.exe
    # （当前 1.2.0 路径不在上方候选表中；以后版本升级也无需改脚本）
    $found = Get-ChildItem -Path "$env:USERPROFILE\.workbuddy\binaries\PortableGit\versions" -Filter "git.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) { $Git = $found.FullName }
}

if (-not $Git) {
    $GitCmd = Get-Command git -ErrorAction SilentlyContinue
    if ($GitCmd) { $Git = $GitCmd.Source }
}

if (-not $Git) {
    $regPaths = @(
        "HKLM:\SOFTWARE\GitForWindows",
        "HKLM:\SOFTWARE\Wow6432Node\GitForWindows",
        "HKCU:\SOFTWARE\GitForWindows"
    )
    foreach ($rp in $regPaths) {
        if (Test-Path $rp) {
            try {
                $installPath = (Get-ItemProperty $rp -Name InstallPath -ErrorAction SilentlyContinue).InstallPath
                if ($installPath) {
                    $candidate = Join-Path $installPath "bin\git.exe"
                    if (Test-Path $candidate) { $Git = $candidate; break }
                    $candidate = Join-Path $installPath "cmd\git.exe"
                    if (Test-Path $candidate) { $Git = $candidate; break }
                }
            } catch { }
        }
    }
}

if (-not $Git) {
    Write-Host "[X] 未找到 Git，请安装 Git for Windows"
    Write-Host "    下载地址：https://git-scm.com/download/win"
    Write-Host ""
    Read-Host "按回车退出"
    exit 1
}

Write-Host "[Git] $Git"
& "$Git" --version
if ($LASTEXITCODE -ne 0) {
    Write-Host "[X] Git 无法运行"
    Read-Host "按回车退出"
    exit 1
}
Write-Host ""

# ── 克隆或更新仓库 ──
if (Test-Path (Join-Path $Cache ".git")) {
    Write-Host ""
    Write-Host "[1/4] 更新本地仓库缓存..."
    Set-Location -LiteralPath $Cache
    Remove-Item -Path ".git\index.lock" -Force -ErrorAction SilentlyContinue
    # 补全老浅克隆（宝塔 force push 拒绝浅仓）
    if (Test-Path ".git\shallow") { & "$Git" fetch --unshallow 2>&1 | Out-Null }
    & "$Git" pull origin $Branch
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  [警告] 拉取失败，尝试重置..."
        & "$Git" fetch origin $Branch
        & "$Git" reset --hard "origin/$Branch"
    }
    Write-Host "  [OK]"
} else {
    Write-Host ""
    Write-Host "[1/4] 首次使用，克隆仓库..."
    if (Test-Path $Cache) { Remove-Item -Path $Cache -Recurse -Force }
    # 完整克隆（force push 给宝塔时服务器拒绝浅克隆）
    & "$Git" clone $Repo $Cache
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[X] 克隆失败！"
        Write-Host "请确认："
        Write-Host "  1. 已安装 Git for Windows"
        Write-Host "  2. SSH 已配置：ssh -T git@github.com"
        Write-Host "  3. 网络可以访问 GitHub"
        Write-Host ""
        Read-Host "按回车退出"
        exit 1
    }
    Write-Host "  [OK]"
}

# ── 复制数据文件 ──
Write-Host ""
Write-Host "[2/4] 复制数据文件..."
Set-Location -LiteralPath $Cache
$DestData = Join-Path $Cache "data\after-sale-data-compact.json"
Copy-Item -Path $SourceFile -Destination $DestData -Force
Write-Host "  [OK] data/after-sale-data-compact.json"

# ── 提交 ──
Write-Host ""
Write-Host "[2/4] 提交..."
& "$Git" config user.email "zhongshanms@github.com"
& "$Git" config user.name "门锁数据同步"
& "$Git" add data/after-sale-data-compact.json
$commitMsg = "data sync: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
& "$Git" commit -m $commitMsg
if ($LASTEXITCODE -ne 0) {
    Write-Host "  (内容未变化，跳过提交)"
} else {
    Write-Host "  [OK]"
}

# ── 推送（双远程：GitHub + 宝塔） ──
$ServerRemote = "server"
$ServerUrl = "root@119.29.107.118:/www/wwwroot/doorlock.zhongshanzhiliang.top"
$MaxRetries = 3

function Push-WithRetry {
    param($gitPath, [string]$remoteName = 'origin', [switch]$Force)
    for ($i = 1; $i -le $MaxRetries; $i++) {
        if ($i -gt 1) { Write-Host "  [重试] $i/$MaxRetries，5秒后..." ; Start-Sleep 5 }
        $saveEAP = $ErrorActionPreference
        $ErrorActionPreference = 'SilentlyContinue'
        if ($Force) {
            $pushOutput = & $gitPath push --force $remoteName $Branch 2>&1 | Out-String
        } else {
            $pushOutput = & $gitPath push $remoteName $Branch 2>&1 | Out-String
        }
        $ErrorActionPreference = $saveEAP
        if ($LASTEXITCODE -eq 0) { return $true }
        Write-Host "  [WARN] git push 失败 (第$i/$MaxRetries 次):"
        $pushOutput.Trim() -split "`n" | Where-Object { $_ } | ForEach-Object { Write-Host "    $_" }
    }
    return $false
}

function Push-ToRemote {
    param($gitPath, [string]$remoteName, [string]$label, [string]$successUrl, [switch]$Force)
    Write-Host "  推送 $label ..."
    $ok = Push-WithRetry $gitPath $remoteName -Force:$Force
    if ($ok) {
        Write-Host "  [OK] $label 推送成功"
        if ($successUrl) { Write-Host "       $successUrl" }
    } else {
        Write-Host "  [FAIL] $label 推送失败（已重试 $MaxRetries 次）"
    }
    return $ok
}

Write-Host ""
Write-Host "[3/4] 推送数据..."

# 确保 server 远程已添加
$remotes = & "$Git" remote 2>$null
if ($remotes -notcontains $ServerRemote) {
    & "$Git" remote add $ServerRemote $ServerUrl 2>$null
}

$results = @{}
$results['github'] = Push-ToRemote "$Git" 'origin' 'GitHub' 'https://zhongshanms.github.io/doorlock-after-sale-analysis/'
$results['server'] = Push-ToRemote "$Git" $ServerRemote '宝塔服务器' 'https://doorlock.zhongshanzhiliang.top/' -Force

Write-Host ""
Write-Host "============================================"
if ($results['github'] -and $results['server']) {
    Write-Host "  [SUCCESS] 双端推送完成"
} elseif ($results['github']) {
    Write-Host "  [PARTIAL] GitHub [OK] | 宝塔 [FAIL]"
} elseif ($results['server']) {
    Write-Host "  [PARTIAL] 宝塔 [OK] | GitHub [FAIL]"
} else {
    Write-Host "  [FAILED] 双端推送均失败"
}
Write-Host "============================================"
Write-Host ""
Read-Host "按回车退出"
