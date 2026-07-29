﻿﻿param(
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

# 宝塔服务器配置
$ServerRemote = "server"
$ServerUrl = "ssh://root@119.29.107.118:53844/www/wwwroot/doorlock.zhongshanzhiliang.top"
$PemPath = Join-Path $Cache "tencentbaota.pem"
$MaxRetries = 3

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

# ── 读取源 JSON 获取统计数据，生成 version.json ──
Write-Host "[0/4] 生成版本信息..."
try {
    $jsonContent = Get-Content -LiteralPath $SourceFile -Raw -Encoding UTF8
    $data = $jsonContent | ConvertFrom-Json
    $afterCount = $data.ar.Count
    $salesCount = $data.sr.Count
    $totalSales = 0
    $totalOrders = 0
    foreach ($s in $data.sr) {
        $totalSales += [int]$s.sq
        $totalOrders += [int]$s.oq
    }
    $now = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $autoVersion = "sync-" + (Get-Date).ToString("yyyyMMdd-HHmmss")
    $verJson = @{
        version = $autoVersion
        generated_at = $now
        total_after_sale = $afterCount
        total_sales = $totalSales
        total_orders = $totalOrders
    } | ConvertTo-Json
    $verPath = Join-Path $env:TEMP "doorlock-version.json"
    $verJson | Out-File -LiteralPath $verPath -Encoding UTF8
    Write-Host "  售后=$afterCount 销量=$salesCount 总销量=$totalSales 总订单=$totalOrders"
} catch {
    Write-Host "  [警告] 无法解析 JSON，跳过版本生成"
    Write-Host "  错误: $_"
    $verPath = $null
}

# ── 克隆或更新仓库 ──
if (Test-Path (Join-Path $Cache ".git")) {
    Write-Host ""
    Write-Host "[1/4] 更新本地仓库缓存..."
    Set-Location -LiteralPath $Cache
    Remove-Item -Path ".git\index.lock" -Force -ErrorAction SilentlyContinue
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
    & "$Git" clone --depth 1 $Repo $Cache
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

if ($verPath -and (Test-Path $verPath)) {
    $DestVer = Join-Path $Cache "data\version.json"
    Copy-Item -Path $verPath -Destination $DestVer -Force
    Write-Host "  [OK] data/version.json"
}

# ── 提交 ──
Write-Host ""
Write-Host "[3/4] 提交..."
& "$Git" config user.email "zhongshanms@github.com"
& "$Git" config user.name "门锁数据同步"
& "$Git" add data/after-sale-data-compact.json data/version.json
$commitMsg = "data sync: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - 售后$afterCount条 销量$salesCount条"
& "$Git" commit -m $commitMsg
if ($LASTEXITCODE -ne 0) {
    Write-Host "  (内容未变化，跳过提交)"
} else {
    Write-Host "  [OK]"
}

# ── 推送（双远程） ──
function Push-ToRemote {
    param([string]$RemoteName, [string]$SshKey, [string]$Label)
    Write-Host "[$Label] 正在推送..."

    if ($RemoteName -eq $ServerRemote) {
        if (-not (Test-Path -LiteralPath $SshKey)) {
            Write-Host "  [SKIP] 宝塔SSH密钥未找到: $SshKey"
            return $false
        }
        $env:GIT_SSH_COMMAND = "ssh -i `"$SshKey`" -o StrictHostKeyChecking=no -o UserKnownHostsFile=NUL"
    } else {
        $env:GIT_SSH_COMMAND = $null
    }

    & "$Git" push $RemoteName $Branch 2>&1 | ForEach-Object { Write-Host "  $_" }
    $ok = ($LASTEXITCODE -eq 0)
    if ($ok) { Write-Host "  [OK] $Label 推送成功" } else { Write-Host "  [FAIL] $Label 推送失败" }

    $env:GIT_SSH_COMMAND = $null
    return $ok
}

function Push-WithRetry {
    param([string]$RemoteName, [string]$SshKey, [string]$Label)
    for ($i = 1; $i -le $MaxRetries; $i++) {
        if ($i -gt 1) { Write-Host "  [重试] 第 $i/$MaxRetries 次..." }
        $result = Push-ToRemote -RemoteName $RemoteName -SshKey $SshKey -Label $Label
        if ($result) { return $true }
        if ($i -lt $MaxRetries) { Start-Sleep -Seconds 3 }
    }
    Write-Host "  [X] $Label 已重试 $MaxRetries 次，全部失败"
    return $false
}

Write-Host ""
Write-Host "[4/4] 推送到远端..."
# 确保 server 远程已添加
$remotes = & "$Git" remote 2>$null
if ($remotes -notcontains $ServerRemote) {
    & "$Git" remote add $ServerRemote $ServerUrl 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  [警告] 无法添加服务器远程: $ServerUrl"
    }
}

$ghOk = Push-WithRetry -RemoteName "origin" -SshKey $null -Label "GitHub"
Write-Host ""
$btOk = Push-WithRetry -RemoteName $ServerRemote -SshKey $PemPath -Label "宝塔"

Write-Host ""
Write-Host "============================================"
if ($ghOk -and $btOk) {
    Write-Host "  SUCCESS - 双端推送完成"
} elseif ($ghOk) {
    Write-Host "  PARTIAL - GitHub [OK]  | 宝塔 [FAIL]"
} elseif ($btOk) {
    Write-Host "  PARTIAL - GitHub [FAIL] | 宝塔 [OK]"
} else {
    Write-Host "  FAILED - 双端推送均失败"
    Write-Host "  常见原因：1. 网络不通  2. SSH 密钥未配置  3. 仓库权限"
    Write-Host ""
    Read-Host "按回车退出"
    exit 1
}
Write-Host ""
Write-Host "  站点地址："
Write-Host "  GitHub: https://zhongshanms.github.io/doorlock-after-sale-analysis/"
Write-Host "  宝塔:   https://doorlock.zhongshanzhiliang.top/"
Write-Host "============================================"
Write-Host ""
Read-Host "按回车退出"
