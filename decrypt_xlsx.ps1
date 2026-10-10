param([string]$SourcePath)

if (-not $SourcePath -or -not (Test-Path $SourcePath -PathType Leaf)) {
    [Console]::Error.WriteLine("[decrypt] 文件不存在: $SourcePath")
    exit 1
}

# 读取文件头前4字节
$bytes = [IO.File]::ReadAllBytes($SourcePath)
$isEncrypted = ($bytes.Count -ge 4) -and ($bytes[0] -eq 0x88) -and ($bytes[1] -eq 0x7d) -and ($bytes[2] -eq 0x1c)

if (-not $isEncrypted) {
    # 非 DLP 加密，直接返回原路径
    Write-Output $SourcePath
    exit 0
}

Write-Host "[decrypt] 检测到 DLP 加密，使用 WPS 解密..." -ForegroundColor Yellow
$dest = "$env:TEMP\doorlock_decrypted_$(Get-Date -Format 'yyyyMMddHHmmss').xlsx"

try {
    $wps = New-Object -ComObject Ket.Application
    $wb = $wps.Workbooks.Open($SourcePath, 0, $true)
    $wb.SaveCopyAs($dest)
    $wb.Close($false)
    $wps.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    [Runtime.InteropServices.Marshal]::ReleaseComObject($wps) | Out-Null
} catch {
    [Console]::Error.WriteLine("[decrypt] WPS COM 解密失败: $_")
    exit 1
}

Write-Host "[decrypt] 已解密到: $dest" -ForegroundColor Green
Write-Output $dest
exit 0
