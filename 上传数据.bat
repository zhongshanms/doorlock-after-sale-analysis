@echo off
title 门锁售后数据同步 - 拖拽 JSON 文件到此处
chcp 65001 >nul

set "PS1=%~dp0sync.ps1"

if "%~1"=="" (
    echo 把 after-sale-data-compact.json 拖到这个图标上！
    echo.
    echo 或直接双击运行（会自动查找桌面文件）
    echo.
    pause
    exit /b 1
)

if not exist "%~1" (
    echo 文件不存在: %~1
    pause
    exit /b 1
)

echo ============================================
echo   亚马逊门锁售后分析系统 - 数据同步
echo ============================================
echo 源文件: %~1
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" "%~1"

echo.
pause
