@echo off
title 亚马逊门锁售后分析 - 一键更新
chcp 65001 >nul

set "PYTHON=%~dp0..\..\..\.workbuddy\binaries\python\versions\3.13.12\python.exe"
set "CONVERT=%~dp0convert_doorlock_to_compact.py"
set "SYNC=%~dp0sync.ps1"

echo ============================================
echo   亚马逊门锁售后分析系统 - 一键更新
echo ============================================
echo.

REM 拖入文件优先：xlsx → 转换脚本；json → 跳过转换直接同步
set "ARG=%~1"
if not "%ARG%"=="" goto HAS_ARG
echo [1/2] 从桌面 Excel 转换数据...
goto RUN_CONVERT

:HAS_ARG
if /i "%ARG:~-5%"==".json" (
    echo [1/2] 跳过转换（已检测到 JSON，直接同步）
    set "SKIP_CONVERT=1"
    goto RUN_SYNC
)
echo [1/2] 从拖入的 Excel 转换数据...
echo       文件: %ARG%

:RUN_CONVERT
if "%SKIP_CONVERT%"=="1" goto RUN_SYNC

if not exist "%PYTHON%" (
    echo [X] Python 未找到: %PYTHON%
    pause
    exit /b 1
)

if defined ARG (
    "%PYTHON%" "%CONVERT%" "%ARG%"
) else (
    "%PYTHON%" "%CONVERT%"
)
if %ERRORLEVEL% neq 0 (
    echo.
    echo [X] 数据转换失败！
    pause
    exit /b 1
)

echo.
echo [2/2] 推送到 GitHub...
echo.
:RUN_SYNC
powershell -NoProfile -ExecutionPolicy Bypass -File "%SYNC%"
if %ERRORLEVEL% neq 0 (
    echo [X] 同步失败！
    pause
    exit /b 1
)

echo.
echo ============================================
echo   完成！1-2 分钟后所有设备可刷新查看
echo   https://zhongshanms.github.io/doorlock-after-sale-analysis/
echo ============================================
echo.
pause
