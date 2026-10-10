@echo off
title 亚马逊门锁售后分析 - 一键更新
chcp 65001 >nul

REM ── 定位 Python（多级探测：配置覆盖 → WorkBuddy 相对路径 → 用户目录任意版本 → PATH → 常见位置）──
set "PYTHON="
if exist "%~dp0python_path.txt" for /f "usebackq delims=" %%p in ("%~dp0python_path.txt") do if not defined PYTHON if exist "%%p" set "PYTHON=%%p"
if not defined PYTHON if exist "%~dp0..\..\..\.workbuddy\binaries\python\versions\3.13.12\python.exe" set "PYTHON=%~dp0..\..\..\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not defined PYTHON for /d %%d in ("%USERPROFILE%\.workbuddy\binaries\python\versions\*") do if not defined PYTHON if exist "%%d\python.exe" set "PYTHON=%%d\python.exe"
if not defined PYTHON for /f "delims=" %%p in ('where python 2^>nul') do if not defined PYTHON set "PYTHON=%%p"
if not defined PYTHON if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined PYTHON if exist "C:\Python313\python.exe" set "PYTHON=C:\Python313\python.exe"

set "CONVERT=%~dp0convert_doorlock_to_compact.py"
set "SYNC=%~dp0sync.ps1"
set "DECRYPT=%~dp0decrypt_xlsx.ps1"

echo ============================================
echo   亚马逊门锁售后分析系统 - 一键更新
echo ============================================
echo.

REM 拖入文件优先：xlsx → 解密→转换；json → 跳过转换直接同步
set "ARG=%~1"
if not "%ARG%"=="" goto HAS_ARG
echo [1/2] 从桌面 Excel 转换数据...
goto RUN_CONVERT

:HAS_ARG
if /i "%ARG:~-5%"==".json" (
    echo [1/2] 跳过转换（JSON，直接同步）
    goto RUN_SYNC
)
echo [1/2] 从拖入的 Excel 转换数据...
echo       文件: %ARG%
call :DECRYPT_PASS "%ARG%"
if errorlevel 1 goto FAIL_CONVERT
goto RUN_SYNC

:RUN_CONVERT
if not defined PYTHON (
    echo [X] 未找到 Python。请任选其一:
    echo     1^) 安装 WorkBuddy 后重试（其自带 Python 会被自动探测）
    echo     2^) 在项目目录新建 python_path.txt，写入 python.exe 完整路径（单行）
    pause
    exit /b 1
)
if not exist "%PYTHON%" (
    echo [X] Python 路径无效: %PYTHON%
    echo     请修改 python_path.txt，或运行"新电脑初始化.bat"重新配置
    pause
    exit /b 1
)

REM 自动模式：找桌面文件
set "AUTO_AFTERSALE="
set "AUTO_SALES="
if exist "%USERPROFILE%\Desktop\亚马逊门锁售后数据_合并.xlsx" set "AUTO_AFTERSALE=%USERPROFILE%\Desktop\亚马逊门锁售后数据_合并.xlsx" & goto AUTO_FOUND
for /f "delims=" %%f in ('powershell -NoProfile -Command "Get-ChildItem '%USERPROFILE%\Desktop\亚马逊门锁售后工单*.xlsx' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName"') do set "AUTO_AFTERSALE=%%f"
if not defined AUTO_AFTERSALE (
    echo [X] 桌面未找到售后工单/合并文件
    echo    预期: 亚马逊门锁售后工单*.xlsx 或 亚马逊门锁售后数据_合并.xlsx
    pause
    exit /b 1
)

:AUTO_FOUND
call :DECRYPT_PASS "%AUTO_AFTERSALE%"
if errorlevel 1 goto FAIL_CONVERT
goto RUN_SYNC

REM ── 子程序：解密 xlsx 然后传给 Python ──
:DECRYPT_PASS
set "SRC=%~1"
for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%DECRYPT%" -SourcePath "%SRC%"`) do set "DECRYPTED=%%i"
if errorlevel 1 exit /b 1
if not defined DECRYPTED (
    echo [X] 解密返回空路径
    exit /b 1
)
echo [convert] 传入: %DECRYPTED%
"%PYTHON%" "%CONVERT%" "%DECRYPTED%"
if errorlevel 1 exit /b 1
exit /b 0

:FAIL_CONVERT
echo.
echo [X] 数据转换失败！
pause
exit /b 1

:RUN_SYNC
echo.
echo [2/3] 刷新供应商映射（SKU->供应商）...
if not exist "%PYTHON%" (
    echo [警告] Python 未找到，跳过供应商映射刷新
) else (
    "%PYTHON%" "%~dp0build_supplier_map.py"
    if errorlevel 1 echo [警告] 供应商映射刷新失败（不影响数据更新）
)
echo.
echo [3/3] 双端推送...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%SYNC%"
if %ERRORLEVEL% neq 0 (
    echo [X] 同步失败！
    pause
    exit /b 1
)

echo.
echo ============================================
echo   完成！
echo   GitHub: https://zhongshanms.github.io/doorlock-after-sale-analysis/
echo ============================================
echo.
pause
