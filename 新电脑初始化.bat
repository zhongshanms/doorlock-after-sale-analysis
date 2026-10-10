@echo off
chcp 65001 >nul
title 新电脑初始化 - 亚马逊门锁售后分析系统
setlocal enabledelayedexpansion

echo ============================================
echo   新电脑初始化（部署参数 + 环境自检）
echo ============================================
echo.
echo 直接回车 = 保留方括号中的默认值。
echo.

set "DEF_SRV=root@119.29.107.118:/www/wwwroot/doorlock.zhongshanzhiliang.top"
set "DEF_SITE=https://doorlock.zhongshanzhiliang.top/"
set "DEF_REPO=git@github.com:zhongshanms/doorlock-after-sale-analysis.git"
set "DEF_PAGES=https://zhongshanms.github.io/doorlock-after-sale-analysis/"

set /p SRV=服务器地址 [%DEF_SRV%]:
if "!SRV!"=="" set "SRV=%DEF_SRV%"

set /p SITE=站点网址 [%DEF_SITE%]:
if "!SITE!"=="" set "SITE=%DEF_SITE%"

set /p REPO=GitHub 仓库 [%DEF_REPO%]:
if "!REPO!"=="" set "REPO=%DEF_REPO%"

set /p PAGES=GitHub Pages 网址 [%DEF_PAGES%]:
if "!PAGES!"=="" set "PAGES=%DEF_PAGES%"

echo.
echo [1/3] 写入 deploy-config.json ...
(
  echo {
  echo   "github_repo": "!REPO!",
  echo   "server_url": "!SRV!",
  echo   "pages_url": "!PAGES!",
  echo   "site_url": "!SITE!"
  echo }
) > "%~dp0deploy-config.json"
if exist "%~dp0deploy-config.json" (echo   [OK] deploy-config.json) else (echo   [X] 写入失败 & pause & exit /b 1)

echo.
echo [2/3] 探测 Python ...
set "PYTHON="
if exist "%~dp0python_path.txt" for /f "usebackq delims=" %%p in ("%~dp0python_path.txt") do if not defined PYTHON if exist "%%p" set "PYTHON=%%p"
if not defined PYTHON if exist "%~dp0..\..\..\.workbuddy\binaries\python\versions\3.13.12\python.exe" set "PYTHON=%~dp0..\..\..\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not defined PYTHON for /d %%d in ("%USERPROFILE%\.workbuddy\binaries\python\versions\*") do if not defined PYTHON if exist "%%d\python.exe" set "PYTHON=%%d\python.exe"
if not defined PYTHON for /f "delims=" %%p in ('where python 2^>nul') do if not defined PYTHON set "PYTHON=%%p"
if not defined PYTHON if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined PYTHON if exist "C:\Python313\python.exe" set "PYTHON=C:\Python313\python.exe"

if defined PYTHON (
    > "%~dp0python_path.txt" echo !PYTHON!
    echo   [OK] !PYTHON!
    echo   [OK] 已记录到 python_path.txt
) else (
    echo   [X] 未找到 Python
    echo       请安装 WorkBuddy（自带 Python）或 Python 3.x 后重新运行本脚本
    echo       也可手动新建 python_path.txt，写入 python.exe 完整路径
)

echo.
echo [3/3] 依赖与工具自检 ...
set "GIT="
for /f "delims=" %%p in ('where git 2^>nul') do if not defined GIT set "GIT=%%p"
if not defined GIT if exist "%USERPROFILE%\.workbuddy\binaries\PortableGit\versions" for /f "delims=" %%p in ('dir /b /s "%USERPROFILE%\.workbuddy\binaries\PortableGit\versions\git.exe" 2^>nul') do if not defined GIT set "GIT=%%p"
if not defined GIT if exist "C:\Program Files\Git\cmd\git.exe" set "GIT=C:\Program Files\Git\cmd\git.exe"
if defined GIT (echo   [OK] Git: !GIT!) else (echo   [X] 未找到 Git —— 需安装 Git for Windows: https://git-scm.com/download/win)

if defined PYTHON (
    "!PYTHON!" -c "import openpyxl" >nul 2>&1
    if errorlevel 1 (
        echo   [..] 正在安装 openpyxl ...
        "!PYTHON!" -m pip install openpyxl
        "!PYTHON!" -c "import openpyxl" >nul 2>&1
        if errorlevel 1 (echo   [X] openpyxl 安装失败，请检查网络) else (echo   [OK] openpyxl 已安装)
    ) else (
        echo   [OK] openpyxl 已就绪
    )
    "!PYTHON!" -c "import importlib,site,sys;usp=site.getusersitepackages();sys.path.append(usp);importlib.invalidate_caches();import openpyxl" >nul 2>&1
    if errorlevel 1 (echo   [警告] openpyxl 可导入但 user site 需加固，见部署指南"常见问题") else (echo   [OK] user site 依赖链正常)
)

echo.
echo ============================================
echo   初始化完成。
echo   下一步: 双击"一键更新.bat"（可拖入 xlsx/json）
echo   首次会在 .sync_cache 克隆仓库，需要 GitHub SSH 已配置。
echo ============================================
pause
