@echo off
chcp 65001 >nul
title 涉诉画像工作台
cd /d "%~dp0"

if not exist "runtime\python\python.exe" (
  echo.
  echo [错误] 缺少 runtime\python，压缩包可能没有解压完整。
  echo 请把整个文件夹重新解压一次，不要只拷其中一部分。
  echo.
  pause
  exit /b 1
)
if not exist "runtime\node\tyc.cmd" (
  echo.
  echo [错误] 缺少 runtime\node，压缩包可能没有解压完整。
  echo.
  pause
  exit /b 1
)

set "PATH=%~dp0runtime\node;%PATH%"

echo 正在启动工作台，浏览器会自动打开 http://127.0.0.1:8777
echo 关掉这个黑窗口即停止工作台。
echo.
"runtime\python\python.exe" app.py

echo.
echo 工作台已停止。
pause
