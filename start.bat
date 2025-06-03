@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
title 正在运行Python Web服务

echo 正在启动Python Web服务...
echo 请确保已安装Python和所需依赖库

:: 检查Python是否可用
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到Python，请先安装Python并添加到PATH
    pause
    exit /b
)

:: 运行Python脚本
python run_web.py

if %errorlevel% neq 0 (
    echo.
    echo Web服务启动失败，请检查错误信息
) else (
    echo.
    echo Web服务已正常启动
)

pause