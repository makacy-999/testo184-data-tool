@echo off
chcp 65001 >nul 2>&1
echo.
echo ╔══════════════════════════════════════════════════╗
echo ║   Testo 184 温度计数据读取汇总工具              ║
echo ╚══════════════════════════════════════════════════╝
echo.

:: 检查 Python
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ❌ 未找到 Python
    echo    请访问 https://www.python.org/downloads/ 下载安装
    echo    安装时请勾选 "Add Python to PATH"
    pause
    exit /b 1
)

echo ✅ Python:
python --version
echo.

:: 安装依赖
echo 📦 检查依赖...
pip install -q flask openpyxl 2>nul
if %ERRORLEVEL% neq 0 (
    echo ⚠️  pip 安装失败，尝试 --user 模式...
    pip install --user -q flask openpyxl
)
echo ✅ 依赖就绪
echo.

:: 启动
echo 🚀 启动服务...
echo.
python "%~dp0app.py"

pause
