#!/bin/bash
# Testo 184 温度计数据读取汇总工具 - macOS 启动脚本

set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo ""
echo "╔══════════════════════════════════════════════════╗"
echo "║   Testo 184 温度计数据读取汇总工具              ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""

# 优先使用项目内虚拟环境
if [ -x "$SCRIPT_DIR/venv/bin/python" ]; then
    PYTHON_CMD="$SCRIPT_DIR/venv/bin/python"
    echo "✅ 使用项目虚拟环境"
else
    # 检查系统 Python
    PYTHON_CMD=""
    if command -v python3 &>/dev/null; then
        PYTHON_CMD="python3"
    elif command -v python &>/dev/null; then
        MAJOR=$(python --version 2>&1 | awk '{print $2}' | cut -d. -f1)
        [ "$MAJOR" -ge 3 ] && PYTHON_CMD="python"
    fi

    if [ -z "$PYTHON_CMD" ]; then
        echo "❌ 未找到 Python 3"
        echo "   安装方式: brew install python3"
        exit 1
    fi

    echo "✅ Python: $($PYTHON_CMD --version 2>&1)"

    # 首次运行：创建虚拟环境并安装依赖
    echo "📦 首次运行，创建虚拟环境并安装依赖（使用清华镜像）..."
    $PYTHON_CMD -m venv "$SCRIPT_DIR/venv"
    "$SCRIPT_DIR/venv/bin/pip" install -q -i https://pypi.tuna.tsinghua.edu.cn/simple flask openpyxl
    PYTHON_CMD="$SCRIPT_DIR/venv/bin/python"
    echo "✅ 依赖就绪"
fi

echo ""
echo "🚀 启动服务，请在浏览器打开: http://localhost:8000"
echo "   （注意：不是 5000 端口，macOS 的 AirPlay 占用了 5000）"
echo ""
$PYTHON_CMD "$SCRIPT_DIR/app.py"
