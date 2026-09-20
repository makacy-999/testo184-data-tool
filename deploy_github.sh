#!/bin/bash
# ─────────────────────────────────────────────────────────
# Testo 184 工具 - 一键部署到 GitHub
# 用法: ./deploy_github.sh [仓库名]
# 例:   ./deploy_github.sh testo184-data-tool
# ─────────────────────────────────────────────────────────

set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

REPO_NAME="${1:-testo184-data-tool}"
REPO_DESC="Testo 184 温度计数据读取汇总工具（Flask + SQLite，跨平台）"

echo ""
echo "╔══════════════════════════════════════════════════╗"
echo "║   部署 Testo 184 工具到 GitHub                    ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""

# 1. 确保 git 仓库已初始化
if [ ! -d ".git" ]; then
    git init -b main
    echo "✅ 已初始化 git 仓库"
fi

# 2. 提交当前代码
git add -A
if git diff --cached --quiet 2>/dev/null; then
    echo "✅ 代码无变更，跳过提交"
else
    git commit -m "Testo 184 温度计数据读取汇总工具 v1.0"
    echo "✅ 已提交代码"
fi

# 3. 检查 gh CLI
if ! command -v gh &>/dev/null; then
    echo "📦 安装 GitHub CLI..."
    if command -v brew &>/dev/null; then
        brew install gh
    else
        echo "❌ 未找到 brew，请先安装 GitHub CLI："
        echo "   方式1: 安装 Homebrew 后重试本脚本"
        echo "   方式2: 手动下载 https://cli.github.com/"
        exit 1
    fi
fi
echo "✅ GitHub CLI 就绪"

# 4. 登录 GitHub（浏览器授权）
if ! gh auth status &>/dev/null; then
    echo ""
    echo "🔐 需要登录 GitHub（将在浏览器中打开授权页面）..."
    echo "   按回车继续，按 Ctrl+C 取消"
    read -r
    gh auth login -p https -w
fi
echo "✅ 已登录: $(gh api user -q .login)"

# 5. 创建远程仓库并推送
echo ""
echo "🚀 创建仓库 $REPO_NAME 并推送..."
if gh repo view "$REPO_NAME" &>/dev/null 2>&1; then
    echo "⚠️  仓库已存在，推送更新..."
    git remote remove origin 2>/dev/null || true
    git remote add origin "https://github.com/$(gh api user -q .login)/$REPO_NAME.git"
    git push -u origin main
else
    gh repo create "$REPO_NAME" --public --source=. --remote=origin --push --description "$REPO_DESC"
fi

echo ""
echo "╔══════════════════════════════════════════════════╗"
echo "║   ✅ 部署完成！                                   ║"
echo "╠══════════════════════════════════════════════════╣"
echo "║   仓库地址:                                       ║"
echo "║   https://github.com/$(gh api user -q .login)/$REPO_NAME"
echo "╚══════════════════════════════════════════════════╝"
echo ""
