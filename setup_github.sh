#!/bin/bash
# GitHub 仓库快速配置脚本
# 运行此脚本自动完成 Git 配置和推送

set -e

echo "========================================"
echo "  GitHub 仓库快速配置"
echo "========================================"
echo

# 进入项目目录
cd "$(dirname "$0")"

# 检查是否已配置 Git
if [ -z "$(git config user.name)" ]; then
    echo "📝 配置 Git 用户信息"
    echo
    read -p "请输入你的 GitHub 用户名: " GIT_USERNAME
    read -p "请输入你的 GitHub 邮箱: " GIT_EMAIL
    
    git config user.name "$GIT_USERNAME"
    git config user.email "$GIT_EMAIL"
    echo "✓ Git 用户信息已配置"
    echo
else
    echo "✓ Git 用户信息已配置: $(git config user.name) <$(git config user.email)>"
    echo
fi

# 检查远程仓库
if git remote -v | grep -q origin; then
    echo "✓ 远程仓库已配置:"
    git remote -v
    echo
    read -p "是否要重新配置远程仓库? (y/N): " -n 1 -r
    echo
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        git remote remove origin
    else
        echo "跳过远程仓库配置"
    fi
fi

# 配置远程仓库
if ! git remote -v | grep -q origin; then
    echo "🔗 配置远程仓库"
    echo
    echo "请在浏览器中创建 GitHub 仓库："
    echo "  1. 访问 https://github.com/new"
    echo "  2. 填写仓库名称（例如：excel-merge-tool）"
    echo "  3. 选择 Public 或 Private"
    echo "  4. 点击 Create repository"
    echo "  5. 复制仓库地址（格式：https://github.com/用户名/仓库名.git）"
    echo
    
    read -p "粘贴你的 GitHub 仓库地址: " REPO_URL
    
    git remote add origin "$REPO_URL"
    echo "✓ 远程仓库已添加"
    echo
fi

# 添加文件
echo "📦 准备提交文件..."
git add .

# 检查是否有更改
if git diff --staged --quiet; then
    echo "ℹ️  没有新的更改需要提交"
else
    read -p "提交信息 (直接回车使用默认): " COMMIT_MSG
    
    if [ -z "$COMMIT_MSG" ]; then
        COMMIT_MSG="更新：$(date '+%Y-%m-%d %H:%M:%S')"
    fi
    
    git commit -m "$COMMIT_MSG"
    echo "✓ 文件已提交"
fi

# 推送
echo
echo "🚀 推送到 GitHub..."
echo

# 检查分支
CURRENT_BRANCH=$(git branch --show-current)
if [ -z "$CURRENT_BRANCH" ]; then
    git branch -M main
    CURRENT_BRANCH="main"
fi

# 推送
if git push -u origin $CURRENT_BRANCH 2>&1; then
    echo
    echo "========================================"
    echo "  ✅ 推送成功！"
    echo "========================================"
    echo
    echo "📋 下一步："
    echo
    echo "1. 打开浏览器查看你的仓库"
    REPO_URL=$(git remote get-url origin | sed 's/\.git$//')
    echo "   $REPO_URL"
    echo
    echo "2. 点击 'Actions' 标签查看打包进度"
    echo "   $REPO_URL/actions"
    echo
    echo "3. 等待 5-10 分钟，打包完成后下载文件"
    echo
    echo "💡 提示："
    echo "  - 首次推送可能需要登录 GitHub"
    echo "  - 如果推送失败，检查仓库地址是否正确"
    echo "  - 查看 GitHub配置指南.md 获取详细帮助"
    echo
else
    echo
    echo "========================================"
    echo "  ❌ 推送失败"
    echo "========================================"
    echo
    echo "可能的原因："
    echo "  1. 仓库地址错误"
    echo "  2. 未授权访问"
    echo "  3. 网络问题"
    echo
    echo "解决方案："
    echo "  1. 检查仓库地址：git remote -v"
    echo "  2. 重新登录 GitHub"
    echo "  3. 查看 GitHub配置指南.md"
    echo
    exit 1
fi
