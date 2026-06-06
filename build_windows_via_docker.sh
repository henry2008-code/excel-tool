#!/bin/bash
# macOS 上通过 Docker 打包 Windows exe
# 需要安装 Docker Desktop for Mac

set -e

echo "========================================"
echo "  macOS 交叉编译 Windows EXE"
echo "  使用 Docker + Wine + PyInstaller"
echo "========================================"
echo

# 检查 Docker
if ! command -v docker &> /dev/null; then
    echo "❌ 错误: 未检测到 Docker"
    echo
    echo "请先安装 Docker Desktop for Mac:"
    echo "  1. 访问 https://www.docker.com/products/docker-desktop"
    echo "  2. 下载并安装 Docker Desktop"
    echo "  3. 启动 Docker Desktop"
    echo
    exit 1
fi

echo "✓ Docker 已就绪"
echo

# 项目路径
PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
IMAGE_NAME="pyinstaller-windows"

# 检查镜像是否存在
if ! docker images --format '{{.Repository}}' | grep -q "^${IMAGE_NAME}$"; then
    echo "[1/3] 构建 Docker 镜像..."
    echo "这可能需要 10-20 分钟（首次需要下载基础镜像和依赖）"
    echo
    
    docker build -t ${IMAGE_NAME} -f Dockerfile.windows .
    
    echo "✓ Docker 镜像构建完成"
    echo
else
    echo "✓ Docker 镜像已存在"
    echo
fi

echo "[2/3] 清理旧的打包文件..."
rm -rf "${PROJECT_DIR}/dist"
rm -rf "${PROJECT_DIR}/build"
rm -f "${PROJECT_DIR}"/*.spec
echo

echo "[3/3] 开始打包 Windows EXE..."
echo

docker run --rm \
    -v "${PROJECT_DIR}:/app" \
    -w /app \
    ${IMAGE_NAME} \
    pyinstaller --name "加密Excel合并工具" \
        --windowed \
        --onefile \
        --hidden-import openpyxl \
        --hidden-import olefile \
        --hidden-import cryptography \
        --hidden-import msoffcrypto \
        --hidden-import PyQt6 \
        --hidden-import PyQt6.sip \
        --exclude-module tkinter \
        --exclude-module unittest \
        --noconfirm \
        main.py

echo
echo "========================================"
echo "  ✅ 打包完成！"
echo "========================================"
echo

if [ -f "${PROJECT_DIR}/dist/加密Excel合并工具.exe" ]; then
    EXE_SIZE=$(du -h "${PROJECT_DIR}/dist/加密Excel合并工具.exe" | cut -f1)
    echo "📦 输出文件: dist/加密Excel合并工具.exe"
    echo "📏 文件大小: ${EXE_SIZE}"
    echo
    echo "💡 提示:"
    echo "  - 该 exe 文件可在 Windows 电脑上直接运行"
    echo "  - 目标电脑无需安装 Python 或任何依赖"
    echo "  - 建议先复制到其他 Windows 电脑测试"
    echo
    
    # 询问是否打开目录
    read -p "是否打开 dist 目录? (y/N): " -n 1 -r
    echo
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        open "${PROJECT_DIR}/dist"
    fi
else
    echo "❌ 错误: 未找到生成的 exe 文件"
    echo "请检查上方的错误信息"
    exit 1
fi
