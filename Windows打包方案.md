# 在 macOS 上打包 Windows EXE 的完整方案

## ⚠️ 重要说明

**macOS 无法直接打包 Windows .exe 文件**，原因：
- PyInstaller 不能交叉编译（macOS → Windows）
- Docker + Wine 方案复杂且不稳定
- 需要 Windows Python 解释器和 Windows 依赖库

## ✅ 推荐方案（按优先级排序）

### 方案 1：使用 Windows 电脑（最简单）⭐⭐⭐⭐⭐

**步骤：**
1. 将项目文件夹复制到 Windows 电脑（通过 U 盘、网络共享、云盘等）
2. 在 Windows 上双击运行 `build_windows.bat`
3. 等待 2-5 分钟，生成 `dist/加密Excel合并工具.exe`
4. 将 exe 文件复制回 macOS 或直接分发

**优点：**
- ✅ 最简单、最稳定
- ✅ 不需要额外安装软件
- ✅ 打包速度快

**如何传输文件：**
```bash
# macOS 上压缩项目
cd /Users/lyzz/PycharmProjects
zip -r excel-tool.zip excel/ -x "*.pyc" "__pycache__/*" ".venv/*" "dist/*" "build/*"

# 通过以下方式传输到 Windows：
# - U 盘
# - 网络共享 (SMB)
# - 云盘 (百度网盘、OneDrive 等)
# - AirDrop (如果有 Windows 电脑支持)
# - 微信/QQ 文件传输
```

---

### 方案 2：使用虚拟机（推荐开发环境）⭐⭐⭐⭐

**步骤：**

1. **安装虚拟机软件**（选择其一）：
   - VirtualBox（免费）：https://www.virtualbox.org/
   - Parallels Desktop（付费，性能好）：https://www.parallels.com/
   - VMware Fusion（付费）：https://www.vmware.com/

2. **安装 Windows 虚拟机**：
   - 下载 Windows 10/11 ISO 镜像
   - 在虚拟机中安装 Windows
   - 安装 Python 3.12+

3. **在虚拟机中打包**：
   ```bash
   # 在 Windows 虚拟机中
   # 1. 共享文件夹或复制项目到虚拟机
   # 2. 打开命令提示符
   cd C:\项目路径
   pip install pyinstaller
   pip install -r requirements.txt
   build_windows.bat
   ```

**优点：**
- ✅ 完全模拟 Windows 环境
- ✅ 可以测试 exe 运行效果
- ✅ 适合长期开发

**缺点：**
- ⚠️ 需要 Windows 许可证
- ⚠️ 占用磁盘空间（20GB+）
- ⚠️ 首次设置较复杂

---

### 方案 3：GitHub Actions 自动打包（云端）⭐⭐⭐⭐

**步骤：**

1. 将代码推送到 GitHub 仓库
2. 创建 `.github/workflows/build-windows.yml`
3. GitHub 自动在 Windows 虚拟机上打包
4. 下载生成的 exe 文件

**GitHub Actions 配置文件：**

```yaml
name: Build Windows EXE

on:
  push:
    tags:
      - 'v*'
  workflow_dispatch:  # 允许手动触发

jobs:
  build-windows:
    runs-on: windows-latest
    
    steps:
    - uses: actions/checkout@v3
    
    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.12'
    
    - name: Install dependencies
      run: |
        pip install -r requirements.txt
        pip install pyinstaller
    
    - name: Build EXE
      run: |
        pyinstaller --name "加密Excel合并工具" `
          --windowed `
          --onefile `
          --hidden-import openpyxl `
          --hidden-import olefile `
          --hidden-import cryptography `
          --hidden-import msoffcrypto `
          --hidden-import PyQt6 `
          --hidden-import PyQt6.sip `
          --exclude-module tkinter `
          --exclude-module unittest `
          --noconfirm `
          main.py
    
    - name: Upload artifact
      uses: actions/upload-artifact@v3
      with:
        name: 加密Excel合并工具
        path: dist/加密Excel合并工具.exe
    
    - name: Create Release
      if: startsWith(github.ref, 'refs/tags/')
      uses: softprops/action-gh-release@v1
      with:
        files: dist/加密Excel合并工具.exe
      env:
        GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

**使用方法：**
```bash
# 1. 初始化 git 仓库（如果没有）
cd /Users/lyzz/PycharmProjects/excel
git init
git add .
git commit -m "Initial commit"

# 2. 创建 GitHub 仓库并推送
git remote add origin https://github.com/你的用户名/excel-tool.git
git push -u origin main

# 3. 创建标签触发包
git tag v1.0.0
git push origin v1.0.0
```

**优点：**
- ✅ 不需要 Windows 电脑
- ✅ 自动化，一键打包
- ✅ 免费版可用
- ✅ 支持版本管理和发布

**缺点：**
- ⚠️ 需要 GitHub 账号
- ⚠️ 需要网络
- ⚠️ 首次配置稍复杂

---

### 方案 4：使用 CI/CD 服务 ⭐⭐⭐

类似 GitHub Actions，其他平台：
- **GitLab CI** - 免费的 Windows Runner
- **Azure Pipelines** - 每月免费 1800 分钟
- **AppVeyor** - 专为 Windows 设计

---

## ❌ 不推荐的方案

### Docker + Wine（复杂且不稳定）
- Wine 对 PyQt6 支持不完善
- Windows Python 嵌入版缺少必要模块
- 打包过程可能失败
- 生成的 exe 可能无法正常运行

### 交叉编译工具
- 目前没有可靠的 Python → Windows 交叉编译方案
- PyInstaller 不支持交叉编译
- cx_Freeze 等工具也不支持

---

## 🎯 最佳实践建议

### 短期方案（立即使用）
1. 找一台 Windows 电脑
2. 复制项目文件
3. 运行 `build_windows.bat`
4. 获取 exe 文件

### 中期方案（开发测试）
1. 安装 VirtualBox（免费）
2. 安装 Windows 10/11 虚拟机
3. 在虚拟机中开发和打包
4. 可以随时测试

### 长期方案（持续交付）
1. 使用 GitHub Actions
2. 配置自动打包流程
3. 每次发布新版本自动构建
4. 自动生成 Release 和下载链接

---

## 📋 快速检查清单

在 Windows 电脑上打包前，确保：

- [ ] 安装 Python 3.8+（推荐 3.12）
- [ ] 安装 PyInstaller：`pip install pyinstaller`
- [ ] 安装项目依赖：`pip install -r requirements.txt`
- [ ] 项目文件完整复制（包括 main.py、requirements.txt）
- [ ] 测试过 Python 脚本可以正常运行

打包后测试：

- [ ] exe 文件可以双击运行
- [ ] GUI 界面正常显示
- [ ] 拖拽文件功能正常
- [ ] 密码配置功能正常
- [ ] 合并功能正常
- [ ] 大文件处理正常（可选）

---

## 💡 常见问题

### Q1: 我没有 Windows 电脑怎么办？
A: 使用方案 3（GitHub Actions）最方便，完全免费。

### Q2: VirtualBox 安装 Windows 需要什么？
A: 
- VirtualBox（免费）
- Windows 10/11 ISO 镜像（可从微软官网下载）
- 至少 30GB 磁盘空间
- 至少 4GB 内存分配给虚拟机

### Q3: GitHub Actions 是免费的吗？
A: 是的，公共仓库完全免费，私有仓库每月有 2000 分钟免费额度。

### Q4: 打包的 exe 可以在任何 Windows 电脑上运行吗？
A: 是的，Windows 7/8/10/11 都可以，无需安装 Python。

### Q5: exe 文件有多大？
A: 约 80-150 MB，因为包含了 Python 运行时和所有依赖库。

---

## 📞 需要帮助？

如果遇到问题，可以：
1. 查看 [打包说明.md](./打包说明.md)
2. 检查 GitHub Issues
3. 联系开发者
