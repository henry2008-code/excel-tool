# macOS 打包说明

## ✅ 已完成打包

生成的 macOS 应用程序位于：`dist/加密Excel合并工具.app`

## 📦 打包信息

- **文件大小**：34 MB
- **应用类型**：macOS .app 应用程序包
- **架构**：x86_64 (Intel)
- **Python 版本**：3.12
- **PyInstaller 版本**：6.20.0

## 🚀 使用方法

### 方式一：直接运行
1. 打开 `dist` 文件夹
2. 双击 `加密Excel合并工具.app`
3. 如果提示"无法验证开发者"，前往 系统设置 > 隐私与安全性，点击"仍要打开"

### 方式二：命令行运行
```bash
./dist/加密Excel合并工具
```

## ⚠️ 首次运行提示

macOS 可能阻止未签名的应用运行，解决方法：

```bash
# 方法 1：在终端中直接运行（绕过 Gatekeeper）
./dist/加密Excel合并工具

# 方法 2：移除隔离属性
xattr -cr dist/加密Excel合并工具.app

# 方法 3：系统设置中允许
系统设置 > 隐私与安全性 > 安全性 > 允许
```

## 📝 重新打包

修改代码后重新打包：

```bash
cd /Users/lyzz/PycharmProjects/excel
source .venv/bin/activate

# 清理旧文件
rm -rf dist build *.spec

# 重新打包
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
```

## 🔧 优化选项

### 减小文件大小
```bash
# 使用 UPX 压缩（需先安装 upx: brew install upx）
pyinstaller --onefile --windowed --upx-dir=/usr/local/bin ...

# 或使用 onedir 模式（启动更快，体积更小）
pyinstaller --onedir --windowed ...
```

### 添加应用图标
```bash
# 准备 .icns 图标文件
pyinstaller --onefile --windowed --icon=app.icns ...
```

### 支持 Apple Silicon (M1/M2)
```bash
# 在 M1/M2 Mac 上运行打包命令即可自动打包 arm64 版本
pyinstaller --name "加密Excel合并工具" --windowed --onefile ...
```

## 📤 分发

将 `dist/加密Excel合并工具.app` 压缩后分发：

```bash
cd dist
zip -r 加密Excel合并工具.zip 加密Excel合并工具.app
```

## 🐛 常见问题

### Q: 双击无反应？
A: 在终端中运行查看错误：
```bash
./dist/加密Excel合并工具
```

### Q: 提示缺少模块？
A: 添加 `--hidden-import` 参数重新打包。

### Q: 如何在其他 Mac 上运行？
A: 直接复制 .app 文件即可，目标电脑无需安装 Python。

### Q: 打包后启动很慢？
A: 这是 `--onefile` 模式的特性（需解压到临时目录）。可改用 `--onedir` 模式。

## 📋 文件结构

```
dist/
├── 加密Excel合并工具          # 单个可执行文件（onefile 模式）
└── 加密Excel合并工具.app/     # macOS 应用包（windowed 模式）
    └── Contents/
        ├── MacOS/             # 可执行文件
        ├── Resources/         # 资源文件
        ├── Frameworks/        # 依赖框架
        └── Info.plist         # 应用信息
```
