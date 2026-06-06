# GitHub Actions 自动打包 Windows EXE 配置指南

## 📋 前置检查清单

在开始之前，请确认：

- [ ] 已安装 Git（✅ 已确认）
- [ ] 有 GitHub 账号（如没有，访问 https://github.com 注册）
- [ ] 项目代码在 `/Users/lyzz/PycharmProjects/excel`

---

## 🚀 完整配置步骤（约 10 分钟）

### 步骤 1：配置 Git 用户信息

```bash
cd /Users/lyzz/PycharmProjects/excel

# 设置你的 Git 用户名和邮箱（替换成你自己的）
git config user.name "你的GitHub用户名"
git config user.email "你的GitHub邮箱"
```

**示例：**
```bash
git config user.name "zhangsan"
git config user.email "zhangsan@example.com"
```

---

### 步骤 2：在 GitHub 创建仓库

1. **打开浏览器**，访问 https://github.com/new

2. **填写仓库信息**：
   - **Repository name**: `excel-merge-tool`（或你喜欢的名字）
   - **Description**: `加密Excel合并工具 - 支持大文件处理`
   - **Visibility**: 选择 `Public`（公开，免费）或 `Private`（私有）
   - ❌ **不要勾选** "Add a README file"
   - ❌ **不要勾选** "Add .gitignore"
   - ❌ **不要勾选** "Choose a license"

3. **点击** `Create repository` 按钮

4. **复制仓库地址**，格式如下：
   ```
   https://github.com/你的用户名/excel-merge-tool.git
   ```

---

### 步骤 3：连接本地仓库到 GitHub

```bash
cd /Users/lyzz/PycharmProjects/excel

# 添加远程仓库（替换成你的仓库地址）
git remote add origin https://github.com/你的用户名/excel-merge-tool.git

# 添加所有文件
git add .

# 提交
git commit -m "初始提交：加密Excel合并工具"

# 推送到 GitHub
git branch -M main
git push -u origin main
```

**💡 首次推送需要授权：**
- 浏览器会弹出 GitHub 登录页面
- 输入你的 GitHub 账号密码
- 授权访问

---

### 步骤 4：验证推送成功

1. 打开浏览器访问：`https://github.com/你的用户名/excel-merge-tool`
2. 你应该能看到所有代码文件
3. 点击 **Actions** 标签页
4. 你会看到打包任务正在运行（黄色圆点）

---

### 步骤 5：等待打包完成

- ⏱️ **首次打包**：约 5-10 分钟
- 📦 **打包内容**：
  - ✅ Windows EXE 文件
  - ✅ macOS APP 文件

**查看进度：**
1. 点击 **Actions** 标签
2. 点击正在运行的任务
3. 查看详细日志

---

### 步骤 6：下载生成的文件

打包完成后（绿色对勾 ✓）：

**方式一：从 Actions 下载**
1. 点击完成的运行记录
2. 滚动到页面底部 **Artifacts** 区域
3. 点击 `加密Excel合并工具-Windows` 下载 exe
4. 点击 `加密Excel合并工具-macOS` 下载 app

**方式二：从 Releases 下载（打标签时）**
1. 点击仓库页面的 **Releases**
2. 找到最新版本
3. 下载附件中的 exe 文件

---

## 🎯 后续使用

### 日常开发后自动打包

```bash
# 1. 修改代码
# ... 编辑 main.py 等文件 ...

# 2. 提交并推送
git add .
git commit -m "修改了XXX功能"
git push

# 3. GitHub 自动打包
# 等待 5-10 分钟，在 Actions 页面下载
```

### 发布正式版本（打标签）

```bash
# 创建版本标签
git tag v1.0.0
git push origin v1.0.0

# GitHub 会自动：
# - 打包 Windows exe 和 macOS app
# - 创建 GitHub Release
# - 附加生成的文件
```

---

## 🔧 常见问题

### Q1: 推送时要求登录怎么办？
**A:** 按照浏览器提示登录 GitHub 即可。或使用 GitHub CLI：
```bash
# 安装 GitHub CLI
brew install gh

# 登录
gh auth login

# 推送
git push
```

### Q2: Actions 没有自动运行？
**A:** 检查：
1. 仓库页面 → Settings → Actions → General
2. 确保选择 `Allow all actions and reusable workflows`
3. 重新推送一次代码

### Q3: 打包失败了怎么办？
**A:** 
1. 点击 Actions 中的失败任务
2. 查看错误日志
3. 常见错误：
   - 依赖安装失败 → 检查 requirements.txt
   - 导入错误 → 检查代码是否有语法错误
   - 内存不足 → 联系 GitHub 支持

### Q4: 如何只打包 Windows 不打包 macOS？
**A:** 编辑 `.github/workflows/build.yml`，删除 `build-macos` 部分。

### Q5: 打包的 exe 文件在哪？
**A:** 三种获取方式：
1. Actions → 点击运行记录 → 底部 Artifacts 下载
2. Releases → 选择版本 → 下载附件
3. 本地 Windows 电脑上运行 `build_windows.bat`

---

## 📊 GitHub Actions 免费额度

| 项目 | 额度 |
|------|------|
| 公共仓库 | 无限次使用 |
| 私有仓库 | 每月 2000 分钟 |
| 存储空间 | 500 MB |
| Artifact 保留 | 90 天 |

**你的项目预计使用：**
- 每次打包：约 5-8 分钟
- 每月打包 10 次：约 50-80 分钟
- **完全在免费额度内** ✅

---

## 💡 高级用法

### 手动触发包

1. 进入仓库 → Actions
2. 点击左侧 `Build Windows EXE`
3. 点击右侧 `Run workflow` 按钮
4. 选择分支 → 点击 `Run workflow`

### 自定义打包配置

编辑 `.github/workflows/build.yml`：

```yaml
# 修改 Python 版本
python-version: '3.11'  # 改为 3.11

# 添加应用图标
- name: Build EXE
  run: |
    pyinstaller ... --icon=app.ico ...

# 包含额外文件
- name: Build EXE
  run: |
    pyinstaller ... --add-data "config.json;." ...
```

---

## 📞 需要帮助？

如果遇到问题：

1. **查看日志**：Actions → 失败任务 → 查看错误信息
2. **检查配置**：确保仓库地址正确
3. **重新推送**：有时网络问题导致推送失败
4. **联系支持**：GitHub Community 论坛

---

## ✅ 验证清单

完成后确认：

- [ ] 代码已推送到 GitHub
- [ ] Actions 页面显示打包任务
- [ ] 打包任务成功（绿色对勾）
- [ ] 已下载 Windows exe 文件
- [ ] 在 Windows 电脑上测试 exe 可以运行

---

## 🎉 完成！

配置完成后，你将拥有：
- ✅ 自动化的 Windows EXE 打包流程
- ✅ 自动化的 macOS APP 打包流程
- ✅ 版本管理和发布系统
- ✅ 无需 Windows 电脑即可打包
- ✅ 完全免费

**下次修改代码只需：**
```bash
git add .
git commit -m "描述修改"
git push
# 等待 5-10 分钟，下载新的 exe
```
