# macOS (Apple Silicon) 打包指南

本指南将帮助你在 Mac mini (M4) 上打包微博热搜抓取工具。

## 1. 准备工作

由于 Windows 和 macOS 的可执行文件不兼容，你必须在 Mac 上进行打包。

### 1.1 安装 Python
确保 Mac 上安装了 Python 3.10 或更高版本。
推荐使用 Homebrew 安装：
```bash
brew install python
```

### 1.2 获取代码
将整个项目文件夹复制到你的 Mac 上。

## 2. 环境配置

打开终端 (Terminal)，进入项目目录：
```bash
cd /path/to/your/project
```
*(提示：可以在终端输入 `cd ` 然后把文件夹拖进去)*

### 2.1 创建并激活虚拟环境
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 2.2 安装依赖
```bash
pip install -r requirements.txt
pip install pyinstaller
```

### 2.3 安装 Playwright 浏览器
这一步非常关键，Mac 上的浏览器文件与 Windows 不同。
```bash
playwright install chromium
```

## 3. 执行打包

使用以下命令进行打包：

```bash
pyinstaller WeiboHotGUI.spec --noconfirm
```

## 4. 整理发布包

打包完成后，`dist/WeiboHotGUI` 文件夹就是你的程序。但还需要手动复制浏览器文件。

### 4.1 查找浏览器路径
在终端运行：
```bash
ls ~/Library/Caches/ms-playwright
```
你会看到类似 `chromium-xxxx` 和 `chromium_headless_shell-xxxx` 的文件夹。

### 4.2 复制浏览器
我们需要将这些浏览器复制到 `dist/WeiboHotGUI/browsers/` 下。

```bash
# 1. 创建目录
mkdir -p dist/WeiboHotGUI/browsers

# 2. 复制浏览器 (注意：请将下方命令中的 '*' 替换为实际的版本号，或者直接运行命令让它自动匹配)
cp -R ~/Library/Caches/ms-playwright/chromium-* dist/WeiboHotGUI/browsers/
cp -R ~/Library/Caches/ms-playwright/chromium_headless_shell-* dist/WeiboHotGUI/browsers/
# 可选：复制 ffmpeg
cp -R ~/Library/Caches/ms-playwright/ffmpeg-* dist/WeiboHotGUI/browsers/
```

### 4.3 复制说明文件
```bash
cp README.md dist/WeiboHotGUI/使用说明.md
```

## 5. 运行与常见问题

### 5.1 运行
在 Finder 中打开 `dist/WeiboHotGUI` 文件夹，双击 `WeiboHotGUI` 可执行文件。
或者在终端运行：
```bash
./dist/WeiboHotGUI/WeiboHotGUI
```

### 5.2 权限问题 ("文件已损坏" 或 "无法打开")
由于应用未签名，macOS 可能会阻止运行。
**解决方法：**
在终端执行以下命令（移除隔离属性）：
```bash
xattr -cr dist/WeiboHotGUI/WeiboHotGUI
```
或者在“系统设置” -> “隐私与安全性”中，找到被拦截的提示并点击“仍要打开”。

### 5.3 托盘图标
在 macOS 上，程序图标会出现在**顶部菜单栏的右侧**，而不是底部的 Dock 或 Windows 的右下角。

## 6. 验证
双击运行后，程序应能自动识别 `browsers` 目录下的浏览器并正常抓取。
