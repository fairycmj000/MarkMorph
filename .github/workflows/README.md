# GitHub Actions 自动化打包说明

## 功能
此工作流可以自动在 Windows 和 macOS 系统上打包应用程序。

## 触发方式

### 1. 手动触发
在 GitHub 仓库页面：
1. 进入 `Actions` 标签页
2. 选择 `Build and Release` 工作流
3. 点击 `Run workflow` 按钮
4. 选择分支后点击绿色的 `Run workflow`

构建完成后，在工作流运行页面下载构建产物（Artifacts）。

### 2. 标签触发（自动发布）
当推送带有版本标签时自动触发并创建 Release：

```bash
git tag v1.0.1
git push origin v1.0.1
```

构建完成后会自动创建 GitHub Release，并附带 Windows 和 macOS 安装包。

## 使用步骤

1. **初始化 Git 仓库**（如果还没有）
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   ```

2. **创建 GitHub 仓库**
   - 在 GitHub 上创建新仓库
   - 不要添加 README、.gitignore 或 license

3. **推送代码**
   ```bash
   git remote add origin https://github.com/你的用户名/仓库名.git
   git branch -M main
   git push -u origin main
   ```

4. **触发构建**
   - 手动触发：在 GitHub Actions 页面点击 `Run workflow`
   - 或推送标签：`git tag v1.0.1 && git push origin v1.0.1`

## 输出文件
- **Windows**: `格式转换工具 Setup 1.0.1.exe`
- **macOS**: `格式转换工具-1.0.1.dmg` 和 `格式转换工具-1.0.1-arm64.dmg`

## 注意事项
- macOS 版本不会进行代码签名（需要 Apple Developer 证书）
- 用户首次打开可能需要在"系统偏好设置 > 安全性与隐私"中允许运行
- 如需代码签名，需要配置 GitHub Secrets 并修改工作流
