# Mac 版打包说明

版本：**v1.0.20** | 应用名称：MarkMorph（格式转换工具）

麻烦在 **macOS 系统**上帮忙打包，步骤如下。

## 前置条件
- 一台 Mac（Intel 或 Apple Silicon 都行）
- 安装 Node.js 18 或更高版本（https://nodejs.org 下载 LTS 版）

## 打包步骤

1. 把本文件夹整个拷到 Mac 上，打开「终端」进入该目录：
   ```bash
   cd 本文件夹路径
   ```

2. 安装依赖（第一次需要，联网，几分钟）：
   ```bash
   npm install
   ```

3. 打包 Mac 版：
   ```bash
   npm run dist:mac
   ```

4. 打包完成后，成品在自动生成的 `dist/` 文件夹里，是一个 `.dmg` 文件
   （已配好 Intel x64 + Apple Silicon arm64 双架构，一个 dmg 通用）。

把 `dist/` 里的 `.dmg` 发回来即可。

## 说明
- 图标 `build/icon.icns` 已备好，Mac 打包会自动嵌入，无需额外操作。
- 如果 `npm run dist:mac` 报签名相关的错，可以先试着忽略签名打包：
  ```bash
  CSC_IDENTITY_AUTO_DISCOVERY=false npm run dist:mac
  ```
  （内部分享用，不做苹果公证也能装，首次打开时右键「打开」即可绕过 Gatekeeper。）
