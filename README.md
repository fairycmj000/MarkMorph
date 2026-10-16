# MarkMorph（格式转换工具）

一款基于 Electron 的桌面工具，把 Markdown（`.md`）一键转换为 **PDF / HTML / Word**，也能把 Word（`.docx`）无损转换为 **Markdown / HTML**，并支持**多文件分屏预览与合并导出**。便捷美观，无需联网。

## 功能

- **MD → PDF**：先预览排版，再下载 PDF，自动生成书签目录
- **MD → HTML**：转换为网页并在浏览器中打开
- **MD → Word**：转换为 `.docx`，保留表格、标题、加粗/斜体、换行等格式
- **Word → MD**：把 `.docx` 无损转换为 Markdown，标题/表格/嵌套表格/图片/超链接完整保留，图片提取到 `assets/` 子文件夹，嵌入的 Excel 电子表格渲染为表格
- **Word → HTML**：把 `.docx` 无损转换为 HTML 网页，图片内联、单文件自包含，转换范围与 Word → MD 一致
- **多文件分屏预览 + 合并下载**：多个 Markdown 并列分屏预览（可编辑保存 / 取消、拖动排序、关闭分片、布局切换 1×8 ↔ 2×4 并拖动分隔线），合并下载支持 **HTML / PDF / Word**（HTML 为标签页 + 汇总目录 + 阅读进度条，PDF/Word 按顺序从上到下拼接）
- **预览内编辑并保存 / 取消**：单文件与多文件分屏均可在预览内编辑并保存回源 `.md` 文件，或点「取消」放弃修改；编辑时可直接粘贴或选择本地图片插入（base64 内联，PDF/HTML/Word 均可显示）
- **HTML 查看增强**：目录导航栏可折叠/拖宽、「编辑」同行缩放按钮（50%–300%）、图片/流程图全屏放大查看并可单独下载
- **明亮 / 黑暗主题 + 全屏**：预览面板支持 🌙/☀️ 主题切换与全屏模式（全屏隐藏上传入口）
- **预览实时同步 + 冲突保护**：预览打开期间监听源文件，外部编辑保存后自动刷新；应用内正在编辑时检测到外部修改会提示并回退只读，不提交未保存的编辑
- **流程图 / 图片 / 表格**：mermaid 流程图内置引擎渲染，本地相对路径图片自动内嵌，宽表格自动 A3 横向适配
- **自定义背景图**：🎨 按钮选择 / 重置背景
- **`.md` 文件关联**：双击 `.md` 文件直接用 MarkMorph 以 HTML 预览（默认全屏）并可下载

## 技术栈

- Electron 35
- marked（Markdown 解析）、pdf-lib（PDF 书签）、docx（Word 生成）、jszip + fast-xml-parser（docx 解析，Word → MD/HTML）、sharp（流程图/图片处理）、mermaid（流程图渲染，本地化）

## 开发

```bash
npm install
npm start        # 本地运行
```

## 打包

```bash
npm run dist       # Windows（NSIS 安装包）
npm run dist:mac   # macOS（dmg，需在 Mac 上执行）
```

打包产物输出到 `dist/`，安装包归档到 `安装包/Windows` 与 `安装包/Mac`。

## 文档

- 完整安装与使用说明：[docs/安装使用说明.md](docs/安装使用说明.md)
- Mac 打包步骤：[packaging/mac-source/Mac打包说明.md](packaging/mac-source/Mac打包说明.md)
