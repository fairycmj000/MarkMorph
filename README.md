# MarkMorph（格式转换工具）

一款基于 Electron 的桌面工具，把 Markdown（`.md`）文档一键转换为 **PDF / HTML / Word**，并支持**多文件合并导出**。便捷美观，无需联网。

## 功能

- **MD → PDF**：先预览排版，再下载 PDF，自动生成书签目录
- **MD → HTML**：转换为网页并在浏览器中打开
- **MD → Word**：转换为 `.docx`，保留表格、标题、加粗/斜体、换行等格式
- **多文件合并导出**：多个 Markdown 合并为一个 HTML（顶部标签页 + 汇总目录 + 阅读进度条），顺序可拖拽调整
- **HTML 查看增强**：目录导航栏可折叠/拖宽、仅内容区域缩放、图片/流程图全屏放大查看
- **预览实时同步**：预览打开期间监听源文件，外部编辑保存后自动刷新
- **流程图 / 图片 / 表格**：mermaid 流程图内置引擎渲染，本地相对路径图片自动内嵌，宽表格自动 A3 横向适配
- **自定义背景图**：🎨 按钮选择 / 重置背景
- **`.md` 文件关联**：双击 `.md` 文件直接用 MarkMorph 以 HTML 预览并可下载

## 技术栈

- Electron 35
- marked（Markdown 解析）、pdf-lib（PDF 书签）、docx（Word 生成）、sharp（流程图/图片处理）、mermaid（流程图渲染，本地化）

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
