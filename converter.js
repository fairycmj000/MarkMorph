const { marked } = require('marked');
const fs = require('fs');
const path = require('path');
const { pathToFileURL } = require('url');
const { Document, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle, PageOrientation, ImageRun, TableLayoutType } = require('docx');

// 将本地图片读取为 base64 data URI，使其在 data: 页面 / PDF / 离线 HTML 中都能显示
// 远程图片(http/https)和已内联的 data: 图片保持原样
function inlineImage(href, basedir) {
  try {
    if (!href) return href;
    if (/^(https?:|data:)/i.test(href)) return href;

    // 去掉 file:// 前缀并解码 URL 转义
    let p = href.replace(/^file:\/{0,3}/i, '');
    try { p = decodeURIComponent(p); } catch (e) {}

    // 相对路径 → 基于 markdown 所在目录解析
    if (!path.isAbsolute(p) && basedir) {
      p = path.resolve(basedir, p);
    }

    if (!fs.existsSync(p)) return href;

    const buf = fs.readFileSync(p);
    const ext = path.extname(p).toLowerCase().replace('.', '');
    const mimeMap = {
      jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png',
      gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp',
      svg: 'image/svg+xml', ico: 'image/x-icon'
    };
    const mime = mimeMap[ext] || 'image/png';
    return `data:${mime};base64,${buf.toString('base64')}`;
  } catch (e) {
    return href;
  }
}

// 把 data: URI 解码为 Buffer（用于 Word 内嵌粘贴/插入的图片）
function dataUrlToBuffer(dataUrl) {
  try {
    const m = /^data:([^;,]*)(;base64)?,([\s\S]*)$/.exec(dataUrl);
    if (!m) return null;
    return m[2] ? Buffer.from(m[3], 'base64') : Buffer.from(decodeURIComponent(m[3]), 'utf8');
  } catch (e) {
    return null;
  }
}

// 安全的日志函数，避免 EPIPE 错误
function safeLog(...args) {
  try {
    console.log(...args);
  } catch (e) {
    // 忽略 EPIPE 错误
  }
}

// 从图片 buffer 读取像素尺寸（支持 PNG / JPEG / GIF / BMP），无需外部依赖
function readImageSize(buf) {
  try {
    if (!buf || buf.length < 24) return null;
    // PNG: 89 50 4E 47
    if (buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4E && buf[3] === 0x47) {
      return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
    }
    // GIF: "GIF8"
    if (buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x38) {
      return { width: buf.readUInt16LE(6), height: buf.readUInt16LE(8) };
    }
    // BMP: "BM"
    if (buf[0] === 0x42 && buf[1] === 0x4D) {
      return { width: buf.readInt32LE(18), height: Math.abs(buf.readInt32LE(22)) };
    }
    // JPEG: FF D8
    if (buf[0] === 0xFF && buf[1] === 0xD8) {
      let off = 2;
      while (off < buf.length - 8) {
        if (buf[off] !== 0xFF) { off++; continue; }
        const marker = buf[off + 1];
        // SOF0..SOF15 (排除 C4/C8/CC)
        if (marker >= 0xC0 && marker <= 0xCF && marker !== 0xC4 && marker !== 0xC8 && marker !== 0xCC) {
          return { height: buf.readUInt16BE(off + 5), width: buf.readUInt16BE(off + 7) };
        }
        off += 2 + buf.readUInt16BE(off + 2);
      }
    }
  } catch (e) {}
  return null;
}

// 去掉 markdown 反斜杠转义，还原特殊字符（如 \_ → _，\( → (，\) → )）
function unescapeMarkdown(text) {
  return text.replace(/\\([!"#$%&'()*+,\-./:;<=>?@[\\\]^_`{|}~])/g, '$1');
}

function sanitizeMermaid(code) {
  return code
    .replace(/[\u{1F300}-\u{1FFFF}]/gu, '')
    .replace(/⚠️|✅|❌|⭐|🔴|🟡|🟢/g, '')
    .replace(/~/g, '-');
}

function extractHeadings(mdContent) {
  const headings = [];
  const tokens = marked.lexer(mdContent);
  for (const token of tokens) {
    if (token.type === 'heading') {
      headings.push({ level: token.depth, text: token.text });
    }
  }
  return headings;
}

// 把数组按 size 分块
function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

// 后处理：把连续的 <img> 分组为「图片行」，每行不超过 3 张，行内均分宽度（只缩不放）
function groupImageRows(html) {
  return html.replace(/<p>([\s\S]*?)<\/p>/g, (whole, inner) => {
    const imgs = inner.match(/<img\s[^>]*>/g);
    // 纯图片段落：整体按 ≤3 张分组成 flex 行
    if (imgs && imgs.length > 0 && /^(?:\s*<img\s[^>]*>\s*)+$/.test(inner)) {
      return chunk(imgs, 3)
        .map(row => `<div class="img-row" style="--n:${row.length}">${row.join('')}</div>`)
        .join('');
    }
    // 混合段落：连续 2 张及以上的图片序列也分组（单张保持原样）
    const mixed = inner.replace(/(?:<img\s[^>]*>\s*){2,}/g, run => {
      const tags = run.match(/<img\s[^>]*>/g);
      return chunk(tags, 3)
        .map(row => `<span class="img-row" style="--n:${row.length}">${row.join('')}</span>`)
        .join('');
    });
    return `<p>${mixed}</p>`;
  });
}

// 把单个 markdown 渲染成正文 HTML（内联图片 + mermaid 占位 + 图片行分组）
// 抽成独立函数，便于单文件预览与多文件合并共用同一套渲染逻辑
function renderBody(mdContent, basedir) {
  const renderer = new marked.Renderer();
  renderer.code = function({ text, lang }) {
    if (lang === 'mermaid') {
      return `<div class="mermaid">${sanitizeMermaid(text)}</div>`;
    }
    const escaped = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return `<pre><code class="language-${lang || ''}">${escaped}</code></pre>`;
  };
  // 将本地图片内联为 base64，确保在 data: 页面和离线 HTML 中正常显示
  renderer.image = function({ href, title, text }) {
    const src = inlineImage(href, basedir);
    const titleAttr = title ? ` title="${title.replace(/"/g, '&quot;')}"` : '';
    const altAttr = (text || '').replace(/"/g, '&quot;');
    return `<img src="${src}" alt="${altAttr}"${titleAttr}>`;
  };
  return groupImageRows(marked.parse(mdContent, { renderer }));
}

function buildHtml(mdContent, basedir, mermaidSrc) {
  const useA3 = hasWideTable(mdContent);
  // Mermaid 库本地化：默认加载随应用分发的本地文件，避免依赖 CDN（jsdelivr 在国内常被重置连接）
  const MERMAID_SRC = mermaidSrc || pathToFileURL(path.join(__dirname, 'renderer', 'mermaid.min.js')).href;
  const bodyHtml = renderBody(mdContent, basedir);
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    font-family: "Microsoft YaHei","PingFang SC",Arial,sans-serif;
    font-size: 14px;
    line-height: 1.7;
    color: #24292e;
    background: #ffffff;
  }
  /* 顶部阅读进度条 */
  #progress-bar {
    position: fixed;
    top: 0;
    left: 0;
    height: 3px;
    background: linear-gradient(90deg, #0366d6, #2188ff);
    z-index: 200;
    width: 0;
    transition: width 0.1s ease-out;
  }
  /* 左侧导航栏 */
  #sidebar {
    width: var(--sidebar-w, 280px);
    background: #f8f9fa;
    border-right: 1px solid #e1e4e8;
    height: 100vh;
    position: fixed;
    top: 0;
    left: 0;
    bottom: 0;
    display: flex;
    flex-direction: column;
    z-index: 100;
  }
  #sidebar-header {
    padding: 18px 20px 10px;
    border-bottom: 1px solid #e1e4e8;
    flex-shrink: 0;
  }
  #sidebar-header h2 {
    font-size: 12px;
    font-weight: 600;
    color: #6a737d;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin: 0 0 4px;
    padding: 0;
    border: none;
  }
  #progress-text { font-size: 11px; color: #8b949e; }
  #toc-wrapper {
    flex: 1;
    overflow-y: auto;
    padding: 10px 0 20px;
  }
  #toc { list-style: none; padding: 0; margin: 0; }
  #toc li { margin: 0; }
  .toc-row { display: flex; align-items: center; }
  .toc-toggle {
    flex: 0 0 18px;
    text-align: center;
    cursor: pointer;
    color: #8b949e;
    font-size: 11px;
    line-height: 1;
    user-select: none;
  }
  .toc-toggle:hover { color: #0366d6; }
  .toc-toggle.leaf { visibility: hidden; cursor: default; }
  .toc-children { list-style: none; padding: 0 0 0 16px; margin: 0; }
  .toc-item.collapsed > .toc-children { display: none; }
  #toc a {
    flex: 1;
    display: block;
    padding: 5px 10px 5px 4px;
    color: #586069;
    text-decoration: none;
    font-size: 12.5px;
    border-left: 3px solid transparent;
    transition: all 0.15s ease;
    line-height: 1.4;
  }
  #toc a:hover { background: #e6e9ec; color: #0366d6; }
  #toc a.active {
    color: #0366d6;
    font-weight: 600;
    background: #e6f0ff;
    border-left-color: #0366d6;
  }
  #toc .lvl-1 { font-weight: 700; color: #1a1a1a; font-size: 13px; }
  #toc .lvl-2 { font-weight: 600; color: #24292e; }
  #toc .lvl-3 { font-size: 12px; }
  #toc .lvl-4 { font-size: 11.5px; color: #6a737d; }
  #toc .lvl-5 { font-size: 11px; color: #8a939c; }

  /* 主内容区 */
  #main {
    margin-left: var(--sidebar-w, 280px);
    /* 自适应：撑满可视区域，避免宽屏右侧留白（打印时由 @media print 覆盖回 100%） */
    max-width: none;
  }
  #content {
    zoom: var(--zoom, 1);
    padding: 36px 50px;
    /* 用 JS 直接写入具体长度，避免 calc() 除以 var() 的兼容性坑（会导致 max-width 失效 → 放大后内容溢出到右侧留白） */
    max-width: var(--content-max-w, none);
  }
  h1, h2, h3, h4, h5, h6 {
    margin-top: 24px;
    margin-bottom: 12px;
    font-weight: 600;
    line-height: 1.3;
    color: #1a1a1a;
    scroll-margin-top: 20px;
  }
  h1 { font-size: 22px; border-bottom: 2px solid #333; padding-bottom: 8px; margin-top: 0; }
  h2 { font-size: 18px; border-bottom: 1px solid #ccc; padding-bottom: 4px; margin-top: 32px; }
  h3 { font-size: 15px; margin-top: 24px; }
  h4 { font-size: 13px; margin-top: 16px; }
  h5 { font-size: 12px; margin-top: 14px; }
  p { margin: 8px 0; }
  ul, ol { padding-left: 24px; }
  li { margin: 3px 0; }
  blockquote {
    border-left: 4px solid #0366d6;
    background: #f1f8ff;
    margin: 12px 0;
    padding: 8px 14px;
    color: #444;
  }
  table {
    border-collapse: collapse;
    width: 100%;
    margin: 12px 0;
    font-size: 12px;
  }
  th, td {
    border: 1px solid #ccc;
    padding: 6px 10px;
    text-align: left;
    vertical-align: top;
  }
  th { background: #f0f0f0; font-weight: bold; }
  tr:nth-child(even) td { background: #fafafa; }
  /* 表格字段筛选栏 */
  .table-filter-bar {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 6px 12px;
    margin: 4px 0 8px;
    padding: 8px 10px;
    background: #f6f8fa;
    border: 1px solid #e1e4e8;
    border-radius: 6px;
    font-size: 12px;
  }
  .table-filter-bar .filter-field { display: inline-flex; align-items: center; gap: 4px; white-space: nowrap; }
  .table-filter-bar .filter-field label { color: #586069; }
  .table-filter-bar .filter-multi { position: relative; display: inline-block; vertical-align: middle; }
  .table-filter-bar .filter-multi-trigger {
    display: inline-block;
    min-width: 90px;
    max-width: 180px;
    padding: 3px 6px;
    font-size: 12px;
    border: 1px solid #d1d5da;
    border-radius: 4px;
    background: #fff;
    color: #24292e;
    cursor: pointer;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .table-filter-bar .filter-multi-trigger::after { content: ' ▾'; color: #6a737d; }
  .table-filter-bar .filter-multi-panel {
    display: none;
    position: absolute;
    top: 100%;
    left: 0;
    z-index: 60;
    min-width: 150px;
    max-width: 280px;
    max-height: 240px;
    overflow: auto;
    margin-top: 2px;
    padding: 4px;
    background: #fff;
    border: 1px solid #d1d5da;
    border-radius: 4px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12);
    white-space: nowrap;
  }
  .table-filter-bar .filter-multi.open .filter-multi-panel { display: block; }
  .table-filter-bar .filter-opt {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 2px 4px;
    border-radius: 3px;
    cursor: pointer;
  }
  .table-filter-bar .filter-opt:hover { background: #f0f6ff; }
  .table-filter-bar .filter-opt input { margin: 0; }
  .table-filter-bar .filter-clear {
    display: block;
    width: 100%;
    margin: 0 0 4px;
    padding: 3px 0;
    font-size: 12px;
    text-align: center;
    border: none;
    border-bottom: 1px solid #eaecef;
    background: transparent;
    color: #0366d6;
    cursor: pointer;
  }
  .table-filter-bar .filter-clear:hover { background: #e6f0ff; }
  .table-filter-bar .filter-count { color: #6a737d; margin-left: auto; }
  .table-filter-bar button {
    padding: 3px 10px;
    font-size: 12px;
    border: 1px solid #d1d5da;
    border-radius: 4px;
    background: #fff;
    color: #24292e;
    cursor: pointer;
  }
  .table-filter-bar button:hover { background: #e6f0ff; border-color: #0366d6; color: #0366d6; }
  code {
    background: #f4f4f4;
    padding: 1px 4px;
    border-radius: 3px;
    font-family: Consolas, monospace;
    font-size: 12px;
  }
  pre { background: #f4f4f4; padding: 12px; border-radius: 4px; overflow: auto; }
  pre code { background: none; padding: 0; }
  hr { border: none; border-top: 1px solid #ddd; margin: 24px 0; }
  strong { color: #1a1a1a; }
  a { color: #0366d6; text-decoration: none; }
  a:hover { text-decoration: underline; }
  .mermaid { text-align: center; margin: 20px 0; }
  .mermaid svg { max-width: 100%; height: auto; }
  img { max-width: 100%; height: auto; display: block; margin: 12px 0; }
  /* 图片行：一组连续图片并排，每行不超过 3 张，只缩不放（小图保持原大小） */
  .img-row { display: flex; align-items: flex-start; gap: 8px; margin: 12px 0; page-break-inside: avoid; }
  .img-row img { flex: 0 1 auto; min-width: 0; max-width: calc(100% / var(--n, 3) - 8px); height: auto; margin: 0; }

  /* 正文图片可点击放大 */
  #main img { cursor: zoom-in; }

  /* 右上角正文缩放工具条 */
  #zoom-toolbar {
    position: fixed;
    top: 14px;
    right: 14px;
    z-index: 160;
    display: flex;
    align-items: center;
    gap: 2px;
    padding: 4px;
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid #e1e4e8;
    border-radius: 8px;
    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.12);
  }
  #zoom-toolbar button {
    border: none;
    background: transparent;
    cursor: pointer;
    min-width: 28px;
    height: 28px;
    border-radius: 5px;
    font-size: 15px;
    line-height: 1;
    color: #24292e;
  }
  #zoom-toolbar button:hover { background: #e6f0ff; color: #0366d6; }
  #zoom-level {
    min-width: 46px;
    text-align: center;
    font-size: 12px;
    color: #586069;
    font-variant-numeric: tabular-nums;
  }
  #zoom-reset { font-size: 12px; padding: 0 8px; margin-left: 5px; }
  #zoom-in, #zoom-out { font-weight: 700; font-size: 18px; min-width: 30px; height: 30px; }

  /* 导航栏折叠 / 展开 */
  #sidebar-header-top { display: flex; align-items: center; justify-content: space-between; }
  #sidebar-header-top h2 { margin-bottom: 0; }
  .toc-title { display: flex; align-items: center; gap: 4px; }
  #toc-toggle-all {
    border: none; background: transparent; color: #6a737d; cursor: pointer;
    font-size: 11px; line-height: 1; padding: 4px 6px; border-radius: 4px;
  }
  #toc-toggle-all:hover { background: #e6e9ec; color: #0366d6; }
  #collapse-btn {
    border: none;
    background: transparent;
    color: #6a737d;
    cursor: pointer;
    font-size: 16px;
    line-height: 1;
    padding: 4px 8px;
    border-radius: 4px;
  }
  #collapse-btn:hover { background: #e6e9ec; color: #0366d6; }
  #expand-btn {
    display: none;
    position: fixed;
    top: 14px;
    left: 14px;
    z-index: 160;
    width: 36px;
    height: 36px;
    border: 1px solid #e1e4e8;
    background: #fff;
    border-radius: 6px;
    cursor: pointer;
    font-size: 18px;
    color: #0366d6;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
  }
  body.collapsed #sidebar { display: none; }
  body.collapsed #sidebar-resizer { display: none; }
  body.collapsed #main { margin-left: 0; }
  body.collapsed #expand-btn { display: block; }

  /* 导航栏宽度拖动条 */
  #sidebar-resizer {
    position: fixed;
    top: 0;
    bottom: 0;
    left: calc(var(--sidebar-w, 280px) - 3px);
    width: 6px;
    cursor: col-resize;
    z-index: 101;
  }
  #sidebar-resizer:hover { background: rgba(3, 102, 214, 0.25); }

  /* 图片全屏 lightbox */
  #lightbox {
    position: fixed;
    inset: 0;
    z-index: 300;
    background: rgba(0, 0, 0, 0.9);
  }
  #lightbox.hidden { display: none; }
  #lightbox-stage {
    position: absolute;
    inset: 0;
    overflow: hidden;
    cursor: grab;
  }
  #lightbox-stage:active { cursor: grabbing; }
  #lightbox-img {
    position: absolute;
    left: 0;
    top: 0;
    transform-origin: 0 0;
    max-width: none;
    max-height: none;
    margin: 0;
    user-select: none;
    -webkit-user-drag: none;
  }
  #lightbox-close {
    position: absolute;
    top: 16px;
    right: 16px;
    z-index: 310;
    width: 40px;
    height: 40px;
    border: none;
    border-radius: 50%;
    background: rgba(255, 255, 255, 0.15);
    color: #fff;
    font-size: 22px;
    line-height: 1;
    cursor: pointer;
  }
  #lightbox-close:hover { background: rgba(255, 255, 255, 0.3); }
  #lightbox-controls {
    position: absolute;
    bottom: 20px;
    left: 50%;
    transform: translateX(-50%);
    z-index: 310;
    display: flex;
    align-items: center;
    gap: 2px;
    padding: 4px;
    background: rgba(255, 255, 255, 0.12);
    border-radius: 8px;
  }
  #lightbox-controls button {
    border: none;
    background: transparent;
    color: #fff;
    cursor: pointer;
    min-width: 32px;
    height: 32px;
    border-radius: 6px;
    font-size: 16px;
    line-height: 1;
  }
  #lightbox-controls button:hover { background: rgba(255, 255, 255, 0.25); }
  #lightbox-controls #lb-level {
    min-width: 48px;
    text-align: center;
    color: #fff;
    font-size: 12px;
    font-variant-numeric: tabular-nums;
  }
  #lightbox-controls #lb-reset { font-size: 12px; padding: 0 10px; }
  #lightbox-controls #lb-download { font-size: 12px; padding: 0 12px; }

  @media (max-width: 800px) {
    #sidebar, #sidebar-resizer, #expand-btn { display: none; }
    #main { margin-left: 0; max-width: 100%; }
    #content { padding: 20px; }
  }
  @page {
    size: ${useA3 ? 'A3' : 'A4'};
    margin: 20mm;
  }
  @media print {
    #sidebar, #progress-bar, #sidebar-resizer, #zoom-toolbar, #expand-btn, #collapse-btn, #lightbox { display: none !important; }
    #main { margin-left: 0 !important; max-width: 100% !important; }
    #content { zoom: 1 !important; padding: 20px !important; max-width: 100% !important; }
    .mermaid { page-break-inside: avoid; }
    table { page-break-inside: avoid; }

    /* 打印字号：屏幕字号（px）按 96DPI 1:1 打印会偏小（14px≈10.5pt），
       这里用 pt 放大到接近 Word 默认档位 */
    body { font-size: 12pt; }          /* 正文 12pt */
    h1 { font-size: 22pt; }
    h2 { font-size: 17pt; }
    h3 { font-size: 14pt; }
    h4 { font-size: 12.5pt; }
    table { font-size: 11pt; }         /* 表格 11pt */
    code { font-size: 10.5pt; }        /* 代码 10.5pt */
  }
</style>
</head>
<body>
<div id="progress-bar"></div>
<aside id="sidebar">
  <div id="sidebar-header">
    <div id="sidebar-header-top">
      <div class="toc-title">
        <h2>目录</h2>
        <button id="toc-toggle-all" title="全部收起" aria-label="全部收起">收起</button>
      </div>
      <button id="collapse-btn" title="收起导航栏" aria-label="收起导航栏">«</button>
    </div>
    <div id="progress-text">阅读进度 0%</div>
  </div>
  <div id="toc-wrapper">
    <ul id="toc"></ul>
  </div>
</aside>
<button id="expand-btn" title="展开导航栏" aria-label="展开导航栏">»</button>
<div id="sidebar-resizer"></div>
<div id="zoom-toolbar">
  <button id="zoom-out" title="缩小" aria-label="缩小">−</button>
  <span id="zoom-level">100%</span>
  <button id="zoom-in" title="放大" aria-label="放大">+</button>
  <button id="zoom-reset" title="重置缩放">重置</button>
</div>
<main id="main">
  <div id="content">
${bodyHtml}
  </div>
</main>
<div id="lightbox" class="hidden">
  <button id="lightbox-close" title="关闭" aria-label="关闭">×</button>
  <div id="lightbox-stage">
    <div id="lightbox-img"></div>
  </div>
  <div id="lightbox-controls">
    <button id="lb-zoom-out" title="缩小" aria-label="缩小">−</button>
    <span id="lb-level">100%</span>
    <button id="lb-zoom-in" title="放大" aria-label="放大">+</button>
    <button id="lb-reset" title="重置缩放">重置</button>
    <button id="lb-download" title="下载当前图片 / 流程图">下载</button>
  </div>
</div>
<script>
  // 生成目录 + 滚动追踪（独立运行，不依赖 mermaid）
  (function() {
    var toc = document.getElementById('toc');
    var tocWrapper = document.getElementById('toc-wrapper');
    var headings = Array.from(document.querySelectorAll('#main h1, #main h2, #main h3, #main h4, #main h5'));
    var tocItems = [];

    headings.forEach(function(h, i) {
      if (!h.id) h.id = 'heading-' + i;
    });

    // 滚动到目标标题：手动计算 scrollTop 以兼容正文 zoom 缩放（scrollIntoView 在 zoom 下会定位偏差）
    function scrollToTarget(el) {
      var scrollEl = document.scrollingElement || document.documentElement;
      var current = window.pageYOffset || scrollEl.scrollTop || document.documentElement.scrollTop || 0;
      var margin = 0;
      try { margin = parseFloat(getComputedStyle(el).scrollMarginTop) || 0; } catch (e) {}
      var top = el.getBoundingClientRect().top + current - margin;
      window.scrollTo({ top: top, behavior: 'smooth' });
    }

    // 按标题层级构建目录树（虚拟根 level=0），支持逐级折叠/展开
    function buildTocTree() {
      var stack = [{ level: 0, children: [] }];
      headings.forEach(function(h) {
        var level = parseInt(h.tagName.substring(1));
        var node = { level: level, heading: h, children: [], parent: null, li: null, toggle: null };
        while (stack.length > 1 && stack[stack.length - 1].level >= level) stack.pop();
        node.parent = stack[stack.length - 1];
        stack[stack.length - 1].children.push(node);
        stack.push(node);
      });
      return stack[0];
    }

    // 展开某节点到根，保证当前激活项始终可见
    function expandTo(node) {
      var n = node;
      while (n) {
        if (n.li && n.li.classList.contains('collapsed')) {
          n.li.classList.remove('collapsed');
          if (n.toggle) n.toggle.textContent = '▾';
        }
        n = n.parent;
      }
    }

    function renderToc(nodes, parentUl) {
      nodes.forEach(function(node) {
        var li = document.createElement('li');
        li.className = 'toc-item';
        node.li = li;

        var row = document.createElement('div');
        row.className = 'toc-row';

        var toggle = document.createElement('span');
        toggle.className = 'toc-toggle' + (node.children.length ? '' : ' leaf');
        toggle.textContent = node.children.length ? '▾' : '';
        node.toggle = toggle;
        if (node.children.length) {
          toggle.addEventListener('click', function() {
            var collapsed = li.classList.toggle('collapsed');
            toggle.textContent = collapsed ? '▸' : '▾';
          });
        }

        var a = document.createElement('a');
        a.href = '#' + node.heading.id;
        a.className = 'lvl-' + node.level;
        a.textContent = node.heading.textContent;
        a.dataset.target = node.heading.id;
        a.addEventListener('click', function(e) {
          e.preventDefault();
          scrollToTarget(node.heading);
          setTimeout(updateActive, 50);
        });

        row.appendChild(toggle);
        row.appendChild(a);
        li.appendChild(row);

        if (node.children.length) {
          var ul = document.createElement('ul');
          ul.className = 'toc-children';
          renderToc(node.children, ul);
          li.appendChild(ul);
        }

        parentUl.appendChild(li);
        tocItems.push({ heading: node.heading, link: a, node: node });
      });
    }

    // 渲染前清空已有目录：下载后的 HTML 已把目录内联，再次打开时脚本会再跑一遍，
    // 若不先清空会重复追加导致目录翻倍
    toc.innerHTML = '';
    renderToc(buildTocTree().children, toc);

    // 全部展开 / 全部收起（单按钮切换，位于「目录」右侧）
    var tocToggleAll = document.getElementById('toc-toggle-all');
    function setAllToc(collapsed) {
      tocItems.forEach(function(item) {
        var node = item.node;
        if (!node.children.length) return;   // 叶子节点无需展开/收起
        if (collapsed) {
          node.li.classList.add('collapsed');
          node.toggle.textContent = '▸';
        } else {
          node.li.classList.remove('collapsed');
          node.toggle.textContent = '▾';
        }
      });
      if (tocToggleAll) {
        tocToggleAll.textContent = collapsed ? '展开' : '收起';
        tocToggleAll.title = collapsed ? '全部展开' : '全部收起';
        tocToggleAll.setAttribute('aria-label', collapsed ? '全部展开' : '全部收起');
      }
    }
    tocToggleAll.addEventListener('click', function() {
      setAllToc(tocToggleAll.textContent === '收起');
    });

    var progressBar = document.getElementById('progress-bar');
    var progressText = document.getElementById('progress-text');

    function getCurrentHeadingIndex() {
      var offset = 80;
      var idx = 0;
      for (var i = 0; i < headings.length; i++) {
        var top = headings[i].getBoundingClientRect().top;
        if (top - offset <= 0) idx = i; else break;
      }
      return idx;
    }

    function updateActive() {
      var scrollEl = document.scrollingElement || document.documentElement;
      var bodyH = document.body.scrollHeight || 0;
      var htmlH = document.documentElement.scrollHeight || 0;
      var fullHeight = Math.max(bodyH, htmlH, scrollEl.scrollHeight);
      var docHeight = fullHeight - window.innerHeight;
      var scrolled = window.pageYOffset || scrollEl.scrollTop || document.documentElement.scrollTop || 0;
      var pct = docHeight > 0 ? Math.min(100, Math.round((scrolled / docHeight) * 100)) : 0;
      progressBar.style.width = pct + '%';
      progressText.textContent = '阅读进度 ' + pct + '%';

      var currentIdx = getCurrentHeadingIndex();
      var activeLink = null;
      tocItems.forEach(function(item, i) {
        var isActive = i === currentIdx;
        item.link.classList.toggle('active', isActive);
        if (isActive) activeLink = item.link;
      });
      if (currentIdx >= 0 && tocItems[currentIdx]) expandTo(tocItems[currentIdx].node);

      if (activeLink) {
        var wrapperRect = tocWrapper.getBoundingClientRect();
        var linkRect = activeLink.getBoundingClientRect();
        var wrapperMid = wrapperRect.top + wrapperRect.height / 2;
        var linkMid = linkRect.top + linkRect.height / 2;
        var diff = linkMid - wrapperMid;
        if (Math.abs(diff) > wrapperRect.height * 0.3) {
          tocWrapper.scrollTop += diff;
        }
      }
    }

    var ticking = false;
    window.addEventListener('scroll', function() {
      if (!ticking) {
        window.requestAnimationFrame(function() {
          updateActive();
          ticking = false;
        });
        ticking = true;
      }
    }, { passive: true });
    window.addEventListener('resize', function() {
      updateActive();
    });

    // 延迟到首次布局完成后计算，确保 scrollHeight 准确
    window.requestAnimationFrame(function() {
      updateActive();
    });
  })();
</script>
<script>
  // ===== 正文缩放 / 导航栏宽度·折叠 / 图片 lightbox =====
  (function() {
    var main = document.getElementById('main');
    var content = document.getElementById('content');
    var body = document.body;

    // ---- 正文缩放（仅 #content，不影响导航栏）----
    var zoomPct = 100;
    var zoomLevelEl = document.getElementById('zoom-level');
    var zoomInIframe = (function(){ try { return window.parent !== window.self; } catch (e) { return false; } })();
    function naturalWidth() {
      // 可视内容区宽度：#main 已用 margin-left 扣除侧栏；异常时回退窗口宽度
      return (main && main.clientWidth) ? main.clientWidth : window.innerWidth;
    }
    function reportZoom() {
      if (zoomLevelEl) zoomLevelEl.textContent = zoomPct + '%';
      if (zoomInIframe) { try { window.parent.postMessage({ type: 'mm-zoom-level', pct: zoomPct }, '*'); } catch (e) {} }
    }
    function applyZoom() {
      var z = zoomPct / 100;
      content.style.setProperty('--zoom', String(z));
      // 缩放后内容渲染宽度 = 布局宽度 × z，因此布局宽度要设为 naturalWidth()/z，放大后正好撑满、不溢出到右侧留白
      content.style.setProperty('--content-max-w', (naturalWidth() / z) + 'px');
      reportZoom();
    }
    function setZoom(pct) {
      zoomPct = Math.min(300, Math.max(50, Math.round(pct / 10) * 10));
      applyZoom();
    }
    // 预览内缩放改由外部工具栏（同「编辑」一行的 − / + / 重置）通过 postMessage 控制；
    // 独立下载打开的 HTML 仍保留右上角缩放按钮。
    if (zoomInIframe) {
      var zt = document.getElementById('zoom-toolbar');
      if (zt) zt.style.display = 'none';
      window.addEventListener('message', function(e) {
        var d = e && e.data;
        if (!d) return;
        if (d.type === 'mm-zoom-set') setZoom(Number(d.pct) || 100);
        else if (d.type === 'mm-zoom-delta') setZoom(zoomPct + (Number(d.delta) || 0));
      });
    } else {
      document.getElementById('zoom-in').addEventListener('click', function() { setZoom(zoomPct + 10); });
      document.getElementById('zoom-out').addEventListener('click', function() { setZoom(zoomPct - 10); });
      document.getElementById('zoom-reset').addEventListener('click', function() { setZoom(100); });
    }
    applyZoom();
    window.addEventListener('resize', applyZoom);

    // ---- 导航栏宽度 / 折叠 ----
    var SIDEBAR_MIN = 160, SIDEBAR_MAX = 800;
    var sidebarW = 280, collapsed = false;
    var resizer = document.getElementById('sidebar-resizer');

    function applySidebar() {
      document.documentElement.style.setProperty('--sidebar-w', sidebarW + 'px');
      body.classList.toggle('collapsed', collapsed);
      applyZoom();
    }
    function setCollapsed(c) { collapsed = c; applySidebar(); }
    document.getElementById('collapse-btn').addEventListener('click', function() { setCollapsed(true); });
    document.getElementById('expand-btn').addEventListener('click', function() { setCollapsed(false); });

    resizer.addEventListener('mousedown', function(e) {
      e.preventDefault();
      var startX = e.clientX, startW = sidebarW;
      function onMove(ev) {
        var maxW = Math.min(SIDEBAR_MAX, window.innerWidth * 0.6);
        sidebarW = Math.min(maxW, Math.max(SIDEBAR_MIN, startW + (ev.clientX - startX)));
        document.documentElement.style.setProperty('--sidebar-w', sidebarW + 'px');
        applyZoom();
      }
      function onUp() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        document.body.style.cursor = '';
      }
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
      document.body.style.cursor = 'col-resize';
    });
    applySidebar();

    // ---- 图片 lightbox ----
    var lightbox = document.getElementById('lightbox');
    var lbImg = document.getElementById('lightbox-img');
    var lbLevel = document.getElementById('lb-level');
    var stage = document.getElementById('lightbox-stage');
    var lb = { base: 1, k: 1, tx: 0, ty: 0, iw: 1, ih: 1, downloadHref: null, downloadSvg: null };
    var didDrag = false;

    function lbApply() {
      lbImg.style.transform = 'translate(' + lb.tx + 'px,' + lb.ty + 'px) scale(' + (lb.base * lb.k) + ')';
      lbLevel.textContent = Math.round(lb.k * 100) + '%';
    }
    function lbReset() {
      var sw = stage.clientWidth, sh = stage.clientHeight;
      lb.base = Math.min(sw / lb.iw, sh / lb.ih, 1);
      lb.k = 1;
      lb.tx = (sw - lb.iw * lb.base) / 2;
      lb.ty = (sh - lb.ih * lb.base) / 2;
      lbApply();
    }
    function lbShow() {
      lightbox.classList.remove('hidden');
      document.body.style.overflow = 'hidden';
    }
    function lbOpenImage(img) {
      var im = document.createElement('img');
      im.src = img.src;
      lb.downloadHref = img.src;
      lb.downloadSvg = null;
      im.style.maxWidth = 'none';
      im.style.maxHeight = 'none';
      lbImg.innerHTML = '';
      lbImg.appendChild(im);
      lbShow();
      if (im.complete && im.naturalWidth) {
        lb.iw = im.naturalWidth; lb.ih = im.naturalHeight; lbReset();
      } else {
        im.onload = function() {
          lb.iw = im.naturalWidth || im.width || 1;
          lb.ih = im.naturalHeight || im.height || 1;
          lbReset();
        };
      }
    }
    function lbOpenSvg(svg) {
      var rect = svg.getBoundingClientRect();
      var vb = svg.viewBox && svg.viewBox.baseVal ? svg.viewBox.baseVal : null;
      var w = rect.width || (vb && vb.width) || (svg.width && svg.width.baseVal && svg.width.baseVal.value) || 600;
      var h = rect.height || (vb && vb.height) || (svg.height && svg.height.baseVal && svg.height.baseVal.value) || 400;
      var clone = svg.cloneNode(true);
      clone.style.maxWidth = 'none';
      clone.style.maxHeight = 'none';
      clone.style.width = w + 'px';
      clone.style.height = h + 'px';
      // 放大查看时铺白底，避免黑色 lightbox 背景下灰色线条难以辨认
      clone.style.background = '#ffffff';
      lb.downloadSvg = new XMLSerializer().serializeToString(clone);
      lb.downloadHref = null;
      lbImg.innerHTML = '';
      lbImg.appendChild(clone);
      lb.iw = w; lb.ih = h;
      lbShow();
      lbReset();
    }
    function lbClose() {
      lightbox.classList.add('hidden');
      document.body.style.overflow = '';
      lbImg.innerHTML = '';
    }
    function lbZoom(factor, cx, cy) {
      var newK = Math.min(10, Math.max(0.1, lb.k * factor));
      var ratio = newK / lb.k;
      lb.tx = cx - (cx - lb.tx) * ratio;
      lb.ty = cy - (cy - lb.ty) * ratio;
      lb.k = newK;
      lbApply();
    }
    function lbDownloadName(href) {
      var ext = 'png';
      if (href) {
        var m = /^data:image[/]([a-zA-Z0-9.+-]+)/.exec(href);
        if (m) ext = m[1].split('+')[0];
        else {
          var clean = href.split('?')[0];
          var base = clean.split('/').pop();
          var em = /[.]([a-zA-Z0-9]+)$/.exec(base);
          if (em) ext = em[1];
        }
      }
      return '图片.' + ext;
    }
    function lbDownload() {
      var a = document.createElement('a');
      a.style.display = 'none';
      document.body.appendChild(a);
      if (lb.downloadSvg != null) {
        var blob = new Blob([lb.downloadSvg], { type: 'image/svg+xml;charset=utf-8' });
        a.href = URL.createObjectURL(blob);
        a.download = '流程图.svg';
        a.click();
        setTimeout(function() { URL.revokeObjectURL(a.href); }, 1000);
      } else if (lb.downloadHref) {
        a.href = lb.downloadHref;
        a.download = lbDownloadName(lb.downloadHref);
        a.click();
      }
      a.remove();
    }

    document.querySelectorAll('#main img').forEach(function(img) {
      img.addEventListener('click', function() { lbOpenImage(img); });
    });
    document.querySelectorAll('#main .mermaid').forEach(function(m) {
      m.style.cursor = 'zoom-in';
      m.addEventListener('click', function(e) {
        var svg = m.querySelector('svg');
        if (!svg) return;
        e.stopPropagation();
        lbOpenSvg(svg);
      });
    });

    document.getElementById('lightbox-close').addEventListener('click', lbClose);
    document.getElementById('lb-download').addEventListener('click', lbDownload);
    document.getElementById('lb-zoom-in').addEventListener('click', function() {
      lbZoom(1.25, stage.clientWidth / 2, stage.clientHeight / 2);
    });
    document.getElementById('lb-zoom-out').addEventListener('click', function() {
      lbZoom(0.8, stage.clientWidth / 2, stage.clientHeight / 2);
    });
    document.getElementById('lb-reset').addEventListener('click', lbReset);

    // 点击空白处关闭（拖拽后忽略一次点击，避免误关）
    lightbox.addEventListener('click', function(e) {
      if (didDrag) { didDrag = false; return; }
      if (e.target === lightbox || e.target === stage) lbClose();
    });

    // 滚轮缩放（围绕光标位置）
    stage.addEventListener('wheel', function(e) {
      e.preventDefault();
      var rect = stage.getBoundingClientRect();
      lbZoom(e.deltaY < 0 ? 1.2 : 0.8333, e.clientX - rect.left, e.clientY - rect.top);
    }, { passive: false });

    // 拖动平移（放大后）
    stage.addEventListener('mousedown', function(e) {
      if (lb.k <= 1) return;
      e.preventDefault();
      var sx = e.clientX, sy = e.clientY, ox = lb.tx, oy = lb.ty;
      didDrag = false;
      function onMove(ev) {
        if (Math.abs(ev.clientX - sx) > 3 || Math.abs(ev.clientY - sy) > 3) didDrag = true;
        lb.tx = ox + (ev.clientX - sx);
        lb.ty = oy + (ev.clientY - sy);
        lbApply();
      }
      function onUp() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
      }
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });

    // Esc 关闭
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape' && !lightbox.classList.contains('hidden')) lbClose();
    });
  })();
</script>
<script>
  // ===== 表格字段筛选：对 data-filterable 表格按字段多选筛选 =====
  (function() {
    // 点击筛选栏以外区域时收起下拉面板
    document.addEventListener('click', function(e) {
      var inside = e.target.closest ? e.target.closest('.filter-multi') : null;
      document.querySelectorAll('.filter-multi.open').forEach(function(m) {
        if (m !== inside) m.classList.remove('open');
      });
    });
    function initFilters() {
      document.querySelectorAll('table[data-filterable]').forEach(function(table) {
        if (table._mmFilterInit) return;
        table._mmFilterInit = true;
        var meta;
        try { meta = JSON.parse(table.getAttribute('data-filter-meta')); } catch (e) { return; }
        if (!meta || !meta.cols || !meta.cols.length) return;
        var cols = meta.cols;
        var labels = (meta.labels && meta.labels.length === cols.length) ? meta.labels : cols.map(function(c) { return '列' + (c + 1); });

        // 数据行 = 带首个筛选列 data-f 属性的 <tr>（表头/前置标题行不参与筛选）
        var firstKey = 'data-f' + cols[0];
        var dataRows = Array.prototype.filter.call(table.rows, function(tr) { return tr.hasAttribute(firstKey); });

        var optionSets = cols.map(function() { return {}; });
        dataRows.forEach(function(tr) {
          cols.forEach(function(col, i) {
            var v = (tr.getAttribute('data-f' + col) || '').trim();
            if (v) optionSets[i][v] = true;
          });
        });

        var bar = document.createElement('div');
        bar.className = 'table-filter-bar';

        var multis = [];
        cols.forEach(function(col, i) {
          var field = document.createElement('span');
          field.className = 'filter-field';
          var lab = document.createElement('label');
          lab.textContent = labels[i] + '：';

          var multi = document.createElement('span');
          multi.className = 'filter-multi';

          var trigger = document.createElement('span');
          trigger.className = 'filter-multi-trigger';
          trigger.textContent = '全部';

          var panel = document.createElement('span');
          panel.className = 'filter-multi-panel';

          var clearBtn = document.createElement('button');
          clearBtn.type = 'button';
          clearBtn.className = 'filter-clear';
          clearBtn.textContent = '清空';
          panel.appendChild(clearBtn);

          var checkboxes = [];
          var vals = Object.keys(optionSets[i]);
          vals.sort(function(a, b) {
            var na = Number(a), nb = Number(b);
            if (!isNaN(na) && !isNaN(nb)) return na - nb;
            return String(a).localeCompare(String(b), 'zh-CN');
          });
          vals.forEach(function(v) {
            var opt = document.createElement('label');
            opt.className = 'filter-opt';
            var cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = v;
            opt.appendChild(cb);
            opt.appendChild(document.createTextNode(' ' + v));
            panel.appendChild(opt);
            checkboxes.push(cb);
          });

          multi.appendChild(trigger);
          multi.appendChild(panel);
          field.appendChild(lab);
          field.appendChild(multi);
          bar.appendChild(field);
          multis.push({ col: col, multi: multi, trigger: trigger, panel: panel, checkboxes: checkboxes, clearBtn: clearBtn });
        });

        var countEl = document.createElement('span');
        countEl.className = 'filter-count';
        countEl.textContent = '共 ' + dataRows.length + ' 条';

        var resetBtn = document.createElement('button');
        resetBtn.type = 'button';
        resetBtn.textContent = '重置';

        bar.appendChild(resetBtn);
        bar.appendChild(countEl);

        function apply() {
          var shown = 0;
          dataRows.forEach(function(tr) {
            var match = true;
            for (var i = 0; i < multis.length; i++) {
              var m = multis[i];
              var activeVals = m.checkboxes.filter(function(cb) { return cb.checked; }).map(function(cb) { return cb.value; });
              if (!activeVals.length) continue;
              var v = (tr.getAttribute('data-f' + m.col) || '').trim();
              if (activeVals.indexOf(v) === -1) { match = false; break; }
            }
            tr.style.display = match ? '' : 'none';
            if (match) shown++;
          });
          countEl.textContent = '共 ' + shown + ' 条';
          multis.forEach(function(m) {
            var n = m.checkboxes.filter(function(cb) { return cb.checked; }).length;
            m.trigger.textContent = n === 0 ? '全部' : '已选 ' + n + ' 项';
          });
        }

        multis.forEach(function(m) {
          m.trigger.addEventListener('click', function(e) {
            e.stopPropagation();
            var isOpen = m.multi.classList.contains('open');
            document.querySelectorAll('.filter-multi.open').forEach(function(x) { x.classList.remove('open'); });
            if (!isOpen) m.multi.classList.add('open');
          });
          m.panel.addEventListener('click', function(e) { e.stopPropagation(); });
          m.checkboxes.forEach(function(cb) { cb.addEventListener('change', apply); });
          m.clearBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            m.checkboxes.forEach(function(cb) { cb.checked = false; });
            apply();
          });
        });
        resetBtn.addEventListener('click', function() {
          multis.forEach(function(m) {
            m.checkboxes.forEach(function(cb) { cb.checked = false; });
          });
          apply();
        });

        table.parentNode.insertBefore(bar, table);
      });
    }
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', initFilters);
    } else {
      initFilters();
    }
  })();
</script>
<script src="${MERMAID_SRC}"></script>
<script>
  try {
    if (typeof mermaid !== 'undefined') {
      mermaid.initialize({ startOnLoad: true, theme: 'default', flowchart: { useMaxWidth: true } });
    }
  } catch(e) {}
</script>
</body>
</html>`;
}

// 多个 markdown 文件合并为一个自包含 HTML：
// 顶部可拖拽排序、可重命名的标签页（一次只显示一个文件），
// 左侧目录按文件名分组汇总，支持跨文件跳转；文件名写入导出后的 HTML。
function buildMergedHtml(files, mermaidSrc, editable) {
  const list = (files || []).filter(f => f && typeof f.content === 'string');
  const useA3 = list.some(f => hasWideTable(f.content));
  const MERMAID_SRC = mermaidSrc || pathToFileURL(path.join(__dirname, 'renderer', 'mermaid.min.js')).href;

  function basedirOf(srcPath) {
    try { return srcPath ? path.dirname(srcPath) : ''; } catch (e) { return ''; }
  }

  const sectionsHtml = list.map((f, i) =>
    '<section class="doc-section' + (i === 0 ? ' active' : '') + '" id="doc-' + i + '" data-idx="' + i + '">\n' +
    '  <div class="doc-file-title"></div>\n' +
    '  <div class="doc-body">' + renderBody(f.content, basedirOf(f.srcPath)) + '</div>\n' +
    '</section>'
  ).join('\n');

  const docsJson = JSON.stringify(list.map(f => ({ name: f.name || '' }))).replace(/</g, '\\u003c');

  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    font-family: "Microsoft YaHei","PingFang SC",Arial,sans-serif;
    font-size: 14px;
    line-height: 1.7;
    color: #24292e;
    background: #ffffff;
  }

  /* 顶部阅读进度条 */
  #progress-bar {
    position: fixed;
    top: 0;
    left: 0;
    height: 3px;
    background: linear-gradient(90deg, #0366d6, #2188ff);
    z-index: 200;
    width: 0;
    transition: width 0.1s ease-out;
  }
  #progress-text { font-size: 11px; color: #8b949e; }

  /* 顶部标签栏（可拖拽排序、双击重命名） */
  #tab-bar {
    position: fixed;
    top: 0;
    left: var(--sidebar-w, 280px);
    right: 0;
    height: 44px;
    z-index: 150;
    display: flex;
    align-items: center;
    background: #fff;
    border-bottom: 1px solid #e1e4e8;
    overflow-x: auto;
    overflow-y: hidden;
  }
  #tabs { display: flex; align-items: center; gap: 4px; padding: 0 12px; }
  .tab {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 5px 12px;
    border: 1px solid #e1e4e8;
    border-radius: 8px;
    background: #f6f8fa;
    cursor: grab;
    font-size: 12.5px;
    color: #24292e;
    white-space: nowrap;
    user-select: none;
    flex-shrink: 0;
  }
  .tab:hover { border-color: #0366d6; }
  .tab.active { background: #0366d6; color: #fff; border-color: #0366d6; }
  .tab.dragging { opacity: 0.4; }
  .tab-rename-input {
    font-size: 12.5px;
    padding: 2px 6px;
    width: 140px;
    border: 1px solid #0366d6;
    border-radius: 4px;
    outline: none;
  }

  /* 左侧目录（按文件名分组汇总） */
  #sidebar {
    width: var(--sidebar-w, 280px);
    background: #f8f9fa;
    border-right: 1px solid #e1e4e8;
    height: 100vh;
    position: fixed;
    top: 0;
    left: 0;
    bottom: 0;
    display: flex;
    flex-direction: column;
    z-index: 100;
  }
  #sidebar-header {
    padding: 18px 20px 10px;
    border-bottom: 1px solid #e1e4e8;
    flex-shrink: 0;
  }
  #sidebar-header h2 {
    font-size: 12px;
    font-weight: 600;
    color: #6a737d;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin: 0;
    padding: 0;
    border: none;
  }
  #sidebar-header-top { display: flex; align-items: center; justify-content: space-between; }
  #sidebar-header-top h2 { margin-bottom: 0; }
  #toc-wrapper { flex: 1; overflow-y: auto; padding: 10px 0 20px; }
  #toc { list-style: none; padding: 0; margin: 0; }
  #toc .toc-group { margin: 0 0 6px; }
  #toc .toc-file {
    display: block;
    padding: 6px 20px 6px 14px;
    color: #1a1a1a;
    font-weight: 700;
    font-size: 13px;
    text-decoration: none;
    background: #eef1f4;
    border-left: 3px solid #0366d6;
    line-height: 1.4;
  }
  #toc .toc-file:hover { color: #0366d6; }
  #toc .toc-sub { list-style: none; padding: 0; margin: 0; }
  #toc .toc-sub a {
    display: block;
    padding: 4px 20px 4px 18px;
    color: #586069;
    text-decoration: none;
    font-size: 12px;
    border-left: 3px solid transparent;
    line-height: 1.4;
  }
  #toc .toc-sub a:hover { background: #e6e9ec; color: #0366d6; }
  #toc .toc-sub a.active {
    color: #0366d6;
    font-weight: 600;
    background: #e6f0ff;
    border-left-color: #0366d6;
  }
  #toc .toc-sub .lvl-1 { font-weight: 600; color: #1a1a1a; padding-left: 20px; }
  #toc .toc-sub .lvl-2 { padding-left: 26px; }
  #toc .toc-sub .lvl-3 { padding-left: 40px; font-size: 11.5px; }
  #toc .toc-sub .lvl-4 { padding-left: 54px; font-size: 11px; color: #6a737d; }
  #toc .toc-sub .lvl-5 { padding-left: 68px; font-size: 10.5px; color: #8a939c; }

  /* 主内容区 */
  #main {
    margin-left: var(--sidebar-w, 280px);
    /* 自适应：撑满可视区域，避免宽屏右侧留白（打印时由 @media print 覆盖回 100%） */
    max-width: none;
  }
  #content {
    zoom: var(--zoom, 1);
    padding: 64px 50px 36px;
    max-width: var(--content-max-w, none);
  }
  .doc-section { display: none; scroll-margin-top: 60px; }
  .doc-section.active { display: block; }
  .doc-file-title {
    font-size: 20px;
    font-weight: 700;
    color: #0366d6;
    border-bottom: 2px solid #0366d6;
    padding-bottom: 8px;
    margin-bottom: 16px;
  }

  h1, h2, h3, h4, h5, h6 {
    margin-top: 24px;
    margin-bottom: 12px;
    font-weight: 600;
    line-height: 1.3;
    color: #1a1a1a;
    scroll-margin-top: 64px;
  }
  h1 { font-size: 22px; border-bottom: 2px solid #333; padding-bottom: 8px; margin-top: 0; }
  h2 { font-size: 18px; border-bottom: 1px solid #ccc; padding-bottom: 4px; margin-top: 32px; }
  h3 { font-size: 15px; margin-top: 24px; }
  h4 { font-size: 13px; margin-top: 16px; }
  h5 { font-size: 12px; margin-top: 14px; }
  p { margin: 8px 0; }
  ul, ol { padding-left: 24px; }
  li { margin: 3px 0; }
  blockquote {
    border-left: 4px solid #0366d6;
    background: #f1f8ff;
    margin: 12px 0;
    padding: 8px 14px;
    color: #444;
  }
  table {
    border-collapse: collapse;
    width: 100%;
    margin: 12px 0;
    font-size: 12px;
  }
  th, td {
    border: 1px solid #ccc;
    padding: 6px 10px;
    text-align: left;
    vertical-align: top;
  }
  th { background: #f0f0f0; font-weight: bold; }
  tr:nth-child(even) td { background: #fafafa; }
  /* 表格字段筛选栏 */
  .table-filter-bar {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 6px 12px;
    margin: 4px 0 8px;
    padding: 8px 10px;
    background: #f6f8fa;
    border: 1px solid #e1e4e8;
    border-radius: 6px;
    font-size: 12px;
  }
  .table-filter-bar .filter-field { display: inline-flex; align-items: center; gap: 4px; white-space: nowrap; }
  .table-filter-bar .filter-field label { color: #586069; }
  .table-filter-bar .filter-multi { position: relative; display: inline-block; vertical-align: middle; }
  .table-filter-bar .filter-multi-trigger {
    display: inline-block;
    min-width: 90px;
    max-width: 180px;
    padding: 3px 6px;
    font-size: 12px;
    border: 1px solid #d1d5da;
    border-radius: 4px;
    background: #fff;
    color: #24292e;
    cursor: pointer;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .table-filter-bar .filter-multi-trigger::after { content: ' ▾'; color: #6a737d; }
  .table-filter-bar .filter-multi-panel {
    display: none;
    position: absolute;
    top: 100%;
    left: 0;
    z-index: 60;
    min-width: 150px;
    max-width: 280px;
    max-height: 240px;
    overflow: auto;
    margin-top: 2px;
    padding: 4px;
    background: #fff;
    border: 1px solid #d1d5da;
    border-radius: 4px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12);
    white-space: nowrap;
  }
  .table-filter-bar .filter-multi.open .filter-multi-panel { display: block; }
  .table-filter-bar .filter-opt {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 2px 4px;
    border-radius: 3px;
    cursor: pointer;
  }
  .table-filter-bar .filter-opt:hover { background: #f0f6ff; }
  .table-filter-bar .filter-opt input { margin: 0; }
  .table-filter-bar .filter-clear {
    display: block;
    width: 100%;
    margin: 0 0 4px;
    padding: 3px 0;
    font-size: 12px;
    text-align: center;
    border: none;
    border-bottom: 1px solid #eaecef;
    background: transparent;
    color: #0366d6;
    cursor: pointer;
  }
  .table-filter-bar .filter-clear:hover { background: #e6f0ff; }
  .table-filter-bar .filter-count { color: #6a737d; margin-left: auto; }
  .table-filter-bar button {
    padding: 3px 10px;
    font-size: 12px;
    border: 1px solid #d1d5da;
    border-radius: 4px;
    background: #fff;
    color: #24292e;
    cursor: pointer;
  }
  .table-filter-bar button:hover { background: #e6f0ff; border-color: #0366d6; color: #0366d6; }
  code {
    background: #f4f4f4;
    padding: 1px 4px;
    border-radius: 3px;
    font-family: Consolas, monospace;
    font-size: 12px;
  }
  pre { background: #f4f4f4; padding: 12px; border-radius: 4px; overflow: auto; }
  pre code { background: none; padding: 0; }
  hr { border: none; border-top: 1px solid #ddd; margin: 24px 0; }
  strong { color: #1a1a1a; }
  a { color: #0366d6; text-decoration: none; }
  a:hover { text-decoration: underline; }
  .mermaid { text-align: center; margin: 20px 0; }
  .mermaid svg { max-width: 100%; height: auto; }
  img { max-width: 100%; height: auto; display: block; margin: 12px 0; }
  .img-row { display: flex; align-items: flex-start; gap: 8px; margin: 12px 0; page-break-inside: avoid; }
  .img-row img { flex: 0 1 auto; min-width: 0; max-width: calc(100% / var(--n, 3) - 8px); height: auto; margin: 0; }
  #main img { cursor: zoom-in; }

  /* 右上角缩放工具条（下移到标签栏下方） */
  #zoom-toolbar {
    position: fixed;
    top: 54px;
    right: 14px;
    z-index: 160;
    display: flex;
    align-items: center;
    gap: 2px;
    padding: 4px;
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid #e1e4e8;
    border-radius: 8px;
    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.12);
  }
  #zoom-toolbar button {
    border: none;
    background: transparent;
    cursor: pointer;
    min-width: 28px;
    height: 28px;
    border-radius: 5px;
    font-size: 15px;
    line-height: 1;
    color: #24292e;
  }
  #zoom-toolbar button:hover { background: #e6f0ff; color: #0366d6; }
  #zoom-level { min-width: 46px; text-align: center; font-size: 12px; color: #586069; font-variant-numeric: tabular-nums; }
  #zoom-reset { font-size: 12px; padding: 0 8px; margin-left: 5px; }
  #zoom-in, #zoom-out { font-weight: 700; font-size: 18px; min-width: 30px; height: 30px; }

  #collapse-btn {
    border: none;
    background: transparent;
    color: #6a737d;
    cursor: pointer;
    font-size: 16px;
    line-height: 1;
    padding: 4px 8px;
    border-radius: 4px;
  }
  #collapse-btn:hover { background: #e6e9ec; color: #0366d6; }
  #expand-btn {
    display: none;
    position: fixed;
    top: 54px;
    left: 14px;
    z-index: 160;
    width: 36px;
    height: 36px;
    border: 1px solid #e1e4e8;
    background: #fff;
    border-radius: 6px;
    cursor: pointer;
    font-size: 18px;
    color: #0366d6;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
  }
  body.collapsed #sidebar { display: none; }
  body.collapsed #sidebar-resizer { display: none; }
  body.collapsed #tab-bar { left: 0; }
  body.collapsed #main { margin-left: 0; }
  body.collapsed #expand-btn { display: block; }

  #sidebar-resizer {
    position: fixed;
    top: 0;
    bottom: 0;
    left: calc(var(--sidebar-w, 280px) - 3px);
    width: 6px;
    cursor: col-resize;
    z-index: 101;
  }
  #sidebar-resizer:hover { background: rgba(3, 102, 214, 0.25); }

  /* 图片全屏 lightbox（与单文件一致） */
  #lightbox { position: fixed; inset: 0; z-index: 300; background: rgba(0, 0, 0, 0.9); }
  #lightbox.hidden { display: none; }
  #lightbox-stage { position: absolute; inset: 0; overflow: hidden; cursor: grab; }
  #lightbox-stage:active { cursor: grabbing; }
  #lightbox-img { position: absolute; left: 0; top: 0; transform-origin: 0 0; max-width: none; max-height: none; margin: 0; user-select: none; -webkit-user-drag: none; }
  #lightbox-close { position: absolute; top: 16px; right: 16px; z-index: 310; width: 40px; height: 40px; border: none; border-radius: 50%; background: rgba(255, 255, 255, 0.15); color: #fff; font-size: 22px; line-height: 1; cursor: pointer; }
  #lightbox-close:hover { background: rgba(255, 255, 255, 0.3); }
  #lightbox-controls { position: absolute; bottom: 20px; left: 50%; transform: translateX(-50%); z-index: 310; display: flex; align-items: center; gap: 2px; padding: 4px; background: rgba(255, 255, 255, 0.12); border-radius: 8px; }
  #lightbox-controls button { border: none; background: transparent; color: #fff; cursor: pointer; min-width: 32px; height: 32px; border-radius: 6px; font-size: 16px; line-height: 1; }
  #lightbox-controls button:hover { background: rgba(255, 255, 255, 0.25); }
  #lightbox-controls #lb-level { min-width: 48px; text-align: center; color: #fff; font-size: 12px; font-variant-numeric: tabular-nums; }
  #lightbox-controls #lb-reset { font-size: 12px; padding: 0 10px; }
  #lightbox-controls #lb-download { font-size: 12px; padding: 0 12px; }

  @media (max-width: 800px) {
    #sidebar, #sidebar-resizer, #expand-btn { display: none; }
    #tab-bar { left: 0; }
    #main { margin-left: 0; max-width: 100%; }
    #content { padding: 64px 20px 20px; }
  }
  @page {
    size: ${useA3 ? 'A3' : 'A4'};
    margin: 20mm;
  }
  @media print {
    #sidebar, #tab-bar, #sidebar-resizer, #zoom-toolbar, #expand-btn, #collapse-btn, #lightbox { display: none !important; }
    #main { margin-left: 0 !important; max-width: 100% !important; }
    #content { zoom: 1 !important; padding: 20px !important; max-width: 100% !important; }
    .doc-section { display: block !important; }
    .mermaid { page-break-inside: avoid; }
    table { page-break-inside: avoid; }
    body { font-size: 12pt; }
    h1 { font-size: 22pt; }
    h2 { font-size: 17pt; }
    h3 { font-size: 14pt; }
    h4 { font-size: 12.5pt; }
    table { font-size: 11pt; }
    code { font-size: 10.5pt; }
  }
</style>
</head>
<body>
<div id="progress-bar"></div>
<div id="tab-bar"><div id="tabs"></div></div>
<aside id="sidebar">
  <div id="sidebar-header">
    <div id="sidebar-header-top">
      <h2>目录（汇总所有文件）</h2>
      <button id="collapse-btn" title="收起导航栏" aria-label="收起导航栏">«</button>
    </div>
    <div id="progress-text">阅读进度 0%</div>
  </div>
  <div id="toc-wrapper"><ul id="toc"></ul></div>
</aside>
<button id="expand-btn" title="展开导航栏" aria-label="展开导航栏">»</button>
<div id="sidebar-resizer"></div>
<div id="zoom-toolbar">
  <button id="zoom-out" title="缩小" aria-label="缩小">−</button>
  <span id="zoom-level">100%</span>
  <button id="zoom-in" title="放大" aria-label="放大">+</button>
  <button id="zoom-reset" title="重置缩放">重置</button>
</div>
<main id="main">
  <div id="content">
${sectionsHtml}
  </div>
</main>
<div id="lightbox" class="hidden">
  <button id="lightbox-close" title="关闭" aria-label="关闭">×</button>
  <div id="lightbox-stage"><div id="lightbox-img"></div></div>
  <div id="lightbox-controls">
    <button id="lb-zoom-out" title="缩小" aria-label="缩小">−</button>
    <span id="lb-level">100%</span>
    <button id="lb-zoom-in" title="放大" aria-label="放大">+</button>
    <button id="lb-reset" title="重置缩放">重置</button>
    <button id="lb-download" title="下载当前图片 / 流程图">下载</button>
  </div>
</div>
<script>
  // ===== 标签页（拖拽排序 / 双击重命名）+ 汇总目录（按文件分组，跨文件跳转） =====
  (function() {
    var EDITABLE = ${editable ? 'true' : 'false'};
    var docs = ${docsJson};
    var docOrder = docs.map(function(_, i) { return i; });
    var active = 0;
    var dragFromPos = null;
    var tabsEl = document.getElementById('tabs');
    var tocEl = document.getElementById('toc');
    var sections = Array.from(document.querySelectorAll('.doc-section'));

    // 滚动到目标：手动计算 scrollTop 以兼容正文 zoom 缩放（scrollIntoView 在 zoom 下会定位偏差）
    function scrollToTarget(el) {
      var scrollEl = document.scrollingElement || document.documentElement;
      var current = window.pageYOffset || scrollEl.scrollTop || document.documentElement.scrollTop || 0;
      var margin = 0;
      try { margin = parseFloat(getComputedStyle(el).scrollMarginTop) || 0; } catch (e) {}
      var top = el.getBoundingClientRect().top + current - margin;
      window.scrollTo({ top: top, behavior: 'smooth' });
    }

    function switchTo(docIdx) {
      active = docIdx;
      sections.forEach(function(sec) {
        sec.classList.toggle('active', Number(sec.getAttribute('data-idx')) === docIdx);
      });
      updateTabActive();
      updateActiveToc();
      updateProgress();
    }

    // 切换时只改高亮、不重建 DOM，避免重渲染打断标签的「双击」检测
    function updateTabActive() {
      document.querySelectorAll('#tabs .tab').forEach(function(tab) {
        tab.classList.toggle('active', Number(tab.getAttribute('data-doc-idx')) === active);
      });
    }

    function renderTabs() {
      tabsEl.innerHTML = '';
      docOrder.forEach(function(docIdx) {
        var tab = document.createElement('div');
        tab.className = 'tab' + (docIdx === active ? ' active' : '');
        tab.setAttribute('data-doc-idx', docIdx);
        if (EDITABLE) {
          tab.setAttribute('draggable', 'true');
          tab.title = '双击重命名，拖拽排序';
        } else {
          tab.style.cursor = 'pointer';
        }
        var label = document.createElement('span');
        label.className = 'tab-label';
        label.textContent = docs[docIdx].name || ('文档 ' + (docIdx + 1));
        tab.appendChild(label);
        tab.addEventListener('click', function() { switchTo(docIdx); });
        if (EDITABLE) {
          tab.addEventListener('dblclick', function() { startRename(tab, docIdx); });
          tab.addEventListener('dragstart', function(e) {
            dragFromPos = docOrder.indexOf(docIdx);
            tab.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
          });
          tab.addEventListener('dragend', function() { tab.classList.remove('dragging'); dragFromPos = null; });
          tab.addEventListener('dragover', function(e) { e.preventDefault(); });
          tab.addEventListener('drop', function(e) {
            e.preventDefault();
            var from = dragFromPos;
            var to = docOrder.indexOf(docIdx);
            if (from == null || from === to) return;
            var moved = docOrder.splice(from, 1)[0];
            docOrder.splice(to, 0, moved);
            dragFromPos = null;
            renderTabs();
            buildToc();
          });
        }
        tabsEl.appendChild(tab);
      });
    }

    function startRename(tab, docIdx) {
      var label = tab.querySelector('.tab-label');
      if (!label) return;
      var oldName = docs[docIdx].name;
      var input = document.createElement('input');
      input.className = 'tab-rename-input';
      input.value = oldName;
      tab.replaceChild(input, label);
      input.focus();
      input.select();
      var committed = false;
      function commit() {
        if (committed) return;
        committed = true;
        var val = input.value.trim();
        if (!val) val = oldName;
        if (val !== oldName) {
          docs[docIdx].name = val;
          updateDocTitle(docIdx);
          buildToc();
        }
        renderTabs();
      }
      input.addEventListener('blur', commit);
      input.addEventListener('keydown', function(e) {
        if (e.key === 'Enter') { input.blur(); }
        else if (e.key === 'Escape') { input.value = oldName; input.blur(); }
      });
    }

    function updateDocTitle(docIdx) {
      var sec = sections[docIdx];
      if (sec) {
        var t = sec.querySelector('.doc-file-title');
        if (t) t.textContent = docs[docIdx].name;
      }
    }

    function buildToc() {
      tocEl.innerHTML = '';
      docOrder.forEach(function(docIdx) {
        var name = docs[docIdx].name || ('文档 ' + (docIdx + 1));
        var sec = sections[docIdx];
        var group = document.createElement('li');
        group.className = 'toc-group';

        var fileA = document.createElement('a');
        fileA.className = 'toc-file';
        fileA.href = '#';
        fileA.textContent = name;
        fileA.addEventListener('click', function(e) {
          e.preventDefault();
          switchTo(docIdx);
          scrollToTarget(sec);
        });
        group.appendChild(fileA);

        var sub = document.createElement('ul');
        sub.className = 'toc-sub';
        var headings = sec.querySelectorAll('.doc-body h1, .doc-body h2, .doc-body h3, .doc-body h4, .doc-body h5');
        headings.forEach(function(h, i) {
          if (!h.id) h.id = 'doc-' + docIdx + '-h-' + i;
          var level = parseInt(h.tagName.substring(1), 10);
          var li = document.createElement('li');
          var a = document.createElement('a');
          a.href = '#' + h.id;
          a.className = 'lvl-' + level;
          a.textContent = h.textContent;
          a.setAttribute('data-doc-idx', docIdx);
          a.setAttribute('data-target', h.id);
          a.addEventListener('click', function(e) {
            e.preventDefault();
            switchTo(docIdx);
            setTimeout(function() { scrollToTarget(h); }, 30);
          });
          li.appendChild(a);
          sub.appendChild(li);
        });
        if (headings.length > 0) group.appendChild(sub);
        tocEl.appendChild(group);
      });
    }

    function currentHeadings() {
      var sec = sections[active];
      return sec ? sec.querySelectorAll('.doc-body h1, .doc-body h2, .doc-body h3, .doc-body h4, .doc-body h5') : [];
    }

    function updateActiveToc() {
      var headings = currentHeadings();
      var idx = -1;
      var offset = 120;
      for (var i = 0; i < headings.length; i++) {
        if (headings[i].getBoundingClientRect().top - offset <= 0) idx = i; else break;
      }
      var activeTarget = (idx >= 0 && headings[idx]) ? headings[idx].id : null;
      document.querySelectorAll('#toc a[data-target]').forEach(function(a) {
        var isActive = Number(a.getAttribute('data-doc-idx')) === active && a.getAttribute('data-target') === activeTarget;
        a.classList.toggle('active', isActive);
      });
    }

    var progressBar = document.getElementById('progress-bar');
    var progressText = document.getElementById('progress-text');
    function updateProgress() {
      var scrollEl = document.scrollingElement || document.documentElement;
      var fullHeight = Math.max(document.body.scrollHeight, document.documentElement.scrollHeight, scrollEl.scrollHeight);
      var docHeight = fullHeight - window.innerHeight;
      var scrolled = window.pageYOffset || scrollEl.scrollTop || document.documentElement.scrollTop || 0;
      var pct = docHeight > 0 ? Math.min(100, Math.round((scrolled / docHeight) * 100)) : 0;
      progressBar.style.width = pct + '%';
      progressText.textContent = '阅读进度 ' + pct + '%';
    }

    var ticking = false;
    window.addEventListener('scroll', function() {
      if (!ticking) {
        window.requestAnimationFrame(function() { updateActiveToc(); updateProgress(); ticking = false; });
        ticking = true;
      }
    }, { passive: true });
    window.addEventListener('resize', function() { updateActiveToc(); updateProgress(); });

    docs.forEach(function(_, i) { updateDocTitle(i); });
    renderTabs();
    buildToc();
    window.requestAnimationFrame(function() { updateActiveToc(); updateProgress(); });
  })();
</script>
<script>
  // ===== 正文缩放 / 导航栏宽度·折叠 / 图片 lightbox（与单文件一致） =====
  (function() {
    var main = document.getElementById('main');
    var content = document.getElementById('content');
    var body = document.body;

    var zoomPct = 100;
    var zoomLevelEl = document.getElementById('zoom-level');
    var zoomInIframe = (function(){ try { return window.parent !== window.self; } catch (e) { return false; } })();
    function naturalWidth() {
      // 可视内容区宽度：#main 已用 margin-left 扣除侧栏；异常时回退窗口宽度
      return (main && main.clientWidth) ? main.clientWidth : window.innerWidth;
    }
    function reportZoom() {
      if (zoomLevelEl) zoomLevelEl.textContent = zoomPct + '%';
      if (zoomInIframe) { try { window.parent.postMessage({ type: 'mm-zoom-level', pct: zoomPct }, '*'); } catch (e) {} }
    }
    function applyZoom() {
      var z = zoomPct / 100;
      content.style.setProperty('--zoom', String(z));
      content.style.setProperty('--content-max-w', (naturalWidth() / z) + 'px');
      reportZoom();
    }
    function setZoom(pct) {
      zoomPct = Math.min(300, Math.max(50, Math.round(pct / 10) * 10));
      applyZoom();
    }
    // 预览内缩放改由外部工具栏（同「编辑」一行的 − / + / 重置）通过 postMessage 控制；
    // 独立下载打开的 HTML 仍保留右上角缩放按钮。
    if (zoomInIframe) {
      var zt = document.getElementById('zoom-toolbar');
      if (zt) zt.style.display = 'none';
      window.addEventListener('message', function(e) {
        var d = e && e.data;
        if (!d) return;
        if (d.type === 'mm-zoom-set') setZoom(Number(d.pct) || 100);
        else if (d.type === 'mm-zoom-delta') setZoom(zoomPct + (Number(d.delta) || 0));
      });
    } else {
      document.getElementById('zoom-in').addEventListener('click', function() { setZoom(zoomPct + 10); });
      document.getElementById('zoom-out').addEventListener('click', function() { setZoom(zoomPct - 10); });
      document.getElementById('zoom-reset').addEventListener('click', function() { setZoom(100); });
    }
    applyZoom();
    window.addEventListener('resize', applyZoom);

    var SIDEBAR_MIN = 160, SIDEBAR_MAX = 800;
    var sidebarW = 280, collapsed = false;
    var resizer = document.getElementById('sidebar-resizer');

    function applySidebar() {
      document.documentElement.style.setProperty('--sidebar-w', sidebarW + 'px');
      body.classList.toggle('collapsed', collapsed);
      applyZoom();
    }
    function setCollapsed(c) { collapsed = c; applySidebar(); }
    document.getElementById('collapse-btn').addEventListener('click', function() { setCollapsed(true); });
    document.getElementById('expand-btn').addEventListener('click', function() { setCollapsed(false); });

    resizer.addEventListener('mousedown', function(e) {
      e.preventDefault();
      var startX = e.clientX, startW = sidebarW;
      function onMove(ev) {
        var maxW = Math.min(SIDEBAR_MAX, window.innerWidth * 0.6);
        sidebarW = Math.min(maxW, Math.max(SIDEBAR_MIN, startW + (ev.clientX - startX)));
        document.documentElement.style.setProperty('--sidebar-w', sidebarW + 'px');
        applyZoom();
      }
      function onUp() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        document.body.style.cursor = '';
      }
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
      document.body.style.cursor = 'col-resize';
    });
    applySidebar();

    var lightbox = document.getElementById('lightbox');
    var lbImg = document.getElementById('lightbox-img');
    var lbLevel = document.getElementById('lb-level');
    var stage = document.getElementById('lightbox-stage');
    var lb = { base: 1, k: 1, tx: 0, ty: 0, iw: 1, ih: 1, downloadHref: null, downloadSvg: null };
    var didDrag = false;

    function lbApply() {
      lbImg.style.transform = 'translate(' + lb.tx + 'px,' + lb.ty + 'px) scale(' + (lb.base * lb.k) + ')';
      lbLevel.textContent = Math.round(lb.k * 100) + '%';
    }
    function lbReset() {
      var sw = stage.clientWidth, sh = stage.clientHeight;
      lb.base = Math.min(sw / lb.iw, sh / lb.ih, 1);
      lb.k = 1;
      lb.tx = (sw - lb.iw * lb.base) / 2;
      lb.ty = (sh - lb.ih * lb.base) / 2;
      lbApply();
    }
    function lbShow() {
      lightbox.classList.remove('hidden');
      document.body.style.overflow = 'hidden';
    }
    function lbOpenImage(img) {
      var im = document.createElement('img');
      im.src = img.src;
      lb.downloadHref = img.src;
      lb.downloadSvg = null;
      im.style.maxWidth = 'none';
      im.style.maxHeight = 'none';
      lbImg.innerHTML = '';
      lbImg.appendChild(im);
      lbShow();
      if (im.complete && im.naturalWidth) {
        lb.iw = im.naturalWidth; lb.ih = im.naturalHeight; lbReset();
      } else {
        im.onload = function() {
          lb.iw = im.naturalWidth || im.width || 1;
          lb.ih = im.naturalHeight || im.height || 1;
          lbReset();
        };
      }
    }
    function lbOpenSvg(svg) {
      var rect = svg.getBoundingClientRect();
      var vb = svg.viewBox && svg.viewBox.baseVal ? svg.viewBox.baseVal : null;
      var w = rect.width || (vb && vb.width) || (svg.width && svg.width.baseVal && svg.width.baseVal.value) || 600;
      var h = rect.height || (vb && vb.height) || (svg.height && svg.height.baseVal && svg.height.baseVal.value) || 400;
      var clone = svg.cloneNode(true);
      clone.style.maxWidth = 'none';
      clone.style.maxHeight = 'none';
      clone.style.width = w + 'px';
      clone.style.height = h + 'px';
      // 放大查看时铺白底，避免黑色 lightbox 背景下灰色线条难以辨认
      clone.style.background = '#ffffff';
      lb.downloadSvg = new XMLSerializer().serializeToString(clone);
      lb.downloadHref = null;
      lbImg.innerHTML = '';
      lbImg.appendChild(clone);
      lb.iw = w; lb.ih = h;
      lbShow();
      lbReset();
    }
    function lbClose() {
      lightbox.classList.add('hidden');
      document.body.style.overflow = '';
      lbImg.innerHTML = '';
    }
    function lbZoom(factor, cx, cy) {
      var newK = Math.min(10, Math.max(0.1, lb.k * factor));
      var ratio = newK / lb.k;
      lb.tx = cx - (cx - lb.tx) * ratio;
      lb.ty = cy - (cy - lb.ty) * ratio;
      lb.k = newK;
      lbApply();
    }
    function lbDownloadName(href) {
      var ext = 'png';
      if (href) {
        var m = /^data:image[/]([a-zA-Z0-9.+-]+)/.exec(href);
        if (m) ext = m[1].split('+')[0];
        else {
          var clean = href.split('?')[0];
          var base = clean.split('/').pop();
          var em = /[.]([a-zA-Z0-9]+)$/.exec(base);
          if (em) ext = em[1];
        }
      }
      return '图片.' + ext;
    }
    function lbDownload() {
      var a = document.createElement('a');
      a.style.display = 'none';
      document.body.appendChild(a);
      if (lb.downloadSvg != null) {
        var blob = new Blob([lb.downloadSvg], { type: 'image/svg+xml;charset=utf-8' });
        a.href = URL.createObjectURL(blob);
        a.download = '流程图.svg';
        a.click();
        setTimeout(function() { URL.revokeObjectURL(a.href); }, 1000);
      } else if (lb.downloadHref) {
        a.href = lb.downloadHref;
        a.download = lbDownloadName(lb.downloadHref);
        a.click();
      }
      a.remove();
    }

    document.querySelectorAll('#main img').forEach(function(img) {
      img.addEventListener('click', function() { lbOpenImage(img); });
    });
    document.querySelectorAll('#main .mermaid').forEach(function(m) {
      m.style.cursor = 'zoom-in';
      m.addEventListener('click', function(e) {
        var svg = m.querySelector('svg');
        if (!svg) return;
        e.stopPropagation();
        lbOpenSvg(svg);
      });
    });

    document.getElementById('lightbox-close').addEventListener('click', lbClose);
    document.getElementById('lb-download').addEventListener('click', lbDownload);
    document.getElementById('lb-zoom-in').addEventListener('click', function() { lbZoom(1.25, stage.clientWidth / 2, stage.clientHeight / 2); });
    document.getElementById('lb-zoom-out').addEventListener('click', function() { lbZoom(0.8, stage.clientWidth / 2, stage.clientHeight / 2); });
    document.getElementById('lb-reset').addEventListener('click', lbReset);

    lightbox.addEventListener('click', function(e) {
      if (didDrag) { didDrag = false; return; }
      if (e.target === lightbox || e.target === stage) lbClose();
    });

    stage.addEventListener('wheel', function(e) {
      e.preventDefault();
      var rect = stage.getBoundingClientRect();
      lbZoom(e.deltaY < 0 ? 1.2 : 0.8333, e.clientX - rect.left, e.clientY - rect.top);
    }, { passive: false });

    stage.addEventListener('mousedown', function(e) {
      if (lb.k <= 1) return;
      e.preventDefault();
      var sx = e.clientX, sy = e.clientY, ox = lb.tx, oy = lb.ty;
      didDrag = false;
      function onMove(ev) {
        if (Math.abs(ev.clientX - sx) > 3 || Math.abs(ev.clientY - sy) > 3) didDrag = true;
        lb.tx = ox + (ev.clientX - sx);
        lb.ty = oy + (ev.clientY - sy);
        lbApply();
      }
      function onUp() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
      }
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });

    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape' && !lightbox.classList.contains('hidden')) lbClose();
    });
  })();
</script>
<script>
  // ===== 表格字段筛选：对 data-filterable 表格按字段多选筛选 =====
  (function() {
    // 点击筛选栏以外区域时收起下拉面板
    document.addEventListener('click', function(e) {
      var inside = e.target.closest ? e.target.closest('.filter-multi') : null;
      document.querySelectorAll('.filter-multi.open').forEach(function(m) {
        if (m !== inside) m.classList.remove('open');
      });
    });
    function initFilters() {
      document.querySelectorAll('table[data-filterable]').forEach(function(table) {
        if (table._mmFilterInit) return;
        table._mmFilterInit = true;
        var meta;
        try { meta = JSON.parse(table.getAttribute('data-filter-meta')); } catch (e) { return; }
        if (!meta || !meta.cols || !meta.cols.length) return;
        var cols = meta.cols;
        var labels = (meta.labels && meta.labels.length === cols.length) ? meta.labels : cols.map(function(c) { return '列' + (c + 1); });

        // 数据行 = 带首个筛选列 data-f 属性的 <tr>（表头/前置标题行不参与筛选）
        var firstKey = 'data-f' + cols[0];
        var dataRows = Array.prototype.filter.call(table.rows, function(tr) { return tr.hasAttribute(firstKey); });

        var optionSets = cols.map(function() { return {}; });
        dataRows.forEach(function(tr) {
          cols.forEach(function(col, i) {
            var v = (tr.getAttribute('data-f' + col) || '').trim();
            if (v) optionSets[i][v] = true;
          });
        });

        var bar = document.createElement('div');
        bar.className = 'table-filter-bar';

        var multis = [];
        cols.forEach(function(col, i) {
          var field = document.createElement('span');
          field.className = 'filter-field';
          var lab = document.createElement('label');
          lab.textContent = labels[i] + '：';

          var multi = document.createElement('span');
          multi.className = 'filter-multi';

          var trigger = document.createElement('span');
          trigger.className = 'filter-multi-trigger';
          trigger.textContent = '全部';

          var panel = document.createElement('span');
          panel.className = 'filter-multi-panel';

          var clearBtn = document.createElement('button');
          clearBtn.type = 'button';
          clearBtn.className = 'filter-clear';
          clearBtn.textContent = '清空';
          panel.appendChild(clearBtn);

          var checkboxes = [];
          var vals = Object.keys(optionSets[i]);
          vals.sort(function(a, b) {
            var na = Number(a), nb = Number(b);
            if (!isNaN(na) && !isNaN(nb)) return na - nb;
            return String(a).localeCompare(String(b), 'zh-CN');
          });
          vals.forEach(function(v) {
            var opt = document.createElement('label');
            opt.className = 'filter-opt';
            var cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = v;
            opt.appendChild(cb);
            opt.appendChild(document.createTextNode(' ' + v));
            panel.appendChild(opt);
            checkboxes.push(cb);
          });

          multi.appendChild(trigger);
          multi.appendChild(panel);
          field.appendChild(lab);
          field.appendChild(multi);
          bar.appendChild(field);
          multis.push({ col: col, multi: multi, trigger: trigger, panel: panel, checkboxes: checkboxes, clearBtn: clearBtn });
        });

        var countEl = document.createElement('span');
        countEl.className = 'filter-count';
        countEl.textContent = '共 ' + dataRows.length + ' 条';

        var resetBtn = document.createElement('button');
        resetBtn.type = 'button';
        resetBtn.textContent = '重置';

        bar.appendChild(resetBtn);
        bar.appendChild(countEl);

        function apply() {
          var shown = 0;
          dataRows.forEach(function(tr) {
            var match = true;
            for (var i = 0; i < multis.length; i++) {
              var m = multis[i];
              var activeVals = m.checkboxes.filter(function(cb) { return cb.checked; }).map(function(cb) { return cb.value; });
              if (!activeVals.length) continue;
              var v = (tr.getAttribute('data-f' + m.col) || '').trim();
              if (activeVals.indexOf(v) === -1) { match = false; break; }
            }
            tr.style.display = match ? '' : 'none';
            if (match) shown++;
          });
          countEl.textContent = '共 ' + shown + ' 条';
          multis.forEach(function(m) {
            var n = m.checkboxes.filter(function(cb) { return cb.checked; }).length;
            m.trigger.textContent = n === 0 ? '全部' : '已选 ' + n + ' 项';
          });
        }

        multis.forEach(function(m) {
          m.trigger.addEventListener('click', function(e) {
            e.stopPropagation();
            var isOpen = m.multi.classList.contains('open');
            document.querySelectorAll('.filter-multi.open').forEach(function(x) { x.classList.remove('open'); });
            if (!isOpen) m.multi.classList.add('open');
          });
          m.panel.addEventListener('click', function(e) { e.stopPropagation(); });
          m.checkboxes.forEach(function(cb) { cb.addEventListener('change', apply); });
          m.clearBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            m.checkboxes.forEach(function(cb) { cb.checked = false; });
            apply();
          });
        });
        resetBtn.addEventListener('click', function() {
          multis.forEach(function(m) {
            m.checkboxes.forEach(function(cb) { cb.checked = false; });
          });
          apply();
        });

        table.parentNode.insertBefore(bar, table);
      });
    }
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', initFilters);
    } else {
      initFilters();
    }
  })();
</script>
<script src="${MERMAID_SRC}"></script>
<script>
  try {
    if (typeof mermaid !== 'undefined') {
      mermaid.initialize({ startOnLoad: true, theme: 'default', flowchart: { useMaxWidth: true } });
    }
  } catch(e) {}
</script>
</body>
</html>`;
}

module.exports = { buildHtml, buildMergedHtml, extractHeadings, buildDocx, extractMermaidBlocks, hasWideTable };

// 提取 Markdown 中所有 mermaid 代码块（供主进程渲染成图片）
function extractMermaidBlocks(mdContent) {
  const blocks = [];
  const tokens = marked.lexer(mdContent);
  for (const token of tokens) {
    if (token.type === 'code' && token.lang === 'mermaid') {
      blocks.push(sanitizeMermaid(token.text));
    }
  }
  return blocks;
}

// 检查是否有超过4列的表格
function hasWideTable(mdContent, threshold = 4) {
  const tokens = marked.lexer(mdContent);
  for (const token of tokens) {
    if (token.type === 'table' && token.header && token.header.length > threshold) {
      return true;
    }
  }
  return false;
}

// 把段落文本拆成「文字 / 图片组」片段；连续图片（中间只有空白）算作一组
function extractImageSegments(text) {
  const segments = [];
  const imgRegex = /!\[[^\]]*\]\(([^)\s]+)(?:\s+[^)]*)?\)/g;
  let lastIndex = 0;
  let match;
  let pending = null; // { srcs: [] }

  while ((match = imgRegex.exec(text)) !== null) {
    const between = text.slice(lastIndex, match.index);
    if (pending) {
      if (/^\s*$/.test(between)) {
        pending.srcs.push(match[1]);
      } else {
        segments.push({ type: 'images', srcs: pending.srcs });
        pending = null;
        if (between.trim()) segments.push({ type: 'text', text: between });
        pending = { srcs: [match[1]] };
      }
    } else {
      if (between.trim()) segments.push({ type: 'text', text: between });
      pending = { srcs: [match[1]] };
    }
    lastIndex = match.index + match[0].length;
  }
  if (pending) segments.push({ type: 'images', srcs: pending.srcs });
  const tail = text.slice(lastIndex);
  if (tail.trim()) segments.push({ type: 'text', text: tail });
  return segments;
}

function buildDocx(mdContent, mermaidImages = [], basedir) {
  const tokens = marked.lexer(mdContent);
  const sections = [];
  let mermaidIndex = 0;

  // 读取本地图片为 buffer 及像素尺寸（用于 Word 内嵌）
  function loadDocxImage(href) {
    try {
      if (!href) return null;
      if (/^data:/i.test(href)) return dataUrlToBuffer(href);
      if (/^https?:/i.test(href)) return null;
      let p = href.replace(/^file:\/{0,3}/i, '');
      try { p = decodeURIComponent(p); } catch (e) {}
      if (!path.isAbsolute(p) && basedir) p = path.resolve(basedir, p);
      if (!fs.existsSync(p)) return null;
      return fs.readFileSync(p);
    } catch (e) {
      return null;
    }
  }

  // 生成一行图片：1 张铺满可用宽度；2~3 张用一行表格并排；只缩不放（小图保持原大小）
  function makeImageRow(buffers) {
    const availWidth = hasWideTable(mdContent) ? 740 : 540;
    const totalDxa = hasWideTable(mdContent) ? 12500 : 8700;
    const n = buffers.length;

    // 按最大宽度等比缩放（只缩不放：宽度不超过自然宽度与 maxPx 中较小者）
    function scaled(buf, maxPx) {
      const dims = readImageSize(buf);
      let w = dims ? dims.width : maxPx;
      let h = dims ? dims.height : Math.round(maxPx * 0.6);
      if (w > maxPx) { h = Math.round(h * maxPx / w); w = maxPx; }
      return { width: w, height: h };
    }

    if (n === 1) {
      const s = scaled(buffers[0], availWidth);
      return [new Paragraph({
        children: [new ImageRun({ data: buffers[0], transformation: { width: s.width, height: s.height } })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 }
      })];
    }

    const cellDxa = Math.floor(totalDxa / n);
    const slotPx = Math.floor(availWidth / n);
    const cells = buffers.map(buf => {
      const s = scaled(buf, slotPx);
      return new TableCell({
        children: [new Paragraph({
          children: [new ImageRun({ data: buf, transformation: { width: s.width, height: s.height } })],
          alignment: AlignmentType.CENTER
        })],
        width: { size: cellDxa, type: WidthType.DXA }
      });
    });
    return [new Table({
      rows: [new TableRow({ children: cells })],
      width: { size: totalDxa, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: cells.map(() => cellDxa)
    })];
  }

  // 将单元格文字解析为 Paragraph 数组（处理图片 + <br> 换行 + 行内格式）
  // colWidthDxa: 该列的 DXA 宽度，用于等比缩放图片
  // bold: 表头单元格传 true，文字全部加粗
  function makeCellParagraphs(cellText, colWidthDxa, bold = false) {
    const paragraphs = [];
    const imgRegex = /!\[[^\]]*\]\(([^)\s]+)(?:\s+[^)]*)?\)/g;
    let lastIndex = 0;
    let match;

    while ((match = imgRegex.exec(cellText)) !== null) {
      // 图片前的文字
      const before = cellText.slice(lastIndex, match.index).trim();
      if (before) {
        paragraphs.push(new Paragraph({ children: parseInline(before, bold), alignment: AlignmentType.LEFT }));
      }
      // 图片：按列宽的 90% 等比缩放
      const buffer = loadDocxImage(match[1]);
      if (buffer) {
        const maxWidth = Math.max(40, Math.floor((colWidthDxa / 20) * 0.9));
        const dims = readImageSize(buffer);
        let w = dims ? dims.width : maxWidth;
        let h = dims ? dims.height : Math.round(maxWidth * 0.6);
        if (w > maxWidth) { h = Math.round(h * (maxWidth / w)); w = maxWidth; }
        paragraphs.push(new Paragraph({
          children: [new ImageRun({ data: buffer, transformation: { width: w, height: h } })],
          alignment: AlignmentType.CENTER
        }));
      }
      lastIndex = match.index + match[0].length;
    }

    // 剩余文字
    const remaining = cellText.slice(lastIndex).trim();
    if (remaining) {
      paragraphs.push(new Paragraph({ children: parseInline(remaining, bold), alignment: AlignmentType.LEFT }));
    }

    // 单元格不能为空（docx 规范要求至少一个段落）
    if (paragraphs.length === 0) {
      paragraphs.push(new Paragraph({ text: '' }));
    }
    return paragraphs;
  }

  safeLog(`buildDocx called with ${mermaidImages.length} mermaid images`);

  // A3 尺寸: 宽 16838 twips × 高 23811 twips (297mm × 420mm)
  // A4 尺寸: 宽 11906 twips × 高 16838 twips (210mm × 297mm)
  const useA3 = hasWideTable(mdContent);
  const pageSize = useA3
    ? { width: 16838, height: 23811, orientation: PageOrientation.PORTRAIT }
    : { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT };

  for (const token of tokens) {
    if (token.type === 'heading') {
      const level = token.depth === 1 ? HeadingLevel.HEADING_1 :
                    token.depth === 2 ? HeadingLevel.HEADING_2 :
                    token.depth === 3 ? HeadingLevel.HEADING_3 :
                    token.depth === 4 ? HeadingLevel.HEADING_4 :
                    token.depth === 5 ? HeadingLevel.HEADING_5 : HeadingLevel.HEADING_6;
      sections.push(new Paragraph({
        children: parseInline(token.text),
        heading: level,
        spacing: { before: 240, after: 120 }
      }));
    } else if (token.type === 'paragraph') {
      // 把段落拆成「文字 / 图片组」，图片组按每行 ≤3 张处理
      const segments = extractImageSegments(token.text);
      for (const seg of segments) {
        if (seg.type === 'text') {
          const runs = parseInline(seg.text);
          sections.push(new Paragraph({ children: runs, spacing: { after: 120 } }));
        } else {
          const buffers = seg.srcs.map(loadDocxImage).filter(Boolean);
          for (let i = 0; i < buffers.length; i += 3) {
            sections.push(...makeImageRow(buffers.slice(i, i + 3)));
          }
        }
      }
    } else if (token.type === 'list') {
      for (const item of token.items) {
        const runs = parseInline(item.text);
        sections.push(new Paragraph({
          children: runs,
          bullet: { level: 0 },
          spacing: { after: 60 }
        }));
      }
    } else if (token.type === 'code') {
      // Mermaid 代码块 → 插入预渲染的图片
      if (token.lang === 'mermaid') {
        safeLog(`Processing mermaid block ${mermaidIndex}, available images: ${mermaidImages.length}`);
        const imgData = mermaidImages[mermaidIndex++];
        if (imgData && imgData.buffer) {
          safeLog(`Inserting mermaid image: ${imgData.width}x${imgData.height}, buffer size: ${imgData.buffer.length}`);

          // 按纸张可用宽度和高度同时限制，等比例缩放
          // A4: 可用宽约 540pt, 可用高约 720pt (去掉页边距)
          // A3: 可用宽约 740pt, 可用高约 1030pt
          const maxWidth = useA3 ? 740 : 540;
          const maxHeight = useA3 ? 1030 : 720;
          const ratio = imgData.height / imgData.width;

          // 先按宽度缩放
          let displayWidth = Math.min(imgData.width, maxWidth);
          let displayHeight = displayWidth * ratio;

          // 如果高度超出，再按高度重新缩放
          if (displayHeight > maxHeight) {
            displayHeight = maxHeight;
            displayWidth = displayHeight / ratio;
          }

          safeLog(`Mermaid display size: ${displayWidth}x${displayHeight} (maxW=${maxWidth}, maxH=${maxHeight})`);

          sections.push(new Paragraph({
            children: [new ImageRun({
              data: imgData.buffer,
              transformation: {
                width: Math.round(displayWidth),
                height: Math.round(displayHeight)
              }
            })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 200 }
          }));
          continue;
        } else {
          safeLog(`Mermaid image ${mermaidIndex - 1} is null or missing buffer, falling back to code`);
        }
        // 图片渲染失败，回退到源码显示
      }
      const lines = token.text.split('\n');
      lines.forEach((line, idx) => {
        sections.push(new Paragraph({
          children: [new TextRun({
            text: line || ' ',
            font: 'Consolas',
            size: 20
          })],
          shading: { type: 'clear', fill: 'F4F4F4' },
          border: {
            left: { style: BorderStyle.SINGLE, size: 12, color: 'CCCCCC', space: 4 }
          },
          indent: { left: 240 },
          spacing: {
            before: idx === 0 ? 120 : 0,
            after: idx === lines.length - 1 ? 120 : 0,
            line: 260
          }
        }));
      });
    } else if (token.type === 'blockquote') {
      const text = token.tokens.map(t => t.type === 'paragraph' ? t.text : '').join(' ');
      sections.push(new Paragraph({
        text: text,
        italics: true,
        indent: { left: 720 },
        spacing: { before: 120, after: 120 }
      }));
    } else if (token.type === 'table') {
      const rows = [];
      const numCols = token.header ? token.header.length : (token.rows[0] ? token.rows[0].length : 1);

      // 使用固定 DXA 单位设置列宽，更可靠
      // A4 可用宽度约 9000 DXA (11906 - 2*1080边距) / 20 * 15 ≈ 8721 DXA
      // A3 可用宽度约 13000 DXA (16838 - 2*1080边距) / 20 * 15 ≈ 12507 DXA
      const useA3 = hasWideTable(mdContent);
      const totalWidth = useA3 ? 12500 : 8700;
      const maxColWidth = Math.floor(totalWidth / 3); // 单列最大 1/3
      const minColWidth = Math.floor(totalWidth * 0.08); // 单列最小 8%

      // 统计每列最大内容长度
      const colContentLengths = new Array(numCols).fill(0);
      if (token.header) {
        token.header.forEach((cell, i) => {
          colContentLengths[i] = Math.max(colContentLengths[i], (cell.text || '').length);
        });
      }
      token.rows.forEach(row => {
        row.forEach((cell, i) => {
          if (i < numCols) {
            colContentLengths[i] = Math.max(colContentLengths[i], (cell.text || '').length);
          }
        });
      });

      // 按内容比例初始分配
      const totalLength = colContentLengths.reduce((sum, len) => sum + len, 0);
      let colWidths = colContentLengths.map(len =>
        totalLength > 0 ? Math.round((len / totalLength) * totalWidth) : Math.floor(totalWidth / numCols)
      );

      // 限制最大和最小宽度
      let exceeded = 0;
      let deficit = 0;

      colWidths = colWidths.map(w => {
        if (w > maxColWidth) {
          exceeded += (w - maxColWidth);
          return maxColWidth;
        } else if (w < minColWidth) {
          deficit += (minColWidth - w);
          return minColWidth;
        }
        return w;
      });

      // 平衡超出和不足
      const adjustable = colWidths.map((w, i) => (w > minColWidth && w < maxColWidth) ? i : -1).filter(i => i >= 0);
      const netAdjust = exceeded - deficit;

      if (netAdjust !== 0 && adjustable.length > 0) {
        const perColAdjust = Math.floor(netAdjust / adjustable.length);
        adjustable.forEach(i => {
          colWidths[i] = Math.max(minColWidth, Math.min(maxColWidth, colWidths[i] + perColAdjust));
        });
      }

      // 确保总和等于 totalWidth
      const currentTotal = colWidths.reduce((sum, w) => sum + w, 0);
      if (currentTotal !== totalWidth && numCols > 0) {
        colWidths[numCols - 1] = Math.max(minColWidth, colWidths[numCols - 1] + (totalWidth - currentTotal));
      }

      safeLog(`Table with ${numCols} cols, DXA widths:`, colWidths, `total: ${colWidths.reduce((s,w)=>s+w,0)}`);

      // Header row
      if (token.header && token.header.length > 0) {
        rows.push(new TableRow({
          children: token.header.map((cell, idx) => new TableCell({
            children: makeCellParagraphs(cell.text, colWidths[idx], true),
            shading: { fill: 'E8E8E8' },
            width: { size: colWidths[idx], type: WidthType.DXA }
          }))
        }));
      }
      // Body rows
      for (const row of token.rows) {
        rows.push(new TableRow({
          children: row.map((cell, idx) => new TableCell({
            children: makeCellParagraphs(cell.text, colWidths[idx]),
            width: { size: colWidths[idx], type: WidthType.DXA }
          }))
        }));
      }
      sections.push(new Table({
        rows,
        width: { size: totalWidth, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,
        columnWidths: colWidths
      }));
    } else if (token.type === 'hr') {
      sections.push(new Paragraph({
        text: '',
        border: { bottom: { color: '000000', space: 1, style: BorderStyle.SINGLE, size: 6 } },
        spacing: { before: 120, after: 120 }
      }));
    }
  }

  return new Document({
    sections: [{
      properties: {
        page: {
          size: pageSize,
          margin: { top: 1440, right: 1080, bottom: 1440, left: 1080 }
        }
      },
      children: sections
    }]
  });
}

function parseInline(text, bold = false) {
  // 先按 <br> 分割，段内换行用 TextRun break 实现
  const brParts = text.split(/<br\s*\/?>/i);
  const runs = [];
  brParts.forEach((part, partIdx) => {
    if (partIdx > 0) {
      runs.push(new TextRun({ break: 1 }));
    }
    if (!part) return;
    // 去掉 markdown 反斜杠转义（\_ → _，\( → ( 等）
    const unescaped = unescapeMarkdown(part);
    // 解析 **bold**, *italic*, `code`
    const regex = /(\*\*(.+?)\*\*)|(\*(.+?)\*)|(`(.+?)`)|([\s\S]+?(?=\*\*|\*|`|$))/g;
    let match;
    while ((match = regex.exec(unescaped)) !== null) {
      if (match[2]) {
        runs.push(new TextRun({ text: match[2], bold: true }));
      } else if (match[4]) {
        runs.push(new TextRun({ text: match[4], italics: true, bold }));
      } else if (match[6]) {
        runs.push(new TextRun({ text: match[6], font: 'Courier New', bold }));
      } else if (match[7]) {
        const cleaned = match[7].replace(/\*\*|\*|`/g, '');
        if (cleaned) runs.push(new TextRun({ text: cleaned, bold }));
      }
    }
  });
  return runs.length > 0 ? runs : [new TextRun({ text: unescapeMarkdown(text), bold })];
}
