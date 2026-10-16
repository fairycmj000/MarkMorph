const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const { PDFDocument, PDFName, PDFArray, PDFDict, PDFHexString, PDFNull, PDFNumber } = require('pdf-lib');
const { buildHtml, buildMergedHtml, extractHeadings, buildDocx, extractMermaidBlocks, hasWideTable } = require('./converter');
const { Packer } = require('docx');
const sharp = require('sharp');
const { pathToFileURL } = require('url');

// Mermaid 库本地化：随应用分发的本地 mermaid.min.js，避免依赖 CDN
const MERMAID_SRC = pathToFileURL(path.join(__dirname, 'renderer', 'mermaid.min.js')).href;

let mainWindow = null;
let pendingOpenPath = null; // 双击 .md 关联打开时，待预览的文件绝对路径

// 从命令行参数中提取 markdown 文件路径（文件关联 / 拖到 exe 打开时由系统传入）
function extractMdPath(args) {
  for (const a of args || []) {
    if (!a || a.startsWith('-')) continue;
    if (/\.(md|markdown|mdown|mkd)$/i.test(a)) return a;
  }
  return null;
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1100,
    height: 700,
    minWidth: 600,
    minHeight: 500,
    title: 'MarkMorph',
    icon: path.join(__dirname, 'build', 'icon.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });
  const indexPath = path.join(__dirname, 'renderer', 'index.html');
  if (pendingOpenPath) {
    // 首次启动即带着文件：通过 query 把路径交给渲染层，自动进入 HTML 预览
    win.loadFile(indexPath, { query: { open: pendingOpenPath } });
    pendingOpenPath = null;
  } else {
    win.loadFile(indexPath);
  }
  win.setMenuBarVisibility(false);
  mainWindow = win;
}

// 单实例：文件关联双击打开时，第二个实例把文件路径转交给已存在的窗口
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on('second-instance', (event, argv) => {
    const filePath = extractMdPath(argv);
    if (filePath && mainWindow && !mainWindow.isDestroyed()) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
      mainWindow.webContents.send('open-file', filePath);
    }
  });

  // 兼容 open-file 事件（macOS 及部分 Windows 场景）
  app.on('open-file', (event, filePath) => {
    event.preventDefault();
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('open-file', filePath);
    } else {
      pendingOpenPath = filePath;
    }
  });

  pendingOpenPath = extractMdPath(process.argv.slice(1));

  app.whenReady().then(createWindow);
  app.on('window-all-closed', () => app.quit());
}

// 窗口布局切换
ipcMain.handle('resize-layout', async (_, mode) => {
  if (!mainWindow) return;
  const [w, h] = mainWindow.getSize();
  if (mode === 'portrait') {
    // 保持高度，收窄宽度到竖屏比例
    const newW = Math.max(500, Math.round(h * 0.7));
    mainWindow.setSize(newW, h);
  } else {
    // 保持宽度，降低高度到横屏比例
    const newH = Math.max(450, Math.round(w * 0.7));
    mainWindow.setSize(w, newH);
  }
  mainWindow.center();
});

// 背景图片配置路径
const bgConfigPath = path.join(app.getPath('userData'), 'bg-config.json');

function getBgConfig() {
  try {
    if (fs.existsSync(bgConfigPath)) {
      return JSON.parse(fs.readFileSync(bgConfigPath, 'utf-8'));
    }
  } catch (e) {}
  return null;
}

// 选择背景图片
ipcMain.handle('select-background', async () => {
  const { filePaths, canceled } = await dialog.showOpenDialog({
    title: '选择背景图片',
    filters: [
      { name: '图片', extensions: ['jpg', 'jpeg', 'png', 'webp', 'bmp', 'gif'] },
      { name: '所有文件', extensions: ['*'] }
    ],
    properties: ['openFile']
  });
  if (canceled || !filePaths || filePaths.length === 0) return null;

  const srcPath = filePaths[0];
  const ext = path.extname(srcPath);
  const destPath = path.join(app.getPath('userData'), 'custom-bg' + ext);
  fs.copyFileSync(srcPath, destPath);
  fs.writeFileSync(bgConfigPath, JSON.stringify({ path: destPath }), 'utf-8');
  return destPath;
});

// 获取当前背景图片
ipcMain.handle('get-background', async () => {
  const config = getBgConfig();
  if (config && config.path && fs.existsSync(config.path)) {
    return config.path;
  }
  return null;
});

// 重置背景图片
ipcMain.handle('reset-background', async () => {
  try {
    if (fs.existsSync(bgConfigPath)) {
      fs.unlinkSync(bgConfigPath);
    }
    // 删除自定义背景图片
    const userDataPath = app.getPath('userData');
    const files = fs.readdirSync(userDataPath);
    files.forEach(file => {
      if (file.startsWith('custom-bg')) {
        fs.unlinkSync(path.join(userDataPath, file));
      }
    });
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// 应用信息：版本 + 打包日期（水印自动显示最新版本的构建日期）
ipcMain.handle('get-app-info', async () => {
  let buildDate = '';
  try {
    // 打包后 app.getAppPath() 指向 app.asar，其修改时间即打包时间；开发态为源码目录
    const mtime = fs.statSync(app.getAppPath()).mtime;
    buildDate = '' + mtime.getFullYear() + String(mtime.getMonth() + 1).padStart(2, '0') + String(mtime.getDate()).padStart(2, '0');
  } catch (e) {
    buildDate = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  }
  return { version: app.getVersion(), buildDate };
});

// 从 markdown 文件路径推导所在目录（用于解析相对图片路径）
function basedirOf(srcPath) {
  try {
    return srcPath ? path.dirname(srcPath) : '';
  } catch (e) {
    return '';
  }
}

// 将 HTML 写入临时文件并用 loadFile 加载。
// 图片内嵌 base64 后 HTML 很大，data:URL 会超长导致 ERR_INVALID_URL，故改用临时文件。
let __tmpCounter = 0;
async function loadHtmlIntoWindow(win, html) {
  const tmpDir = app.getPath('temp');
  __tmpCounter += 1;
  const tmpFile = path.join(tmpDir, `md2pdf-${process.pid}-${__tmpCounter}.html`);
  fs.writeFileSync(tmpFile, html, 'utf-8');
  try {
    await win.loadFile(tmpFile);
  } finally {
    try { fs.unlinkSync(tmpFile); } catch (e) {}
  }
}

// 读取本地 .md 文件并返回内容（文件关联双击打开时由渲染层调用）
ipcMain.handle('read-md-file', async (_, filePath) => {
  try {
    if (!filePath || typeof filePath !== 'string') return { error: 'no-path' };
    const content = fs.readFileSync(filePath, 'utf-8');
    return { content, filename: path.basename(filePath), srcPath: filePath };
  } catch (e) {
    return { error: e.message };
  }
});

// 预览内编辑保存：把 markdown 内容写回源文件
ipcMain.handle('write-md-file', async (_, { filePath, content }) => {
  try {
    if (!filePath || typeof filePath !== 'string') return { success: false, error: 'no-path' };
    if (typeof content !== 'string') return { success: false, error: 'no-content' };
    fs.writeFileSync(filePath, content, 'utf-8');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// ====== 文件实时同步：监听正在预览的 .md 文件，外部编辑保存后通知渲染层刷新 ======
const watchedFiles = new Map(); // filePath -> dirPath
const dirWatchers = new Map();  // dirPath -> { watcher, files:Set<basename> }
const watchTimers = new Map();  // filePath -> timeout（防抖）

function debouncedNotify(filePath) {
  if (watchTimers.has(filePath)) clearTimeout(watchTimers.get(filePath));
  watchTimers.set(filePath, setTimeout(() => {
    watchTimers.delete(filePath);
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('file-changed', filePath);
    }
  }, 300));
}

// 监听单个文件。实际监听其所在目录并按文件名过滤：
// 兼容「先写临时文件再 rename 覆盖」的原子保存（直接 fs.watch(file) 在文件被替换后会失效）。
function watchFile(filePath) {
  if (!filePath || watchedFiles.has(filePath)) return;
  const dir = path.dirname(filePath);
  const base = path.basename(filePath);

  let entry = dirWatchers.get(dir);
  if (!entry) {
    const files = new Set();
    const watcher = fs.watch(dir, (eventType, filename) => {
      if (!filename) {
        files.forEach(f => debouncedNotify(path.join(dir, f)));
      } else if (files.has(filename)) {
        debouncedNotify(path.join(dir, filename));
      }
    });
    watcher.on('error', () => {});
    entry = { watcher, files };
    dirWatchers.set(dir, entry);
  }
  entry.files.add(base);
  watchedFiles.set(filePath, dir);
}

function unwatchFile(filePath) {
  if (!filePath) return;
  const dir = watchedFiles.get(filePath);
  if (!dir) return;
  const base = path.basename(filePath);
  const entry = dirWatchers.get(dir);
  if (entry) {
    entry.files.delete(base);
    if (entry.files.size === 0) {
      entry.watcher.close();
      dirWatchers.delete(dir);
    }
  }
  watchedFiles.delete(filePath);
  if (watchTimers.has(filePath)) {
    clearTimeout(watchTimers.get(filePath));
    watchTimers.delete(filePath);
  }
}

ipcMain.handle('watch-file', async (_, filePath) => {
  try { watchFile(filePath); } catch (e) {}
});

ipcMain.handle('unwatch-file', async (_, filePath) => {
  try { unwatchFile(filePath); } catch (e) {}
});

// 预览：返回渲染好的 HTML 字符串
ipcMain.handle('preview-md', async (_, { mdContent, srcPath }) => {
  return buildHtml(mdContent, basedirOf(srcPath));
});

// 预览：多个 markdown 合并为一个 HTML（顶部标签页 + 汇总目录）
ipcMain.handle('preview-merge', async (_, { files }) => {
  return buildMergedHtml(files || [], undefined, true);
});

// 转换：渲染 HTML → printToPDF → 添加书签 → 保存文件
ipcMain.handle('convert-md', async (event, { mdContent, filename, srcPath }) => {
  const html = buildHtml(mdContent, basedirOf(srcPath));
  const headings = extractHeadings(mdContent);

  const hidden = new BrowserWindow({
    show: false,
    webPreferences: { offscreen: true }
  });

  try {
    await loadHtmlIntoWindow(hidden, html);

    // 等待 mermaid 渲染完成（最多 15 秒）
    await hidden.webContents.executeJavaScript(`
      new Promise((resolve) => {
        const check = () => {
          const divs = document.querySelectorAll('.mermaid');
          if (divs.length === 0 || Array.from(divs).every(d => d.querySelector('svg'))) {
            resolve();
          } else {
            setTimeout(check, 200);
          }
        };
        setTimeout(check, 500);
      })
    `);

    // 获取每个标题在页面中的 Y 坐标（像素）及文档总高度
    const headingData = await hidden.webContents.executeJavaScript(`
      (() => {
        const hs = document.querySelectorAll('#main h1, #main h2, #main h3, #main h4, #main h5, #main h6');
        return {
          totalHeight: document.documentElement.scrollHeight,
          headings: Array.from(hs).map(h => ({
            top: h.getBoundingClientRect().top + window.scrollY,
            level: parseInt(h.tagName.substring(1)),
            text: h.textContent
          }))
        };
      })()
    `);

    const pdfBuffer = await hidden.webContents.printToPDF({
      pageSize: hasWideTable(mdContent) ? 'A3' : 'A4',
      marginType: 'default',
      printBackground: true
    });

    // 使用 pdf-lib 添加书签
    const pdfWithBookmarks = await addBookmarks(pdfBuffer, headingData.headings, headingData.totalHeight);

    // 弹出保存对话框
    const baseName = path.basename(filename, path.extname(filename));
    const { filePath, canceled } = await dialog.showSaveDialog({
      defaultPath: baseName + '.pdf',
      filters: [{ name: 'PDF', extensions: ['pdf'] }]
    });

    if (canceled || !filePath) return { success: false };

    fs.writeFileSync(filePath, pdfWithBookmarks);
    return { success: true, filePath };
  } finally {
    hidden.destroy();
  }
});

// 下载 PDF 到指定目录（独立下载按钮）
ipcMain.handle('download-pdf', async (event, { mdContent, filename, srcPath }) => {
  const html = buildHtml(mdContent, basedirOf(srcPath));
  const headings = extractHeadings(mdContent);

  // 让用户选择保存目录
  const { filePaths, canceled } = await dialog.showOpenDialog({
    title: '选择 PDF 保存目录',
    properties: ['openDirectory']
  });

  if (canceled || !filePaths || filePaths.length === 0) return { success: false };

  const baseName = path.basename(filename, path.extname(filename));
  const savePath = path.join(filePaths[0], baseName + '.pdf');

  const hidden = new BrowserWindow({
    show: false,
    webPreferences: { offscreen: true }
  });

  try {
    await loadHtmlIntoWindow(hidden, html);

    await hidden.webContents.executeJavaScript(`
      new Promise((resolve) => {
        const check = () => {
          const divs = document.querySelectorAll('.mermaid');
          if (divs.length === 0 || Array.from(divs).every(d => d.querySelector('svg'))) {
            resolve();
          } else {
            setTimeout(check, 200);
          }
        };
        setTimeout(check, 500);
      })
    `);

    const headingData = await hidden.webContents.executeJavaScript(`
      (() => {
        const hs = document.querySelectorAll('#main h1, #main h2, #main h3, #main h4, #main h5, #main h6');
        return {
          totalHeight: document.documentElement.scrollHeight,
          headings: Array.from(hs).map(h => ({
            top: h.getBoundingClientRect().top + window.scrollY,
            level: parseInt(h.tagName.substring(1)),
            text: h.textContent
          }))
        };
      })()
    `);

    const pdfBuffer = await hidden.webContents.printToPDF({
      pageSize: hasWideTable(mdContent) ? 'A3' : 'A4',
      marginType: 'default',
      printBackground: true
    });

    const pdfWithBookmarks = await addBookmarks(pdfBuffer, headingData.headings, headingData.totalHeight);
    fs.writeFileSync(savePath, pdfWithBookmarks);
    shell.openPath(savePath);
    return { success: true, filePath: savePath };
  } finally {
    hidden.destroy();
  }
});

// 为 PDF 添加书签（大纲）
async function addBookmarks(pdfBuffer, headingPositions, totalHeight) {
  const pdfDoc = await PDFDocument.load(pdfBuffer);
  const pages = pdfDoc.getPages();

  if (pages.length === 0 || headingPositions.length === 0) {
    return pdfBuffer;
  }

  const pageCount = pages.length;
  const pageHeightPt = pages[0].getHeight();
  // 用文档实际总高度除以页数，得到每页对应的 CSS 像素高度
  const cssHeightPerPage = totalHeight / pageCount;

  const outlineItems = headingPositions.map(h => {
    const pageIndex = Math.max(0, Math.min(pageCount - 1, Math.floor(h.top / cssHeightPerPage)));
    const posInPage = h.top - pageIndex * cssHeightPerPage;
    const yPt = pageHeightPt - (posInPage / cssHeightPerPage) * pageHeightPt;
    return {
      title: h.text,
      level: h.level,
      pageIndex,
      yPt: Math.max(0, Math.min(pageHeightPt, yPt))
    };
  });

  const context = pdfDoc.context;
  const outlineRef = context.nextRef();
  const refs = outlineItems.map(() => context.nextRef());

  const tree = buildOutlineTree(outlineItems, refs, outlineRef);

  for (let i = 0; i < outlineItems.length; i++) {
    const item = outlineItems[i];
    const node = tree[i];
    const page = pages[item.pageIndex];

    const destArray = PDFArray.withContext(context);
    destArray.push(page.ref);
    destArray.push(PDFName.of('XYZ'));
    destArray.push(PDFNumber.of(0));
    destArray.push(PDFNumber.of(Math.round(item.yPt)));
    destArray.push(PDFNull);

    const dict = PDFDict.withContext(context);
    dict.set(PDFName.of('Title'), PDFHexString.fromText(item.title));
    dict.set(PDFName.of('Parent'), node.parent);
    dict.set(PDFName.of('Dest'), destArray);

    if (node.prev) dict.set(PDFName.of('Prev'), node.prev);
    if (node.next) dict.set(PDFName.of('Next'), node.next);
    if (node.first) dict.set(PDFName.of('First'), node.first);
    if (node.last) dict.set(PDFName.of('Last'), node.last);
    if (node.count !== undefined) dict.set(PDFName.of('Count'), PDFNumber.of(node.count));

    context.assign(refs[i], dict);
  }

  const topLevelIndices = tree.reduce((acc, n, i) => { if (n.parentIsRoot) acc.push(i); return acc; }, []);
  const outlineDict = PDFDict.withContext(context);
  outlineDict.set(PDFName.of('Type'), PDFName.of('Outlines'));
  outlineDict.set(PDFName.of('First'), refs[topLevelIndices[0]]);
  outlineDict.set(PDFName.of('Last'), refs[topLevelIndices[topLevelIndices.length - 1]]);
  outlineDict.set(PDFName.of('Count'), PDFNumber.of(topLevelIndices.length));
  context.assign(outlineRef, outlineDict);

  pdfDoc.catalog.set(PDFName.of('Outlines'), outlineRef);

  return Buffer.from(await pdfDoc.save());
}

function buildOutlineTree(items, refs, outlineRef) {
  const nodes = items.map(() => ({ parent: null, prev: null, next: null, first: null, last: null, parentIsRoot: false, count: undefined }));

  const stack = [];

  for (let i = 0; i < items.length; i++) {
    const level = items[i].level;

    while (stack.length > 0 && stack[stack.length - 1].level >= level) {
      stack.pop();
    }

    if (stack.length === 0) {
      nodes[i].parentIsRoot = true;
      nodes[i].parent = outlineRef;
    } else {
      const parentIdx = stack[stack.length - 1].index;
      nodes[i].parent = refs[parentIdx];

      if (!nodes[parentIdx].first) {
        nodes[parentIdx].first = refs[i];
      }
      nodes[parentIdx].last = refs[i];
      nodes[parentIdx].count = (nodes[parentIdx].count || 0) + 1;
    }

    for (let j = i - 1; j >= 0; j--) {
      if (items[j].level === level) {
        const sameParent = (nodes[i].parentIsRoot && nodes[j].parentIsRoot) ||
          (nodes[i].parent && nodes[j].parent && nodes[i].parent === nodes[j].parent);
        if (sameParent) {
          nodes[i].prev = refs[j];
          nodes[j].next = refs[i];
          break;
        }
      }
      if (items[j].level < level) break;
    }

    stack.push({ level, index: i });
  }

  return nodes;
}

// 转换：渲染 HTML → 内联 mermaid SVG → 保存为 .html 文件
ipcMain.handle('convert-html', async (event, { mdContent, filename, srcPath }) => {
  const html = buildHtml(mdContent, basedirOf(srcPath));

  const hidden = new BrowserWindow({
    show: false,
    webPreferences: { offscreen: true }
  });

  try {
    await loadHtmlIntoWindow(hidden, html);

    // 等待 mermaid 渲染完成
    await hidden.webContents.executeJavaScript(`
      new Promise((resolve) => {
        const check = () => {
          const divs = document.querySelectorAll('.mermaid');
          if (divs.length === 0 || Array.from(divs).every(d => d.querySelector('svg'))) {
            resolve();
          } else {
            setTimeout(check, 200);
          }
        };
        setTimeout(check, 500);
      })
    `);

    // 移除 mermaid CDN 脚本，仅保留已渲染的 SVG，使 HTML 离线可用
    const finalHtml = await hidden.webContents.executeJavaScript(`
      (() => {
        document.querySelectorAll('script[src]').forEach(s => s.remove());
        return '<!DOCTYPE html>\\n' + document.documentElement.outerHTML;
      })()
    `);

    const baseName = path.basename(filename, path.extname(filename));
    const { filePath, canceled } = await dialog.showSaveDialog({
      defaultPath: baseName + '.html',
      filters: [{ name: 'HTML', extensions: ['html'] }]
    });

    if (canceled || !filePath) return { success: false };

    fs.writeFileSync(filePath, finalHtml, 'utf-8');
    shell.openPath(filePath);
    return { success: true, filePath };
  } finally {
    hidden.destroy();
  }
});

// 导出：多个 markdown 合并后的 HTML（内联 mermaid SVG，离线可用）
ipcMain.handle('convert-merge', async (event, { files, filename }) => {
  const html = buildMergedHtml(files || [], undefined, false);

  const hidden = new BrowserWindow({
    show: false,
    webPreferences: { offscreen: true }
  });

  try {
    await loadHtmlIntoWindow(hidden, html);

    await hidden.webContents.executeJavaScript(`
      new Promise((resolve) => {
        const check = () => {
          const divs = document.querySelectorAll('.mermaid');
          if (divs.length === 0 || Array.from(divs).every(d => d.querySelector('svg'))) {
            resolve();
          } else {
            setTimeout(check, 200);
          }
        };
        setTimeout(check, 500);
      })
    `);

    const finalHtml = await hidden.webContents.executeJavaScript(`
      (() => {
        document.querySelectorAll('script[src]').forEach(s => s.remove());
        return '<!DOCTYPE html>\\n' + document.documentElement.outerHTML;
      })()
    `);

    const baseName = (filename || '合并文档').replace(/\.[^.]+$/, '');
    const { filePath, canceled } = await dialog.showSaveDialog({
      defaultPath: baseName + '.html',
      filters: [{ name: 'HTML', extensions: ['html'] }]
    });

    if (canceled || !filePath) return { success: false };

    fs.writeFileSync(filePath, finalHtml, 'utf-8');
    shell.openPath(filePath);
    return { success: true, filePath };
  } finally {
    hidden.destroy();
  }
});

// 多文件合并为单一 markdown：每个文件按自身目录解析相对图片，再按顺序自上而下拼接
function mergeMarkdown(files) {
  return (files || []).map(f => {
    const content = f.content || '';
    const basedir = f.srcPath ? path.dirname(f.srcPath) : '';
    // 相对图片路径 → 绝对 file:// URL，避免拼接后因所在目录不同而失效
    return content.replace(/!\[([^\]]*)\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g, (whole, alt, href) => {
      if (/^(https?:|data:|file:)/i.test(href) || path.isAbsolute(href)) return whole;
      return `![${alt}](${pathToFileURL(path.resolve(basedir, href)).href})`;
    });
  }).join('\n\n---\n\n');
}

// 导出：多个 markdown 自上而下合并为 PDF
ipcMain.handle('convert-merge-pdf', async (event, { files, filename }) => {
  const mergedMd = mergeMarkdown(files);
  const html = buildHtml(mergedMd, '');
  const headings = extractHeadings(mergedMd);

  const hidden = new BrowserWindow({
    show: false,
    webPreferences: { offscreen: true }
  });

  try {
    await loadHtmlIntoWindow(hidden, html);

    await hidden.webContents.executeJavaScript(`
      new Promise((resolve) => {
        const check = () => {
          const divs = document.querySelectorAll('.mermaid');
          if (divs.length === 0 || Array.from(divs).every(d => d.querySelector('svg'))) {
            resolve();
          } else {
            setTimeout(check, 200);
          }
        };
        setTimeout(check, 500);
      })
    `);

    const headingData = await hidden.webContents.executeJavaScript(`
      (() => {
        const hs = document.querySelectorAll('#main h1, #main h2, #main h3, #main h4, #main h5, #main h6');
        return {
          totalHeight: document.documentElement.scrollHeight,
          headings: Array.from(hs).map(h => ({
            top: h.getBoundingClientRect().top + window.scrollY,
            level: parseInt(h.tagName.substring(1)),
            text: h.textContent
          }))
        };
      })()
    `);

    const pdfBuffer = await hidden.webContents.printToPDF({
      pageSize: hasWideTable(mergedMd) ? 'A3' : 'A4',
      marginType: 'default',
      printBackground: true
    });

    const pdfWithBookmarks = await addBookmarks(pdfBuffer, headingData.headings, headingData.totalHeight);

    const baseName = (filename || '合并文档').replace(/\.[^.]+$/, '');
    const { filePath, canceled } = await dialog.showSaveDialog({
      defaultPath: baseName + '.pdf',
      filters: [{ name: 'PDF', extensions: ['pdf'] }]
    });

    if (canceled || !filePath) return { success: false };

    fs.writeFileSync(filePath, pdfWithBookmarks);
    shell.openPath(filePath);
    return { success: true, filePath };
  } finally {
    hidden.destroy();
  }
});

// 导出：多个 markdown 自上而下合并为 Word
ipcMain.handle('convert-merge-word', async (event, { files, filename }) => {
  try {
    const mergedMd = mergeMarkdown(files);
    const mermaidImages = await renderMermaidToPngs(mergedMd);
    const doc = buildDocx(mergedMd, mermaidImages, '');

    const baseName = (filename || '合并文档').replace(/\.[^.]+$/, '');
    const { filePath, canceled } = await dialog.showSaveDialog({
      defaultPath: baseName + '.docx',
      filters: [{ name: 'Word', extensions: ['docx'] }]
    });

    if (canceled || !filePath) return { success: false };

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);
    shell.openPath(filePath);
    return { success: true, filePath };
  } catch (error) {
    safeLog('convert-merge-word error:', error);
    throw error;
  }
});

// 安全的日志函数，避免 EPIPE 错误
function safeLog(...args) {
  try {
    console.log(...args);
  } catch (e) {
    // 忽略 EPIPE 错误
  }
}

// 渲染 Mermaid 代码块为 PNG 图片数组（独立轻量渲染，避免 GPU 崩溃）
async function renderMermaidToPngs(mdContent) {
  try {
    safeLog('renderMermaidToPngs: Extracting mermaid blocks...');
    // 1. 只提取 mermaid 代码块，构建独立的最小 HTML
    const mermaidBlocks = extractMermaidBlocks(mdContent);
    safeLog(`renderMermaidToPngs: Found ${mermaidBlocks.length} mermaid blocks`);

    if (mermaidBlocks.length === 0) return [];

    // 独立的轻量 HTML，只包含 mermaid，避免大页面导致 GPU 崩溃
    const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body { margin: 0; padding: 20px; background: #fff; font-family: "Microsoft YaHei", Arial, sans-serif; }
  .mermaid-wrap { padding: 20px; margin-bottom: 20px; }
</style>
</head>
<body>
${mermaidBlocks.map((code, i) => `<div class="mermaid-wrap"><div class="mermaid" id="m${i}">${code}</div></div>`).join('\n')}
<script src="${MERMAID_SRC}"></script>
<script>
  mermaid.initialize({
    startOnLoad: true,
    theme: 'default',
    flowchart: { useMaxWidth: false, htmlLabels: false },
    themeVariables: { fontFamily: 'Arial, sans-serif' }
  });
</script>
</body>
</html>`;

    safeLog('renderMermaidToPngs: HTML built, length:', html.length);

    const hidden = new BrowserWindow({
      show: false,
      width: 1600,
      height: 3000,
      webPreferences: {
        offscreen: false,
        nodeIntegration: false,
        contextIsolation: true
      }
    });

    safeLog('renderMermaidToPngs: Hidden window created');

    try {
      await loadHtmlIntoWindow(hidden, html);
      safeLog('renderMermaidToPngs: HTML loaded');

      // 2. 等待 mermaid 渲染完成
      await hidden.webContents.executeJavaScript(`
        new Promise((resolve) => {
          const check = () => {
            const divs = document.querySelectorAll('.mermaid');
            if (divs.length === 0 || Array.from(divs).every(d => d.querySelector('svg'))) {
              setTimeout(resolve, 300);
            } else {
              setTimeout(check, 200);
            }
          };
          setTimeout(check, 500);
        })
      `);
      safeLog('renderMermaidToPngs: Mermaid rendered');

      // 3. 提取每个 mermaid 的位置和尺寸
      const pageInfo = await hidden.webContents.executeJavaScript(`
        (() => {
          const rects = Array.from(document.querySelectorAll('.mermaid')).map(div => {
            const rect = div.getBoundingClientRect();
            return {
              x: Math.round(rect.left),
              y: Math.round(rect.top),
              width: Math.round(rect.width),
              height: Math.round(rect.height)
            };
          });
          return {
            rects,
            totalHeight: document.documentElement.scrollHeight,
            totalWidth: document.documentElement.scrollWidth
          };
        })()
      `);

      const rects = pageInfo.rects;
      safeLog(`Found ${rects.length} mermaid blocks, page: ${pageInfo.totalWidth}x${pageInfo.totalHeight}, rects:`, rects);

      // 4. 提取每个 SVG 的完整 XML，用 sharp 直接渲染（避免截图局限）
      const svgs = await hidden.webContents.executeJavaScript(`
        Array.from(document.querySelectorAll('.mermaid svg')).map(svg => {
          // 确保有明确的宽高
          const bbox = svg.getBBox();
          const width = svg.width.baseVal.value || bbox.width;
          const height = svg.height.baseVal.value || bbox.height;

          // 强制设置尺寸和 viewBox
          svg.setAttribute('width', width);
          svg.setAttribute('height', height);
          if (!svg.getAttribute('viewBox')) {
            svg.setAttribute('viewBox', '0 0 ' + width + ' ' + height);
          }

          // 关键：将所有 foreignObject 中的 HTML 文本转换成原生 SVG text（sharp 无法渲染 foreignObject）
          svg.querySelectorAll('foreignObject').forEach(fo => {
            const foWidth = parseFloat(fo.getAttribute('width')) || 100;
            const foHeight = parseFloat(fo.getAttribute('height')) || 20;
            const foX = parseFloat(fo.getAttribute('x')) || 0;
            const foY = parseFloat(fo.getAttribute('y')) || 0;

            // 递归提取文本内容
            const text = fo.textContent.trim();
            if (!text) { fo.remove(); return; }

            // 按换行拆分成多行
            const lines = text.split(/\\n|<br\\s*\\/?>/i).filter(l => l.trim());
            const displayLines = lines.length > 0 ? lines : [text];

            const svgNS = 'http://www.w3.org/2000/svg';
            const textEl = document.createElementNS(svgNS, 'text');
            textEl.setAttribute('x', foX + foWidth / 2);
            textEl.setAttribute('y', foY + foHeight / 2);
            textEl.setAttribute('text-anchor', 'middle');
            textEl.setAttribute('dominant-baseline', 'middle');
            textEl.setAttribute('font-family', 'Arial, "Microsoft YaHei", sans-serif');
            textEl.setAttribute('font-size', '14');
            textEl.setAttribute('fill', '#333');

            const lineHeight = 16;
            const startY = (foY + foHeight / 2) - ((displayLines.length - 1) * lineHeight) / 2;
            displayLines.forEach((line, i) => {
              const tspan = document.createElementNS(svgNS, 'tspan');
              tspan.setAttribute('x', foX + foWidth / 2);
              tspan.setAttribute('y', startY + i * lineHeight);
              tspan.textContent = line.trim();
              textEl.appendChild(tspan);
            });

            fo.parentNode.replaceChild(textEl, fo);
          });

          // 内联所有 style（避免依赖外部样式）
          const styleElements = document.querySelectorAll('style');
          let styleContent = '';
          styleElements.forEach(s => { styleContent += s.textContent; });

          // 添加 xmlns 命名空间
          if (!svg.getAttribute('xmlns')) {
            svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
          }

          // 序列化并包含 style
          const serializer = new XMLSerializer();
          let svgXml = serializer.serializeToString(svg);

          // 插入 style 到 SVG 头部（保留字体家族，去掉 foreignObject 相关样式即可）
          if (styleContent) {
            const safeStyle = styleContent.replace(/</g, '&lt;').replace(/>/g, '&gt;');
            svgXml = svgXml.replace(/<svg([^>]*)>/, '<svg$1><style>' + safeStyle + '</style>');
          }

          return { xml: svgXml, width: Math.ceil(width), height: Math.ceil(height) };
        })
      `);

      const images = [];
      for (let i = 0; i < svgs.length; i++) {
        const svg = svgs[i];
        try {
          // 用 sharp 将 SVG 渲染为高分辨率 PNG（2x 密度）
          const density = 2;
          const pngBuffer = await sharp(Buffer.from(svg.xml, 'utf-8'), {
            density: 72 * density
          })
            .resize(svg.width * density, svg.height * density, {
              fit: 'contain',
              background: { r: 255, g: 255, b: 255, alpha: 1 }
            })
            .png()
            .toBuffer();

          safeLog(`Mermaid ${i} rendered via sharp: ${svg.width}x${svg.height}, buffer size: ${pngBuffer.length}`);

          images.push({
            buffer: pngBuffer,
            width: svg.width,
            height: svg.height
          });
        } catch (e) {
          safeLog(`Failed to render mermaid ${i} via sharp:`, e.message);
          // 回退：尝试截图
          try {
            const fullImage = await hidden.webContents.capturePage();
            const rect = rects[i];
            if (rect && fullImage) {
              const pngBuffer = await sharp(fullImage.toPNG())
                .extract({
                  left: Math.max(0, rect.x),
                  top: Math.max(0, rect.y),
                  width: rect.width,
                  height: rect.height
                })
                .png()
                .toBuffer();
              images.push({
                buffer: pngBuffer,
                width: rect.width,
                height: rect.height
              });
              safeLog(`Mermaid ${i} fallback captured: ${rect.width}x${rect.height}`);
            } else {
              images.push(null);
            }
          } catch (e2) {
            safeLog(`Mermaid ${i} fallback also failed:`, e2.message);
            images.push(null);
          }
        }
      }
      return images;
    } catch (err) {
      safeLog('Mermaid render error in hidden window:', err);
      return [];
    } finally {
      hidden.destroy();
      safeLog('renderMermaidToPngs: Hidden window destroyed');
    }
  } catch (err) {
    safeLog('renderMermaidToPngs: Fatal error:', err);
    return [];
  }
}

// 转换：MD → Word
ipcMain.handle('convert-word', async (event, { mdContent, filename, srcPath }) => {
  try {
    safeLog('convert-word handler called, mdContent length:', mdContent.length);

    // 1. 渲染完整 HTML 并截取 mermaid 图片
    safeLog('Starting renderMermaidToPngs...');
    const mermaidImages = await renderMermaidToPngs(mdContent);
    safeLog(`Rendered ${mermaidImages.length} mermaid images`);

    // 2. 构建 Word 文档
    safeLog('Starting buildDocx...');
    const doc = buildDocx(mdContent, mermaidImages, basedirOf(srcPath));
    const baseName = path.basename(filename, path.extname(filename));

    const { filePath, canceled } = await dialog.showSaveDialog({
      defaultPath: baseName + '.docx',
      filters: [{ name: 'Word', extensions: ['docx'] }]
    });

    if (canceled || !filePath) return { success: false };

    safeLog('Packing document...');
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);
    shell.openPath(filePath);
    safeLog('Word conversion completed');
    return { success: true, filePath };
  } catch (error) {
    safeLog('convert-word error:', error);
    throw error;
  }
});
