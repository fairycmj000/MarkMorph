const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('api', {
  // 获取拖入/选择文件的本地绝对路径（用于解析 markdown 中的相对图片路径）
  getFilePath: (file) => {
    try {
      if (webUtils && webUtils.getPathForFile) return webUtils.getPathForFile(file);
    } catch (e) {}
    return (file && file.path) || '';
  },
  previewMd: (mdContent, srcPath) => ipcRenderer.invoke('preview-md', { mdContent, srcPath }),
  convertMd: (mdContent, filename, srcPath) => ipcRenderer.invoke('convert-md', { mdContent, filename, srcPath }),
  convertHtml: (mdContent, filename, srcPath) => ipcRenderer.invoke('convert-html', { mdContent, filename, srcPath }),
  previewMerge: (files) => ipcRenderer.invoke('preview-merge', { files }),
  convertMerge: (files, filename) => ipcRenderer.invoke('convert-merge', { files, filename }),
  convertMergePdf: (files, filename) => ipcRenderer.invoke('convert-merge-pdf', { files, filename }),
  convertMergeWord: (files, filename) => ipcRenderer.invoke('convert-merge-word', { files, filename }),
  convertWord: (mdContent, filename, srcPath) => ipcRenderer.invoke('convert-word', { mdContent, filename, srcPath }),
  convertDocxToMd: (srcPath) => ipcRenderer.invoke('convert-docx-to-md', srcPath),
  saveDocxMd: (srcPath) => ipcRenderer.invoke('save-docx-md', srcPath),
  // Word → HTML：预览走专用 HTML 转换（表格/嵌入 Excel 保留颜色与合并），下载时另存为自包含 .html
  convertDocxToHtml: (srcPath) => ipcRenderer.invoke('convert-docx-to-html', srcPath),
  saveDocxHtml: (srcPath) => ipcRenderer.invoke('save-docx-html', srcPath),
  downloadPdf: (mdContent, filename, srcPath) => ipcRenderer.invoke('download-pdf', { mdContent, filename, srcPath }),
  resizeLayout: (mode) => ipcRenderer.invoke('resize-layout', mode),
  selectBackground: () => ipcRenderer.invoke('select-background'),
  getBackground: () => ipcRenderer.invoke('get-background'),
  getAppInfo: () => ipcRenderer.invoke('get-app-info'),
  openExternal: (url) => ipcRenderer.invoke('open-external', url),
  resetBackground: () => ipcRenderer.invoke('reset-background'),
  readMdFile: (filePath) => ipcRenderer.invoke('read-md-file', filePath),
  writeMdFile: (filePath, content) => ipcRenderer.invoke('write-md-file', { filePath, content }),
  onOpenFile: (cb) => ipcRenderer.on('open-file', (_event, filePath) => cb(filePath)),
  watchFile: (filePath) => ipcRenderer.invoke('watch-file', filePath),
  unwatchFile: (filePath) => ipcRenderer.invoke('unwatch-file', filePath),
  onFileChanged: (cb) => ipcRenderer.on('file-changed', (_event, filePath) => cb(filePath))
});
