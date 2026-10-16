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
  convertWord: (mdContent, filename, srcPath) => ipcRenderer.invoke('convert-word', { mdContent, filename, srcPath }),
  downloadPdf: (mdContent, filename, srcPath) => ipcRenderer.invoke('download-pdf', { mdContent, filename, srcPath }),
  resizeLayout: (mode) => ipcRenderer.invoke('resize-layout', mode),
  selectBackground: () => ipcRenderer.invoke('select-background'),
  getBackground: () => ipcRenderer.invoke('get-background'),
  getAppInfo: () => ipcRenderer.invoke('get-app-info'),
  resetBackground: () => ipcRenderer.invoke('reset-background'),
  readMdFile: (filePath) => ipcRenderer.invoke('read-md-file', filePath),
  onOpenFile: (cb) => ipcRenderer.on('open-file', (_event, filePath) => cb(filePath))
});
