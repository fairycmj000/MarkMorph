// docx2md.js —— 把 Word (.docx) 无损转换为 Markdown (.md)
//
// 能力：
//   - 标题（pStyle 1–6 → H1–H6）、段落、编号/项目列表、加粗/斜体
//   - Word 原生表格 → Markdown 表格（处理 gridSpan 横向合并 / vMerge 纵向合并）
//   - 图片（word/media/*）→ 提取为独立文件，md 里用 ./assets/ 相对路径引用
//   - 嵌入的 Excel OLE（ProgID 含 Excel、target 为 .xlsx）→ 渲染成 Markdown 表格（不保留 .xlsx 原文件）
//   - 外部超链接 → [文字](链接)
//
// 返回：{ markdown, images:[{filename,buffer,mime}], stats:{headings,tables,embeddedSheets,images,hyperlinks} }
//
// 依赖：jszip（读 zip 容器）+ fast-xml-parser（解析 OOXML / SpreadsheetML）。

'use strict';
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { XMLParser } = require('fast-xml-parser');

// preserveOrder：按文档顺序保留 <w:p> 与 <w:tbl> 的交错顺序（正文与表格混排时必须）。
// removeNSPrefix：去掉 w:/a:/r:/v:/o: 前缀，元素名与属性名都变简洁（如 w:p → p、r:id → @_id）。
// trimValues:false / parseTagValue:false：保留文本原始空白、数值按字符串处理（避免 "0"/"12.50" 被改写）。
const XML_OPTS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true,
  preserveOrder: true,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false
};

const MIME = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.bmp': 'image/bmp',
  '.webp': 'image/webp',
  '.svg': 'image/svg+xml'
};

function mimeFor(name) {
  return MIME[path.extname(name).toLowerCase()] || 'application/octet-stream';
}

// 按文件魔数识别真实图片格式（飞书导出常把 JPEG 存成 .png，扩展名不可靠）
function detectImageMime(buf) {
  if (buf.length >= 8 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) return { ext: '.png', mime: 'image/png' };
  if (buf.length >= 3 && buf[0] === 0xff && buf[1] === 0xd8) return { ext: '.jpg', mime: 'image/jpeg' };
  if (buf.length >= 6 && buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46) return { ext: '.gif', mime: 'image/gif' };
  if (buf.length >= 2 && buf[0] === 0x42 && buf[1] === 0x4d) return { ext: '.bmp', mime: 'image/bmp' };
  if (buf.length >= 12 && buf.slice(0, 4).toString('latin1') === 'RIFF' && buf.slice(8, 12).toString('latin1') === 'WEBP') return { ext: '.webp', mime: 'image/webp' };
  return null;
}

// ===================== preserveOrder 树的读取辅助 =====================
// 每个元素节点形如 { tagName: [子节点...], (有属性时) ":@": { @_属性: 值 } }；
// 文本节点形如 { "#text": "字符串" }。
function nodeTag(node) {
  const keys = Object.keys(node || {});
  for (const k of keys) if (k !== ':@') return k;
  return null;
}
function nodeVal(node) { return node[nodeTag(node)]; }
function nodeAttrs(node) { return (node && node[':@']) || {}; }
function attr(node, name) { return node ? nodeAttrs(node)[name] : undefined; }
function childrenOf(node) {
  const v = nodeVal(node);
  return Array.isArray(v) ? v : [];
}
function textOf(node) {
  let s = '';
  for (const c of childrenOf(node)) {
    if (nodeTag(c) === '#text') s += nodeVal(c);
  }
  return decodeNumericEntities(s);
}

// fast-xml-parser 只解码命名实体（&amp;/&lt;/&gt;/&quot;/&apos;），不解码数值字符引用
// （&#xA;/&#10;/&#20013; 等），导致 Excel/Word 文本里的换行等以字面量 &#xA; 泄漏到输出。
// 这里手动解码数值字符引用，还原为真实字符（换行 → \n，再由各渲染器转成 <br>）。
function decodeNumericEntities(s) {
  if (s == null || s.indexOf('&#') === -1) return s;
  return String(s).replace(/&#(x[0-9A-Fa-f]+|\d+);/g, (m, n) => {
    const code = (n[0] === 'x' || n[0] === 'X') ? parseInt(n.slice(1), 16) : parseInt(n, 10);
    if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return m;
    try { return String.fromCodePoint(code); } catch (e) { return m; }
  });
}
function findChild(node, tag) {
  if (!node) return null;
  for (const c of childrenOf(node)) if (nodeTag(c) === tag) return c;
  return null;
}
function findDescendant(node, tag) {
  if (!node) return null;
  if (nodeTag(node) === tag) return node;
  for (const c of childrenOf(node)) {
    const r = findDescendant(c, tag);
    if (r) return r;
  }
  return null;
}
function findAllDescendants(node, tag, out) {
  out = out || [];
  if (nodeTag(node) === tag) out.push(node);
  for (const c of childrenOf(node)) findAllDescendants(c, tag, out);
  return out;
}

// 从 zip 里读字符串（缺失返回 ''）
async function zipString(zip, p) {
  const f = zip.file(p);
  return f ? await f.async('string') : '';
}

// ===================== 关系映射（rId → 目标） =====================
function parseRels(relsXml) {
  const map = {};
  if (!relsXml) return map;
  const tree = new XMLParser(XML_OPTS).parse(relsXml);
  const root = tree.find((n) => nodeTag(n) === 'Relationships');
  if (!root) return map;
  for (const c of childrenOf(root)) {
    if (nodeTag(c) !== 'Relationship') continue;
    const id = attr(c, '@_Id');
    if (id) {
      map[id] = {
        target: attr(c, '@_Target') || '',
        external: attr(c, '@_TargetMode') === 'External',
        type: attr(c, '@_Type') || ''
      };
    }
  }
  return map;
}

// ===================== 编号列表映射（numId → 各级 numFmt） =====================
function parseNumbering(numberingXml) {
  const map = { numToAbstract: {}, abstract: {} };
  if (!numberingXml) return { format: () => 'decimal' };
  const tree = new XMLParser(XML_OPTS).parse(numberingXml);
  const root = tree.find((n) => nodeTag(n) === 'numbering');
  if (!root) return { format: () => 'decimal' };
  for (const c of childrenOf(root)) {
    const t = nodeTag(c);
    if (t === 'abstractNum') {
      const aid = attr(c, '@_abstractNumId');
      const levels = {};
      for (const lvl of childrenOf(c).filter((x) => nodeTag(x) === 'lvl')) {
        const ilvl = attr(lvl, '@_ilvl');
        const numFmt = attr(findChild(lvl, 'numFmt'), '@_val');
        levels[ilvl] = numFmt;
      }
      map.abstract[aid] = levels;
    } else if (t === 'num') {
      const nid = attr(c, '@_numId');
      const ref = findChild(c, 'abstractNumId');
      map.numToAbstract[nid] = attr(ref, '@_val');
    }
  }
  return {
    format(numId, ilvl) {
      const aid = map.numToAbstract[numId];
      const lv = aid != null ? map.abstract[aid] : null;
      return (lv && lv[ilvl]) || 'decimal';
    }
  };
}

// ===================== 内联格式渲染 =====================
function wrap(s, bold, italic) {
  if (!bold && !italic) return s;
  // 用内联 HTML（<strong>/<em>）而非 Markdown 的 **/*：中文文本里加粗若以标点结尾、
  // 后接汉字（如 `**标签：**正文`），CommonMark 左右侧翼规则会判定结束标记无法闭合，
  // 把 ** 泄漏成字面字符；HTML 标签不受此规则影响，且 marked 会原样透传。
  // 内容先做 HTML 转义，避免 Word 文本里的 < > & 被误当成标签或实体。
  const lead = (s.match(/^\s*/) || [''])[0];
  const trail = (s.match(/\s*$/) || [''])[0];
  const core = s.slice(lead.length, s.length - trail.length);
  if (!core) return s;
  const esc = escapeHtml(core);
  let out;
  if (italic && bold) out = '<strong><em>' + esc + '</em></strong>';
  else if (bold) out = '<strong>' + esc + '</strong>';
  else out = '<em>' + esc + '</em>';
  return lead + out + trail;
}
function hasOn(rPr, tag) {
  if (!rPr) return false;
  const el = findChild(rPr, tag);
  if (!el) return false;
  const v = attr(el, '@_val');
  return !(v === 'false' || v === '0' || v === 'off');
}

// rel target（相对 word/）→ zip 内真实路径（posix 分隔）
function resolveZipPath(target) {
  const t = String(target || '').replace(/^\/+/, '');
  return path.posix.normalize(path.posix.join('word', t));
}

async function ensureImage(target, filename, ctx) {
  if (ctx.images.has(filename)) return filename;
  const zipPath = resolveZipPath(target);
  const f = ctx.zip.file(zipPath);
  if (!f) return filename;
  const buf = await f.async('nodebuffer');
  // 飞书导出常把 JPEG 存成 .png：按魔数修正扩展名与 MIME，避免内联/落盘格式错误
  const det = detectImageMime(buf);
  const origMime = mimeFor(filename);
  let outName = filename;
  let mime = origMime;
  if (det && det.mime !== origMime) {
    outName = path.posix.basename(filename, path.posix.extname(filename)) + det.ext;
    mime = det.mime;
  }
  if (ctx.images.has(outName)) return outName;
  ctx.images.set(outName, { buffer: buf, mime });
  ctx.imageOrder.push(outName);
  ctx.stats.images++;
  return outName;
}

// ===================== HTML 输出辅助（Word→HTML 无损表格用） =====================
function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// 提取已渲染 HTML 片段的纯文本（用于表头识别与字段筛选值）
function htmlToText(html) {
  return String(html == null ? '' : html)
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

// ARGB（如 FFFFF5EB）→ CSS 颜色（#FFF5EB）；6 位 RGB 原样；无效返回 ''
function argbToCss(rgb) {
  if (!rgb) return '';
  let s = String(rgb).replace(/^#/, '');
  if (s.length === 8) s = s.slice(2); // 去掉 ARGB 的 alpha（FF=不透明）
  if (!/^[0-9A-Fa-f]{6}$/.test(s)) return '';
  return '#' + s.toUpperCase();
}

// Word 单元格底纹（tcPr/shd）→ 背景色；无底纹或 clear 返回 ''
function cellFillColor(tcPr) {
  if (!tcPr) return '';
  const shd = findChild(tcPr, 'shd');
  if (!shd) return '';
  const val = attr(shd, '@_val') || '';
  if (val === 'clear') return '';
  const fill = attr(shd, '@_fill');
  if (!fill || fill === 'auto') return '';
  return argbToCss(fill);
}

// 生成 <td> 的 style：始终给背景色（未着色用白色），以盖掉预览 CSS 的斑马纹，忠实还原原表
function tdStyle(fill, font) {
  const parts = [];
  parts.push(`background-color:${fill || '#ffffff'}`);
  if (font) parts.push(`color:${font}`);
  return parts.join(';');
}

async function renderImageByRid(rId, ctx) {
  const rel = ctx.rels[rId];
  if (!rel || !rel.target) return '';
  const filename = path.posix.basename(rel.target);
  const outName = await ensureImage(rel.target, filename, ctx);
  return `![](./assets/${outName})`;
}

async function renderDrawing(dNode, ctx) {
  const blip = findDescendant(dNode, 'blip');
  if (!blip) return '';
  const rId = attr(blip, '@_embed');
  if (!rId) return '';
  return await renderImageByRid(rId, ctx);
}

async function renderPict(pNode, ctx) {
  const img = findDescendant(pNode, 'imagedata');
  if (img) {
    const rId = attr(img, '@_id');
    if (rId) return await renderImageByRid(rId, ctx);
  }
  return '';
}

async function renderObject(oNode, ctx) {
  const ole = findDescendant(oNode, 'OLEObject');
  if (!ole) return '';
  const rId = attr(ole, '@_id');
  const progId = attr(ole, '@_ProgID') || '';
  const rel = ctx.rels[rId];
  if (!rel || !rel.target) return '';
  // 嵌入 Excel：渲染成 Markdown 表格（不保留 .xlsx）
  if (/excel/i.test(progId) && /\.xlsx$/i.test(rel.target)) {
    const label = path.posix.basename(rel.target, '.xlsx').replace(/^Microsoft_Excel_/, '');
    const zipPath = resolveZipPath(rel.target);
    const f = ctx.zip.file(zipPath);
    if (f) {
      const buf = await f.async('nodebuffer');
      const table = await renderXlsx(buf, label);
      if (table) {
        ctx.stats.embeddedSheets++;
        return table;
      }
    }
  }
  // 非 Excel 的 OLE：回退为预览图
  const img = findDescendant(oNode, 'imagedata');
  if (img) {
    const irId = attr(img, '@_id');
    if (irId) return await renderImageByRid(irId, ctx);
  }
  return '';
}

// 提取一个 run 的纯文本、内嵌对象与加粗/斜体格式（供合并相邻同格式 run 使用）
async function runParts(rNode, ctx) {
  const rPr = findChild(rNode, 'rPr');
  const bold = hasOn(rPr, 'b');
  const italic = hasOn(rPr, 'i');
  const kids = childrenOf(rNode).filter((c) => nodeTag(c) !== 'rPr');
  let text = '';
  let extra = '';
  for (const k of kids) {
    const t = nodeTag(k);
    if (t === 't') text += textOf(k);
    else if (t === 'br' || t === 'cr') text += '\n';
    else if (t === 'tab') text += ' ';
    else if (t === 'noBreakHyphen') text += '-';
    else if (t === 'drawing') extra += await renderDrawing(k, ctx);
    else if (t === 'object') extra += await renderObject(k, ctx);
    else if (t === 'pict') extra += await renderPict(k, ctx);
    // instrText / fldChar / delText / sym 等忽略
  }
  return { text, extra, bold, italic };
}

async function renderInlines(children, ctx) {
  let out = '';
  // 相邻、加粗/斜体格式完全相同的纯文本 run 先合并再包裹，避免出现
  // `**a****b**` 这类让 marked 误解析、把 `**` 泄漏成字面字符的序列
  let buf = '';
  let bufBold = false;
  let bufItalic = false;
  const flush = () => {
    if (buf) {
      out += wrap(buf, bufBold, bufItalic);
      buf = '';
      bufBold = false;
      bufItalic = false;
    }
  };
  for (const c of children) {
    const t = nodeTag(c);
    if (t === '#text') {
      flush();
      out += nodeVal(c);
    } else if (t === 'r') {
      const { text, extra, bold, italic } = await runParts(c, ctx);
      if (extra) {
        // 含图片/对象等非文本内容：不参与合并，先 flush 再单独输出
        flush();
        if (text) out += wrap(text, bold, italic);
        if (text && extra) out += ' ';
        out += extra;
      } else if (text === '') {
        // 空 run：忽略
      } else if (buf !== '' && (bold !== bufBold || italic !== bufItalic)) {
        flush();
        buf = text;
        bufBold = bold;
        bufItalic = italic;
      } else {
        buf += text;
        bufBold = bold;
        bufItalic = italic;
      }
    } else if (t === 'hyperlink') {
      flush();
      out += await renderHyperlink(c, ctx);
    } else if (t === 'sdt') {
      const sc = findChild(c, 'sdtContent');
      if (sc) { flush(); out += await renderInlines(childrenOf(sc), ctx); }
    } else if (t === 'smartTag') {
      flush();
      out += await renderInlines(childrenOf(c), ctx);
    }
    // bookmarkStart/End、proofErr、commentRange* 等忽略
  }
  flush();
  return out;
}

async function renderHyperlink(hNode, ctx) {
  const rId = attr(hNode, '@_id');
  const rel = rId ? ctx.rels[rId] : null;
  const url = rel ? rel.target : '';
  const text = await renderInlines(childrenOf(hNode), ctx);
  if (!text) return '';
  ctx.stats.hyperlinks++;
  if (url) return `[${text}](${url})`;
  return text;
}

// ===================== HTML 内联渲染（表格单元格 / 嵌入对象用） =====================
async function renderHtmlImageByRid(rId, ctx) {
  const rel = ctx.rels[rId];
  if (!rel || !rel.target) return '';
  const filename = path.posix.basename(rel.target);
  const outName = await ensureImage(rel.target, filename, ctx);
  return `<img src="./assets/${outName}" alt="">`;
}

async function renderHtmlDrawing(dNode, ctx) {
  const blip = findDescendant(dNode, 'blip');
  if (!blip) return '';
  const rId = attr(blip, '@_embed');
  if (!rId) return '';
  return await renderHtmlImageByRid(rId, ctx);
}

async function renderHtmlPict(pNode, ctx) {
  const img = findDescendant(pNode, 'imagedata');
  if (img) {
    const rId = attr(img, '@_id');
    if (rId) return await renderHtmlImageByRid(rId, ctx);
  }
  return '';
}

async function renderHtmlObject(oNode, ctx) {
  const ole = findDescendant(oNode, 'OLEObject');
  if (!ole) return '';
  const rId = attr(ole, '@_id');
  const progId = attr(ole, '@_ProgID') || '';
  const rel = ctx.rels[rId];
  if (!rel || !rel.target) return '';
  // 嵌入 Excel：渲染成带颜色/合并的 HTML 表格
  if (/excel/i.test(progId) && /\.xlsx$/i.test(rel.target)) {
    const label = path.posix.basename(rel.target, '.xlsx').replace(/^Microsoft_Excel_/, '');
    const zipPath = resolveZipPath(rel.target);
    const f = ctx.zip.file(zipPath);
    if (f) {
      const buf = await f.async('nodebuffer');
      const table = await renderXlsxHtml(buf, label);
      if (table) {
        ctx.stats.embeddedSheets++;
        return table;
      }
    }
  }
  // 非 Excel 的 OLE：回退为预览图
  const img = findDescendant(oNode, 'imagedata');
  if (img) {
    const irId = attr(img, '@_id');
    if (irId) return await renderHtmlImageByRid(irId, ctx);
  }
  return '';
}

async function renderHtmlRun(rNode, ctx) {
  const rPr = findChild(rNode, 'rPr');
  const bold = hasOn(rPr, 'b');
  const italic = hasOn(rPr, 'i');
  const underline = hasOn(rPr, 'u');
  const strike = hasOn(rPr, 'strike');
  const colorEl = rPr ? findChild(rPr, 'color') : null;
  const colorVal = colorEl ? (attr(colorEl, '@_val') || '') : '';
  const kids = childrenOf(rNode).filter((c) => nodeTag(c) !== 'rPr');
  let text = '';
  let extra = '';
  for (const k of kids) {
    const t = nodeTag(k);
    if (t === 't') text += textOf(k);
    else if (t === 'br' || t === 'cr') text += '\n';
    else if (t === 'tab') text += ' ';
    else if (t === 'noBreakHyphen') text += '-';
    else if (t === 'drawing') extra += await renderHtmlDrawing(k, ctx);
    else if (t === 'object') extra += await renderHtmlObject(k, ctx);
    else if (t === 'pict') extra += await renderHtmlPict(k, ctx);
  }
  let s = escapeHtml(text).replace(/\n/g, '<br>');
  if (underline) s = `<u>${s}</u>`;
  if (strike) s = `<s>${s}</s>`;
  if (italic) s = `<em>${s}</em>`;
  if (bold) s = `<strong>${s}</strong>`;
  if (/^[0-9A-Fa-f]{6}$/.test(colorVal)) s = `<span style="color:#${colorVal.toUpperCase()}">${s}</span>`;
  if (extra) s += (s ? ' ' : '') + extra;
  return s;
}

async function renderHtmlHyperlink(hNode, ctx) {
  const rId = attr(hNode, '@_id');
  const rel = rId ? ctx.rels[rId] : null;
  const url = rel ? rel.target : '';
  const text = await renderHtmlInlines(childrenOf(hNode), ctx);
  if (!text) return '';
  ctx.stats.hyperlinks++;
  if (url) return `<a href="${escapeHtml(url)}">${text}</a>`;
  return text;
}

async function renderHtmlInlines(children, ctx) {
  let out = '';
  for (const c of children) {
    const t = nodeTag(c);
    if (t === '#text') out += escapeHtml(nodeVal(c));
    else if (t === 'r') out += await renderHtmlRun(c, ctx);
    else if (t === 'hyperlink') out += await renderHtmlHyperlink(c, ctx);
    else if (t === 'sdt') {
      const sc = findChild(c, 'sdtContent');
      if (sc) out += await renderHtmlInlines(childrenOf(sc), ctx);
    } else if (t === 'smartTag') out += await renderHtmlInlines(childrenOf(c), ctx);
  }
  return out;
}

// ===================== 段落 / 标题 / 列表 =====================
// 判断段落是否「仅含嵌入 Excel OLE 对象、无任何文本/图片/其他对象」
function isPureExcelParagraph(kids) {
  let excelCount = 0;
  let ok = true;
  const scan = (nodes) => {
    for (const n of nodes) {
      const t = nodeTag(n);
      if (t === '#text') { if (String(nodeVal(n)).trim()) ok = false; }
      else if (t === 'r') {
        for (const k of childrenOf(n).filter((x) => nodeTag(x) !== 'rPr')) {
          const kt = nodeTag(k);
          if (kt === 't') { if (textOf(k).trim()) ok = false; }
          else if (kt === 'object') {
            const ole = findDescendant(k, 'OLEObject');
            const progId = ole ? (attr(ole, '@_ProgID') || '') : '';
            if (/excel/i.test(progId)) excelCount++;
            else ok = false;
          } else if (kt === 'drawing' || kt === 'pict') ok = false;
        }
      } else if (t === 'hyperlink' || t === 'smartTag') scan(childrenOf(n));
      else if (t === 'sdt') {
        const sc = findChild(n, 'sdtContent');
        if (sc) scan(childrenOf(sc));
      }
    }
  };
  scan(kids);
  return ok && excelCount > 0;
}

async function renderParagraph(pNode, ctx) {
  const pPr = findChild(pNode, 'pPr');
  const kids = childrenOf(pNode).filter((c) => nodeTag(c) !== 'pPr');

  // （Word→HTML）段落仅含嵌入 Excel 对象时，直接输出带颜色/合并的 HTML 表格
  if (ctx.htmlTables && isPureExcelParagraph(kids)) {
    const parts = [];
    for (const c of kids) {
      if (nodeTag(c) !== 'r') continue;
      for (const k of childrenOf(c).filter((x) => nodeTag(x) !== 'rPr')) {
        if (nodeTag(k) === 'object') {
          const html = await renderHtmlObject(k, ctx);
          if (html) parts.push(html);
        }
      }
    }
    if (parts.length) return parts.join('\n');
  }

  // 标题：飞书导出用自定义段落样式 1–6 表示 H1–H6
  if (pPr) {
    const pStyle = findChild(pPr, 'pStyle');
    if (pStyle) {
      const v = attr(pStyle, '@_val');
      const lvl = /^[1-6]$/.test(v) ? parseInt(v, 10) : 0;
      if (lvl) {
        const text = (await renderInlines(kids, ctx)).trim();
        if (!text) return '';
        ctx.stats.headings++;
        return '#'.repeat(lvl) + ' ' + text;
      }
    }
  }

  // 编号 / 项目列表
  if (pPr && findChild(pPr, 'numPr')) {
    const numPr = findChild(pPr, 'numPr');
    const ilvl = parseInt(attr(findChild(numPr, 'ilvl'), '@_val') || '0', 10);
    const numId = attr(findChild(numPr, 'numId'), '@_val');
    const fmt = ctx.numbering.format(numId, ilvl);
    const text = (await renderInlines(kids, ctx)).trim();
    if (!text) return '';
    const indent = '  '.repeat(ilvl);
    if (fmt === 'bullet') return indent + '- ' + text;
    const key = numId + ':' + ilvl;
    ctx.counters[key] = (ctx.counters[key] || 0) + 1;
    return indent + ctx.counters[key] + '. ' + text;
  }

  // 普通段落
  const text = (await renderInlines(kids, ctx)).trim();
  return text || '';
}

// ===================== 表格渲染 =====================
function escapeCell(s) {
  return String(s || '')
    .replace(/\\/g, '\\\\')
    .replace(/\|/g, '\\|')
    .replace(/\r?\n/g, '<br>');
}

// 嵌套表格（表格单元格里再套表格）：Markdown pipe 表格无法真正嵌套，
// 按「行间 <br>、列间 |」扁平化，外层 escapeCell 会把 | 转义，内容完整保留。
async function renderNestedTable(tblNode, ctx) {
  ctx.stats.tables++;
  const trs = childrenOf(tblNode).filter((c) => nodeTag(c) === 'tr');
  const rows = [];
  for (const tr of trs) {
    const cells = [];
    for (const tc of childrenOf(tr).filter((x) => nodeTag(x) === 'tc')) {
      cells.push(await renderCellText(tc, ctx));
    }
    rows.push(cells.join(' | '));
  }
  return rows.join('<br>');
}

async function renderCellText(tc, ctx) {
  const parts = [];
  for (const c of childrenOf(tc)) {
    const t = nodeTag(c);
    if (t === 'p') {
      const pPr = findChild(c, 'pPr');
      const kids = childrenOf(c).filter((x) => nodeTag(x) !== 'pPr');
      const txt = (await renderInlines(kids, ctx)).trim();
      if (txt) parts.push(txt);
    } else if (t === 'tbl') {
      const nested = await renderNestedTable(c, ctx);
      if (nested) parts.push(nested);
    }
  }
  return parts.join('<br>');
}

// 单格表格（1 行 1 单元格）判断：飞书把「同步块」导出成这种包裹结构
function isSingleCellTable(tblNode) {
  const trs = childrenOf(tblNode).filter((c) => nodeTag(c) === 'tr');
  if (trs.length !== 1) return false;
  return childrenOf(trs[0]).filter((c) => nodeTag(c) === 'tc').length === 1;
}
function singleCellOf(tblNode) {
  const tr = findChild(tblNode, 'tr');
  return tr ? findChild(tr, 'tc') : null;
}

async function renderTable(tblNode, ctx) {
  const trs = childrenOf(tblNode).filter((c) => nodeTag(c) === 'tr');
  const grid = [];
  let ncols = 0;
  for (const tr of trs) {
    const row = [];
    let c = 0;
    for (const tc of childrenOf(tr).filter((x) => nodeTag(x) === 'tc')) {
      const tcPr = findChild(tc, 'tcPr');
      const span = parseInt(attr(tcPr ? findChild(tcPr, 'gridSpan') : null, '@_val') || '1', 10);
      const vm = tcPr ? findChild(tcPr, 'vMerge') : null;
      const vmVal = vm ? attr(vm, '@_val') || 'continue' : null;
      const text = await renderCellText(tc, ctx);
      for (let k = 0; k < span; k++) {
        if (k === 0 && vmVal !== 'continue') row[c + k] = text;
        else if (row[c + k] === undefined) row[c + k] = '';
      }
      c += span;
    }
    grid.push(row);
    if (c > ncols) ncols = c;
  }
  for (const row of grid) {
    for (let i = 0; i < ncols; i++) if (row[i] === undefined) row[i] = '';
    row.length = ncols;
  }
  if (!grid.length) return '';

  const lines = [];
  lines.push('| ' + grid[0].map(escapeCell).join(' | ') + ' |');
  lines.push('| ' + grid[0].map(() => '---').join(' | ') + ' |');
  for (let i = 1; i < grid.length; i++) {
    lines.push('| ' + grid[i].map(escapeCell).join(' | ') + ' |');
  }
  return lines.join('\n');
}

// ===================== 嵌入 Excel → Markdown 表格 =====================
function parseCellRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref || '');
  if (!m) return { r: 0, col: 0 };
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { r: parseInt(m[2], 10), col };
}
function parseMergeRef(ref) {
  const m = /^([A-Z]+\d+):([A-Z]+\d+)$/.exec(ref || '');
  if (!m) return null;
  const a = parseCellRef(m[1]);
  const b = parseCellRef(m[2]);
  return { r1: a.r, c1: a.col, r2: b.r, c2: b.col };
}

function parseSharedStrings(xml) {
  const out = [];
  if (!xml) return out;
  const tree = new XMLParser(XML_OPTS).parse(xml);
  const root = tree.find((n) => nodeTag(n) === 'sst');
  if (!root) return out;
  for (const si of childrenOf(root).filter((c) => nodeTag(c) === 'si')) {
    let s = '';
    for (const t of findAllDescendants(si, 't')) s += textOf(t);
    out.push(s);
  }
  return out;
}

function pickSheet(zip) {
  if (zip.file('xl/worksheets/sheet1.xml')) return 'xl/worksheets/sheet1.xml';
  const names = Object.keys(zip.files)
    .filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))
    .sort();
  return names[0] || 'xl/worksheets/sheet1.xml';
}

function parseSheet(xml, shared) {
  const tree = new XMLParser(XML_OPTS).parse(xml);
  const root = tree.find((n) => nodeTag(n) === 'worksheet');
  if (!root) return [];
  const grid = {}; // "r,c" → 文本
  let maxR = 0;
  let maxC = 0;
  const sheetData = findChild(root, 'sheetData');
  if (sheetData) {
    for (const row of childrenOf(sheetData).filter((c) => nodeTag(c) === 'row')) {
      const rIdx = parseInt(attr(row, '@_r') || '0', 10);
      for (const c of childrenOf(row).filter((x) => nodeTag(x) === 'c')) {
        const { r, col } = parseCellRef(attr(c, '@_r') || '');
        if (r <= 0 || col <= 0) continue;
        const t = attr(c, '@_t');
        let val = '';
        if (t === 's') {
          const v = findChild(c, 'v');
          const idx = v ? parseInt(textOf(v), 10) : -1;
          val = idx >= 0 && shared[idx] != null ? shared[idx] : '';
        } else if (t === 'inlineStr') {
          const is = findChild(c, 'is');
          if (is) {
            let s = '';
            for (const tt of findAllDescendants(is, 't')) s += textOf(tt);
            val = s;
          }
        } else {
          const v = findChild(c, 'v');
          val = v ? textOf(v) : '';
        }
        grid[r + ',' + col] = val;
        if (r > maxR) maxR = r;
        if (col > maxC) maxC = col;
      }
    }
  }

  const merges = [];
  const mergeCells = findChild(root, 'mergeCells');
  if (mergeCells) {
    for (const mc of childrenOf(mergeCells).filter((c) => nodeTag(c) === 'mergeCell')) {
      const m = parseMergeRef(attr(mc, '@_ref') || '');
      if (m) {
        merges.push(m);
        if (m.r2 > maxR) maxR = m.r2;
        if (m.c2 > maxC) maxC = m.c2;
      }
    }
  }

  if (maxR <= 0 || maxC <= 0) return [];
  const rows = [];
  for (let r = 1; r <= maxR; r++) {
    const row = [];
    for (let c = 1; c <= maxC; c++) row.push(grid[r + ',' + c] || '');
    rows.push(row);
  }
  // 合并区只保留左上角，其余留空（pipe 表格不支持 colspan/rowspan，留空是标准做法）
  for (const m of merges) {
    for (let r = m.r1; r <= m.r2; r++) {
      for (let c = m.c1; c <= m.c2; c++) {
        if (r === m.r1 && c === m.c1) continue;
        rows[r - 1][c - 1] = '';
      }
    }
  }
  return trimGrid(rows);
}

function trimGrid(rows) {
  if (!rows.length) return rows;
  const H = rows.length;
  const W = rows[0].length;
  const rowEmpty = (i) => rows[i].every((v) => v === '');
  const colEmpty = (j) => rows.every((r) => r[j] === '');
  let r1 = 0;
  let r2 = H - 1;
  let c1 = 0;
  let c2 = W - 1;
  while (r1 <= r2 && rowEmpty(r1)) r1++;
  while (r2 >= r1 && rowEmpty(r2)) r2--;
  while (c1 <= c2 && colEmpty(c1)) c1++;
  while (c2 >= c1 && colEmpty(c2)) c2--;
  if (r1 > r2 || c1 > c2) return [];
  return rows.slice(r1, r2 + 1).map((r) => r.slice(c1, c2 + 1));
}

function gridToMarkdown(rows) {
  const lines = [];
  lines.push('| ' + rows[0].map(escapeCell).join(' | ') + ' |');
  lines.push('| ' + rows[0].map(() => '---').join(' | ') + ' |');
  for (let i = 1; i < rows.length; i++) {
    lines.push('| ' + rows[i].map(escapeCell).join(' | ') + ' |');
  }
  return lines.join('\n');
}

async function renderXlsx(buf, label) {
  const zip = await JSZip.loadAsync(buf);
  const shared = parseSharedStrings(await zipString(zip, 'xl/sharedStrings.xml'));
  const sheetXml = await zipString(zip, pickSheet(zip));
  const rows = parseSheet(sheetXml, shared);
  if (!rows.length) return '';
  return `**嵌入表格（${label}）**\n\n${gridToMarkdown(rows)}`;
}

// ===================== Word 表格 → HTML（colspan/rowspan + 底纹 + 字体颜色） =====================
async function renderHtmlCellText(tc, ctx) {
  const parts = [];
  for (const c of childrenOf(tc)) {
    const t = nodeTag(c);
    if (t === 'p') {
      const pPr = findChild(c, 'pPr');
      const kids = childrenOf(c).filter((x) => nodeTag(x) !== 'pPr');
      const txt = (await renderHtmlInlines(kids, ctx)).trim();
      if (txt) parts.push(txt);
    } else if (t === 'tbl') {
      const nested = await renderHtmlTable(c, ctx);
      if (nested) {
        ctx.stats.tables++;
        parts.push(nested);
      }
    }
  }
  return parts.join('<br>');
}

// 统计从某行起，同一列位置连续出现 vMerge="continue" 的行数（用于计算 rowspan）
function countVMergeContinues(rows, r, startCol) {
  let n = 0;
  for (let rr = r + 1; rr < rows.length; rr++) {
    const hit = rows[rr].some((c) => c.start === startCol && c.vMerge === 'continue');
    if (!hit) break;
    n++;
  }
  return n;
}

async function renderHtmlTable(tblNode, ctx) {
  const trs = childrenOf(tblNode).filter((c) => nodeTag(c) === 'tr');
  const rows = [];
  for (const tr of trs) {
    const cells = [];
    let c = 0;
    for (const tc of childrenOf(tr).filter((x) => nodeTag(x) === 'tc')) {
      const tcPr = findChild(tc, 'tcPr');
      const span = parseInt(attr(tcPr ? findChild(tcPr, 'gridSpan') : null, '@_val') || '1', 10);
      const vm = tcPr ? findChild(tcPr, 'vMerge') : null;
      const vmVal = vm ? (attr(vm, '@_val') || 'restart') : null;
      const fill = cellFillColor(tcPr);
      const html = await renderHtmlCellText(tc, ctx);
      cells.push({ start: c, colspan: span, vMerge: vmVal, fill, html });
      c += span;
    }
    rows.push(cells);
  }
  if (!rows.length) return '';

  // 前向填充的逻辑网格（纵向/横向合并均铺开），用于表头识别与字段筛选值
  let W = 0;
  for (const r of rows) for (const cell of r) W = Math.max(W, cell.start + cell.colspan);
  const grid = rows.map(() => new Array(W).fill(''));
  for (let r = 0; r < rows.length; r++) {
    for (const cell of rows[r]) {
      if (cell.vMerge === 'continue') continue; // 由上方 restart 前向填充
      const text = htmlToText(cell.html);
      const spanRows = cell.vMerge === 'restart' ? 1 + countVMergeContinues(rows, r, cell.start) : 1;
      for (let rr = r; rr < r + spanRows && rr < rows.length; rr++) {
        for (let cc = cell.start; cc < cell.start + cell.colspan; cc++) grid[rr][cc] = text;
      }
    }
  }

  // 标题行 = 第一行所有列均非空、且至少含 2 个不同字段名（跳过整行合并的单值标题行）
  let headerIdx = -1;
  for (let r = 0; r < grid.length; r++) {
    if (grid[r].every((v) => v !== '')) {
      if (new Set(grid[r]).size >= 2) { headerIdx = r; break; }
    }
  }

  // 筛选列 = 标题行非空的列；数据行 = 标题行以下；超过 15 行且存在筛选列 → 可筛选
  const filterCols = [];
  const filterLabels = [];
  if (headerIdx >= 0) {
    for (let c = 0; c < W; c++) {
      const label = grid[headerIdx][c].trim();
      if (label) { filterCols.push(c); filterLabels.push(label); }
    }
  }
  // 筛选项列数上限 25 列（超出取前 25 列）
  if (filterCols.length > 25) { filterCols.length = 25; filterLabels.length = 25; }
  const dataRowCount = headerIdx >= 0 ? grid.length - headerIdx - 1 : 0;
  const filterable = dataRowCount > 15 && filterCols.length > 0;

  const tableAttrs = [];
  if (filterable) {
    tableAttrs.push('data-filterable="1"');
    tableAttrs.push(`data-filter-meta='${escapeHtml(JSON.stringify({ cols: filterCols, labels: filterLabels }))}'`);
  }

  const out = [tableAttrs.length ? `<table ${tableAttrs.join(' ')}>` : '<table>'];
  for (let r = 0; r < rows.length; r++) {
    const rowAttrs = [];
    if (filterable && r > headerIdx) {
      for (const col of filterCols) rowAttrs.push(`data-f${col}="${escapeHtml(grid[r][col])}"`);
    }
    out.push(rowAttrs.length ? `<tr ${rowAttrs.join(' ')}>` : '<tr>');
    for (const cell of rows[r]) {
      if (cell.vMerge === 'continue') continue; // 已被上方 rowspan 覆盖
      let rowspan = 1;
      if (cell.vMerge === 'restart') rowspan = 1 + countVMergeContinues(rows, r, cell.start);
      const attrs = [];
      if (cell.colspan > 1) attrs.push(`colspan="${cell.colspan}"`);
      if (rowspan > 1) attrs.push(`rowspan="${rowspan}"`);
      attrs.push(`style="${tdStyle(cell.fill, '')}"`);
      out.push(`<td ${attrs.join(' ')}>${cell.html || '&nbsp;'}</td>`);
    }
    out.push('</tr>');
  }
  out.push('</table>');
  return out.join('\n');
}

// ===================== 嵌入 Excel → HTML（颜色 + 合并单元格） =====================
function parseStyles(xml) {
  const fills = [];
  const fonts = [];
  const cellXfs = [];
  if (!xml) return { fills, fonts, cellXfs };
  const tree = new XMLParser(XML_OPTS).parse(xml);
  const root = tree.find((n) => nodeTag(n) === 'styleSheet');
  if (!root) return { fills, fonts, cellXfs };
  const fillsEl = findChild(root, 'fills');
  if (fillsEl) {
    for (const fill of childrenOf(fillsEl).filter((c) => nodeTag(c) === 'fill')) {
      const pf = findChild(fill, 'patternFill');
      if (!pf) { fills.push(null); continue; }
      const fg = findChild(pf, 'fgColor');
      let color = fg ? argbToCss(attr(fg, '@_rgb')) : '';
      if (!color) {
        const bg = findChild(pf, 'bgColor');
        color = bg ? argbToCss(attr(bg, '@_rgb')) : '';
      }
      fills.push(color || null);
    }
  }
  const fontsEl = findChild(root, 'fonts');
  if (fontsEl) {
    for (const font of childrenOf(fontsEl).filter((c) => nodeTag(c) === 'font')) {
      const colorEl = findChild(font, 'color');
      fonts.push(colorEl ? argbToCss(attr(colorEl, '@_rgb')) : null);
    }
  }
  const cellXfsEl = findChild(root, 'cellXfs');
  if (cellXfsEl) {
    for (const xf of childrenOf(cellXfsEl).filter((c) => nodeTag(c) === 'xf')) {
      const fillId = parseInt(attr(xf, '@_fillId') || '0', 10);
      const fontId = parseInt(attr(xf, '@_fontId') || '0', 10);
      cellXfs.push({ fill: fills[fillId] || null, font: fonts[fontId] || null });
    }
  }
  return { fills, fonts, cellXfs };
}

function parseSheetHtml(xml, shared, cellXfs) {
  const tree = new XMLParser(XML_OPTS).parse(xml);
  const root = tree.find((n) => nodeTag(n) === 'worksheet');
  if (!root) return null;
  const grid = {};
  let maxR = 0;
  let maxC = 0;
  const sheetData = findChild(root, 'sheetData');
  if (sheetData) {
    for (const row of childrenOf(sheetData).filter((c) => nodeTag(c) === 'row')) {
      for (const c of childrenOf(row).filter((x) => nodeTag(x) === 'c')) {
        const { r, col } = parseCellRef(attr(c, '@_r') || '');
        if (r <= 0 || col <= 0) continue;
        const t = attr(c, '@_t');
        let val = '';
        if (t === 's') {
          const v = findChild(c, 'v');
          const idx = v ? parseInt(textOf(v), 10) : -1;
          val = idx >= 0 && shared[idx] != null ? shared[idx] : '';
        } else if (t === 'inlineStr') {
          const is = findChild(c, 'is');
          if (is) {
            let s = '';
            for (const tt of findAllDescendants(is, 't')) s += textOf(tt);
            val = s;
          }
        } else {
          const v = findChild(c, 'v');
          val = v ? textOf(v) : '';
        }
        const si = parseInt(attr(c, '@_s') || '0', 10);
        const st = cellXfs[si] || {};
        grid[r + ',' + col] = { text: val, fill: st.fill || null, font: st.font || null };
        if (r > maxR) maxR = r;
        if (col > maxC) maxC = col;
      }
    }
  }
  const merges = [];
  const mergeCells = findChild(root, 'mergeCells');
  if (mergeCells) {
    for (const mc of childrenOf(mergeCells).filter((c) => nodeTag(c) === 'mergeCell')) {
      const m = parseMergeRef(attr(mc, '@_ref') || '');
      if (m) {
        merges.push(m);
        if (m.r2 > maxR) maxR = m.r2;
        if (m.c2 > maxC) maxC = m.c2;
      }
    }
  }
  if (maxR <= 0 || maxC <= 0) return null;
  const rows = [];
  for (let r = 1; r <= maxR; r++) {
    const row = [];
    for (let c = 1; c <= maxC; c++) row.push(grid[r + ',' + c] || { text: '', fill: null, font: null });
    rows.push(row);
  }
  return { rows, merges };
}

function trimSheetHtml(rows, merges) {
  if (!rows.length) return { rows, merges };
  const H = rows.length;
  const W = rows[0].length;
  const empty = (cell) => !cell || cell.text === '';
  const rowEmpty = (i) => rows[i].every(empty);
  const colEmpty = (j) => rows.every((r) => empty(r[j]));
  let r1 = 0;
  let r2 = H - 1;
  let c1 = 0;
  let c2 = W - 1;
  while (r1 <= r2 && rowEmpty(r1)) r1++;
  while (r2 >= r1 && rowEmpty(r2)) r2--;
  while (c1 <= c2 && colEmpty(c1)) c1++;
  while (c2 >= c1 && colEmpty(c2)) c2--;
  if (r1 > r2 || c1 > c2) return { rows: [], merges: [] };
  const newRows = rows.slice(r1, r2 + 1).map((r) => r.slice(c1, c2 + 1));
  const newMerges = merges
    .map((m) => ({
      r1: Math.max(m.r1, r1 + 1),
      c1: Math.max(m.c1, c1 + 1),
      r2: Math.min(m.r2, r2 + 1),
      c2: Math.min(m.c2, c2 + 1)
    }))
    .filter((m) => m.r1 <= m.r2 && m.c1 <= m.c2);
  return { rows: newRows, merges: newMerges };
}

function renderXlsxHtmlTable(rows, merges) {
  const H = rows.length;
  const W = rows[0].length;
  const covered = {};
  for (const m of merges) {
    for (let r = m.r1; r <= m.r2; r++) {
      for (let c = m.c1; c <= m.c2; c++) {
        if (r === m.r1 && c === m.c1) continue;
        covered[r + ',' + c] = true;
      }
    }
  }

  // 前向填充的逻辑网格（合并区域用左上角值铺满、空白规范化），用于表头识别与字段筛选值
  const norm = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  const grid = rows.map((row) => row.map((cell) => norm(cell && cell.text)));
  for (const m of merges) {
    const val = grid[m.r1 - 1][m.c1 - 1];
    for (let r = m.r1 - 1; r < m.r2; r++) {
      for (let c = m.c1 - 1; c < m.c2; c++) grid[r][c] = val;
    }
  }

  // 标题行 = 第一行所有列均非空、且至少含 2 个不同字段名（跳过整行合并的单值标题行）
  let headerIdx = -1;
  for (let r = 0; r < H; r++) {
    if (grid[r].every((v) => v !== '')) {
      if (new Set(grid[r]).size >= 2) { headerIdx = r; break; }
    }
  }

  // 筛选列 = 标题行非空的列；数据行 = 标题行以下；超过 15 行且存在筛选列 → 可筛选
  const filterCols = [];
  const filterLabels = [];
  if (headerIdx >= 0) {
    for (let c = 0; c < W; c++) {
      if (grid[headerIdx][c] !== '') { filterCols.push(c); filterLabels.push(grid[headerIdx][c]); }
    }
  }
  // 筛选项列数上限 25 列（超出取前 25 列）
  if (filterCols.length > 25) { filterCols.length = 25; filterLabels.length = 25; }
  const dataRowCount = headerIdx >= 0 ? H - headerIdx - 1 : 0;
  const filterable = dataRowCount > 15 && filterCols.length > 0;

  const tableAttrs = [];
  if (filterable) {
    tableAttrs.push('data-filterable="1"');
    tableAttrs.push(`data-filter-meta='${escapeHtml(JSON.stringify({ cols: filterCols, labels: filterLabels }))}'`);
  }

  const lines = [tableAttrs.length ? `<table ${tableAttrs.join(' ')}>` : '<table>'];
  for (let r = 0; r < H; r++) {
    const rowAttrs = [];
    if (filterable && r > headerIdx) {
      for (const col of filterCols) rowAttrs.push(`data-f${col}="${escapeHtml(grid[r][col])}"`);
    }
    lines.push(rowAttrs.length ? `<tr ${rowAttrs.join(' ')}>` : '<tr>');
    for (let c = 0; c < W; c++) {
      const key = (r + 1) + ',' + (c + 1);
      if (covered[key]) continue;
      const cell = rows[r][c];
      let colspan = 1;
      let rowspan = 1;
      for (const m of merges) {
        if (m.r1 === r + 1 && m.c1 === c + 1) {
          colspan = m.c2 - m.c1 + 1;
          rowspan = m.r2 - m.r1 + 1;
          break;
        }
      }
      const attrs = [];
      if (colspan > 1) attrs.push(`colspan="${colspan}"`);
      if (rowspan > 1) attrs.push(`rowspan="${rowspan}"`);
      attrs.push(`style="${tdStyle(cell.fill, cell.font)}"`);
      const cellHtml = escapeHtml(cell.text).replace(/\r?\n/g, '<br>');
      lines.push(`<td ${attrs.join(' ')}>${cellHtml || '&nbsp;'}</td>`);
    }
    lines.push('</tr>');
  }
  lines.push('</table>');
  return lines.join('\n');
}

async function renderXlsxHtml(buf, label) {
  const zip = await JSZip.loadAsync(buf);
  const shared = parseSharedStrings(await zipString(zip, 'xl/sharedStrings.xml'));
  const sheetXml = await zipString(zip, pickSheet(zip));
  const styles = parseStyles(await zipString(zip, 'xl/styles.xml'));
  const parsed = parseSheetHtml(sheetXml, shared, styles.cellXfs);
  if (!parsed || !parsed.rows.length) return '';
  const { rows, merges } = trimSheetHtml(parsed.rows, parsed.merges);
  if (!rows.length) return '';
  const table = renderXlsxHtmlTable(rows, merges);
  return `<p><strong>嵌入表格（${escapeHtml(label)}）</strong></p>\n${table}`;
}

// ===================== 文档主体遍历 =====================
async function renderBlocks(children, ctx, out) {
  for (const c of children) {
    const t = nodeTag(c);
    if (t === 'p') {
      const md = await renderParagraph(c, ctx);
      if (md) out.push(md);
    } else if (t === 'tbl') {
      // 飞书「同步块」导出为单格表格：把单格表格当作透明容器，直接渲染其内部内容，
      // 避免把真正的表格/段落扁平化成一个大单元格。
      if (isSingleCellTable(c)) {
        const tc = singleCellOf(c);
        if (tc) await renderBlocks(childrenOf(tc), ctx, out);
      } else if (ctx.htmlTables) {
        const html = await renderHtmlTable(c, ctx);
        if (html) out.push(html);
      } else {
        const md = await renderTable(c, ctx);
        if (md) out.push(md);
      }
      ctx.stats.tables++;
    } else if (t === 'sdt') {
      const sc = findChild(c, 'sdtContent');
      if (sc) await renderBlocks(childrenOf(sc), ctx, out);
    }
    // sectPr / bookmarkStart 等忽略
  }
}

// ===================== 主入口 =====================
async function convertDocxToMarkdown(docxPath, opts) {
  opts = opts || {};
  const data = fs.readFileSync(docxPath);
  const zip = await JSZip.loadAsync(data);
  const docXml = await zipString(zip, 'word/document.xml');
  if (!docXml) throw new Error('不是有效的 .docx 文件（缺少 word/document.xml）');

  const rels = parseRels(await zipString(zip, 'word/_rels/document.xml.rels'));
  const numbering = parseNumbering(await zipString(zip, 'word/numbering.xml'));

  const docTree = new XMLParser(XML_OPTS).parse(docXml);
  const docNode = docTree.find((n) => nodeTag(n) === 'document');
  if (!docNode) throw new Error('解析 .docx 失败：找不到 document 节点');
  const body = findChild(docNode, 'body');
  if (!body) throw new Error('解析 .docx 失败：找不到 body 节点');

  const ctx = {
    zip,
    rels,
    numbering,
    images: new Map(),
    imageOrder: [],
    counters: {},
    htmlTables: !!opts.htmlTables,
    stats: { headings: 0, tables: 0, embeddedSheets: 0, images: 0, hyperlinks: 0 }
  };

  const blocks = [];
  await renderBlocks(childrenOf(body), ctx, blocks);
  const markdown = blocks.filter((b) => b != null && b !== '').join('\n\n') + '\n';
  const images = ctx.imageOrder.map((fn) => {
    const img = ctx.images.get(fn);
    return { filename: fn, buffer: img.buffer, mime: img.mime };
  });

  return { markdown, images, stats: ctx.stats };
}

// Word → HTML：与 markdown 版共用同一套遍历，区别是表格/嵌入 Excel 直接输出 HTML（保留颜色与合并）
async function convertDocxToHtml(docxPath) {
  return convertDocxToMarkdown(docxPath, { htmlTables: true });
}

module.exports = { convertDocxToMarkdown, convertDocxToHtml };
