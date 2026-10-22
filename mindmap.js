// ===================== Excel/CSV → SVG 思维导图 =====================
// 解析 .xlsx（复用 docx2md.js 的 jszip + fast-xml-parser 原语）与 .csv（手写 RFC-4180），
// 把表格数据构造成树，再用手写 SVG 布局引擎产出放射状思维导图。
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { XMLParser } = require('fast-xml-parser');
const {
  XML_OPTS, nodeTag, attr, childrenOf, findChild,
  zipString, parseRels, parseCellRef, parseSharedStrings, parseSheet, trimGrid
} = require('./docx2md');

// ===================== 列号 / A1 记法 =====================
function colLetterToNum(s) {
  let n = 0;
  for (const ch of String(s).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
function numToColLetter(n) {
  let s = '';
  let x = n;
  while (x > 0) {
    const m = (x - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

// ===================== 范围解析 =====================
// 接受 null（自动）、"A1:B20" / "B5" 字符串、或 {r1,c1,r2,c2}；1-based 含端点。
function normalizeRange(input, fullRows, fullCols) {
  if (input == null || input === '') return null;
  let r1, c1, r2, c2;
  if (typeof input === 'string') {
    const m = /^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/.exec(input.trim());
    if (!m) return null;
    c1 = colLetterToNum(m[1]); r1 = parseInt(m[2], 10);
    if (m[3]) { c2 = colLetterToNum(m[3]); r2 = parseInt(m[4], 10); }
    else { c2 = c1; r2 = r1; }
  } else {
    r1 = input.r1; c1 = input.c1; r2 = input.r2; c2 = input.c2;
  }
  if (!r1 || !c1 || !r2 || !c2) return null;
  if (r1 > r2) { const t = r1; r1 = r2; r2 = t; }
  if (c1 > c2) { const t = c1; c1 = c2; c2 = t; }
  const fr = fullRows || 1, fc = fullCols || 1;
  r1 = Math.max(1, Math.min(r1, fr)); c1 = Math.max(1, Math.min(c1, fc));
  r2 = Math.max(1, Math.min(r2, fr)); c2 = Math.max(1, Math.min(c2, fc));
  return { r1, c1, r2, c2 };
}

// ===================== CSV 编码探测 + RFC-4180 解析 =====================
function decodeBuffer(buf) {
  if (buf.length >= 3 && buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) return buf.slice(3).toString('utf-8');
  if (buf.length >= 2 && buf[0] === 0xFF && buf[1] === 0xFE) return new TextDecoder('utf-16le').decode(buf);
  if (buf.length >= 2 && buf[0] === 0xFE && buf[1] === 0xFF) return new TextDecoder('utf-16be').decode(buf);
  const utf8 = buf.toString('utf-8');
  if (utf8.includes('�')) {
    try { return new TextDecoder('gb18030').decode(buf); } catch (e) { /* 忽略 */ }
  }
  return utf8;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  const n = text.length;
  while (i < n) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
        inQuotes = false; i++; continue;
      }
      field += ch; i++; continue;
    }
    if (ch === '"') { inQuotes = true; i++; continue; }
    if (ch === ',') { row.push(field); field = ''; i++; continue; }
    if (ch === '\r') { if (text[i + 1] === '\n') i++; row.push(field); field = ''; rows.push(row); row = []; i++; continue; }
    if (ch === '\n') { row.push(field); field = ''; rows.push(row); row = []; i++; continue; }
    field += ch; i++;
  }
  if (field !== '' || row.length > 0) { row.push(field); rows.push(row); }
  return rows;
}

function padRows(rows) {
  const W = rows.reduce((m, r) => Math.max(m, r.length), 0);
  return rows.map((r) => { const c = r.slice(); while (c.length < W) c.push(''); return c; });
}

function sliceGrid(rows, range) {
  if (!rows.length || !range) return rows;
  const sliced = rows.slice(range.r1 - 1, range.r2).map((r) => r.slice(range.c1 - 1, range.c2));
  return trimGrid(sliced);
}

// ===================== XLSX 解析（复用 docx2md 原语，新增全表枚举） =====================
function resolveXlsxTarget(target) {
  if (!target) return '';
  let t = String(target).replace(/\\/g, '/');
  if (t.startsWith('/')) t = t.slice(1);
  else t = 'xl/' + t;
  return path.posix.normalize(t);
}

async function listXlsxSheets(zip) {
  const wbXml = await zipString(zip, 'xl/workbook.xml');
  if (!wbXml) return [];
  const tree = new XMLParser(XML_OPTS).parse(wbXml);
  const root = tree.find((n) => nodeTag(n) === 'workbook');
  if (!root) return [];
  const sheetsEl = findChild(root, 'sheets');
  if (!sheetsEl) return [];
  const rels = parseRels(await zipString(zip, 'xl/_rels/workbook.xml.rels'));
  const sheets = [];
  let idx = 0;
  for (const s of childrenOf(sheetsEl).filter((c) => nodeTag(c) === 'sheet')) {
    const name = attr(s, '@_name') || ('Sheet' + (idx + 1));
    const rId = attr(s, '@_id'); // removeNSPrefix 后 r:id → @_id
    let sheetPath = rId && rels[rId] ? resolveXlsxTarget(rels[rId].target) : '';
    if (!sheetPath) sheetPath = 'xl/worksheets/sheet' + (idx + 1) + '.xml';
    sheets.push({ index: idx, name, path: sheetPath });
    idx++;
  }
  return sheets;
}

async function readXlsxGrid(zip, sheetIndex, range) {
  const sheets = await listXlsxSheets(zip);
  const sheet = sheets[sheetIndex];
  if (!sheet) return { grid: [], error: '工作表不存在' };
  const shared = parseSharedStrings(await zipString(zip, 'xl/sharedStrings.xml'));
  const sheetXml = await zipString(zip, sheet.path);
  if (!sheetXml) return { grid: [], error: '无法读取工作表内容' };
  let rows = parseSheet(sheetXml, shared, { fillMerged: true }); // 已 trim，合并单元格向下/向右填充
  if (range) rows = sliceGrid(rows, normalizeRange(range, rows.length, rows[0] ? rows[0].length : 0));
  return { grid: rows };
}

// ===================== 树构建 =====================
function buildTree(grid, opts) {
  const levels = (opts.levels || []).map(Number);
  const rootCol = (opts.rootCol != null && opts.rootCol >= 0) ? opts.rootCol : -1;
  const rootTitle = opts.rootTitle || '';
  const maxDepth = opts.maxDepth || 8;
  const maxRows = opts.maxRows || 2000;
  const skipHeader = opts.skipHeader !== false; // 默认跳过表头行（第 0 行）
  const stats = { nodes: 0, depth: 0, rowsUsed: 0, truncated: false };

  if (!grid || !grid.length) return { tree: null, stats };

  let rootText = rootTitle || '中心主题';
  if (rootCol >= 0) {
    const headerVal = (grid[0] && (grid[0][rootCol] || '')).trim();
    if (headerVal) rootText = headerVal;
    else {
      for (let r = 1; r < grid.length; r++) {
        const v = grid[r] && (grid[r][rootCol] || '').trim();
        if (v) { rootText = v; break; }
      }
    }
  }

  const root = { text: rootText, depth: 0, children: [], count: 0 };
  const maps = new Map(); // node -> Map(label -> child)
  function getMap(node) {
    let m = maps.get(node);
    if (!m) { m = new Map(); maps.set(node, m); }
    return m;
  }
  function insertPath(parent, p, idx) {
    if (idx >= p.length) return;
    const label = p[idx];
    const m = getMap(parent);
    let child = m.get(label);
    if (!child) {
      child = { text: label, depth: parent.depth + 1, children: [], count: 0 };
      parent.children.push(child);
      m.set(label, child);
    }
    child.count++;
    insertPath(child, p, idx + 1);
  }

  const dataStart = skipHeader ? 1 : 0;
  const limit = Math.min(grid.length, dataStart + maxRows);
  let maxPathLen = 0;
  let rowsUsed = 0;
  for (let r = dataStart; r < limit; r++) {
    const row = grid[r];
    if (!row || !levels.length) continue;
    const raw = [];
    for (let li = 0; li < levels.length && li < maxDepth; li++) {
      const col = levels[li];
      raw.push((row[col] == null ? '' : String(row[col])).trim());
    }
    let a = 0, b = raw.length - 1;
    while (a <= b && raw[a] === '') a++;
    while (b >= a && raw[b] === '') b--;
    if (a > b) continue; // 整行空
    const p = [];
    // 空值生成「（空）」占位节点，保留层级位置（不再截断）
    for (let i = a; i <= b; i++) {
      p.push(raw[i] === '' ? '（空）' : raw[i]);
    }
    if (!p.length) continue;
    insertPath(root, p, 0);
    root.count++;
    rowsUsed++;
    if (p.length > maxPathLen) maxPathLen = p.length;
  }

  stats.rowsUsed = rowsUsed;
  stats.depth = maxPathLen;
  stats.nodes = countNodes(root);
  stats.truncated = grid.length > limit;
  return { tree: root, stats };
}

// ===================== 矩阵模式（列=主体，行=属性） =====================
// 适用于「宽矩阵」：列头=实体（如门店类型），标签列=属性名，单元格=属性值。
// 树：根 → 实体（列头） → 属性（标签） → 值。labelCol 为 -1 时无标签列：每列=实体，单元格=子节点。
function buildMatrixTree(grid, opts) {
  const labelCol = (opts.labelCol != null) ? Number(opts.labelCol) : 0;
  const rootTitle = opts.rootTitle || '中心主题';
  const includeValues = opts.includeValues !== false;
  const maxRows = opts.maxRows || 2000;
  const stats = { nodes: 0, depth: 0, rowsUsed: 0, truncated: false };

  if (!grid || !grid.length || !grid[0]) return { tree: null, stats };
  const header = grid[0];

  const root = { text: rootTitle, depth: 0, children: [], count: 0 };
  const maps = new Map();
  function getMap(node) {
    let m = maps.get(node);
    if (!m) { m = new Map(); maps.set(node, m); }
    return m;
  }
  function insertPath(parent, p, idx) {
    if (idx >= p.length) return;
    const label = p[idx];
    const m = getMap(parent);
    let child = m.get(label);
    if (!child) {
      child = { text: label, depth: parent.depth + 1, children: [], count: 0 };
      parent.children.push(child);
      m.set(label, child);
    }
    child.count++;
    insertPath(child, p, idx + 1);
  }

  // 实体列：优先用调用方指定的顺序列表（支持拖拽排序）；否则除标签列外的所有列
  let entityCols;
  if (Array.isArray(opts.entityCols) && opts.entityCols.length) {
    entityCols = opts.entityCols
      .map(Number)
      .filter((n) => Number.isFinite(n) && n >= 0 && n < header.length && n !== labelCol);
  } else {
    entityCols = [];
    for (let c = 0; c < header.length; c++) {
      if (c !== labelCol) entityCols.push(c);
    }
  }

  const dataStart = 1; // 第 0 行是列头
  const limit = Math.min(grid.length, dataStart + maxRows);
  let maxPathLen = 0;
  let rowsUsed = 0;

  for (const c of entityCols) {
    const entityName = (header[c] == null ? '' : String(header[c])).trim();
    if (!entityName) continue;
    for (let r = dataStart; r < limit; r++) {
      const row = grid[r];
      if (!row) continue;
      const p = [entityName];
      if (labelCol >= 0) {
        const attrName = (row[labelCol] == null ? '' : String(row[labelCol])).trim();
        if (attrName) p.push(attrName);
      }
      const val = (row[c] == null ? '' : String(row[c])).trim();
      if (includeValues && val) p.push(val);
      if (p.length <= 1) continue;
      insertPath(root, p, 0);
      root.count++;
      rowsUsed++;
      if (p.length > maxPathLen) maxPathLen = p.length;
    }
  }

  stats.rowsUsed = rowsUsed;
  stats.depth = maxPathLen;
  stats.nodes = countNodes(root);
  return { tree: root, stats };
}

// ===================== 决策表模式（条件 → 结果） =====================
// 结果列归一化：支持单列 resultCol 或多列级联 resultCols（按顺序组成层级）。
function resultColsOf(opts) {
  const arr = Array.isArray(opts.resultCols)
    ? opts.resultCols
    : (opts.resultCol != null && opts.resultCol !== '' && Number(opts.resultCol) >= 0 ? [opts.resultCol] : []);
  return arr.map(Number).filter((n) => Number.isFinite(n) && n >= 0);
}

// 适用于「决策表」：若干条件列 + 结果列（可为多列级联，如 大类→中类→小类）。
// 空条件 = 不关心（忽略该条件，结果仍作为叶子显示，不截断）；空结果层级 = 忽略该层级。
// view: 'tree' 根→条件值→结果级联；'grouped' 根→结果级联→「条件名=值」。
function buildDecisionTree(grid, opts) {
  const resultCols = resultColsOf(opts);
  const condCols = (opts.condCols || []).map(Number);
  const rootTitle = opts.rootTitle || '中心主题';
  const view = opts.view || 'tree';
  const maxRows = opts.maxRows || 2000;
  const stats = { nodes: 0, depth: 0, rowsUsed: 0, truncated: false };

  if (!grid || !grid.length || !grid[0]) return { tree: null, stats };
  const header = grid[0];

  const root = { text: rootTitle, depth: 0, children: [], count: 0 };
  const maps = new Map();
  function getMap(node) {
    let m = maps.get(node);
    if (!m) { m = new Map(); maps.set(node, m); }
    return m;
  }
  function insertPath(parent, p, idx) {
    if (idx >= p.length) return null;
    const label = p[idx];
    const m = getMap(parent);
    let child = m.get(label);
    if (!child) {
      child = { text: label, depth: parent.depth + 1, children: [], count: 0 };
      parent.children.push(child);
      m.set(label, child);
    }
    child.count++;
    if (idx === p.length - 1) return child;
    return insertPath(child, p, idx + 1);
  }

  const dataStart = 1; // 第 0 行是表头
  const limit = Math.min(grid.length, dataStart + maxRows);
  let maxPathLen = 0;
  let rowsUsed = 0;

  for (let r = dataStart; r < limit; r++) {
    const row = grid[r];
    if (!row) continue;
    const resultPath = [];     // 结果值序列（分组视图用）
    const resultPathCols = []; // 对应结果列下标（树视图用于带出结果字段名）
    for (const c of resultCols) { // 结果级联：按顺序取各结果列，跳过空层级
      const v = (row[c] == null ? '' : String(row[c])).trim();
      if (v) { resultPath.push(v); resultPathCols.push(c); }
    }
    if (!resultPath.length) continue; // 结果全空 → 跳过该行

    if (view === 'grouped') {
      // 根 → 结果级联（大类→中类→小类）→ 各条件「名=值」（条件并列，不嵌套成链）
      const resultLeaf = insertPath(root, resultPath, 0);
      let hasCond = false;
      for (const c of condCols) {
        const v = (row[c] == null ? '' : String(row[c])).trim();
        if (!v) continue; // 忽略空条件
        const name = (header[c] == null ? '' : String(header[c])).trim();
        insertPath(resultLeaf, [name ? (name + '=' + v) : v], 0);
        hasCond = true;
      }
      const len = resultPath.length + (hasCond ? 1 : 0);
      if (maxPathLen < len) maxPathLen = len;
    } else {
      const p = [];
      for (const c of condCols) { // 条件：与结果节点样式统一，用「列名=具体值」
        const v = (row[c] == null ? '' : String(row[c])).trim();
        if (!v) continue;
        const name = (header[c] == null ? '' : String(header[c])).trim();
        p.push(name ? (name + '=' + v) : v);
      }
      for (let i = 0; i < resultPath.length; i++) { // 叶子 = 结果级联，带出结果字段名
        const c = resultPathCols[i];
        const name = (header[c] == null ? '' : String(header[c])).trim();
        const v = resultPath[i];
        p.push(name ? (name + '=' + v) : v);
      }
      insertPath(root, p, 0);
      if (p.length > maxPathLen) maxPathLen = p.length;
    }
    root.count++;
    rowsUsed++;
  }

  stats.rowsUsed = rowsUsed;
  stats.depth = maxPathLen;
  stats.nodes = countNodes(root);
  stats.truncated = grid.length > limit;
  return { tree: root, stats };
}

// ===================== 决策表冲突校验 =====================
// 检测「同一条件组合（忽略空条件）对应多个不同结果」的冲突行。
// 返回 [{ conds, results: [{ result, rows }] }]，rows 为 1 起算的行号（表头=第 1 行）；
// 把每个结果与产生它的具体行号对应起来，便于直接回查源数据。
function findDecisionConflicts(grid, opts) {
  const resultCols = resultColsOf(opts);
  const condCols = (opts.condCols || []).map(Number);
  if (!grid || !grid.length || !grid[0]) return [];
  const header = grid[0];
  const byKey = new Map(); // key -> { conds, groups: Map(result -> rows[]) }
  for (let r = 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row) continue;
    const resultPath = [];
    for (const c of resultCols) { // 结果级联：冲突按「层级1 / 层级2 / …」判定
      const v = (row[c] == null ? '' : String(row[c])).trim();
      if (v) resultPath.push(v);
    }
    const result = resultPath.join(' / ');
    if (!result) continue;
    const condParts = [];
    for (const c of condCols) {
      const v = (row[c] == null ? '' : String(row[c])).trim();
      if (!v) continue; // 忽略空条件
      const name = (header[c] == null ? '' : String(header[c])).trim();
      condParts.push(name ? (name + '=' + v) : v);
    }
    if (!condParts.length) continue; // 无条件组合的行不参与冲突校验
    const key = condParts.join('\u0001');
    let entry = byKey.get(key);
    if (!entry) {
      entry = { conds: condParts.join('，'), groups: new Map() };
      byKey.set(key, entry);
    }
    let rows = entry.groups.get(result);
    if (!rows) { rows = []; entry.groups.set(result, rows); }
    rows.push(r + 1);
  }
  const conflicts = [];
  for (const entry of byKey.values()) {
    if (entry.groups.size > 1) {
      conflicts.push({
        conds: entry.conds,
        results: Array.from(entry.groups.entries()).map(([result, rows]) => ({ result, rows }))
      });
    }
  }
  return conflicts;
}

// ===================== 数据筛选过滤 =====================
// filters: [{ col, op, value, logic }]；第 0 行视为表头始终保留。
// op: contains | notContains | equals | notEmpty | empty
// logic: 与上一条条件的连接符 'and'（默认）| 'or'；首条条件的 logic 忽略。
function rowMatches(row, f) {
  const cell = (row && row[f.col] != null ? String(row[f.col]) : '').trim();
  const c = cell.toLowerCase();
  const v = (f.value != null ? String(f.value) : '').trim().toLowerCase();
  switch (f.op) {
    case 'contains': return v !== '' && c.includes(v);
    case 'notContains': return !(v !== '' && c.includes(v));
    case 'equals': return c === v;
    case 'notEmpty': return c !== '';
    case 'empty': return c === '';
    default: return true;
  }
}

function filterGrid(grid, filters) {
  const active = (filters || []).filter((f) => f && f.col != null && f.col >= 0 && f.op);
  if (!active.length) return grid;
  if (!grid || !grid.length) return grid;
  const out = [grid[0]];
  for (let r = 1; r < grid.length; r++) {
    const row = grid[r];
    let ok = true;
    for (let i = 0; i < active.length; i++) {
      const pass = rowMatches(row, active[i]);
      if (i === 0) ok = pass;
      else if (active[i].logic === 'or') ok = ok || pass;
      else ok = ok && pass;
    }
    if (ok) out.push(row);
  }
  return out;
}

function countNodes(node) {
  let n = 1;
  for (const c of node.children) n += countNodes(c);
  return n;
}

// ===================== CJK 宽度测算 =====================
function textW(str, fs) {
  let w = 0;
  for (const ch of str) {
    const cp = ch.codePointAt(0);
    if ((cp >= 0x4E00 && cp <= 0x9FFF) ||   // CJK 统一汉字
        (cp >= 0x3000 && cp <= 0x303F) ||   // CJK 标点
        (cp >= 0xFF00 && cp <= 0xFFEF) ||   // 全角形式
        (cp >= 0x3040 && cp <= 0x30FF) ||   // 假名
        (cp >= 0xAC00 && cp <= 0xD7AF)) {   // 谚文
      w += fs;
    } else {
      w += 0.55 * fs;
    }
  }
  return w;
}

// ===================== SVG 布局引擎 =====================
const STYLE = {
  0: { fill: '#C08030', stroke: '#8A5A1A', text: '#FFFFFF', edge: '#C08030' },
  1: {
    L: { fill: '#DBEAFE', stroke: '#3B82F6', text: '#1E3A8A', edge: '#93C5FD' },
    R: { fill: '#FEF3C7', stroke: '#F59E0B', text: '#7C2D12', edge: '#FCD34D' }
  },
  2: {
    L: { fill: '#EFF6FF', stroke: '#60A5FA', text: '#1D4ED8', edge: '#93C5FD' },
    R: { fill: '#FFFBEB', stroke: '#FBBF24', text: '#92400E', edge: '#FCD34D' }
  },
  3: {
    L: { fill: '#F0F9FF', stroke: '#38BDF8', text: '#075985', edge: '#93C5FD' },
    R: { fill: '#FFF7ED', stroke: '#FB923C', text: '#7C2D12', edge: '#FCD34D' }
  }
};

function styleFor(node) {
  if (node.depth === 0) return STYLE[0];
  const d = Math.min(node.depth, 3);
  const side = node.dir === -1 ? 'L' : 'R';
  return STYLE[d][side];
}

function escapeXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function wrapLines(text, fs, maxLabelWidth) {
  if (!text) return [''];
  const lines = [];
  let cur = '';
  for (const ch of text) {
    if (cur !== '' && textW(cur + ch, fs) > maxLabelWidth) { lines.push(cur); cur = ch; }
    else cur += ch;
  }
  if (cur !== '' || lines.length === 0) lines.push(cur);
  return lines;
}

function r(n) { return Math.round(n * 10) / 10; }

function measure(node, fs, maxLabelWidth, padX, padY, lineHeight) {
  node.lines = wrapLines(node.label || node.text, fs, maxLabelWidth);
  let w = 0;
  for (const l of node.lines) w = Math.max(w, textW(l, fs));
  node.w = w + 2 * padX;
  node.h = node.lines.length * lineHeight + 2 * padY;
  for (const c of node.children) measure(c, fs, maxLabelWidth, padX, padY, lineHeight);
}

// 后序计算每棵子树的垂直总高（subH）与子树顶相对自身盒顶的偏移（subTop，≤0）。
// 兄弟子树必须按 subH 而非节点自身高度堆叠，否则高个子树会与相邻兄弟重叠。
function computeSubtree(node, vgap) {
  if (!node.children.length) { node.subH = node.h; node.subTop = 0; return; }
  let childrenH = 0;
  for (const c of node.children) { computeSubtree(c, vgap); childrenH += c.subH; }
  childrenH += vgap * (node.children.length - 1);
  const top = Math.min(0, node.h / 2 - childrenH / 2);
  const bottom = Math.max(node.h, node.h / 2 + childrenH / 2);
  node.subH = bottom - top;
  node.subTop = top;
}

function placeSide(children, parent, dir, hgap, vgap) {
  if (!children.length) return;
  let childrenH = 0;
  for (const c of children) childrenH += c.subH;
  childrenH += vgap * (children.length - 1);
  let cursor = parent.y + parent.h / 2 - childrenH / 2;
  for (const c of children) {
    if (dir === 1) c.x = parent.x + parent.w + hgap;
    else c.x = parent.x - hgap - c.w;
    c.y = cursor - c.subTop;
    place(c, dir, hgap, vgap);
    cursor += c.subH + vgap;
  }
}

function place(node, dir, hgap, vgap) {
  node.dir = dir;
  if (!node.children.length) return;
  if (dir === 0) {
    const right = [], left = [];
    node.children.forEach((c, i) => (i % 2 === 0 ? right : left).push(c));
    placeSide(right, node, 1, hgap, vgap);
    placeSide(left, node, -1, hgap, vgap);
  } else {
    placeSide(node.children, node, dir, hgap, vgap);
  }
}

// ===================== 上下布局（td，自上而下） =====================
// 后序计算每棵子树的水平总宽 subW（兄弟按 subW 横向堆叠）
function computeSubtreeW(node, hgap) {
  if (!node.children.length) { node.subW = node.w; return; }
  let childrenW = 0;
  for (const c of node.children) { computeSubtreeW(c, hgap); childrenW += c.subW; }
  childrenW += hgap * (node.children.length - 1);
  node.subW = Math.max(node.w, childrenW);
}

// 上下布局：把子树放到父节点正下方，兄弟按 subW 横向居中排成一行
function placeTopDown(node, hgap, vgap) {
  node.dir = 1; // 统一取右侧配色（上下布局无左右之分）
  if (!node.children.length) return;
  let childrenW = 0;
  for (const c of node.children) childrenW += c.subW;
  childrenW += hgap * (node.children.length - 1);
  const cx = node.x + node.w / 2;
  let cursor = cx - childrenW / 2;
  for (const c of node.children) {
    c.x = cursor + (c.subW - c.w) / 2;
    c.y = node.y + node.h + vgap;
    placeTopDown(c, hgap, vgap);
    cursor += c.subW + hgap;
  }
}

// 生成内嵌在 SVG 里的折叠/展开脚本（参照「项目管理经验.html」：静态布局 + 运行时重排）。
// 树结构在服务端已知，这里把每个节点的 pid/x/w/h/dir 编码成 DATA，脚本据此重排并隐藏子树。
// layout: 'lr' = 左右放射（重排 y）；'td' = 上下（重排 x）。
function collapseScript(layout, hgap, vgap, pad, dataArray) {
  const isTd = layout === 'td';
  return `(function(){
  var svg = document.querySelector('svg');
  if(!svg) return;
  var LAYOUT = '${isTd ? 'td' : 'lr'}';
  var HGAP = ${hgap};
  var VGAP = ${vgap};
  var PAD = ${pad};
  var DATA = ${JSON.stringify(dataArray)};
  var nodes = [], i;
  for(i=0;i<DATA.length;i++){
    var d = DATA[i];
    var n = { id:i, pid:d[0], x:d[1], w:d[2], h:d[3], dir:d[4], hasKids:!!d[5], children:[], parent:null,
      el:document.getElementById('n'+i), edge:document.getElementById('e'+i), badge:document.getElementById('b'+i) };
    nodes.push(n);
  }
  var root = null;
  for(i=0;i<nodes.length;i++){
    var n = nodes[i];
    if(n.pid < 0){ root = n; }
    else { n.parent = nodes[n.pid]; n.parent.children.push(n); }
  }
  if(!root) return;
  var collapsed = {};
  function rr(v){ return Math.round(v*10)/10; }
  function computeSub(n){
    if(collapsed[n.id] || !n.children.length){ n.subH = n.h; n.subTop = 0; n.subW = n.w; return; }
    var j;
    if(LAYOUT === 'td'){
      var cw = 0;
      for(j=0;j<n.children.length;j++){ computeSub(n.children[j]); cw += n.children[j].subW; }
      cw += HGAP * (n.children.length - 1);
      n.subW = Math.max(n.w, cw);
    } else {
      var ch = 0;
      for(j=0;j<n.children.length;j++){ computeSub(n.children[j]); ch += n.children[j].subH; }
      ch += VGAP * (n.children.length - 1);
      var top = Math.min(0, n.h/2 - ch/2);
      var bottom = Math.max(n.h, n.h/2 + ch/2);
      n.subH = bottom - top;
      n.subTop = top;
    }
  }
  function stackX(n){
    if(!n.children.length) return;
    var cw = 0, j;
    for(j=0;j<n.children.length;j++) cw += n.children[j].subW;
    cw += HGAP * (n.children.length - 1);
    var cursor = n.x + n.w/2 - cw/2;
    for(j=0;j<n.children.length;j++){
      var c = n.children[j];
      c.x = cursor + (c.subW - c.w)/2;
      placeX(c);
      cursor += c.subW + HGAP;
    }
  }
  function placeX(n){ if(n.children.length) stackX(n); }
  function stackY(children, parent){
    if(!children.length) return;
    var ch = 0, j;
    for(j=0;j<children.length;j++) ch += children[j].subH;
    ch += VGAP * (children.length - 1);
    var cursor = parent.y + parent.h/2 - ch/2;
    for(j=0;j<children.length;j++){
      var c = children[j];
      c.y = cursor - c.subTop;
      placeY(c);
      cursor += c.subH + VGAP;
    }
  }
  function placeY(n){
    if(!n.children.length) return;
    if(n.dir === 0){
      var right = [], left = [], j;
      for(j=0;j<n.children.length;j++){ if(j % 2 === 0) right.push(n.children[j]); else left.push(n.children[j]); }
      stackY(right, n);
      stackY(left, n);
    } else {
      stackY(n.children, n);
    }
  }
  function drawEdge(n){
    if(!n.edge || !n.parent) return;
    var p = n.parent, x1, x2, y1, y2, dd;
    if(LAYOUT === 'td'){
      x1 = p.x + p.w/2; y1 = p.y + p.h;
      x2 = n.x + n.w/2; y2 = n.y;
      dd = Math.max(16, (y2 - y1) * 0.5);
      n.edge.setAttribute('d', 'M' + rr(x1) + ' ' + rr(y1) + ' C ' + rr(x1) + ' ' + rr(y1+dd) + ', ' + rr(x2) + ' ' + rr(y2-dd) + ', ' + rr(x2) + ' ' + rr(y2));
    } else if(n.dir === -1){
      x1 = p.x; y1 = p.y + p.h/2; x2 = n.x + n.w; y2 = n.y + n.h/2;
      dd = Math.max(16, (x1 - x2) * 0.5);
      n.edge.setAttribute('d', 'M' + rr(x1) + ' ' + rr(y1) + ' C ' + rr(x1-dd) + ' ' + rr(y1) + ', ' + rr(x2+dd) + ' ' + rr(y2) + ', ' + rr(x2) + ' ' + rr(y2));
    } else {
      x1 = p.x + p.w; y1 = p.y + p.h/2; x2 = n.x; y2 = n.y + n.h/2;
      dd = Math.max(16, (x2 - x1) * 0.5);
      n.edge.setAttribute('d', 'M' + rr(x1) + ' ' + rr(y1) + ' C ' + rr(x1+dd) + ' ' + rr(y1) + ', ' + rr(x2-dd) + ' ' + rr(y2) + ', ' + rr(x2) + ' ' + rr(y2));
    }
  }
  function refresh(){
    if(LAYOUT === 'td'){ root.x = 0; root.y = 0; }
    else { root.x = -root.w/2; root.y = -root.h/2; }
    computeSub(root);
    if(LAYOUT === 'td'){ placeX(root); } else { placeY(root); }
    var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    (function walk(n, hidden){
      var vis = !hidden;
      n.el.style.display = vis ? '' : 'none';
      n.el.setAttribute('transform', 'translate(' + rr(n.x) + ' ' + rr(n.y) + ')');
      if(n.edge){ n.edge.style.display = vis ? '' : 'none'; if(vis) drawEdge(n); }
      if(n.badge){
        n.badge.style.display = vis ? '' : 'none';
        if(vis){
          var bx, cy;
          if(LAYOUT === 'td'){ bx = n.x + n.w/2; cy = n.y + n.h + 9; }
          else { bx = (n.dir === -1) ? (n.x - 9) : (n.x + n.w + 9); cy = n.y + n.h/2; }
          var c = n.badge.querySelector('circle');
          var t = n.badge.querySelector('text');
          c.setAttribute('cx', bx); c.setAttribute('cy', cy);
          t.setAttribute('x', bx); t.setAttribute('y', cy);
          t.textContent = collapsed[n.id] ? '+' : '−';
        }
      }
      if(vis){
        minX = Math.min(minX, n.x); minY = Math.min(minY, n.y);
        maxX = Math.max(maxX, n.x + n.w); maxY = Math.max(maxY, n.y + n.h);
      }
      var childHidden = hidden || collapsed[n.id];
      for(var j=0;j<n.children.length;j++) walk(n.children[j], childHidden);
    })(root, false);
    svg.setAttribute('viewBox', rr(minX-PAD) + ' ' + rr(minY-PAD) + ' ' + rr(maxX-minX+2*PAD) + ' ' + rr(maxY-minY+2*PAD));
    svg.setAttribute('width', rr(maxX-minX+2*PAD));
    svg.setAttribute('height', rr(maxY-minY+2*PAD));
  }
  function toggle(n){
    if(collapsed[n.id]) delete collapsed[n.id]; else collapsed[n.id] = true;
    refresh();
  }
  nodes.forEach(function(n){
    // 折叠/展开只在「链接点」徽标上生效；点击节点内容区不折叠，避免干扰选择/复制文本
    if(n.badge){
      n.badge.addEventListener('click', function(ev){ ev.preventDefault(); ev.stopPropagation(); toggle(n); });
    }
  });
  refresh();
})();`;
}

function buildMindMapSvg(tree, options) {
  const opts = options || {};
  const fs = opts.fontSize || 16;
  const hgap = opts.hgap || 48;
  const vgap = opts.vgap || 14;
  const maxLabelWidth = opts.maxLabelWidth || 220;
  const padX = 12, padY = 8;
  const lineHeight = fs * 1.5;
  const showCount = !!opts.showCount;
  const layout = opts.layout === 'td' ? 'td' : 'lr';
  if (!tree) return '';

  if (showCount) {
    (function tag(node) {
      node.label = node.text + (node.count && node.count > 1 ? ' (' + node.count + ')' : '');
      for (const c of node.children) tag(c);
    })(tree);
  }

  measure(tree, fs, maxLabelWidth, padX, padY, lineHeight);
  if (layout === 'td') {
    computeSubtreeW(tree, hgap);
    tree.x = 0;
    tree.y = 0;
    placeTopDown(tree, hgap, vgap);
  } else {
    computeSubtree(tree, vgap);
    tree.x = -tree.w / 2;
    tree.y = -tree.h / 2;
    place(tree, 0, hgap, vgap);
  }

  const acc = { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity };
  (function bbox(node) {
    acc.minX = Math.min(acc.minX, node.x);
    acc.minY = Math.min(acc.minY, node.y);
    acc.maxX = Math.max(acc.maxX, node.x + node.w);
    acc.maxY = Math.max(acc.maxY, node.y + node.h);
    for (const c of node.children) bbox(c);
  })(tree);

  const pad = 24;
  const vbW = r(acc.maxX - acc.minX + 2 * pad);
  const vbH = r(acc.maxY - acc.minY + 2 * pad);
  const vbX = r(acc.minX - pad);
  const vbY = r(acc.minY - pad);

  const out = [];
  out.push('<svg xmlns="http://www.w3.org/2000/svg" viewBox="' + vbX + ' ' + vbY + ' ' + vbW + ' ' + vbH + '" width="' + vbW + '" height="' + vbH + '">');
  out.push('<style>text{font-family:\'Microsoft YaHei\',\'PingFang SC\',\'Noto Sans CJK SC\',sans-serif;}</style>');

  const nodes = [];
  const edges = [];
  (function walk(node, pid) {
    node.id = nodes.length;
    node.pid = pid;
    nodes.push(node);
    for (const c of node.children) { edges.push({ parent: node, child: c }); walk(c, node.id); }
  })(tree, -1);

  // 边（顶层，便于折叠时按子树隐藏/重画）
  for (const e of edges) {
    const st = styleFor(e.child);
    if (layout === 'td') {
      const x1 = e.parent.x + e.parent.w / 2, y1 = e.parent.y + e.parent.h;
      const x2 = e.child.x + e.child.w / 2, y2 = e.child.y;
      const d = Math.max(16, (y2 - y1) * 0.5);
      out.push('<path id="e' + e.child.id + '" d="M' + r(x1) + ' ' + r(y1) + ' C ' + r(x1) + ' ' + r(y1 + d) + ', ' + r(x2) + ' ' + r(y2 - d) + ', ' + r(x2) + ' ' + r(y2) + '" fill="none" stroke="' + st.edge + '" stroke-width="2" stroke-linecap="round"/>');
    } else {
      const pcy = e.parent.y + e.parent.h / 2;
      const ccy = e.child.y + e.child.h / 2;
      if (e.child.dir === -1) {
        const x1 = e.parent.x, y1 = pcy, x2 = e.child.x + e.child.w, y2 = ccy;
        const d = Math.max(16, (x1 - x2) * 0.5);
        out.push('<path id="e' + e.child.id + '" d="M' + r(x1) + ' ' + r(y1) + ' C ' + r(x1 - d) + ' ' + r(y1) + ', ' + r(x2 + d) + ' ' + r(y2) + ', ' + r(x2) + ' ' + r(y2) + '" fill="none" stroke="' + st.edge + '" stroke-width="2" stroke-linecap="round"/>');
      } else {
        const x1 = e.parent.x + e.parent.w, y1 = pcy, x2 = e.child.x, y2 = ccy;
        const d = Math.max(16, (x2 - x1) * 0.5);
        out.push('<path id="e' + e.child.id + '" d="M' + r(x1) + ' ' + r(y1) + ' C ' + r(x1 + d) + ' ' + r(y1) + ', ' + r(x2 - d) + ' ' + r(y2) + ', ' + r(x2) + ' ' + r(y2) + '" fill="none" stroke="' + st.edge + '" stroke-width="2" stroke-linecap="round"/>');
      }
    }
  }

  // 节点（用 transform 定位，折叠时只需改 translate 的 y）
  for (const node of nodes) {
    const st = styleFor(node);
    let tspan = '';
    for (let i = 0; i < node.lines.length; i++) {
      const y = padY + fs * 1.2 + i * lineHeight;
      tspan += '<tspan x="' + r(node.w / 2) + '" y="' + r(y) + '">' + escapeXml(node.lines[i]) + '</tspan>';
    }
    out.push('<g class="nd" id="n' + node.id + '" transform="translate(' + r(node.x) + ' ' + r(node.y) + ')">');
    out.push('<rect class="box" x="0" y="0" width="' + r(node.w) + '" height="' + r(node.h) + '" rx="10" fill="' + st.fill + '" stroke="' + st.stroke + '" stroke-width="1.5"/>');
    out.push('<text class="txt" text-anchor="middle" font-size="' + fs + '" fill="' + st.text + '">' + tspan + '</text>');
    out.push('</g>');
  }

  // 折叠/展开徽标（仅容器节点）
  for (const node of nodes) {
    if (!node.children.length) continue;
    const bx = layout === 'td' ? node.x + node.w / 2 : (node.dir === -1 ? node.x - 9 : node.x + node.w + 9);
    const cy = layout === 'td' ? node.y + node.h + 9 : node.y + node.h / 2;
    out.push('<g class="bdg" id="b' + node.id + '" style="cursor:pointer"><circle cx="' + r(bx) + '" cy="' + r(cy) + '" r="7" fill="#ffffff" stroke="#94a3b8" stroke-width="1"/><text x="' + r(bx) + '" y="' + r(cy) + '" text-anchor="middle" dominant-baseline="central" font-size="11" fill="#475569" font-weight="bold">−</text></g>');
  }

  // 节点折叠/展开（参照业务知识库 HTML 的做法：静态布局 + 内嵌脚本重排 y）
  const dataArray = nodes.map((n) => [n.pid, r(n.x), r(n.w), r(n.h), n.dir, n.children.length ? 1 : 0]);
  out.push('<script>\n//<![CDATA[\n');
  out.push(collapseScript(layout, hgap, vgap, pad, dataArray));
  out.push('\n//]]>\n</script>');

  out.push('</svg>');
  return out.join('');
}

// ===================== 对外 API =====================
function buildPreview(grid) {
  const maxR = Math.min(grid.length, 20);
  const maxC = Math.min(grid[0] ? grid[0].length : 0, 12);
  const cells = [];
  for (let i = 0; i < maxR; i++) {
    const row = [];
    for (let j = 0; j < maxC; j++) {
      let v = grid[i] && grid[i][j] != null ? String(grid[i][j]) : '';
      if (v.length > 50) v = v.slice(0, 50) + '…';
      row.push(v);
    }
    cells.push(row);
  }
  return { rows: maxR, cols: maxC, cells };
}

function lastValueCell(grid) {
  let lastR = 0, lastC = 0;
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      if (row[c] != null && String(row[c]).trim() !== '') {
        if (r + 1 > lastR) lastR = r + 1;
        if (c + 1 > lastC) lastC = c + 1;
      }
    }
  }
  return { lastR, lastC };
}

function buildGridResult(kind, sheets, sheetIndex, grid) {
  const rows = grid.length;
  const cols = rows ? (grid[0] ? grid[0].length : 0) : 0;
  // 「数据范围默认 A1:最后一个有值的单元格」：显式扫描最大非空行/列，
  // 不依赖 trimGrid 的包围盒右下角（后者在角上是空单元格时仍会多算）。
  const { lastR, lastC } = lastValueCell(grid);
  return {
    kind,
    sheets: sheets.map((s) => ({ index: s.index, name: s.name })),
    sheetIndex,
    grid,
    rows, cols,
    fullRows: rows, fullCols: cols,
    startCell: lastR >= 1 && lastC >= 1 ? 'A1' : '',
    endCell: lastR >= 1 && lastC >= 1 ? (numToColLetter(lastC) + lastR) : '',
    preview: buildPreview(grid)
  };
}

async function parseSpreadsheetFile(srcPath) {
  try {
    const buf = fs.readFileSync(srcPath);
    const isXlsx = buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4B && buf[2] === 0x03 && buf[3] === 0x04;
    if (isXlsx) {
      const zip = await JSZip.loadAsync(buf);
      const sheets = await listXlsxSheets(zip);
      if (!sheets.length) return { kind: 'xlsx', error: '未找到工作表' };
      const { grid, error } = await readXlsxGrid(zip, 0, null);
      if (error) return { kind: 'xlsx', sheets, error };
      return buildGridResult('xlsx', sheets, 0, grid);
    }
    const text = decodeBuffer(buf);
    const grid = trimGrid(padRows(parseCsv(text)));
    return buildGridResult('csv', [{ index: 0, name: 'Sheet1' }], 0, grid);
  } catch (err) {
    return { error: err.message };
  }
}

async function readSheetGrid(srcPath, opts) {
  const { kind, sheetIndex = 0, range = null } = opts || {};
  try {
    const buf = fs.readFileSync(srcPath);
    if (kind === 'xlsx') {
      const zip = await JSZip.loadAsync(buf);
      const sheets = await listXlsxSheets(zip);
      const { grid, error } = await readXlsxGrid(zip, sheetIndex, range);
      if (error) return { kind, sheets, error };
      return buildGridResult('xlsx', sheets, sheetIndex, grid);
    }
    const text = decodeBuffer(buf);
    let grid = trimGrid(padRows(parseCsv(text)));
    if (range) grid = sliceGrid(grid, normalizeRange(range, grid.length, grid[0] ? grid[0].length : 0));
    return buildGridResult('csv', [{ index: 0, name: 'Sheet1' }], 0, grid);
  } catch (err) {
    return { error: err.message };
  }
}

module.exports = {
  parseSpreadsheetFile,
  readSheetGrid,
  buildTree,
  buildMatrixTree,
  buildDecisionTree,
  findDecisionConflicts,
  buildMindMapSvg,
  filterGrid,
  textW,
  // 供测试/验证与 renderer 使用
  normalizeRange, parseCsv, decodeBuffer, numToColLetter
};
