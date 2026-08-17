'use strict';
/* DeepSeek Excel 助手 — MVP 核心逻辑 */

const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
function md(t) {
  let out = '';
  String(t ?? '').split('```').forEach((p, i) => {
    if (i % 2 === 1) out += '<pre>' + esc(p.replace(/^[a-z0-9_+-]+\r?\n/i, '')) + '</pre>';
    else out += mdInline(p);
  });
  return out;
}
/* 安全的 Markdown 子集：标题、列表、引用、分隔线、表格、粗斜体与行内代码。 */
function mdInlineText(value) {
  return esc(value)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>')
    .replace(/__([^_]+)__/g, '<b>$1</b>')
    .replace(/(^|[^*])\*([^*\n]+)\*/g, '$1<i>$2</i>')
    .replace(/(^|[^_])_([^_\n]+)_/g, '$1<i>$2</i>');
}
function mdInline(p) {
  const lines = String(p ?? '').replace(/\r\n?/g, '\n').split('\n');
  const blocks = [];
  let plain = [];
  let listType = '';
  let listItems = [];
  const flushPlain = () => {
    if (plain.length) blocks.push(plain.map(mdInlineText).join('<br>'));
    plain = [];
  };
  const flushList = () => {
    if (listType && listItems.length) blocks.push('<' + listType + '>' + listItems.map(x => '<li>' + mdInlineText(x) + '</li>').join('') + '</' + listType + '>');
    listType = ''; listItems = [];
  };
  const cells = line => line.replace(/^\s*\||\|\s*$/g, '').split('|').map(x => x.trim());
  const isDivider = line => /^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|?\s*$/.test(line);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.includes('|') && i + 1 < lines.length && isDivider(lines[i + 1])) {
      flushPlain(); flushList();
      const tableRows = [cells(line)];
      i += 2;
      while (i < lines.length && lines[i].includes('|') && lines[i].trim()) {
        tableRows.push(cells(lines[i])); i++;
      }
      i--;
      let table = '<table class="md-table">';
      tableRows.forEach((row, rowIndex) => {
        table += '<tr>' + row.map(cell => rowIndex === 0 ? '<th>' + mdInlineText(cell) + '</th>' : '<td>' + mdInlineText(cell) + '</td>').join('') + '</tr>';
      });
      blocks.push(table + '</table>');
      continue;
    }
    const heading = line.match(/^\s*(#{1,6})\s+(.+?)\s*#*\s*$/);
    if (heading) {
      flushPlain(); flushList();
      const level = heading[1].length;
      blocks.push('<h' + level + ' class="md-heading">' + mdInlineText(heading[2]) + '</h' + level + '>');
      continue;
    }
    const list = line.match(/^\s*(?:([-+*])|(\d+)\.)\s+(.+)$/);
    if (list) {
      flushPlain();
      const nextType = list[2] ? 'ol' : 'ul';
      if (listType && listType !== nextType) flushList();
      listType = nextType; listItems.push(list[3]);
      continue;
    }
    const quote = line.match(/^\s*>\s?(.*)$/);
    if (quote) {
      flushPlain(); flushList();
      const quoteLines = [quote[1]];
      while (i + 1 < lines.length && /^\s*>/.test(lines[i + 1])) quoteLines.push(lines[++i].replace(/^\s*>\s?/, ''));
      blocks.push('<blockquote>' + quoteLines.map(mdInlineText).join('<br>') + '</blockquote>');
      continue;
    }
    if (/^\s*(?:-{3,}|\*{3,}|_{3,})\s*$/.test(line)) {
      flushPlain(); flushList(); blocks.push('<hr>'); continue;
    }
    if (!line.trim()) {
      flushPlain(); flushList(); continue;
    }
    if (listType) flushList();
    plain.push(line);
  }
  flushPlain(); flushList();
  return blocks.join('');
}

/* 上下文截断：超长内容保留头尾，避免截断在关键信息处 */
function trimContext(t, limit) {
  if (t.length <= limit) return { text: t, truncated: false };
  const head = t.slice(0, Math.floor(limit * 0.6));
  const tail = t.slice(-Math.floor(limit * 0.35));
  return { text: head + '\n\n…（中间内容已截断，原共 ' + t.length + ' 字符）…\n\n' + tail, truncated: true };
}
/* 网格数据行列上限 */
function capGrid(values, maxRows, maxCols) {
  const arr = values || [];
  const truncated = arr.length > maxRows || arr.some(r => Array.isArray(r) && r.length > maxCols);
  return { values: arr.slice(0, maxRows).map(r => (Array.isArray(r) ? r.slice(0, maxCols) : r)), truncated };
}

/* ================= 常量 ================= */
const LSP = 'dsx-mvp-';
const PROVIDERS = {
  'https://api.deepseek.com': { name: 'DeepSeek 官方', model: 'deepseek-v4-flash' },
  'https://api.siliconflow.cn/v1': { name: 'SiliconFlow', model: 'deepseek-ai/DeepSeek-V4-Flash' },
  'https://openrouter.ai/api/v1': { name: 'OpenRouter', model: 'deepseek/deepseek-v4-flash' }
};
const MODEL_OPTIONS = {
  'https://api.deepseek.com': [['deepseek-v4-flash', 'V4-Flash（快）'], ['deepseek-v4-pro', 'V4-Pro（强）']],
  'https://api.siliconflow.cn/v1': [['deepseek-ai/DeepSeek-V4-Flash', 'V4-Flash'], ['deepseek-ai/DeepSeek-V4-Pro', 'V4-Pro']],
  'https://openrouter.ai/api/v1': [['deepseek/deepseek-v4-flash', 'V4-Flash'], ['deepseek/deepseek-v4-pro', 'V4-Pro']]
};
const DEFAULT_SKILLS = [
  { id: 'da', name: '数据分析师', instruction: '你以数据分析师身份工作：先描述数据结构，再给出关键统计量（求和/均值/最大最小/趋势），最后给出可执行的建议。' },
  { id: 'fx', name: '公式助手', instruction: '你以 Excel 公式专家身份工作：优先用函数公式解决问题，给出可直接粘贴的公式并解释；善用动态数组函数（SEQUENCE/FILTER/SORT/XLOOKUP 等）。' },
  { id: 'vba', name: 'VBA 助手', instruction: '你以 VBA 专家身份工作：给出可直接运行的 VBA 代码块，并说明用途与注意事项。' }
];
const TOOLS = [
  { type: 'function', function: { name: 'read_selection', description: '读取 Excel 当前选中单元格/区域的值和公式', parameters: { type: 'object', properties: {}, required: [] } } },
  { type: 'function', function: { name: 'read_sheet', description: '读取当前工作表已用区域的前 N 行（默认 200 行）', parameters: { type: 'object', properties: { maxRows: { type: 'number', description: '最多读取的行数，默认 200' } }, required: [] } } },
  { type: 'function', function: { name: 'read_range', description: '读取指定工作表的任意区域（常用于查看表头或局部数据；创建透视表/排序前先用它确认表头名与列序）', parameters: { type: 'object', properties: { sheet: { type: 'string', description: '工作表名称' }, address: { type: 'string', description: '区域地址，如 A1:D2' }, maxRows: { type: 'number', description: '最多读取行数，默认 50' } }, required: ['sheet', 'address'] } } },
  { type: 'function', function: { name: 'list_sheets', description: '列出工作簿中所有工作表的名称', parameters: { type: 'object', properties: {}, required: [] } } },
  { type: 'function', function: { name: 'set_values', description: '把二维数组写入指定工作表的区域（写入前会向用户弹确认框，仅当用户明确要求修改单元格时调用）', parameters: { type: 'object', properties: { sheet: { type: 'string', description: '工作表名称' }, address: { type: 'string', description: '区域地址，如 A1:B2' }, values: { type: 'array', items: { type: 'array' }, description: '二维数组，行数×列数必须与区域一致' } }, required: ['sheet', 'address', 'values'] } } },
  { type: 'function', function: { name: 'format_range', description: '设置指定区域的格式：填充色（背景色）、字体色、加粗；支持隔行相间填充（斑马纹）。写入前会向用户弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string', description: '工作表名称' }, address: { type: 'string', description: '区域地址，如 A1:D20' }, fillColor: { type: 'string', description: '填充色十六进制，如 #FFE599；不设置则保持原样' }, fontColor: { type: 'string', description: '字体色十六进制，如 #000000；不设置则保持原样' }, bold: { type: 'boolean', description: '是否加粗；不设置则保持原样' }, alternating: { type: 'boolean', description: '是否隔行相间填充（偶数行用 fillColor，奇数行留白）' } }, required: ['sheet', 'address'] } } },
  { type: 'function', function: { name: 'list_pivots', description: '列出工作簿中所有数据透视表的名称及其所在工作表', parameters: { type: 'object', properties: {}, required: [] } } },
  { type: 'function', function: { name: 'create_pivot', description: '在指定位置创建数据透视表并配置行/列/值字段。字段名必须与源区域表头完全一致，不确定时先用 read_range 查看表头（创建前会向用户弹确认框）', parameters: { type: 'object', properties: { sheet: { type: 'string', description: '透视表所在的目标工作表名称' }, destAddress: { type: 'string', description: '透视表左上角位置，如 A10' }, sourceSheet: { type: 'string', description: '数据源所在工作表名称' }, sourceAddress: { type: 'string', description: '数据源区域，如 A1:C100（首行须为列标题）' }, name: { type: 'string', description: '透视表名称，可选' }, rows: { type: 'array', items: { type: 'string' }, description: '行字段（列标题名）数组' }, columns: { type: 'array', items: { type: 'string' }, description: '列字段（列标题名）数组' }, values: { type: 'array', items: { type: 'object' }, description: '值字段数组，每项 {field: 字段名, summarizeBy: sum|count|average|max|min}' } }, required: ['sheet', 'destAddress', 'sourceSheet', 'sourceAddress'] } } },
  { type: 'function', function: { name: 'read_pivot', description: '读取指定数据透视表的数据区域内容', parameters: { type: 'object', properties: { sheet: { type: 'string', description: '工作表名称' }, name: { type: 'string', description: '透视表名称' } }, required: ['sheet', 'name'] } } },
  { type: 'function', function: { name: 'refresh_pivot', description: '刷新数据透视表；不指定 name 时尝试刷新全部透视表', parameters: { type: 'object', properties: { sheet: { type: 'string', description: '工作表名称' }, name: { type: 'string', description: '透视表名称，可选' } }, required: [] } } },
  { type: 'function', function: { name: 'refresh_connections', description: '刷新工作簿全部数据连接（含 Power Query 查询结果）；注意：无法创建/编辑查询本身', parameters: { type: 'object', properties: {}, required: [] } } },
  { type: 'function', function: { name: 'format_conditional', description: '添加条件格式：cellValue（数值/文本规则+填充色/字体色）、colorScale（三色色阶）、dataBar（数据条）、iconSet（图标集）。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' }, type: { type: 'string', description: 'cellValue / colorScale / dataBar / iconSet，默认 cellValue' }, operator: { type: 'string', description: 'greaterThan/lessThan/equalTo/notEqualTo/greaterThanOrEqual/lessThanOrEqual/between/notBetween' }, formula1: { type: 'string', description: '规则公式1，如 "10"' }, formula2: { type: 'string', description: '规则公式2（between 用）' }, fillColor: { type: 'string', description: '如 #F8CBAD' }, fontColor: { type: 'string' }, iconSet: { type: 'string', description: '图标集样式，如 ThreeTrafficLights1/ThreeArrows/ThreeFlags' } }, required: ['sheet', 'address'] } } },
  { type: 'function', function: { name: 'create_chart', description: '在工作表创建图表。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, sourceAddress: { type: 'string', description: '数据源区域，如 A1:B10' }, destAddress: { type: 'string', description: '图表左上角位置' }, chartType: { type: 'string', description: 'ColumnClustered/ColumnStacked/BarClustered/Line/LineMarkers/Pie/Area/Scatter，默认 ColumnClustered' }, title: { type: 'string', description: '图表标题' }, seriesBy: { type: 'string', description: 'Auto/Columns/Rows，默认 Auto' }, legend: { type: 'boolean', description: '是否显示图例，默认 true' } }, required: ['sheet', 'sourceAddress', 'destAddress'] } } },
  { type: 'function', function: { name: 'delete_chart', description: '按索引删除工作表上的图表（0 为第一个）。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, index: { type: 'number' } }, required: ['sheet', 'index'] } } },
  { type: 'function', function: { name: 'list_charts', description: '列出工作表上所有图表（索引与标题）', parameters: { type: 'object', properties: { sheet: { type: 'string' } }, required: ['sheet'] } } },
  { type: 'function', function: { name: 'create_table', description: '把区域转换为 Excel 表格。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' }, name: { type: 'string', description: '表格名称，可选' }, hasHeaders: { type: 'boolean', description: '首行是否为标题，默认 true' }, style: { type: 'string', description: '样式名如 TableStyleMedium2，可选' } }, required: ['sheet', 'address'] } } },
  { type: 'function', function: { name: 'list_tables', description: '列出工作表中所有表格的名称', parameters: { type: 'object', properties: { sheet: { type: 'string' } }, required: ['sheet'] } } },
  { type: 'function', function: { name: 'sort_range', description: '对区域排序（可逆操作，不弹确认）', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' }, hasHeaders: { type: 'boolean' }, keys: { type: 'array', items: { type: 'object' }, description: '[{column: 区域内列序号(0起), ascending: 布尔}]' } }, required: ['sheet', 'address', 'keys'] } } },
  { type: 'function', function: { name: 'filter_range', description: '给区域添加/移除自动筛选（可逆操作，不弹确认）', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' }, on: { type: 'boolean', description: 'true 添加筛选，false 移除' } }, required: ['sheet', 'address', 'on'] } } },
  { type: 'function', function: { name: 'freeze_panes', description: '冻结/取消冻结窗格（可逆操作，不弹确认）', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string', description: '冻结到该单元格左上角，如 B2；不填则取消冻结' } }, required: ['sheet'] } } },
  { type: 'function', function: { name: 'set_validation', description: '设置数据验证（下拉列表/整数/小数/日期）。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' }, type: { type: 'string', description: 'list/whole/decimal/date' }, formula1: { type: 'string', description: 'list 时为选项逗号分隔如 "A,B,C"；其他为下界' }, formula2: { type: 'string', description: '上界（Between 用）' }, operator: { type: 'string', description: 'Between/NotBetween/EqualTo/GreaterThan/LessThan 等，默认 Between' }, errorMessage: { type: 'string' } }, required: ['sheet', 'address', 'type', 'formula1'] } } },
  { type: 'function', function: { name: 'clear_validation', description: '清除区域的数据验证。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['sheet', 'address'] } } },
  { type: 'function', function: { name: 'add_comment', description: '给单元格添加批注。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' }, text: { type: 'string' } }, required: ['sheet', 'address', 'text'] } } },
  { type: 'function', function: { name: 'delete_comment', description: '删除单元格批注。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' } }, required: ['sheet', 'address'] } } },
  { type: 'function', function: { name: 'add_sparkline', description: '在目标单元格添加迷你图。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, sourceAddress: { type: 'string', description: '数据源（单行或单列），如 B2:G2' }, destAddress: { type: 'string', description: '目标单元格，如 H2' }, type: { type: 'string', description: 'line/column/winloss，默认 line' } }, required: ['sheet', 'sourceAddress', 'destAddress'] } } },
  { type: 'function', function: { name: 'define_name', description: '定义名称（名称管理器）。写入前弹确认框', parameters: { type: 'object', properties: { name: { type: 'string' }, refersTo: { type: 'string', description: '如 =Sheet1!$A$1:$C$10' } }, required: ['name', 'refersTo'] } } },
  { type: 'function', function: { name: 'list_names', description: '列出工作簿所有已定义名称', parameters: { type: 'object', properties: {}, required: [] } } },
  { type: 'function', function: { name: 'delete_name', description: '删除已定义名称。写入前弹确认框', parameters: { type: 'object', properties: { name: { type: 'string' } }, required: ['name'] } } },
  { type: 'function', function: { name: 'set_size', description: '设置行高/列宽，支持自动调整（不弹确认，可逆）', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string', description: '目标区域，如 A1:D10 或整列 A:C' }, columnWidth: { type: 'string', description: '数字(字符宽度)或 auto(自动调整列宽)' }, rowHeight: { type: 'string', description: '数字(磅)或 auto(自动调整行高)' } }, required: ['sheet', 'address'] } } },
  { type: 'function', function: { name: 'merge_cells', description: '合并或拆分单元格。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' }, merge: { type: 'boolean', description: 'true 合并，false 拆分' } }, required: ['sheet', 'address', 'merge'] } } },
  { type: 'function', function: { name: 'clear_range', description: '清除区域内容/格式。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string' }, what: { type: 'string', description: 'contents(内容)/formats(格式)/all(全部)，默认 contents' } }, required: ['sheet', 'address'] } } },
  { type: 'function', function: { name: 'hide_rows_cols', description: '隐藏/取消隐藏行或列（可逆操作，不弹确认）', parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string', description: '目标行/列区域，如 2:5 或 B:D' }, mode: { type: 'string', description: 'rows 或 columns' }, hidden: { type: 'boolean' } }, required: ['sheet', 'address', 'mode', 'hidden'] } } },
  { type: 'function', function: { name: 'add_sheet', description: '新建工作表。写入前弹确认框', parameters: { type: 'object', properties: { name: { type: 'string', description: '工作表名，可选' } }, required: [] } } },
  { type: 'function', function: { name: 'rename_sheet', description: '重命名工作表。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' }, newName: { type: 'string' } }, required: ['sheet', 'newName'] } } },
  { type: 'function', function: { name: 'delete_sheet', description: '删除工作表（不可恢复）。写入前弹确认框', parameters: { type: 'object', properties: { sheet: { type: 'string' } }, required: ['sheet'] } } },
  { type: 'function', function: { name: 'set_page_setup', description: '设置打印与页面：打印区域、纸张方向、缩放比例或按页宽/页高适配（可逆设置，不弹确认）', parameters: { type: 'object', properties: { sheet: { type: 'string' }, printArea: { type: 'string', description: '打印区域，如 A1:D20；不填则不改' }, orientation: { type: 'string', description: 'portrait(纵向) 或 landscape(横向)' }, scale: { type: 'number', description: '缩放百分比，如 90' }, fitToWidth: { type: 'number', description: '横向适配页数（1 表示一页宽）' }, fitToHeight: { type: 'number', description: '纵向适配页数' } }, required: ['sheet'] } } }
];

/* ================= 状态 ================= */
let SETTINGS = {
  provider: 'https://api.deepseek.com', customBase: '', customProviderName: '', apikey: '', model: 'deepseek-v4-flash',
  thinking: true, effort: 'high', ctxMode: 'selection', autoAttach: true,
  skillId: '', accent: '#4f7cff', fontSize: 13, dark: true, sidebarOpen: false, permission: 'ask'
};
const DEFAULT_QUICK = [
  { id: 'qa1', label: '📊 分析选区', prompt: '读取我的选区并做简要分析（先总结结构，再给关键统计）' },
  { id: 'qa2', label: '📈 生成图表', prompt: '根据我的选区数据生成合适的图表，放在选区右侧空白处' },
  { id: 'qa3', label: '🔢 数据透视', prompt: '根据我的选区数据创建数据透视表，请先读取表头确认字段名，位置放在 G1（若被占用请换空位）' },
  { id: 'qa4', label: '🎨 格式美化', prompt: '把我的选区美化一下：隔行浅色填充，表头加粗' }
];
let CONVS = [];
let currentId = null;
let attachedFiles = []; // {name, text}
let skills = [...DEFAULT_SKILLS];
let convFilter = ''; // 会话搜索过滤词
let quickPrompts = [...DEFAULT_QUICK];
let customProviders = [];
let editingCustomProviderId = null;

const uid = () => (crypto.randomUUID ? crypto.randomUUID() : 'c' + Date.now() + Math.random().toString(16).slice(2));
const curConv = () => CONVS.find(c => c.id === currentId);
const customProviderKey = id => 'custom:' + id;
const getCustomProvider = (value = SETTINGS.provider) => {
  if (!String(value).startsWith('custom:')) return null;
  const id = String(value).slice(7);
  return customProviders.find(p => p.id === id) || null;
};
const getBaseUrl = () => {
  const custom = getCustomProvider();
  return (custom ? (SETTINGS.customBase || custom.baseUrl) : SETTINGS.provider).trim().replace(/\/+$/, '');
};

/* 生成控制 */
let activeAbort = null;
let userAborted = false;
function setGenerating(on) { const b = $('btnStop'); if (b) b.style.display = on ? '' : 'none'; }

/* ================= 持久化 ================= */
function loadState() {
  try {
    const s = JSON.parse(localStorage.getItem(LSP + 'settings')); if (s) SETTINGS = Object.assign({}, SETTINGS, s);
    const cs = JSON.parse(localStorage.getItem(LSP + 'convs')); if (Array.isArray(cs)) CONVS = cs;
    const sk = JSON.parse(localStorage.getItem(LSP + 'skills')); if (Array.isArray(sk)) skills = sk;
    const qp = JSON.parse(localStorage.getItem(LSP + 'quick')); if (Array.isArray(qp)) quickPrompts = qp;
    const cp = JSON.parse(localStorage.getItem(LSP + 'providers')); if (Array.isArray(cp)) customProviders = cp.filter(p => p && p.id && p.name && p.baseUrl);
    currentId = localStorage.getItem(LSP + 'current') || null;
  } catch (e) { /* 忽略损坏数据 */ }
  if (SETTINGS.provider === 'custom' && SETTINGS.customBase) {
    const migrated = { id: uid(), name: SETTINGS.customProviderName || '自定义供应商', baseUrl: SETTINGS.customBase, model: SETTINGS.model };
    customProviders.push(migrated); SETTINGS.provider = customProviderKey(migrated.id);
  }
  if (String(SETTINGS.provider).startsWith('custom:') && !getCustomProvider()) SETTINGS.provider = 'https://api.deepseek.com';
  if (!CONVS.length) { CONVS = [{ id: uid(), title: '新对话', messages: [], updatedAt: Date.now() }]; currentId = CONVS[0].id; }
  if (!curConv()) currentId = CONVS[0].id;
}
function saveState() {
  try {
    localStorage.setItem(LSP + 'settings', JSON.stringify(SETTINGS));
    localStorage.setItem(LSP + 'skills', JSON.stringify(skills));
    localStorage.setItem(LSP + 'quick', JSON.stringify(quickPrompts));
    localStorage.setItem(LSP + 'providers', JSON.stringify(customProviders));
    localStorage.setItem(LSP + 'current', currentId);
  } catch (e) { /* 极小项失败可忽略 */ }
  try {
    localStorage.setItem(LSP + 'convs', JSON.stringify(CONVS));
  } catch (e) {
    setStatus('⚠ 对话记录存储空间不足：请导出 JSON 备份后删除部分历史对话');
  }
}
function renderStorageInfo() {
  const el = $('storageInfo');
  if (!el) return;
  let bytes = 0;
  try { bytes = (localStorage.getItem(LSP + 'convs') || '').length * 2; } catch (e) { bytes = 0; }
  el.textContent = '对话记录占用约 ' + (bytes / 1024).toFixed(0) + ' KB（浏览器上限约 5 MB，接近上限时请导出备份）';
}
function bumpConv() { const c = curConv(); if (c) { c.updatedAt = Date.now(); CONVS.sort((a, b) => b.updatedAt - a.updatedAt); } }

/* ================= 主题 ================= */
function applyTheme() {
  const root = document.documentElement;
  root.style.setProperty('--accent', SETTINGS.accent);
  root.style.setProperty('--fs', SETTINGS.fontSize + 'px');
  root.setAttribute('data-theme', SETTINGS.dark ? 'dark' : 'light');
  $('btnTheme').textContent = SETTINGS.dark ? '🌙' : '☀️';
}

/* ================= 侧边栏折叠（默认收起） ================= */
function applySidebar() {
  const open = !!SETTINGS.sidebarOpen;
  $('sidebar').style.display = open ? '' : 'none';
  const button = $('btnSidebar');
  const label = open ? '收起会话记录' : '显示会话记录';
  button.title = label;
  button.classList[open ? 'add' : 'remove']('is-open');
  if (button.setAttribute) {
    button.setAttribute('aria-label', label);
    button.setAttribute('aria-expanded', String(open));
  }
}

/* ================= 模态框 ================= */
let modalCb = null;
let modalExtraCb = null;
function openModal({ title, body = '', input = false, okText = '确定', cancelText = '取消', extraText = '', extraCb = null }) {
  return new Promise(res => {
    modalCb = res;
    $('modalTitle').textContent = title;
    $('modalBody').textContent = body || '';
    $('modalBody').style.display = body ? '' : 'none';
    $('modalInputRow').style.display = input ? '' : 'none';
    $('modalInput').value = '';
    $('modalOk').textContent = okText;
    $('modalCancel').textContent = cancelText;
    $('modalCancel').style.display = cancelText ? '' : 'none';
    modalExtraCb = extraCb;
    $('modalExtra').style.display = extraText ? '' : 'none';
    $('modalExtra').textContent = extraText;
    $('modal').style.display = 'flex';
  });
}
function closeModal(v) { $('modal').style.display = 'none'; const c = modalCb; modalCb = null; if (c) c(v); }
$('modalOk').onclick = () => closeModal({ ok: true, text: $('modalInput').value });
$('modalCancel').onclick = () => closeModal({ ok: false });
$('modalExtra').onclick = () => { const cb = modalExtraCb; closeModal({ ok: false }); if (cb) cb(); };

/* 敏感操作确认：权限为"替我批准"时自动放行 */
function confirmAction({ title, body, okText = '允许', cancelText = '拒绝' }) {
  if (SETTINGS.permission === 'auto') return Promise.resolve({ ok: true });
  return openModal({ title, body, okText, cancelText });
}

/* 工具清单帮助弹窗 */
function showToolCatalog() {
  const examples = [
    '「把 A 列大于 100 的单元格标红」→ format_conditional',
    '「用 A1:B10 做柱状图放在 D1」→ create_chart',
    '「按年级汇总成绩生成透视表在 F1」→ create_pivot',
    '「给 D 列加下拉选项：是/否」→ set_validation',
    '「按第一列降序排序」→ sort_range',
    '「把 A 到 D 列自动调整列宽」→ set_size',
    '「合并 A1:B2」→ merge_cells'
  ];
  const lines = TOOLS.map(t => '• ' + t.function.name + '：' + (t.function.description || ''));
  openModal({
    title: '🧰 工具清单（共 ' + TOOLS.length + ' 个）',
    body: '对话示例：\n' + examples.map(e => '  ' + e).join('\n') + '\n\n全部工具：\n' + lines.join('\n'),
    okText: '关闭', cancelText: ''
  });
}

/* ================= Office.js ================= */
let officeReady = false;
Office.onReady(() => {
  officeReady = true;
  // 顶部不再显示品牌、版本或 API 状态；只保留屏幕阅读器可识别的交互控件。
});

/* ================= Excel 上下文（默认：当前单元格/选区） ================= */
async function collectContext(mode) {
  try {
    return await Excel.run(async ctx => {
      if (mode === 'selection') {
        const r = ctx.workbook.getSelectedRange();
        r.load('address, values, formulas');
        await ctx.sync();
        const v = capGrid(r.values, 300, 100);
        const f = capGrid(r.formulas, 300, 100);
        return { type: 'selection', address: r.address, values: v.values, formulas: f.values, truncated: v.truncated || f.truncated };
      }
      if (mode === 'sheet') {
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        ws.load('name');
        await ctx.sync();
        try {
          const r = ws.getUsedRange();
          r.load('address, values, rowCount, columnCount');
          await ctx.sync();
          const cap = 200;
          return { type: 'sheet', name: ws.name, address: r.address, rowCount: r.rowCount, columnCount: r.columnCount, truncated: r.rowCount > cap, values: r.values.slice(0, cap) };
        } catch (e) { return { type: 'sheet', name: ws.name, empty: true }; }
      }
      if (mode === 'workbook') {
        const sheets = ctx.workbook.worksheets;
        sheets.load('items/name');
        await ctx.sync();
        const out = { type: 'workbook', sheets: [] };
        for (const ws of sheets.items) {
          const info = { name: ws.name };
          try {
            const r = ws.getUsedRange();
            r.load('address, values, rowCount, columnCount');
            await ctx.sync();
            info.address = r.address; info.rowCount = r.rowCount; info.columnCount = r.columnCount;
            info.values = r.values.slice(0, 30);
          } catch (e) { info.empty = true; }
          out.sheets.push(info);
        }
        return out;
      }
    });
  } catch (e) { return { error: '读取失败：' + (e.message || e) }; }
}

/* ================= 工具执行 ================= */
async function execTool(name, args) {
  if (name === 'read_selection') {
    return await Excel.run(async ctx => {
      const r = ctx.workbook.getSelectedRange();
      r.load('address, values');
      await ctx.sync();
      return { address: r.address, values: r.values };
    });
  }
  if (name === 'read_sheet') {
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load('name');
      await ctx.sync();
      const r = ws.getUsedRange();
      r.load('address, values, rowCount');
      await ctx.sync();
      const cap = Math.min((args && args.maxRows) || 200, 500);
      return { sheet: ws.name, address: r.address, rowCount: r.rowCount, truncated: r.rowCount > cap, values: r.values.slice(0, cap) };
    });
  }
  if (name === 'read_range') {
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const r = ws.getRange(args.address);
      r.load('address, values, rowCount, columnCount');
      await ctx.sync();
      const cap = Math.min((args && args.maxRows) || 50, 300);
      return { sheet: args.sheet, address: r.address, rowCount: r.rowCount, columnCount: r.columnCount, truncated: r.rowCount > cap, values: (r.values || []).slice(0, cap) };
    });
  }
  if (name === 'list_sheets') {
    return await Excel.run(async ctx => {
      const sheets = ctx.workbook.worksheets;
      sheets.load('items/name');
      await ctx.sync();
      return { sheets: sheets.items.map(s => s.name) };
    });
  }
  if (name === 'set_values') {
    const r = await openModal({
      title: '模型请求写入 Excel',
      body: '工作表：' + args.sheet + '\n区域：' + args.address + '\n值：\n' + JSON.stringify(args.values).slice(0, 800),
      okText: '允许写入', cancelText: '拒绝'
    });
    if (!r.ok) return { denied: true, message: '用户拒绝了写入' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const range = ws.getRange(args.address);
      range.values = args.values;
      await ctx.sync();
      return { written: true, sheet: args.sheet, address: args.address };
    });
  }
  if (name === 'format_range') {
    const r = await openModal({
      title: '模型请求设置格式',
      body: '工作表：' + args.sheet + '\n区域：' + args.address +
        '\n填充色：' + (args.fillColor || '不变') +
        '\n字体色：' + (args.fontColor || '不变') +
        '\n加粗：' + (args.bold ? '是' : '不变') +
        '\n隔行相间：' + (args.alternating ? '是' : '否'),
      okText: '允许设置', cancelText: '拒绝'
    });
    if (!r.ok) return { denied: true, message: '用户拒绝了格式设置' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const range = ws.getRange(args.address);
      range.load('rowCount');
      await ctx.sync();
      if (args.alternating && args.fillColor && range.rowCount > 1) {
        const n = Math.min(range.rowCount, 1000);
        for (let i = 0; i < n; i += 2) { range.getRow(i).format.fill.color = args.fillColor; }
      } else if (args.fillColor) {
        range.format.fill.color = args.fillColor;
      }
      if (args.fontColor) range.format.font.color = args.fontColor;
      if (args.bold !== undefined) range.format.font.bold = !!args.bold;
      await ctx.sync();
      return { formatted: true, sheet: args.sheet, address: args.address, alternating: !!args.alternating };
    });
  }
  if (name === 'list_pivots') {
    return await Excel.run(async ctx => {
      const sheets = ctx.workbook.worksheets;
      sheets.load('items/name');
      await ctx.sync();
      const out = [];
      for (const ws of sheets.items) {
        const pts = ws.pivotTables;
        pts.load('items/name');
        await ctx.sync();
        for (const pt of pts.items) out.push({ sheet: ws.name, name: pt.name });
      }
      return { pivots: out };
    });
  }
  if (name === 'create_pivot') {
    const r = await confirmAction({
      title: '模型请求创建数据透视表',
      body: '目标位置：' + args.sheet + '!' + args.destAddress +
        '\n数据源：' + args.sourceSheet + '!' + args.sourceAddress +
        '\n行字段：' + (args.rows || []).join('、') +
        '\n列字段：' + (args.columns || []).join('、') +
        '\n值字段：' + JSON.stringify(args.values || []),
      okText: '允许创建', cancelText: '拒绝'
    });
    if (!r.ok) return { denied: true, message: '用户拒绝了创建数据透视表' };
    const baseName = args.name || '数据透视表';
    // 第一步：创建透视表（重名时自动加后缀重试）
    const addPivot = async nm => {
      return await Excel.run(async ctx => {
        const destWs = ctx.workbook.worksheets.getItem(args.sheet);
        const srcWs = ctx.workbook.worksheets.getItem(args.sourceSheet);
        const pt = destWs.pivotTables.add(nm, srcWs.getRange(args.sourceAddress), destWs.getRange(args.destAddress));
        pt.load('name');
        await ctx.sync();
        return pt.name;
      });
    };
    let ptName;
    try { ptName = await addPivot(baseName); }
    catch (e) {
      try { ptName = await addPivot(baseName + '_' + Math.floor(Math.random() * 10000)); }
      catch (e2) {
        let existing = '';
        try {
          const info = await Excel.run(async ctx2 => {
            const w2 = ctx2.workbook.worksheets.getItem(args.sheet);
            const ps2 = w2.pivotTables;
            ps2.load('items/name');
            await ctx2.sync();
            return ps2.items.map(p => p.name);
          });
          existing = '；当前已有透视表：' + (info.join('、') || '无');
        } catch (e3) { /* 忽略 */ }
        return { error: '创建透视表失败：' + (e2.message || e2) + '。提示：源区域首行需为列标题；目标位置不要与已有数据/透视表重叠（可换个位置如 G1）' + existing };
      }
    }
    // 第二步：取透视表（枚举已加载集合，避免未初始化代理的字段集合缺失）+ 逐字段匹配配置
    try {
      return await Excel.run(async ctx => {
        const ws = ctx.workbook.worksheets.getItem(args.sheet);
        const pts = ws.pivotTables;
        pts.load('items/name');
        await ctx.sync();
        let pt = pts.items.find(p => p.name === ptName) || null;
        if (!pt) pt = ws.pivotTables.getItem(ptName);
        // 字段名来源①：层次结构集合（官方推荐用法）；失败则②：直接读源区域首行（最可靠）
        let names = [];
        try {
          if (pt.hierarchies) {
            pt.hierarchies.load('items/name');
            await ctx.sync();
            names = pt.hierarchies.items.map(h => h.name);
          }
        } catch (e) { names = []; }
        if (!names.length) {
          try {
            const srcWs = ctx.workbook.worksheets.getItem(args.sourceSheet);
            const header = srcWs.getRange(args.sourceAddress).getRow(0);
            header.load('values');
            await ctx.sync();
            names = ((header.values && header.values[0]) || []).map(v => String(v));
          } catch (e) { names = []; }
        }
        const warnings = [], appliedRows = [], appliedColumns = [], appliedValues = [];
        if (!names.length) warnings.push('无法读取字段名：请确认源区域首行是列标题');
        const find = want => {
          if (!want) return null;
          if (names.includes(want)) return want;
          const low = String(want).toLowerCase();
          return names.find(n => String(n).toLowerCase() === low) || null;
        };
        const getField = n => {
          try { return pt.hierarchies ? pt.hierarchies.getItem(n) : null; } catch (e) { return null; }
        };
        for (const f of (args.rows || [])) {
          const hit = find(f);
          if (hit) { try { pt.rowHierarchies.add(getField(hit)); await ctx.sync(); appliedRows.push(hit); } catch (e) { warnings.push('行字段「' + f + '」添加失败'); } }
          else warnings.push('未找到行字段「' + f + '」，实际字段：' + names.join('、'));
        }
        for (const f of (args.columns || [])) {
          const hit = find(f);
          if (hit) { try { pt.columnHierarchies.add(getField(hit)); await ctx.sync(); appliedColumns.push(hit); } catch (e) { warnings.push('列字段「' + f + '」添加失败'); } }
          else warnings.push('未找到列字段「' + f + '」，实际字段：' + names.join('、'));
        }
        const AGG = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min' };
        for (const v of (args.values || [])) {
          const hit = find(v && v.field);
          if (hit) { try { const h = pt.dataHierarchies.add(getField(hit)); h.summarizeBy = AGG[v.summarizeBy] || 'Sum'; await ctx.sync(); appliedValues.push(hit + ':' + (v.summarizeBy || 'sum')); } catch (e) { warnings.push('值字段「' + (v && v.field) + '」添加失败'); } }
          else warnings.push('未找到值字段「' + (v && v.field) + '」，实际字段：' + names.join('、'));
        }
        if (!appliedRows.length && !appliedColumns.length && !appliedValues.length) {
          warnings.push('未配置任何字段：透视表已创建但字段配置失败，可能需在 Excel 中手动拖拽字段');
        }
        return { created: true, sheet: args.sheet, name: ptName, appliedRows, appliedColumns, appliedValues, warnings };
      });
    } catch (e) {
      return { error: '创建透视表失败：' + (e.message || e) + '。提示：源区域首行需为列标题；目标位置不要与数据重叠' };
    }
  }
  if (name === 'read_pivot') {
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const pt = ws.pivotTables.getItem(args.name);
      const range = (typeof pt.getDataBodyRange === 'function') ? pt.getDataBodyRange() : pt.layout.getRange();
      range.load('address, values, rowCount, columnCount');
      await ctx.sync();
      const cap = 200;
      return { sheet: args.sheet, name: args.name, address: range.address, rowCount: range.rowCount, columnCount: range.columnCount, truncated: range.rowCount > cap, values: (range.values || []).slice(0, cap) };
    });
  }
  if (name === 'refresh_pivot') {
    return await Excel.run(async ctx => {
      if (args && args.name) {
        const ws = ctx.workbook.worksheets.getItem(args.sheet);
        const pt = ws.pivotTables.getItem(args.name);
        pt.refresh();
        await ctx.sync();
        return { refreshed: true, sheet: args.sheet, name: args.name };
      }
      if (typeof ctx.workbook.refreshAllPivotTables === 'function') {
        ctx.workbook.refreshAllPivotTables();
        await ctx.sync();
        return { refreshedAllPivots: true };
      }
      return { error: '当前 Excel 版本不支持一键刷新全部透视表，请指定 sheet 与 name' };
    });
  }
  if (name === 'refresh_connections') {
    return await Excel.run(async ctx => {
      if (typeof ctx.workbook.refreshAllDataConnections !== 'function') {
        return { error: '当前 Excel 版本不支持该操作' };
      }
      ctx.workbook.refreshAllDataConnections();
      await ctx.sync();
      return { refreshed: true, note: '已触发刷新全部数据连接（含 Power Query 查询结果）；查询本身的创建/编辑需在 Excel 的「数据-查询和连接」中手动进行' };
    });
  }
  if (name === 'format_conditional') {
    const r = await confirmAction({ title: '模型请求添加条件格式', body: '工作表：' + args.sheet + '\n区域：' + args.address + '\n类型：' + (args.type || 'cellValue') + '\n规则：' + JSON.stringify({ operator: args.operator, formula1: args.formula1, formula2: args.formula2 }), okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了条件格式' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const range = ws.getRange(args.address);
      const type = args.type || 'cellValue';
      if (type === 'colorScale') {
        const cs = range.conditionalFormats.add('ColorScale');
        cs.colorScale.criteria = { minimum: { color: '#F8696B', type: 'LowestValue' }, midpoint: { color: '#FFEB84', type: 'Percentile', formula: '50' }, maximum: { color: '#63BE7B', type: 'HighestValue' } };
      } else if (type === 'dataBar') {
        const db = range.conditionalFormats.add('DataBar');
        if (args.fillColor) db.dataBar.format.fill.color = args.fillColor;
      } else if (type === 'iconSet') {
        const ic = range.conditionalFormats.add('IconSet');
        ic.iconSet.iconSet = args.iconSet || 'ThreeTrafficLights1';
      } else {
        const cf = range.conditionalFormats.add('CellValue');
        if (args.fillColor) cf.cellValue.format.fill.color = args.fillColor;
        if (args.fontColor) cf.cellValue.format.font.color = args.fontColor;
        const OP = { greaterThan: 'GreaterThan', lessThan: 'LessThan', equalTo: 'EqualTo', notEqualTo: 'NotEqualTo', greaterThanOrEqual: 'GreaterThanOrEqual', lessThanOrEqual: 'LessThanOrEqual', between: 'Between', notBetween: 'NotBetween' };
        cf.cellValue.rule = { formula1: args.formula1 || '0', formula2: args.formula2, operator: OP[args.operator] || 'GreaterThan' };
      }
      await ctx.sync();
      return { formatted: true, sheet: args.sheet, address: args.address, type };
    });
  }
  if (name === 'create_chart') {
    const r = await confirmAction({ title: '模型请求创建图表', body: '工作表：' + args.sheet + '\n数据源：' + args.sourceAddress + '\n位置：' + args.destAddress + '\n类型：' + (args.chartType || 'ColumnClustered') + '\n标题：' + (args.title || ''), okText: '允许创建', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了创建图表' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const chart = ws.charts.add(args.chartType || 'ColumnClustered', ws.getRange(args.sourceAddress), args.seriesBy || 'Auto');
      if (args.title) chart.title.text = args.title;
      if (args.legend === false) chart.legend.visible = false;
      chart.setPosition(args.destAddress);
      await ctx.sync();
      return { created: true, sheet: args.sheet, type: args.chartType || 'ColumnClustered' };
    });
  }
  if (name === 'delete_chart') {
    if (args.index === undefined || args.index === null) return { error: 'delete_chart 需要 index 参数（工作表上的图表序号，从 0 开始）' };
    const r = await confirmAction({ title: '模型请求删除图表', body: '工作表：' + args.sheet + '\n图表索引：' + args.index, okText: '允许删除', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了删除图表' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      ws.charts.getItemAt(args.index).delete();
      await ctx.sync();
      return { deleted: true, sheet: args.sheet, index: args.index };
    });
  }
  if (name === 'list_charts') {
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const charts = ws.charts;
      charts.load('items/title/text');
      await ctx.sync();
      return { sheet: args.sheet, charts: charts.items.map((c, i) => ({ index: i, title: (c.title && c.title.text) || '' })) };
    });
  }
  if (name === 'create_table') {
    const r = await confirmAction({ title: '模型请求创建表格', body: '工作表：' + args.sheet + '\n区域：' + args.address + '\n名称：' + (args.name || '自动'), okText: '允许创建', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了创建表格' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const t = ws.tables.add(args.address, args.hasHeaders !== false);
      let usedName = '';
      if (args.name) {
        // 表格名只允许字母/数字/下划线/点，且须以字母或下划线开头（中文/空格会导致 Excel 报错）
        let nm = String(args.name).replace(/[^A-Za-z0-9_.]/g, '');
        if (!nm || !/^[A-Za-z_]/.test(nm)) nm = 'T_' + nm;
        t.name = nm;
        usedName = nm;
      }
      if (args.style) t.style = args.style;
      await ctx.sync();
      return { created: true, sheet: args.sheet, address: args.address, name: usedName || undefined };
    });
  }
  if (name === 'list_tables') {
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const tables = ws.tables;
      tables.load('items/name');
      await ctx.sync();
      return { sheet: args.sheet, tables: tables.items.map(t => t.name) };
    });
  }
  if (name === 'sort_range') {
    return await Excel.run(async ctx => {
      const fields = (args.keys || []).map(k => ({ key: k.column, ascending: k.ascending !== false }));
      if (!fields.length) return { error: 'sort_range 需要 keys 参数（至少一个排序列）' };
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const range = ws.getRange(args.address);
      range.sort.apply(fields, false, !!args.hasHeaders);
      await ctx.sync();
      return { sorted: true, sheet: args.sheet, address: args.address, keys: args.keys };
    });
  }
  if (name === 'filter_range') {
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const range = ws.getRange(args.address);
      if (args.on) range.autoFilter.apply(range);
      else range.autoFilter.remove();
      await ctx.sync();
      return { filtered: !!args.on, sheet: args.sheet, address: args.address };
    });
  }
  if (name === 'freeze_panes') {
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      if (args.address) ws.freezePanes.freezeAt(ws.getRange(args.address));
      else ws.freezePanes.unfreeze();
      await ctx.sync();
      return { frozen: !!args.address, sheet: args.sheet };
    });
  }
  if (name === 'set_validation') {
    const r = await confirmAction({ title: '模型请求设置数据验证', body: '工作表：' + args.sheet + '\n区域：' + args.address + '\n类型：' + args.type + '\n规则：' + args.formula1 + (args.formula2 ? ' ~ ' + args.formula2 : ''), okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了数据验证' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const dv = ws.getRange(args.address).dataValidation;
      const OP = args.operator || 'Between';
      if (args.type === 'list') dv.rule = { list: { inCellDropDown: true, source: args.formula1 } };
      else if (args.type === 'whole') dv.rule = { wholeNumber: { operator: OP, formula1: args.formula1, formula2: args.formula2 } };
      else if (args.type === 'decimal') dv.rule = { decimal: { operator: OP, formula1: args.formula1, formula2: args.formula2 } };
      else if (args.type === 'date') dv.rule = { date: { operator: OP, formula1: args.formula1, formula2: args.formula2 } };
      else throw new Error('不支持的类型：' + args.type);
      if (args.errorMessage) dv.errorAlert = { message: args.errorMessage, showAlert: true, style: 'Stop', title: '输入无效' };
      await ctx.sync();
      return { validated: true, sheet: args.sheet, address: args.address, type: args.type };
    });
  }
  if (name === 'clear_validation') {
    const r = await confirmAction({ title: '模型请求清除数据验证', body: '工作表：' + args.sheet + '\n区域：' + args.address, okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了清除数据验证' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      ws.getRange(args.address).dataValidation.clear();
      await ctx.sync();
      return { cleared: true, sheet: args.sheet, address: args.address };
    });
  }
  if (name === 'add_comment') {
    const r = await confirmAction({ title: '模型请求添加批注', body: '单元格：' + args.sheet + '!' + args.address + '\n内容：' + args.text, okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了添加批注' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      ws.comments.add(args.address, args.text);
      await ctx.sync();
      return { added: true, sheet: args.sheet, address: args.address };
    });
  }
  if (name === 'delete_comment') {
    const r = await confirmAction({ title: '模型请求删除批注', body: '单元格：' + args.sheet + '!' + args.address, okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了删除批注' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      ws.comments.getItemByCell(args.address).delete();
      await ctx.sync();
      return { deleted: true, sheet: args.sheet, address: args.address };
    });
  }
  if (name === 'add_sparkline') {
    const r = await confirmAction({ title: '模型请求添加迷你图', body: '数据源：' + args.sourceAddress + '\n目标：' + args.sheet + '!' + args.destAddress + '\n类型：' + (args.type || 'line'), okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了添加迷你图' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const T = { line: 'Line', column: 'Column', winloss: 'Stacked' };
      ws.getRange(args.destAddress).addSparkline(T[args.type] || 'Line', ws.getRange(args.sourceAddress));
      await ctx.sync();
      return { added: true, sheet: args.sheet, destAddress: args.destAddress };
    });
  }
  if (name === 'define_name') {
    const r = await confirmAction({ title: '模型请求定义名称', body: '名称：' + args.name + '\n引用：' + args.refersTo, okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了定义名称' };
    return await Excel.run(async ctx => {
      ctx.workbook.names.add(args.name, args.refersTo);
      await ctx.sync();
      return { defined: true, name: args.name, refersTo: args.refersTo };
    });
  }
  if (name === 'list_names') {
    return await Excel.run(async ctx => {
      const names = ctx.workbook.names;
      names.load('items/name,items/refersTo');
      await ctx.sync();
      return { names: names.items.map(n => ({ name: n.name, refersTo: n.refersTo })) };
    });
  }
  if (name === 'delete_name') {
    const r = await confirmAction({ title: '模型请求删除名称', body: '名称：' + args.name, okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了删除名称' };
    return await Excel.run(async ctx => {
      ctx.workbook.names.getItem(args.name).delete();
      await ctx.sync();
      return { deleted: true, name: args.name };
    });
  }
  if (name === 'set_size') {
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const range = ws.getRange(args.address);
      if (args.columnWidth === 'auto') range.format.autofitColumns();
      else if (args.columnWidth) range.format.columnWidth = Number(args.columnWidth);
      if (args.rowHeight === 'auto') range.format.autofitRows();
      else if (args.rowHeight) range.format.rowHeight = Number(args.rowHeight);
      await ctx.sync();
      return { sized: true, sheet: args.sheet, address: args.address };
    });
  }
  if (name === 'merge_cells') {
    const r = await confirmAction({ title: '模型请求' + (args.merge ? '合并' : '拆分') + '单元格', body: '工作表：' + args.sheet + '\n区域：' + args.address, okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了' + (args.merge ? '合并' : '拆分') + '单元格' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const range = ws.getRange(args.address);
      if (args.merge) range.merge(true);
      else range.unmerge();
      await ctx.sync();
      return { merged: !!args.merge, sheet: args.sheet, address: args.address };
    });
  }
  if (name === 'clear_range') {
    const r = await confirmAction({ title: '模型请求清除区域', body: '工作表：' + args.sheet + '\n区域：' + args.address + '\n清除：' + (args.what || 'contents'), okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了清除' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const what = args.what === 'formats' ? 'Formats' : (args.what === 'all' ? 'All' : 'Contents');
      ws.getRange(args.address).clear(what);
      await ctx.sync();
      return { cleared: true, sheet: args.sheet, address: args.address, what };
    });
  }
  if (name === 'hide_rows_cols') {
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const range = ws.getRange(args.address);
      if (args.mode === 'columns') range.entireColumn.hidden = !!args.hidden;
      else range.entireRow.hidden = !!args.hidden;
      await ctx.sync();
      return { hidden: !!args.hidden, mode: args.mode, sheet: args.sheet, address: args.address };
    });
  }
  if (name === 'add_sheet') {
    const r = await confirmAction({ title: '模型请求新建工作表', body: '名称：' + (args.name || '自动'), okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了新建工作表' };
    return await Excel.run(async ctx => {
      const ws = args.name ? ctx.workbook.worksheets.add(args.name) : ctx.workbook.worksheets.add();
      ws.load('name');
      await ctx.sync();
      return { created: true, name: ws.name };
    });
  }
  if (name === 'rename_sheet') {
    const r = await confirmAction({ title: '模型请求重命名工作表', body: '工作表：' + args.sheet + ' → ' + args.newName, okText: '允许', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了重命名' };
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      ws.name = args.newName;
      await ctx.sync();
      return { renamed: true, sheet: args.sheet, newName: args.newName };
    });
  }
  if (name === 'delete_sheet') {
    const r = await confirmAction({ title: '模型请求删除工作表', body: '工作表：' + args.sheet + '\n⚠ 删除后不可恢复', okText: '允许删除', cancelText: '拒绝' });
    if (!r.ok) return { denied: true, message: '用户拒绝了删除工作表' };
    return await Excel.run(async ctx => {
      ctx.workbook.worksheets.getItem(args.sheet).delete();
      await ctx.sync();
      return { deleted: true, sheet: args.sheet };
    });
  }
  if (name === 'set_page_setup') {
    return await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(args.sheet);
      const pl = ws.pageLayout;
      if (args.printArea) pl.printArea = args.printArea;
      if (args.orientation === 'landscape') pl.orientation = 'Landscape';
      else if (args.orientation === 'portrait') pl.orientation = 'Portrait';
      if (args.scale) pl.zoom = { scale: Number(args.scale) };
      else if (args.fitToWidth || args.fitToHeight) pl.zoom = { horizontalFitToPages: Number(args.fitToWidth || 0), verticalFitToPages: Number(args.fitToHeight || 0) };
      await ctx.sync();
      return { set: true, sheet: args.sheet, printArea: args.printArea || null, orientation: args.orientation || null };
    });
  }
  return { error: '未知工具：' + name };
}

/* ================= API 客户端 ================= */
async function callChat(messages, onDelta) {
  const url = getBaseUrl() + '/chat/completions';
  const body = { model: SETTINGS.model.trim(), messages, stream: true, tools: TOOLS };
  if (getBaseUrl().includes('deepseek.com')) {
    body.thinking = { type: SETTINGS.thinking ? 'enabled' : 'disabled' };
    if (SETTINGS.thinking && SETTINGS.effort) body.reasoning_effort = SETTINGS.effort;
  }

  const ctrl = new AbortController();
  activeAbort = ctrl;
  userAborted = false;
  const timer = setTimeout(() => ctrl.abort(), 240000);
  let res;
  try {
    try {
      res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + SETTINGS.apikey },
        body: JSON.stringify(body),
        signal: ctrl.signal
      });
    } catch (e) {
      if (e.name === 'AbortError') throw new Error(userAborted ? '已停止生成' : '请求超时（240 秒）');
      throw new Error('网络请求失败：' + (e.message || e) + '（若为 CORS/failed to fetch，检查 Base URL 与网络）');
    }
    if (!res.ok) {
      const t = await res.text().catch(() => '');
      let msg = 'HTTP ' + res.status + '：' + t.slice(0, 300);
      if (res.status === 401) msg = '认证失败(401)：API Key 与供应商不匹配（官方 key 只能用于官方地址）';
      if (res.status === 402) msg = '余额不足(402)：请为 API 账户充值';
      if (res.status === 400 && t.includes('reasoning_content')) msg = 'HTTP 400：工具调用轮次缺少 reasoning_content 回传';
      throw new Error(msg);
    }

  const reader = res.body.getReader();
  const dec = new TextDecoder();
  let buf = '', content = '', reasoning = '', finishReason = null;
  const toolCalls = {};
  const processLine = line => {
    line = line.trim();
    if (!line.startsWith('data:')) return;
    const payload = line.slice(5).trim();
    if (payload === '[DONE]') return;
    let j; try { j = JSON.parse(payload); } catch (e) { return; }
    const d = j.choices && j.choices[0] && j.choices[0].delta;
    if (!d) return;
    if (d.reasoning_content) { reasoning += d.reasoning_content; onDelta && onDelta({ type: 'think', text: d.reasoning_content }); }
    if (d.content) { content += d.content; onDelta && onDelta({ type: 'text', text: d.content }); }
    if (d.tool_calls) for (const tc of d.tool_calls) {
      const k = tc.index || 0;
      if (!toolCalls[k]) toolCalls[k] = { id: '', name: '', args: '' };
      if (tc.id) toolCalls[k].id = tc.id;
      if (tc.function && tc.function.name) toolCalls[k].name += tc.function.name;
      if (tc.function && tc.function.arguments) toolCalls[k].args += tc.function.arguments;
    }
    if (j.choices[0].finish_reason) finishReason = j.choices[0].finish_reason;
  };
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += dec.decode(value, { stream: true });
    let i;
    while ((i = buf.indexOf('\n')) >= 0) { const line = buf.slice(0, i); buf = buf.slice(i + 1); processLine(line); }
  }
  if (buf.trim()) processLine(buf);
  return { content, reasoning, finishReason, toolCalls: Object.values(toolCalls) };
  } finally {
    clearTimeout(timer);
    if (activeAbort === ctrl) activeAbort = null;
  }
}

/* ================= 系统提示（含技能） ================= */
function systemPrompt() {
  let p = '你是运行在 Excel 加载项中的 AI 助手，可以调用工具读取/写入用户的 Excel 数据以及设置格式。\n' +
    '规则：\n1. 需要数据时调用 read_selection / read_sheet / list_sheets，不要凭空猜测。\n' +
    '2. set_values（写值/公式）与 format_range（填充色/字体色/加粗/隔行相间着色等格式）仅在用户明确要求修改时调用（会向用户弹确认框）。\n' +
    '3. 用简体中文回答，简洁结构化，涉及数据时给出关键数字。\n' +
    '4. 数据量大时先总结结构再分析。\n' +
    '5. 记住整个对话上下文，用户的追问（如"那 A2 呢""那平均呢"）都指代之前讨论的内容。\n' +
    '6. 数据透视表用 list_pivots / create_pivot / read_pivot / refresh_pivot；刷新 Power Query 等外部数据连接用 refresh_connections（Office.js 无法创建/编辑 Power Query 查询本身，只能触发刷新；若用户要求新建/修改查询，请说明该限制并建议手动操作）。创建透视表/排序前，先用 read_range 或 read_sheet 查看源区域表头，字段名必须与表头完全一致。\n' +
    '7. 其他能力：条件格式 format_conditional；图表 create_chart / delete_chart / list_charts；表格 create_table / list_tables；排序 sort_range；筛选 filter_range；冻结窗格 freeze_panes；数据验证 set_validation / clear_validation；批注 add_comment / delete_comment；迷你图 add_sparkline；名称管理器 define_name / list_names / delete_name；行高列宽 set_size；合并/拆分单元格 merge_cells；清除内容/格式 clear_range；隐藏行/列 hide_rows_cols；工作表管理 add_sheet / rename_sheet / delete_sheet；打印设置 set_page_setup。创建/删除/修改类操作会向用户弹确认框，排序/筛选/冻结/行高列宽/隐藏/打印设置为可逆操作不弹框。';
  const sk = skills.find(s => s.id === SETTINGS.skillId);
  if (sk) p += '\n\n[当前技能：' + sk.name + ']\n' + sk.instruction;
  return p;
}

/* ================= Agent 循环 ================= */
async function runAgent() {
  const c = curConv();
  if (!c) return;
  const maxRounds = 6;
  setGenerating(true);
  try {
  for (let step = 0; step < maxRounds; step++) {
    setStatus('第 ' + (step + 1) + ' 轮 API 调用…');
    let acc = '', think = '', box = null, bubble = null, thinkEl = null;
    const r = await callChat([{ role: 'system', content: systemPrompt() }, ...c.messages], d => {
      if (d.type === 'think') {
        think += d.text;
        box = box || appendRespBox();
        if (!thinkEl) { thinkEl = makeThinkEl(); box.appendChild(thinkEl); }
        thinkEl.querySelector('.thinkBody').textContent = think;
      } else {
        acc += d.text;
        box = box || appendRespBox();
        bubble = bubble || appendBubble(box);
        bubble.innerHTML = md(acc);
      }
      if (nearBottom()) $('chat').scrollTop = $('chat').scrollHeight;
    });
    const assistant = { role: 'assistant', content: r.content, reasoning_content: r.reasoning };
    if (r.toolCalls.length) assistant.tool_calls = r.toolCalls.map(t => ({ id: t.id, type: 'function', function: { name: t.name, arguments: t.args } }));
    c.messages.push(assistant);
    saveState();
    if (!r.toolCalls.length) {
      if (bubble && r.content) bubble.appendChild(copyBtn(r.content));
      setStatus('完成（第 ' + (step + 1) + ' 轮结束）');
      return;
    }
    for (const t of r.toolCalls) {
      const detBody = logToolStart(t.name, t.args);
      let args = {};
      try { args = JSON.parse(t.args || '{}'); } catch (e) { args = {}; }
      let result;
      try { result = await execTool(t.name, args); } catch (e) { result = { error: String(e.message || e) }; }
      const txt = JSON.stringify(result);
      logToolEnd(detBody, txt);
      c.messages.push({ role: 'tool', tool_call_id: t.id, content: txt.slice(0, 6000) });
    }
    if (c.messages.length > 60) c.messages.splice(0, c.messages.length - 60);
    saveState();
  }
  setStatus('达到最大轮数（' + maxRounds + ' 轮）');
  } finally {
    setGenerating(false);
  }
}

function copyBtn(text) {
  const wrap = document.createElement('div');
  wrap.className = 'msg-actions';
  const b = document.createElement('button');
  b.className = 'copy-btn';
  b.textContent = '📋 复制';
  b.title = '复制到剪贴板';
  b.onclick = async () => {
    try { await navigator.clipboard.writeText(text); b.textContent = '✓ 已复制'; setTimeout(() => { b.textContent = '📋 复制'; }, 1500); }
    catch (e) { b.textContent = '✗ 复制失败'; }
  };
  wrap.appendChild(b);
  return wrap;
}

/* ================= 渲染 ================= */
function setStatus(t) { $('status').textContent = t; }
function setSkillResult(t) { const el = $('skillResult'); if (el) el.textContent = t; }
function nearBottom() {
  const c = $('chat');
  return (c.scrollHeight - c.scrollTop - (c.clientHeight || 0)) < 80;
}
function log(cls, t) {
  const d = document.createElement('div');
  d.className = cls || '';
  d.textContent = t;
  $('log').appendChild(d);
  $('log').scrollTop = $('log').scrollHeight;
}
/* 工具调用日志：默认折叠，摘要行只显示工具名 */
function logToolStart(name, args) {
  const det = document.createElement('details');
  det.className = 'log-tool';
  const sum = document.createElement('summary');
  sum.textContent = '🔧 ' + name + ' · ' + String(args).slice(0, 60);
  det.appendChild(sum);
  const body = document.createElement('pre');
  body.className = 'log-body';
  body.textContent = '→ 工具调用：' + name + '(' + String(args).slice(0, 200) + ')';
  det.appendChild(body);
  $('log').appendChild(det);
  $('log').scrollTop = $('log').scrollHeight;
  return body;
}
function logToolEnd(bodyEl, txt) {
  bodyEl.textContent += '\n← 结果：' + txt.slice(0, 250) + (txt.length > 250 ? '…' : '');
  $('log').scrollTop = $('log').scrollHeight;
}
function appendRespBox() { const b = document.createElement('div'); b.className = 'respBox'; $('chat').appendChild(b); return b; }
function appendBubble(box) { const d = document.createElement('div'); d.className = 'bubble assistant'; (box || $('chat')).appendChild(d); return d; }
function makeThinkEl() {
  const det = document.createElement('details');
  det.className = 'think';
  det.innerHTML = '<summary>💭 思考过程</summary><pre class="thinkBody"></pre>';
  return det;
}

function renderChat() {
  $('chat').innerHTML = '';
  const c = curConv();
  if (!c) return;
  for (const m of c.messages) {
    if (m.role === 'user') {
      const d = document.createElement('div');
      d.className = 'bubble user';
      d.textContent = m.display || m.content;
      d.appendChild(copyBtn(m.display || m.content));
      $('chat').appendChild(d);
    } else if (m.role === 'assistant') {
      if (m.reasoning_content) {
        const det = makeThinkEl();
        det.querySelector('.thinkBody').textContent = m.reasoning_content;
        $('chat').appendChild(det);
      }
      if (m.content) {
        const d = document.createElement('div');
        d.className = 'bubble assistant';
        d.innerHTML = md(m.content);
        d.appendChild(copyBtn(m.content));
        $('chat').appendChild(d);
      } else if (m.tool_calls && m.tool_calls.length) {
        const d = document.createElement('div');
        d.className = 'hint';
        d.textContent = '🔧 调用工具：' + m.tool_calls.map(t => t.function.name).join(', ');
        $('chat').appendChild(d);
      }
    }
  }
  refreshEmptyState();
  $('chat').scrollTop = $('chat').scrollHeight;
}

async function renameConv(id) {
  const c = CONVS.find(x => x.id === id);
  if (!c) return;
  const r = await openModal({ title: '重命名对话', input: true, okText: '保存' });
  if (r.ok && r.text.trim()) { c.title = r.text.trim(); saveState(); renderConvList(); }
}

function renderConvList() {
  $('convList').innerHTML = '';
  const q = convFilter.trim().toLowerCase();
  const list = q ? CONVS.filter(c => {
    if (c.title && String(c.title).toLowerCase().includes(q)) return true;
    return (c.messages || []).some(m => {
      const t = m.role === 'user' ? (m.display || m.content) : (m.content || '');
      return String(t).toLowerCase().includes(q);
    });
  }) : CONVS;
  if (q && !list.length) {
    const empty = document.createElement('div');
    empty.className = 'hint';
    empty.textContent = '无匹配的对话';
    $('convList').appendChild(empty);
    return;
  }
  for (const c of list) {
    const item = document.createElement('div');
    item.className = 'conv-item' + (c.id === currentId ? ' active' : '');
    const ttl = document.createElement('span');
    ttl.className = 'ttl';
    ttl.textContent = c.title || '（空对话）';
    ttl.title = '点击切换，双击重命名';
    const del = document.createElement('span');
    del.className = 'del';
    del.textContent = '✕';
    del.title = '删除';
    del.onclick = e => { e.stopPropagation(); deleteConv(c.id); };
    const ren = document.createElement('span');
    ren.className = 'del';
    ren.textContent = '✎';
    ren.title = '重命名';
    ren.onclick = e => { e.stopPropagation(); renameConv(c.id); };
    item.appendChild(ttl);
    item.appendChild(ren);
    item.appendChild(del);
    item.onclick = () => switchConv(c.id);
    item.ondblclick = () => renameConv(c.id);
    $('convList').appendChild(item);
  }
}

function deleteConv(id) {
  if (CONVS.length <= 1) { setStatus('至少保留一个对话'); return; }
  CONVS = CONVS.filter(c => c.id !== id);
  if (currentId === id) currentId = CONVS[0].id;
  saveState(); renderConvList(); renderChat();
}
function switchConv(id) {
  currentId = id; saveState(); renderConvList(); renderChat();
  $('log').innerHTML = ''; setStatus('已切换对话');
}

/* ================= 导出到工作表 ================= */
async function writeConvToSheet() {
  const c = curConv(); if (!c) return;
  const rows = [['角色', '内容']];
  for (const m of c.messages) {
    if (m.role === 'user') rows.push(['用户', m.display || m.content]);
    else if (m.role === 'assistant' && m.content) rows.push(['助手', m.content]);
    else if (m.role === 'tool') rows.push(['工具', String(m.content).slice(0, 500)]);
  }
  if (rows.length <= 1) { setStatus('当前对话为空，无需导出'); return; }
  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.add('对话导出');
      sheet.getRangeByIndexes(0, 0, rows.length, 2).values = rows;
      sheet.getRange('A:A').format.columnWidth = 120;
      sheet.getRange('B:B').format.columnWidth = 100;
      sheet.getRange('B:B').format.wrapText = true;
      await ctx.sync();
    });
    setStatus('已写入新工作表「对话导出」');
  } catch (e) { setStatus('写入工作表失败：' + (e.message || e)); }
}

/* ================= 技能 ================= */
function renderSkillSelect() {
  $('skillSel').innerHTML = '<option value="">（不使用技能）</option>';
  for (const s of skills) {
    const o = document.createElement('option');
    o.value = s.id; o.textContent = s.name;
    $('skillSel').appendChild(o);
  }
  $('skillSel').value = SETTINGS.skillId || '';
}

function renderSkillList() {
  const box = $('skillList');
  if (!box) return;
  box.innerHTML = '';
  for (const s of skills) {
    const item = document.createElement('div');
    item.className = 'skill-item';
    const nm = document.createElement('span');
    nm.textContent = s.name;
    const del = document.createElement('button');
    del.className = 'ghost small';
    del.textContent = '✕';
    del.title = '删除技能';
    del.onclick = () => deleteSkill(s.id);
    item.appendChild(nm);
    item.appendChild(del);
    box.appendChild(item);
  }
}

function deleteSkill(id) {
  const sk = skills.find(s => s.id === id);
  skills = skills.filter(s => s.id !== id);
  if (SETTINGS.skillId === id) SETTINGS.skillId = '';
  saveState(); renderSkillSelect(); renderSkillList();
  setSkillResult(sk ? '已删除技能「' + sk.name + '」' : '已删除');
}

async function exportSkills() {
  const json = JSON.stringify(skills, null, 2);
  const r = await openModal({ title: '导出技能（JSON）', body: json, okText: '复制', cancelText: '关闭' });
  if (!r.ok) return;
  try { await navigator.clipboard.writeText(json); setSkillResult('✓ 技能 JSON 已复制到剪贴板'); }
  catch (e) { setSkillResult('复制失败，请在上方文本框全选复制'); }
}

async function importSkills() {
  const r = await openModal({ title: '导入技能（粘贴 JSON）', input: true, okText: '导入' });
  if (!r.ok) return;
  try {
    const arr = JSON.parse(r.text);
    if (!Array.isArray(arr)) throw new Error('不是数组');
    let added = 0, updated = 0;
    for (const s of arr) {
      if (!s || !s.name || !s.instruction) continue;
      const exist = skills.find(x => x.name === s.name);
      if (exist) { exist.instruction = s.instruction; updated++; }
      else { skills.push({ id: uid(), name: s.name, instruction: s.instruction }); added++; }
    }
    saveState(); renderSkillSelect(); renderSkillList();
    setSkillResult('✓ 导入完成：新增 ' + added + ' 个，更新 ' + updated + ' 个');
  } catch (e) { setSkillResult('⚠ 导入失败：' + (e.message || e)); }
}

/* ================= 文件 ================= */
async function handleFiles(fileList) {
  for (const f of fileList) {
    try {
      const name = f.name.toLowerCase();
      let text = '';
      if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
        if (typeof XLSX === 'undefined') { attachedFiles.push({ name: f.name, text: '[xlsx 解析失败：SheetJS 未加载]' }); continue; }
        const buf = await f.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        text = XLSX.utils.sheet_to_csv(ws);
      } else if (name.endsWith('.png') || name.endsWith('.jpg') || name.endsWith('.jpeg') || name.endsWith('.gif') || name.endsWith('.webp')) {
        attachedFiles.push({ name: f.name, text: '[图片文件暂不支持直接入上下文]' });
        continue;
      } else {
        text = await f.text();
      }
      if (text.length > 20000) text = text.slice(0, 20000) + '\n…（已截断）';
      attachedFiles.push({ name: f.name, text });
    } catch (e) {
      attachedFiles.push({ name: f.name, text: '[读取失败：' + (e.message || e) + ']' });
    }
  }
  if (attachedFiles.length > 5) attachedFiles = attachedFiles.slice(0, 5);
  renderFileList();
}
function renderFileList() {
  $('fileList').textContent = attachedFiles.length
    ? '已附加 ' + attachedFiles.length + ' 个：' + attachedFiles.map(f => f.name).join('、')
    : '';
  $('clearFiles').style.display = attachedFiles.length ? '' : 'none';
}

/* ================= 设置同步与 UI 填充 ================= */
function syncSettings() {
  SETTINGS.apikey = $('apikey').value.trim();
  SETTINGS.customBase = $('customBase').value.trim();
  SETTINGS.customProviderName = $('customProviderName').value.trim();
  SETTINGS.model = $('model').value.trim();
  updateModelHint();
}

function buildModelOptions() {
  const sel = $('model');
  sel.innerHTML = '';
  const opts = [...(MODEL_OPTIONS[SETTINGS.provider] || MODEL_OPTIONS['https://api.deepseek.com'])];
  if (getCustomProvider() && SETTINGS.model && !opts.some(p => p[0] === SETTINGS.model)) opts.push([SETTINGS.model, SETTINGS.model + '（自定义）']);
  for (const [v, label] of opts) {
    const o = document.createElement('option');
    o.value = v; o.textContent = label;
    sel.appendChild(o);
  }
  if (!opts.some(p => p[0] === SETTINGS.model)) SETTINGS.model = opts[0][0];
  sel.value = SETTINGS.model;
}

function populateSettings() {
  const sel = $('provider');
  sel.innerHTML = '';
  for (const key of Object.keys(PROVIDERS)) {
    const o = document.createElement('option');
    o.value = key; o.textContent = PROVIDERS[key].name;
    sel.appendChild(o);
  }
  for (const provider of customProviders) {
    const option = document.createElement('option');
    option.value = customProviderKey(provider.id); option.textContent = provider.name;
    sel.appendChild(option);
  }
  const addOption = document.createElement('option');
  addOption.value = '__add_custom__'; addOption.textContent = '＋ 添加自定义供应商…';
  sel.appendChild(addOption);
  if (!(SETTINGS.provider in PROVIDERS) && !getCustomProvider()) SETTINGS.provider = 'https://api.deepseek.com';
  sel.value = SETTINGS.provider;
  renderCustomProviderEditor(getCustomProvider());
  $('apikey').value = SETTINGS.apikey;
  buildModelOptions();
  $('thinking').checked = SETTINGS.thinking;
  $('effort').value = SETTINGS.effort;
  $('ctxMode').value = SETTINGS.ctxMode;
  $('autoAttach').checked = SETTINGS.autoAttach;
  $('permission').value = SETTINGS.permission;
  $('accent').value = SETTINGS.accent;
  $('fontSize').value = String(SETTINGS.fontSize);
  renderSkillSelect();
  renderSkillList();
  refreshEmptyState();
  renderQuickList();
  renderFileList();
  renderStorageInfo();
  updateModelHint();
}

function renderCustomProviderEditor(provider, creating = false) {
  const row = $('customProviderFields');
  if (!row) return;
  const show = creating || !!provider;
  row.style.display = show ? 'flex' : 'none';
  editingCustomProviderId = provider ? provider.id : null;
  $('customProviderName').value = provider ? provider.name : '';
  $('customBase').value = provider ? provider.baseUrl : '';
  $('saveCustomProvider').textContent = provider ? '更新供应商' : '添加供应商';
  if (provider) {
    SETTINGS.customProviderName = provider.name;
    SETTINGS.customBase = provider.baseUrl;
  }
}

function saveCustomProviderFromForm() {
  const name = $('customProviderName').value.trim();
  const baseUrl = $('customBase').value.trim().replace(/\/+$/, '');
  const result = $('providerResult');
  if (!name || !baseUrl) { result.textContent = '⚠ 请填写供应商名称和 API 地址'; return; }
  try {
    const parsed = new URL(baseUrl);
    if (!/^https?:$/.test(parsed.protocol)) throw new Error('protocol');
  } catch (e) { result.textContent = '⚠ API 地址必须是完整的 http:// 或 https:// 地址'; return; }
  const duplicate = customProviders.find(p => p.id !== editingCustomProviderId && p.name.toLowerCase() === name.toLowerCase());
  if (duplicate) { result.textContent = '⚠ 已存在同名供应商'; return; }
  let provider = customProviders.find(p => p.id === editingCustomProviderId);
  if (provider) {
    provider.name = name; provider.baseUrl = baseUrl; provider.model = SETTINGS.model;
  } else {
    provider = { id: uid(), name, baseUrl, model: SETTINGS.model };
    customProviders.push(provider);
  }
  SETTINGS.provider = customProviderKey(provider.id);
  SETTINGS.customProviderName = name; SETTINGS.customBase = baseUrl;
  saveState(); populateSettings();
  result.textContent = '✓ 已保存供应商「' + name + '」，可继续从下拉框添加其他供应商';
}

function updateModelHint() {
  const m = SETTINGS.model.toLowerCase();
  if (m.includes('reasoner') || m.includes('r1') || m === 'deepseek-chat') {
    $('modelHint').textContent = '注意：deepseek-chat / deepseek-reasoner 旧模型名已进入停用流程，请用 V4 模型';
  } else if (!getBaseUrl().includes('deepseek.com')) {
    $('modelHint').textContent = '网关提示：思考模式与推理强度参数仅官方 API 生效，网关按各自默认行为执行';
  } else {
    $('modelHint').textContent = 'V4 说明：思考模式默认开启；推理强度 low/high/max（medium、xhigh 映射为 high）；思考模式下 temperature 等不生效';
  }
}

/* 快捷指令 */
function applyQuickPrompt(text) {
  $('input').value = text;
  setStatus('已填入快捷指令，可直接发送或修改');
}
/* 空态快捷操作（新对话空白区显示，选择后消失） */
function updateEmptyState(show) {
  const el = $('emptyState');
  if (!el) return;
  if (show) {
    el.innerHTML = '';
    const title = document.createElement('div');
    title.className = 'empty-title';
    title.textContent = '👋 想对 Excel 做什么？';
    el.appendChild(title);
    const sub = document.createElement('div');
    sub.className = 'empty-sub';
    sub.textContent = '选择一个快捷操作开始，或直接在下方输入';
    el.appendChild(sub);
    const grid = document.createElement('div');
    grid.className = 'empty-grid';
    for (const q of quickPrompts) {
      const b = document.createElement('button');
      b.textContent = q.label;
      b.title = q.prompt;
      b.onclick = () => { applyQuickPrompt(q.prompt); updateEmptyState(false); };
      grid.appendChild(b);
    }
    el.appendChild(grid);
    el.style.display = '';
  } else {
    el.style.display = 'none';
  }
}
function refreshEmptyState() {
  const c = curConv();
  updateEmptyState(!(c && c.messages && c.messages.length));
}
function renderQuickList() {
  const box = $('quickList');
  if (!box) return;
  box.innerHTML = '';
  for (const q of quickPrompts) {
    const item = document.createElement('div');
    item.className = 'skill-item';
    const nm = document.createElement('span');
    nm.textContent = q.label;
    nm.title = q.prompt;
    const del = document.createElement('button');
    del.className = 'ghost small';
    del.textContent = '✕';
    del.title = '删除快捷指令';
    del.onclick = () => deleteQuick(q.id);
    item.appendChild(nm);
    item.appendChild(del);
    box.appendChild(item);
  }
}
function deleteQuick(id) {
  quickPrompts = quickPrompts.filter(q => q.id !== id);
  saveState(); refreshEmptyState(); renderQuickList();
  const r = $('quickResult'); if (r) r.textContent = '已删除快捷指令';
}

/* 输入框回车发送（中文输入法组合中回车不发送；Ctrl/Cmd+Enter 强制发送） */
function handleInputKey(e) {
  const isEnter = e.key === 'Enter';
  const plainEnter = isEnter && !e.shiftKey && !e.isComposing && e.keyCode !== 229;
  const ctrlEnter = isEnter && (e.ctrlKey || e.metaKey);
  if (plainEnter || ctrlEnter) {
    e.preventDefault();
    $('send').click();
  }
}

/* ================= 事件绑定 ================= */
function bindEvents() {
  $('btnTheme').onclick = () => { SETTINGS.dark = !SETTINGS.dark; saveState(); applyTheme(); };
  $('btnSidebar').onclick = () => { SETTINGS.sidebarOpen = !SETTINGS.sidebarOpen; saveState(); applySidebar(); };
  $('convSearch').oninput = e => { convFilter = e.target.value; renderConvList(); };
  $('btnStop').onclick = () => {
    userAborted = true;
    if (activeAbort) activeAbort.abort();
    setStatus('⏹ 已停止生成');
  };
  $('btnSettings').onclick = () => { renderStorageInfo(); $('settingsModal').style.display = 'flex'; };
  $('btnCloseSettings').onclick = () => { $('settingsModal').style.display = 'none'; };

  $('saveKey').onclick = () => { SETTINGS.apikey = $('apikey').value.trim(); saveState(); setStatus('API Key 已保存'); };
  $('clearKey').onclick = () => { SETTINGS.apikey = ''; $('apikey').value = ''; saveState(); setStatus('API Key 已清除'); };
  $('customBase').onchange = e => { SETTINGS.customBase = e.target.value.trim(); saveState(); };
  $('customProviderName').onchange = e => { SETTINGS.customProviderName = e.target.value.trim(); };
  $('saveCustomProvider').onclick = saveCustomProviderFromForm;

  $('provider').onchange = e => {
    if (e.target.value === '__add_custom__') {
      renderCustomProviderEditor(null, true);
      $('providerResult').textContent = '填写名称和 API 地址后保存；保存后可继续添加。';
      return;
    }
    SETTINGS.provider = e.target.value;
    const custom = getCustomProvider(e.target.value);
    renderCustomProviderEditor(custom);
    if (e.target.value in PROVIDERS) SETTINGS.model = PROVIDERS[e.target.value].model;
    else if (custom && custom.model) SETTINGS.model = custom.model;
    buildModelOptions();
    saveState(); updateModelHint();
  };
  $('model').onchange = e => { SETTINGS.model = e.target.value; const custom = getCustomProvider(); if (custom) custom.model = e.target.value; saveState(); updateModelHint(); };
  $('effort').onchange = e => { SETTINGS.effort = e.target.value; saveState(); };
  $('thinking').onchange = e => { SETTINGS.thinking = e.target.checked; saveState(); };
  $('ctxMode').onchange = e => { SETTINGS.ctxMode = e.target.value; saveState(); };
  $('autoAttach').onchange = e => { SETTINGS.autoAttach = e.target.checked; saveState(); };
  $('permission').onchange = e => { SETTINGS.permission = e.target.value; saveState(); setStatus('敏感操作权限：' + (e.target.value === 'auto' ? '替我批准（自动放行）' : '请求批准（每次确认）')); };

  $('skillSel').onchange = e => { SETTINGS.skillId = e.target.value; saveState(); };
  $('addSkill').onclick = () => {
    const name = $('skillName').value.trim(), instr = $('skillInstr').value.trim();
    if (!name || !instr) { setSkillResult('⚠ 请填写技能名称与指令'); return; }
    skills.push({ id: uid(), name, instruction: instr });
    $('skillName').value = ''; $('skillInstr').value = '';
    saveState(); renderSkillSelect(); renderSkillList();
    setSkillResult('✓ 已添加技能「' + name + '」，可在聊天工具栏的「技能」下拉中选择使用');
  };
  $('addQuick').onclick = () => {
    const label = $('quickLabel').value.trim(), prompt = $('quickPrompt').value.trim();
    const out = t => { const r = $('quickResult'); if (r) r.textContent = t; };
    if (!label || !prompt) { out('⚠ 请填写按钮名与指令内容'); return; }
    quickPrompts.push({ id: uid(), label, prompt });
    $('quickLabel').value = ''; $('quickPrompt').value = '';
    saveState(); refreshEmptyState(); renderQuickList();
    out('✓ 已添加快捷指令「' + label + '」，已出现在空白对话的快捷操作区');
  };
  $('btnExportSkills').onclick = () => exportSkills();
  $('btnImportSkills').onclick = () => importSkills();
  $('btnTools').onclick = () => showToolCatalog();

  $('btnAttach').onclick = () => $('fileInput').click();
  $('fileInput').onchange = e => { handleFiles(e.target.files); e.target.value = ''; };
  $('clearFiles').onclick = () => { attachedFiles = []; renderFileList(); };

  $('accent').onchange = e => { SETTINGS.accent = e.target.value; saveState(); applyTheme(); };
  $('fontSize').onchange = e => { SETTINGS.fontSize = Number(e.target.value); saveState(); applyTheme(); };

  $('testConn').onclick = async () => {
    syncSettings();
    const out = t => { if ($('connResult')) $('connResult').textContent = t; setStatus(t); };
    setGenerating(true);
    out('测试连接中…');
    try {
      const t0 = Date.now();
      const r = await callChat([{ role: 'system', content: systemPrompt() }, { role: 'user', content: '请只回复：ok' }], null);
      out('✓ 连接成功（' + (Date.now() - t0) + ' ms）：' + (r.content || '(空)').slice(0, 40) + (r.reasoning ? '；思考链 ' + r.reasoning.length + ' 字符' : ''));
    } catch (e) { out('✗ 失败：' + (e.message || e)); } finally { setGenerating(false); }
  };

  $('btnNew').onclick = () => {
    convFilter = '';
    if ($('convSearch')) $('convSearch').value = '';
    const c = { id: uid(), title: '新对话', messages: [], updatedAt: Date.now() };
    CONVS.unshift(c); currentId = c.id; saveState(); renderConvList(); renderChat();
    $('log').innerHTML = ''; setStatus('已新建对话');
  };

  $('btnExport').onclick = async () => {
    const json = JSON.stringify(CONVS, null, 2);
    const r = await openModal({ title: '导出对话（JSON）', body: json, okText: '复制', cancelText: '关闭' });
    if (r.ok) {
      try { await navigator.clipboard.writeText(json); setStatus('已复制到剪贴板'); }
      catch (e) { setStatus('复制失败，请在上方文本框全选复制'); }
    }
  };
  $('btnExportSheet').onclick = () => writeConvToSheet();
  $('btnImport').onclick = async () => {
    const r = await openModal({ title: '导入对话（粘贴 JSON）', input: true, okText: '导入' });
    if (!r.ok) return;
    try {
      const arr = JSON.parse(r.text);
      if (!Array.isArray(arr)) throw new Error('不是数组');
      CONVS = arr;
      if (!CONVS.length) CONVS = [{ id: uid(), title: '新对话', messages: [], updatedAt: Date.now() }];
      currentId = CONVS[0].id;
      saveState(); renderConvList(); renderChat();
      setStatus('导入成功（' + CONVS.length + ' 个对话）');
    } catch (e) { setStatus('导入失败：' + (e.message || e)); }
  };

  $('send').onclick = async () => {
    syncSettings();
    const text = $('input').value.trim();
    if (!text) return;
    if (!SETTINGS.apikey) { setStatus('请先在 ⚙ 设置 中填写 API Key'); return; }
    const c = curConv(); if (!c) return;

    let msg = text;
    const attachParts = [];
    if (attachedFiles.length) {
      const ft = trimContext(attachedFiles.map(f => '=== ' + f.name + ' ===\n' + f.text).join('\n\n'), 30000);
      msg += '\n\n[已上传文件' + (ft.truncated ? '（过长已截断）' : '') + ']\n' + ft.text;
      attachParts.push(attachedFiles.length + ' 个文件');
    }
    if (SETTINGS.autoAttach) {
      setStatus('正在读取 Excel 上下文（' + SETTINGS.ctxMode + '）…');
      const data = await collectContext(SETTINGS.ctxMode);
      const raw = JSON.stringify(data);
      const ct = trimContext(raw, 20000);
      msg += '\n\n[Excel 上下文（' + SETTINGS.ctxMode + '，已自动附加' + (ct.truncated ? '，过长已截断' : '') + '）]\n' + ct.text;
      attachParts.push('上下文 ' + (raw.length / 1024).toFixed(1) + ' KB' + (ct.truncated ? '（已截断）' : ''));
    }

    const ub = document.createElement('div');
    ub.className = 'bubble user';
    ub.textContent = text;
    $('chat').appendChild(ub);
    $('input').value = '';
    $('log').innerHTML = '';
    if (attachParts.length) log('ctx', '📎 已附加：' + attachParts.join(' · '));

    c.messages.push({ role: 'user', content: msg, display: text });
    refreshEmptyState();
    if (c.title === '新对话') c.title = text.slice(0, 20);
    saveState(); bumpConv(); renderConvList();
    $('chat').scrollTop = $('chat').scrollHeight;
    try { await runAgent(); } catch (e) {
      const m = e.message || e;
      setStatus(String(m).includes('已停止') ? '⏹ 已停止生成' : '✗ 出错：' + m);
    }
  };

  $('input').addEventListener('keydown', handleInputKey);
}

/* ================= 启动 ================= */
/* 保活：任务窗格打开期间每 2 分钟 ping 本地服务器，窗格关闭后服务器将在 15 分钟空闲后自动退出 */
if (typeof location !== 'undefined' && location.origin) {
  setInterval(() => {
    try { fetch(location.origin + '/ping').catch(() => {}); } catch (e) { /* 忽略 */ }
  }, 120000);
}
loadState();
applyTheme();
applySidebar();
populateSettings();
bindEvents();
renderConvList();
renderChat();

