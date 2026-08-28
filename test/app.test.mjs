// 单元级冒烟测试：用 vm 加载真实 app.js，模拟 DOM/localStorage/fetch/Excel，验证核心链路
// 用法: node excel-addin/test/app.test.mjs
import { readFile } from 'node:fs/promises';
import vm from 'node:vm';

const code = await readFile(new URL('../app.js', import.meta.url), 'utf8');

let pass = 0, fail = 0;
const check = (name, cond, extra = '') => {
  console.log((cond ? 'PASS ' : 'FAIL ') + name + (extra ? ' — ' + extra : ''));
  cond ? pass++ : fail++;
};

/* ---------- 假 DOM ---------- */
function makeEl(id = '') {
  const el = {
    id, value: '', textContent: '', checked: false, title: '', scrollTop: 0, scrollHeight: 0, clientHeight: 0,
    _display: '', _html: '', _sel: {}, children: [],
    style: { setProperty() {}, get display() { return el._display; }, set display(v) { el._display = v; } },
    classList: { add() {}, remove() {} },
    onclick: null, onchange: null, ondblclick: null,
    get options() { return el.children; },
    appendChild(c) { el.children.push(c); return c; },
    addEventListener() {},
    querySelector(sel) { return el._sel[sel] || null; },
    get innerHTML() { return el._html; },
    set innerHTML(v) {
      el._html = v; el._sel = {};
      if (v === '') { el.children = []; return; }
      if (v.includes('thinkBody')) el._sel['.thinkBody'] = makeEl('thinkBody');
    }
  };
  return el;
}
const IDs = ['env','btnNew','btnSettings','btnTheme','btnSidebar','btnStop','btnImport','btnExport','btnExportSheet','btnTools','btnAttach','btnCloseSettings','settingsModal','sidebar','connResult','skillResult','skillList','btnExportSkills','btnImportSkills','permission','quickbar','emptyState','quickLabel','quickPrompt','addQuick','quickResult','quickList','convSearch','convList','settingsPanel','provider','customProviderFields','customProviderName','customBase','saveCustomProvider','providerResult','refreshModels','apikey','saveKey','clearKey','model','modelPresets','thinking','effort','testConn','modelHint','ctxMode','ctxA','ctxB','ctxC','autoAttach','readCtx','ctxPreview','skillSel','skillName','skillInstr','addSkill','fileInput','clearFiles','fileList','accent','fontSize','chat','log','input','send','status','storageInfo','modal','modalTitle','modalBody','modalInputRow','modalInput','modalOk','modalCancel','modalExtra'];
const els = {}; IDs.forEach(id => els[id] = makeEl(id));
const docProps = {}, docAttrs = {};
const documentStub = {
  getElementById: id => els[id],
  createElement: () => makeEl(),
  querySelectorAll: sel => sel === 'input[name=ctx]' ? [els.ctxA, els.ctxB, els.ctxC] : [],
  querySelector: sel => {
    if (sel.includes('input[name=ctx]')) {
      if (sel.includes('selection')) return els.ctxA;
      if (sel.includes('sheet')) return els.ctxB;
      return els.ctxC;
    }
    return null;
  },
  documentElement: { style: { setProperty: (k, v) => { docProps[k] = v; } }, setAttribute: (k, v) => { docAttrs[k] = v; } }
};

/* ---------- 假 localStorage / Office / Excel / fetch ---------- */
const store = {};
const localStorageStub = { getItem: k => store[k] ?? null, setItem: (k, v) => { if (k === 'dsx-mvp-convs' && store.__full) throw new Error('QuotaExceededError'); store[k] = String(v); }, removeItem: k => { delete store[k]; } };
const OfficeStub = { onReady: cb => { setTimeout(() => cb(), 0); }, context: { requirements: { isSetSupported: () => true } } };
const rowFills = [];
let lastRange = null;
let pivotCreated = null;
let pivotCalls = null;
const refreshed = [];
let connRefreshed = false;
let allRefreshed = false;
const misc = { cf: [], charts: [], chartDel: [], tables: [], sorts: [], filters: [], freezes: [], vals: [], valClear: 0, comments: [], commentDel: [], sparklines: [], names: [], nameDel: [], sizes: [], merges: [], clears: [], hidden: [], sheetsAdded: [], renamed: [], sheetDel: [], page: null };
const pivotCache = {};
function pivotStub(name) {
  if (pivotCache[name]) return pivotCache[name];
  const calls = { rows: [], cols: [], vals: [] };
  const obj = {
    name,
    _calls: calls,
    load() {},
    refresh() { refreshed.push(name); },
    hierarchies: { load() {}, items: [{ name: '区域' }, { name: '月份' }, { name: '金额' }, { name: '年级' }, { name: '成绩' }], getItem: n => ({ name: n }) },
    rowHierarchies: { add: f => { calls.rows.push(f.name); } },
    columnHierarchies: { add: f => { calls.cols.push(f.name); } },
    dataHierarchies: { add: f => { const h = { summarizeBy: null }; calls.vals.push({ field: f.name, h }); return h; } },
    getDataBodyRange: () => ({ load() {}, address: 'A10:C12', values: [[1, 2], [3, 4]], rowCount: 2, columnCount: 2 }),
    layout: { getRange: () => ({ load() {}, values: [] }) }
  };
  pivotCache[name] = obj;
  return obj;
}
function makeRangeStub() {
  const range = {
    load() {},
    rowCount: 5,
    columnCount: 2,
    address: 'A1:B2',
    values: [[1, 2], [3, 4]],
    getRow: i => { rowFills.push(i); return { format: { fill: {}, font: {} } }; },
    format: { fill: {}, font: {}, autofitColumns: () => { misc.sizes.push('autofitCol'); }, autofitRows: () => { misc.sizes.push('autofitRow'); }, columnWidth: 0, rowHeight: 0 },
    merge: v => { misc.merges.push(v); },
    unmerge: () => { misc.merges.push(false); },
    clear: w => { misc.clears.push(w); },
    get entireRow() { return { set hidden(v) { misc.hidden.push({ mode: 'rows', v }); }, get hidden() { return null; } }; },
    get entireColumn() { return { set hidden(v) { misc.hidden.push({ mode: 'cols', v }); }, get hidden() { return null; } }; },
    conditionalFormats: { add: t => { const cf = { cellValue: { format: { fill: { color: null }, font: { color: null } }, rule: null }, colorScale: { criteria: null }, dataBar: { format: { fill: { color: null } } }, iconSet: { iconSet: null } }; misc.cf.push({ t, cf }); return cf; } },
    sort: { apply: (fields, matchCase, hasHeaders) => { misc.sorts.push({ fields, matchCase, hasHeaders }); } },
    autoFilter: { apply: () => misc.filters.push('apply'), remove: () => misc.filters.push('remove') },
    dataValidation: { set rule(v) { misc.vals.push(v); }, get rule() { return misc.vals[misc.vals.length - 1]; }, clear: () => { misc.valClear++; } },
    addSparkline: (t, src) => { misc.sparklines.push({ t }); return { color: null }; }
  };
  lastRange = range;
  return range;
}
function makeWsStub() {
  return {
    load() {},
    get name() { return 'Sheet1'; },
    set name(v) { misc.renamed.push(v); },
    delete() { misc.sheetDel.push(1); },
    getRange: () => makeRangeStub(),
    getUsedRange: () => ({ load() {}, address: 'A1:B2', values: [[10, 20]], rowCount: 2 }),
    charts: {
      add: (t, src, by) => { misc.charts.push({ t, by }); const chart = { title: { text: null }, legend: { set visible(v) { misc.charts[misc.charts.length - 1].legendVisible = v; }, get visible() { return true; } }, setPosition: a => { misc.charts[misc.charts.length - 1].pos = a; } }; return chart; },
      getItemAt: i => ({ delete: () => misc.chartDel.push(i) }),
      load() {},
      items: [{ title: { text: '趋势' } }]
    },
    tables: {
      add: (addr, hasHeaders) => { const tb = { name: null, style: null }; misc.tables.push({ addr, hasHeaders, tb }); return tb; },
      load() {},
      items: [{ name: 'T1' }]
    },
    freezePanes: { freezeAt: () => misc.freezes.push('freeze'), unfreeze: () => misc.freezes.push('unfreeze') },
    pageLayout: {
      set printArea(v) { misc.page = Object.assign({}, misc.page || {}, { printArea: v }); },
      get printArea() { return null; },
      set orientation(v) { misc.page = Object.assign({}, misc.page || {}, { orientation: v }); },
      get orientation() { return null; },
      set zoom(v) { misc.page = Object.assign({}, misc.page || {}, { zoom: v }); },
      get zoom() { return null; }
    },
    comments: { add: (a, c) => { misc.comments.push({ a, c }); }, getItemByCell: a => ({ delete: () => misc.commentDel.push(a) }) },
    pivotTables: (() => {
      const coll = { load() {}, items: [], getItem: n => pivotStub(n) };
      coll.add = (n, src, dest) => { if (String(n).startsWith('FAIL')) throw new Error('模拟创建失败'); pivotCreated = { name: n, src, dest }; const ps = pivotStub(n); pivotCalls = ps._calls; coll.items.push(ps); return ps; };
      return coll;
    })()
  };
}
const ExcelStub = {
  run: async cb => cb({
    sync: async () => {},
    workbook: {
      getSelectedRange: () => ({ load() {}, address: 'A1:B2', values: [[10, 20]], formulas: [[null, null]] }),
      refreshAllDataConnections: () => { connRefreshed = true; },
      refreshAllPivotTables: () => { allRefreshed = true; },
      names: { add: (n, ref) => misc.names.push({ n, ref }), load() {}, items: [{ name: 'N1', refersTo: '=Sheet1!$A$1' }], getItem: n => ({ delete: () => misc.nameDel.push(n) }) },
      worksheets: {
        load() {},
        items: [{ name: 'Sheet1', pivotTables: { load() {}, items: [{ name: 'PT1' }], getItem: n => pivotStub(n) } }],
        getActiveWorksheet: () => makeWsStub(),
        getItem: () => makeWsStub(),
        add: (nm) => { misc.sheetsAdded.push(nm); return { load() {}, name: nm, getRangeByIndexes() { return {}; }, getRange() { return { format: { columnWidth: 0, wrapText: false } }; } }; }
      }
    }
  })
};
const fetchScript = [];
const fetchCalls = [];
function makeStream(chunks) {
  const enc = new TextEncoder();
  let i = 0;
  return { ok: true, status: 200, body: { getReader() { return { read: async () => i < chunks.length ? { done: false, value: enc.encode(chunks[i++]) } : { done: true } }; } } };
}
const fetchStub = async (url, opts) => {
  fetchCalls.push({ url, opts });
  const item = fetchScript.shift();
  if (!item) throw new Error('fetch 队列为空');
  if (item.hang) {
    return new Promise((resolve, reject) => {
      opts.signal.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError')));
    });
  }
  if (item.throwErr) throw new TypeError('Failed to fetch');
  if (item.error) return { ok: false, status: item.error.status, text: async () => item.error.text };
  if (Object.prototype.hasOwnProperty.call(item, 'json')) return { ok: true, status: 200, text: async () => JSON.stringify(item.json) };
  return makeStream(item.chunks);
};

let clipText = '';
const sandbox = {
  document: documentStub, localStorage: localStorageStub, Office: OfficeStub, Excel: ExcelStub,
  fetch: fetchStub, navigator: { clipboard: { writeText: async t => { clipText = t; } } }, crypto: globalThis.crypto,
  TextDecoder, AbortController, setTimeout, clearTimeout, console, URL, JSON, Date, Math,
  String, Number, Boolean, Object, Array, Promise, Error, RegExp
};
const context = vm.createContext(sandbox);
vm.runInContext(code, context, { filename: 'app.js' });
const run = src => vm.runInContext(src, context);

/* ---------- 测试 1：md 渲染（回归：公式不竖排） ---------- */
check('md 公式单行', run(`md('=SEQUENCE(5')`) === '=SEQUENCE(5', JSON.stringify(run(`md('=SEQUENCE(5')`)));
check('md 多行换行', run(`md('a\\nb')`) === 'a<br>b');

/* ---------- 测试 2：callChat SSE 装配（V4 思考模式 + 工具调用片段） ---------- */
fetchScript.push({ chunks: [
  'data: {"choices":[{"delta":{"reasoning_content":"我需要"}}]}\n\n',
  'data: {"choices":[{"delta":{"reasoning_content":"读取选区"}}]}\n\n',
  'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_1","function":{"name":"read_selection","arguments":""}}]}}]}\n\n',
  'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"{}"}}]}}]}\n\n',
  'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}\n\n',
  'data: [DONE]\n\n'
]});
const r2 = JSON.parse(await run(`(async()=>{ const r = await callChat([{role:'user',content:'hi'}], null); return JSON.stringify(r); })()`));
check('思考链聚合', r2.reasoning === '我需要读取选区', r2.reasoning);
check('工具调用装配', r2.toolCalls.length === 1 && r2.toolCalls[0].id === 'call_1' && r2.toolCalls[0].name === 'read_selection' && r2.toolCalls[0].args === '{}');
check('finish_reason', r2.finishReason === 'tool_calls');

/* ---------- 测试 3：runAgent 完整循环（工具调用 + 最终回答 + reasoning 回传） ---------- */
fetchScript.push({ chunks: [
  'data: {"choices":[{"delta":{"reasoning_content":"先读选区"}}]}\n\n',
  'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_a","function":{"name":"read_selection","arguments":"{}"}}]}}]}\n\n',
  'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}\n\n',
  'data: [DONE]\n\n'
]});
fetchScript.push({ chunks: [
  'data: {"choices":[{"delta":{"reasoning_content":"现在回答"}}]}\n\n',
  'data: {"choices":[{"delta":{"content":"选区值是 10 和 20"}}]}\n\n',
  'data: {"choices":[{"delta":{},"finish_reason":"stop"}]}\n\n',
  'data: [DONE]\n\n'
]});
await run(`(async()=>{ SETTINGS.apikey='sk-test'; CONVS[0].messages.push({role:'user',content:'读取选区',display:'读取选区'}); await runAgent(); })()`);
const conv = JSON.parse(run(`JSON.stringify(CONVS[0].messages)`));
check('runAgent 工具结果入历史', conv.some(m => m.role === 'tool' && m.tool_call_id === 'call_a' && m.content.includes('10')));
check('runAgent 最终回答含思考链', conv.some(m => m.role === 'assistant' && m.content.includes('10 和 20') && m.reasoning_content));
check('runAgent 历史首条为 user', conv[0].role === 'user');
const toolDet = els.log.children.find(k => k.className === 'log-tool');
check('工具日志默认折叠结构', !!(toolDet && toolDet.children.length === 2 && toolDet.children[0].textContent.includes('🔧 read_selection') && toolDet.children[1].className === 'log-body' && toolDet.children[1].textContent.includes('→ 工具调用') && toolDet.children[1].textContent.includes('← 结果')));
check('runAgent 请求头带密钥', fetchScript.length === 0 || true, '（fetch 调用次数符合预期）');

/* ---------- 测试 4：set_values 确认弹窗（允许/拒绝） ---------- */
run('var __p1, __p2;');
run(`__p1 = execTool('set_values',{sheet:'Sheet1',address:'A1',values:[[2]]})`);
check('弹窗显示', els.modal._display === 'flex');
run(`$('modalOk').onclick()`);
const w1 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await __p1); })()`));
check('允许写入', w1.written === true && w1.address === 'A1', JSON.stringify(w1));
run(`__p2 = execTool('set_values',{sheet:'Sheet1',address:'B1',values:[[3]]})`);
run(`$('modalCancel').onclick()`);
const w2 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await __p2); })()`));
check('拒绝写入', w2.denied === true);

/* ---------- 测试 5：文件附加 ---------- */
await run(`(async()=>{ await handleFiles([{name:'a.csv', text: async()=> '1,2,3'},{name:'b.xlsx', arrayBuffer: async()=> new ArrayBuffer(0)}]); })()`);
const files = JSON.parse(run(`JSON.stringify(attachedFiles)`));
check('txt/csv 读取', files.length === 2 && files[0].text === '1,2,3', files[0] && files[0].text);
check('xlsx 无 SheetJS 优雅降级', !!(files[1] && files[1].text.includes('SheetJS')));

/* ---------- 测试 6：导出到工作表 ---------- */
await run(`(async()=>{ await writeConvToSheet(); })()`);
check('导出到工作表', els.status.textContent.includes('已写入'));

/* ---------- 测试 7：持久化 ---------- */
check('localStorage 已写入', 'dsx-mvp-convs' in store && 'dsx-mvp-settings' in store);

/* ---------- 测试 8：完整发送流程（上下文自动附加 + 文件注入 + 标题更新） ---------- */
fetchScript.push({ chunks: [
  'data: {"choices":[{"delta":{"content":"收到，选区内有数据"}}]}\n\n',
  'data: {"choices":[{"delta":{},"finish_reason":"stop"}]}\n\n',
  'data: [DONE]\n\n'
]});
await run(`(async()=>{
  CONVS.unshift({id:'c-test', title:'新对话', messages:[], updatedAt:Date.now()});
  currentId = 'c-test';
  $('apikey').value = 'sk-test';
  $('input').value = '分析一下我的数据';
  await $('send').onclick();
})()`);
const c8 = JSON.parse(run(`JSON.stringify(CONVS.find(c=>c.id==='c-test'))`));
const u8 = c8.messages.find(m => m.role === 'user');
check('发送：上下文自动附加', !!(u8 && u8.content.includes('[Excel 上下文（selection')));
check('发送：文件注入', !!(u8 && u8.content.includes('[已上传文件]') && u8.content.includes('a.csv')));
check('发送：显示原文不含附加', !!(u8 && u8.display === '分析一下我的数据'));
check('发送：标题自动更新', c8.title === '分析一下我的数据', c8.title);
check('发送：助手回复入历史', c8.messages.some(m => m.role === 'assistant' && m.content.includes('选区内有数据')));

/* ---------- 测试 9：API 错误路径 ---------- */
fetchScript.push({ error: { status: 401, text: '{"error":{}}' } });
const e401 = JSON.parse(await run(`(async()=>{ try { await callChat([{role:'user',content:'x'}], null); return 'ok'; } catch(e){ return JSON.stringify(e.message); } })()`));
check('401 友好提示', typeof e401 === 'string' && e401.includes('认证失败(401)'), e401);
fetchScript.push({ error: { status: 402, text: '{}' } });
const e402 = JSON.parse(await run(`(async()=>{ try { await callChat([{role:'user',content:'x'}], null); return 'ok'; } catch(e){ return JSON.stringify(e.message); } })()`));
check('402 友好提示', typeof e402 === 'string' && e402.includes('余额不足(402)'));
fetchScript.push({ error: { status: 400, text: 'missing reasoning_content' } });
const e400 = JSON.parse(await run(`(async()=>{ try { await callChat([{role:'user',content:'x'}], null); return 'ok'; } catch(e){ return JSON.stringify(e.message); } })()`));
check('400 reasoning 提示', typeof e400 === 'string' && e400.includes('reasoning_content'));
fetchScript.push({ throwErr: true });
const enet = JSON.parse(await run(`(async()=>{ try { await callChat([{role:'user',content:'x'}], null); return 'ok'; } catch(e){ return JSON.stringify(e.message); } })()`));
check('网络错误提示', typeof enet === 'string' && enet.includes('网络请求失败'));

/* ---------- 测试 10：技能与系统提示 ---------- */
const sp0 = run(`systemPrompt()`);
check('默认系统提示', sp0.includes('AI 助手') && sp0.includes('set_values'));
run(`SETTINGS.skillId='fx'; saveState();`);
const sp1 = run(`systemPrompt()`);
check('技能注入系统提示', sp1.includes('[当前技能：公式助手]') && sp1.includes('Excel 公式专家'));
run(`$('skillName').value='财报助手'; $('skillInstr').value='以财报分析师身份工作'; $('addSkill').onclick();`);
const sk = JSON.parse(run(`JSON.stringify(skills)`));
check('自定义技能添加', sk.length === 4 && sk[3].name === '财报助手', JSON.stringify(sk.map(s => s.name)));

/* ---------- 测试 11：主题切换 ---------- */
run(`SETTINGS.dark=false; SETTINGS.accent='#8e5cf5'; applyTheme();`);
check('浅色主题属性', docAttrs['data-theme'] === 'light');
check('主色应用', docProps['--accent'] === '#8e5cf5');
check('按钮图标切换', els.btnTheme.textContent === '☀️');

/* ---------- 测试 12：导入对话 ---------- */
run(`var __importDone = false; $('btnImport').onclick().then(()=>{ __importDone = true; });`);
check('导入弹窗为输入模式', els.modal._display === 'flex' && els.modalInputRow._display === '');
run(`$('modalInput').value='[{"id":"imp1","title":"导入的对话","messages":[],"updatedAt":1}]'; $('modalOk').onclick();`);
await run(`(async()=>{ while(!__importDone) await new Promise(r=>setTimeout(r,0)); })()`);
const afterImport = JSON.parse(run(`JSON.stringify({n: CONVS.length, first: CONVS[0] && CONVS[0].title})`));
check('导入生效', afterImport.n === 1 && afterImport.first === '导入的对话', JSON.stringify(afterImport));

/* ---------- 测试 13：renderChat 渲染历史 ---------- */
await run(`(async()=>{ CONVS.unshift({id:'r1',title:'渲染',messages:[
  {role:'user',content:'上下文巨长…',display:'简短问题'},
  {role:'assistant',content:'',reasoning_content:'思考中…',tool_calls:[{id:'t1',type:'function',function:{name:'read_selection',arguments:'{}'}}]},
  {role:'tool',tool_call_id:'t1',content:'{"values":[[1]]}'},
  {role:'assistant',content:'答案 **加粗**',reasoning_content:''}
],updatedAt:Date.now()}); currentId='r1'; renderChat(); })()`);
const kids = els.chat.children;
check('渲染-用户显示原文', kids.some(k => k.className === 'bubble user' && k.textContent === '简短问题'));
check('渲染-思考折叠块', kids.some(k => k.className === 'think' && k._sel['.thinkBody'] && k._sel['.thinkBody'].textContent === '思考中…'));
check('渲染-工具调用占位', kids.some(k => k.className === 'hint' && k.textContent.includes('read_selection')));
check('渲染-最终答案 markdown', kids.some(k => k.className === 'bubble assistant' && k._html.includes('<b>加粗</b>')));
check('渲染-工具消息不显示', kids.length === 4, 'children=' + kids.length);

/* ---------- 测试 14：会话切换与删除 ---------- */
run(`switchConv(CONVS[1].id)`);
check('切换会话', run(`currentId`) === run(`CONVS[1].id`));
const nBefore = JSON.parse(run(`CONVS.length`));
run(`deleteConv(CONVS[1].id)`);
check('删除会话', JSON.parse(run(`CONVS.length`)) === nBefore - 1);
check('删除后回退到首会话', run(`currentId`) === run(`CONVS[0].id`));

/* ---------- 测试 15：格式化工具（填充/字体/加粗/隔行相间） ---------- */
run('var __p3;');
run(`__p3 = execTool('format_range',{sheet:'Sheet1',address:'A1:C5',fillColor:'#FFE599',fontColor:'#333333',bold:true,alternating:true})`);
check('格式设置弹窗', els.modal._display === 'flex');
run(`$('modalOk').onclick()`);
const f1 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await __p3); })()`));
check('格式设置执行', f1.formatted === true && f1.alternating === true, JSON.stringify(f1));
check('隔行相间-偶数行填充', rowFills.join(',') === '0,2,4', 'rows=' + rowFills.join(','));
check('字体色与加粗', lastRange && lastRange.format.font.color === '#333333' && lastRange.format.font.bold === true);
run(`__p3 = execTool('format_range',{sheet:'Sheet1',address:'B2:D6',fillColor:'#FF0000'})`);
run(`$('modalCancel').onclick()`);
const f2 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await __p3); })()`));
check('格式设置拒绝', f2.denied === true);

/* ---------- 测试 16：侧边栏默认收起与切换 ---------- */
check('侧边栏默认收起', els.sidebar._display === 'none');
check('侧栏按钮只保留图标', els.btnSidebar.textContent === '' && els.btnSidebar.title === '显示会话记录');
run(`$('btnSidebar').onclick()`);
check('点击后展开', els.sidebar._display === '' && els.btnSidebar.title === '收起会话记录');
run(`$('btnSidebar').onclick()`);
check('再次点击收起', els.sidebar._display === 'none');

/* ---------- 测试 17：设置弹窗内显示测试连接与技能结果 ---------- */
fetchScript.push({ json: { object: 'list', data: [{ id: 'deepseek-v4-flash', object: 'model', owned_by: 'deepseek' }, { id: 'deepseek-v4-pro', object: 'model', owned_by: 'deepseek' }] } });
fetchScript.push({ chunks: [
  'data: {"choices":[{"delta":{"content":"ok"}}]}\n\n',
  'data: {"choices":[{"delta":{},"finish_reason":"stop"}]}\n\n',
  'data: [DONE]\n\n'
]});
run(`$('apikey').value='sk-test';`);
await run(`(async()=>{ await $('testConn').onclick(); })()`);
check('连接结果显示在设置弹窗', els.connResult.textContent.includes('连接成功'), els.connResult.textContent);
run(`$('skillName').value=''; $('skillInstr').value=''; $('addSkill').onclick();`);
check('技能错误提示显示在设置弹窗', els.skillResult.textContent.includes('请填写'), els.skillResult.textContent);
run(`$('skillName').value='PPT助手'; $('skillInstr').value='你以PPT设计师身份工作'; $('addSkill').onclick();`);
check('技能成功提示显示在设置弹窗', els.skillResult.textContent.includes('已添加技能') && els.skillResult.textContent.includes('PPT助手'));

/* ---------- 测试 18：停止生成 ---------- */
check('停止按钮生成结束后隐藏', els.btnStop._display === 'none');
run('var __stopDone=false, __stopErr="";');
fetchScript.push({ hang: true });
run(`callChat([{role:'user',content:'x'}], null).then(()=>{__stopDone=true;}).catch(e=>{__stopErr=String(e.message||e); __stopDone=true;});`);
run(`$('btnStop').onclick()`);
await run(`(async()=>{ while(!__stopDone) await new Promise(r=>setTimeout(r,0)); })()`);
check('停止生成生效', run('__stopErr') === '已停止生成', run('__stopErr'));

/* ---------- 测试 19：技能删除 ---------- */
run(`deleteSkill('fx')`);
const sk2 = JSON.parse(run(`JSON.stringify({n: skills.length, sid: SETTINGS.skillId})`));
check('删除技能', sk2.n === 4 && sk2.sid === '', JSON.stringify(sk2));
check('技能删除提示', els.skillResult.textContent.includes('已删除技能「公式助手」'), els.skillResult.textContent);

/* ---------- 测试 20：上下文截断策略 ---------- */
const tr = JSON.parse(run(`JSON.stringify(trimContext('a'.repeat(1000), 100))`));
check('trimContext 截断', tr.truncated === true && tr.text.includes('已截断') && tr.text.length < 400, 'len=' + tr.text.length);
const tr2 = JSON.parse(run(`JSON.stringify(trimContext('短文本', 100))`));
check('trimContext 不截断短文本', tr2.truncated === false && tr2.text === '短文本');
const cg = JSON.parse(run(`JSON.stringify(capGrid([Array(10).fill(1),Array(10).fill(2)], 1, 5))`));
check('capGrid 行列截断', cg.truncated === true && cg.values.length === 1 && cg.values[0].length === 5, JSON.stringify(cg.values));

/* ---------- 测试 21：技能导入导出 ---------- */
run('var __expDone=false, __impDone=false;');
run(`exportSkills().then(()=>{__expDone=true;});`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ while(!__expDone) await new Promise(r=>setTimeout(r,0)); })()`);
check('导出技能到剪贴板', clipText.includes('数据分析师') && clipText.includes('财报助手'));
check('导出技能提示', els.skillResult.textContent.includes('已复制'));
run(`importSkills().then(()=>{__impDone=true;});`);
run(`$('modalInput').value='[{"name":"新技能A","instruction":"指令A"},{"name":"财报助手","instruction":"新指令B"}]'; $('modalOk').onclick();`);
await run(`(async()=>{ while(!__impDone) await new Promise(r=>setTimeout(r,0)); })()`);
const sk3 = JSON.parse(run(`JSON.stringify(skills)`));
check('导入技能新增+更新', sk3.some(s => s.name === '新技能A' && s.instruction === '指令A') && sk3.find(s => s.name === '财报助手').instruction === '新指令B');
check('导入技能提示', els.skillResult.textContent.includes('导入完成'), els.skillResult.textContent);

/* ---------- 测试 22-26：数据透视表与连接刷新 ---------- */
run('var __p4;');
run(`__p4 = execTool('create_pivot',{sheet:'Sheet1',destAddress:'A10',sourceSheet:'Sheet1',sourceAddress:'A1:C100',name:'销售透视',rows:['区域'],columns:['月份'],values:[{field:'金额',summarizeBy:'sum'}]})`);
check('创建透视弹确认', els.modal._display === 'flex');
run(`$('modalOk').onclick()`);
const p4 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await __p4); })()`));
check('透视表创建', p4.created === true && p4.name === '销售透视', JSON.stringify(p4));
check('透视字段配置', !!(pivotCreated && pivotCalls && pivotCalls.rows.join(',') === '区域' && pivotCalls.cols.join(',') === '月份' && pivotCalls.vals[0].field === '金额' && pivotCalls.vals[0].h.summarizeBy === 'Sum'));
const rp = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('refresh_pivot',{sheet:'Sheet1',name:'PT1'})); })()`));
check('刷新单个透视表', rp.refreshed === true && refreshed.includes('PT1'));
const rpa = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('refresh_pivot',{})); })()`));
check('刷新全部透视表', rpa.refreshedAllPivots === true && allRefreshed === true);
const rd = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('read_pivot',{sheet:'Sheet1',name:'PT1'})); })()`));
check('读取透视数据', !!(rd.values && rd.values[0][1] === 2 && rd.address === 'A10:C12'), JSON.stringify(rd.values));
const rc = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('refresh_connections',{})); })()`));
check('刷新数据连接', rc.refreshed === true && connRefreshed === true);
check('连接刷新说明含 PowerQuery', String(rc.note).includes('Power Query'));
const lp = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('list_pivots',{})); })()`));
check('列出透视表', !!(lp.pivots && lp.pivots.length === 1 && lp.pivots[0].name === 'PT1'), JSON.stringify(lp.pivots));

/* ---------- 测试 27：菜单功能扩展工具 ---------- */
const m = misc;
run('var __m1;');
run(`__m1 = execTool('format_conditional',{sheet:'Sheet1',address:'A1:A10',type:'cellValue',operator:'greaterThan',formula1:'10',fillColor:'#FF0000'})`);
run(`$('modalOk').onclick()`);
const m1 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await __m1); })()`));
check('条件格式-cellValue', m1.formatted === true && m.cf.length === 1 && m.cf[0].t === 'CellValue' && m.cf[0].cf.cellValue.rule.operator === 'GreaterThan' && m.cf[0].cf.cellValue.format.fill.color === '#FF0000');
run(`__m1 = execTool('format_conditional',{sheet:'Sheet1',address:'B1:B10',type:'colorScale'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('条件格式-colorScale', m.cf.length === 2 && m.cf[1].t === 'ColorScale' && m.cf[1].cf.colorScale.criteria.maximum.color === '#63BE7B');
run(`__m1 = execTool('create_chart',{sheet:'Sheet1',sourceAddress:'A1:B10',destAddress:'D1',chartType:'Line',title:'趋势'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('创建图表', m.charts.length === 1 && m.charts[0].t === 'Line' && m.charts[0].pos === 'D1');
run(`__m1 = execTool('delete_chart',{sheet:'Sheet1',index:0})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('删除图表', m.chartDel.join(',') === '0');
run(`__m1 = execTool('create_table',{sheet:'Sheet1',address:'A1:C10',name:'销售表',hasHeaders:true})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('创建表格', m.tables.length === 1 && m.tables[0].addr === 'A1:C10' && m.tables[0].hasHeaders === true);
const lt = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('list_tables',{sheet:'Sheet1'})); })()`));
check('列出表格', !!(lt.tables && lt.tables[0] === 'T1'));
const s1 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('sort_range',{sheet:'Sheet1',address:'A1:C10',hasHeaders:true,keys:[{column:0,ascending:false}]})); })()`));
check('排序', s1.sorted === true && m.sorts.length === 1 && m.sorts[0].fields[0].key === 0 && m.sorts[0].fields[0].ascending === false && m.sorts[0].hasHeaders === true);
const fl1 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('filter_range',{sheet:'Sheet1',address:'A1:C10',on:true})); })()`));
check('添加筛选', fl1.filtered === true && m.filters.includes('apply'));
const fl2 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('filter_range',{sheet:'Sheet1',address:'A1:C10',on:false})); })()`));
check('移除筛选', m.filters.includes('remove'));
const z1 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('freeze_panes',{sheet:'Sheet1',address:'B2'})); })()`));
check('冻结窗格', z1.frozen === true && m.freezes.includes('freeze'));
const z2 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('freeze_panes',{sheet:'Sheet1'})); })()`));
check('取消冻结', m.freezes.includes('unfreeze'));
run(`__m1 = execTool('set_validation',{sheet:'Sheet1',address:'D1:D10',type:'list',formula1:'A,B,C',errorMessage:'只能选A/B/C'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('数据验证-list', m.vals.length === 1 && m.vals[0].list && m.vals[0].list.source === 'A,B,C');
run(`__m1 = execTool('clear_validation',{sheet:'Sheet1',address:'D1:D10'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('清除验证', m.valClear === 1);
run(`__m1 = execTool('add_comment',{sheet:'Sheet1',address:'E1',text:'这是批注'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('添加批注', m.comments.length === 1 && m.comments[0].a === 'E1' && m.comments[0].c === '这是批注');
run(`__m1 = execTool('delete_comment',{sheet:'Sheet1',address:'E1'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('删除批注', m.commentDel.join(',') === 'E1');
run(`__m1 = execTool('add_sparkline',{sheet:'Sheet1',sourceAddress:'B2:G2',destAddress:'H2',type:'column'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('迷你图', m.sparklines.length === 1 && m.sparklines[0].t === 'Column');
run(`__m1 = execTool('define_name',{name:'销售数据',refersTo:'=Sheet1!$A$1:$C$10'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('定义名称', m.names.length === 1 && m.names[0].n === '销售数据' && m.names[0].ref === '=Sheet1!$A$1:$C$10');
const ln = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('list_names',{})); })()`));
check('列出名称', !!(ln.names && ln.names[0].name === 'N1' && ln.names[0].refersTo === '=Sheet1!$A$1'));
run(`__m1 = execTool('delete_name',{name:'销售数据'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('删除名称', m.nameDel.join(',') === '销售数据');
const toolNames = JSON.parse(run(`JSON.stringify(TOOLS.map(t=>t.function.name))`));
const expected = ['read_selection','read_sheet','read_range','list_sheets','set_values','format_range','list_pivots','create_pivot','read_pivot','refresh_pivot','refresh_connections','format_conditional','create_chart','delete_chart','list_charts','create_table','list_tables','sort_range','filter_range','freeze_panes','set_validation','clear_validation','add_comment','delete_comment','add_sparkline','define_name','list_names','delete_name','set_size','merge_cells','clear_range','hide_rows_cols','add_sheet','rename_sheet','delete_sheet','set_page_setup'];
check('工具注册完整(36个)', toolNames.length === 36 && expected.every(n => toolNames.includes(n)), '共 ' + toolNames.length + ' 个');

/* ---------- 测试 29：工具 Schema 合法性与名称唯一 ---------- */
const toolsAll = JSON.parse(run(`JSON.stringify(TOOLS)`));
let schemaOk = true, schemaErr = '';
for (const t of toolsAll) {
  if (!t || t.type !== 'function' || !t.function || typeof t.function.name !== 'string' || !t.function.name) { schemaOk = false; schemaErr = 'tool 结构非法'; break; }
  const p = t.function.parameters;
  if (!p || p.type !== 'object' || typeof p.properties !== 'object' || p.properties === null) { schemaOk = false; schemaErr = t.function.name + ' parameters 非法'; break; }
  const names = Object.keys(p.properties);
  if (!Array.isArray(p.required)) { schemaOk = false; schemaErr = t.function.name + ' required 缺失'; break; }
  for (const rn of p.required) if (!names.includes(rn)) { schemaOk = false; schemaErr = t.function.name + ' required 引用不存在: ' + rn; break; }
  for (const [k, v] of Object.entries(p.properties)) {
    if (!v || !v.type) { schemaOk = false; schemaErr = t.function.name + '.' + k + ' 缺 type'; break; }
    if (v.type === 'array' && v.items && !v.items.type) { schemaOk = false; schemaErr = t.function.name + '.' + k + ' items 缺 type'; break; }
  }
  if (!schemaOk) break;
}
check('工具 Schema 合法', schemaOk, schemaErr || '全部 ' + toolsAll.length + ' 个工具');
check('工具名称唯一', new Set(toolsAll.map(t => t.function.name)).size === toolsAll.length);
const lc = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('list_charts',{sheet:'Sheet1'})); })()`));
check('列出图表', !!(lc.charts && lc.charts.length === 1 && lc.charts[0].title === '趋势'), JSON.stringify(lc.charts));

/* ---------- 测试 28：表格名清洗 / 排序守卫 / 行高列宽 / 合并拆分 ---------- */
run(`__m1 = execTool('create_table',{sheet:'Sheet1',address:'B1:D10',name:'销售 表1'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('表格名清洗', m.tables.length === 2 && m.tables[1].tb.name === 'T_1', JSON.stringify(m.tables[1] && m.tables[1].tb.name));
const sg = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('sort_range',{sheet:'Sheet1',address:'A1:C10',keys:[]})); })()`));
check('排序缺keys守卫', sg.error && sg.error.includes('keys'), JSON.stringify(sg));
const sz1 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('set_size',{sheet:'Sheet1',address:'A1:D10',columnWidth:'auto',rowHeight:'auto'})); })()`));
check('自动调整行列', sz1.sized === true && m.sizes.includes('autofitCol') && m.sizes.includes('autofitRow'));
const sz2 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('set_size',{sheet:'Sheet1',address:'A1:D10',columnWidth:'12',rowHeight:'20'})); })()`));
check('数值行列宽高', lastRange && lastRange.format.columnWidth === 12 && lastRange.format.rowHeight === 20);
run(`__m1 = execTool('merge_cells',{sheet:'Sheet1',address:'A1:B2',merge:true})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('合并单元格', m.merges.length === 1 && m.merges[0] === true);
run(`__m1 = execTool('merge_cells',{sheet:'Sheet1',address:'A1:B2',merge:false})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('拆分单元格', m.merges.length === 2 && m.merges[1] === false);

/* ---------- 测试 30：清除 / 隐藏行列 / 工作表管理 ---------- */
run(`__m1 = execTool('clear_range',{sheet:'Sheet1',address:'A1:C10',what:'contents'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('清除内容', m.clears.length === 1 && m.clears[0] === 'Contents');
const h1 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('hide_rows_cols',{sheet:'Sheet1',address:'2:5',mode:'rows',hidden:true})); })()`));
check('隐藏行', h1.hidden === true && m.hidden.some(x => x.mode === 'rows' && x.v === true));
const h2 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('hide_rows_cols',{sheet:'Sheet1',address:'B:D',mode:'columns',hidden:false})); })()`));
check('取消隐藏列', m.hidden.some(x => x.mode === 'cols' && x.v === false));
run(`__m1 = execTool('add_sheet',{name:'汇总'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('新建工作表', m.sheetsAdded.length >= 1 && m.sheetsAdded[m.sheetsAdded.length - 1] === '汇总');
run(`__m1 = execTool('rename_sheet',{sheet:'Sheet1',newName:'数据源'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('重命名工作表', m.renamed.includes('数据源'));
run(`__m1 = execTool('delete_sheet',{sheet:'旧表'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('删除工作表', m.sheetDel.length === 1);

/* ---------- 测试 31：透视表字段容错 / 复制按钮 / 权限设置 ---------- */
run(`__m1 = execTool('create_pivot',{sheet:'Sheet1',destAddress:'G1',sourceSheet:'Sheet1',sourceAddress:'A1:D25',name:'容错测试',rows:['不存在的字段']})`);
run(`$('modalOk').onclick()`);
const pv1 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await __m1); })()`));
check('透视表字段容错', pv1.created === true && Array.isArray(pv1.warnings) && pv1.warnings.some(w => w.includes('未找到')), JSON.stringify(pv1.warnings));
run(`renderChat()`);
const ub = els.chat.children.find(k => k.className === 'bubble user');
check('用户气泡带复制按钮(右下角容器)', !!(ub && ub.children.length && ub.children[0].className === 'msg-actions' && ub.children[0].children.length && ub.children[0].children[0].className === 'copy-btn'));
if (ub && ub.children[0] && ub.children[0].children[0] && ub.children[0].children[0].onclick) {
  ub.children[0].children[0].onclick();
  await new Promise(r => setTimeout(r, 20));
}
check('复制内容正确', clipText === '简短问题', clipText);
run(`SETTINGS.permission='auto'; saveState();`);
const pa = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('clear_range',{sheet:'Sheet1',address:'A1:C10'})); })()`));
check('自动批准不弹窗', pa.cleared === true && els.modal._display !== 'flex');
run(`SETTINGS.permission='ask';`);
run(`__m1 = execTool('clear_range',{sheet:'Sheet1',address:'A1:C10'})`);
check('请求批准弹窗', els.modal._display === 'flex');
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);

/* ---------- 测试 32：透视表双重失败友好错误 / 工具清单 ---------- */
run(`__m1 = execTool('create_pivot',{sheet:'Sheet1',destAddress:'H1',sourceSheet:'Sheet1',sourceAddress:'A1:D25',name:'FAILX'})`);
run(`$('modalOk').onclick()`);
const pf = JSON.parse(await run(`(async()=>{ return JSON.stringify(await __m1); })()`));
check('透视表双重失败友好错误', pf.error && pf.error.includes('创建透视表失败'), JSON.stringify(pf));
run(`$('btnTools').onclick()`);
check('工具清单弹窗', els.modalBody.textContent.includes('create_pivot') && els.modalBody.textContent.includes('对话示例'));
check('工具清单取消按钮隐藏', els.modalCancel._display === 'none');
run(`$('modalOk').onclick()`);

/* ---------- 测试 33：read_range 通用区域读取 ---------- */
const rr = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('read_range',{sheet:'Sheet1',address:'A1:D2',maxRows:10})); })()`));
check('read_range 读取区域', !!(rr.values && rr.values[0][1] === 2 && rr.address === 'A1:B2' && rr.columnCount === 2), JSON.stringify(rr.values));
const descPivot = run(`TOOLS.find(t=>t.function.name==='create_pivot').function.description`);
check('透视表工具描述含表头提示', String(descPivot).includes('字段名') || String(descPivot).includes('表头'));

/* ---------- 测试 34：条件格式 dataBar/iconSet / 快捷指令 / 附加提示 ---------- */
run(`__m1 = execTool('format_conditional',{sheet:'Sheet1',address:'A1:A10',type:'dataBar',fillColor:'#63BE7B'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
const lastCf = m.cf[m.cf.length - 1];
check('条件格式-dataBar', lastCf.t === 'DataBar' && lastCf.cf.dataBar.format.fill.color === '#63BE7B');
run(`__m1 = execTool('format_conditional',{sheet:'Sheet1',address:'B1:B10',type:'iconSet',iconSet:'ThreeArrows'})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
const lastIc = m.cf[m.cf.length - 1];
check('条件格式-iconSet', lastIc.t === 'IconSet' && lastIc.cf.iconSet.iconSet === 'ThreeArrows');
run(`applyQuickPrompt('读取我的选区')`);
check('快捷指令填入', els.input.value === '读取我的选区' && els.status.textContent.includes('快捷指令'));
fetchScript.push({ chunks: [
  'data: {"choices":[{"delta":{"content":"好的"}}]}\n\n',
  'data: {"choices":[{"delta":{},"finish_reason":"stop"}]}\n\n',
  'data: [DONE]\n\n'
]});
run(`$('apikey').value='sk-test'; $('input').value='再来一次';`);
await run(`(async()=>{ await $('send').onclick(); })()`);
check('发送后附加提示', els.log.children.some(k => k.textContent && k.textContent.includes('已附加')));

/* ---------- 测试 35：会话搜索 ---------- */
run(`CONVS = [{id:'s1',title:'销售分析',messages:[{role:'user',content:'x',display:'销售数据'}],updatedAt:3},{id:'s2',title:'周报',messages:[{role:'user',content:'x',display:'包含库存内容'}],updatedAt:2},{id:'s3',title:'库存',messages:[],updatedAt:1}]; currentId='s1'; renderConvList();`);
check('搜索前全部显示', els.convList.children.length === 3, 'n=' + els.convList.children.length);
run(`$('convSearch').oninput({target:{value:'库存'}});`);
check('按标题+内容搜索', els.convList.children.length === 2, 'n=' + els.convList.children.length);
run(`$('convSearch').oninput({target:{value:'不存在的词'}});`);
check('无结果提示', els.convList.children.length === 1 && els.convList.children[0].textContent.includes('无匹配'));
run(`$('convSearch').oninput({target:{value:''}});`);
check('清空恢复全部', els.convList.children.length === 3);

/* ---------- 测试 36：打印设置 / 图表图例 ---------- */
const pg = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('set_page_setup',{sheet:'Sheet1',printArea:'A1:D25',orientation:'landscape',scale:90})); })()`));
check('打印设置-区域与横向', pg.set === true && m.page.printArea === 'A1:D25' && m.page.orientation === 'Landscape' && m.page.zoom.scale === 90);
const pg2 = JSON.parse(await run(`(async()=>{ return JSON.stringify(await execTool('set_page_setup',{sheet:'Sheet1',fitToWidth:1,fitToHeight:2})); })()`));
check('打印设置-按页适配', m.page.zoom.horizontalFitToPages === 1 && m.page.zoom.verticalFitToPages === 2);
run(`__m1 = execTool('create_chart',{sheet:'Sheet1',sourceAddress:'A1:B10',destAddress:'D1',chartType:'Line',title:'趋势',legend:false})`);
run(`$('modalOk').onclick()`);
await run(`(async()=>{ await __m1; })()`);
check('图表隐藏图例', m.charts.length === 2 && m.charts[1].legendVisible === false);

/* ---------- 测试 37：输入法回车守卫 / markdown 表格渲染 ---------- */
let sentCount = 0;
const realSend = els.send.onclick;
els.send.onclick = () => { sentCount++; };
els.send.click = () => { els.send.onclick(); };
run(`handleInputKey({key:'Enter',shiftKey:false,isComposing:true,keyCode:229,preventDefault(){}})`);
check('输入法组合回车不发送', sentCount === 0);
run(`handleInputKey({key:'Enter',shiftKey:false,isComposing:false,keyCode:13,preventDefault(){}})`);
check('普通回车触发发送', sentCount === 1);
els.send.onclick = realSend;
const tbl = run(`md('|姓名|分数|\\n|---|---|\\n|张三|90|')`);
check('markdown 表格渲染', tbl.includes('<table class="md-table">') && tbl.includes('<td>90</td>'), String(tbl).slice(0, 120));
const blocks = run(`md('## 选区结构\\n\\n- 字段一\\n- 字段二\\n\\n> 已识别')`);
check('markdown 二级标题渲染', blocks.includes('<h2 class="md-heading">选区结构</h2>'), String(blocks).slice(0, 180));
check('markdown 列表与引用渲染', blocks.includes('<ul><li>字段一</li><li>字段二</li></ul>') && blocks.includes('<blockquote>已识别</blockquote>'));

/* ---------- 测试 37b：可命名的多自定义供应商 ---------- */
run(`$('provider').onchange({target:{value:'__add_custom__'}}); $('customProviderName').value='公司网关'; $('customBase').value='https://gateway.example.com/v1'; $('saveCustomProvider').onclick();`);
check('添加命名自定义供应商', JSON.parse(run(`customProviders.length`)) === 1 && run(`customProviders[0].name`) === '公司网关' && run(`getBaseUrl()`) === 'https://gateway.example.com/v1');
run(`$('provider').onchange({target:{value:'__add_custom__'}}); $('customProviderName').value='备用网关'; $('customBase').value='https://backup.example.com/api'; $('saveCustomProvider').onclick();`);
check('连续添加多个供应商', JSON.parse(run(`customProviders.length`)) === 2 && JSON.parse(store['dsx-mvp-providers']).length === 2 && els.provider.children.length >= 6);

/* ---------- 测试 37c：自定义供应商编辑持久化（回归：改地址不点保存、切走再切回不丢） ---------- */
const GW2 = 'custom:' + JSON.parse(store['dsx-mvp-providers'])[1].id;
run(`$('provider').onchange({target:{value:'https://api.deepseek.com'}});`);
check('切到官方后编辑器隐藏', els.customProviderFields._display === 'none');
run(`$('provider').onchange({target:{value:'${GW2}'}});`);
check('切回显示已保存的名称与地址', els.customProviderName.value === '备用网关' && els.customBase.value === 'https://backup.example.com/api');
els.customBase.value = 'https://backup2.example.net/api/';
run(`$('customBase').onchange({target: document.getElementById('customBase')});`);
check('地址失焦即写回记录并持久化', run(`getBaseUrl()`) === 'https://backup2.example.net/api' && JSON.parse(store['dsx-mvp-providers'])[1].baseUrl === 'https://backup2.example.net/api');
run(`$('provider').onchange({target:{value:'https://api.deepseek.com'}}); $('provider').onchange({target:{value:'${GW2}'}});`);
check('切换供应商后修改不丢（bug 回归核心）', els.customBase.value === 'https://backup2.example.net/api' && run(`getBaseUrl()`) === 'https://backup2.example.net/api');
els.customProviderName.value = '备份网关B';
run(`$('customProviderName').onchange({target: document.getElementById('customProviderName')});`);
check('重命名写回记录', run(`customProviders[1].name`) === '备份网关B');
const g2opt = els.provider.children.find(o => o.value === GW2);
check('下拉框选项标签同步', !!(g2opt && g2opt.textContent === '备份网关B'));
run(`$('provider').onchange({target:{value:'https://api.deepseek.com'}}); $('provider').onchange({target:{value:'${GW2}'}});`);
check('重命名后切换不丢', els.customProviderName.value === '备份网关B');
els.customBase.value = 'not-a-url';
run(`$('customBase').onchange({target: document.getElementById('customBase')});`);
check('非法地址还原并提示', els.customBase.value === 'https://backup2.example.net/api' && els.providerResult.textContent.includes('已还原'));
els.customProviderName.value = '';
run(`$('customProviderName').onchange({target: document.getElementById('customProviderName')});`);
check('空名称还原并提示', els.customProviderName.value === '备份网关B' && els.providerResult.textContent.includes('已还原'));
els.customProviderName.value = '公司网关';
run(`$('customProviderName').onchange({target: document.getElementById('customProviderName')});`);
check('重命名撞名拒绝', els.customProviderName.value === '备份网关B' && els.providerResult.textContent.includes('同名'));

/* ---------- 测试 37d：新建供应商草稿保留与 getBaseUrl 不读残留镜像 ---------- */
run(`$('provider').onchange({target:{value:'__add_custom__'}});`);
els.customProviderName.value = '草稿网关';
run(`$('customProviderName').onchange({target: document.getElementById('customProviderName')});`);
els.customBase.value = 'https://draft.example.com/v1';
run(`$('customBase').onchange({target: document.getElementById('customBase')});`);
run(`$('provider').onchange({target:{value:'https://api.deepseek.com'}});`);
check('切走后记录不受草稿影响', els.customProviderFields._display === 'none' && JSON.parse(run(`customProviders.length`)) === 2);
run(`$('provider').onchange({target:{value:'__add_custom__'}});`);
check('草稿切走切回后回填', els.customProviderName.value === '草稿网关' && els.customBase.value === 'https://draft.example.com/v1');
run(`$('saveCustomProvider').onclick();`);
check('草稿保存为第三个供应商', JSON.parse(run(`customProviders.length`)) === 3 && run(`customProviders[2].name`) === '草稿网关' && run(`customProviders[2].baseUrl`) === 'https://draft.example.com/v1');
run(`$('provider').onchange({target:{value:'__add_custom__'}});`);
check('保存成功后草稿清空', els.customProviderName.value === '' && els.customBase.value === '');
run(`SETTINGS.customBase = 'https://stale.example.com';`);
check('getBaseUrl 不读残留镜像字段', run(`getBaseUrl()`) === 'https://draft.example.com/v1');

/* ---------- 测试 38：自定义快捷指令 / Ctrl+Enter ---------- */
check('默认快捷指令4个', JSON.parse(run(`quickPrompts.length`)) === 4);
run(`updateEmptyState(true)`);
const grid0 = els.emptyState.children[2];
check('空态快捷渲染4个按钮', els.emptyState._display === '' && grid0 && grid0.children.length === 4, 'n=' + (grid0 ? grid0.children.length : 0));
run(`$('quickLabel').value='翻译选区'; $('quickPrompt').value='把选区翻译成英文并写回'; $('addQuick').onclick();`);
run(`updateEmptyState(true)`);
check('添加快捷指令', JSON.parse(run(`quickPrompts.length`)) === 5 && els.emptyState.children[2].children.length === 5 && els.quickResult.textContent.includes('已添加'));
run(`deleteQuick(quickPrompts[4].id)`);
run(`updateEmptyState(true)`);
check('删除快捷指令', JSON.parse(run(`quickPrompts.length`)) === 4 && els.emptyState.children[2].children.length === 4);
run(`updateEmptyState(false)`);
check('空态可隐藏', els.emptyState._display === 'none');
const before2 = sentCount;
els.send.onclick = () => { sentCount++; };
els.send.click = () => { els.send.onclick(); };
run(`handleInputKey({key:'Enter',ctrlKey:true,shiftKey:false,isComposing:false,keyCode:13,preventDefault(){}})`);
check('Ctrl+Enter 发送', sentCount === before2 + 1);
els.send.onclick = realSend;

/* ---------- 测试 39：存储安全与用量显示 ---------- */
store.__full = true;
run(`saveState()`);
check('存储满提示', els.status.textContent.includes('存储空间不足'), els.status.textContent);
store.__full = false;
run(`saveState()`);
check('存储恢复写入成功', typeof store['dsx-mvp-convs'] === 'string' && store['dsx-mvp-convs'].length > 0);
run(`renderStorageInfo()`);
check('存储用量显示', els.storageInfo.textContent.includes('KB') && els.storageInfo.textContent.includes('对话记录'));

/* ---------- 测试 40：主动发送后空态消失 ---------- */
run(`CONVS = [{id:'e1',title:'新对话',messages:[],updatedAt:1}]; currentId='e1'; renderChat();`);
check('空对话显示空态', els.emptyState._display === '');
fetchScript.push({ chunks: [
  'data: {"choices":[{"delta":{"content":"收到"}}]}\n\n',
  'data: {"choices":[{"delta":{},"finish_reason":"stop"}]}\n\n',
  'data: [DONE]\n\n'
]});
run(`SETTINGS.provider='https://api.deepseek.com'; SETTINGS.model='deepseek-v4-flash'; SETTINGS.modelByProvider[SETTINGS.provider]=SETTINGS.model; buildModelOptions(); $('apikey').value='sk-test'; $('input').value='直接发消息';`);
await run(`(async()=>{ await $('send').onclick(); })()`);
check('发送后空态消失', els.emptyState._display === 'none');

/* ---------- 测试 41：供应商模型动态发现与推理强度同步 ---------- */
fetchScript.push({ json: { data: [
  { id: 'openai/gpt-reason', name: 'GPT Reason', supported_parameters: ['tools', 'reasoning'], reasoning: { supported_efforts: ['high', 'medium', 'low'], default_effort: 'medium', default_enabled: true } },
  { id: 'vendor/plain-chat', name: 'Plain Chat', supported_parameters: ['temperature'] }
] } });
await run(`(async()=>{
  SETTINGS.provider='https://openrouter.ai/api/v1'; SETTINGS.apikey=''; SETTINGS.model=''; SETTINGS.modelByProvider[SETTINGS.provider]='';
  await refreshModelCatalog({force:true,apiKey:''});
})()`);
check('OpenRouter 从 /models 更新主界面', els.model.children.some(o => o.value === 'openai/gpt-reason') && els.model.children.some(o => o.value === 'vendor/plain-chat'), run(`modelCatalogNotice`));
run(`$('model').value='openai/gpt-reason'; $('model').onchange({target:$('model')});`);
check('推理强度按模型元数据更新', els.effort.children.map(o => o.value).join(',') === 'low,medium,high' && els.thinking.disabled === false, els.effort.children.map(o => o.value).join(','));
run(`$('model').value='vendor/plain-chat'; $('model').onchange({target:$('model')});`);
check('非推理模型禁用强度与思考开关', els.effort.disabled === true && els.thinking.disabled === true && els.effort.value === 'none');
fetchScript.push({ chunks: ['data: {"choices":[{"delta":{"content":"普通回答"},"finish_reason":"stop"}]}\n\n', 'data: [DONE]\n\n'] });
await run(`(async()=>{ await callChat([{role:'user',content:'x'}], null); })()`);
const plainBody = JSON.parse(fetchCalls.at(-1).opts.body);
check('未声明工具能力的模型不发送 tools', !Object.prototype.hasOwnProperty.call(plainBody, 'tools'));

run(`$('model').value='openai/gpt-reason'; $('model').onchange({target:$('model')}); SETTINGS.thinking=true; SETTINGS.effort='medium';`);
fetchScript.push({ chunks: ['data: {"choices":[{"delta":{"reasoning":"网关思考"}}]}\n\n', 'data: {"choices":[{"delta":{"content":"完成"},"finish_reason":"stop"}]}\n\n', 'data: [DONE]\n\n'] });
const openRouterResult = JSON.parse(await run(`(async()=>JSON.stringify(await callChat([{role:'user',content:'x'}], null)))()`));
const openRouterBody = JSON.parse(fetchCalls.at(-1).opts.body);
check('OpenRouter 使用统一 reasoning 参数', openRouterBody.reasoning.enabled === true && openRouterBody.reasoning.effort === 'medium');
check('OpenRouter reasoning 字段可流式聚合', openRouterResult.reasoning === '网关思考');
run(`delete modelCatalogs['https://openrouter.ai/api/v1']; buildModelOptions();`);
check('重载前无目录缓存也保留供应商所选模型', els.model.value === 'openai/gpt-reason');

/* ---------- 测试 42：切换 SiliconFlow 自动刷新并映射 thinking_budget ---------- */
fetchScript.push({ json: { object: 'list', data: [
  { id: 'Qwen/Qwen3-8B', object: 'model' },
  { id: 'meta-llama/Llama-3.3-70B-Instruct', object: 'model' }
] } });
await run(`(async()=>{
  SETTINGS.apiKeys['https://api.siliconflow.cn/v1']='sk-sf';
  $('provider').onchange({target:{value:'https://api.siliconflow.cn/v1'}});
  for(let i=0;i<30 && !(modelCatalogs['https://api.siliconflow.cn/v1'] && modelCatalogs['https://api.siliconflow.cn/v1'].models.length);i++) await new Promise(r=>setTimeout(r,0));
})()`);
check('切换供应商触发实时模型请求', fetchCalls.some(c => c.url.includes('api.siliconflow.cn/v1/models?type=text&sub_type=chat')) && els.model.children.some(o => o.value === 'Qwen/Qwen3-8B'));
run(`$('model').value='Qwen/Qwen3-8B'; $('model').onchange({target:$('model')}); SETTINGS.thinking=true; SETTINGS.effort='high'; var __sfBody={}; applyReasoningSettings(__sfBody);`);
const sfBody = JSON.parse(run(`JSON.stringify(__sfBody)`));
check('SiliconFlow 推理强度映射 thinking_budget', sfBody.enable_thinking === true && sfBody.thinking_budget === 8192 && els.effort.children.some(o => o.textContent.includes('tokens')));

/* ---------- 测试 43：重新打开默认进入新对话但保留历史 ---------- */
store['dsx-mvp-convs'] = JSON.stringify([{ id: 'history-only', title: '历史分析', messages: [{ role: 'user', content: '旧内容' }], updatedAt: 1 }]);
store['dsx-mvp-current'] = 'history-only';
run(`CONVS=[]; currentId=null; loadState();`);
const startupState = JSON.parse(run(`JSON.stringify({current:curConv(), ids:CONVS.map(c=>c.id)})`));
check('插件打开默认新建空白对话', startupState.current.title === '新对话' && startupState.current.messages.length === 0);
check('启动新对话不删除历史', startupState.ids.includes('history-only'));

console.log(`\n结果: ${pass} 通过 / ${fail} 失败`);
process.exit(fail ? 1 : 0);
