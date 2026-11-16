/* 生成 docs/agents/office-compat-matrix.md（收口+模板层版，分母 315） */
const fs = require('fs');
const routesSrc = fs.readFileSync('C:/Users/金东权/wps-skills/wps-multitool-mcp/src/client/action-routes.generated.ts', 'utf8');
const routes = {};
for (const m of routesSrc.matchAll(/"([A-Za-z0-9_]+)":\s*'(\w+)'/g)) routes[m[1]] = m[2];
const cases = Object.keys(routes);
if (cases.length !== 317) { console.error('路由表数量异常: ' + cases.length); process.exit(1); }

const live = new Set([
  'createDocument', 'insertText', 'addHeading', 'setFont', 'insertTable', 'getTableData',
  'getDocumentText', 'getDocumentStats', 'addPageNumber',
  'createWorkbook', 'setCellValue', 'getCellValue', 'setRangeData', 'getRangeData', 'setFormula',
  'setCellFormat', 'createChart', 'createExcelTable', 'autoFilter', 'sortRange',
  'createPresentation', 'addSlide', 'addTextBox', 'setTextBoxStyle', 'insertPptTable',
  'insertPptChart', 'getShapes', 'getSlideInfo', 'exportSlideAsImage', 'convertToPDF',
  'findReplace', 'addComment', 'getComments', 'addFootnote', 'insertHeader', 'setPageSetup',
  'addCaption', 'saveAs', 'getOpenDocuments', 'closeDocument', 'openDocument',
  'createSheet', 'switchSheet', 'renameSheet', 'setNumberFormat', 'setBorder', 'addCellComment',
  'getCellComments', 'freezePanes', 'protectSheet', 'unprotectSheet', 'copySheet', 'deleteSheet',
  'closeWorkbook', 'openWorkbook',
  'setSlideTitle', 'getSlideTitle', 'addShape', 'setShapePosition', 'setShapeRoundness',
  'setSlideBackground', 'setSlideNotes', 'getSlideNotes', 'duplicateSlide', 'moveSlide',
  'closePresentation', 'openPresentation', 'findInSheet',
  'addList', 'addPptHyperlink', 'removePptHyperlink', 'addMasterElement', 'getSlideMaster', 'setMasterBackground',
  'renderTemplate', 'extractTemplate', 'renderTemplatePpt', 'renderTemplateXlsx',
  'wpsStatus',
]);

const quirkLanded = {
  __slides: 'pwsh7 JSON 数值→Int64，Office Slides.Item 拒收 VT_I8；已统一 [int] 强转（97 处）',
  setCellValue: 'Sheets.Item(null)→Resolve-Sheet；NumberFormat General office 分支',
  getCellValue: 'Sheets.Item(null)→Resolve-Sheet',
  getRangeData: 'Sheets.Item(null)→Resolve-Sheet',
  setRangeData: "Office 拒 NumberFormat='General' 重置；office host 跳过",
  createChart: '字符串 chartType 过枚举表（Office 严格绑定）',
  createWorkbook: 'Resolve-Book 记账链（F3\' 范式）',
  openWorkbook: 'Resolve-Book 记账链（open 点）',
  switchWorkbook: 'Resolve-Book 记账链（switch 点）',
  closeWorkbook: '记账簿关闭清记账',
  quitWps: 'ProgID/残留清点/文案按 host 分支',
  wpsStatus: 'ProgID 与产品显示名按 host 分支；janitor 统计回显',
  wpsRepair: '诊断表 ProgID 按 host 取',
  getAppInfo: 'getter 已接 host 感知',
  saveContextSnapshot: 'Get-ActiveWpsApp 位点已接 Get-ProgidForApp',
  openFile: 'getter 已接 host 感知；Open 原生调用直通',
  cleanData: '裸 Sheets.Item($p.sheet) → Resolve-Sheet 化',
  switchSheet: '裸 Sheets.Item($p.sheet) 无守卫 → Resolve-Sheet 化（抽验真机绿）',
  updateChart: '手写 Sheets.Item → Resolve-Sheet 化',
  exportChartAsImage: '同 updateChart',
  exportRangeAsImage: '同 updateChart',
  deleteSheet: 'Resolve-Sheet 化（抽验真机绿）',
  renameSheet: 'Resolve-Sheet 化（抽验真机绿）',
  copySheet: 'Resolve-Sheet 化 + Copy $null 参数改 [Type]::Missing（抽验真机绿）',
  moveSheet: 'Resolve-Sheet 化',
  protectSheet: 'Resolve-Sheet 化 + Protect VT_NULL 假保护修正（抽验真机绿）',
  unprotectSheet: 'Resolve-Sheet 化（抽验真机绿）',
  setSlideTitle: 'Shapes.HasTitle 非标准属性 → try Shapes.Title（三轮迭代）',
  getSlideTitle: '同 setSlideTitle',
  addSlide: 'title 写入 try Shapes.Title；无占位符落文本框（等效实现+warning）',
  setPageSetup: '×28.35 是 WPS 边距单位语义；office host 不换行（抽验真机绿）',
  setShapeRoundness: 'Adjustments.Count=0 手柄防护（WPS 假成功/Office 诚实抛——抽验真机绿）',
  freezePanes: 'headless 实例 ActiveWindow=null——补可见窗口等效实现（抽验真机绿）',
  saveAs: '同名覆盖确认框卡死通道——DisplayAlerts 护栏（抽验真机绿）',
  convertToPDF: 'app 参数短路（抽验三 app 真机绿）',
  addCaption: 'Office 原生 SEQ 域路径真机通——对账结论：直通',
  findInSheet: '收口轮 B1 落地：office host 查找交 PowerShell 5.1 子进程（wps-find51.ps1）。终极破案：跨区域格式/批注使 UsedRange 变【多区域 Range】——Row/Value2 语义全变，四轮四错+两宿主差异统一到此根因；修法 = Areas foreach 枚举 + SAFEARRAY 非零基 Array.GetValue。三格式卡真机复验绿',
  renderTemplate: '模板层大卡落地：书签锚点优先+{{key}}兜底+缺 key 列全+落盘校验+已存在拒绝+慢档；双宿主真机 PASS',
  extractTemplate: '模板层大卡落地：复制底稿保样式+正文段转占位+标题保序+maxSlots 上限；双宿主真机 PASS',
  renderTemplatePpt: '模板三格式卡落地：ppt 渲染路径（组合递归 GroupItems + TextRange.Replace 标量两参）核心交 5.1 子进程（wps-render-ppt51.ps1，自包含 CoCreate——pwsh7 daemon STA 会话态绑定不稳实锤）；导图对照验收像素实证',
  renderTemplateXlsx: '模板三格式卡落地：xlsx 渲染路径（UsedRange 多区域 Value2 归一 + SAFEARRAY GetValue + Range.Replace 标量两参 + ReadOnly 打开）；读回验收真机绿',
  setFontColor: 'R1：空文本跳过+warning（deckkit table 收尾上色的兜底，不再中止 build_deck）',
};

const quirkPending = {};
const notes = {
  convertToPDF: 'app 参数已落地（excel/word/ppt/auto）',
  sortRange: '参数口径：column=数字/keyColumn=列地址；Office Range.Sort 真机直通',
  getOpenDocuments: 'Get-WpsWordFromRot 已接 host 感知',
  getOpenWorkbooks: '同 getOpenDocuments',
  getOpenPresentations: '同 getOpenDocuments',
  mergeCells: 'common 组共享 case（三面同构）',
  save: '对象模型同构；DisplayAlerts 差异归快照层联动设计',
  saveAs: '格式枚举待对账注记；DisplayAlerts 护栏已落地（抽验真机绿）',
  convertFormat: '格式枚举对账（saveAs 同源）',
  ping: '无 COM，双 host 直通',
  wireCheck: '无 COM，双 host 直通',
  __sleep: '内部探针，无 COM',
  __sleepRead: '内部探针，无 COM',
};

const wpsOnly = {
  showWps: 'WPS 窗口置前 + App Paths 拉起；office host 守卫拦截（真机绿）',
  __probeCaps: '能力探针；office host 守卫拦截',
  __probeCaptionSeq: '能力探针；office host 守卫拦截',
  __probeClose: '能力探针；office host 守卫拦截',
  __probeDoc: '能力探针；office host 守卫拦截',
  __probeStyleAdd: '能力探针；office host 守卫拦截',
  __probeWordCaps: '能力探针；office host 守卫拦截',
  __probeExcelCaps: '能力探针；office host 守卫拦截',
  __probeDialogs: '模态探测；office host 守卫拦截',
  __probeParams: '参数形状探针；office host 守卫拦截',
  __probeDeckkit: 'deckkit 探针；office host 守卫拦截',
};

const slidesCases = new Set(['getActivePresentation','deleteSlide','duplicateSlide','moveSlide','getSlideInfo','setSlideLayout','getSlideNotes','setSlideNotes','addTextBox','deleteTextBox','getTextBoxes','setTextBoxText','setTextBoxStyle','setSlideTitle','getSlideTitle','setSlideSubtitle','setSlideContent','addShape','deleteShape','batchUpsertShapes','getShapes','setShapeStyle','setFontColor','setShapeFill','setShapeText','setShapePosition','insertPptImage','deletePptImage','replacePptImage','setImageStyle','exportSlideAsImage','insertPptTable','setPptTableCell','getPptTableCell','setPptTableStyle','setPptTableCellStyle','setPptTableRowStyle','setShapeShadow','setBackgroundGradient','setShapeGradient','setShapeBorder','setShapeTransparency','setShapeRoundness','setShapeFullStyle','alignShapes','distributeShapes','groupShapes','duplicateShape','setShapeZOrder','insertPptChart','addAnimation','setSlideTransition','setSlideBackground','addConnector','addArrow','applyColorScheme','autoBeautifySlide','createKpiCards','createStyledTable','addTitleDecoration','addPageIndicator','beautifyAllSlides','createProgressBar','createGauge','createMiniCharts','createDonutChart','autoLayout','smartDistribute','createGrid','addAnimationPreset','addEmphasisAnimation','createFlowChart','createOrgChart','createTimeline','set3DRotation','set3DDepth','set3DMaterial','create3DText','setPptChartData','setPptChartStyle','removeAnimation','getAnimations','setAnimationOrder','removeSlideTransition','applyTransitionToAll','setBackgroundColor','setBackgroundImage','addPptHyperlink','removePptHyperlink','findPptText','replacePptText','unifyFont','beautifySlide','__probeDeckkit']);

function tagOf(c) {
  if (wpsOnly[c]) return { tag: 'WPS 专用', note: wpsOnly[c], v: 'static' };
  if (quirkPending[c]) return { tag: '需 quirk 分支(未落地)', note: quirkPending[c], v: 'static' };
  if (quirkLanded[c]) {
    const lf = live.has(c);
    return { tag: '需 quirk 分支(已落地)', note: quirkLanded[c] + (lf ? '；真机绿 ✓' : ''), v: lf ? 'live' : 'static' };
  }
  if (slidesCases.has(c)) {
    const lf = live.has(c);
    return { tag: '需 quirk 分支(已落地)', note: quirkLanded.__slides + (lf ? '；试点真机绿 ✓' : ''), v: lf ? 'live' : 'static' };
  }
  const lf = live.has(c);
  if (lf) return { tag: '直通', note: notes[c] ? notes[c] + '；真机绿 ✓' : '真机绿 ✓', v: 'live' };
  if (notes[c]) return { tag: '直通', note: notes[c] + '；static', v: 'static' };
  return { tag: '直通', note: '', v: 'static' };
}

const byApp = { word: [], excel: [], ppt: [], common: [] };
for (const c of cases.sort()) byApp[routes[c]].push(c);

const stat = { 直通: 0, '需 quirk 分支(已落地)': 0, '需 quirk 分支(未落地)': 0, 'WPS 专用': 0 };
let liveCount = 0;
const rows = {};
for (const app of Object.keys(byApp)) {
  rows[app] = byApp[app].map((c) => {
    const t = tagOf(c);
    stat[t.tag]++;
    if (t.v === 'live') liveCount++;
    return `| \`${c}\` | ${t.tag} | ${t.v} | ${t.note || '—'} |`;
  });
}

const md = [];
md.push('# Office 兼容度对账表（一期+二期+收口+模板层+三格式 · 全量施工图）');
md.push('');
md.push('> 分母 = ps1 全部 case **317**（word 75 / excel 105 / ppt 113 / common 24——模板层 renderTemplate/extractTemplate/renderTemplatePpt/renderTemplateXlsx 四 case 入 common，与 gen-action-routes 口径一致）。');
md.push('> 标注三态：**直通**｜**需 quirk 分支(已落地/未落地)**（注明差异）｜**WPS 专用(Office 不做)**。');
md.push('> 验证状态两档：**live**（真机实测）/ **static**（静态对账依据）。');
md.push('> 一期：试点 32 case（36/36）；二期：抽验 44 case（60/61）；收口：双场景 42/42 + findInSheet 9/9 + B2 六项 14/14 + 孤儿桥 10/10；三格式：ppt/xlsx 渲染+导图像素+读回全绿（WPS 宿主主验证）。');
md.push('');
md.push('## 一、总统计');
md.push('');
md.push('| 标注 | case 数 | 占比 |');
md.push('|---|---|---|');
for (const k of Object.keys(stat)) md.push(`| ${k} | ${stat[k]} | ${(stat[k] * 100 / 317).toFixed(1)}% |`);
md.push(`| **合计** | **317** | 100% |`);
md.push('');
md.push(`live 标注 ${liveCount} 个 case。`);
md.push('');
md.push('## 二、quirk 清单（一期六项 + 二期十项 + 收口/模板层持续，全部真机证据驱动，已落地）');
md.push('');
md.push('一期：①VT_I8 拒收 ②Sheets.Item(null)×3 ③NumberFormat General ④chartType 字符串映射 ⑤Resolve-Book 记账链 ⑥Test-AppIdentity 瞬态不定态。');
md.push('二期：⑦Resolve-Sheet 化×11 ⑧HasTitle 三态→try Shapes.Title ⑨Copy $null→Missing ⑩边距单位 ⑪手柄防护 ⑫freezePanes 补窗 ⑬saveAs DisplayAlerts 护栏 ⑭convertToPDF app 参数 ⑮Resolve-Doc 记账 ⑯WPS 专用守卫。');
md.push('收口：⑰绑定一致性（缓存探活优先于 ROT）⑱create 记账健康门控 ⑲findInSheet 5.1 子进程。');
md.push('模板层/三格式：⑳renderTemplate 双轨 ㉑extractTemplate 保样式 ㉒renderTemplatePpt 5.1 子进程 ㉓renderTemplateXlsx SAFEARRAY GetValue+Range.Replace。');
md.push('');
md.push('## 三、施工清单');
md.push('');
md.push('未落地清零。存量注记：saveAs/convertFormat 的格式枚举细目对账（直通主体已真机验）。');
md.push('');
md.push('## 四、分应用对账表');
md.push('');
for (const app of ['word', 'excel', 'ppt', 'common']) {
  md.push(`### ${app}（${byApp[app].length} case）`);
  md.push('');
  md.push('| case | 标注 | 验证 | 说明 |');
  md.push('|---|---|---|---|');
  md.push(rows[app].join('\n'));
  md.push('');
}
md.push('## 五、真机证据索引');
md.push('');
md.push('- 一期试点：`%TEMP%\\office-pilot\\pilot.cjs`（36/36）。二期抽验：`%TEMP%\\office-spot\\spot.cjs`（60/61）。');
md.push('- 三格式锚：`%TEMP%\\office-probe\\trifmt-anchor.cjs`（WPS 宿主：ppt fill=5+导图 30888B+像素非空白 / xlsx fill=4+读回 / docx 回归+缺 key skipped 列全）。');
md.push('- 双宿主渲染：`%TEMP%\\office-probe\\dual-host-render.cjs`（off/on 各 render+extract 全绿）。');
md.push('- 收口：`scripts/manual/word-binding-repro.js`（钉板）/ `binding-accept.js`（双场景 42/42）/ `orphan-bridge-verify.js`（10/10）。');
md.push('');

fs.writeFileSync('C:/Users/金东权/wps-skills/docs/agents/office-compat-matrix.md', md.join('\n'), 'utf8');
console.log('对账表已更新（三格式+模板层版）');
console.log('统计: ' + JSON.stringify(stat) + ' | live=' + liveCount);
