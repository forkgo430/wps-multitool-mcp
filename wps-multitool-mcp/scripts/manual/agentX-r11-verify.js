// 第 1 期第 2 章第 1 轮 · 验证者 独立复核真机探针
// 第 1 条：自己写读回探针验证两个 Word 工具真修好（读回走独立 PowerShell 进程，不走 MCP）
// 第 7 条：确认 4 个调了不存在动作的破损工具 + acceptAll/rejectAll 只是 handleRevisions 的 mode

const { execSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const PIPE = `\\\\.\\pipe\\wps-multitool-mcp-agentX-r11-${Date.now()}`;
process.env.WPS_BRIDGE_PIPE = PIPE;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_IDLE_MS = '120000';
process.env.WPS_DAEMON_TIMEOUT = '30000';

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);
const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => (r && r.content && r.content[0] && r.content[0].text) || (r && r.error) || '';

const results = [];
const check = (name, pass, detail) => {
  results.push({ name, pass });
  console.log(`${pass ? '  ✓' : '  ✗'} ${name} — ${detail}`);
};

// 独立读回：把 PowerShell 脚本写到临时 .ps1，单独起进程跑 GetActiveObject，绝不走 MCP 桥
function readbackWord() {
  const ps = [
    '$app = [Runtime.Interopservices.Marshal]::GetActiveObject(\'Kwps.Application\')',
    '$d = $app.ActiveDocument',
    '$r = $d.Content',
    '[PSCustomObject]@{',
    '  name = $d.Name',
    '  color = [int]$r.Font.Color',
    '  lineSpacing = [double]$r.ParagraphFormat.LineSpacing',
    '  lineRule = [int]$r.ParagraphFormat.LineSpacingRule',
    '} | ConvertTo-Json -Compress',
  ].join('\n');
  const tmp = path.join(os.tmpdir(), `r11-rb-${Date.now()}.ps1`);
  fs.writeFileSync(tmp, ps, 'utf8');
  try {
    const out = execSync(`powershell -NoProfile -File "${tmp}"`, { encoding: 'utf8', timeout: 20000 });
    return JSON.parse(out.trim());
  } finally {
    try { fs.unlinkSync(tmp); } catch { /* */ }
  }
}

(async () => {
  const mine = [];
  try {
    console.log('=== 第 1 条：两个 Word 工具是否真修好（自建读回，独立于 MCP）===');
    const cr = await call('wps_word_create_document', {});
    if (!cr.success) { check('建文档', false, txt(cr).slice(0, 60)); return; }
    const name = (txt(cr).match(/"name":"([^"]*)"/) || [])[1];
    if (name) mine.push(name);
    await call('wps_word_insert_text', { text: 'R11读回测试文本' });

    const before = readbackWord();
    console.log('  基线: ' + JSON.stringify(before));

    const c1 = await call('wps_word_set_text_color', { color: '#FF0000', range: 'all' });
    const after1 = readbackWord();
    check('set_text_color 返回成功(动作名已改对)', c1.success === true && !/Unknown action/.test(txt(c1)), txt(c1).split('\n')[0].slice(0, 50));
    check('set_text_color 读回颜色真变了', after1.color !== before.color, `color ${before.color} → ${after1.color}`);

    const c2 = await call('wps_word_set_line_spacing', { lineSpacing: 2, range: 'all' });
    const after2 = readbackWord();
    check('set_line_spacing 返回成功(动作名已改对)', c2.success === true && !/Unknown action/.test(txt(c2)), txt(c2).split('\n')[0].slice(0, 50));
    check('set_line_spacing 读回行距真变了', after2.lineSpacing !== before.lineSpacing, `ls ${before.lineSpacing} → ${after2.lineSpacing} (rule=${after2.lineRule})`);

    console.log('\n=== 第 7 条：调了不存在动作的破损工具 + acceptAll/rejectAll ===');
    // Excel 环境
    const xb = await call('wps_excel_create_workbook', {});
    const xname = (txt(xb).match(/"name":"([^"]*)"/) || [])[1];
    if (xname) mine.push(xname);
    // 第 1 期第 2 章第 2 轮：这组断言被实现者改过 —— 修完 4 个工具后，现实变了：
    //   autoSum         仍然不存在（正解是 TS 改调 setFormula），断言保留；
    //   evaluateFormula 现在【存在】（本轮新补 case，只算不写），断言必须反过来；
    //   setZoom         现在【存在】（本轮新补 case，ActiveWindow.Zoom），断言必须反过来。
    // 不反过来就会对着"已修好的东西"报红，那种红会把人骗去回滚正确的修复。
    for (const a of ['autoSum']) {
      const r = await call('wps_execute_method', { method: a, appType: 'excel', params: {} });
      check(`动作 "${a}" 在 ps1 里不存在 → 回 Unknown action`, /Unknown action/.test(txt(r)), txt(r).split('\n')[0].slice(0, 60));
    }
    for (const a of ['evaluateFormula', 'setZoom']) {
      const r = await call('wps_execute_method', { method: a, appType: 'excel', params: {} });
      check(
        `动作 "${a}" 本轮已补 case → 不再回 Unknown action（空参应得到"先改参数"式报错）`,
        !/Unknown action/.test(txt(r)) && /先改参数|必须/.test(txt(r)),
        txt(r).split('\n')[0].slice(0, 60)
      );
    }
    // Word 环境
    for (const a of ['insertSectionBreak', 'acceptAll', 'rejectAll']) {
      const r = await call('wps_execute_method', { method: a, appType: 'word', params: {} });
      check(`动作 "${a}" 在 ps1 顶层 switch 里不存在 → 回 Unknown action（证明它不是可分发动作）`, /Unknown action/.test(txt(r)), txt(r).split('\n')[0].slice(0, 60));
    }

    // 对照组：handleRevisions(mode:'acceptAll') 应能走通（acceptAll 是它的 mode，不是动作）
    const hr = await call('wps_word_handle_revisions', { mode: 'acceptAll' });
    check('handleRevisions({mode:"acceptAll"}) 是正确用法（mode 不是动作）', hr.success === true || /没有.*修订|0/.test(txt(hr)), txt(hr).split('\n')[0].slice(0, 60));
  } catch (e) {
    check('运行异常', false, e && e.message ? e.message : String(e));
  } finally {
    for (const n of mine) {
      try {
        if (n.startsWith('文字')) await call('wps_word_close_document', { name: n, save: false });
        else await call('wps_excel_close_workbook', { name: n, saveChanges: false });
      } catch { /* 已关 */ }
    }
  }
  const pass = results.filter((r) => r.pass).length;
  console.log(`\n汇总：${pass}/${results.length} 通过`);
  process.exit(0);
})();
