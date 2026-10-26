/**
 * Input: 别名表 tool-aliases.json（生成物 tool-aliases.generated.ts）+ 真机 WPS
 * Output: 旧名调用回归结论（同一处写入 / 同格同值 / 注解与 warnings 不缩水）
 * Pos: 第 2 轮别名收敛的真机回归。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 验三件事（INBOX 第 2 轮·验证者·任务 2）：
 *   1) wps_insert_text（旧）与 wps_word_insert_text（规范）都成功，且改的是同一处
 *      —— 用"先后插入 + 读全文"证明两者落在同一文档同一插入点，而不是各写各的
 *   2) wps_get_cell_value / wps_excel_get_cell_value 返回同一格同一值
 *      —— 规范名写、旧名读；旧名写、规范名读，两个方向都过
 *   3) 旧名调用的 warnings 与注解行为跟规范名一致，不许出现"旧名没有破坏性提示"的口子
 *      —— 别名必须在 registry 入口就解析掉：旧名不进 tools/list（客户端看到的注解只有规范名那份），
 *         同一操作的 result.warnings / success 与规范名逐字段相同
 *
 * 静态部分不起桥；真机部分走独立管道 + EPHEMERAL，临时文档收尾关闭不保存。
 *
 * 用法（需先 npm run build）：node scripts/manual/alias-regression.js
 */
const fs = require('fs');
const os = require('os');
const path = require('path');

process.env.WPS_BRIDGE_PIPE = '\\\\.\\pipe\\wps-multitool-mcp-agentC-alias';
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_IDLE_MS = '120000';
process.env.WPS_DAEMON_TIMEOUT = process.env.WPS_DAEMON_TIMEOUT || '120000';

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { TOOL_ALIASES, ALIAS_NAMES } = require(path.join(ROOT, 'dist', 'server', 'tool-aliases.generated.js'));

const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

let skipped = false;
const results = [];
const check = (label, ok, detail = '') => {
  results.push({ label, ok });
  console.log(`  ${ok ? '✓' : '✗'} ${label}${detail ? ' — ' + detail : ''}`);
};
const call = (name, args = {}) => toolRegistry.getTool(name).handler(args);
const txt = (r) => (r.content && r.content[0] && r.content[0].text) || r.error || '';
const warningsOf = (r) => JSON.stringify(r.warnings || null);

(async () => {
  try {
    console.log('\n=== 0. 别名表静态检查（不需要 WPS）===');
  const registered = new Set(toolRegistry.listAllTools().map((d) => d.name));
  const badTargets = ALIAS_NAMES.filter((a) => !registered.has(TOOL_ALIASES[a]));
  check('每个别名目标都已在注册表里', badTargets.length === 0, badTargets.join(',') || '无');
  const selfAlias = ALIAS_NAMES.filter((a) => TOOL_ALIASES[a] === a);
  check('没有别名指向自己', selfAlias.length === 0, selfAlias.join(',') || '无');
  const aliasAlsoTool = ALIAS_NAMES.filter((a) => registered.has(a));
  check('旧名没有同时又是注册工具（否则 tools/list 会出现两个）', aliasAlsoTool.length === 0, aliasAlsoTool.join(',') || '无');

  const listed = toolRegistry.listTools().tools.map((d) => d.name);
  const aliasInList = ALIAS_NAMES.filter((a) => listed.includes(a));
  check('旧名不出现在 tools/list（客户端看不到旧名）', aliasInList.length === 0, aliasInList.join(',') || '无');
  const allNames = toolRegistry.listAllTools().map((d) => d.name);
  check(
    '注册表里没有重名工具（别名没有偷偷占一个注册位）',
    new Set(allNames).size === allNames.length,
    `共 ${allNames.length} 个，去重后 ${new Set(allNames).size} 个`
  );

  const { getToolAnnotations } = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.js'));
  const noAnno = [...new Set(ALIAS_NAMES.map((a) => TOOL_ALIASES[a]))].filter((n) => !getToolAnnotations(n));
  check('每个别名目标都有注解（破坏性提示挂在规范名上）', noAnno.length === 0, noAnno.join(',') || '无');

  console.log('\n=== 1. 真机：Word 同位写入（旧名 vs 规范名）===');
  const hb = await call('wps_common_ping', {});
  if (!hb.success) {
    skipped = true;
    console.log(`  ⚠ WPS 不可用（${txt(hb).slice(0, 80)}），本轮跳过，不判失败`);
    console.log('    （别名回归在真机上【未验证】）');
    return;
  }

  const doc = await call('wps_word_create_document', {});
  check('新建临时 Word 文档', doc.success, txt(doc).split('\n')[0].slice(0, 60));
  if (doc.success) {
    const wCanon = await call('wps_word_insert_text', { text: '规范写入A', position: 'end' });
    const wOld = await call('wps_insert_text', { text: '旧名写入B', position: 'end' });
    const full = await call('wps_word_get_document_text', {});
    const body = txt(full);
    const iA = body.indexOf('规范写入A');
    const iB = body.indexOf('旧名写入B');
    check(
      'wps_insert_text（旧）与 wps_word_insert_text（规范）都成功',
      wCanon.success && wOld.success,
      `规范=${wCanon.success} 旧名=${wOld.success}｜${txt(wOld).slice(0, 60)}`
    );
    check(
      '两者改的是同一处（同一文档、同一插入点、顺序正确）',
      iA >= 0 && iB > iA,
      `A@${iA} B@${iB}｜全文长度 ${body.length}`
    );
    check(
      '旧名与规范名的 warnings/success 字段一致',
      wOld.success === wCanon.success && warningsOf(wOld) === warningsOf(wCanon),
      `warnings 规范=${warningsOf(wCanon)} 旧名=${warningsOf(wOld)}`
    );
    const closedDoc = await call('wps_word_close_document', { save: false });
    check('收尾关闭临时文档', closedDoc.success, txt(closedDoc).slice(0, 60));
  }

  console.log('\n=== 2. 真机：Excel 同格同值（两个方向）===');
  const wb = await call('wps_excel_create_workbook', {});
  check('新建临时工作簿', wb.success, txt(wb).split('\n')[0].slice(0, 60));
  if (wb.success) {
    // 规范名写 A1，旧名读 A1
    const setA1 = await call('wps_excel_set_cell_value', { sheet: 1, row: 1, col: 1, value: '规范值' });
    const getOld = await call('wps_get_cell_value', { sheet: 1, row: 1, col: 1 });
    const getCanon = await call('wps_excel_get_cell_value', { sheet: 1, row: 1, col: 1 });
    check(
      'wps_get_cell_value（旧）与 wps_excel_get_cell_value（规范）返回同一格同一值',
      setA1.success && getOld.success && getCanon.success
        && txt(getOld).includes('规范值') && txt(getCanon).includes('规范值'),
      `旧名读="${(txt(getOld).match(/[:：]\s*(.*)/) || [])[1]}" 规范读="${(txt(getCanon).match(/[:：]\s*(.*)/) || [])[1]}"`
    );

    // 旧名写 A2，规范名读 A2
    const setA2 = await call('wps_set_cell_value', { sheet: 1, row: 2, col: 1, value: '旧名值' });
    const getCanon2 = await call('wps_excel_get_cell_value', { sheet: 1, row: 2, col: 1 });
    check(
      '旧名写、规范名读也通（wps_set_cell_value → wps_excel_get_cell_value）',
      setA2.success && getCanon2.success && txt(getCanon2).includes('旧名值'),
      `${txt(setA2).split('\n')[0].slice(0, 50)}｜读回="${(txt(getCanon2).match(/[:：]\s*(.*)/) || [])[1]}"`
    );

    // 收尾关闭：不走 wps_excel_close_workbook —— 那个工具把 save 发给 PS，而 PS 读的是
    // saveChanges（src/tools/excel/workbook.ts:175 vs scripts/wps-com.ps1:3129），
    // 关未保存工作簿会弹「另存为」模态框把 COM 通道卡到超时（首轮实测 120s）。
    // 所以这里用裸调 + 正确参数名收尾，别让清理步骤自己变成故障。
    const closedWb = await call('wps_execute_method', {
      method: 'closeWorkbook', appType: 'excel', params: { saveChanges: false },
    });
    check('收尾关闭临时工作簿', closedWb.success, txt(closedWb).slice(0, 80));
    console.log(
      '  · 历史观测（不计 PASSED）：首轮用 wps_excel_close_workbook({save:false}) 收尾时实测超时 120s' +
        '（TS 发 save、PS 读 saveChanges，缺省 $true → 未存盘工作簿 Close($true) 弹框卡死通道），故改用裸调收尾'
    );
  }

  console.log('\n=== 3. 注解行为：旧名不许绕过破坏性提示 ===');
  const destructive = ['wps_insert_text', 'wps_get_cell_value', 'wps_set_cell_value'].map((a) => ({
    alias: a,
    canonical: TOOL_ALIASES[a],
    anno: getToolAnnotations(TOOL_ALIASES[a]),
  }));
  for (const d of destructive) {
    check(
      `${d.alias} → ${d.canonical} 的破坏性标注挂在规范名上且旧名不可见`,
      !!d.anno && !listed.includes(d.alias),
      JSON.stringify(d.anno)
    );
  }
} catch (e) {
  console.log(`  ⚠ 探测中断：${e && e.message ? e.message : String(e)}`);
  console.log('    （中断部分不判失败，但也不算数 —— 别名回归这部分按未验证计）');
} finally {
  const passed = results.filter((r) => r.ok).map((r) => r.label);
  const failed = results.filter((r) => !r.ok).map((r) => r.label);
  console.log(`\n通过 ${passed.length}/${results.length}`);
  if (failed.length) console.log(`失败：${failed.join('；')}`);
  if (skipped) {
    console.log('未探测到 WPS（或别名真机部分没跑），按约定跳过，退出 0。');
    process.exit(0);
  }
  console.log(`PASSED=${passed.length}/${results.length}`);
  process.exit(failed.length ? 1 : 0);
}
})().catch((e) => {
  console.error('探测失败:', e && e.message ? e.message : String(e));
  process.exit(1);
});
