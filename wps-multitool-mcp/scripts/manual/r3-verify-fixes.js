/**
 * 第 1 期 · 第 2 章 · 第 3 轮 · 4 个修复的独立真机复核（决策者持锁跑）
 *
 * 为什么单独写一个，而不是复用 agentX-r11-verify.js
 * --------------------------------------------------
 * 那个脚本用 `toolRegistry.getTool(n).handler(a)` 直调 handler —— 绕过了 callTool，
 * 别名解析 / 参数映射 / 必填校验全都不经过（铁律 8：绕过 callTool 判"通过"是假绿）。
 * 本脚本一律走 `toolRegistry.callTool({id,name,arguments})`。
 *
 * 判据口径（每条都断言【具体值】，不断言 success 或回包长度）
 * ----------------------------------------------------------
 * F1 auto_sum       : A11 的 formula 恰为 =SUM(A1:A10)，值 55；改 A1 后跟着变（不是死数字）
 * F2 evaluate_formula: 不给 cell 时整表零改动；给 cell 时写【值】不写公式
 * F3 set_zoom       : 独立 PowerShell 进程直读 ActiveWindow.Zoom（不经 MCP 桥）
 * F4 insert_section_break : 节数 +1；evenPage 不退化成普通分页
 * F5 set_formula    : 不传 sheet 不再抛 Sheets.Item($null)（本轮连带修的旧 bug）
 * F6 insertPageBreak: page/column/section/sectionContinuous 四种映射不回归（ps1 裸调，工具层不暴露 type）
 *
 * 临时文稿一律 save:false / saveChanges:false，只关自己创建的（打开列表差集风格：只用自己拿到的名字）。
 */

const { execSync } = require('child_process');
const { randomUUID } = require('crypto');
const fs = require('fs');
const os = require('os');
const path = require('path');

const STAMP = Date.now();
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-r3-${STAMP}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_IDLE_MS = '120000';
process.env.WPS_DAEMON_TIMEOUT = '30000';

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));

new WpsMcpServer();

const call = (name, args = {}) => toolRegistry.callTool({ id: randomUUID(), name, arguments: args });
const txt = (r) =>
  (r && r.content && r.content.map((c) => c.text).join('\n')) || (r && r.error) || '';

const results = [];
function check(name, pass, detail) {
  results.push({ name, pass });
  console.log(`  ${pass ? 'PASS' : 'FAIL'} ${name}`);
  if (detail) console.log(`        ${String(detail).replace(/\n/g, ' | ').slice(0, 200)}`);
}

function runPs(lines) {
  const tmp = path.join(os.tmpdir(), `r3-ps-${Date.now()}-${Math.random().toString(36).slice(2)}.ps1`);
  // 必须带 BOM：PowerShell 5.1 读【无 BOM】的 UTF-8 脚本会把中文（工作簿19 这类名字）
  // 解码成乱码，于是名字比对永远不匹配、脚本静默什么都不做。
  // 这个坑 HANDOFF 里写过（"改名用 cmd ren"那一条），我在这里又踩了一次：
  // 兜底关闭"执行了但没关掉"，没有任何报错。
  fs.writeFileSync(tmp, '\uFEFF' + lines.join('\n'), 'utf8');
  try {
    return execSync(`powershell -NoProfile -File "${tmp}"`, { encoding: 'utf8', timeout: 25000 }).trim();
  } finally {
    try { fs.unlinkSync(tmp); } catch { /* */ }
  }
}

/**
 * 认领自己创建的文稿：拿【打开列表差集】，不猜名字。
 *
 * 走过的两个坑（都真踩了）：
 *   v1 从 create_* 的返回文本里抓 "name": —— 返回里没有这个名字字段 → 一个都没关，
 *      现场留了 工作簿12 / 文字文稿14。
 *   v2 改用 wps_excel_get_open_workbooks 的文本再正则 —— 那个 handler 把 data 格式化成
 *      中文列表，不是原始 JSON → 仍然抓不到，还打印出"残留 0"这种【假绿】。
 * 现在改成独立 PowerShell 读 COM 的 Workbooks/Documents 集合（与读 Zoom 同一路数），
 * 空集合明确回 {"names":[]}，读不到（WPS 没开）也回空数组，绝不会"看起来干净"。
 */
function openNames(app, collection) {
  const script = [
    'try {',
    `  $a = [Runtime.Interopservices.Marshal]::GetActiveObject('${app}.Application')`,
    `  @{ names = @($a.${collection} | ForEach-Object { $_.Name }) } | ConvertTo-Json -Compress`,
    "} catch { '{\"names\":[]}' }",
  ];
  try {
    const j = JSON.parse(runPs(script));
    return Array.isArray(j.names) ? j.names : [];
  } catch {
    return [];
  }
}

/** 独立读回：另起 PowerShell 进程走 COM，绝不经过 MCP 桥 */
function readbackExcelZoom() {
  const ps = [
    "$app = [Runtime.Interopservices.Marshal]::GetActiveObject('Ket.Application')",
    '$w = $app.ActiveWindow',
    "[PSCustomObject]@{ zoom = [int]$w.Zoom; book = $app.ActiveWorkbook.Name } | ConvertTo-Json -Compress",
  ].join('\n');
  const tmp = path.join(os.tmpdir(), `r3-zoom-${Date.now()}.ps1`);
  fs.writeFileSync(tmp, ps, 'utf8');
  try {
    const out = execSync(`powershell -NoProfile -File "${tmp}"`, { encoding: 'utf8', timeout: 20000 });
    return JSON.parse(out.trim());
  } finally {
    try { fs.unlinkSync(tmp); } catch { /* */ }
  }
}

const num = (s) => {
  const m = String(s).match(/-?\d+(\.\d+)?/);
  return m ? Number(m[0]) : NaN;
};

(async () => {
  const beforeBooks = new Set(openNames('Ket', 'Workbooks'));
  const beforeDocs = new Set(openNames('Kwps', 'Documents'));
  console.log(`开跑前：Excel 已开 ${beforeBooks.size} 个 [${[...beforeBooks].join(', ')}]，Word 已开 ${beforeDocs.size} 个 [${[...beforeDocs].join(', ')}]`);

  let sheet = 'Sheet1';

  try {
    // ────────────────────────── Excel ──────────────────────────
    console.log('\n=== Excel 靶子 ===');
    const wb = await call('wps_excel_create_workbook', {});
    check('建工作簿', wb.success === true, txt(wb).split('\n')[0]);

    const sl = await call('wps_excel_get_sheet_list', {});
    const sheetName = (txt(sl).match(/"name"\s*:\s*"([^"]*)"/) || [])[1];
    if (sheetName) sheet = sheetName;
    console.log(`  当前工作表: ${sheet}`);

    // 写 1..10 到 A1:A10
    for (let i = 1; i <= 10; i++) {
      const r = await call('wps_excel_set_cell_value', { sheet, row: i, col: 1, value: String(i) });
      if (!r.success) { check(`写 A${i}`, false, txt(r)); break; }
    }
    const a11Before = await call('wps_excel_get_cell_value', { sheet, row: 11, col: 1 });
    check('基线：A11 为空', /null|""|undefined/.test(txt(a11Before)) || num(txt(a11Before)) === 0, txt(a11Before));

    // ── F1 auto_sum ──
    console.log('\n=== F1 wps_excel_auto_sum ===');
    const f1 = await call('wps_excel_auto_sum', { range: 'A1:A10', targetCell: 'A11' });
    check('F1 调用成功（不再是 Unknown action）', f1.success === true && !/Unknown action/.test(txt(f1)), txt(f1).split('\n')[0]);

    const f1f = await call('wps_excel_get_formula', { sheet, cell: 'A11' });
    check('F1 A11 的公式恰为 =SUM(A1:A10)', /=\s*SUM\(A1:A10\)/i.test(txt(f1f)), txt(f1f));

    const f1v = await call('wps_excel_get_cell_value', { sheet, row: 11, col: 1 });
    check('F1 A11 的值 = 55', num(txt(f1v)) === 55, txt(f1v));

    // 改 A1 → 公式结果必须跟着变（证明写的是公式不是死数字）
    // 期望值算式：1..10 的和是 55；把 A1 从 1 改成 100 → 55 - 1 + 100 = 154。
    // 【上一版这里写成 145，是我算错了，不是工具错】——判据错会把人骗去回滚正确的实现。
    await call('wps_excel_set_cell_value', { sheet, row: 1, col: 1, value: '100' });
    const f1v2 = await call('wps_excel_get_cell_value', { sheet, row: 11, col: 1 });
    check('F1 改 A1=100 后 A11 = 154（是活公式，不是死数字）', num(txt(f1v2)) === 154, txt(f1v2));

    // ── F2 evaluate_formula ──
    console.log('\n=== F2 wps_excel_evaluate_formula ===');
    const e1 = await call('wps_excel_evaluate_formula', { formula: '=SUM(A1:A10)' });
    check('F2a 不给 cell：计算成功且结果 = 154', e1.success === true && num(txt(e1)) === 154, txt(e1).split('\n')[0]);
    const b1 = await call('wps_excel_get_cell_value', { sheet, row: 1, col: 2 });
    check('F2a 不给 cell：B1 仍为空（只算不写）', /null|""|undefined/.test(txt(b1)) || num(txt(b1)) === 0, txt(b1));

    const e2 = await call('wps_excel_evaluate_formula', { formula: '=1+2', cell: 'B1' });
    check('F2b 给 cell：调用成功', e2.success === true, txt(e2).split('\n')[0]);
    const b1v = await call('wps_excel_get_cell_value', { sheet, row: 1, col: 2 });
    check('F2b B1 的值 = 3', num(txt(b1v)) === 3, txt(b1v));
    const b1f = await call('wps_excel_get_formula', { sheet, cell: 'B1' });
    check('F2b B1 没有公式（写的是值不是公式）', !/=/.test(txt(b1f)) || /null|空/.test(txt(b1f)), txt(b1f));

    // ── F2b 诊断：setRangeData 到底是"数字不行"还是"所有类型都不行" ──
    // 这条不是判据，是给实现者的定位证据：同一个动作，字符串能写、数字不能写 → 修 PS 侧类型处理。
    console.log('\n=== F2b 诊断 setRangeData 的类型处理 ===');
    const dStr = await call('wps_execute_method', {
      method: 'setRangeData', appType: 'excel', params: { range: 'D1', data: [['hello']] },
    });
    check('诊断：setRangeData 写【字符串】成功', dStr.success === true, txt(dStr).replace(/\n/g, ' ').slice(0, 140));
    const dNum = await call('wps_execute_method', {
      method: 'setRangeData', appType: 'excel', params: { range: 'D2', data: [[3]] },
    });
    check('诊断：setRangeData 写【数字】成功', dNum.success === true, txt(dNum).replace(/\n/g, ' ').slice(0, 140));

    // ── F5 set_formula 不传 sheet ──
    console.log('\n=== F5 wps_excel_set_formula（连带修的 sheet 解析）===');
    const s1 = await call('wps_excel_set_formula', { range: 'C1', formula: '=2+2' });
    check('F5a 不传 sheet 成功（改前抛 Sheets.Item($null)）', s1.success === true && !/Sheets.Item|Item\(\$null\)|Exception/i.test(txt(s1)), txt(s1).split('\n')[0]);
    const c1v = await call('wps_excel_get_cell_value', { sheet, row: 1, col: 3 });
    check('F5a C1 = 4', num(txt(c1v)) === 4, txt(c1v));

    const s2 = await call('wps_excel_set_formula', { range: 'C2', formula: '=3+4', sheet });
    const c2v = await call('wps_excel_get_cell_value', { sheet, row: 2, col: 3 });
    check(`F5b 显式 sheet="${sheet}" 成功且 C2 = 7`, s2.success === true && num(txt(c2v)) === 7, txt(c2v));

    // ── F3 set_zoom ──
    console.log('\n=== F3 wps_excel_set_zoom ===');
    const z1 = await call('wps_excel_set_zoom', { percent: 150 });
    check('F3a 调用成功且回显 150', z1.success === true && /150/.test(txt(z1)), txt(z1).split('\n')[0]);
    let rb = null;
    try { rb = readbackExcelZoom(); } catch (e) { rb = { error: String(e.message).slice(0, 80) }; }
    check('F3a 独立 PowerPoint 进程直读 Zoom = 150', rb && rb.zoom === 150, JSON.stringify(rb));

    const z2 = await call('wps_excel_set_zoom', { percent: 5 });
    check('F3b percent=5 → 明确报错（不静默夹取）', z2.success === false && /10-400|10~400/.test(txt(z2)), txt(z2).split('\n')[0]);

    const z3 = await call('wps_execute_method', { method: 'setZoom', appType: 'excel', params: { percent: 5 } });
    check('F3c 裸调 ps1 setZoom(5) → PS 侧护栏也报错', z3.success === false && /10-400/.test(txt(z3)), txt(z3).split('\n')[0]);

    // ── F7 wps_excel_write_range：同一个 setRangeData 底层，数字bug 的影响面 ──
    // description 里的示例是 [["姓名","年龄"],["张三",25],["李四",30]] —— 示例本身带数字。
    console.log('\n=== F7 wps_excel_write_range（同样走 setRangeData）===');
    const w1 = await call('wps_excel_write_range', { range: 'E1:F2', data: [['姓名', '年龄'], ['张三', 25]] });
    check('F7a 按 description 的示例（含数字）写入 → 成功', w1.success === true, txt(w1).split('\n')[0]);

    const w2 = await call('wps_excel_write_range', { range: 'H1:I2', data: [['姓名', '年龄'], ['张三', '25']] });
    check('F7b 同类数据全字符串 → 成功', w2.success === true, txt(w2).split('\n')[0]);

    // 关键判据：字符串写进去的 "25" 是【文本】还是【数值】——用 SUM 能否算进去来分。
    // 文本 "25" → SUM = 0；数值 25 → SUM = 25。这条比"读回值是不是 25"硬得多。
    await call('wps_excel_set_formula', { range: 'J1', formula: '=SUM(I2)', sheet });
    const j1 = await call('wps_excel_get_cell_value', { sheet, row: 1, col: 10 });
    check('F7c 全字符串写入后 SUM(I2) = 25（写进去的是数值，不是文本）', num(txt(j1)) === 25, txt(j1));

    // ────────────────────────── Word ──────────────────────────
    console.log('\n=== Word 靶子 ===');
    const doc = await call('wps_word_create_document', {});
    check('建文档', doc.success === true, txt(doc).split('\n')[0]);
    await call('wps_word_insert_text', { text: '第一段：用于分节与分页复核。\n第二段：继续写一些内容。\n第三段：再来一行。' });

    const st0 = await call('wps_word_get_document_stats', {});
    const sec0 = num(txt(st0).match(/节数[^\d-]*(-?\d+)/) ? txt(st0).match(/节数[^\d-]*(-?\d+)/)[1] : 'NaN');
    check('基线：读到节数', Number.isFinite(sec0), txt(st0).replace(/\n/g, ' ').slice(0, 120));

    // ── F4 insert_section_break ──
    console.log('\n=== F4 wps_word_insert_section_break ===');
    const sb1 = await call('wps_word_insert_section_break', { breakType: 'continuous' });
    check('F4a continuous 调用成功（不再是 Unknown action）', sb1.success === true && !/Unknown action/.test(txt(sb1)), txt(sb1).split('\n')[0]);
    const st1 = await call('wps_word_get_document_stats', {});
    const sec1 = num((txt(st1).match(/节数[^\d-]*(-?\d+)/) || [])[1]);
    check('F4a continuous 后节数 +1', sec1 === sec0 + 1, `${sec0} → ${sec1}`);

    const sb2 = await call('wps_word_insert_section_break', { breakType: 'evenPage' });
    const st2 = await call('wps_word_get_document_stats', {});
    const sec2 = num((txt(st2).match(/节数[^\d-]*(-?\d+)/) || [])[1]);
    check('F4b evenPage 后节数 +1（没退化成普通分页）', sb2.success === true && sec2 === sec1 + 1, `${sec1} → ${sec2}`);

    const sb3 = await call('wps_word_insert_section_break', { breakType: 'nextPage' });
    const st3 = await call('wps_word_get_document_stats', {});
    const sec3 = num((txt(st3).match(/节数[^\d-]*(-?\d+)/) || [])[1]);
    check('F4c nextPage（映射成 section，不是 page）后节数 +1', sb3.success === true && sec3 === sec2 + 1, `${sec2} → ${sec3}`);

    // ── F6 insertPageBreak 四种映射不回归（ps1 裸调，工具层不暴露 type）──
    console.log('\n=== F6 insertPageBreak 的四种映射（ps1 裸调）===');
    for (const [type, code] of [['page', 7], ['column', 8], ['section', 2], ['sectionContinuous', 3]]) {
      const r = await call('wps_execute_method', { method: 'insertPageBreak', appType: 'word', params: { type } });
      const got = num((txt(r).match(/"code"\s*:\s*(-?\d+)/) || [])[1]);
      check(`F6 type=${type} → 断点码 ${code}`, r.success === true && got === code, txt(r).replace(/\n/g, ' ').slice(0, 120));
    }
    // 工具层：默认分页仍可用，且不引入节（分页符不是分节符）
    const stBefore = await call('wps_word_get_document_stats', {});
    const secBefore = num((txt(stBefore).match(/节数[^\d-]*(-?\d+)/) || [])[1]);
    const pb = await call('wps_word_insert_page_break', {});
    const stAfter = await call('wps_word_get_document_stats', {});
    const secAfter = num((txt(stAfter).match(/节数[^\d-]*(-?\d+)/) || [])[1]);
    check('F6 wps_word_insert_page_break 默认仍能用，且不改变节数', pb.success === true && secAfter === secBefore, `节数 ${secBefore} → ${secAfter}`);
  } catch (e) {
    check('运行异常', false, e && e.stack ? e.stack.split('\n')[0] + ' @ ' + (e.stack.split('\n')[1] || '') : String(e));
  } finally {
    // 清理：用【打开列表差集】认领，不猜名字；残留数也由独立进程读，不用工具输出（那会假绿）。
    console.log('\n=== 清理（只关自己创建的、不保存）===');
    const mineBooks = openNames('Ket', 'Workbooks').filter((n) => !beforeBooks.has(n));
    const stubborn = [];
    for (const n of mineBooks) {
      try {
        const r = await call('wps_excel_close_workbook', { name: n, saveChanges: false });
        console.log(`  关工作簿 ${n}: ${r.success ? 'ok' : txt(r).slice(0, 110)}`);
        // 实测：从未保存过的工作簿会被这条护栏拒绝（Word 侧同名调用却能关）。
        // 这里不判它失败，记下来交给 STATE 的 B2；兜底用独立进程 Close($false)（实测不弹框）。
        if (!r.success) stubborn.push(n);
      } catch (e) { console.log(`  关工作簿 ${n} 抛错: ${e.message}`); stubborn.push(n); }
    }
    if (stubborn.length) {
      const list = stubborn.map((n) => `'${String(n).replace(/'/g, "''")}'`).join(',');
      const script = [
        'try {',
        "  $a = [Runtime.Interopservices.Marshal]::GetActiveObject('Ket.Application')",
        `  $targets = @(${list})`,
        '  foreach ($b in @($a.Workbooks)) { if ($targets -contains $b.Name) { $b.Close($false); "  兜底关闭: $($b.Name)" } }',
        '} catch { "  兜底关闭失败: $($_.Exception.Message)" }',
      ];
      console.log(runPs(script));
    }
    const mineDocs = openNames('Kwps', 'Documents').filter((n) => !beforeDocs.has(n));
    for (const n of mineDocs) {
      try {
        const r = await call('wps_word_close_document', { name: n, save: false });
        console.log(`  关文档 ${n}: ${r.success ? 'ok' : txt(r).slice(0, 80)}`);
      } catch (e) { console.log(`  关文档 ${n} 抛错: ${e.message}`); }
    }
    const leftBooks = openNames('Ket', 'Workbooks');
    const leftDocs = openNames('Kwps', 'Documents');
    console.log(`  残留（独立进程实测）：Excel ${leftBooks.length} [${leftBooks.join(', ')}] / Word ${leftDocs.length} [${leftDocs.join(', ')}]`);
    check('收尾现场干净（Excel 0 / Word 0）', leftBooks.length === 0 && leftDocs.length === 0, `Excel=${leftBooks.length} Word=${leftDocs.length}`);
  }

  const pass = results.filter((r) => r.pass).length;
  console.log(`\n══════ 汇总：${pass}/${results.length} 通过 ══════`);
  const failed = results.filter((r) => !r.pass);
  if (failed.length) {
    console.log('未通过项：');
    for (const f of failed) console.log(`  - ${f.name}`);
  }
  process.exit(0);
})();
