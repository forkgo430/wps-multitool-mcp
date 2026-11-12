/* 模板三格式补全卡 · 真机锚（WPS 宿主主验证 + 导图验收） */
process.env.WPS_OFFICE = 'off';   // WPS 宿主主验证（卡面红线：WPS 主验证）
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-trifmt-${Date.now()}`;
const path = require('path');
const fs = require('fs');
const { wpsClient } = require(path.join('C:/Users/金东权/wps-skills/wps-multitool-mcp', 'dist', 'client', 'wps-client.js'));
const TPL = 'C:/Users/金东权/wps-skills/wps-multitool-mcp/templates';
const TMP = 'C:/Users/金东权/AppData/Local/Temp/trifmt';
const TS = Date.now();
fs.mkdirSync(TMP, { recursive: true });
let pass = 0, fail = 0;
async function run(label, action, params, app, assert) {
  try {
    const r = await wpsClient.executeMethod(action, params || {}, app);
    if (r && r.success === true) { let n = ''; if (assert) { try { n = assert(r); } catch (e) { fail++; console.log(`  ✗ ${label} 断言: ${e.message}`); return r; } } pass++; console.log(`  ✓ ${label} ${n}`); }
    else { fail++; console.log(`  ✗ ${label}: ${(r && (r.error || JSON.stringify(r.warnings || '')) || '').slice(0, 160)}`); }
    return r;
  } catch (e) { fail++; console.log(`  ✗ ${label} EXC: ${(e.message || '').slice(0, 150)}`); return null; }
}
(async () => {
  console.log('=== 模板三格式真机锚（WPS 宿主）===');
  // PPT 渲染：{{brief_title}}/{{presenter}}/{{brief_point_N}}
  const outPpt = path.join(TMP, `rendered-ppt-${TS}.pptx`);
  const rp = await run('PPT renderTemplate（{{brief_*}} 四锚）', 'renderTemplatePpt', {
    templatePath: path.join(TPL, 'briefing.pptx'), outputPath: outPpt,
    data: { brief_title: '三格式锚验证', presenter: '执行者', brief_point_1: '要点一', brief_point_2: '要点二', brief_point_3: '要点三' },
  }, 'ppt', (r) => {
    if (!(r.data && r.data.filledCount >= 4)) throw new Error('filled<4: ' + JSON.stringify(r.data).slice(0, 100));
    if (!fs.existsSync(outPpt) || fs.statSync(outPpt).size === 0) throw new Error('未落盘');
    return `落盘 ${fs.statSync(outPpt).size}B filled=${r.data.filledCount}`;
  });
  // 导图对照验收（exportSlideAsImage 现成）+ 非空白像素比例（pwsh7 System.Drawing）
  if (rp && rp.success) {
    const png = path.join(TMP, 'slide1.png');
    await run('导图前重开产物', 'openPresentation', { path: outPpt }, 'ppt');
    const ex = await run('PPT 导图对照验收（exportSlideAsImage）', 'exportSlideAsImage', { slideIndex: 1, outputPath: png, format: 'PNG', presentationName: outPpt }, 'ppt', () => {
      if (!fs.existsSync(png) || fs.statSync(png).size === 0) throw new Error('导图未落盘');
      return `导图 ${fs.statSync(png).size}B`;
    });
    if (ex && ex.success) {
      const px = spawnSync2();
      function spawnSync2() {
        try {
          const { execFileSync } = require('child_process');
          const out = execFileSync('pwsh', ['-NoProfile', '-Command', `
            Add-Type -AssemblyName System.Drawing
            \$bmp = [System.Drawing.Bitmap]::FromFile('${png.replace(/'/g, "''")}')
            \$nonWhite = 0; \$total = 0
            for (\$x = 0; \$x -lt \$bmp.Width; \$x += 4) { for (\$y = 0; \$y -lt \$bmp.Height; \$y += 4) {
              \$p = \$bmp.GetPixel(\$x, \$y); \$total++
              if (-not (\$p.R -gt 240 -and \$p.G -gt 240 -and \$p.B -gt 240)) { \$nonWhite++ }
            } }
            \$bmp.Dispose()
            Write-Output ("PIXELS nonWhite=\$nonWhite total=\$total")
          `], { encoding: 'utf8', timeout: 60000 });
          const m = out.match(/nonWhite=(\d+) total=(\d+)/);
          if (m) {
            const ratio = parseInt(m[1], 10) / parseInt(m[2], 10);
            console.log(`  [像素] 非空白比例 ${(ratio * 100).toFixed(1)}%（>5% 即非空白页）`);
            return ratio > 0.05;
          }
          return false;
        } catch { return false; }
      }
      ok('pptx 导图像素验收（非空白比例 >5%）', spawnSync2(), '');
      await run('导图后关闭产物', 'closePresentation', {}, 'ppt');
    }
  }
  // XLSX 渲染：{{ledger_title}}/{{ledger_item_N}}
  const outX = path.join(TMP, `rendered-xlsx-${TS}.xlsx`);
  const rx = await run('XLSX renderTemplate（{{ledger_*}} 四锚）', 'renderTemplateXlsx', {
    templatePath: path.join(TPL, 'ledger.xlsx'), outputPath: outX,
    data: { ledger_title: '三格式台账锚验证', ledger_item_1: '条目甲', ledger_item_2: '条目乙', ledger_item_3: '条目丙' },
  }, 'excel', (r) => {
    if (!(r.data && r.data.filledCount >= 3)) throw new Error('filled<3: ' + JSON.stringify(r.data).slice(0, 100));
    if (!fs.existsSync(outX) || fs.statSync(outX).size === 0) throw new Error('未落盘');
    return `落盘 ${fs.statSync(outX).size}B filled=${r.data.filledCount}`;
  });
  // XLSX 读回验收（渲染值实在）
  if (rx && rx.success) {
    const renderedName = path.basename(outX);
    await run('XLSX 读回验收（openWorkbook 产物）', 'openWorkbook', { path: outX }, 'excel', (r) => 'opened=' + renderedName);
    await run('XLSX getCellValue A1', 'getCellValue', { row: 1, col: 1, workbookName: renderedName }, 'excel', (r) => {
      const v = r.data && (r.data.value ?? r.data);
      if (String(v).indexOf('三格式台账锚验证') === -1) throw new Error('A1 未替换: ' + JSON.stringify(v));
      return 'A1=' + v;
    });
  }
  // 缺 key 不静默（data 只给 1 锚 → 其余 skipped 列全）
  const outM = path.join(TMP, `rendered-missing-${TS}.pptx`);
  const rm = await run('PPT 缺 key 不静默（skipped 列全）', 'renderTemplatePpt', {
    templatePath: path.join(TPL, 'briefing.pptx'), outputPath: outM,
    data: { brief_title: '只有标题' },
  }, 'ppt', (r) => {
    if (!(r.data && r.data.skipped && r.data.skipped.length >= 3)) throw new Error('skipped<3: ' + JSON.stringify(r.data && r.data.skipped));
    return 'skipped=' + r.data.skipped.join(',');
  });
  // docx 路径不变（回归：同款调用仍走 renderTemplate）
  const outD = path.join(TMP, `rendered-docx-${TS}.docx`);
  await run('DOCX 回归（双轨不变）', 'renderTemplate', {
    templatePath: path.join(TPL, 'weekly-report.docx'), outputPath: outD,
    data: { week_summary: 'docx 回归', task_1: 'T1', task_2: 'T2', task_3: 'T3', next_plan: 'N1' },
  }, 'word', (r) => {
    if (!(r.data && r.data.filledCount >= 5)) throw new Error('docx filled<5');
    return 'filled=' + r.data.filledCount + '（docx 行为不变）';
  });
  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
