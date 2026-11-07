/**
 * 双装机绑定一致性 · 双场景验收（Office 收口合并卡 A 段第 4 项）
 *
 * 多实例场（每 app ≥3 实例在场）跑 word/excel/ppt 生命周期全绿；
 * 修复 A（缓存探活优先于 ROT）后同会话同 app 必须绑同一实例——生命周期不断链即为过。
 * 单实例场（清场后）再跑一轮。
 *
 * 判据：两场景 × 三 app 生命周期全绿。live 验收脚本（需要真机 + 真 Office），判定归决策者。
 */
process.env.WPS_OFFICE = process.env.WPS_OFFICE || 'on';
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_PIPE = process.env.WPS_BRIDGE_PIPE || `\\\\.\\pipe\\wps-multitool-mcp-bindacc-${Date.now()}`;

const path = require('path');
const fs = require('fs');
const { execFileSync } = require('child_process');
const ROOT = path.resolve(__dirname, '..', '..');
const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));

const args = process.argv.slice(2);
const SCENE = args.includes('--single') ? 'single' : 'multi';
const TMP = path.join(process.env.TEMP || 'C:/Temp', 'binding-accept');
fs.mkdirSync(TMP, { recursive: true });

const APP_SPEC = {
  word: { progid: 'Word.Application', doc: 'Documents' },
  excel: { progid: 'Excel.Application', doc: 'Workbooks' },
  ppt: { progid: 'PowerPoint.Application', doc: 'Presentations' },
};

function spawnInstance(app, i) {
  const spec = APP_SPEC[app];
  const script = `
    $w = New-Object -ComObject '${spec.progid}'
    $w.Visible = $true
    Start-Sleep -Milliseconds 500
  `;
  const f = path.join(process.env.TEMP || 'C:/Temp', `bindacc-${app}-${i}.ps1`);
  fs.writeFileSync(f, script, 'utf8');
  try { execFileSync('powershell', ['-NoProfile', '-ExecutionPolicy', '-Bypass', '-File', f].slice(0, 3).concat(['-ExecutionPolicy', 'Bypass', '-File', f]), { timeout: 60000, stdio: 'ignore' }); return true; }
  catch { try { execFileSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', f], { timeout: 60000, stdio: 'ignore' }); return true; } catch { return false; } }
}

function appPids(app) {
  const name = { word: 'WINWORD', excel: 'EXCEL', ppt: 'POWERPNT' }[app];
  try {
    const out = execFileSync('powershell', ['-NoProfile', '-Command', `Get-Process ${name} -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }`], { encoding: 'utf8', timeout: 20000 });
    return out.split(/\s+/).filter(Boolean).map(Number);
  } catch { return []; }
}

function killPids(pids) {
  for (const p of pids) { try { execFileSync('taskkill', ['/F', '/PID', String(p)], { timeout: 10000, stdio: 'ignore' }); } catch { /* 已不在 */ } }
}

let pass = 0, fail = 0;
async function run(label, action, params, app, assert) {
  try {
    const r = await wpsClient.executeMethod(action, params || {}, app);
    if (r && r.success === true) {
      let note = '';
      if (assert) { try { note = assert(r) || ''; } catch (e) { fail++; console.log(`  ✗ ${label} 断言: ${e.message}`); return r; } }
      pass++; console.log(`  ✓ ${label} ${note}`);
    } else { fail++; console.log(`  ✗ ${label}: ${(r && r.error || '').slice(0, 140)}`); }
    return r;
  } catch (e) { fail++; console.log(`  ✗ ${label} EXC: ${(e.message || '').slice(0, 140)}`); return null; }
}

async function lifecycle(app, tmpName) {
  console.log(`--- ${app} 生命周期（${SCENE} 场）---`);
  if (app === 'word') {
    await run('create', 'createDocument', {}, 'word');
    await run('insert', 'insertText', { text: `绑定验收 ${SCENE}`, position: 'end' }, 'word');
    await run('stats（读回对账）', 'getDocumentStats', {}, 'word', (r) => {
      if (!(r.data && r.data.paragraphs > 0)) throw new Error('paragraphs=0');
      return 'paragraphs=' + r.data.paragraphs;
    });
    const f = path.join(TMP, tmpName + '.docx');
    await run('saveAs', 'saveAs', { path: f }, 'word', () => { if (!fs.existsSync(f)) throw new Error('未落盘'); return 'saved'; });
    await run('close', 'closeDocument', { saveChanges: true }, 'word');
    await run('reopen', 'openDocument', { path: f }, 'word', (r) => 'name=' + (r.data && r.data.name));
    await run('close（收尾）', 'closeDocument', { saveChanges: false }, 'word');
  } else if (app === 'excel') {
    await run('create', 'createWorkbook', {}, 'excel');
    await run('set', 'setCellValue', { row: 1, col: 1, value: '绑定验收' }, 'excel');
    await run('get（读回对账）', 'getCellValue', { row: 1, col: 1 }, 'excel', (r) => {
      const v = r.data && (r.data.value ?? r.data);
      if (v !== '绑定验收') throw new Error('值漂移: ' + JSON.stringify(v));
      return 'value OK';
    });
    const f = path.join(TMP, tmpName + '.xlsx');
    await run('saveAs', 'saveAs', { path: f }, 'excel', () => { if (!fs.existsSync(f)) throw new Error('未落盘'); return 'saved'; });
    await run('close', 'closeWorkbook', { saveChanges: true }, 'excel');
    await run('reopen', 'openWorkbook', { path: f }, 'excel', (r) => 'name=' + (r.data && r.data.name));
    await run('close（收尾）', 'closeWorkbook', { saveChanges: false }, 'excel');
  } else {
    await run('create', 'createPresentation', {}, 'ppt');
    await run('addSlide', 'addSlide', { layout: 'blank' }, 'ppt');
    await run('info（读回）', 'getSlideInfo', { slideIndex: 1 }, 'ppt');
    const f = path.join(TMP, tmpName + '.pptx');
    await run('saveAs', 'saveAs', { path: f }, 'ppt', () => { if (!fs.existsSync(f)) throw new Error('未落盘'); return 'saved'; });
    await run('close', 'closePresentation', {}, 'ppt');
    await run('reopen', 'openPresentation', { path: f }, 'ppt', (r) => 'name=' + (r.data && (r.data.name || '')));
    await run('close（收尾）', 'closePresentation', {}, 'ppt');
  }
}

(async () => {
  console.log(`=== 双场景绑定验收（${SCENE} 场）===`);
  const spawned = [];
  if (SCENE === 'multi') {
    for (const app of Object.keys(APP_SPEC)) {
      const pre = new Set(appPids(app));
      for (let i = 1; i <= 3; i++) {
        if (spawnInstance(app, i)) { await new Promise((r) => setTimeout(r, 700)); }
      }
      const now = appPids(app);
      const mine = now.filter((p) => !pre.has(p));
      spawned.push(...mine.map((p) => ({ app, pid: p })));
      console.log(`  造场 ${app}: +${mine.length} 实例（在场 ${now.length}）`);
    }
    await new Promise((r) => setTimeout(r, 1500));
  }

  for (const app of Object.keys(APP_SPEC)) await lifecycle(app, `bind-${SCENE}-${app}`);

  console.log(`\n${SCENE} 场汇总：${pass} 通过 / ${fail} 失败`);
  // 清场：只收本脚本造的
  for (const { pid } of spawned) killPids([pid]);
  if (spawned.length) console.log(`清场：已收本脚本造的 ${spawned.length} 个实例`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('验收异常：', e && e.stack ? e.stack : e); process.exit(2); });
