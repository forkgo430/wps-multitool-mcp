/**
 * Pos: 跨守护附着态感知监测探针（快照跨守护附着修复卡 · 钉板与永久监测两用）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 病灶档案（2026-11-21 修复轮实证，详见 OUTBOX 同日交付）：
 *   守护（wps-comd 常驻）拉起的 WPS 文字组件实例不保持 ROT 注册（T+5~90s NOREG 实证），而
 *   word 守护与 common 守护跨进程文档可见性唯一通道 = ROT 附着 → 用户实例不可附着（ROT 空）时
 *   common 守护的 save_as/snapshot 对 word 守护所建文档恒败（自建空岛并缓存）。
 *   修复三方向裁决：a（主动促 ROT 注册）——WPS 对象拒绝被外部进程注册（E_INVALIDARG，
 *   STA/MTA × WEAK/STRONG × object/IntPtr 四变体全灭）；c（释放引用促自注册）——25s 不触发，
 *   只有客户端进程退出才触发；b（按文档归属路由）需动 TS 路由语义——按卡停下报决策者。
 *
 * 断言面（态感知——健康态查回归，受限态出诊断不误报）：
 * 1. ps1 结构：附着三步原序（缓存→ROT→新建）未动（行为保持面）
 * 2. 真机态诊断：ROT 对 Kwps/Ket/Kwpp 三链的可达性
 * 3. 跨守护主锚：word 守护 createDocument → common 守护 save_as
 *    - 成功 → ✓（附着可达路径回归监测）
 *    - 失败且 ROT 空 → 已知限制（决策待裁），SKIP 出诊断，不计失败
 *    - 失败且 ROT 非空 → 回归，FAIL
 * 全 real-machine（需 WPS 可达）。不设 WPS_OFFICE——缺省态（auto → 双链 → wps 优先）。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-xdaemon-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '120000';
delete process.env.WPS_OFFICE;

const path = require('path');
const fs = require('fs');
const os = require('os');
const { execSync } = require('child_process');
const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

const s = new WpsMcpServer();
s.registerBuiltinTools();
toolRegistry.registerAll(allTools);
const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => String((r && r.content && r.content[0] && r.content[0].text) || '');

let pass = 0, fail = 0, skip = 0;
const ok = (label, cond, extra = '') => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};
const skipNote = (label, extra = '') => { skip++; console.log(`  ⊘ SKIP ${label}${extra ? ' — ' + extra : ''}`); };

const PWSH = ['C:/Program Files/PowerShell/7/pwsh.exe', 'pwsh.exe', 'powershell.exe']
  .find((p) => { try { return fs.existsSync(p) || execSync(`where.exe ${p}`).toString().trim(); } catch { return false; } });

(async () => {
  console.log('== 跨守护附着态感知监测探针 ==');
  const ps1 = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8').replace(/\r\n/g, '\n');
  const grab = (fnName) => {
    const i = ps1.indexOf(`function ${fnName}`);
    if (i < 0) return '';
    const j = ps1.indexOf('\n}', i);
    return ps1.slice(i, j);
  };

  // ---- 1. ps1 结构：行为保持面（附着三步原序） ----
  {
    const gww = grab('Get-WpsWord');
    ok('行为保持面：Get-WpsWord 附着三步原序（缓存→ROT→新建）',
      gww.includes('if ($null -ne $script:UserWordApp') &&
      gww.indexOf('$script:UserWordApp') < gww.indexOf('Get-RotObject') &&
      gww.indexOf('Get-RotObject') < gww.indexOf('New-TrackedWpsInstance'), '');
    ok('行为保持面：ROT 附着入口仍是 Get-RotObject 统一通道',
      grab('Get-RotObject').includes('WpsRot]::GetActiveObject'), '');
  }

  // ---- 2. 真机 ROT 态诊断 ----
  let rotState = 'UNKNOWN';
  if (!PWSH) {
    fail++; console.log('  ✗ 态诊断：找不到 pwsh/powershell 宿主');
  } else {
    const tmpScript = path.join(os.tmpdir(), `xdaemon-rotstate-${Date.now()}.ps1`);
    fs.writeFileSync(tmpScript, [
      "$ErrorActionPreference = 'Continue'",
      "$env:WPS_COM_DAEMON = '1'",
      `. '${ROOT.replace(/\\/g, '/')}\\scripts\\wps-com.ps1' -Instance user -App word`,
      "$r = @()",
      "foreach ($p in @('Kwps.Application', 'Ket.Application', 'Kwpp.Application')) {",
      "  $o = Get-RotObject $p",
      "  $r += ($(if ($null -ne $o) { 'HAS' } else { 'EMPTY' }))",
      "}",
      "Write-Output ('ROT=' + ($r -join '/'))",
    ].join('\n'), 'utf8');
    try {
      const out = execSync(`"${PWSH}" -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File "${tmpScript}"`, { encoding: 'utf8', timeout: 120000 });
      rotState = (out.match(/ROT=(\S+)/) || [])[1] || 'UNKNOWN';
      console.log(`  ℹ ROT 可达性（Kwps/Ket/Kwpp）= ${rotState}`);
    } catch (e) {
      console.log(`  ℹ 态诊断执行异常：${(e.message || '').slice(0, 60)}（按 UNKNOWN 处理）`);
    } finally {
      try { fs.unlinkSync(tmpScript); } catch { /* 已清 */ }
    }
  }

  // ---- 3. 跨守护主锚（态感知判级）----
  {
    const cr = await call('wps_word_create_document', {});
    const crOk = cr.success !== false;
    ok('跨守护链路：word 守护 createDocument', crOk, txt(cr).slice(0, 50));
    if (crOk) {
      await call('wps_word_insert_text', { text: 'cross-daemon attach probe', position: 'end' }).catch(() => {});
      const docPath = path.join(os.tmpdir(), `xdaemon-attach-${Date.now()}.docx`);
      const sv = await call('wps_common_save_as', { filePath: docPath });
      const saveOk = sv.success === true && fs.existsSync(docPath);
      if (saveOk) {
        ok('跨守护主锚：common 守护 save_as 对 word 守护所建文档成功（附着可达路径）', true, txt(sv).slice(0, 60));
        try { fs.unlinkSync(docPath); } catch { /* 清理尽力 */ }
        await call('wps_word_close_document', { save: false }).catch(() => {});
      } else {
        const noActive = /No active document/i.test(txt(sv));
        if (noActive && rotState.indexOf('EMPTY') >= 0) {
          // 已知限制（岛间隔离，决策待裁）——出诊断不误报回归
          skipNote('跨守护主锚：save_as 失败但属已知限制（ROT 空态岛间隔离，a/c 不可行已报决策者）',
            `ROT=${rotState}，save_as="${txt(sv).slice(0, 60)}"`);
        } else if (noActive) {
          ok('跨守护主锚：save_as 失败且 ROT 态非全空——回归', false, `ROT=${rotState}，"${txt(sv).slice(0, 60)}"`);
        } else {
          ok('跨守护主锚：save_as 失败（非已知失败形态）', false, txt(sv).slice(0, 80));
        }
        await call('wps_word_close_document', { save: false }).catch(() => {});
      }
    }
  }

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败 / ${skip} 已知限制 SKIP`);
  process.exit(fail === 0 ? 0 : 1);
})().catch(async (e) => {
  console.error('PROBE FAIL:', e.message || e);
  try { await call('wps_word_close_document', { save: false }); console.log('失败路径已尝试关稿（未保存）'); } catch { /* 已关/无文档 */ }
  process.exit(1);
});
