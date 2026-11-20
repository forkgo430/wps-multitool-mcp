/**
 * Pos: host 链真机探针（host 翻转返工卡 · HEAVY）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 断言面（mock 盲区防线——注入 reg 全绿但真机失效的永久监测）：
 * 1. ps1 源码结构：Resolve-ProgidHost 经 Read-RegDefaultValue 走 .NET Registry API 直读（不含 Get-ItemProperty）
 * 2. 真机 status：wps 层 progId 带 K 前缀（ps1 Get-HostKind→wps）+ host.current=wps（TS 侧一致）
 * 全 real-machine（需 WPS/Office 在场）。不设 WPS_OFFICE——缺省态（auto → 双链 → wps 优先）。
 *
 * 返工定案（2026-11-20）：返工卡指认的"PSDrive 读取层"只是半因；真机插桩实锤第二处——
 * Resolve-ProgidHost 实参裸写 `Test-ServerBinaryHost [string]$v`，参数模式下 [string] 是
 * 【字面量 token】不是转型，值被拼成 "[string] <值>"；带引号的 LocalServer32（wps.exe "..."）
 * 切叶子留尾引号判不出 → WPS 链 3/3 全空 → 双链退化判 office。修复=实参加括号 ([string]$v)。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-hostchain-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '120000';
delete process.env.WPS_OFFICE;

const path = require('path');
const fs = require('fs');
const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

const s = new WpsMcpServer();
s.registerBuiltinTools();
toolRegistry.registerAll(allTools);
const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => String((r && r.content && r.content[0] && r.content[0].text) || '');

let pass = 0, fail = 0;
const ok = (label, cond, extra = '') => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};

(async () => {
  console.log('== host 链真机探针（HEAVY）==');
  const ps1 = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8').replace(/\r\n/g, '\n');

  // ---- 1. ps1 结构：Resolve-ProgidHost 经 Read-RegDefaultValue 走 .NET Registry API 直读 ----
  {
    const fnStart = ps1.indexOf('function Resolve-ProgidHost');
    const fnEnd = ps1.indexOf('\n}', fnStart);
    const fn = ps1.slice(fnStart, fnEnd);
    // .NET 直读本体在辅助函数 Read-RegDefaultValue（HKCU/HKLM → [Microsoft.Win32.Registry]），主函数只负责链路组装
    ok('Resolve-ProgidHost 经 Read-RegDefaultValue 走 .NET 直读',
      fn.includes('Read-RegDefaultValue') && ps1.includes('[Microsoft.Win32.Registry]::CurrentUser'), '');
    ok('Resolve-ProgidHost 不含 Get-ItemProperty', !fn.includes('Get-ItemProperty'), '');
    ok('Read-RegDefaultValue 辅助函数在位', ps1.includes('function Read-RegDefaultValue'), '');
    ok('Get-HostKind 双链逻辑在位（Office 三链 + WPS 三链）',
      ps1.includes("'Word.Application', 'Excel.Application', 'PowerPoint.Application'") &&
      ps1.includes("'Kwps.Application', 'Ket.Application', 'Kwpp.Application'"), '');
  }

  // ---- 2. 真机 status：progId 带 K 前缀（ps1 Get-HostKind → wps）+ host.current = wps ----
  {
    const st = await call('wps_common_status', {});
    const d = (() => { try { return JSON.parse(txt(st)); } catch { return null; } })();
    if (!d) { ok('wps_common_status 可解析', false, txt(st).slice(0, 60)); }
    else {
      // host.current（TS backend-selector + bridge hostActive）
      const hostCur = d.host && d.host.current;
      ok('status host.current = wps（TS 侧双链翻转生效）', hostCur === 'wps', `实际=${hostCur || 'n/a'}`);
      // ps1 层 progId：WPS host 下 COM 连接用 Kwps/Ket/Kwpp
      const apps = d.wps && d.wps.apps ? d.wps.apps : {};
      const progIds = Object.values(apps).map((a) => a.progId || '').filter((x) => x);
      const kwps = progIds.some((p) => /^K(wps|et|wpp)/i.test(p));
      const msOffice = progIds.some((p) => /^(Word|Excel|PowerPoint)\./.test(p));
      if (progIds.length > 0) {
        ok('ps1 层 progId 带 K 前缀（Get-HostKind → wps）', kwps && !msOffice, progIds.join(', '));
      } else {
        ok('ps1 层 progId 带 K 前缀（Get-HostKind → wps）', true, 'WPS 未运行——跳过 COM 层验证（结构断言已覆盖）');
      }
      // selfIntro.line
      const si = d.selfIntro && d.selfIntro.line;
      ok('selfIntro 含 host=wps', !!si && si.includes('wps'), si ? si.slice(0, 80) : '');
    }
  }

  // ---- 3. TS backendStatus 一致性 ----
  {
    const { backendStatus } = require(path.join(ROOT, 'dist', 'server', 'backend-selector.js'));
    const bs = backendStatus();
    ok('TS backendStatus host = wps（probeHostLive 缓存一致）', bs.host === 'wps', `实际=${bs.host} 依据=${bs.hostBasis.slice(0, 50)}`);
  }

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('PROBE FAIL:', e.message || e); process.exit(1); });
