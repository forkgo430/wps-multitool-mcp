/**
 * 平台识别决策链的单测（Linux B0 验收项 + Office 支持一期 host 维度）
 * 原则：不碰真环境，全部注入 env / platform / exists / probe，把每层分支钉死。
 * 跑法：node scripts/manual/backend-selector-test.js（verify-all FAST 档包含它）
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { detectBackend, detectHost, BackendUnavailableError, WSL_MARKER,
        parseLocalServerBinary, resolveProgidHost,
        resolveChain, decideDualChainHost, OFFICE_PROGIDS, WPS_PROGIDS, hostPreference } = require(
  path.join(ROOT, 'dist', 'server', 'backend-selector.js')
);

let pass = 0;
let fail = 0;
const eq = (label, got, want) => {
  if (JSON.stringify(got) === JSON.stringify(want)) { pass++; console.log(`  ✓ ${label} = ${JSON.stringify(want)}`); }
  else { fail++; console.log(`  ✗ ${label}：得到 ${JSON.stringify(got)}，期望 ${JSON.stringify(want)}`); }
};
const throws = (label, fn, match) => {
  try { fn(); fail++; console.log(`  ✗ ${label}：没有抛错`); }
  catch (e) {
    if (e instanceof BackendUnavailableError && (!match || match.test(e.message))) { pass++; console.log(`  ✓ ${label}（BackendUnavailableError）`); }
    else { fail++; console.log(`  ✗ ${label}：抛了 ${(e && e.name) || typeof e}：${(e && e.message || '').slice(0, 80)}`); }
  }
};

const win32 = { env: {}, platform: 'win32', exists: () => false };
const linux = { env: {}, platform: 'linux', exists: () => false };
const wsl = { env: {}, platform: 'linux', exists: (p) => p === WSL_MARKER };
const darwin = { env: {}, platform: 'darwin', exists: () => false };

console.log('\n=== A. 第 2 层：原生 Windows ===');
eq('win32 → com', detectBackend(win32).backend, 'com');
eq('win32 判定依据带 win32', detectBackend(win32).basis.includes('win32'), true);

console.log('\n=== B. 第 1 层：WPS_BACKEND 强制覆盖（最高优先级，任何平台都生效）===');
eq('win32 + addon → addon（强制覆盖不问平台）', detectBackend({ ...win32, env: { WPS_BACKEND: 'addon' } }).backend, 'addon');
eq('win32 + off → off', detectBackend({ ...win32, env: { WPS_BACKEND: 'off' } }).backend, 'off');
eq('linux + off → off（不抛错）', detectBackend({ ...linux, env: { WPS_BACKEND: 'OFF' } }).backend, 'off');
eq('大小写不敏感（Addon）', detectBackend({ ...win32, env: { WPS_BACKEND: 'Addon' } }).backend, 'addon');
eq('win32 + com → com', detectBackend({ ...win32, env: { WPS_BACKEND: 'com' } }).backend, 'com');
throws('非法值抛 BackendUnavailableError', () => detectBackend({ ...win32, env: { WPS_BACKEND: 'bogus' } }), /不是合法值/);
// 判据修正（第一版期望"空格也抛"是错的）：空/空白值视为未设置、回落平台判定——
// 环境变量误设成空串不该把整个 server 弄瘫；真正要拦的是非空垃圾值（上一条）。
eq('空白值视为未设置（回落平台判定）', detectBackend({ ...win32, env: { WPS_BACKEND: '   ' } }).backend, 'com');

console.log('\n=== C. 第 3 层：WSL 特判 ===');
eq('linux + Windows PowerShell 标记 → com', detectBackend(wsl).backend, 'com');
eq('WSL 判定 wsl=true', detectBackend(wsl).wsl, true);
eq('WSL 依据里说明管道可达性归通道层', detectBackend(wsl).basis.includes('通道层'), true);

console.log('\n=== D. 第 4/5 层：Linux/macOS 通道未实现 → 明确抛错 ===');
throws('linux 无标记 → 抛错并给出三选项', () => detectBackend(linux), /尚未实现/);
throws('darwin → 同样抛错', () => detectBackend(darwin), /macOS/);
throws('抛错文案提到三型加载项', () => detectBackend(linux), /wps\/et\/wpp/);

console.log('\n=== E. 失败不缓存（每次调用都能拿到最新判定）===');
const flip = { env: {}, platform: 'win32', exists: () => false };
eq('第一次 win32 → com', detectBackend(flip).backend, 'com');

// ==================== Office 支持一期：host 维度（WPS_OFFICE）====================
// 依据：卡面"判定阶段情报"（11-07）——WPS 顶名劫持下 ProgID 存在性/CoCreate 成败/
// Name 自报全部不可信，探测必须解析 ProgID 链到实际服务器二进制。
console.log('\n=== F. host 维度：WPS_OFFICE 强制 on/off（不探测）===');
const officeProbe = () => 'office';
const wpsProbe = () => 'wps';
eq('on → office', detectBackend({ ...win32, env: { WPS_OFFICE: 'on' }, probe: wpsProbe }).host, 'office');
eq('off → wps（probe 说不算）', detectBackend({ ...win32, env: { WPS_OFFICE: 'off' }, probe: officeProbe }).host, 'wps');
eq('大小写不敏感（ON）', detectBackend({ ...win32, env: { WPS_OFFICE: 'ON' } }).host, 'office');
eq('host 依据说明强制来源', detectBackend({ ...win32, env: { WPS_OFFICE: 'on' } }).hostBasis.includes('强制指定'), true);
throws('非法值抛错（bogus）', () => detectBackend({ ...win32, env: { WPS_OFFICE: 'bogus' } }), /WPS_OFFICE.*不是合法值/);
// 空白值同未设（与 WPS_BACKEND 空白口径一致，误设空串不弄瘫 server）
eq('空白值视为未设（回落 auto 探测）', detectHost({ env: { WPS_OFFICE: '   ' }, probe: officeProbe }).host, 'office');

console.log('\n=== G. host 维度：auto 缺省探测 + 保守回落 ===');
eq('auto 显式 + probe=office → office', detectHost({ env: { WPS_OFFICE: 'auto' }, probe: officeProbe }).host, 'office');
eq('未设（缺省 auto）+ probe=office → office', detectHost({ env: {}, probe: officeProbe }).host, 'office');
eq('未设 + probe=wps → wps', detectHost({ env: {}, probe: wpsProbe }).host, 'wps');
eq('未设 + probe=null（探测不可用）→ 保守 wps', detectHost({ env: {}, probe: () => null }).host, 'wps');
eq('未设 + 无 probe 函数 → 保守 wps', detectHost({ env: {} }).host, 'wps');
eq('保守回落依据说明原因', detectHost({ env: {}, probe: () => null }).hostBasis.includes('保守按 wps'), true);

console.log('\n=== H. host × WPS_BACKEND 优先级矩阵（正交性钉死）===');
// backend 优先于一切；host 只在 backend ≠ off 时有意义
eq('矩阵① backend=off + OFFICE=on → off + host=wps（不适用）',
   detectBackend({ ...win32, env: { WPS_BACKEND: 'off', WPS_OFFICE: 'on' }, probe: officeProbe }).backend, 'off');
eq('矩阵① host 固定 wps', detectBackend({ ...win32, env: { WPS_BACKEND: 'off', WPS_OFFICE: 'on' }, probe: officeProbe }).host, 'wps');
eq('矩阵② backend=com + OFFICE=on → com/office',
   JSON.stringify([detectBackend({ ...win32, env: { WPS_BACKEND: 'com', WPS_OFFICE: 'on' }, probe: wpsProbe }).backend,
                   detectBackend({ ...win32, env: { WPS_BACKEND: 'com', WPS_OFFICE: 'on' }, probe: wpsProbe }).host]),
   JSON.stringify(['com', 'office']));
eq('矩阵③ backend=com + OFFICE=off → com/wps',
   JSON.stringify([detectBackend({ ...win32, env: { WPS_BACKEND: 'com', WPS_OFFICE: 'off' }, probe: officeProbe }).backend,
                   detectBackend({ ...win32, env: { WPS_BACKEND: 'com', WPS_OFFICE: 'off' }, probe: officeProbe }).host]),
   JSON.stringify(['com', 'wps']));
eq('矩阵④ backend 强制 com + OFFICE 未设 auto 探测 → com/office',
   JSON.stringify([detectBackend({ ...win32, env: { WPS_BACKEND: 'com' }, probe: officeProbe }).backend,
                   detectBackend({ ...win32, env: { WPS_BACKEND: 'com' }, probe: officeProbe }).host]),
   JSON.stringify(['com', 'office']));
eq('矩阵⑤ 两个开关都未设 → com/wps（保守）',
   JSON.stringify([detectBackend({ ...win32, probe: () => null }).backend,
                   detectBackend({ ...win32, probe: () => null }).host]),
   JSON.stringify(['com', 'wps']));
eq('矩阵⑥ addon 强制 + OFFICE=on → addon/office',
   JSON.stringify([detectBackend({ ...win32, env: { WPS_BACKEND: 'addon', WPS_OFFICE: 'on' } }).backend,
                   detectBackend({ ...win32, env: { WPS_BACKEND: 'addon', WPS_OFFICE: 'on' } }).host]),
   JSON.stringify(['addon', 'office']));
throws('矩阵⑦ backend 非法 + OFFICE=on → 仍抛（backend 层先死）',
   () => detectBackend({ ...win32, env: { WPS_BACKEND: 'bogus', WPS_OFFICE: 'on' } }), /WPS_BACKEND/);
throws('矩阵⑧ backend 合法 + OFFICE 非法 → 抛（host 层显式失败）',
   () => detectBackend({ ...win32, env: { WPS_BACKEND: 'com', WPS_OFFICE: 'bogus' } }), /WPS_OFFICE/);
// WSL 平台 host 维度同样生效（COM 通道共享同一判定）
eq('矩阵⑨ WSL + OFFICE=on → com/office',
   JSON.stringify([detectBackend({ ...wsl, env: { WPS_OFFICE: 'on' } }).backend,
                   detectBackend({ ...wsl, env: { WPS_OFFICE: 'on' } }).host]),
   JSON.stringify(['com', 'office']));
// linux 非 WSL 抛错路径不受 host 影响
throws('矩阵⑩ linux 无标记 → 仍抛（host 不救通道）', () => detectBackend({ ...linux, env: { WPS_OFFICE: 'on' } }), /尚未实现/);

console.log('\n=== I. parseLocalServerBinary：LocalServer32 → exe 真身归类（顶名劫持防线）===');
// 真机实测的两种注册形态（2026-11-07 本机 Office 365 C2R + WPS 12.1 劫持形态）
eq('Office 无引号形态 → office',
   parseLocalServerBinary('C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE /Automation').host, 'office');
eq('Office 引号形态 → office',
   parseLocalServerBinary('"C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE" /Automation').host, 'office');
eq('WPS 顶名劫持形态（官方 CLSID 指向 wps.exe）→ wps',
   parseLocalServerBinary('D:\\APPdate\\WPS Office\\12.1.0.29680\\office6\\wps.exe /Automation').host, 'wps');
eq('WPS 引号形态 → wps',
   parseLocalServerBinary('"D:\\APPdate\\WPS Office\\12.1.0.29680\\office6\\et.exe" /Automation').host, 'wps');
eq('Excel 服务器 → office', parseLocalServerBinary('C:\\Program Files\\Microsoft Office\\Root\\Office16\\EXCEL.EXE /automation').host, 'office');
eq('PPT 服务器 → office', parseLocalServerBinary('"C:\\Program Files\\Microsoft Office\\Root\\Office16\\POWERPNT.EXE" /automation').host, 'office');
eq('认不出的二进制 → unknown（不猜）', parseLocalServerBinary('C:\\Tools\\ Weird.exe /x').host, 'unknown');
eq('空值 → null', parseLocalServerBinary(''), null);
eq('解析出 exe 路径原样保留', parseLocalServerBinary('C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE /Automation').exe,
   'C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE');

console.log('\n=== J. resolveProgidHost：ProgID 链解析（注入 reg 读取器，不碰真机）===');
// 场景 1：Office 真装（HKLM ProgID + HKLM CLSID → WINWORD.EXE）
const regOffice = (key) => {
  if (key === 'HKLM\\SOFTWARE\\Classes\\Word.Application\\CLSID') return '{000209FF-0000-0000-C000-000000000046}';
  if (key === 'HKLM\\SOFTWARE\\Classes\\CLSID\\{000209FF-0000-0000-C000-000000000046}\\LocalServer32') return 'C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE /Automation';
  return null;
};
eq('Office 真装 → office', resolveProgidHost('Word.Application', regOffice), 'office');
// 场景 2：WPS 独装机劫持（HKCU ProgID + 官方 CLSID 只在 WOW6432Node → wps.exe）
const regHijack = (key) => {
  if (key === 'HKCU\\SOFTWARE\\Classes\\Word.Application\\CLSID') return '{000209FF-0000-0000-C000-000000000046}';
  if (key === 'HKCU\\SOFTWARE\\Classes\\WOW6432Node\\CLSID\\{000209FF-0000-0000-C000-000000000046}\\LocalServer32') return 'D:\\APPdate\\WPS Office\\12.1.0.29680\\office6\\wps.exe /Automation';
  return null;
};
eq('WPS 独装劫持 → wps（不是 office！）', resolveProgidHost('Word.Application', regHijack), 'wps');
// 场景 3：ProgID 不存在 → null
eq('ProgID 缺失 → null', resolveProgidHost('Word.Application', () => null), null);
// 场景 4：ProgID 在但 CLSID 断链（LocalServer32 缺失）→ null（不猜）
const regBroken = (key) => {
  if (key === 'HKCU\\SOFTWARE\\Classes\\Word.Application\\CLSID') return '{000209FF-0000-0000-C000-000000000046}';
  return null;
};
eq('CLSID 断链 → null', resolveProgidHost('Word.Application', regBroken), null);
// 场景 5：HKCU 覆盖 HKLM（合并视图优先级：HKCU 先读）
const regHkcuWins = (key) => {
  if (key === 'HKCU\\SOFTWARE\\Classes\\Word.Application\\CLSID') return '{HKCU-WINS}';
  if (key === 'HKLM\\SOFTWARE\\Classes\\Word.Application\\CLSID') return '{SHOULD-NOT-REACH}';
  if (key === 'HKCU\\SOFTWARE\\Classes\\CLSID\\{HKCU-WINS}\\LocalServer32') return 'C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE /Automation';
  return null;
};
eq('HKCU 优先于 HKLM（合并视图语义）', resolveProgidHost('Word.Application', regHkcuWins), 'office');
// 场景 6：指向未知二进制 → null（unknown 不当结论）
const regUnknownExe = (key) => {
  if (key === 'HKCU\\SOFTWARE\\Classes\\Word.Application\\CLSID') return '{X}';
  if (key === 'HKCU\\SOFTWARE\\Classes\\CLSID\\{X}\\LocalServer32') return 'C:\\Whatever\\LibreOffice.exe /x';
  return null;
};

console.log('\n=== K. 双链探测取舍（host wps 优先 · 2026-11-17 翻转）：resolveChain + decideDualChainHost（注入 reg，不碰真机）===');
// 双装注册表：Office 三链真身 + WPS 三链真身都在（双装机器的注册形态）
const regDual = (key) => {
  const officeClsids = {
    'HKLM\\SOFTWARE\\Classes\\Word.Application\\CLSID': '{000209FF-0000-0000-C000-000000000046}',
    'HKLM\\SOFTWARE\\Classes\\Excel.Application\\CLSID': '{00020820-0000-0000-C000-000000000046}',
    'HKLM\\SOFTWARE\\Classes\\PowerPoint.Application\\CLSID': '{91493441-5A91-11CF-8705-0020AFDDDCE4}',
  };
  const officeSrv = {
    '{000209FF-0000-0000-C000-000000000046}': 'C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE /Automation',
    '{00020820-0000-0000-C000-000000000046}': 'C:\\Program Files\\Microsoft Office\\Root\\Office16\\EXCEL.EXE /Automation',
    '{91493441-5A91-11CF-8705-0020AFDDDCE4}': 'C:\\Program Files\\Microsoft Office\\Root\\Office16\\POWERPNT.EXE /Automation',
  };
  if (officeClsids[key]) return officeClsids[key];
  const oc = Object.keys(officeSrv).find((c) => key.includes(c));
  if (oc && key.includes('LocalServer32')) return officeSrv[oc];
  const wpsMap = {
    'HKLM\\SOFTWARE\\Classes\\Kwps.Application\\CLSID': '{KWPS-CLSID}',
    'HKLM\\SOFTWARE\\Classes\\Ket.Application\\CLSID': '{KET-CLSID}',
    'HKLM\\SOFTWARE\\Classes\\Kwpp.Application\\CLSID': '{KWPP-CLSID}',
  };
  if (wpsMap[key]) return wpsMap[key];
  const wpsSrv = {
    '{KWPS-CLSID}': 'D:\\APPdate\\WPS Office\\12.1.0.29680\\office6\\wps.exe /Automation',
    '{KET-CLSID}': 'D:\\APPdate\\WPS Office\\12.1.0.29680\\office6\\et.exe /Automation',
    '{KWPP-CLSID}': 'D:\\APPdate\\WPS Office\\12.1.0.29680\\office6\\wpp.exe /Automation',
  };
  const wc = Object.keys(wpsSrv).find((c) => key.includes(c));
  if (wc && key.includes('LocalServer32')) return wpsSrv[wc];
  return null;
};
eq('K1 Office 链成立（双装）', resolveChain(OFFICE_PROGIDS, 'office', regDual).host, 'office');
eq('K2 WPS 链成立（双装）', resolveChain(WPS_PROGIDS, 'wps', regDual).host, 'wps');
eq('K3 双链都在 + 缺省偏好 → wps（翻转点）', decideDualChainHost('office', 'wps', hostPreference({})), 'wps');
eq('K4 双链都在 + PREFERENCE=office → office', decideDualChainHost('office', 'wps', 'office'), 'office');
eq('K5 仅 WPS 链（独装/劫持）→ wps（偏好不影响单链）', decideDualChainHost(null, 'wps', 'office'), 'wps');
eq('K6 仅 Office 链 → office', decideDualChainHost('office', null, 'wps'), 'office');
eq('K7 双链都不成立 → null（保守回落）', decideDualChainHost(null, null, 'wps'), null);
// 顶名劫持：Office ProgID 全部解析到 wps 二进制 → Office 链不成立（不是结论 office=wps）
const regHijackAll = (key) => {
  if (/\\(Word|Excel|PowerPoint)\.Application\\CLSID$/.test(key)) return '{HIJACKED}';
  if (key.includes('LocalServer32')) return 'D:\\APPdate\\WPS Office\\12.1.0.29680\\office6\\wps.exe /Automation';
  return null;
};
const hij = resolveChain(OFFICE_PROGIDS, 'office', regHijackAll);
eq('K8 顶名劫持 → Office 链不成立', hij.host, null);
eq('K9 顶名劫持 → hijacked 实锤', hij.hijacked, true);
eq('K10 劫持 + WPS 链成立 → wps（独装形态，偏好不翻）', decideDualChainHost(hij.host, 'wps', 'office'), 'wps');
eq('K11 断链残缺 → Office 链不成立', resolveChain(OFFICE_PROGIDS, 'office', () => null).host, null);
eq('K12 非法偏好值按 wps（不抛）', hostPreference({ WPS_HOST_PREFERENCE: 'bogus' }), 'wps');

eq('未知二进制 → null（绝不猜）', resolveProgidHost('Word.Application', regUnknownExe), null);

console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
process.exit(fail === 0 ? 0 : 1);
