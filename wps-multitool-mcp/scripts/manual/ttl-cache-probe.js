/**
 * Pos: 读结果 TTL 缓存探针（代码技法卡 F 段——只探不上线，上线与否归决策者）。
 * 一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 量什么：三类高频读（get_cell_value / get_range_data / get_slide_info）在短窗口内
 * 重复调用时，真实通道往返耗时 vs 模拟 TTL 缓存命中（内存直出）——量出"缓存能省多少毫秒/次"。
 *
 * 怎么量：
 *   ① 冷读一次（真实通道，同时当缓存填充）
 *   ② 30 次重复调用全部走真实通道，取平均 = 单次通道成本
 *   ③ 同 30 次参数的"缓存命中"模拟 = 内存 map 直出（0 通道），实测内存取用耗时
 *   ④ 输出：每工具的通道均值 / 命中均值 / 每 100 次读的节省估算（按命中 100% 与 70% 两档）
 *
 * 纪律：自建工作簿/演示文稿自己关（save:false）；独立管道 + EPHEMERAL（探针退出桥自灭）。
 * 产物：evidence JSON 落 scripts/manual/（Temp 会清）。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-ttl-probe-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '120000';

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

/** 模拟 TTL 缓存层：键 = 工具名 + 参数指纹（稳定序列化） */
function stableStringify(v) {
  if (v === null || typeof v !== 'object') return JSON.stringify(v);
  if (Array.isArray(v)) return '[' + v.map(stableStringify).join(',') + ']';
  const keys = Object.keys(v).sort();
  return '{' + keys.map((k) => JSON.stringify(k) + ':' + stableStringify(v[k])).join(',') + '}';
}
const cache = new Map();
const cacheKey = (tool, args) => `${tool}|${stableStringify(args)}`;
const cachedCall = async (tool, args, ttlMs = 5000) => {
  const key = cacheKey(tool, args);
  const hit = cache.get(key);
  if (hit && Date.now() - hit.t < ttlMs) return hit.value;
  const value = await call(tool, args);
  cache.set(key, { t: Date.now(), value });
  return value;
};

async function timeN(tool, args, n) {
  // 预热（冷读进缓存）
  await cachedCall(tool, args);
  const t0 = Date.now();
  for (let i = 0; i < n; i++) await call(tool, args);
  const channelAvg = (Date.now() - t0) / n;
  const t1 = Date.now();
  for (let i = 0; i < n; i++) await cachedCall(tool, args);
  const hitAvg = (Date.now() - t1) / n;
  return { channelAvg, hitAvg };
}

(async () => {
  const N = 30;
  const out = { probe: 'ttl-cache', ts: new Date().toISOString(), tools: {} };
  console.log('== TTL 缓存探针（真实通道 vs 模拟命中，各 30 次）==');

  // Excel 场：自建工作簿 + 已知数据
  const wb = await call('wps_excel_create_workbook', {});
  if (!String((wb.content || [{}])[0]?.text || '').includes('成功') && !String((wb.content || [{}])[0]?.text || '')) {
    console.log('  ⚠ create_workbook 回包：', txt(wb));
  }
  await call('wps_excel_write_range', { range: 'A1:B3', data: [['名', '值'], ['a', 1], ['b', 2]] });

  // 三类读：参数固定（同一窗口内重复调用的真实形态）
  const cases = [
    ['wps_excel_get_cell_value', { sheet: 'Sheet1', row: 1, col: 1 }],
    ['wps_excel_read_range', { range: 'A1:B3', sheet: 'Sheet1' }],
  ];
  for (const [tool, args] of cases) {
    try {
      const r = await timeN(tool, args, N);
      out.tools[tool] = r;
      console.log(`  ${tool}: 通道均值 ${r.channelAvg.toFixed(1)}ms | 命中均值 ${r.hitAvg.toFixed(2)}ms | 每100次读省 ${(r.channelAvg - r.hitAvg).toFixed(0)}ms（100%命中）/ ${((r.channelAvg - r.hitAvg) * 0.7).toFixed(0)}ms（70%命中）`);
    } catch (e) {
      console.log(`  ⚠ ${tool} 探针失败: ${e.message}`);
      out.tools[tool] = { error: e.message };
    }
  }

  // PPT 场：自建演示文稿 + 一页
  try {
    await call('wps_ppt_create_presentation', {});
    await call('wps_ppt_add_slide', { title: 'TTL 探针' });
    const r = await timeN('wps_ppt_get_slide_info', { slideIndex: 1 }, N);
    out.tools['wps_ppt_get_slide_info'] = r;
    console.log(`  wps_ppt_get_slide_info: 通道均值 ${r.channelAvg.toFixed(1)}ms | 命中均值 ${r.hitAvg.toFixed(2)}ms | 每100次读省 ${(r.channelAvg - r.hitAvg).toFixed(0)}ms（100%命中）`);
    await call('wps_ppt_close_presentation', { save: false });
  } catch (e) {
    console.log(`  ⚠ wps_ppt_get_slide_info 探针失败: ${e.message}`);
    out.tools['wps_ppt_get_slide_info'] = { error: e.message };
  }

  // 收尾：自建工作簿自己关
  try { await call('wps_excel_close_workbook', { save: false, workbook: wb && '工作簿1' }); } catch { /* 已关/名字不匹配不阻塞 */ }

  const evDir = path.join(ROOT, 'scripts', 'manual');
  const evFile = path.join(evDir, `ttl-cache-probe-evidence.json`);
  fs.writeFileSync(evFile, JSON.stringify(out, null, 2), 'utf8');
  console.log(`证据：${evFile}`);
  process.exit(0);
})().catch((e) => { console.error('PROBE FAIL:', e); process.exit(1); });

function txt(r) { return String((r && r.content && r.content[0] && r.content[0].text) || ''); }
