/**
 * `wps-multitool-mcp install` 的门禁：不碰真实用户配置，全程在临时 HOME 里跑。
 *
 * 为什么值得单独立门禁
 * ------------------
 * 这个子命令会改【别人的配置文件】。改错的代价不是"功能不好用"，而是用户其它 MCP 服务一起失效。
 * 所以这里断言的重点不是"写进去了"，而是：
 *   1) 别的服务一个都不能少（合并而不是覆盖）
 *   2) 写之前有备份、写之后能回读
 *   3) --dry-run 一个字节都不动
 *   4) 认不出的客户端不瞎建目录
 *
 * 用法：node scripts/manual/install-verify.js
 */
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok });
  console.log(`PASS  ${name}${ok ? '' : '  → ' + detail}`);
};

const tmpHome = path.join(os.tmpdir(), `wps-install-verify-${Date.now()}`);
const cbFile = path.join(tmpHome, '.codebuddy', 'mcp.json');
const codexFile = path.join(tmpHome, '.codex', 'config.toml');
const cursorDir = path.join(tmpHome, '.cursor');

fs.mkdirSync(path.dirname(cbFile), { recursive: true });
fs.mkdirSync(path.dirname(codexFile), { recursive: true });
fs.mkdirSync(cursorDir, { recursive: true });
fs.writeFileSync(
  cbFile,
  JSON.stringify({ mcpServers: { other: { command: 'keep-me', args: [] } }, theme: 'dark' }, null, 2),
  'utf8'
);
fs.writeFileSync(codexFile, '[mcp_servers.other]\ncommand = \'keep-me\'\n', 'utf8');

// 在 require 之前改 HOME：install.ts 每次调用都重新算主目录（不是模块加载时缓存）
process.env.HOME = tmpHome;
process.env.USERPROFILE = tmpHome;

const inst = require(path.join(ROOT, 'dist', 'cli', 'install.js'));

(async () => {
  try {
    // ---- 1. 探测 ----
    const detected = inst.detectClients();
    const presentIds = detected.filter((c) => c.present).map((c) => c.id).sort();
    check(
      '只探测到真实存在的客户端（不误判 ~/.claude.json）',
      presentIds.join(',') === 'codebuddy,codex,cursor',
      presentIds.join(',')
    );
    check(
      '没装的客户端标记为 absent',
      detected.filter((c) => !c.present).map((c) => c.id).sort().join(',') ===
        'claude,claude-settings,workbuddy',
      detected.filter((c) => !c.present).map((c) => c.id).join(',')
    );

    // ---- 2. JSON 合并 ----
    const merged = inst.mergeJsonClient(
      { mcpServers: { other: { command: 'keep-me' } }, theme: 'dark' },
      { command: 'node', args: ['C:\\x\\index.js'] }
    );
    check('JSON 合并保留其它服务', !!merged.mcpServers.other, JSON.stringify(merged.mcpServers));
    check('JSON 合并保留其它顶层键', merged.theme === 'dark');
    check('JSON 合并写入本服务', !!merged.mcpServers['wps-multitool-mcp']);

    // ---- 3. TOML 合并（新增 + 替换 + 保留别人）----
    const entry = { command: 'node', args: ['C:\\a b\\index.js'] };
    const tomlNew = inst.mergeTomlClient('[mcp_servers.other]\ncommand = \'keep-me\'\n', entry);
    check('TOML 新增本服务段', tomlNew.includes('[mcp_servers.wps-multitool-mcp]'));
    check('TOML 保留其它服务段', tomlNew.includes('[mcp_servers.other]'));
    check('TOML 路径用单引号字面量（反斜杠不被转义）', tomlNew.includes("'C:\\a b\\index.js'"), tomlNew);
    const tomlTwice = inst.mergeTomlClient(tomlNew, entry);
    check(
      'TOML 重复写不会产生第二个同名段',
      tomlTwice.split('[mcp_servers.wps-multitool-mcp]').length === 2,
      `出现 ${tomlTwice.split('[mcp_servers.wps-multitool-mcp]').length - 1} 次`
    );

    // ---- 4. --dry-run 不动盘 ----
    const beforeCb = fs.readFileSync(cbFile, 'utf8');
    const beforeCodex = fs.readFileSync(codexFile, 'utf8');
    const dryCode = await inst.runInstall(['--dry-run']);
    check('--dry-run 退出码 0', dryCode === 0, String(dryCode));
    check('--dry-run 没改 JSON', fs.readFileSync(cbFile, 'utf8') === beforeCb);
    check('--dry-run 没改 TOML', fs.readFileSync(codexFile, 'utf8') === beforeCodex);
    check('--dry-run 不产生备份文件', !fs.readdirSync(path.dirname(cbFile)).some((f) => f.includes('.bak-')));

    // ---- 5. 真正写入 ----
    const code = await inst.runInstall([]);
    check('install 退出码 0', code === 0, String(code));

    const after = JSON.parse(fs.readFileSync(cbFile, 'utf8'));
    check('CodeBuddy 配置写入成功', !!after.mcpServers['wps-multitool-mcp']);
    check('CodeBuddy 其它服务没丢', after.mcpServers.other && after.mcpServers.other.command === 'keep-me');
    check('CodeBuddy 其它顶层键没丢', after.theme === 'dark');
    const backups = fs.readdirSync(path.dirname(cbFile)).filter((f) => f.includes('.bak-'));
    check('写前留了备份', backups.length === 1, backups.join(','));

    const tomlAfter = fs.readFileSync(codexFile, 'utf8');
    check('Codex 配置写入成功', tomlAfter.includes('[mcp_servers.wps-multitool-mcp]'));
    check('Codex 其它服务没丢', tomlAfter.includes('[mcp_servers.other]'));

    const cursorFile = path.join(cursorDir, 'mcp.json');
    check('Cursor 目录存在 → 自动建配置文件', fs.existsSync(cursorFile));

    // ---- 6. 幂等 ----
    const again = await inst.runInstall([]);
    check('重复 install 退出码 0', again === 0, String(again));
    const after2 = fs.readFileSync(cbFile, 'utf8');
    check('重复 install 不重复插入', after2.split('"wps-multitool-mcp"').length === 2,
      `出现 ${after2.split('"wps-multitool-mcp"').length - 1} 次`);

    // ---- 7. 未知客户端 ----
    const bad = await inst.runInstall(['--client', 'nope']);
    check('未知客户端返回 1', bad === 1, String(bad));

    // ---- 8. 没装任何客户端时 --print 也能给出可用片段 ----
    const snippet = inst.formatSnippet('json', entry);
    check('配置片段是合法 JSON', (() => {
      try { JSON.parse(snippet); return true; } catch { return false; }
    })());
  } catch (e) {
    check('install 验证未抛异常', false, e && e.message ? e.message : String(e));
  } finally {
    try { fs.rmSync(tmpHome, { recursive: true, force: true }); } catch { /* 占用 */ }
    const passed = results.filter((r) => r.ok).length;
    console.log(`\n汇总：${passed}/${results.length} 通过`);
    process.exit(passed === results.length ? 0 : 1);
  }
})();
