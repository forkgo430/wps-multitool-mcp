/**
 * 阶段 4（A/B/C/D）真机验证
 *
 * 背景：之前撞到过 ActiveDocument 返回"非空但空壳"、以及 RPC_E_CALL_REJECTED，
 * 导致原有 word 工具也一起中招。本轮做了四件事，逐条验：
 *
 *   A) 不再依赖 ActiveDocument —— 改走 Resolve-Doc/Resolve-Book
 *      （显式 documentName → 健康检查过的 Active → 遍历集合兜底）
 *   B) 健康检查 Test-DocHealthy —— 不健康不执行半截动作，明确报 document_not_found
 *   C) 模态对话框探测 —— Get-CallError 在 RPC_E_CALL_REJECTED 时附带"该关哪个窗口"
 *   D) wps_common_repair —— 诊断 + 只关无名残留文档
 *
 * ⚠️ 本脚本会【真的打开一个 WPS 文字文档】（用 openDocument 打开本脚本自己造的临时 .txt）。
 *    结束时在 finally 里关掉，不留残留。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const ROOT = path.resolve(__dirname, '../..');
process.env.WPS_BRIDGE_PIPE = '\\\\.\\pipe\\wps-multitool-mcp-repairtest';
process.env.WPS_BRIDGE_IDLE_MS = '30000';
process.env.WPS_DAEMON_TIMEOUT = process.env.WPS_DAEMON_TIMEOUT || '90000';

const { wpsClient, setWordTarget } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
const { WpsAppType } = require(path.join(ROOT, 'dist', 'types', 'wps.js'));
const { classifyWpsError, WpsErrorType: T } = require(path.join(ROOT, 'dist', 'utils', 'error.js'));
const { psDaemon } = require(path.join(ROOT, 'dist', 'client', 'ps-daemon.js'));

const results = [];
function check(name, ok, detail = '') {
  results.push({ name, ok });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  → ' + detail : ''}`);
}

const TMP = path.join(os.tmpdir(), `wps_repair_test_${process.pid}.txt`);

async function call(action, params = {}, app = WpsAppType.WRITER) {
  return wpsClient.executeMethod(action, params, app);
}

async function main() {
  // 造一个临时文档（纯文本，WPS 文字能直接打开）
  fs.writeFileSync(TMP, '阶段4验证用临时文档。\n第二行。\n', 'utf8');
  console.log(`临时文档：${TMP}\n`);

  console.log('=== D) wpsRepair 诊断（WPS 未开时）===');
  const r0 = await call('wpsRepair', {}, undefined);
  check('wpsRepair 在 WPS 未运行时也能返回（不报错）', r0.success === true, JSON.stringify(r0.data?.apps?.word ?? {}).slice(0, 80));

  console.log('\n=== 打开文档，为 A/B/C/D 造出真实场景 ===');
  const open = await call('openDocument', { path: TMP });
  check('openDocument 成功', open.success === true, String(open.error ?? '').slice(0, 80));

  // 等 WPS 把文档真正载入
  await new Promise((r) => setTimeout(r, 3500));

  console.log('\n=== D) 打开后的 wpsRepair 诊断 ===');
  const r1 = await call('wpsRepair', {}, undefined);
  const wordState = r1.data?.apps?.word ?? {};
  console.log('  word 状态：', JSON.stringify(wordState));
  check('诊断到 word 在运行', wordState.running === true);
  check('诊断到文档数 ≥ 1', (wordState.count ?? 0) >= 1, `count=${wordState.count}`);
  const docs = wordState.documents ?? [];
  check(
    '逐个文档报告了健康度',
    docs.length > 0 && docs.every((d) => typeof d.healthy === 'boolean'),
    JSON.stringify(docs)
  );

  console.log('\n=== A) 不依赖 ActiveDocument ===');
  // A-1：不指定 documentName，走"健康检查过的 Active → 遍历兜底"
  const txt = await call('getDocumentText', {});
  check('不指定目标也能取到正文（A 的兜底路径生效）', txt.success === true && typeof txt.data?.text === 'string' && txt.data.text.length > 0, `len=${txt.data?.text?.length}`);

  // A-2：显式指定 documentName（唯一不依赖焦点的方式）
  const base = path.basename(TMP);
  const byName = await call('getDocumentText', { documentName: base });
  check('按 documentName 显式指定能定位', byName.success === true, String(byName.error ?? '').slice(0, 70));

  const byStem = await call('getDocumentText', { documentName: base.replace(/\.txt$/i, '') });
  check('documentName 支持不带扩展名（文件名 stem）', byStem.success === true, String(byStem.error ?? '').slice(0, 70));

  const byFull = await call('getDocumentText', { documentName: TMP });
  check('documentName 支持完整路径', byFull.success === true, String(byFull.error ?? '').slice(0, 70));

  // A-3：指定一个不存在的文档 → 必须明确失败，且【不许】静默回退到活动文档
  const bogus = await call('getDocumentText', { documentName: '__不存在的文档__.docx' });
  check('指定不存在的文档会失败（不静默回退到别的文档）', bogus.success === false, `error=${String(bogus.error).slice(0, 60)}`);
  check(
    '该失败被分类为 document_not_found',
    [T.DOCUMENT_NOT_FOUND, T.UNSUPPORTED_OPERATION].includes(classifyWpsError(String(bogus.error))),
    `classify=${classifyWpsError(String(bogus.error))}`
  );

  console.log('\n=== 目标锁定（switch_document 会设 documentName）===');
  setWordTarget(base);
  const locked = await call('getDocumentText', {});
  check('锁定目标后仍能正确定位', locked.success === true, String(locked.error ?? '').slice(0, 70));
  setWordTarget(undefined); // 解锁，避免影响收尾

  console.log('\n=== C) 模态对话框探测 ===');
  const probe = await call('__probeDialogs', {}, undefined);
  const dlgs = probe.data?.dialogs ?? [];
  console.log('  探测到：', JSON.stringify(dlgs), ' hint=', String(probe.data?.hint ?? 'null').slice(0, 70));
  check('对话框探测可用（无对话框时返回空数组）', probe.success === true && Array.isArray(dlgs), `count=${dlgs.length}`);
  // 关键回归：第一版用"可见+有标题+不占满屏"当判据，把 WPS 主窗口（类名 OpusApp）
  // 误判成了对话框 —— 那会让每次 COM 被拒都附一句错误指令。必须只认对话框类名。
  check(
    '文档窗口【没有】被误判成对话框（判据只看类名）',
    !dlgs.some((d) => /OpusApp|—\s*WPS Office\s*\[/i.test(d)),
    JSON.stringify(dlgs)
  );
  check('主窗口开着时 hint 应为 null（没有真对话框）', probe.data?.hint === null || probe.data?.hint === undefined, String(probe.data?.hint).slice(0, 60));

  console.log('\n=== 写动作验证（A/B 改动后仍正常）===');
  const ins = await call('insertText', { text: '阶段4插入验证', position: 0 });
  check('insertText 仍能写入（A/B 改造没破坏写路径）', ins.success === true, String(ins.error ?? '').slice(0, 70));
  const after = await call('getDocumentText', {});
  check('写入内容可读回', String(after.data?.text ?? '').includes('阶段4插入验证'));

  const stats = await call('getDocumentStats', {});
  check('getDocumentStats 可用', stats.success === true, String(stats.error ?? '').slice(0, 70));

  const passed = results.filter((r) => r.ok).length;
  console.log(`\n汇总：${passed}/${results.length} 通过`);
  console.log(
    '\n⚠️ 验证边界：\n' +
      '   - C 的"有对话框时会不会带上提示"【没有真机验证】——要造出模态框（另存为确认等）才行。\n' +
      '     本轮只验了"无对话框时探测不报错"，Get-CallError 的拼接逻辑是按代码推断。\n' +
      '   - B 的"空壳文档"分支没被真机触发（要 WPS 恰好返回空壳才行），\n' +
      '     Test-DocHealthy 的判据（Name 非空 + Content 可取）是按文档语义写的。'
  );
  process.exitCode = passed === results.length ? 0 : 1;
}

main()
  .catch((e) => {
    console.error('测试异常:', e && e.message ? e.message : e);
    process.exitCode = 1;
  })
  .finally(async () => {
    // 收尾：关掉本脚本打开的文档 + 删临时文件，不留残留（阶段 4 的 D 项纪律）
    try {
      await call('closeDocument', { name: path.basename(TMP), save: false }, undefined);
      console.log('已关闭测试文档');
    } catch {
      console.log('关闭文档失败（可能本来就没打开）');
    }
    try {
      if (fs.existsSync(TMP)) fs.unlinkSync(TMP);
      console.log('已删除临时文件');
    } catch {
      /* 忽略 */
    }
    psDaemon.dispose();
    setTimeout(() => process.exit(process.exitCode || 0), 800);
  });