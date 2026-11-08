/**
 * P3 · 孤儿 {} 块清理（提质小卡 2026-11-08）
 *
 * 判据：handleToolCall 的裸块作用域（无害死结构）已清——函数签名后不再有独立 `{` 行，
 * 函数闭合括号层级归位（tsc 0 错为前置条件）。
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${detail ? '：' + detail : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${detail ? '：' + detail : ''}`); }
};

(async () => {
  console.log('=== 孤儿 {} 块清理 ===');
  const src = fs.readFileSync(path.join(ROOT, 'src', 'server', 'mcp-server.ts'), 'utf8');

  // 1) 函数签名后不再紧跟裸块 `{`
  const sig = src.indexOf('): Promise<{ content: ToolContent[]; isError: boolean }> {');
  ok('handleToolCall 签名存在', sig > 0);
  const afterSig = src.slice(sig, sig + 200);
  ok('签名后无裸块 `{`（孤儿块已清）', !/\{\s*\n\s*\{\s*\n\s*\/\/ 检查Tool/.test(afterSig) && /\/\/ 检查Tool/.test(afterSig));

  // 2) 层级归位：catch 收尾后直接函数闭（原裸块闭 `    }` 已删）
  const tail = src.slice(src.indexOf('注册内置Tools') - 300, src.indexOf('注册内置Tools'));
  ok('函数闭合层级归位（无中间裸块闭）', /      \}\n  \}\n\n  \/\*\*/.test(tail.replace(/\r\n/g, '\n')));

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
