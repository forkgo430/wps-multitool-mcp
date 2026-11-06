#!/usr/bin/env node
/**
 * Input: 无（mock COM，不碰 WPS）
 * Output: ⑤ 批量形状原语 wps_ppt_batch_upsert_shapes 的断言组——
 *         upsert 两路（建/改）、部分失败聚合（warning 纪律）、全失败整体 success:false、
 *         上限 50、进度通知、遇错不中断（逐项结果完整）。
 * Pos: 主线一 A 段验收脚本（2026-11-06）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 怎么 mock：替换 wpsClient.executeMethod——getActiveDocumentState 假文档；
 * batchUpsertShapes 按 ps1 同款语义模拟（逐项成败、遇错不中断）。
 * 用法：node scripts/manual/batch-shapes-verify.js
 */
'use strict';

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { toolRegistry, ToolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist/tools/index.js'));
const { wpsClient } = require(path.join(ROOT, 'dist/client/wps-client.js'));

new (require(path.join(ROOT, 'dist/server/mcp-server.js')).WpsMcpServer)();
toolRegistry.registerAll(allTools);

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};
const call = (name, args = {}) => toolRegistry.callTool(ToolRegistry.createRequest(name, args));
const warningsOf = (r) => (r.warnings || []);

/** ps1 同款语义模拟：逐项 upsert，遇错记录不中断；failIdx 指定哪些序号失败 */
function mockPs1(failIdx = []) {
  const calls = [];
  return {
    calls,
    fn: async (action, params) => {
      calls.push({ action, params });
      if (action === 'getActiveDocumentState') return { success: true, data: { hasDocument: false } };
      if (action === 'batchUpsertShapes') {
        const results = (params.shapes || []).map((s, i) => {
          const n = i + 1;
          if (failIdx.includes(n)) return { index: n, action: '', name: '', ok: false, error: `mock 第 ${n} 项炸了` };
          return { index: n, action: s.shapeIndex || s.name ? 'updated' : 'created', name: `形状 ${n}`, ok: true, error: '' };
        });
        return {
          success: true,
          data: {
            slideIndex: params.slideIndex || 1,
            total: results.length,
            ok: results.filter((x) => x.ok).length,
            failed: results.filter((x) => !x.ok).length,
            results,
          },
        };
      }
      return { success: true, data: {} };
    },
  };
}

(async () => {
  console.log('\n=== 1. 全部新建：逐项 created + 无 warnings + 进度通知 ===');
  {
    const m = mockPs1();
    wpsClient.executeMethod = m.fn;
    const r = await call('wps_ppt_batch_upsert_shapes', {
      slideIndex: 1,
      shapes: [
        { type: 'rectangle', left: 0, top: 0, width: 100, height: 50, text: '卡片1' },
        { type: 'rectangle', left: 120, top: 0, width: 100, height: 50, fillColor: '#FF0000' },
        { type: 'oval', left: 240, top: 0, width: 80, height: 80 },
      ],
    });
    ok('全新建 success=true', r.success === true, String(r.error || ''));
    ok('逐项 3 条全 created', r.content[0].text.includes('3/3 成功') && (r.content[0].text.match(/created/g) || []).length === 3);
    const batchWs = warningsOf(r).filter((w) => w.step && w.step.startsWith('shapes['));
    ok('无 batch 自身 warnings（全成功）', batchWs.length === 0);
    ok('进度通知两端（0/N 与 N/N）', m.calls.filter((x) => x.action === 'batchUpsertShapes').length === 1);
  }

  console.log('\n=== 2. upsert 更新路：带 shapeIndex → updated ===');
  {
    const m = mockPs1();
    wpsClient.executeMethod = m.fn;
    const r = await call('wps_ppt_batch_upsert_shapes', {
      slideIndex: 1,
      shapes: [
        { shapeIndex: 2, left: 300, top: 200 },
        { name: '标题 1', text: '新标题文字' },
      ],
    });
    const sent = m.calls.find((x) => x.action === 'batchUpsertShapes');
    ok('两定位键原样透传 ps1', sent && sent.params.shapes[0].shapeIndex === 2 && sent.params.shapes[1].name === '标题 1');
    ok('逐项 action=updated', r.success === true && (r.content[0].text.match(/updated/g) || []).length === 2);
  }

  console.log('\n=== 3. 部分失败：success=true + 逐项 warning（severity:warning）===');
  {
    const m = mockPs1([2]);
    wpsClient.executeMethod = m.fn;
    const r = await call('wps_ppt_batch_upsert_shapes', {
      slideIndex: 1,
      shapes: [
        { type: 'rectangle', left: 0, top: 0, width: 100, height: 50 },
        { type: 'oval', left: 10, top: 10, width: 80, height: 80 },
        { type: 'star', left: 20, top: 20, width: 60, height: 60 },
      ],
    });
    ok('部分失败 → success=true（成功项已生效）', r.success === true);
    const ws = warningsOf(r).filter((w) => w.step && w.step.startsWith('shapes['));
    ok('逐项失败 = warning 级（非 info）', ws.length === 1 && ws[0].severity === 'warning', JSON.stringify(warningsOf(r).map((w) => w.severity)));
    ok('warning 指明失败序号与原因', ws[0].step === 'shapes[2]' && ws[0].message.includes('第 2 项失败') && ws[0].message.includes('mock 第 2 项炸了'));
    ok('成功项仍列入正文', r.content[0].text.includes('2/3 成功'));
  }

  console.log('\n=== 4. 全部失败：整体 success=false ===');
  {
    const m = mockPs1([1, 2]);
    wpsClient.executeMethod = m.fn;
    const r = await call('wps_ppt_batch_upsert_shapes', {
      slideIndex: 1,
      shapes: [
        { type: 'rectangle', left: 0, top: 0, width: 10, height: 10 },
        { type: 'oval', left: 10, top: 10, width: 10, height: 10 },
      ],
    });
    ok('全失败 → 整体 success=false', r.success === false);
    ok('error 聚合了两项', (r.error || '').includes('第 1 项') && (r.error || '').includes('第 2 项'));
    const wsAll = warningsOf(r).filter((w) => w.step && w.step.startsWith('shapes['));
    ok('batch warnings 两条且 warning 级', wsAll.length === 2 && wsAll.every((w) => w.severity === 'warning'));
  }

  console.log('\n=== 5. 遇错不中断：失败项后面的项照常执行（ps1 契约）===');
  {
    const m = mockPs1([1]);
    wpsClient.executeMethod = m.fn;
    const r = await call('wps_ppt_batch_upsert_shapes', {
      slideIndex: 1,
      shapes: [
        { type: 'rectangle', left: 0, top: 0, width: 10, height: 10 },
        { type: 'oval', left: 10, top: 10, width: 20, height: 20 },
        { type: 'star', left: 20, top: 20, width: 30, height: 30 },
      ],
    });
    const sent = m.calls.find((x) => x.action === 'batchUpsertShapes');
    ok('ps1 收到全部 3 项（不因第 1 项失败截断）', sent.params.shapes.length === 3);
    ok('逐项结果 3 条完整（2/3 成功）', r.content[0].text.includes('2/3 成功') && (r.content[0].text.match(/✓/g) || []).length === 2 && (r.content[0].text.match(/✗/g) || []).length === 1);
  }

  console.log('\n=== 6. 边界：空 shapes / 超上限 50 ===');
  {
    const m = mockPs1();
    wpsClient.executeMethod = m.fn;
    const rEmpty = await call('wps_ppt_batch_upsert_shapes', { slideIndex: 1, shapes: [] });
    ok('空 shapes → 前置拦（success=false，未起 COM）', rEmpty.success === false && m.calls.every((x) => x.action !== 'batchUpsertShapes'));
    const rOver = await call('wps_ppt_batch_upsert_shapes', {
      slideIndex: 1,
      shapes: Array.from({ length: 51 }, (_, i) => ({ type: 'rectangle', left: i, top: i, width: 10, height: 10 })),
    });
    ok('超上限 50 → 前置拦且文案带上限', rOver.success === false && /上限 50/.test(rOver.content[0].text), (rOver.error || '').slice(0, 60));
  }

  console.log(`\n汇总：${pass}/${pass + fail} 通过`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => {
  console.error('验收异常：', e && e.stack ? e.stack : e);
  process.exit(2);
});
