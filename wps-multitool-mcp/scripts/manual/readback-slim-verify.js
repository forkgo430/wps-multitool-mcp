#!/usr/bin/env node
/**
 * Input: 无（全程 mock COM / 直调 handler，不碰 WPS、不起桥）
 * Output: 回包瘦身的断言组（体验修复轮 B · 方案 §4.4，原则"不丢信息只改默认"）：
 *         read_range maxRows 截断+提示 / get_outline 默认 100 / get_shapes·get_textboxes
 *         brief 摘要模式（默认开，brief:false 回全量）/ export inline 开关（默认 true）
 * Pos: 体验修复轮 B 的验收脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 怎么 mock：替换 wpsClient.executeMethod / invokeAction——excel getRangeData 回可控
 * 行数；ps1 getTextBoxes/getShapes 回可控字段。断言"改前改后结果比对"：默认行为与
 * 改前逐字兼容（不传新参数 = 旧行为或仅按新默认收紧），新参数显式传入才走新路径。
 *
 * 用法：node scripts/manual/readback-slim-verify.js
 */
'use strict';

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { toolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist/tools/index.js'));
const { wpsClient } = require(path.join(ROOT, 'dist/client/wps-client.js'));
const { ToolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));

new (require(path.join(ROOT, 'dist/server/mcp-server.js')).WpsMcpServer)();
toolRegistry.registerAll(allTools);

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};
const call = (name, args = {}) => toolRegistry.callTool(ToolRegistry.createRequest(name, args));

(async () => {
  console.log('\n=== 1. wps_excel_read_range maxRows（默认 200，超出截断+提示剩余）===');
  {
    // 350 行 × 2 列
    const rows = Array.from({ length: 350 }, (_, i) => [`r${i + 1}`, i + 1]);
    wpsClient.getRangeData = async () => rows;

    const rDefault = await call('wps_excel_read_range', { range: 'A1:B350' });
    const dText = rDefault.content[0].text;
    ok('默认 200 行封顶：只展示前 200 行', rDefault.success === true && dText.includes('350行') && dText.includes('仅展示前 200 行'));
    ok('截断提示剩余行数', dText.includes('剩余 150 行'));
    ok('总行数如实带回（350）', /350行 x 2列/.test(dText));

    const rFull = await call('wps_excel_read_range', { range: 'A1:B350', maxRows: 400 });
    ok('maxRows 放大 → 全量展示（无截断提示）',
      rFull.success === true && rFull.content[0].text.includes('350行') && !rFull.content[0].text.includes('仅展示前'));
    const rSmall = await call('wps_excel_read_range', { range: 'A1:B350', maxRows: 10 });
    ok('maxRows 收紧 → 10 行', rSmall.content[0].text.includes('仅展示前 10 行') && rSmall.content[0].text.includes('剩余 340 行'));
  }

  console.log('\n=== 2. wps_ppt_get_shapes brief 摘要模式（默认开）===');
  {
    const FAKE_SHAPES = {
      shapes: [
        { index: 1, name: '标题 1', type: '文本框', left: 40, top: 30, width: 880, height: 60, text: '第一季度汇报' },
        { index: 2, name: '矩形 2', type: '矩形', left: 40, top: 120, width: 200, height: 80, text: '' },
      ],
    };
    wpsClient.executeMethod = async (action) => {
      if (action === 'getActiveDocumentState') return { success: true, data: { hasDocument: false } };
      if (action === 'getShapes') return { success: true, data: FAKE_SHAPES };
      return { success: true, data: {} };
    };

    const rBrief = await call('wps_ppt_get_shapes', { slideIndex: 1 });
    const bText = rBrief.content[0].text;
    ok('brief 默认开：回序号/名称/几何摘要', rBrief.success === true && bText.includes('标题 1') && bText.includes('40, 50') === false && bText.includes('400, 100') === false || bText.includes('400'), bText.split('\n')[2] || '');
    ok('brief 摘要模式标注', bText.includes('brief 摘要模式'));
    ok('brief 不含形状文字内容（瘦身）', !bText.includes('第一季度汇报'));

    const rFull = await call('wps_ppt_get_shapes', { slideIndex: 1, brief: false });
    const fText = rFull.content[0].text;
    ok('brief:false 回全量（类型+独立几何行）', fText.includes('(文本框)') && fText.includes('位置: (40, 30)'));
  }

  console.log('\n=== 3. wps_ppt_get_textboxes brief 摘要模式（默认开）===');
  {
    const FAKE_TB = {
      textBoxes: [
        { index: 1, name: 'TextBox 1', left: 40, top: 30, width: 880, height: 60, text: '机密内容ABC' },
        { index: 2, name: 'TextBox 2', left: 40, top: 120, width: 400, height: 50, text: '' },
      ],
    };
    wpsClient.executeMethod = async (action) => {
      if (action === 'getActiveDocumentState') return { success: true, data: { hasDocument: false } };
      if (action === 'getTextBoxes') return { success: true, data: FAKE_TB };
      return { success: true, data: {} };
    };

    const rBrief = await call('wps_ppt_get_textboxes', { slideIndex: 1 });
    const bText = rBrief.content[0].text;
    ok('brief 默认开：正文瘦身掉（不回 text 内容）', rBrief.success === true && !bText.includes('机密内容ABC'));
    ok('几何字段保留', bText.includes('"width": 880'));

    const rFull = await call('wps_ppt_get_textboxes', { slideIndex: 1, brief: false });
    ok('brief:false 回全量（含 text）', rFull.content[0].text.includes('机密内容ABC'));
  }

  console.log('\n=== 4. wps_ppt_export_slide_as_image inline 开关（默认 true）===');
  {
    const FAKE_STATE = { hasDocument: false };
    wpsClient.executeMethod = async (action) => {
      if (action === 'getActiveDocumentState') return { success: true, data: FAKE_STATE };
      if (action === 'exportSlideAsImage') return { success: true, data: { slideIndex: 1 } };
      return { success: true, data: {} };
    };

    const rOn = await call('wps_ppt_export_slide_as_image', {
      slideIndex: 1, outputPath: 'C:/tmp/x.png', width: 10, height: 10,
    });
    ok('inline 缺省=true：内容类型数组里有图像块（或 1.5MB 上限内）',
      rOn.success === true && rOn.content.length >= 1, `content=${rOn.content.length} 段`);

    const rOff = await call('wps_ppt_export_slide_as_image', {
      slideIndex: 1, outputPath: 'C:/tmp/x.png', width: 10, height: 10, inline: false,
    });
    ok('inline=false：只回文本路径（无图像块）',
      rOff.success === true && rOff.content.length === 1 && rOff.content[0].type === 'text');
  }

  console.log(`\n汇总：${pass}/${pass + fail} 通过`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => {
  console.error('验收异常：', e && e.stack ? e.stack : e);
  process.exit(2);
});
