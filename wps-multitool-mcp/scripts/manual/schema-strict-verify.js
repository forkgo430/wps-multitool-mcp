/**
 * Pos: 入参真校验门禁（代码技法卡 A 段）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 * 断言面：类型错/枚举错/缺 required/数组元素错/范围错/批量上限错/非隐式转换/
 *        合法透传逐字节/错误文案形态/全工具 schema 可校验性扫描/注册表集成。
 * 全 mock：validateToolArgs 纯函数 + 注册表集成路径（校验失败不进 handler，零 COM）。
 */
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { validateToolArgs } = require(path.join(ROOT, 'dist', 'server', 'schema-validate.js'));
const { toolRegistry, ToolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));

const results = [];
const check = (name, cond, info = '') => {
  results.push({ name, ok: !!cond });
  console.log(`  ${cond ? '✓' : '✗'} ${name}${info ? ` — ${info}` : ''}`);
};

(async () => {
  // 真注册路径（gen-doc-stats 同款口径）
  new WpsMcpServer();
  const defOf = (name) => toolRegistry.getTool(name)?.definition;

  // ---- 1. 类型错：number 参数收 string
  {
    const def = defOf('wps_excel_set_cell_value');
    const err = validateToolArgs(def, { sheet: 'S', row: '1', col: 1, value: 'x' });
    check('类型错被拒（number 参数收 string）', !!err && err.includes('应为 number') && err.includes('row'), err?.slice(0, 80));
  }

  // ---- 2. 非隐式转换："3" 不静默转 number
  {
    const def = defOf('wps_excel_get_cell_value');
    const err = validateToolArgs(def, { sheet: 'S', row: '3', col: 1 });
    check('非隐式转换（"3" 给 number 参数 = 报错）', !!err && /应为 number/.test(err), err?.slice(0, 80));
  }

  // ---- 3. 枚举错
  {
    const def = defOf('wps_execute_method');
    const err = validateToolArgs(def, { method: 'ping', appType: 'word' });
    check('枚举错被拒（appType 收 word）', !!err && /应为枚举值之一/.test(err) && err.includes('appType'), err?.slice(0, 80));
  }

  // ---- 4. 缺 required + 文案形态
  {
    const def = defOf('wps_excel_set_cell_value');
    const err = validateToolArgs(def, { sheet: 'S', row: 1 });
    check('缺 required 被拒', !!err && err.includes('缺少必填参数') && err.includes('col'), err?.slice(0, 80));
    check('错误文案形态（[勿重试·先改参数] 前缀）', !!err && err.startsWith('[勿重试·先改参数]'), err?.slice(0, 40));
  }

  // ---- 5. 数组元素类型错（一层嵌套）
  {
    const def = defOf('wps_ppt_batch_upsert_shapes');
    const err = validateToolArgs(def, { slideIndex: 1, shapes: 'not-array' });
    check('数组参数收标量被拒', !!err && /应为 array/.test(err), err?.slice(0, 80));
  }

  // ---- 6. minItems 批量上限（wps_common_batch steps ≥ 1）
  {
    const def = defOf('wps_common_batch');
    const err = validateToolArgs(def, { steps: [] });
    check('minItems 下限被拒（steps ≥ 1）', !!err && /至少需要/.test(err), err?.slice(0, 80));
  }

  // ---- 7. 数值范围（minimum/maximum 子集支持——合成 schema 验证校验器能力）
  {
    const fakeDef = {
      name: 'wps_hygiene_range_probe',
      inputSchema: {
        type: 'object',
        properties: { alpha: { type: 'number', minimum: 1, maximum: 100 } },
        required: [],
      },
    };
    const hi = validateToolArgs(fakeDef, { alpha: 101 });
    const lo = validateToolArgs(fakeDef, { alpha: 0 });
    const pass = validateToolArgs(fakeDef, { alpha: 50 });
    check('数值范围（maximum 超界被拒）', !!hi && /应 ≤ 100/.test(hi), hi?.slice(0, 60));
    check('数值范围（minimum 越界被拒）', !!lo && /应 ≥ 1/.test(lo), lo?.slice(0, 60));
    check('数值范围（界内通过）', pass === null, String(pass));
  }

  // ---- 8. 合法透传零拦截（含二维数组、target 对象、可选缺省）
  {
    const def = defOf('wps_excel_write_range');
    const pass = validateToolArgs(def, {
      range: 'A1:B2',
      data: [['姓名', '年龄'], ['张三', 25]],
      sheet: 'Sheet1',
      target: { instance: 'sandbox', doc: 'x.xlsx' },
    });
    check('合法透传（二维数组+target 对象+可选参数）零拦截', pass === null, String(pass));
    const def2 = defOf('wps_word_insert_text');
    const pass2 = validateToolArgs(def2, { text: 'hello' });
    check('合法透传（word 单参数）零拦截', pass2 === null, String(pass2));
  }

  // ---- 9. 全工具 schema 可校验性扫描（287 工具的已声明参数都有可识别 type）
  {
    let tools = 0, params = 0;
    const badTypes = [];
    for (const definition of toolRegistry.listAllTools()) {
      tools++;
      const props = definition.inputSchema.properties || {};
      for (const [k, p] of Object.entries(props)) {
        params++;
        const known = ['string', 'number', 'boolean', 'object', 'array'];
        if (!known.includes(p.type)) badTypes.push(`${definition.name}.${k}(${p.type})`);
      }
    }
    check(`全工具 schema 可校验性扫描（${tools} 工具 ${params} 参数无未知 type）`, tools === 287 && badTypes.length === 0, badTypes.join(', ').slice(0, 80));
  }

  // ---- 10. 注册表集成：校验失败不进 handler（零 COM 路径）
  {
    const r = await toolRegistry.callTool(ToolRegistry.createRequest('wps_excel_set_cell_value', { sheet: 'S', row: 'bad', col: 1, value: 'x' }));
    const text = r.content?.[0]?.text || '';
    check('注册表集成（非法入参当场回，不进 COM）', r.success === false && text.startsWith('[勿重试·先改参数]'), text.slice(0, 70));
  }

  const passed = results.filter((x) => x.ok).length;
  console.log(`汇总：${passed}/${results.length} 通过`);
  if (passed !== results.length) process.exit(1);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
