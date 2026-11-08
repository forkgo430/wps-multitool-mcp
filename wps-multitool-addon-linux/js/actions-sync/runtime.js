/**
 * B5T4 同步门面动作层 · 运行时（新增文件，不改 11.x 基线）
 *
 * 模型：12.x 的 webview 里 window.Application（或 wps.Application）是 WPS 注入的
 * 「同步 JSAPI 门面」，属性/方法直接读写，返回值即时可得（S1-REPORT §2/§3 实测）。
 * 与 11.x 的 jsAsynCall(cefQuery) 异步桥是两套模型；本文件只服务同步模型。
 *
 * 回包结构照抄包内 actions/*.js 的 guard 形态：Promise<{ok:true,data}|{ok:false,error}>，
 * 异常不吞、原文带回。S1 判「抛错 / 不支持（属性不存在）」的条目显式 unsupported。
 *
 * S1 §4 语义差异适配（18 条里与本层相关的）：
 *   #1 Cells/Rows/Columns/Names 是集合对象，必须 .Item(...)，不可当函数调用
 *   #2 ws.Range('A1:B2') 反而是函数，可直接调用
 *   #3 写单元格 .Value= 静默失效 → .Value2= 或 .Formula=；读取用 .Value()
 *   #4 range.Value 是函数 → range.Value()
 *   #5 sel.Address 是函数 → sel.Address()
 *   #8 文本含裸 \r → 按 \n 规整
 *   #6 ws.PivotTables.Count=undefined（如实回传，不假装 0）
 *   #15 ws.ChartObjects.Add 不存在 → ws.Shapes.AddChart
 */
(function (global) {
  'use strict';

  function getApp() {
    try {
      if (typeof global.Application !== 'undefined' && global.Application) return global.Application;
      if (typeof global.wps !== 'undefined' && global.wps && global.wps.Application) return global.wps.Application;
    } catch (e) {}
    return null;
  }

  function tstr(v) {
    try {
      if (v === null) return 'null';
      if (v === undefined) return 'undefined';
      var t = typeof v, s;
      if (t === 'function') return 'function';
      if (t === 'object') {
        try { s = JSON.stringify(v); } catch (e) { s = String(v); }
      } else s = String(v);
      // S1 §4 #8：裸 \r 会破坏按行解析，规整为 \\n
      s = s.replace(/\r/g, '\\n');
      if (s.length > 200) s = s.slice(0, 200) + '…(截断)';
      return t + ':' + s;
    } catch (e) { return '<unprintable>'; }
  }

  function has(o, k) {
    if (o === null || o === undefined) return false;
    try { return (typeof o === 'object' || typeof o === 'function') && (k in o); }
    catch (e) { return false; }
  }

  /** 取集合项：优先 .Item(i)，兼容少数真是函数的集合 */
  function item(coll, a, b) {
    if (!coll) throw new Error('集合不存在');
    if (typeof coll.Item === 'function') return b === undefined ? coll.Item(a) : coll.Item(a, b);
    if (typeof coll === 'function') return b === undefined ? coll(a) : coll(a, b);
    throw new Error('集合既无 .Item 也不是函数（typeof=' + typeof coll + '）');
  }

  /** 读单元格值：.Value 在 WPS Linux 里是函数（S1 §4 #3/#4） */
  function val(cell) {
    try { if (typeof cell.Value === 'function') return cell.Value(); } catch (e) {}
    try { if (typeof cell.Value2 === 'function') return cell.Value2(); } catch (e) {}
    return cell.Value;
  }

  /** 写单元格值：.Value= 静默失效，优先 .Value2= / .Formula= */
  function setVal(cell, v) {
    try { if ('Value2' in cell) { cell.Value2 = v; return 'Value2'; } } catch (e) {}
    try { cell.Formula = v; return 'Formula'; } catch (e) {}
    try { cell.Value = v; return 'Value'; } catch (e) {}
    throw new Error('单元格写入失败：Value2/Formula/Value 均不可用');
  }

  function addr(sel) {
    try { var a = sel.Address; return typeof a === 'function' ? a.call(sel) : a; }
    catch (e) { return ''; }
  }

  /** 构造上下文：三型对象模型隔离，只暴露当前型能取到的对象 */
  function context(app) {
    var c = { app: app, t: tstr, has: has, item: item, val: val, setVal: setVal, addr: addr };
    try { c.wb = app.ActiveWorkbook; } catch (e) { c.wb = null; }
    try { c.ws = c.wb ? c.wb.ActiveSheet : null; } catch (e) { c.ws = null; }
    try { c.doc = app.ActiveDocument; } catch (e) { c.doc = null; }
    try { c.pres = app.ActivePresentation; } catch (e) { c.pres = null; }
    try { c.sel = app.Selection; } catch (e) { c.sel = null; }
    c.cell = function (r, col) {
      if (!c.ws) throw new Error('无活动工作表');
      return item(c.ws.Cells, r, col);
    };
    c.rng = function (a) {
      if (!c.ws) throw new Error('无活动工作表');
      return c.ws.Range(a);   // S1 §4 #2：Range 是函数
    };
    c.rows = function (i) { if (!c.ws) throw new Error('无活动工作表'); return item(c.ws.Rows, i); };
    c.cols = function (i) { if (!c.ws) throw new Error('无活动工作表'); return item(c.ws.Columns, i); };
    c.needWs = function () { if (!c.ws) throw new Error('无活动工作表（ActiveSheet 为 null）'); return c.ws; };
    c.needDoc = function () { if (!c.doc) throw new Error('无活动文档（ActiveDocument 为 null）'); return c.doc; };
    c.needPres = function () { if (!c.pres) throw new Error('无活动演示（ActivePresentation 为 null）'); return c.pres; };
    c.rangeVal = function (rg) {
      var v = rg.Value;
      return typeof v === 'function' ? v.call(rg) : v;  // S1 §4 #4
    };
    return c;
  }

  /** guard：与包内 actions/*.js 同形态，异常不吞 */
  function guard(impl) {
    return function (params) {
      try { return Promise.resolve(impl(params || {})); }
      catch (e) { return Promise.resolve({ ok: false, error: String(e && e.message ? e.message : e) }); }
    };
  }

  /**
   * 把一条声明式规格变成动作函数。
   * spec: { v: S1判定, jsapi: S1§3操作列, evidence: S1证据, code: 同步门面直调 JS 语句 }
   * code 在 new Function('c','p','R', code) 中执行，c=context, p=params, R=辅助函数集。
   */
  function makeAction(spec) {
    return guard(function (p) {
      var app = getApp();
      if (!app) return { ok: false, error: '同步门面不可用：window.Application / wps.Application 均不存在（应走 11.x 异步兜底）' };
      // 任务书 §三.2：判「抛错 / 不支持」→ 显式 unsupported + S1 原文，不假装成功
      if (spec.v === '抛错' || spec.v === '不支持（属性不存在）' || spec.v === '不支持') {
        return { ok: false, unsupported: 'S1:' + spec.v, error: (spec.evidence || spec.jsapi || 'S1 判定不支持') };
      }
      var c = context(app);
      var R = { t: tstr, has: has, item: item, val: val, setVal: setVal, addr: addr };
      var fn = new Function('c', 'p', 'R', spec.code);
      var out = fn(c, p, R);
      return {
        ok: true,
        data: out === undefined ? null : out,
        verdict: spec.v,
        jsapi: spec.jsapi,
      };
    });
  }

  function build(specs) {
    var reg = {};
    for (var k in specs) {
      if (Object.prototype.hasOwnProperty.call(specs, k)) reg[k] = makeAction(specs[k]);
    }
    return reg;
  }

  global.WpsAddonSyncRuntime = {
    getApp: getApp, context: context, tstr: tstr, has: has, item: item,
    val: val, setVal: setVal, addr: addr, guard: guard, makeAction: makeAction, build: build,
  };
})(typeof window !== 'undefined' ? window : this);
