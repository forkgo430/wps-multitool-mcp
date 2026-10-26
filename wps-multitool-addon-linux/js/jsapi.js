/**
 * Input: WPS JSAPI 方法名（method）与参数对象
 * Output: Promise<object>——原生侧回调带回的结果（JSON 解析后）
 * Pos: Linux 加载项通道的 JSAPI 桥（B4：异步 jsAsynCall 模型）。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 为什么是这个模型（B3-REPORT.md §3.3，报告行 96-115）
 * ------------------------------------------------
 * B3 真机探针：在线 jsplugin 模式的 webview 里 typeof wps === "undefined"、
 * 没有 Application，唯一的原生桥是 CEF 的 cefQuery（+cefQueryCancel）。
 * WPS 自有插件（office6/addons/kbarcode/…/kwebpack_api_invoker.js）的调用范式：
 *
 *   global.cefQuery({ request: "jsAsynCall('" + JSON.stringify({method, callback, …params}) + "')" });
 *   // 结果由原生侧回调全局 jsCallbackN(result)
 *
 * B2.5 的同步 COM 门面（app.Documents.Open() 当返回值用）在这个宿主里无对象可调——
 * 这是模型错配，不是漏引脚本。B4 全部动作改走本模块。
 *
 * 设计（B4 任务卡 §1.1）
 * ------------------
 *   - 回调名：'jsCallback' + 自增序号，挂 global；用完 delete 清理，防全局泄漏；
 *   - 超时：WPS_JSAPI_TIMEOUT（默认 30000ms）→ reject('jsAsynCall 超时：method')，
 *     reject 时同样清理回调与 cefQueryCancel 句柄；
 *   - 并发：每次调用独立回调名 + 独立计时器，回调按名取回，互不串线；
 *   - 结果解析：原生回传是 JSON 字符串，解析失败时把原文包在 { raw } 里如实上抛，不吞。
 */

(function (global) {
  'use strict';

  function timeoutMs() {
    var v = parseInt(process && process.env ? process.env.WPS_JSAPI_TIMEOUT : '', 10);
    return Number.isFinite(v) && v > 0 ? v : 30000;
  }

  var seq = 0;

  /**
   * 发起一次 jsAsynCall。
   * @param {string} method  JSAPI 方法名（与 S1-REPORT §3"对应的 JSAPI 操作"列对应，出处见各动作头注）
   * @param {object} params  参数（会与 method/callback 一起序列化进 request）
   * @returns {Promise<object>}
   */
  function jsAsynCall(method, params) {
    return new Promise(function (resolve, reject) {
      if (typeof global.cefQuery !== 'function') {
        reject(new Error('jsAsynCall 不可用：宿主无 cefQuery（B3 §3.3：在线模式 webview 只有 CEF 原生桥；本调用必须在 WPS 加载项 webview 内发起）'));
        return;
      }
      seq += 1;
      var callbackName = 'jsCallback' + seq;
      var settled = false;
      var timer = setTimeout(function () {
        if (settled) return;
        settled = true;
        cleanup();
        reject(new Error('jsAsynCall 超时（' + timeoutMs() + 'ms）：' + method));
      }, timeoutMs());

      function cleanup() {
        try { delete global[callbackName]; } catch (e) { global[callbackName] = undefined; }
      }

      global[callbackName] = function (result) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        cleanup();
        var parsed;
        try {
          parsed = JSON.parse(result);
        } catch (e) {
          // 原文不吞：解析失败把原文带回去，交给上层判断（B3 没有回包样例，格式以真机为准）
          parsed = { raw: result === undefined ? null : String(result), parseError: String(e && e.message ? e.message : e) };
        }
        resolve(parsed);
      };

      var payload = { method: method, callback: callbackName };
      if (params) {
        for (var k in params) {
          if (Object.prototype.hasOwnProperty.call(params, k)) payload[k] = params[k];
        }
      }
      try {
        global.cefQuery({
          request: "jsAsynCall('" + JSON.stringify(payload) + "')",
          onSuccess: function (r) { /* 主通道是全局回调；onSuccess 只作旁证，不重复 settle */ },
          onFailure: function (err) {
            if (settled) return;
            settled = true;
            clearTimeout(timer);
            cleanup();
            reject(new Error('jsAsynCall 失败（cefQuery onFailure）：' + method + ' —— ' + String(err)));
          },
        });
      } catch (e) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        cleanup();
        reject(new Error('jsAsynCall 发送失败：' + method + ' —— ' + String(e && e.message ? e.message : e)));
      }
    });
  }

  /** 测试/诊断用：当前全局残留的 jsCallback 名（应为空数组——回调用完即清） */
  function pendingCallbacks() {
    var out = [];
    for (var k in global) {
      if (Object.prototype.hasOwnProperty.call(global, k) && /^jsCallback\d+$/.test(k)) out.push(k);
    }
    return out;
  }

  var jsapi = { jsAsynCall: jsAsynCall, pendingCallbacks: pendingCallbacks, _timeoutMs: timeoutMs };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = jsapi;
  } else {
    global.WpsAddonJsapi = jsapi;
  }
})(typeof window !== 'undefined' ? window : globalThis);
