/**
 * B5T4 双模型检测（新增文件，11.x 异步基线不动）
 *
 * 任务书 §三.3：boot 时探测同步门面 ——
 *   typeof Application !== 'undefined' → 12.x，走 js/actions-sync/ 同步实现；
 *   探测不到 → 11.x，保留原有 js/actions/*.js 的 jsAsynCall 异步实现（兜底）。
 * 两套共存，运行时选路：apply(type) 只把当前型的动作注册表替换为同步版。
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

  function detect() {
    var app = getApp();
    if (!app) return 'async';
    // 再确认门面可读：ActiveWorkbook / ActiveDocument / ActivePresentation 至少一个存在
    try {
      if ('ActiveWorkbook' in app || 'ActiveDocument' in app || 'ActivePresentation' in app) return 'sync';
    } catch (e) {}
    return 'async';
  }

  var MAP = { wps: 'Word', et: 'Excel', wpp: 'Ppt' };

  function apply(type) {
    var mode = detect();
    global.WpsAddonModel = mode;
    if (mode === 'sync') {
      var k = MAP[type];
      var sync = global['WpsAddonActionsSync' + k];
      if (sync) global['WpsAddonActions' + k] = sync;   // 覆盖注册表，dispatch 运行时查表
    }
    try {
      var x = new XMLHttpRequest();
      x.open('POST', 'http://127.0.0.1:27740/env', true);
      x.setRequestHeader('Content-Type', 'application/json');
      x.send(JSON.stringify({ type: type, phase: 'modelDetect', model: mode,
        Application: typeof global.Application, wps: typeof global.wps,
        actionCount: mode === 'sync' && global['WpsAddonActions' + MAP[type]]
          ? Object.keys(global['WpsAddonActions' + MAP[type]]).length : 0 }));
    } catch (e) {}
    return mode;
  }

  global.WpsAddonModelDetect = { detect: detect, apply: apply };
})(typeof window !== 'undefined' ? window : this);
