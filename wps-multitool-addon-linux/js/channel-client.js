/**
 * Input: 服务端地址与令牌 —— 优先 URL query（?server=http://127.0.0.1:27740&token=…），
 *        其次 window.WPS_ADDON_CONFIG（index.html 里内联的对象），最后默认值。
 * Output: window.WpsAddonChannel = { start(type), dispatch(action, params, app), state() }
 * Pos: 三型加载项（wps/et/wpp）共用的通道客户端（B2 骨架）。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 协议（与 Node 侧 addon-channel.ts 对齐，帧同构 Windows 管道协议 ps-daemon.ts:150）
 * --------------------------------------------------------------------------
 *   POST /handshake  { type, token, addonVersion }          → { sessionId, pollIntervalMs }
 *   GET  /poll?sessionId=…                                  → { id, action, params, app } | { idle: true }
 *   POST /result     { sessionId, id, ok, data?, error? }   → { received: true }
 *
 * 循环：握手一次 → poll（服务端无任务挂起 ≤10s，到点回 idle）→ 取到帧就 dispatch →
 *       回结果 → 立即再 poll。断线/404（会话被覆盖或失效）→ 退避 1s 重新握手。
 *
 * 动作分发器是 B2 骨架：只有 ping / wireCheck 真实现（直接回显，不碰 JSAPI），
 * 其余动作一律回 { ok:false, unsupported:'B3' } —— B3 才在加载项环境里填 JSAPI 实现
 * （S1 覆盖率：支持 199 / 部分支持 45 / 抛错 10 / 不支持 43，实现顺序按 B4 结论）。
 *
 * 已知风险（照计划 §4.6 记录，不装不知道）：
 *   - 定时器节流（R6）：WPS 最小化/省电时本循环的 setTimeout 可能被节流到分钟级，
 *     表现为"不报错只是慢"。上线前按 B4 结论评估换 WebSocket（同一套协议帧，动作层不改）。
 *   - CORS：Node 侧执行通道刻意不做 CORS（C2）。WPS jsaddons 的 webview 若强制走
 *     浏览器同源策略，XHR 会被拦 —— B3 真机联调第一件事就是验证这条；被拦时的出路是
 *     jsaddons 宿主提供的特权请求通道，不是给服务端开 CORS。
 */
(function (global) {
  'use strict';

  function queryParam(name) {
    try {
      var m = new RegExp('[?&]' + name + '=([^&]*)').exec(String(global.location && global.location.search) || '');
      return m ? decodeURIComponent(m[1]) : '';
    } catch (e) {
      return '';
    }
  }

  var cfg = global.WPS_ADDON_CONFIG || {};
  var SERVER = queryParam('server') || cfg.server || 'http://127.0.0.1:27740';
  var TOKEN = queryParam('token') || cfg.token || '';
  var ADDON_VERSION = '0.1.0';
  var RECONNECT_MS = 1000;

  var state = {
    type: '',
    sessionId: '',
    connected: false,
    lastError: '',
    handled: 0,
    startedAt: '',
  };

  function log(msg) {
    try {
      if (global.console && global.console.log) global.console.log('[addon-channel-client] ' + msg);
    } catch (e) { /* 宿主可能没 console */ }
  }

  /* ---- 底层请求：XHR，webview 里 fetch 的可用性不如 XHR 稳（B3 真机复核） ---- */
  function xhr(method, path, body, cb) {
    try {
      var x = new global.XMLHttpRequest();
      x.open(method, SERVER + path, true);
      x.timeout = 30000; // 服务端长轮询挂 10s，客户端超时给足余量
      x.onreadystatechange = function () {
        if (x.readyState !== 4) return;
        if (x.status < 200 || x.status >= 300) {
          cb(null, { status: x.status, text: String(x.responseText || '') });
          return;
        }
        var parsed = null;
        try {
          parsed = JSON.parse(x.responseText);
        } catch (e) {
          cb(null, { status: x.status, text: '非 JSON 响应: ' + String(x.responseText).slice(0, 80) });
          return;
        }
        cb(parsed, null);
      };
      x.onerror = function () {
        cb(null, { status: 0, text: '网络错误（服务端没起/端口不通/CORS 被拦）' });
      };
      x.ontimeout = function () {
        cb(null, { status: 0, text: '请求超时（30s）' });
      };
      if (body !== null && body !== undefined) {
        x.setRequestHeader('Content-Type', 'application/json; charset=utf-8');
        x.send(JSON.stringify(body));
      } else {
        x.send(null);
      }
    } catch (e) {
      cb(null, { status: 0, text: String(e && e.message ? e.message : e) });
    }
  }

  /* ---- 动作分发器（B2 骨架）----
   * 真实现：ping / wireCheck（回显，不碰 JSAPI —— 通道闭环验证用）。
   * 其余一律 unsupported:'B3'：显式拒绝，绝不静默成功（计划 §6 D3 的精神）。 */
  function dispatch(action, params, app) {
    if (action === 'ping') {
      return { ok: true, data: { pong: true, app: app, type: state.type, addonVersion: ADDON_VERSION, ts: Date.now() } };
    }
    if (action === 'wireCheck') {
      return { ok: true, data: { message: 'WPS MCP 加载项通道已连接（' + state.type + ' 型，B2 骨架）' } };
    }
    return { ok: false, unsupported: 'B3', error: '动作 ' + action + ' 在加载项通道尚未实现（B3 填 JSAPI）' };
  }

  function handshake() {
    xhr('POST', '/handshake', { type: state.type, token: TOKEN, addonVersion: ADDON_VERSION }, function (res, err) {
      if (err || !res || !res.sessionId) {
        state.connected = false;
        state.lastError = err ? err.text : '握手失败：' + JSON.stringify(res);
        log(state.lastError + ' —— ' + RECONNECT_MS + 'ms 后重试');
        setTimeout(handshake, RECONNECT_MS);
        return;
      }
      state.sessionId = res.sessionId;
      state.connected = true;
      state.lastError = '';
      state.startedAt = new Date().toISOString();
      log('握手成功（type=' + state.type + '，sessionId=' + res.sessionId.slice(0, 8) + '…），进入轮询');
      pollOnce();
    });
  }

  function pollOnce() {
    if (!state.sessionId) {
      handshake();
      return;
    }
    xhr('GET', '/poll?sessionId=' + encodeURIComponent(state.sessionId), null, function (frame, err) {
      if (err) {
        // 网络断/超时：退避后重试；连续失败较多时重新握手（会话可能已被服务端清掉）
        state.lastError = err.text;
        setTimeout(function () {
          if (state.connected) pollOnce();
          else handshake();
        }, RECONNECT_MS);
        return;
      }
      if (frame && frame.error) {
        // 404/409：会话失效（被同型覆盖 / 服务端重启）→ 重新握手
        state.connected = false;
        state.sessionId = '';
        state.lastError = String(frame.error);
        setTimeout(handshake, RECONNECT_MS);
        return;
      }
      if (!frame || frame.idle) {
        setTimeout(pollOnce, 0); // 长轮询到点回 idle：立即再来
        return;
      }
      var out;
      try {
        out = dispatch(frame.action, frame.params, frame.app);
      } catch (e) {
        out = { ok: false, error: '分发器异常：' + String(e && e.message ? e.message : e) };
      }
      state.handled += 1;
      var payload = { sessionId: state.sessionId, id: frame.id, ok: out.ok === true };
      if (out.ok) {
        payload.data = out.data;
      } else {
        payload.error = out.error || '动作执行失败';
        if (out.unsupported) payload.unsupported = out.unsupported;
      }
      xhr('POST', '/result', payload, function () {
        setTimeout(pollOnce, 0);
      });
    });
  }

  /** 启动：type 是本加载项在 WPS 里的型别（'wps' | 'et' | 'wpp'），由各型的 boot.js 传入 */
  function start(type) {
    if (['wps', 'et', 'wpp'].indexOf(type) < 0) {
      log('start() 的 type 必须是 wps|et|wpp，收到：' + type);
      return;
    }
    state.type = type;
    handshake();
  }

  global.WpsAddonChannel = {
    start: start,
    dispatch: dispatch,
    state: function () {
      return JSON.parse(JSON.stringify(state));
    },
  };
})(typeof window !== 'undefined' ? window : this);
