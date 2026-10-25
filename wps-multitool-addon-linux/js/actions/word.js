/**
 * Input: 宿主 Application 对象（wps 型 = 文字）、动作参数（与 src/tools 的 TS handler 所发参数同名）
 * Output: { ok: true, data } | { ok: false, error }（通道协议回包结构，B2 已定）
 * Pos: Linux 加载项通道的动作层 · wps 型（文字）5 个动作。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 语义依据：S1-REPORT.md §3 逐条明细（每条函数头部标了报告行号）。
 * 实现纪律（B2.5 任务卡）：
 *   - 参数名与 TS 工具层实际发送的键一致（帧里的 params 来自那边，本文件不改名）；
 *   - JSAPI 抛错 → { ok:false, error: 原文 }，不吞、不静默成功（guard() 统一兜）；
 *   - 枚举/颜色换算镜像 scripts/wps-com.ps1 的同名 case（行号在注释里），编号语义 S1 §5 实测与 Windows 恒等。
 * 本文件只在 wps 型加载项里加载（addon-wps）；et/wpp 型没有 Documents（S1 §6 三型隔离）。
 */
(function (global) {
  'use strict';

  /** JSAPI 异常统一兜底：原文带回，不改写、不吞。 */
  function guard(impl, app, params) {
    try {
      return impl(app, params || {});
    } catch (e) {
      return { ok: false, error: String(e && e.message ? e.message : e) };
    }
  }

  /**
   * 十六进制颜色 → COM RGB 整数（R + G*256 + B*65536）。
   * 镜像 scripts/wps-com.ps1:732 Convert-HexColorToRgbInt：支持 #RGB 三位展开；
   * 非法输入返回 null（调用方决定报错还是跳过）。
   * 编码正确性有 S0 实测背书：绿色 #00FF00 → 65280，读回一致（S1-REPORT §2.4 / S0 #18）。
   */
  function hexToRgbInt(hex) {
    if (!hex) return null;
    var s = String(hex).trim();
    if (s.charAt(0) === '#') s = s.slice(1);
    if (s.length === 3) s = s[0] + s[0] + s[1] + s[1] + s[2] + s[2];
    if (!/^[0-9a-fA-F]{6}$/.test(s)) return null;
    var r = parseInt(s.slice(0, 2), 16);
    var g = parseInt(s.slice(2, 4), 16);
    var b = parseInt(s.slice(4, 6), 16);
    return r + g * 256 + b * 65536;
  }

  /** 色名 → COM RGB（镜像 ps1 setFont 的 colorMap；Word 侧 Font.Color 直设整数，不是 .RGB） */
  var COLOR_NAMES = { red: 255, blue: 16711680, green: 65280, black: 0 };

  var wordActions = {

    /**
     * openDocument —— S1-REPORT.md:134（§3 #108 支持）
     * 语义差异（S1 §4 #12，报告行 341）：同文件重复 Open 只激活已有文档、不新建；
     * Open 返回的引用在文档关闭后失活——本动作只在拿到引用后立刻读 Name，不缓存。
     */
    openDocument: function (app, params) {
      return guard(function (a, p) {
        var path = p.path || p.filePath; // TS 两个键都发（word/document.ts:235-239）
        if (!path) {
          return { ok: false, error: '[勿重试·先改参数] 缺少 path/filePath（openDocument 必填）' };
        }
        var doc = a.Documents.Open(path);
        var name = '';
        try { name = String(doc.Name || ''); } catch (e) { /* 引用刚拿到就失活属于异常路径，交给外层语义 */ }
        return { ok: true, data: { documentName: name, name: name } };
      }, app, params);
    },

    /**
     * getActiveDocumentState —— S1-REPORT.md:129（§3 #103 支持）
     * 返回结构与 Windows 侧 ps1:1030 的 DocState 对齐（snapshot.ts 的 getActiveState 消费同名键）。
     * 逐字段防御性读取镜像 ps1：ActiveDocument 是弱引用，单字段读挂不该毁掉整次调用。
     * 注意（S1 §4 #9）：wps 型的 ActiveDocument 跟全局焦点走——"活动"是宿主说了算，不是调用方。
     */
    getActiveDocumentState: function (app, params) {
      return guard(function (a) {
        var doc = a.ActiveDocument;
        var name = '';
        try { name = String(doc.Name || ''); } catch (e) { /* 空壳 */ }
        if (!name) {
          return { ok: false, error: '没有活动文档（wps 型 ActiveDocument 为空壳——弱引用，见 S1 §4 #9）' };
        }
        var fullName = '';
        var dir = '';
        var saved = true;
        var readOnly = false;
        try { fullName = String(doc.FullName || ''); } catch (e) { /* 未保存文档 FullName 可能拼默认路径 */ }
        try { dir = String(doc.Path || ''); } catch (e) { /* 同上 */ }
        try { saved = !!doc.Saved; } catch (e) { /* 默认按已保存处理 */ }
        try { readOnly = !!doc.ReadOnly; } catch (e) { /* 默认非只读 */ }
        return {
          ok: true,
          data: { app: 'word', name: name, fullName: fullName, saved: saved, readOnly: readOnly, hasPath: dir.length > 0 },
        };
      }, app, params);
    },

    /**
     * getSelectedText —— S1-REPORT.md:303（§3 #277：wps 型 Text 支持）
     * 语义差异（S1 §4 #9，报告行 339）：Selection 跟全局焦点走（与 Windows 同病），
     * 钉不住指定文档；wpp 型干脆没有 Selection——本动作只在 wps 型注册。
     * 返回串可能带裸 \r（S1 §4 #8）——原样返回，清洗职责归调用方（按 \n 拆行时注意）。
     */
    getSelectedText: function (app, params) {
      return guard(function (a) {
        var text = a.Selection.Text;
        return { ok: true, data: { text: text === null || text === undefined ? '' : String(text) } };
      }, app, params);
    },

    /**
     * replaceRange —— S1-REPORT.md:163（§3 #137 支持：doc.Range(a,b).Text = '…' 读回验证）
     * 参数名 startPos/endPos/text 与 TS handler 一致（word/proofread.ts:250）。
     */
    replaceRange: function (app, params) {
      return guard(function (a, p) {
        if (p.startPos === undefined || p.endPos === undefined) {
          return { ok: false, error: '[勿重试·先改参数] startPos 与 endPos 必填（replaceRange）' };
        }
        if (p.text === undefined || p.text === null) {
          return { ok: false, error: '[勿重试·先改参数] text 必填（replaceRange）' };
        }
        var doc = a.ActiveDocument;
        var range = doc.Range(Number(p.startPos), Number(p.endPos));
        range.Text = String(p.text);
        return { ok: true, data: { startPos: Number(p.startPos), endPos: Number(p.endPos), length: String(p.text).length } };
      }, app, params);
    },

    /**
     * setFont —— S1-REPORT.md:145（§3 #119 支持：Size/Name 写后读回恒等）
     * 参数与语义镜像 scripts/wps-com.ps1 的 setFont case：
     *   - range='all' → doc.Content.Font；否则 Selection.Font（选区钉不住的护栏同 ps1）；
     *   - 钉住文档（documentName）+ 非 all = 用活动文档选区，会改错地方 → 拒绝（ps1 同款护栏）；
     *   - color：色名表（red/blue/green/black）或 #hex → COM RGB 整数，Word 侧 Font.Color 直设
     *     （注意与 PPT 的 Font.Color.RGB 不同——Word 是 WdColor 整数属性；Font.Color 未在 S1 逐条验证，
     *     B3 真机复核点，见交付报告）。
     */
    setFont: function (app, params) {
      return guard(function (a, p) {
        if (p.documentName && p.range !== 'all') {
          return { ok: false, error: '[勿重试·先改参数] 钉住文档（documentName）时不能用选区范围改字体：选区属于活动文档，会改到别人身上。请改传 range=all' };
        }
        var font = p.range === 'all' ? a.ActiveDocument.Content.Font : a.Selection.Font;
        if (p.fontName !== undefined && p.fontName !== null) font.Name = String(p.fontName);
        if (p.fontSize !== undefined && p.fontSize !== null) font.Size = Number(p.fontSize);
        if (p.bold !== undefined && p.bold !== null) font.Bold = !!p.bold;
        if (p.italic !== undefined && p.italic !== null) font.Italic = !!p.italic;
        if (p.underline !== undefined && p.underline !== null) font.Underline = !!p.underline;
        var appliedColor = null;
        if (p.color) {
          var v = COLOR_NAMES[String(p.color).toLowerCase()];
          if (v === undefined) v = hexToRgbInt(String(p.color));
          if (v === null || v === undefined) {
            return { ok: false, error: '[勿重试·先改参数] 不认识的颜色「' + p.color + '」（可用色名 red/green/blue/black 或 #RRGGBB）' };
          }
          font.Color = v; // Word 侧直设整数（ps1 同款）；B3 复核点
          appliedColor = v;
        }
        return { ok: true, data: { settings: { fontName: p.fontName, fontSize: p.fontSize, bold: p.bold, italic: p.italic, underline: p.underline, colorRgb: appliedColor } } };
      }, app, params);
    },
  };

  /* 浏览器：挂 window；Node（单测/门禁）：挂 module.exports。channel-client.js 在分发时按注册表查找。 */
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = wordActions;
  } else {
    global.WpsAddonActionsWord = wordActions;
  }
})(typeof window !== 'undefined' ? window : this);
