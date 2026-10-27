/**
 * Input: 动作参数（与 src/tools 的 TS handler 所发参数同名；不再需要宿主对象——B4 走 jsAsynCall）
 * Output: Promise<{ ok: true, data } | { ok: false, error }>（通道协议回包结构）
 * Pos: Linux 加载项通道的动作层 · wps 型（文字）5 个动作，B4 起为异步 jsAsynCall 模型。
 *      一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * B4 模型变更（依据 B3-REPORT.md §3.3）：webview 无 Application/wps 全局，唯一原生桥是
 * cefQuery + jsAsynCall（异步方法调用 + 全局回调）。B2.5 的同步 COM 门面在真机上无对象可调。
 *
 * method 名依据（B4 任务卡 §1.1 + 红线 4）：method 编码自 S1-REPORT.md §3"对应的 JSAPI 操作"列
 * 的属性链（去宿主前缀、参数抽进 params）；每个动作头注标 S1 报告行号。
 * ⚠ 诚实声明：S1 的证据是同步门面上的属性链调用，jsAsynCall 的 method 字符串形式在 S1/B3
 * 里没有直接样例——编码规则与缺口清单见交付报告（B4 §5），B3/B5 真机复核后统一修正。
 *
 * S1 有背书的语义照搬：换算（颜色 RGB）、校验（缺参显式拒绝）、错误不吞全部保留在 JS 侧。
 * 本文件只在 wps 型加载项里加载（addon-wps/index.html 引用）；et/wpp 型没有 Documents（S1 §6）。
 */
(function (global) {
  'use strict';

  var jsapi = (typeof module !== 'undefined' && module.exports)
    ? require('../jsapi.js')
    : global.WpsAddonJsapi;

  /** JSAPI 异常统一兜底：原文带回，不改写、不吞。
   * async 形态：jsAsynCall 的 rejection（超时/无 cefQuery/onFailure）也要接住转 ok:false——
   * 同步 try/catch 捕不到异步 reject（B4 首跑踩过：openDocument 的拒绝用例直接炸到顶层）。 */
  function guard(impl) {
    return async function (params) {
      try {
        return await impl(params || {});
      } catch (e) {
        return { ok: false, error: String(e && e.message ? e.message : e) };
      }
    };
  }

  /**
   * 十六进制颜色 → COM RGB 整数（R + G*256 + B*65536）。
   * 镜像 scripts/wps-com.ps1:732；编码有 S0 实测背书（S1-REPORT §2.4：绿色 #00FF00 → 65280 读回一致）。
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

  /** 色名 → COM RGB（镜像 ps1 setFont case 的 colorMap） */
  var COLOR_NAMES = { red: 255, blue: 16711680, green: 65280, black: 0 };

  var wordActions = {

    /**
     * openDocument —— S1-REPORT.md:134（§3 #108 支持）
     * method 'Documents.Open'：S1 列 `app.Documents.Open(path)`。
     * 语义差异（S1 §4 #12，报告行 341）：同文件重复 Open 只激活已有文档；引用在文档关闭后失活——
     * 结果由原生侧回调带回，本层不缓存任何引用。
     */
    openDocument: guard(function (p) {
      var path = p.path || p.filePath; // TS 两个键都发（word/document.ts:235-239）
      if (!path) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 path/filePath（openDocument 必填）' });
      }
      return jsapi.jsAsynCall('Documents.Open', { path: path }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getActiveDocumentState —— S1-REPORT.md:129（§3 #103）
     * method 'ActiveDocument'：S1 列 `app.ActiveDocument.Name` 等（多字段属性读）。
     * ⚠ 属性读在 jsAsynCall 下的确切语义待 S1/B3 补测（method 编码规则见交付报告）。
     * 结果透传原生侧，本层不组装——原生返回的键名以真机为准（B3 复核点）。
     */
    getActiveDocumentState: guard(function () {
      return jsapi.jsAsynCall('ActiveDocument', {}).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getSelectedText —— S1-REPORT.md:303（§3 #277：wps 型 Selection.Text 支持）
     * method 'Selection.Text'：S1 列 `app.Selection.Text`。
     * 语义差异（S1 §4 #9）：Selection 跟全局焦点走，钉不住指定文档；wpp 型无 Selection——本动作只在 wps 型。
     * 返回串可能带裸 \r（S1 §4 #8）——原样透传，清洗职责归调用方。
     */
    getSelectedText: guard(function () {
      return jsapi.jsAsynCall('Selection.Text', {}).then(function (r) {
        var text = r && typeof r === 'object' && !Array.isArray(r) && 'text' in r ? r.text : r;
        return { ok: true, data: { text: text === null || text === undefined ? '' : String(text) } };
      });
    }),

    /**
     * replaceRange —— S1-REPORT.md:163（§3 #137 支持）
     * method 'Range.Text'：S1 列 `doc.Range(a,b).Text = '…'`（属性写编码进 params，待补测）。
     */
    replaceRange: guard(function (p) {
      if (p.startPos === undefined || p.endPos === undefined) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] startPos 与 endPos 必填（replaceRange）' });
      }
      if (p.text === undefined || p.text === null) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] text 必填（replaceRange）' });
      }
      return jsapi.jsAsynCall('Range.Text', {
        startPos: Number(p.startPos),
        endPos: Number(p.endPos),
        text: String(p.text),
      }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * setFont —— S1-REPORT.md:145（§3 #119 支持：Size/Name 写后读回恒等）
     * method 'Font'：S1 列 `…Font.Size/Name`（多属性写编码进 params，待补测）。
     * 换算在 JS 侧完成：color（色名/十六进制）→ COM RGB 整数再下发；
     * range='all' → 作用于 Content.Font，否则选区（ps1 同款语义）；
     * 钉文档（documentName）+ 非 all = 选区会改错地方 → 拒绝（ps1 setFont 同款护栏）。
     * ⚠ Font.Color 通道 S1 未逐条验证（#119 只测了 Size/Name）——B3 复核点。
     */
    setFont: guard(function (p) {
      if (p.documentName && p.range !== 'all') {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 钉住文档（documentName）时不能用选区范围改字体：选区属于活动文档，会改到别人身上。请改传 range=all' });
      }
      var payload = { range: p.range === 'all' ? 'all' : 'selection' };
      if (p.fontName !== undefined && p.fontName !== null) payload.fontName = String(p.fontName);
      if (p.fontSize !== undefined && p.fontSize !== null) payload.fontSize = Number(p.fontSize);
      if (p.bold !== undefined && p.bold !== null) payload.bold = !!p.bold;
      if (p.italic !== undefined && p.italic !== null) payload.italic = !!p.italic;
      if (p.underline !== undefined && p.underline !== null) payload.underline = !!p.underline;
      var appliedColor = null;
      if (p.color) {
        var v = COLOR_NAMES[String(p.color).toLowerCase()];
        if (v === undefined) v = hexToRgbInt(String(p.color));
        if (v === null || v === undefined) {
          return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 不认识的颜色「' + p.color + '」（可用色名 red/green/blue/black 或 #RRGGBB）' });
        }
        payload.colorRgb = v;
        appliedColor = v;
      }
      return jsapi.jsAsynCall('Font', payload).then(function (r) {
        return { ok: true, data: { settings: { fontName: p.fontName, fontSize: p.fontSize, bold: p.bold, italic: p.italic, underline: p.underline, colorRgb: appliedColor } } };
      });
    }),
  };

  // ========================================================================
  // B5-T3：部分支持收编（Word 6 + Common 2 的 wps 型缺口）。S1 §3 行号见各头注。
  // Common 的 createPresentation 已在 ppt.js（B4）实现——不重复注册；getSelectedText
  // 的 et 型缺口在 excel.js（Address 函数读），wpp 型无 Selection 不做（S1 §4 #9）。
  // ========================================================================

  var b5t3Actions = {
    /**
     * closeDocument —— S1-REPORT.md:109（§3 #109 部分支持）
     * 证据：doc.Close() 对干净文档可关；**dirty 文档弹「是否保存」模态框阻塞 JS 线程直至用户点击**。
     * method 'Document.Close'。saveChanges 缺省 false（放弃保存关闭——自动化通道不冒模态框风险）；
     * 要保改动先显式 save 再关。⚠ S1 模态框阻塞语义未在异步模型下复测——B3.5 复核点。
     */
    closeDocument: guard(function (p) {
      var payload = {};
      if (p.name) payload.name = String(p.name);
      payload.saveChanges = p.saveChanges !== undefined ? !!p.saveChanges : false;
      return jsapi.jsAsynCall('Document.Close', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getDocumentStats —— S1-REPORT.md:110（§3 #110 部分支持：Words/Characters/Paragraphs/Sections
     * 有值，**Pages 属性不存在**——4/5 才是完整语义）
     * method 'Document.Stats'（多字段读，只请求存在的 4 项，Pages 不请求；待补测）。
     */
    getDocumentStats: guard(function () {
      return jsapi.jsAsynCall('Document.Stats', {}).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * replaceBookmarkContent —— S1-REPORT.md:131（§3 #131 部分支持：TypeError doc.Styles is not a
     * function 同族——Bookmarks 是集合，Item 按名可取（getBookmarks 实测），内容替换未单独验证）
     * method 'Bookmarks.Item.Range.Text'（Item 语义 + 属性写，待补测）。
     */
    replaceBookmarkContent: guard(function (p) {
      if (!p.name) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 name（replaceBookmarkContent 要替换哪个书签）' });
      }
      if (p.text === undefined || p.text === null) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 text（replaceBookmarkContent 的新内容）' });
      }
      return jsapi.jsAsynCall('Bookmarks.Item.Range.Text', { name: String(p.name), text: String(p.text) }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * modifyStyle —— S1-REPORT.md:149（§3 #149 部分支持：TypeError doc.Styles is not a function 复核
     * ——Styles 集合 Item 可取（listStyles 实测），样式修改未单独验证）
     * method 'Styles.Item.Font'（Item 语义 + 多属性写，待补测）。fontName/fontSize 至少一个。
     */
    modifyStyle: guard(function (p) {
      if (!p.name) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 name（modifyStyle 要改哪个样式）' });
      }
      if (p.fontName === undefined && p.fontSize === undefined) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] fontName/fontSize 至少给一个（modifyStyle）' });
      }
      var payload = { name: String(p.name) };
      if (p.fontName !== undefined) payload.fontName = String(p.fontName);
      if (p.fontSize !== undefined) payload.fontSize = Number(p.fontSize);
      return jsapi.jsAsynCall('Styles.Item.Font', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * deleteStyle —— S1-REPORT.md:150（§3 #150 部分支持：同上 Item 语义，删除未单独验证）
     * method 'Styles.Item.Delete'。内建样式（如 Normal）删除会抛——由原生侧如实回错，不预判。
     */
    deleteStyle: guard(function (p) {
      if (!p.name) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 name（deleteStyle 要删哪个样式）' });
      }
      return jsapi.jsAsynCall('Styles.Item.Delete', { name: String(p.name) }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * setDefaultFont —— S1-REPORT.md:157（§3 #157 部分支持：doc.Styles('Normal').Font.Name）
     * method 'Styles.Item.Font'（target 固定 Normal——默认字体的载体就是 Normal 样式；与 modifyStyle
     * 同 method 不同定位）。fontName 必填。
     */
    setDefaultFont: guard(function (p) {
      if (!p.fontName) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 fontName（setDefaultFont）' });
      }
      return jsapi.jsAsynCall('Styles.Item.Font', { name: 'Normal', fontName: String(p.fontName) }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * setSelectedText —— S1-REPORT.md:278（§3 #278 部分支持：wps 型 TypeText 已调用读回一致；
     * et 无 TypeText（§4 #14）、wpp 无 Selection——本实现只覆盖 wps 型）
     * method 'Selection.TypeText'。text 必填。
     * ⚠ Selection 跟全局焦点走（S1 §4 #9）——写入落当前活动文档，钉不住指定文档。
     */
    setSelectedText: guard(function (p) {
      if (p.text === undefined || p.text === null || p.text === '') {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 text（setSelectedText 的写入内容）' });
      }
      return jsapi.jsAsynCall('Selection.TypeText', { text: String(p.text) }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * createDocument —— S1-REPORT.md:282（§3 #282 部分支持：仅 wps 型 app.Documents.Add() 支持）
     * method 'Documents.Add'。本实现注册在 word.js（wps 型）——et/wpp 型无 Documents（S1 §6 隔离）。
     */
    createDocument: guard(function () {
      return jsapi.jsAsynCall('Documents.Add', {}).then(function (r) {
        return { ok: true, data: r };
      });
    }),
  };

  Object.assign(wordActions, b5t3Actions);

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = wordActions;
  } else {
    global.WpsAddonActionsWord = wordActions;
  }
})(typeof window !== 'undefined' ? window : this);
