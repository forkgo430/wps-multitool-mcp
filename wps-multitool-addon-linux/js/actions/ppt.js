/**
 * Input: 动作参数（与 src/tools 的 TS handler 所发参数同名；不再需要宿主对象——B4 走 jsAsynCall）
 * Output: Promise<{ ok: true, data } | { ok: false, error }>（通道协议回包结构）
 * Pos: Linux 加载项通道的动作层 · wpp 型（演示）15 个动作，B4 起为异步 jsAsynCall 模型。
 *      一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * B4 模型变更（依据 B3-REPORT.md §3.3）：webview 无 Application/wps 全局，唯一原生桥是
 * cefQuery + jsAsynCall。method 名编码自 S1-REPORT.md §3"对应的 JSAPI 操作"列（去宿主前缀、
 * 参数抽进 params），每个动作头注标 S1 报告行号；jsAsynCall 的 method 字符串形式在 S1/B3
 * 无直接样例——编码规则与缺口清单见交付报告（B4 §5），B3/B5 真机复核后统一修正。
 *
 * 铁律不变：
 *   - 动画【只做加】：S1-REPORT §5（报告行 355）——MainSequence.Count 读回时好时坏、坏时
 *     native signal 且桥死锁，Item(n).MoveTo/Delete 稳定抛错。异步模型下同样**绝不发起**任何
 *     动画读侧调用；level 恒 0（ps1 实测传 1 会建出两条重复效果）。
 *   - 枚举一律 JS 侧换算成编号再下发（§5：编号与 Windows 恒等；映射表抄 ps1 同名 case 并标行号）。
 *   - JSAPI 抛错/超时 → { ok:false, error: 原文 }，不吞、不静默成功。
 * 本文件只在 wpp 型加载项里加载（addon-wpp/index.html 引用）；wps/et 型没有 Presentations（S1 §6）。
 */
(function (global) {
  'use strict';

  var jsapi = (typeof module !== 'undefined' && module.exports)
    ? require('../jsapi.js')
    : global.WpsAddonJsapi;

  /** JSAPI 异常统一兜底：原文带回，不改写、不吞。
   * async 形态：jsAsynCall 的 rejection（超时/无 cefQuery/onFailure）也要接住转 ok:false——
   * 同步 try/catch 捕不到异步 reject（B4 首跑踩过：拒绝用例的 rejection 直接炸到顶层）。 */
  function guard(impl) {
    return async function (params) {
      try {
        return await impl(params || {});
      } catch (e) {
        return { ok: false, error: String(e && e.message ? e.message : e) };
      }
    };
  }

  /** 十六进制颜色 → COM RGB（镜像 scripts/wps-com.ps1:732；S0 #18 绿色读回 65280 验证编码） */
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

  function slideIndexOf(p) {
    return p.slideIndex === undefined || p.slideIndex === null ? 1 : Number(p.slideIndex);
  }

  /**
   * 版式名 → PpSlideLayout 编号。抄自 scripts/wps-com.ps1 的 addSlide case（$layouts 表）。
   * S1-REPORT §5（报告行 353-355）：数字枚举与 Windows 恒等；认不出的回落 2（ps1 同款）。
   */
  var SLIDE_LAYOUTS = { title: 1, title_content: 2, blank: 12, two_column: 3, comparison: 34 };

  /**
   * 切换效果名 → PpEntryEffect 编号。抄自 ps1 setSlideTransition case 的 $effectMap。
   * S1 §5（报告行 354）：0-80 全可设无抛错；名→ID 映射层自建（这张表）。
   */
  var TRANSITION_EFFECTS = {
    none: 0, fade: 1, push: 2, wipe: 3, split: 4, reveal: 5, cover: 6,
    dissolve: 7, curtains: 8, random: 6, blinds: 3,
  };

  /**
   * 动画效果名 → MsoAnimEffect 编号。抄自 ps1 addAnimation case 的 $effectMap（入场 + 强调）。
   * S1 §5（报告行 355）：effectId 1-80 全部可调。退场无稳定枚举名（ps1 同款）：认不出就报错。
   */
  var ANIM_EFFECTS = {
    appear: 1, fadeIn: 10, flyIn: 2, wipeIn: 22, zoomIn: 53, bounceIn: 26, spinIn: 15,
    dissolveIn: 9, boxIn: 4, circleIn: 6, diamondIn: 8, wheelIn: 21, plusIn: 13,
    wedgeIn: 20, randomBarsIn: 14, stripsIn: 18, blindsIn: 3, checkerboardIn: 5,
    crawlIn: 7, stretchIn: 17, swivelIn: 19, floatUp: 30, riseUp: 34, thinLine: 36,
    unfold: 37, whip: 38, expand: 50, shrink: 46,
    pulse: 63, spin: 15, grow: 59, teeter: 28, colorPulse: 62,
  };

  /** 形状定位参数规范化：shapeName 优先、shapeIndex 兜底（ps1 addAnimation 同款双键） */
  function shapeRef(p) {
    if (p.shapeName) return { shapeName: String(p.shapeName) };
    if (p.shapeIndex !== undefined && p.shapeIndex !== null) return { shapeIndex: Number(p.shapeIndex) };
    return null;
  }

  var pptActions = {

    /**
     * createPresentation —— S1-REPORT.md:309（§3 #283：app.Presentations.Add() 支持，仅 wpp 型）
     * method 'Presentations.Add'。
     */
    createPresentation: guard(function () {
      return jsapi.jsAsynCall('Presentations.Add', {}).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getSlideCount —— S1-REPORT.md:202（§3 #176）
     * method 'Slides.Count'：S1 列 `pres.Slides.Count`（属性读编码，待补测）。
     */
    getSlideCount: guard(function () {
      return jsapi.jsAsynCall('Slides.Count', {}).then(function (r) {
        var count = r && typeof r === 'object' && !Array.isArray(r) && 'count' in r ? r.count : r;
        return { ok: true, data: { count: Number(count) } };
      });
    }),

    /**
     * addSlide —— S1-REPORT.md:210（§3 #184：Slides.Add(index, layout) 支持）
     * method 'Slides.Add'。layout 名→编号在 JS 侧换算（ps1 $layouts 表）；
     * position 缺省由调用方语义决定（页尾），title/content 原样下发由原生侧落占位符。
     */
    addSlide: guard(function (p) {
      var layoutKey = p.layout || 'title_content';
      var layoutType = SLIDE_LAYOUTS[layoutKey];
      if (layoutType === undefined) layoutType = 2; // ps1 同款回落
      var payload = {
        layout: layoutType,
        layoutName: layoutKey,
        position: p.position !== undefined && p.position !== null ? Number(p.position) : null, // null = 页尾，原生侧按 Count+1
      };
      if (p.title !== undefined && p.title !== null) payload.title = String(p.title);
      if (p.content !== undefined && p.content !== null) payload.content = String(p.content);
      return jsapi.jsAsynCall('Slides.Add', payload).then(function (r) {
        return { ok: true, data: { slideIndex: r && r.slideIndex !== undefined ? r.slideIndex : payload.position, layout: layoutKey } };
      });
    }),

    /**
     * deleteSlide —— S1-REPORT.md:198（§3 #172：slide.Delete() 支持）
     * method 'Slides.Delete'：S1 列 `pres.Slides.Item(index).Delete()`（定位参数抽进 params）。
     */
    deleteSlide: guard(function (p) {
      var index = p.index || p.slideIndex;
      if (!index) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 slideIndex/index（deleteSlide 必填）' });
      }
      return jsapi.jsAsynCall('Slides.Delete', { index: Number(index) }).then(function (r) {
        return { ok: true, data: { deleted: Number(index) } };
      });
    }),

    /**
     * getShapes —— S1-REPORT.md:222（§3 #196：遍历 Shapes 读 Type/Name 支持）
     * method 'Slide.Shapes'。返回透传（原生侧组装的形状数组）。
     */
    getShapes: guard(function (p) {
      return jsapi.jsAsynCall('Slide.Shapes', { slideIndex: slideIndexOf(p) }).then(function (r) {
        return { ok: true, data: { slideIndex: slideIndexOf(p), shapes: r } };
      });
    }),

    /**
     * setShapeText —— S1-REPORT.md:226（§3 #200：TextFrame.TextRange.Text 写后读回支持）
     * method 'Shape.TextFrame.TextRange.Text'：S1 列 `shape.TextFrame.TextRange.Text = '…'`。
     */
    setShapeText: guard(function (p) {
      if (p.text === undefined || p.text === null) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 text（setShapeText 必填）' });
      }
      var ref = shapeRef(p);
      if (!ref) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 找不到目标形状：shapeIndex（1 起）/shapeName 都没给或不存在' });
      }
      var payload = { slideIndex: slideIndexOf(p), text: String(p.text) };
      for (var k in ref) payload[k] = ref[k];
      return jsapi.jsAsynCall('Shape.TextFrame.TextRange.Text', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * addTextBox —— S1-REPORT.md:211（§3 #185：AddTextbox(orientation,l,t,w,h) 支持）
     * method 'Shapes.AddTextbox'。orientation 固定 1（msoTextOrientationHorizontal，ps1 同值）；
     * 缺省 100/100/400/50 镜像 ps1。
     */
    addTextBox: guard(function (p) {
      var payload = {
        slideIndex: slideIndexOf(p),
        orientation: 1,
        left: p.left !== undefined && p.left !== null ? Number(p.left) : 100,
        top: p.top !== undefined && p.top !== null ? Number(p.top) : 100,
        width: p.width !== undefined && p.width !== null ? Number(p.width) : 400,
        height: p.height !== undefined && p.height !== null ? Number(p.height) : 50,
      };
      if (p.text !== undefined && p.text !== null) payload.text = String(p.text);
      if (p.fontSize !== undefined && p.fontSize !== null) payload.fontSize = Number(p.fontSize);
      if (p.fontName) payload.fontName = String(p.fontName);
      return jsapi.jsAsynCall('Shapes.AddTextbox', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * setShapePosition —— S1-REPORT.md:227（§3 #201：Left/Top/Width/Height 支持读回）
     * method 'Shape.Position'：S1 列 `shape.{Left,Top,Width,Height}`（多属性写编码，待补测）。
     * 只把调用方给了的维度放进 params（缺省维度不下发）。
     */
    setShapePosition: guard(function (p) {
      var ref = shapeRef(p);
      if (!ref) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 找不到目标形状：shapeIndex（1 起）/shapeName 都没给或不存在' });
      }
      var payload = { slideIndex: slideIndexOf(p) };
      for (var k in ref) payload[k] = ref[k];
      var touched = false;
      ['left', 'top', 'width', 'height'].forEach(function (dim) {
        if (p[dim] !== undefined && p[dim] !== null) {
          payload[dim] = Number(p[dim]);
          touched = true;
        }
      });
      if (!touched) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] left/top/width/height 至少给一个（setShapePosition）' });
      }
      return jsapi.jsAsynCall('Shape.Position', payload).then(function (r) {
        return { ok: true, data: { left: p.left, top: p.top, width: p.width, height: p.height } };
      });
    }),

    /**
     * insertPptTable —— S1-REPORT.md:233（§3 #207：AddTable(rows,cols,l,t,w,h) 支持）
     * method 'Shapes.AddTable'。缺省 3×3 @ (100,100) 400×200 镜像 ps1。
     */
    insertPptTable: guard(function (p) {
      var payload = {
        slideIndex: slideIndexOf(p),
        rows: p.rows ? Number(p.rows) : 3,
        cols: p.cols ? Number(p.cols) : 3,
        left: p.left !== undefined && p.left !== null ? Number(p.left) : 100,
        top: p.top !== undefined && p.top !== null ? Number(p.top) : 100,
        width: p.width !== undefined && p.width !== null ? Number(p.width) : 400,
        height: p.height !== undefined && p.height !== null ? Number(p.height) : 200,
      };
      return jsapi.jsAsynCall('Shapes.AddTable', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getPptTableCell —— S1-REPORT.md:235（§3 #209：Table.Cell(r,c).Shape.TextFrame.TextRange.Text 支持）
     * method 'Table.Cell.Text'（定位参数 tableIndex/row/col 进 params）。
     */
    getPptTableCell: guard(function (p) {
      return jsapi.jsAsynCall('Table.Cell.Text', {
        slideIndex: slideIndexOf(p),
        tableIndex: Number(p.tableIndex || 1),
        row: Number(p.row),
        col: Number(p.col),
      }).then(function (r) {
        var text = r && typeof r === 'object' && !Array.isArray(r) && 'text' in r ? r.text : r;
        text = text === null || text === undefined ? '' : String(text);
        return { ok: true, data: { value: text, text: text } };
      });
    }),

    /**
     * setPptTableCell —— S1-REPORT.md:234（§3 #208：同链路写入支持）
     * method 'Table.Cell.Text'（写）。
     */
    setPptTableCell: guard(function (p) {
      if (p.text === undefined || p.text === null) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 text（setPptTableCell 必填）' });
      }
      return jsapi.jsAsynCall('Table.Cell.Text', {
        slideIndex: slideIndexOf(p),
        tableIndex: Number(p.tableIndex || 1),
        row: Number(p.row),
        col: Number(p.col),
        text: String(p.text),
      }).then(function (r) {
        return { ok: true, data: { row: Number(p.row), col: Number(p.col) } };
      });
    }),

    /**
     * addAnimation —— S1-REPORT.md:252/271/272（§3 #226/#245/#246：AddEffect 全可加）+ §5 报告行 355
     * method 'TimeLine.MainSequence.AddEffect'（S1 列原链即方法调用）。
     * 【只做加】：异步模型下同样绝不发起动画读侧调用（不读 Count、不 Delete/MoveTo）；
     * level 恒 0（msoAnimateLevelNone，ps1 实测传 1 会建出两条重复效果）；回包无 count（有意差异）。
     */
    addAnimation: guard(function (p) {
      var ref = shapeRef(p);
      if (!ref) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 找不到要加动画的形状：shapeIndex（1 起）/shapeName 都没给或不存在' });
      }
      var raw = p.effect !== undefined && p.effect !== null ? p.effect : p.animationType; // 旧键名仍认（ps1 同款）
      if (raw === undefined || raw === null || raw === '') {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 effect（可用：appear/fadeIn/flyIn/wipeIn/zoomIn/pulse 等，或直接给 MsoAnimEffect 数字）' });
      }
      var effectId;
      if (typeof raw === 'number') {
        effectId = raw;
      } else if (/^\d+$/.test(String(raw))) {
        effectId = parseInt(String(raw), 10);
      } else {
        effectId = ANIM_EFFECTS[String(raw)];
        if (effectId === undefined) {
          return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 不认识的 effect「' + raw + '」；常用：appear/fadeIn/flyIn/wipeIn/zoomIn（入场）、pulse/spin/grow/teeter（强调）。退场效果请直接给 MsoAnimEffect 数字' });
        }
      }
      var payload = { slideIndex: slideIndexOf(p), effectId: effectId, level: 0 };
      for (var k in ref) payload[k] = ref[k];
      return jsapi.jsAsynCall('TimeLine.MainSequence.AddEffect', payload).then(function (r) {
        return { ok: true, data: { slideIndex: slideIndexOf(p), effectId: effectId } };
      });
    }),

    /**
     * setSlideTransition —— S1-REPORT.md:253（§3 #227：EntryEffect 编号直设支持）+ §5 报告行 354
     * method 'SlideShowTransition.EntryEffect'（属性写编码，待补测）。
     * 效果名→编号在 JS 侧换算（ps1 $effectMap 表）；纯数字串/数字直通（0-80 全可设）。
     */
    setSlideTransition: guard(function (p) {
      var raw = p.effect !== undefined && p.effect !== null ? p.effect : p.transition; // 旧键名仍认（ps1 同款）
      var name = raw === undefined || raw === null || raw === '' ? 'fade' : String(raw).toLowerCase();
      var entry = -1;
      if (Object.prototype.hasOwnProperty.call(TRANSITION_EFFECTS, name)) {
        entry = TRANSITION_EFFECTS[name];
      } else if (/^\d+$/.test(name)) {
        entry = parseInt(name, 10);
      }
      if (entry < 0) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 不认识的切换效果：' + name + '（可用：none/fade/push/wipe/split/reveal/cover/dissolve/curtains）' });
      }
      var dur = 0;
      if (p.duration !== undefined && p.duration !== null) {
        var n = Number(p.duration);
        if (!isNaN(n) && n > 0) dur = n;
      }
      var payload = { slideIndex: slideIndexOf(p), entryEffect: entry };
      if (dur > 0) payload.duration = dur;
      return jsapi.jsAsynCall('SlideShowTransition.EntryEffect', payload).then(function (r) {
        return { ok: true, data: { slideIndex: slideIndexOf(p), effect: name, entryEffect: entry, duration: dur } };
      });
    }),

    /**
     * setSlideBackground —— S1-REPORT.md:254（§3 #228：Background.Fill.Solid/ForeColor 支持）+ S0 #18
     * method 'Background.Fill.Solid'：S1 列 `slide.Background.Fill.Solid/ForeColor.RGB` 两步——
     * 编码为单 method，colorRgb 进 params（先 Solid 再赋色的次序由原生侧承担，B3 复核点）。
     * 非法颜色明确报错（ps1 是静默跳过仍回 success——对 ps1 的有意偏离，B2.5 已定）。
     */
    setSlideBackground: guard(function (p) {
      if (!p.color) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 color（setSlideBackground 本轮只实现纯色路径，#RRGGBB）' });
      }
      var rgb = hexToRgbInt(String(p.color));
      if (rgb === null) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 不认识的颜色「' + p.color + '」（要 #RRGGBB 或 #RGB 三位）' });
      }
      return jsapi.jsAsynCall('Background.Fill.Solid', { slideIndex: slideIndexOf(p), colorRgb: rgb }).then(function (r) {
        return { ok: true, data: { slideIndex: slideIndexOf(p), colorRgb: rgb } };
      });
    }),

    /**
     * exportSlideAsImage —— S1-REPORT.md:232（§3 #206）+ S0 #19（真实生成 1334×750 PNG 验证）
     * method 'Slide.Export'：S1 列 `slide.Export(path, 'PNG')`；B3 §5（报告行 137）明确本动作属异步范畴。
     * 带尺寸时四参下发（S0 实测带尺寸调用成功）。
     */
    exportSlideAsImage: guard(function (p) {
      var path = p.outputPath || p.path; // TS 两个键都发（ppt/image.ts:393-399）
      if (!path) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 outputPath/path（exportSlideAsImage 必填）' });
      }
      var format = p.format || 'PNG';
      var payload = { slideIndex: slideIndexOf(p), path: String(path), filter: String(format) };
      if (p.width !== undefined && p.width !== null && p.height !== undefined && p.height !== null) {
        payload.width = Number(p.width);
        payload.height = Number(p.height);
      }
      return jsapi.jsAsynCall('Slide.Export', payload).then(function (r) {
        return { ok: true, data: { slideIndex: slideIndexOf(p), outputPath: String(path), format: String(format) } };
      });
    }),
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = pptActions;
  } else {
    global.WpsAddonActionsPpt = pptActions;
  }
})(typeof window !== 'undefined' ? window : this);
