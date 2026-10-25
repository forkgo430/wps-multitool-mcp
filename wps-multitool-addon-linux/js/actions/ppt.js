/**
 * Input: 宿主 Application 对象（wpp 型 = 演示）、动作参数（与 src/tools 的 TS handler 所发参数同名）
 * Output: { ok: true, data } | { ok: false, error }（通道协议回包结构，B2 已定）
 * Pos: Linux 加载项通道的动作层 · wpp 型（演示）15 个动作。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 语义依据：S1-REPORT.md §3 逐条明细（每条函数头部标了报告行号）+ §5 枚举映射。
 * 实现纪律（B2.5 任务卡）：
 *   - 动画【只做加】（§5 报告行 355：MainSequence.Count 读回时好时坏、坏时 native signal 且桥死锁，
 *     Item(n).MoveTo/Delete 稳定抛错）——本文件不出现任何动画读侧调用，加完也不读 Count 校验；
 *   - 枚举一律编号直设（§5：编号与 Windows 恒等，Linux 无名→ID 映射层，映射表从
 *     scripts/wps-com.ps1 的同名 case 抄录并注明行号）；
 *   - JSAPI 抛错 → { ok:false, error: 原文 }，不吞、不静默成功。
 * 本文件只在 wpp 型加载项里加载（addon-wpp）；wps/et 型没有 Presentations（S1 §6 三型隔离）。
 */
(function (global) {
  'use strict';

  function guard(impl, app, params) {
    try {
      return impl(app, params || {});
    } catch (e) {
      return { ok: false, error: String(e && e.message ? e.message : e) };
    }
  }

  /** 十六进制颜色 → COM RGB（镜像 scripts/wps-com.ps1:732；PPT 侧 ForeColor.RGB 用，S0 #18 读回 65280 验证编码） */
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

  /** 活动演示文稿。S1-REPORT.md:193（§3 #167：ActivePresentation.Name 支持） */
  function getPres(app) {
    var pres = app.ActivePresentation;
    if (!pres) throw new Error('没有活动演示文稿（ActivePresentation 为空）');
    return pres;
  }

  /** 取第 n 页（1 起；缺省 1）。S1 §3 #178：Slides.Item(n) 可用。 */
  function getSlide(pres, p, def) {
    var idx = p.slideIndex === undefined || p.slideIndex === null ? (def || 1) : Number(p.slideIndex);
    return pres.Slides.Item(idx);
  }

  /** 按名或序号找形状（镜像 ps1 addAnimation 的双重定位：shapeName 优先、shapeIndex 兜底） */
  function findShape(slide, p) {
    if (p.shapeName) return slide.Shapes.Item(String(p.shapeName));
    if (p.shapeIndex !== undefined && p.shapeIndex !== null) return slide.Shapes.Item(Number(p.shapeIndex));
    return null;
  }

  /** 数本页第 tableIndex 个表格（HasTable 形状计数，镜像 ps1 insertPptTable/getPptTableCell 的定位算法） */
  function findTableShape(slide, tableIndex) {
    var want = Number(tableIndex || 1);
    var seen = 0;
    var count = slide.Shapes.Count;
    for (var i = 1; i <= count; i++) {
      var s = slide.Shapes.Item(i);
      var has = false;
      try { has = !!s.HasTable; } catch (e) { has = false; }
      if (has) {
        seen += 1;
        if (seen === want) return s;
      }
    }
    return null;
  }

  /**
   * 版式名 → PpSlideLayout 编号。抄自 scripts/wps-com.ps1 的 addSlide case（$layouts 表）。
   * 编号与 Windows 恒等有 S1 §5 背书（EntryEffect/AutoShapeType 同理）；认不出的回落 2（ps1 同款）。
   */
  var SLIDE_LAYOUTS = { title: 1, title_content: 2, blank: 12, two_column: 3, comparison: 34 };

  /**
   * 切换效果名 → PpEntryEffect 编号。抄自 ps1 setSlideTransition case 的 $effectMap。
   * S1-REPORT.md:354（§5）：0-80 全可设无抛错；Linux 只认数字，名→ID 映射层自建（就是这张表）。
   */
  var TRANSITION_EFFECTS = {
    none: 0, fade: 1, push: 2, wipe: 3, split: 4, reveal: 5, cover: 6,
    dissolve: 7, curtains: 8, random: 6, blinds: 3,
  };

  /**
   * 动画效果名 → MsoAnimEffect 编号。抄自 ps1 addAnimation case 的 $effectMap（入场 + 强调）。
   * S1-REPORT.md:355（§5）：effectId 1-80 全部可调（实测 21/3/29 三组均返回 object），编号与 Windows 恒等。
   * 退场效果 WPS 无稳定枚举名（ps1 同款策略）：认不出就明确报错，让调用方直接给数字。
   */
  var ANIM_EFFECTS = {
    appear: 1, fadeIn: 10, flyIn: 2, wipeIn: 22, zoomIn: 53, bounceIn: 26, spinIn: 15,
    dissolveIn: 9, boxIn: 4, circleIn: 6, diamondIn: 8, wheelIn: 21, plusIn: 13,
    wedgeIn: 20, randomBarsIn: 14, stripsIn: 18, blindsIn: 3, checkerboardIn: 5,
    crawlIn: 7, stretchIn: 17, swivelIn: 19, floatUp: 30, riseUp: 34, thinLine: 36,
    unfold: 37, whip: 38, expand: 50, shrink: 46,
    pulse: 63, spin: 15, grow: 59, teeter: 28, colorPulse: 62,
  };

  var pptActions = {

    /**
     * createPresentation —— S1-REPORT.md:309（§3 #283：app.Presentations.Add() 支持，仅 wpp 型）
     */
    createPresentation: function (app, params) {
      return guard(function (a) {
        var pres = a.Presentations.Add();
        var name = '';
        try { name = String(pres.Name || ''); } catch (e) { /* Name 读取失败不算失败 */ }
        return { ok: true, data: { name: name } };
      }, app, params);
    },

    /**
     * getSlideCount —— S1-REPORT.md:202（§3 #176：pres.Slides.Count 支持）
     */
    getSlideCount: function (app, params) {
      return guard(function (a) {
        var pres = getPres(a);
        return { ok: true, data: { count: Number(pres.Slides.Count) } };
      }, app, params);
    },

    /**
     * addSlide —— S1-REPORT.md:210（§3 #184：Slides.Add(index, layout) 支持）
     * layout 名→编号、position 缺省 Count+1、title/content 落占位符，全部镜像 ps1 addSlide case。
     * content 定位：PlaceholderFormat.Type === 2（正文占位符），ps1 同款。
     */
    addSlide: function (app, params) {
      return guard(function (a, p) {
        var pres = getPres(a);
        var layoutKey = p.layout || 'title_content';
        var layoutType = SLIDE_LAYOUTS[layoutKey];
        if (layoutType === undefined) layoutType = 2; // ps1 同款回落：认不出的版式按 title_content
        var position = p.position ? Number(p.position) : Number(pres.Slides.Count) + 1;
        var slide = pres.Slides.Add(position, layoutType);
        if (p.title) {
          try {
            if (slide.Shapes.HasTitle) slide.Shapes.Title.TextFrame.TextRange.Text = String(p.title);
          } catch (e) { /* 无标题占位符的版式（如 blank）——不视为失败，ps1 同款 try */ }
        }
        if (p.content) {
          try {
            for (var i = 1; i <= slide.Shapes.Count; i++) {
              var shape = slide.Shapes.Item(i);
              try {
                if (Number(shape.PlaceholderFormat.Type) === 2) {
                  shape.TextFrame.TextRange.Text = String(p.content);
                  break;
                }
              } catch (e) { /* 非占位符形状没有 PlaceholderFormat */ }
            }
          } catch (e) { /* 整页遍历失败不视为失败，ps1 同款 */ }
        }
        return { ok: true, data: { slideIndex: position, layout: layoutKey } };
      }, app, params);
    },

    /**
     * deleteSlide —— S1-REPORT.md:198（§3 #172：slide.Delete() 支持）
     * 双键 index/slideIndex 与 ps1 deleteSlide case 一致。
     */
    deleteSlide: function (app, params) {
      return guard(function (a, p) {
        var index = p.index || p.slideIndex;
        if (!index) {
          return { ok: false, error: '[勿重试·先改参数] 缺少 slideIndex/index（deleteSlide 必填）' };
        }
        var pres = getPres(a);
        pres.Slides.Item(Number(index)).Delete();
        return { ok: true, data: { deleted: Number(index) } };
      }, app, params);
    },

    /**
     * getShapes —— S1-REPORT.md:222（§3 #196：遍历 Shapes 读 Type/Name 支持）
     * 返回结构与 ps1 getShapes case 对齐（TS 侧按 {name,index,type,left,top,width,height} 消费）。
     */
    getShapes: function (app, params) {
      return guard(function (a, p) {
        var pres = getPres(a);
        var slide = getSlide(pres, p, 1);
        var shapes = [];
        var count = Number(slide.Shapes.Count);
        for (var i = 1; i <= count; i++) {
          var s = slide.Shapes.Item(i);
          shapes.push({
            name: String(s.Name || ''),
            index: i,
            type: s.Type,
            left: s.Left,
            top: s.Top,
            width: s.Width,
            height: s.Height,
          });
        }
        return { ok: true, data: { slideIndex: p.slideIndex || 1, shapes: shapes, count: shapes.length } };
      }, app, params);
    },

    /**
     * setShapeText —— S1-REPORT.md:226（§3 #200：TextFrame.TextRange.Text 写后读回支持）
     */
    setShapeText: function (app, params) {
      return guard(function (a, p) {
        if (p.text === undefined || p.text === null) {
          return { ok: false, error: '[勿重试·先改参数] 缺少 text（setShapeText 必填）' };
        }
        var pres = getPres(a);
        var slide = getSlide(pres, p, 1);
        var shape = findShape(slide, p);
        if (!shape) {
          return { ok: false, error: '[勿重试·先改参数] 找不到目标形状：shapeIndex（1 起）/shapeName 都没给或不存在' };
        }
        shape.TextFrame.TextRange.Text = String(p.text);
        var name = '';
        try { name = String(shape.Name || ''); } catch (e) { /* 同 createPresentation */ }
        return { ok: true, data: { name: name } };
      }, app, params);
    },

    /**
     * addTextBox —— S1-REPORT.md:211（§3 #185：AddTextbox(orientation,l,t,w,h) 支持）
     * orientation 固定 1（msoTextOrientationHorizontal，ps1 addTextBox case 同值）；
     * 默认值 100/100/400/50 镜像 ps1（注意 TS 层自己发 200 宽默认——本层只管缺省直调）。
     */
    addTextBox: function (app, params) {
      return guard(function (a, p) {
        var pres = getPres(a);
        var slide = getSlide(pres, p, 1);
        var left = p.left !== undefined && p.left !== null ? Number(p.left) : 100;
        var top = p.top !== undefined && p.top !== null ? Number(p.top) : 100;
        var width = p.width !== undefined && p.width !== null ? Number(p.width) : 400;
        var height = p.height !== undefined && p.height !== null ? Number(p.height) : 50;
        var shape = slide.Shapes.AddTextbox(1, left, top, width, height);
        if (p.text !== undefined && p.text !== null) {
          shape.TextFrame.TextRange.Text = String(p.text);
        }
        if (p.fontSize !== undefined && p.fontSize !== null) {
          shape.TextFrame.TextRange.Font.Size = Number(p.fontSize);
        }
        if (p.fontName) {
          shape.TextFrame.TextRange.Font.Name = String(p.fontName);
        }
        var name = '';
        try { name = String(shape.Name || ''); } catch (e) { /* 同上 */ }
        return { ok: true, data: { shapeName: name, slideIndex: p.slideIndex || 1 } };
      }, app, params);
    },

    /**
     * setShapePosition —— S1-REPORT.md:227（§3 #201：Left/Top/Width/Height 支持读回）
     * 只设调用方给了的维度（ps1 setShapePosition 同款语义）。
     */
    setShapePosition: function (app, params) {
      return guard(function (a, p) {
        var pres = getPres(a);
        var slide = getSlide(pres, p, 1);
        var shape = findShape(slide, p);
        if (!shape) {
          return { ok: false, error: '[勿重试·先改参数] 找不到目标形状：shapeIndex（1 起）/shapeName 都没给或不存在' };
        }
        if (p.left !== undefined && p.left !== null) shape.Left = Number(p.left);
        if (p.top !== undefined && p.top !== null) shape.Top = Number(p.top);
        if (p.width !== undefined && p.width !== null) shape.Width = Number(p.width);
        if (p.height !== undefined && p.height !== null) shape.Height = Number(p.height);
        return { ok: true, data: { left: p.left, top: p.top, width: p.width, height: p.height } };
      }, app, params);
    },

    /**
     * insertPptTable —— S1-REPORT.md:233（§3 #207：AddTable(rows,cols,l,t,w,h) 支持）
     * 默认 3×3 @ (100,100) 400×200 镜像 ps1；tableIndex 按 HasTable 计数定位
     * （算法与 ps1 同源：新表在 Shapes 末尾，名字比对 + 计数兜底）。
     */
    insertPptTable: function (app, params) {
      return guard(function (a, p) {
        var pres = getPres(a);
        var slide = getSlide(pres, p, 1);
        var rows = p.rows ? Number(p.rows) : 3;
        var cols = p.cols ? Number(p.cols) : 3;
        var left = p.left !== undefined && p.left !== null ? Number(p.left) : 100;
        var top = p.top !== undefined && p.top !== null ? Number(p.top) : 100;
        var width = p.width !== undefined && p.width !== null ? Number(p.width) : 400;
        var height = p.height !== undefined && p.height !== null ? Number(p.height) : 200;
        var table = slide.Shapes.AddTable(rows, cols, left, top, width, height);
        var newName = '';
        try { newName = String(table.Name || ''); } catch (e) { /* Name 读不到就走计数兜底 */ }
        var tableIndex = 0;
        var count = Number(slide.Shapes.Count);
        for (var i = 1; i <= count; i++) {
          var s = slide.Shapes.Item(i);
          var has = false;
          try { has = !!s.HasTable; } catch (e) { has = false; }
          if (has) {
            tableIndex += 1;
            if (newName) {
              var n = '';
              try { n = String(s.Name || ''); } catch (e) { n = ''; }
              if (n && n === newName) break;
            }
          }
        }
        // 没匹配上名字时 tableIndex 自然停在 HasTable 总数 —— 与 ps1 同款兜底
        // （新表追加在 Shapes 末尾，总数恰好就是新表的序号）。
        return { ok: true, data: { name: newName, tableIndex: tableIndex, slideIndex: p.slideIndex || 1 } };
      }, app, params);
    },

    /**
     * getPptTableCell —— S1-REPORT.md:235（§3 #209：Cell(r,c).Shape.TextFrame.TextRange.Text 支持）
     * value/text 双键回传（TS 侧 value 优先，table.ts:273-274）。
     */
    getPptTableCell: function (app, params) {
      return guard(function (a, p) {
        var pres = getPres(a);
        var slide = getSlide(pres, p, 1);
        var table = findTableShape(slide, p.tableIndex);
        if (!table) {
          return { ok: false, error: '[勿重试·先改参数] 找不到第 ' + (p.tableIndex || 1) + ' 个表格（按 HasTable 形状计数，1 起）' };
        }
        var text = table.Table.Cell(Number(p.row), Number(p.col)).Shape.TextFrame.TextRange.Text;
        text = text === null || text === undefined ? '' : String(text);
        return { ok: true, data: { value: text, text: text } };
      }, app, params);
    },

    /**
     * setPptTableCell —— S1-REPORT.md:234（§3 #208：同链路写入支持）
     */
    setPptTableCell: function (app, params) {
      return guard(function (a, p) {
        if (p.text === undefined || p.text === null) {
          return { ok: false, error: '[勿重试·先改参数] 缺少 text（setPptTableCell 必填）' };
        }
        var pres = getPres(a);
        var slide = getSlide(pres, p, 1);
        var table = findTableShape(slide, p.tableIndex);
        if (!table) {
          return { ok: false, error: '[勿重试·先改参数] 找不到第 ' + (p.tableIndex || 1) + ' 个表格（按 HasTable 形状计数，1 起）' };
        }
        table.Table.Cell(Number(p.row), Number(p.col)).Shape.TextFrame.TextRange.Text = String(p.text);
        return { ok: true, data: { row: Number(p.row), col: Number(p.col) } };
      }, app, params);
    },

    /**
     * addAnimation —— S1-REPORT.md:252/271/272（§3 #226/#245/#246：AddEffect 全可加）+ §5 报告行 355
     * 【只做加】的铁律：加完【不读】MainSequence.Count（S1 实测 Count 读回时好时坏，
     * 坏时 native signal 且桥死锁）——所以本动作的回包没有 count 字段（Windows ps1 有，这是有意的差异）。
     * 第三参 level 固定 0（msoAnimateLevelNone）：ps1 实测传 1 会建出两条重复效果。
     */
    addAnimation: function (app, params) {
      return guard(function (a, p) {
        var pres = getPres(a);
        var slide = getSlide(pres, p, 1);
        var shape = findShape(slide, p);
        if (!shape) {
          return { ok: false, error: '[勿重试·先改参数] 找不到要加动画的形状：shapeIndex（1 起）/shapeName 都没给或不存在' };
        }
        var raw = p.effect !== undefined && p.effect !== null ? p.effect : p.animationType; // 旧键名仍认（ps1 同款）
        if (raw === undefined || raw === null || raw === '') {
          return { ok: false, error: '[勿重试·先改参数] 缺少 effect（可用：appear/fadeIn/flyIn/wipeIn/zoomIn/pulse 等，或直接给 MsoAnimEffect 数字）' };
        }
        var effectId;
        if (typeof raw === 'number') {
          effectId = raw;
        } else if (/^\d+$/.test(String(raw))) {
          effectId = parseInt(String(raw), 10);
        } else {
          effectId = ANIM_EFFECTS[String(raw)];
          if (effectId === undefined) {
            return { ok: false, error: '[勿重试·先改参数] 不认识的 effect「' + raw + '」；常用：appear/fadeIn/flyIn/wipeIn/zoomIn（入场）、pulse/spin/grow/teeter（强调）。退场效果请直接给 MsoAnimEffect 数字' };
          }
        }
        slide.TimeLine.MainSequence.AddEffect(shape, effectId, 0);
        return { ok: true, data: { slideIndex: p.slideIndex || 1, effectId: effectId } }; // 无 count：S1 §5 读侧不稳，刻意不读
      }, app, params);
    },

    /**
     * setSlideTransition —— S1-REPORT.md:253（§3 #227：EntryEffect 编号直设支持）+ §5 报告行 354
     * 效果名→编号表抄自 ps1 setSlideTransition case；纯数字串/数字直通（0-80 全可设）。
     */
    setSlideTransition: function (app, params) {
      return guard(function (a, p) {
        var raw = p.effect !== undefined && p.effect !== null ? p.effect : p.transition; // 旧键名仍认（ps1 同款）
        var name = raw === undefined || raw === null || raw === '' ? 'fade' : String(raw).toLowerCase();
        var entry = -1;
        if (Object.prototype.hasOwnProperty.call(TRANSITION_EFFECTS, name)) {
          entry = TRANSITION_EFFECTS[name];
        } else if (/^\d+$/.test(name)) {
          entry = parseInt(name, 10);
        }
        if (entry < 0) {
          return { ok: false, error: '[勿重试·先改参数] 不认识的切换效果：' + name + '（可用：none/fade/push/wipe/split/reveal/cover/dissolve/curtains）' };
        }
        var pres = getPres(a);
        var slide = getSlide(pres, p, 1);
        var tr = slide.SlideShowTransition;
        tr.EntryEffect = entry;
        var dur = 0;
        if (p.duration !== undefined && p.duration !== null) {
          var n = Number(p.duration);
          if (!isNaN(n) && n > 0) dur = n;
        }
        if (dur > 0) {
          try { tr.Duration = dur; } catch (e) { /* Duration 不可设不算失败：主效果已生效，ps1 同款 try */ }
        }
        return { ok: true, data: { slideIndex: p.slideIndex || 1, effect: name, entryEffect: entry, duration: dur } };
      }, app, params);
    },

    /**
     * setSlideBackground —— S1-REPORT.md:254（§3 #228：Background.Fill.Solid/ForeColor 支持）+ S0 #18 读回验证
     * 颜色 #hex → COM RGB（编码同 Word 侧，S0 绿色读回 65280）。
     * 与 ps1 的差异（如实声明）：非法颜色 ps1 静默跳过仍回 success，这里改为明确报错——
     * "设了但没生效还报成功"正是 B2.5 任务卡禁止的静默成功。
     */
    setSlideBackground: function (app, params) {
      return guard(function (a, p) {
        if (!p.color) {
          return { ok: false, error: '[勿重试·先改参数] 缺少 color（setSlideBackground 本轮只实现纯色路径，#RRGGBB）' };
        }
        var rgb = hexToRgbInt(String(p.color));
        if (rgb === null) {
          return { ok: false, error: '[勿重试·先改参数] 不认识的颜色「' + p.color + '」（要 #RRGGBB 或 #RGB 三位）' };
        }
        var pres = getPres(a);
        var slide = getSlide(pres, p, 1);
        var fill = slide.Background.Fill;
        fill.Solid();
        fill.ForeColor.RGB = rgb;
        return { ok: true, data: { slideIndex: p.slideIndex || 1, colorRgb: rgb } };
      }, app, params);
    },

    /**
     * exportSlideAsImage —— S1-REPORT.md:232（§3 #206）+ S0 #19（真实生成 1334×750 PNG 验证）
     * Export(path, filter, width?, height?)：尺寸给了就带（S0 实测带尺寸调用成功）。
     */
    exportSlideAsImage: function (app, params) {
      return guard(function (a, p) {
        var path = p.outputPath || p.path; // TS 两个键都发（ppt/image.ts:393-399）
        if (!path) {
          return { ok: false, error: '[勿重试·先改参数] 缺少 outputPath/path（exportSlideAsImage 必填）' };
        }
        var format = p.format || 'PNG';
        var pres = getPres(a);
        var slide = getSlide(pres, p, 1);
        if (p.width !== undefined && p.width !== null && p.height !== undefined && p.height !== null) {
          slide.Export(String(path), String(format), Number(p.width), Number(p.height));
        } else {
          slide.Export(String(path), String(format));
        }
        return { ok: true, data: { slideIndex: p.slideIndex || 1, outputPath: String(path), format: String(format) } };
      }, app, params);
    },
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = pptActions;
  } else {
    global.WpsAddonActionsPpt = pptActions;
  }
})(typeof window !== 'undefined' ? window : this);
