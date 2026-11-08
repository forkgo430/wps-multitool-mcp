/**
 * B5T4 同步门面动作层 · Ppt 型（新增文件，11.x 异步基线不动）
 *
 * 规格来源：S1-REPORT.md §3 逐条表（判定列 + 对应的 JSAPI 操作列 + 证据列）。
 * 实现体全部为【同步门面直调】window.Application 对象模型；回包结构与包内
 * actions/*.js 的 guard 形态一致：Promise<{ok,data}|{ok:false,error}>，异常不吞。
 * S1 §4 语义差异已适配：Cells/Rows/Columns/Names 用 .Item()；写值用 Value2；
 * 读值用 Value()；Selection.Address()；PivotTables.Count 如实；AddEffect 禁读数量。
 */
(function (global) {
  'use strict';
  var RT = global.WpsAddonSyncRuntime;
  if (!RT) { global.WpsAddonActionsSyncPpt = {}; return; }
  var SPECS = {
    createPresentation: { v: "部分支持", jsapi: "app.Presentations.Add()", evidence: "仅 wpp 型: app.Presentations.Add() 支持; wps/et 型 Presentations 不存在 → 加载项类型隔离所致", code: "\nvar pr=c.app.Presentations.Add();\nreturn '已建 name='+pr.Name+' Count='+R.t(c.app.Presentations.Count);\n" },
    getSlideCount: { v: "支持", jsapi: "pres.Slides.Count", evidence: "string:Count=1", code: "\nreturn 'Count='+R.t(c.needPres().Slides.Count);\n" },
    addSlide: { v: "支持", jsapi: "pres.Slides.Add(index, layout)", evidence: "string:已加; Count=2", code: "\nvar s=c.needPres().Slides;\nvar before=R.t(s.Count);\nvar sl=s.Add(s.Count+1,1);\nreturn '已加; 前='+before+' 后='+R.t(s.Count);\n" },
    deleteSlide: { v: "支持", jsapi: "slide.Delete()", evidence: "string:已删末页; 剩余=2", code: "\nvar s=c.needPres().Slides;\nif(s.Count<2){ s.Add(2,1); }\nvar before=R.t(s.Count);\nR.item(s,s.Count).Delete();\nreturn '已删; 前='+before+' 后='+R.t(s.Count);\n" },
    getShapes: { v: "支持", jsapi: "遍历 Shapes: Type/Name", evidence: "string:Count=3 [#1:T14,#2:T14,#3:T1]", code: "\nvar sl=R.item(c.needPres().Slides,1);\nvar sh=sl.Shapes, a=[];\nfor(var i=1;i<=Math.min(sh.Count,6);i++){ var x=R.item(sh,i); a.push(R.t(x.Name)+'/'+R.t(x.Type)); }\nreturn 'Count='+R.t(sh.Count)+' ['+a.join(', ')+']';\n" },
    setShapeText: { v: "支持", jsapi: "shape.TextFrame.TextRange.Text", evidence: "string:读回=string:探针形状文字", code: "\nvar _sl=R.item(c.needPres().Slides,1);\nvar _sh=_sl.Shapes;\nif(!_sh || !_sh.Count){ _sh=_sl.Shapes; _sh.AddTextbox(1,50,50,200,50); }\n\nvar x=R.item(_sh,p.index||1);\nx.TextFrame.TextRange.Text=p.text||'B5T4';\nreturn 'Text='+R.t(x.TextFrame.TextRange.Text);\n" },
    addTextBox: { v: "支持", jsapi: "slide.Shapes.AddTextbox(orientation, l, t, w, h)", evidence: "string:已加; Shapes=3", code: "\nvar sl=R.item(c.needPres().Slides,1);\nvar x=sl.Shapes.AddTextbox(1,50,50,200,50);\nreturn '已加 Name='+R.t(x.Name)+' Count='+R.t(sl.Shapes.Count);\n" },
    setShapePosition: { v: "支持", jsapi: "shape.Left/Top/Width/Height", evidence: "string:L=300 T=250 W=250 H=120", code: "\nvar _sl=R.item(c.needPres().Slides,1);\nvar _sh=_sl.Shapes;\nif(!_sh || !_sh.Count){ _sh=_sl.Shapes; _sh.AddTextbox(1,50,50,200,50); }\n\nvar x=R.item(_sh,p.index||1);\nx.Left=p.left||100; x.Top=p.top||100;\nreturn 'Left='+R.t(x.Left)+' Top='+R.t(x.Top)+' Width='+R.t(x.Width)+' Height='+R.t(x.Height);\n" },
    insertPptTable: { v: "支持", jsapi: "slide.Shapes.AddTable(rows, cols, l, t, w, h)", evidence: "string:已建表", code: "\nvar sl=R.item(c.needPres().Slides,1);\nvar x=sl.Shapes.AddTable(2,2,50,120,200,100);\nreturn '已建表; Shapes.Count='+R.t(sl.Shapes.Count);\n" },
    getPptTableCell: { v: "支持", jsapi: "table.Cell(r,c).Shape.TextFrame.TextRange.Text", evidence: "string:string:探针单元格", code: "\nvar _sl=R.item(c.needPres().Slides,1);\nvar _sh=_sl.Shapes;\nif(!_sh || !_sh.Count){ _sh=_sl.Shapes; _sh.AddTextbox(1,50,50,200,50); }\n\nvar n=_sh.Count, cell=null;\nfor(var i=1;i<=n;i++){ var x=R.item(_sh,i); if(R.has(x,'Table')&&x.Table){ cell=x.Table.Cell(1,1); break; } }\nif(!cell) return '幻灯片上无表格对象（Shapes.Count='+R.t(n)+'）';\nreturn 'Cell(1,1).Text='+R.t(cell.Shape.TextFrame.TextRange.Text);\n" },
    setPptTableCell: { v: "支持", jsapi: "table.Cell(r,c).Shape.TextFrame.TextRange.Text", evidence: "string:已写", code: "\nvar _sl=R.item(c.needPres().Slides,1);\nvar _sh=_sl.Shapes;\nif(!_sh || !_sh.Count){ _sh=_sl.Shapes; _sh.AddTextbox(1,50,50,200,50); }\n\nvar n=_sh.Count, cell=null;\nfor(var i=1;i<=n;i++){ var x=R.item(_sh,i); if(R.has(x,'Table')&&x.Table){ cell=x.Table.Cell(1,1); break; } }\nif(!cell) return '幻灯片上无表格对象（Shapes.Count='+R.t(n)+'）';\ncell.Shape.TextFrame.TextRange.Text=p.text||'B5T4';\nreturn '写入读回='+R.t(cell.Shape.TextFrame.TextRange.Text);\n" },
    addAnimation: { v: "支持", jsapi: "slide.TimeLine.MainSequence.AddEffect(shape, effectId, …)", evidence: "string:AddEffect 已调用; 返回=object:[object Object]", code: "\nvar _sl=R.item(c.needPres().Slides,1);\nvar _sh=_sl.Shapes;\nif(!_sh || !_sh.Count){ _sh=_sl.Shapes; _sh.AddTextbox(1,50,50,200,50); }\n\n/* 任务书特殊断言 R13：只 AddEffect，绝不读 MainSequence.Count\n   —— S1 §8 #3：Count GET 会触发 native signal 且导致桥死锁 */\nvar sl=_sl, sh=R.item(_sh,p.index||1);\nvar ms=sl.TimeLine.MainSequence;\nvar e=ms.AddEffect(sh,p.effect||21,0,1);\nreturn 'AddEffect 已调用（effect='+R.t(p.effect||21)+'）；按 R13 禁读 MainSequence.Count';\n" },
    setSlideTransition: { v: "支持", jsapi: "slide.SlideShowTransition.EntryEffect = id", evidence: "string:已设; 读回=2", code: "\nvar sl=R.item(c.needPres().Slides,1);\nsl.SlideShowTransition.EntryEffect=p.effect||1;\nreturn 'EntryEffect='+R.t(sl.SlideShowTransition.EntryEffect);\n" },
    setSlideBackground: { v: "支持", jsapi: "slide.Background.Fill.Solid/ForeColor", evidence: "string:已设背景色; 读回=number:0", code: "\nvar sl=R.item(c.needPres().Slides,1);\nvar f=sl.Background.Fill;\nf.Solid();\nreturn 'Background.Fill.Solid 已调用; Type='+R.t(f.Type);\n" },
    exportSlideAsImage: { v: "支持", jsapi: "slide.Export(path, 'PNG')", evidence: "string:Export 已调用", code: "\nvar sl=R.item(c.needPres().Slides,1);\nvar path=p.path||'/tmp/wps_b38t4/targets/slide_export.png';\nsl.Export(path,'PNG');\nreturn 'Export -> '+path;\n" },
  };
  global.WpsAddonActionsSyncPpt = RT.build(SPECS);
})(typeof window !== 'undefined' ? window : this);
