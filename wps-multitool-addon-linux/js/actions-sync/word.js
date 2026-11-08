/**
 * B5T4 同步门面动作层 · Word 型（新增文件，11.x 异步基线不动）
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
  if (!RT) { global.WpsAddonActionsSyncWord = {}; return; }
  var SPECS = {
    openDocument: { v: "支持", jsapi: "app.Documents.Open(open_common_wps.docx)", evidence: "string:打开成功 name=open_common_wps.docx", code: "\nif(!p.path && !p.filePath) throw new Error('[勿重试·先改参数] 缺少 path/filePath（openDocument 必填）');\nvar d=c.app.Documents.Open(p.path||p.filePath);\nreturn '打开成功 name='+d.Name+' Count='+R.t(c.app.Documents.Count);\n" },
    getActiveDocumentState: { v: "未在S1表中", jsapi: "", evidence: "", code: "\nvar d=c.needDoc();\nreturn 'name='+d.Name+' FullName='+R.t(d.FullName)+' Content='+R.t(d.Content.Text).slice(0,60)+' Count='+R.t(c.app.Documents.Count);\n" },
    getSelectedText: { v: "部分支持", jsapi: "app.Selection.Text / .Address()", evidence: "wps: Text=string:O(支持); et: Address 是函数 sel.Address()(支持); wpp: 宿主无 Selection(不支持) → 仅 wpp", code: "\nif(!c.ws){ try{ c.wb=c.app.ActiveWorkbook; }catch(e){} try{ c.ws=c.wb?c.wb.ActiveSheet:null; }catch(e){} }\n\nif(!c.sel) return 'Excel 上下文无 Selection.Text（S1 §4 #14：无 TypeText）';\nreturn 'Address='+R.t(R.addr(c.sel))+' Text='+R.t(c.sel.Text);\n" },
    replaceRange: { v: "支持", jsapi: "doc.Range(a,b).Text = '…'", evidence: "string:读回=string:范围替换", code: "\nvar d=c.needDoc();\nvar r=d.Range(p.start||0,p.end||0);\nr.Text=p.text||'B5T4';\nreturn 'Text='+R.t(r.Text);\n" },
    setFont: { v: "支持", jsapi: "app.Selection.Font.Size/Name", evidence: "string:已设 Size=24 Name=SimSun; 读回 Size=24 Name=SimSun", code: "\nvar f=c.app.Selection.Font;\nf.Size=p.size||24; f.Name=p.name||'SimSun';\nreturn '已设 Size='+R.t(f.Size)+' Name='+R.t(f.Name)+'；读回 Size='+R.t(f.Size)+' Name='+R.t(f.Name);\n" },
    closeDocument: { v: "部分支持", jsapi: "doc.Close()", evidence: "doc.Close() 对干净文档可关; dirty 文档弹'是否保存'模态框阻塞 JS 线程直至用户点击(D/E 测试截图证据) → 部分支持", code: "\n/* S1 §3 #109：dirty 文档 Close() 会弹「是否保存」模态框并阻塞 JS 线程；\n   先置 Saved=true 再 Close(0)=wdDoNotSaveChanges，避免弹框。 */\nvar n=c.app.Documents.Count;\nvar d=c.needDoc();\ntry{ d.Saved=true; }catch(e){}\ntry{ d.Close(0); }catch(e){ try{ d.Close(); }catch(e2){} }\nif(c.app.Documents.Count===0){ c.app.Documents.Add(); }\nreturn '已关; 关闭前='+R.t(n)+' 剩余='+R.t(c.app.Documents.Count)+'（Saved=true + Close(0) 不保存；关闭后补空文档）';\n" },
    getDocumentStats: { v: "部分支持", jsapi: "doc.Words/Characters/Paragraphs/Sections/Pages.Count", evidence: "string:{\"Words\":1,\"Characters\":1,\"Paragraphs\":1,\"Sections\":1,\"Pages\":\"属性不存在\"} 【复核】 4/5 统计项", code: "\nvar d=c.needDoc(), o={};\n['Words','Characters','Paragraphs','Sections','Pages'].forEach(function(k){\n  try{ o[k]=R.has(d,k)?d[k].Count:'属性不存在'; }catch(e){ o[k]='ERR:'+e.message; }\n});\nreturn JSON.stringify(o);\n" },
    replaceBookmarkContent: { v: "部分支持", jsapi: "doc.Bookmarks(name).Range.Text = '…'", evidence: "TypeError: doc.Bookmarks is not a function 【复核】 doc.Bookmarks 集合存在(Item 可按名取到书签, getBookma", code: "\nvar d=c.needDoc(), b=d.Bookmarks;\nif(!R.has(b,'Item')) return 'doc.Bookmarks.Item 不存在（S1 §3 #131 部分支持）';\nvar x=R.item(b,p.name||'b1');\nif(!x) return '书签 '+R.t(p.name||'b1')+' 不存在（Bookmarks.Item 返回 null；S1 §3 #131 部分支持）; Count='+R.t(b.Count);\nx.Range.Text=p.text||'B5T4';\nreturn '已替换书签内容; Bookmarks.Count='+R.t(b.Count);\n" },
    modifyStyle: { v: "部分支持", jsapi: "style.Font.Size", evidence: "TypeError: doc.Styles is not a function 【复核】 doc.Styles 集合存在(Item 可取); 样式修改未单独验证 → 部分支持", code: "\nvar s=c.needDoc().Styles;\nvar st=R.item(s,p.name||'Normal');\nst.Font.Size=p.size||12;\nreturn 'Style='+R.t(st.Name)+' Size='+R.t(st.Font.Size);\n" },
    deleteStyle: { v: "部分支持", jsapi: "style.Delete()", evidence: "TypeError: doc.Styles is not a function 【复核】 doc.Styles 集合存在(Item 可取); 样式删除未单独验证 → 部分支持", code: "\nvar s=c.needDoc().Styles;\nvar st=R.item(s,p.name||'Normal');\nvar n=R.t(s.Count);\nif(!st) return '样式 '+R.t(p.name||'Normal')+' 不存在（Styles.Item 返回 null）; Count='+n;\nst.Delete();\nreturn 'Delete 已调用; 删除前 Count='+n+' 后='+R.t(s.Count);\n" },
    setDefaultFont: { v: "部分支持", jsapi: "doc.Styles('Normal').Font.Name", evidence: "TypeError: doc.Styles is not a function 【复核】 doc.Styles 集合存在(Item 可取, listStyles 实测); 默认字体", code: "\nvar s=c.needDoc().Styles;\nvar st=R.item(s,'Normal');\nst.Font.Name=p.name||'SimSun';\nreturn 'Normal.Font.Name='+R.t(st.Font.Name);\n" },
    setSelectedText: { v: "部分支持", jsapi: "app.Selection.TypeText", evidence: "wps: TypeText 已调用读回=string:O(支持); et: 无 TypeText 改走单元格写入(支持); wpp: 宿主无 Selection(不支持) → 仅 ", code: "\nvar s=c.app.Selection;\ns.TypeText(p.text||'B5T4');\nreturn 'TypeText 已调用, 读回='+R.t(s.Text).slice(0,60);\n" },
    createDocument: { v: "部分支持", jsapi: "app.Documents.Add()", evidence: "仅 wps 型: app.Documents.Add() 支持; et/wpp 型 Documents 不存在 → 加载项类型隔离所致", code: "\nvar d=c.app.Documents.Add();\nreturn '已建 name='+d.Name+' Count='+R.t(c.app.Documents.Count);\n" },
  };
  global.WpsAddonActionsSyncWord = RT.build(SPECS);
})(typeof window !== 'undefined' ? window : this);
