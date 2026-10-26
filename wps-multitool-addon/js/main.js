/**
 * Input: WPS 加载项回调与用户交互请求
 * Output: 状态展示、快照/还原、导出与日志输出
 * Pos: Windows WPS 加载项脚本入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 多面手 MCP - WPS 加载项（v2.0）
 * MCP Server 通过 PowerShell COM 桥接直接操作 WPS；本加载项给人提供
 * 「看得见、能接管、能回退」的入口 —— 这正是 MCP 工具做不到的那一环。
 *
 * 功能矩阵（9 个按钮，见 ribbon.xml）：
 *   窗口：置于前台 / 面板
 *   快照：保存快照 / 还原快照 / 撤销改动
 *   输出：导出 PDF
 *   维护：刷新状态 / 解除卡死 / 关于
 *
 * 已移除：「导出本页」「另存副本」（功能区与面板的入口、代码、注释全部清掉）。
 *
 * 已知限制（如实说明，别让使用者以为有而实际没有）：
 *   1) 「面板」用的是 wps.CreateTaskPane —— 是真任务窗格，不是对话框。
 *      ⚠ 早期注释里写过"没有侧边栏 API"，那是没验证就下的结论，已证伪。
 *   2) 文件系统操作走 ActiveX FSO，WPS 若禁用 ActiveX 则快照功能不可用
 *      （会明确报错，不静默失败）。
 *
 * 当前维护者：赢容
 */

// ==================== 全局状态 ====================
var addonStatus = {
    initialized: false,
    lastRefresh: null,
    lastAction: '(未操作)'
};

var LINE_TOP = '┌─ 多面手 MCP ─────────────────';
var LINE_BOT = '└────────────────────────────';

// ==================== 工具函数 ====================

/** 取当前活动文档信息（Word / Excel / PPT 通用） */
function getActiveDocInfo() {
    if (typeof Application === 'undefined') { return null; }
    try {
        if (Application.ActiveWorkbook) {
            var wb = Application.ActiveWorkbook;
            return { name: wb.Name, type: 'WPS 表格', extra: '工作表数: ' + wb.Sheets.Count, fullPath: wb.FullName || '' };
        }
    } catch (e) {}
    try {
        if (Application.ActivePresentation) {
            var pres = Application.ActivePresentation;
            return { name: pres.Name, type: 'WPS 演示', extra: '幻灯片数: ' + pres.Slides.Count, fullPath: pres.FullName || '' };
        }
    } catch (e) {}
    try {
        if (Application.ActiveDocument) {
            var d = Application.ActiveDocument;
            var extra = '';
            try { extra = '字数: ' + d.Words.Count; } catch (e2) {}
            return { name: d.Name, type: 'WPS 文字', extra: extra, fullPath: d.FullName || '' };
        }
    } catch (e) {}
    return null;
}

/** 时间戳，用于快照/副本命名 */
function stamp() {
    var d = new Date();
    function p(n) { return (n < 10 ? '0' : '') + n; }
    return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' +
           p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds());
}

/** 文件系统（ActiveX FSO）。拿不到就返回 null，调用方要如实告知。 */
function getFso() {
    try { return new ActiveXObject('Scripting.FileSystemObject'); } catch (e) { return null; }
}

/** 快照目录：文档同目录下的 .mcp-snapshots\ */
function snapshotDir(doc) {
    if (!doc || !doc.fullPath) { return null; }
    var dir = doc.fullPath.substring(0, doc.fullPath.lastIndexOf('\\') + 1);
    return dir + '.mcp-snapshots\\';
}

/** 列出某目录下的快照文件（按时间倒序） */
function listSnapshots(dir) {
    var fso = getFso();
    if (!fso || !dir) { return []; }
    var out = [];
    try {
        if (!fso.FolderExists(dir)) { return []; }
        var files = new Enumerator(fso.GetFolder(dir).Files);
        for (; !files.atEnd(); files.moveNext()) {
            var f = files.item();
            out.push({ name: f.Name, path: f.Path, size: f.Size, date: f.DateLastModified });
        }
    } catch (e) {}
    out.sort(function (a, b) { return b.date - a.date; });
    return out;
}

// ==================== ribbon.xml 回调函数 ====================

function OnAddinLoad(ribbonUI) {
    console.log('[多面手MCP] 加载项已加载 - COM 桥接模式');
    addonStatus.initialized = true;
}

// ---------- 窗口组 ----------

// 置于前台
function ribbonShowWindow(control) {
    var lines = [LINE_TOP];
    try {
        if (typeof Application === 'undefined') {
            lines.push('  ⚠ 无法获取 WPS 应用对象');
        } else {
            try { Application.Visible = true; } catch (e1) {}
            try { if (Application.WindowState === 2) { Application.WindowState = 1; } } catch (e3) {}
            var activated = false;
            try {
                if (Application.Activate) { Application.Activate(); activated = true; }
                else if (Application.ActiveWindow && Application.ActiveWindow.Activate) { Application.ActiveWindow.Activate(); activated = true; }
            } catch (e2) {}
            lines.push(activated ? '  ● 已将 WPS 置于前台' : '  ⚠ 窗口激活接口不可用，已恢复可见');
        }
    } catch (e) {
        lines.push('  ⚠ 出错: ' + e.message);
    }
    lines.push(LINE_BOT);
    alert(lines.join('\n'));
}

// 面板：打开【真任务窗格】（WPS JS 加载项支持 wps.CreateTaskPane）
// ----------------------------------------------------------------------
// 历史教训：我第一版写的是"WPS JS 加载项没有侧边栏 API，做不了真任务面板"，
// 于是把面板降级成 alert 对话框。那句话是【没验证就下的结论】——
// WPS 官方社区有现成示例：wps.CreateTaskPane(url, title) + DockPosition/Width/Visible。
// 设计稿（多面手MCP_加载项UI_v2）画的本来就是常驻任务面板，alert 根本做不到。
//
// 面板页面：panel.html（同目录）。它通过桥的 HTTP 端点轮询状态
// （任务窗格是内嵌浏览器，连不到命名管道）。
function ribbonPanel(control) {
    var url = Application.Env.GetAppDataPath() + '/kingsoft/wps/jsaddons/wps-claude-addon_/panel.html';
    var existing = null;
    try { existing = wps.GetTaskPane('mcp-panel'); } catch (e) { existing = null; }

    if (existing) {
        try { existing.Visible = true; } catch (e2) {}
        try { existing.Activate(); } catch (e3) {}
        return;
    }
    try {
        var pane = wps.CreateTaskPane(url, '多面手 MCP');
        pane.ID = 'mcp-panel';
        pane.Width = 460;
        try { pane.DockPosition = wps.Enum.JSKsoEnum_msoCTPDockPositionRight; } catch (e4) {}
        pane.Visible = true;
        addonStatus.lastAction = '打开面板';
    } catch (e) {
        alert('打开任务面板失败: ' + e.message +
              '\n\n面板页面应位于:\n' + url);
    }
}

// ---------- 快照组 ----------

// 保存快照：把当前文档复制到 .mcp-snapshots\<原名>_<时间戳>
function ribbonSnapshotSave(control) {
    var lines = [LINE_TOP, '  保存快照', LINE_BOT, ''];
    var doc = getActiveDocInfo();
    if (!doc) { lines.push('  ⚠ 没有打开的文档'); alert(lines.join('\n')); return; }
    if (!doc.fullPath) {
        lines.push('  ⚠ 文档尚未保存到磁盘，无法建快照。');
        lines.push('  请先保存：Ctrl+S，或让 MCP 调 wps_common_save。');
        alert(lines.join('\n'));
        return;
    }
    var fso = getFso();
    if (!fso) {
        lines.push('  ⚠ 无法创建文件系统对象（ActiveX 被禁用？）');
        lines.push('  快照功能需要 WPS 允许 ActiveX。');
        alert(lines.join('\n'));
        return;
    }
    var dir = snapshotDir(doc);
    try {
        if (!fso.FolderExists(dir)) { fso.CreateFolder(dir); }
        var base = doc.name.replace(/\.[^.]+$/, '');
        var target = dir + base + '_' + stamp() + (doc.name.match(/\.[^.]+$/) || [''])[0];
        fso.CopyFile(doc.fullPath, target, true);
        addonStatus.lastAction = '保存快照 ' + stamp();
        lines.push('  ● 已建快照');
        lines.push('    ' + target);
        lines.push('  大小: ' + Math.round(fso.GetFile(target).Size / 1024) + ' KB');
        lines.push('');
        lines.push('  提示：AI 改坏文档后，用「还原快照」退回。');
    } catch (e) {
        lines.push('  ✗ 快照失败: ' + e.message);
    }
    alert(lines.join('\n'));
}

// 还原快照：列出最近的快照，让用户确认后覆盖当前文档
function ribbonSnapshotRestore(control) {
    var lines = [LINE_TOP, '  还原快照', LINE_BOT, ''];
    var doc = getActiveDocInfo();
    if (!doc) { lines.push('  ⚠ 没有打开的文档'); alert(lines.join('\n')); return; }
    var dir = snapshotDir(doc);
    var snaps = listSnapshots(dir);
    if (!snaps.length) {
        lines.push('  ⚠ 该文档还没有快照');
        lines.push('    ' + (dir || '(文档未保存)'));
        lines.push('');
        lines.push('  先用「保存快照」建一份。');
        alert(lines.join('\n'));
        return;
    }
    lines.push('  最近 5 份快照：');
    for (var i = 0; i < Math.min(5, snaps.length); i++) {
        var s = snaps[i];
        lines.push('   ' + (i + 1) + ') ' + s.name);
        lines.push('       ' + Math.round(s.size / 1024) + ' KB  ' + s.date.toLocaleString());
    }
    lines.push('');
    lines.push('  ⚠ 还原会【覆盖当前文档】，且不可撤销。');
    lines.push('  确认还原请输入序号 1-' + Math.min(5, snaps.length) + '：');
    var ans = prompt(lines.join('\n'), '');
    if (!ans) { return; }
    var idx = parseInt(ans, 10);
    if (isNaN(idx) || idx < 1 || idx > Math.min(5, snaps.length)) {
        alert('序号无效，已取消。');
        return;
    }
    var pick = snaps[idx - 1];
    try {
        // 先存一份"还原前的现场"，避免还原错了没处退
        var fso = getFso();
        if (fso && doc.fullPath) {
            var backup = dir + '_还原前_' + stamp() + '.bak';
            fso.CopyFile(doc.fullPath, backup, true);
            lines.push('  已备份还原前的文档: ' + backup);
        }
        fso.CopyFile(pick.path, doc.fullPath, true);
        addonStatus.lastAction = '还原快照 ' + pick.name;
        // 关闭当前文档让用户重开，否则内存里还是旧内容
        try { doc.close ? doc.close() : Application.ActiveDocument.Close(0); } catch (e2) {}
        alert(lines.join('\n') + '\n\n  ● 已还原，文档已关闭，请重新打开。');
    } catch (e) {
        alert('  还原失败: ' + e.message);
    }
}

// 撤销改动：先试文档级 UndoRecord，再退回 Application.Undo
function ribbonUndo(control) {
    var lines = [LINE_TOP, '  撤销改动', LINE_BOT, ''];
    try {
        var n = 0;
        try { n = Application.UndoRecord.UndoRecordCount || 0; } catch (e1) {}
        lines.push('  可撤销记录数 : ' + n);
        if (n > 0) {
            var ans = prompt('要撤销几步？（1-' + n + '，留空取消）', '1');
            var k = parseInt(ans || '0', 10);
            if (k > 0 && k <= n) {
                for (var i = 0; i < k; i++) { Application.UndoRecord.Undo(); }
                addonStatus.lastAction = '撤销 ' + k + ' 步';
                lines.push('  ● 已撤销 ' + k + ' 步');
            } else {
                lines.push('  已取消');
            }
        } else {
            lines.push('  当前没有可撤销的记录');
        }
    } catch (e) {
        lines.push('  ⚠ 撤销失败: ' + e.message);
        lines.push('  （部分 WPS 版本不支持 UndoRecord，');
        lines.push('   可改用 Ctrl+Z 逐条撤销）');
    }
    alert(lines.join('\n'));
}

// ---------- 输出组 ----------

// 导出 PDF（用 WPS 自带的导出，不走 MCP）
function ribbonExportPdf(control) {
    var lines = [LINE_TOP, '  导出 PDF', LINE_BOT, ''];
    var doc = getActiveDocInfo();
    if (!doc) { lines.push('  ⚠ 没有打开的文档'); alert(lines.join('\n')); return; }
    try {
        var out = doc.fullPath
            ? doc.fullPath.replace(/\.[^.]+$/, '') + '_' + stamp() + '.pdf'
            : 'C:\\' + (doc.name.replace(/\.[^.]+$/, '') || '导出') + '_' + stamp() + '.pdf';
        // wdExportFormatPDF = 17；Excel/PowerPoint 用 ExportAsFixedFormat(xlPDF=0)
        var ok = false;
        try { doc.ExportAsFixedFormat(out, 17, false); ok = true; } catch (e1) {
            try { Application.ActiveDocument.ExportAsFixedFormat(out, 17, false); ok = true; } catch (e2) {}
        }
        if (ok) {
            addonStatus.lastAction = '导出 PDF';
            lines.push('  ● 已导出:');
            lines.push('    ' + out);
        } else {
            lines.push('  ⚠ 本机 WPS 不支持 ExportAsFixedFormat');
            lines.push('  可改用 MCP 工具 wps_convert_to_pdf。');
        }
    } catch (e) {
        lines.push('  ✗ 导出失败: ' + e.message);
    }
    alert(lines.join('\n'));
}

// ---------- 维护组 ----------

// 刷新状态
function ribbonRefresh(control) {
    addonStatus.lastRefresh = new Date();
    addonStatus.lastAction = '刷新状态';
    var lines = [LINE_TOP, '  状态已刷新', LINE_BOT, ''];
    try {
        var doc = getActiveDocInfo();
        lines.push('  时间 : ' + addonStatus.lastRefresh.toLocaleTimeString());
        lines.push('  应用 : ' + (Application.Name || '?'));
        lines.push('  文档 : ' + (doc ? doc.name + '（' + doc.type + '）' : '无'));
    } catch (e) {
        lines.push('  ⚠ ' + e.message);
    }
    alert(lines.join('\n'));
}

// 解除卡死：把 WPS 的计算/重算关掉再打开，清掉挂起状态
function ribbonUnstick(control) {
    var lines = [LINE_TOP, '  解除卡死', LINE_BOT, ''];
    var tried = [];
    try {
        try { Application.Calculation = -4135 /* xlCalculationManual */; tried.push('计算改为手动'); } catch (e1) {}
        try { Application.ScreenUpdating = false; tried.push('关闭屏幕刷新'); } catch (e2) {}
        try {
            // 恢复回去，目的是让 WPS 重走一遍内部状态
            Application.Calculation = -4105; /* xlCalculationAutomatic */
            Application.ScreenUpdating = true;
            tried.push('已恢复正常设置');
        } catch (e3) { tried.push('恢复正常设置失败（可忽略）'); }
        try { if (Application.ActiveDocument) { Application.ActiveDocument.Undo(); } } catch (e4) {}
        lines.push('  已尝试:');
        for (var i = 0; i < tried.length; i++) { lines.push('   · ' + tried[i]); }
        lines.push('');
        lines.push('  若仍卡：');
        lines.push('   1) 关掉 WPS 里的模态对话框（另存确认等）');
        lines.push('   2) 用 MCP 工具 wps_common_repair 诊断');
        lines.push('   3) 都不行就重启 WPS');
    } catch (e) {
        lines.push('  ⚠ 失败: ' + e.message);
    }
    addonStatus.lastAction = '解除卡死';
    alert(lines.join('\n'));
}

// 关于
function ribbonAbout(control) {
    // 「关于」已搬进任务面板底部（原生 alert 排版被批难看）。
    // 这里直接打开面板，让用户看到完整信息。
    ribbonPanel(control);
}


// ==================== WPS 生命周期函数 ====================

function OnAddinNotify(params) {
    console.log('[多面手MCP] 收到通知:', JSON.stringify(params));
}

// ==================== 初始化 ====================
console.log('[多面手MCP] main.js 已加载 v2.0 - COM 桥接模式');
