/**
 * Input: 无（服务端地址/令牌来自 URL query 或 window.WPS_ADDON_CONFIG，见 channel-client.js）
 * Output: 启动 wps 型加载项的通道轮询循环；功能区按钮的 onAction 入口
 * Pos: Linux 加载项（et 型 = 表格）的启动脚本。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 为什么型别写死：S1 实测三型对象模型互相看不见（wps 型只有 Documents），
 * 每个加载项目录只能拿一种应用的对象模型，型别在部署时就是确定的。
 */
(function () {
    'use strict';

    function boot() {
        if (window.WpsAddonChannel) {
            window.WpsAddonChannel.start('et');
        } else {
            // channel-client.js 未加载（相对路径被宿主改写等）——显式失败，不静默
            var el = document.getElementById('channel-state');
            if (el) el.textContent = '启动失败：找不到 WpsAddonChannel（检查 manifest 的 <script src="../js/channel-client.js"> 是否可达）';
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }

    // 功能区「连接状态」按钮：把通道状态亮给用户看（Linux 上加载项要手动打开，
    // 用户需要一个能自己确认"开没开、连没连上"的入口 —— 计划 §12.2）
    window.ribbonChannelStatus = function () {
        var s = window.WpsAddonChannel ? window.WpsAddonChannel.state() : { connected: false, lastError: '通道客户端未加载' };
        window.alert(
            'MCP · Linux 加载项通道（et 型）\n' +
            '连接：' + (s.connected ? '已连接' : '未连接') + '\n' +
            '已处理动作：' + (s.handled || 0) + '\n' +
            '最近错误：' + (s.lastError || '（无）') + '\n\n' +
            '未连接时检查：1) Node 侧服务已启动（端口/令牌）；2) 本加载项已从「工具 → WPS 加载项」打开。'
        );
    };
})();
