#!/usr/bin/env node
/**
 * Input: 运行环境与启动参数
 * Output: MCP Server 导出与启动日志
 * Pos: MCP 服务入口模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS Office MCP Server - 入口文件
 * 赢容出品，让AI能操控WPS Office
 *
 * 启动方式：
 * - 开发模式：npm run dev
 * - 生产模式：npm run build && npm start
 * - npx 模式：npx wps-multitool-mcp（靠 package.json 的 bin 指向本文件）
 *
 * 注意 shebang 必须在第一行（tsc 会原样保留）：npm 的 bin 在 Unix 下
 * 靠它找解释器，删了这一行 `npx wps-multitool-mcp` 会报 SyntaxError。
 */

import { createMcpServer } from './server/mcp-server';
import { log, createChildLogger } from './utils/logger';
import { psDaemon } from './client/ps-daemon';
import { MCP_NAME, MCP_VERSION } from './version';

// 导出所有模块，方便外部使用
export { WpsMcpServer, createMcpServer } from './server/mcp-server';
export { ToolRegistry, toolRegistry, registerTool } from './server/tool-registry';
export { WpsClient, wpsClient } from './client/wps-client';
export { log, logger, createChildLogger } from './utils/logger';
export * from './utils/error';
export * from './types/tools';
export * from './types/wps';

/**
 * 主函数 - 程序入口
 */
async function main(): Promise<void> {
  const mainLogger = createChildLogger('Main');

  mainLogger.info('='.repeat(50));
  mainLogger.info('WPS Office MCP Server');
  mainLogger.info('赢容出品，必属精品');
  mainLogger.info('='.repeat(50));

  // 创建服务器实例
  const server = createMcpServer({
    name: MCP_NAME,
    version: MCP_VERSION,
    debug: process.env.DEBUG === 'true',
  });

  // 优雅关闭处理
  const gracefulShutdown = async (signal: string): Promise<void> => {
    mainLogger.info(`Received ${signal}, shutting down gracefully...`);

    try {
      // 先收掉常驻的 PowerShell 守护进程：stdin 一关它会自行退出，
      // 但显式 dispose 能保证不留孤儿进程（以前漏过实例在后台）
      psDaemon.dispose();
      await server.stop();
      mainLogger.info('Server stopped successfully');
      process.exit(0);
    } catch (error) {
      mainLogger.error('Error during shutdown', error);
      process.exit(1);
    }
  };

  // 监听进程信号
  process.on('SIGINT', () => gracefulShutdown('SIGINT'));
  process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));

  // 未捕获异常处理
  process.on('uncaughtException', (error) => {
    mainLogger.error('Uncaught exception', error);
    process.exit(1);
  });

  process.on('unhandledRejection', (reason) => {
    mainLogger.error('Unhandled rejection', reason);
    process.exit(1);
  });

  try {
    // 启动服务器
    await server.start();

    const status = server.getStatus();
    mainLogger.info('Server is running', status);
    mainLogger.info('Waiting for MCP client connection...');

    // 服务器会一直运行，等待MCP客户端连接
    // stdio传输层会保持进程活跃

  } catch (error) {
    mainLogger.error('Failed to start server', error);
    process.exit(1);
  }
}

// 如果是直接运行而不是被导入，则启动服务器
if (require.main === module) {
  const sub = (process.argv[2] || '').trim();

  if (sub === 'install') {
    // 子命令：写客户端配置后退出，不碰 stdio（MCP 握手靠的是"无参数运行"这条路径）
    import('./cli/install')
      .then((m) => m.runInstall(process.argv.slice(3)))
      .then((code) => process.exit(code))
      .catch((error) => {
        log.error('install failed', error);
        process.exit(1);
      });
  } else if (sub === 'help' || sub === '--help' || sub === '-h') {
    // 无参数 = 起 MCP server（客户端就是这么调的），所以帮助必须显式要
    process.stdout.write(
      [
        'wps-multitool-mcp —— 用 COM 操控本机正开着的 WPS 文字/表格/演示',
        '',
        '用法：',
        '  wps-multitool-mcp                 以 MCP server 启动（stdio，给客户端调用）',
        '  wps-multitool-mcp install         把启动项写进已检测到的 MCP 客户端配置',
        '      --dry-run                     只看会改什么，不落盘',
        '      --print                       只打印配置片段，一个文件都不动',
        '      --client <id>                 只处理某个客户端：' +
          'codebuddy / workbuddy / codex / claude / claude-settings / cursor / dsh',
        '  wps-multitool-mcp help            显示这段帮助',
        '',
        '环境变量：WPS_TOOL_TIER=full 关掉默认档过滤；WPS_TOOL_GROUPS=word,excel 指定初始加载的组',
        '',
      ].join('\n')
    );
    process.exit(0);
  } else {
    main().catch((error) => {
      log.error('Fatal error', error);
      process.exit(1);
    });
  }
}

export default main;
