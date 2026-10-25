/**
 * Input: package.json
 * Output: 单一真相源的 NAME / VERSION / AUTHOR / CONTRIBUTORS
 * Pos: 版本与署名信息。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有这个文件
 * ------------------
 * 改造前版本号硬编码在三处：package.json、mcp-server.ts 的 DEFAULT_CONFIG、
 * index.ts 的启动日志。三处各写一遍，迟早会漏 —— 而 MCP 握手报给客户端的
 * 是 DEFAULT_CONFIG.version，改了 package.json 而忘了改那里，
 * 用户在客户端里看到的还是旧版本，排查时非常费劲。
 *
 * 现在 package.json 是唯一真相源，这里只读一次。
 * 读不到时退回内置常量（打包成单文件、或 dist 布局变了也不会崩）。
 */

import * as fs from 'fs';
import * as path from 'path';

/** 读不到 package.json 时的兜底值 —— 与 package.json 保持一致 */
const FALLBACK = {
  name: 'wps-multitool-mcp',
  version: '1.0.1',
  author: '赢容',
  contributors: [] as string[],
} as const;

interface PackageInfo {
  name: string;
  version: string;
  author: string;
  contributors: string[];
}

function load(): PackageInfo {
  // dist/version.js → 上一级就是包根
  const candidates = [
    path.join(__dirname, '..', 'package.json'),
    path.join(__dirname, '..', '..', 'package.json'),
  ];
  for (const p of candidates) {
    try {
      const raw = JSON.parse(fs.readFileSync(p, 'utf8')) as Record<string, unknown>;
      if (raw && typeof raw.version === 'string') {
        return {
          name: typeof raw.name === 'string' ? raw.name : FALLBACK.name,
          version: raw.version,
          author: typeof raw.author === 'string' ? raw.author : FALLBACK.author,
          contributors: Array.isArray(raw.contributors)
            ? (raw.contributors as unknown[]).map(String)
            : [],
        };
      }
    } catch {
      /* 试下一个候选路径 */
    }
  }
  return {
    ...FALLBACK,
    contributors: [...FALLBACK.contributors],
  };
}

const info = load();

export const MCP_NAME = info.name;
export const MCP_VERSION = info.version;
export const MCP_AUTHOR = info.author;
export const MCP_CONTRIBUTORS: readonly string[] = info.contributors;

/** 一行版署名，给「关于」类展示用 */
export function versionLine(): string {
  const c = MCP_CONTRIBUTORS.length ? `（${MCP_CONTRIBUTORS.join('、')}）` : '';
  return `${MCP_NAME} v${MCP_VERSION} — ${MCP_AUTHOR}${c} · MIT`;
}
