/**
 * Input: 动作名 + 调用点声明的 WpsAppType
 * Output: 桥要用的应用键（excel / word / ppt / common）
 * Pos: 应用路由解析。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 * 路由优先级
 *      覆盖那些没传 appType 的旧调用点（save / openFile / ping / wpsStatus 等）
 */

import { WpsAppType } from '../types/wps';
import { ACTION_ROUTES } from './action-routes.generated';

/** 桥里按应用拆的子进程键 */
export type AppKey = 'excel' | 'word' | 'ppt' | 'common';

/** WpsAppType -> 应用键（WpsAppType 的值就是 WPS 自己的进程名风格：wps / et / wpp） */
export function appKeyFromType(appType?: WpsAppType | AppKey): AppKey | undefined {
  switch (appType) {
    case WpsAppType.SPREADSHEET:
    case 'excel':
      return 'excel';
    case WpsAppType.WRITER:
    case 'word':
      return 'word';
    case WpsAppType.PRESENTATION:
    case 'ppt':
      return 'ppt';
    default:
      return undefined;
  }
}

export function resolveAppKey(action: string, appType?: WpsAppType | AppKey): AppKey {
  const mode = process.env.WPS_BRIDGE_ROUTE_MODE;
  const fromType = appKeyFromType(appType);
  const fromTable = ACTION_ROUTES[action];

  if (mode === 'table') {
    return fromTable ?? fromType ?? 'common';
  }
  return fromType ?? fromTable ?? 'common';
}

/** 供诊断用：这张表里到底认识多少个动作 */
export function routeTableSize(): number {
  return Object.keys(ACTION_ROUTES).length;
}
