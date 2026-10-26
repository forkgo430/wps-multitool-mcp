/**
 * Input: 动作名 + 调用点声明的 WpsAppType
 * Output: 桥要用的应用键（excel / word / ppt / common）
 * Pos: 应用路由解析。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 路由优先级
 * ----------
 *   1) 调用点声明的 appType —— 这是最可靠的（231/238 个调用点都传了），直接映射成应用键
 *   2) 生成的路由表 ACTION_ROUTES[action] —— 从 wps-com.ps1 的 case 体推断而来，
 *      覆盖那些没传 appType 的旧调用点（save / openFile / ping / wpsStatus 等）
 *   3) 兜底 'common'
 *
 * 为什么两套并存
 * --------------
 * appType 是"调用者的意图"，路由表是"脚本里的实际依赖"。二者正常情况下一致；
 * 不一致时以 appType 为准（更接近真实意图），路由表负责补盲。
 * 可用 WPS_BRIDGE_ROUTE_MODE=table 强制只用表，便于排查。
 */

import { WpsAppType } from '../types/wps';
import { ACTION_ROUTES } from './action-routes.generated';

/** 桥里按应用拆的子进程键 */
export type AppKey = 'excel' | 'word' | 'ppt' | 'common';

/** WpsAppType -> 应用键（WpsAppType 的值就是 WPS 自己的进程名风格：wps / et / wpp） */
export function appKeyFromType(appType?: WpsAppType): AppKey | undefined {
  switch (appType) {
    case WpsAppType.SPREADSHEET:
      return 'excel';
    case WpsAppType.WRITER:
      return 'word';
    case WpsAppType.PRESENTATION:
      return 'ppt';
    default:
      return undefined;
  }
}

/**
 * 解析一个动作该走哪一路子进程。
 * @param action 动作名，如 'getCellValue'
 * @param appType 调用点声明的应用类型（可空）
 */
export function resolveAppKey(action: string, appType?: WpsAppType): AppKey {
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
