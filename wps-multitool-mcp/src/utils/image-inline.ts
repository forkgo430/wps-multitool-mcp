/**
 * Input: 导出类工具产出的图片文件路径
 * Output: MCP 图像内容块（base64）——文件过大或不可读时返回 null（只给文本路径）
 * Pos: 让 AI "看见"自己导出的图（评审对标 2026-10-26：头部项目把截图作为工具结果返回，
 *      AI 能自查排版/图表效果）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么有大小上限：图像走 base64 进工具结果，直接进模型上下文。
 * 1.5MB 原图 ≈ 2MB base64，超过就把上下文撑爆——超限时返回 null，
 * 工具结果退化为"只给路径"，调用方需要图时可以自己缩小尺寸重导。
 */
import * as fs from 'fs';
import type { ToolContent } from '../types/tools';

const MAX_INLINE_BYTES = 1_500_000;

export async function imageContentBlock(
  filePath: string,
  mimeType = 'image/png'
): Promise<ToolContent | null> {
  try {
    const stat = await fs.promises.stat(filePath);
    if (!stat.isFile() || stat.size === 0 || stat.size > MAX_INLINE_BYTES) return null;
    const data = await fs.promises.readFile(filePath);
    return { type: 'image', data: data.toString('base64'), mimeType };
  } catch {
    return null;
  }
}
