/**
 * Input: slideIndex + shapes[]（批量 upsert 形状：带 shapeIndex/name 即改，不带即建）
 * Output: 逐项结果数组 + 聚合 warnings——一次 ps1 往返完成 N 项（单职责原语，非 batch 多步执行器）
 * Pos: ⑤ 批量形状原语（主线一 A 段 2026-11-06；接口草案 experience-prep §三）。
 * 为什么单开一个 ps1 case 而不走 wps_common_batch：batch 是通用多步执行器、上限 3 步
 * （COM STA 串行、占通道易撞人工操作）；本原语一次调用在 ps1 内循环 N 项 = 一次 COM 往返，
 * 快照/部分失败语义天然干净。上限 50（超过易把通道占满，且 50 个形状已超出单页合理密度）。
 * TS 侧聚合——部分失败 = success:true + 逐项 warning（severity:'warning'）；
 * 全部失败 = success:false（整体失败）。坐标口径=磅（1.0.5 全组裁定）。
 */
import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';
import { reportProgress } from '../../server/progress-context';

/** 单项上限：50——超出单页合理密度，且防 COM 通道被长时间占用 */
export const BATCH_SHAPES_MAX = 50;

interface ShapeUpsert {
  /** 改的定位键：页内形状序号（与 name 二选一；都不带 = 新建） */
  shapeIndex?: number;
  name?: string;
  type?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  text?: string;
  fillColor?: string;
}

export const batchUpsertShapesDefinition: ToolDefinition = {
  name: 'wps_ppt_batch_upsert_shapes',
  description: `批量创建/更新幻灯片形状（upsert：带 shapeIndex/name 且存在即改，否则新建）。一次调用 = 一次 COM 往返完成 N 项（上限 ${BATCH_SHAPES_MAX}）。

shapes[] 每项字段：
- shapeIndex/name: 定位键（改已有形状用；都不带 = 新建）
- type: 形状类型（rectangle/oval/triangle/diamond/pentagon/hexagon/arrow/star/heart/cloud；新建时用，默认 rectangle）
- left/top/width/height: 位置与大小（磅；新建缺省 100/100/100/100，更新时缺省不动）
- text: 形状内文字（可选）
- fillColor: 填充色十六进制（可选）

返回：逐项结果数组（index/action=created|updated/name/ok/error）+ 聚合计数。
语义：遇错不中断不重试——部分失败回 success:true + 逐项 warnings；全部失败回 success:false。
坐标口径：磅（v1.0.5 全组裁定）。

使用场景：
- "这一页摆 4 个 KPI 卡片底块"（一次建 4 个矩形）
- "把这页的 8 个形状统一改位置"（批量更新）`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始），默认1',
      },
      shapes: {
        type: 'array',
        description: `形状 upsert 清单（上限 ${BATCH_SHAPES_MAX}）`,
        items: {
          type: 'object',
          properties: {
            shapeIndex: { type: 'number', description: '要更新的形状序号（从1开始；与 name 二选一，都不带 = 新建）' },
            name: { type: 'string', description: '要更新的形状名称（与 shapeIndex 二选一）' },
            type: { type: 'string', description: '形状类型（新建时）', enum: ['rectangle', 'oval', 'triangle', 'diamond', 'pentagon', 'hexagon', 'arrow', 'star', 'heart', 'cloud'] },
            left: { type: 'number', description: '左边距（磅）' },
            top: { type: 'number', description: '上边距（磅）' },
            width: { type: 'number', description: '宽度（磅）' },
            height: { type: 'number', description: '高度（磅）' },
            text: { type: 'string', description: '形状内文字' },
            fillColor: { type: 'string', description: '填充色十六进制，如 #FF0000' },
          },
        },
      },
    },
    required: ['slideIndex', 'shapes'],
  },
};

export const batchUpsertShapesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapes } = args as {
    slideIndex?: number;
    shapes?: ShapeUpsert[];
  };

  const list = Array.isArray(shapes) ? shapes : [];
  if (!list.length) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: 'shapes 为空——至少给一项（新建给 type+几何，更新给 shapeIndex/name）' }],
      error: 'shapes is empty',
    };
  }
  if (list.length > BATCH_SHAPES_MAX) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `shapes 上限 ${BATCH_SHAPES_MAX}（收到 ${list.length}）——大批量请分批` }],
      error: `shapes exceeds limit ${BATCH_SHAPES_MAX}`,
    };
  }

  const effSlide = slideIndex || 1;
  reportProgress(0, list.length, `批量形状 0/${list.length} 项`);

  const r = await wpsClient.executeMethod<{
    slideIndex: number;
    total: number;
    ok: number;
    failed: number;
    results: Array<{ index: number; action: string; name: string; ok: boolean; error: string }>;
  }>(
    'batchUpsertShapes',
    { slideIndex: effSlide, shapes: list },
    WpsAppType.PRESENTATION
  );
  reportProgress(list.length, list.length, `批量形状 ${list.length}/${list.length} 项完成`);

  if (!r.success) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `批量形状失败: ${r.error}` }],
      error: r.error,
    };
  }
  const d = r.data;
  if (!d) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '批量形状失败: ps1 未回 data' }],
      error: 'ps1 returned no data',
    };
  }

  // 聚合：逐项失败 = warning（severity:'warning'）；全部失败 = 整体 success:false
  const warnings = d.results
    .filter((x) => !x.ok)
    .map((x) => ({
      step: `shapes[${x.index}]`,
      severity: 'warning' as const,
      message: `第 ${x.index} 项${x.action ? `（${x.action}）` : ''}失败：${x.error || '未知错误'}`,
    }));

  const lines = d.results.map((x) => {
    const mark = x.ok ? '✓' : '✗';
    const act = x.action || (x.ok ? 'ok' : 'failed');
    return `${mark} [${x.index}] ${act}${x.name ? ` ${x.name}` : ''}${x.ok ? '' : ` — ${x.error}`}`;
  });
  const head = `批量形状完成：${d.ok}/${d.total} 成功（页 ${d.slideIndex}）`;
  const text = warnings.length
    ? `${head}\n${lines.join('\n')}\n${warnings.length} 项失败已列入 warnings——成功项已生效，失败项需按逐项错误处理后单独补。`
    : `${head}\n${lines.join('\n')}`;

  if (d.ok === 0 && d.failed > 0) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text }],
      error: warnings.map((w) => w.message).join('；'),
      warnings,
    };
  }

  return {
    id: uuidv4(),
    success: true,
    content: [{ type: 'text', text }],
    ...(warnings.length ? { warnings } : {}),
  };
};

export const batchShapesTools: RegisteredTool[] = [
  { definition: batchUpsertShapesDefinition, handler: batchUpsertShapesHandler },
];

export default batchShapesTools;
