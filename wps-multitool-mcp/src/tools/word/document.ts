/**
 * Input: 文档管理工具参数
 * Output: 文档操作结果
 * Pos: Word 文档管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word文档管理Tools
 * 处理文档的打开、切换、获取列表等管理操作
 *
 * 包含：
 * - wps_word_get_open_documents: 获取所有已打开的文档列表
 * - wps_word_switch_document: 切换到指定文档
 * - wps_word_open_document: 打开指定路径的文档
 * - wps_word_get_document_text: 获取文档文本内容
 * - wps_word_insert_header: 设置页眉内容
 * - wps_word_insert_footer: 设置页脚内容
 * - wps_word_generate_doc_toc: 自动生成文档目录
 * - wps_word_insert_section_break: 插入分节符
 * - wps_word_set_line_spacing: 设置行距
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient, setWordTarget } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';
import { WORD_TARGET_PROPERTY } from './target';

/**
 * 获取所有已打开的文档列表
 */
export const getOpenDocumentsDefinition: ToolDefinition = {
  name: 'wps_word_get_open_documents',
  description: `获取当前WPS Writer中所有已打开的文档列表。

使用场景：
- "看看现在打开了哪些文档"
- "列出所有打开的Word文件"
- "查看当前文档列表"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: { target: WORD_TARGET_PROPERTY, },
  },
};

export const getOpenDocumentsHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      documents: Array<{ name: string; path: string; active: boolean }>;
    }>(
      'getOpenDocuments',
      {},
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const docs = response.data.documents;
      if (!docs || docs.length === 0) {
        return {
          id: uuidv4(),
          success: true,
          content: [{ type: 'text', text: '当前没有打开任何文档' }],
          // 第 1 期第 2 章第 4 轮（A9）：把 PS 侧的 warnings 传出去。
          // 以前这个 handler 只看 data，warnings 被整个丢掉 —— 于是 PS 特意提示的
          // "这可能是 ROT 重注册窗口，别据此判断文档没了"根本到不了模型眼前。
          // warnings 与 success 正交：success 仍为 true，但调用方必须念给用户听。
          warnings: response.warnings,
        };
      }

      const docList = docs
        .map((doc, i) => `${i + 1}. ${doc.name}${doc.active ? ' (当前活动)' : ''}\n   路径: ${doc.path}`)
        .join('\n');

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已打开的文档列表 (共${docs.length}个):\n${docList}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取文档列表失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取文档列表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 切换到指定文档
 */
export const switchDocumentDefinition: ToolDefinition = {
  name: 'wps_word_switch_document',
  description: `切换到指定名称的文档。

使用场景：
- "切换到报告.docx"
- "打开另一个文档窗口"
- "切换到那个合同文档"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      name: {
        type: 'string',
        description: '要切换到的文档名称',
      },
    },
    required: ['name'],
  },
};

export const switchDocumentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name } = args as { name: string };

  if (!name || name.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '文档名称不能为空！' }],
      error: '文档名称为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'switchDocument',
      { name },
      WpsAppType.WRITER
    );

    if (response.success) {
      // 锁定目标：之后所有 word 调用都会自动带上 documentName，
      // 避免 ActiveDocument 在多文档场景下漂移/返回空壳。
      setWordTarget(name);
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已切换到文档: ${name}（后续 word 操作将锁定此文档）`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `切换文档失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `切换文档出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 打开指定路径的文档
 */
export const openDocumentDefinition: ToolDefinition = {
  name: 'wps_word_open_document',
  description: `打开指定路径的Word文档。

使用场景：
- "打开桌面上的报告.docx"
- "帮我打开这个文件路径的文档"
- "加载指定位置的Word文件"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      filePath: {
        type: 'string',
        description: '要打开的文档文件路径',
      },
    },
    required: ['filePath'],
  },
};

export const openDocumentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { filePath } = args as { filePath: string };

  if (!filePath || filePath.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '文件路径不能为空！' }],
      error: '文件路径为空',
    };
  }

  try {
    const params = {
      filePath,
      path: filePath,
    };

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      documentName: string;
    }>(
      'openDocument',
      params,
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文档打开成功！\n文件: ${response.data.documentName || filePath}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `打开文档失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `打开文档出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 获取文档文本内容
 */
export const getDocumentTextDefinition: ToolDefinition = {
  name: 'wps_word_get_document_text',
  description: `获取当前Word文档的文本内容。

使用场景：
- "读取文档内容"
- "获取文档的前100个字符"
- "查看文档从第50到第200个字符的内容"

可指定起始和结束位置来获取部分文本。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      start: {
        type: 'number',
        description: '起始位置（字符索引），默认从头开始',
      },
      end: {
        type: 'number',
        description: '结束位置（字符索引），默认到文档末尾',
      },
    },
  },
};

export const getDocumentTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { start, end } = args as { start?: number; end?: number };

  try {
    const params: Record<string, unknown> = {};
    if (start !== undefined) params.start = start;
    if (end !== undefined) params.end = end;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      text: string;
      length: number;
    }>(
      'getDocumentText',
      params,
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const rangeInfo = start !== undefined || end !== undefined
        ? `\n范围: ${start ?? 0} - ${end ?? '末尾'}`
        : '';

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文档文本内容 (${response.data.length}字符)${rangeInfo}:\n\n${response.data.text}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取文档文本失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取文档文本出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置页眉内容
 */
export const insertHeaderDefinition: ToolDefinition = {
  name: 'wps_word_insert_header',
  description: `设置页眉内容。

使用场景：
- "给文档加个页眉"
- "设置页眉为公司名称"
- "修改第2节的页眉"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      text: {
        type: 'string',
        description: '页眉文本内容',
      },
      section: {
        type: 'number',
        description: '节编号（从1开始），默认1',
        default: 1,
      },
    },
    required: ['text'],
  },
};

export const insertHeaderHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text, section = 1 } = args as { text: string; section?: number };

  if (!text) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '页眉文本不能为空！' }],
      error: '页眉文本为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertHeader',
      { text, section },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `页眉已设置: "${text}" (第${section}节)` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置页眉失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置页眉出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置页脚内容
 */
export const insertFooterDefinition: ToolDefinition = {
  name: 'wps_word_insert_footer',
  description: `设置页脚内容。

使用场景：
- "给文档加个页脚"
- "设置页脚为页码"
- "修改第1节的页脚"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      text: {
        type: 'string',
        description: '页脚文本内容',
      },
      section: {
        type: 'number',
        description: '节编号（从1开始），默认1',
        default: 1,
      },
    },
    required: ['text'],
  },
};

export const insertFooterHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text, section = 1 } = args as { text: string; section?: number };

  if (!text) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '页脚文本不能为空！' }],
      error: '页脚文本为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertFooter',
      { text, section },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `页脚已设置: "${text}" (第${section}节)` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置页脚失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置页脚出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 插入分节符
 */
export const insertSectionBreakDefinition: ToolDefinition = {
  name: 'wps_word_insert_section_break',
  description: `在光标处插入分节符，把文档分成不同的节（各节可以有不同的页边距/纸张/页眉页脚）。

与 wps_word_insert_page_break 的关系：**同一个底层动作（insertPageBreak）**，本工具是它的【分节子集】——
只插入"节"这一种分隔，但多一个 breakType 决定节的起法。只想要换页用 wps_word_insert_page_break。

四种分节符（对应 Word 的 wdSectionBreak*）：
- nextPage：下一页起新节（最常用）
- continuous：同页连续分节（同一页里换栏型/页边距时用）
- evenPage / oddPage：从下一个偶数页/奇数页起新节（双面排版用）

使用场景：
- "插入一个分节符"
- "从下一页开始新的一节"
- "让这部分从奇数页开始"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      breakType: {
        type: 'string',
        description: '分节符类型：nextPage(下一页)、continuous(连续)、evenPage(偶数页)、oddPage(奇数页)，默认nextPage',
        default: 'nextPage',
        enum: ['nextPage', 'continuous', 'evenPage', 'oddPage'],
      },
    },
  },
};

export const insertSectionBreakHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { breakType = 'nextPage' } = args as { breakType?: string };

  // breakType → ps1 的 insertPageBreak 的 type 值。
  // 注意 nextPage 在这里叫 'section'（wdSectionBreakNextPage=2），不是 'page' ——
  // 名字对不上很容易被"顺手"写成 page，那就会变成普通分页符：用户以为分了节，
  // 后面按节设页边距时会发现所有内容还在同一节里。
  const toPsType: Record<string, string> = {
    nextPage: 'section',
    continuous: 'sectionContinuous',
    evenPage: 'evenPage',
    oddPage: 'oddPage',
  };
  const psType = toPsType[breakType];
  if (!psType) {
    return {
      id: uuidv4(),
      success: false,
      content: [
        {
          type: 'text',
          text: `[勿重试·先改参数] 不认识的分节符类型「${breakType}」。可用：nextPage / continuous / evenPage / oddPage`,
        },
      ],
      error: 'breakType 非法',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      type?: string;
      code?: number;
    }>(
      // 第 1 期第 2 章第 2 轮修：原先调 'insertSectionBreak'，ps1 里【没有这个动作】（零命中）
      // → 每次 Unknown action，工具从未生效。insertPageBreak 本来就支持分节（type=section/
      // sectionContinuous），本轮又给它补了 evenPage/oddPage 两个映射值，所以四种都能落。
      // 原来那句 "// NOTE: 仅Windows支持" 是假的：Windows 上也不支持，压根没这个 case。
      'insertPageBreak',
      { type: psType },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `分节符已插入（类型: ${breakType}，WPS 断点码: ${response.data?.code ?? psType}）`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入分节符失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入分节符出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置行距
 */
export const setLineSpacingDefinition: ToolDefinition = {
  name: 'wps_word_set_line_spacing',
  description: `设置段落行距。

使用场景：
- "把行距设为1.5倍"
- "调整行距"

与 wps_word_set_paragraph 的关系：**同一个底层动作（setParagraph）**，本工具是它的【行距子集】——
只暴露 lineSpacing 这一个属性。要同时改对齐方式、段前段后间距、缩进，用 wps_word_set_paragraph。

作用范围：默认改【当前选中的段落】；range="all" 改全文。
要改"第 3 段"，先把光标放到那一段（选中它），再调用本工具 —— 不支持按段落序号定位。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      lineSpacing: {
        type: 'number',
        description: '行距值（如1.0、1.5、2.0等）',
      },
      range: {
        type: 'string',
        description: '应用范围：selection（当前选中的段落，默认）/ all（全文）',
        enum: ['selection', 'all'],
      },
    },
    required: ['lineSpacing'],
  },
};

export const setLineSpacingHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { lineSpacing, range } = args as { lineSpacing: number; range?: string };

  if (lineSpacing === undefined || lineSpacing <= 0) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '行距值必须为正数！' }],
      error: '行距值无效',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      // 第 1 期第 2 章第 1 轮修：原先调 'setLineSpacing'，而 wps-com.ps1 里【没有这个动作】
      // （零命中）→ 每次调用都落到分发器的 default 分支报 Unknown action，工具等于从未生效。
      // 能力本来就在 setParagraph 上（wps-com.ps1:3907），且参数名 lineSpacing 两边一致，
      // 所以只改动作名，PS 侧一行都不动。
      'setParagraph',
      { lineSpacing, range: range || 'selection' },
      WpsAppType.WRITER
    );

    if (response.success) {
      const target = range === 'all' ? '全文' : '当前选中的段落';
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `行距已设置为 ${lineSpacing} 倍（${target}）` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置行距失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置行距出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有文档管理相关的Tools
 */
export const documentTools: RegisteredTool[] = [
  { definition: getOpenDocumentsDefinition, handler: getOpenDocumentsHandler },
  { definition: switchDocumentDefinition, handler: switchDocumentHandler },
  { definition: openDocumentDefinition, handler: openDocumentHandler },
  { definition: getDocumentTextDefinition, handler: getDocumentTextHandler },
  { definition: insertHeaderDefinition, handler: insertHeaderHandler },
  { definition: insertFooterDefinition, handler: insertFooterHandler },
  { definition: insertSectionBreakDefinition, handler: insertSectionBreakHandler },
  { definition: setLineSpacingDefinition, handler: setLineSpacingHandler },
];

export default documentTools;
