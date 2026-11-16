/**
 * Input: 工作表生命周期管理工具参数
 * Output: 工作表生命周期操作结果
 * Pos: Excel 工作表生命周期管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 工作表的创建、删除、重命名、复制、移动、切换等操作（高级操作见 sheet-advanced.ts）
 * - wps_excel_delete_sheet: 删除指定工作表
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

export const createSheetDefinition: ToolDefinition = {
  name: 'wps_excel_create_sheet',
  description: '在当前工作簿中创建新的工作表。可指定名称和插入位置。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '新工作表的名称',
      },
      position: {
        type: 'number',
        description: '插入位置索引（从0开始），不填则添加到末尾',
      },
    },
    required: ['name'],
  },
};

export const createSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, position } = args as {
    name: string;
    position?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      name: string;
      index: number;
    }>(
      'createSheet',
      { name, position },
      WpsAppType.SPREADSHEET
    );

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `工作表创建成功！\n名称: ${response.data.name}\n位置: 第${response.data.index + 1}个`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 删除指定工作表
 */
export const deleteSheetDefinition: ToolDefinition = {
  name: 'wps_excel_delete_sheet',
  description: '删除当前工作簿中的指定工作表。注意：此操作不可撤销。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要删除的工作表名称',
      },
    },
    required: ['name'],
  },
};

export const deleteSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name } = args as { name: string };

  try {
    const response = await wpsClient.executeMethod<{
      deleted: string;
    }>(
      'deleteSheet',
      { name },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `删除工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `工作表 "${name}" 已成功删除`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `删除工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

export const renameSheetDefinition: ToolDefinition = {
  name: 'wps_excel_rename_sheet',
  description: '重命名当前工作簿中的指定工作表。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      oldName: {
        type: 'string',
        description: '当前工作表名称',
      },
      newName: {
        type: 'string',
        description: '新的工作表名称',
      },
    },
    required: ['oldName', 'newName'],
  },
};

export const renameSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { oldName, newName } = args as {
    oldName: string;
    newName: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      oldName: string;
      newName: string;
    }>(
      'renameSheet',
      { oldName, newName },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `重命名工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `工作表重命名成功！\n"${oldName}" → "${newName}"`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `重命名工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

export const copySheetDefinition: ToolDefinition = {
  name: 'wps_excel_copy_sheet',
  description: '复制当前工作簿中的指定工作表。可指定新名称和插入位置。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要复制的工作表名称',
      },
      newName: {
        type: 'string',
        description: '复制后的工作表名称，不填则自动生成',
      },
      position: {
        type: 'number',
        description: '插入位置索引（从0开始），不填则添加到末尾',
      },
    },
    required: ['name'],
  },
};

export const copySheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, newName, position } = args as {
    name: string;
    newName?: string;
    position?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      sourceName: string;
      newName: string;
      index: number;
    }>(
      'copySheet',
      { name, newName, position },
      WpsAppType.SPREADSHEET
    );

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `复制工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `工作表复制成功！\n源工作表: ${response.data.sourceName}\n新工作表: ${response.data.newName}\n位置: 第${response.data.index + 1}个`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `复制工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

export const getSheetListDefinition: ToolDefinition = {
  name: 'wps_excel_get_sheet_list',
  description: '获取当前工作簿的所有工作表列表，包含名称、索引和是否为活动工作表。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const getSheetListHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      sheets: Array<{
        name: string;
        index: number;
        active: boolean;
      }>;
      count: number;
    }>(
      'getSheetList',
      {},
      WpsAppType.SPREADSHEET
    );

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取工作表列表失败: ${response.error}` }],
        error: response.error,
      };
    }

    const { sheets, count } = response.data;
    let output = `当前工作簿共有 ${count} 个工作表：\n\n`;

    sheets.forEach((sheet) => {
      const activeFlag = sheet.active ? ' [活动]' : '';
      output += `${sheet.index + 1}. ${sheet.name}${activeFlag}\n`;
    });

    return {
      id: uuidv4(),
      success: true,
      content: [{ type: 'text', text: output }],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取工作表列表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

export const switchSheetDefinition: ToolDefinition = {
  name: 'wps_excel_switch_sheet',
  description: '切换到指定的工作表，使其成为活动工作表。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要切换到的工作表名称',
      },
    },
    required: ['name'],
  },
};

export const switchSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name } = args as { name: string };

  try {
    const response = await wpsClient.executeMethod<{
      activatedSheet: string;
    }>(
      'switchSheet',
      { name },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `切换工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `已切换到工作表 "${name}"`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `切换工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

export const moveSheetDefinition: ToolDefinition = {
  name: 'wps_excel_move_sheet',
  description: '移动指定工作表到新的位置。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要移动的工作表名称',
      },
      position: {
        type: 'number',
        description: '目标位置索引（从0开始）',
      },
    },
    required: ['name', 'position'],
  },
};

export const moveSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, position } = args as {
    name: string;
    position: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      name: string;
      newPosition: number;
    }>(
      'moveSheet',
      { name, position },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `移动工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `工作表 "${name}" 已移动到第${position + 1}个位置`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `移动工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

export const sheetTools: RegisteredTool[] = [
  { definition: createSheetDefinition, handler: createSheetHandler },
  { definition: deleteSheetDefinition, handler: deleteSheetHandler },
  { definition: renameSheetDefinition, handler: renameSheetHandler },
  { definition: copySheetDefinition, handler: copySheetHandler },
  { definition: getSheetListDefinition, handler: getSheetListHandler },
  { definition: switchSheetDefinition, handler: switchSheetHandler },
  { definition: moveSheetDefinition, handler: moveSheetHandler },
];

export default sheetTools;
