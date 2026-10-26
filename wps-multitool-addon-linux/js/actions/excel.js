/**
 * Input: 动作参数（et 型，Excel；与 src/tools 的 TS handler 语义对齐，B5-T2 接线批再做参数名逐一核对）
 * Output: Promise<{ ok: true, data } | { ok: false, error | unsupported }>（通道协议回包结构）
 * Pos: Linux 加载项通道的动作层 · et 型（表格）第一批 57 实现 + 3 PivotTables 桩。
 *      一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 范围（B5-T1 任务卡 §0）：S1-REPORT §3 Excel 段判定"支持"的 60 条。
 *   - 57 条真实现（jsAsynCall 异步形态，B4 模型原样复用）；
 *   - createPivotTable / getPivotTables / refreshPivotTables 3 条虽判"支持"，但证据是
 *     "PivotTables 存在、Count=undefined、Item 不可调"（S1 §3:35/74/75 + §4 #6，报告行 335）——
 *     按任务卡 §0 标 unsupported:'待S1补测'，绝不猜 method；
 *   - 部分支持 35（集合双语义主因）/抛错 4/不支持 3 本批不做（任务卡：单独一批防审不动）。
 *
 * ExcelAdapter（集合双语义，S1 §4 逐条落码，报告行 330-335）
 * ------------------------------------------------------
 *   - 集合不可当函数：`ws.Cells(r,c)`/`ws.Rows(i)`/`ws.Columns(i)` 实测 TypeError →
 *     任何单元格定位必须走 Item 语义（cellRef 统一出口）；
 *   - `.Value = x` 赋值静默失效 → 写值一律 Value2（writeValue2 强制，payload 里出现 value 键直接抛）；
 *   - 读二维走 .Value2（range.Value 是函数不是数组，§4 #4）；
 *   - `ws.Range('A1:B2')` 是函数可直接调（§4 #2）——ref 参数原样下发。
 *
 * method 名规范沿用 B4：S1 §3"对应的 JSAPI 操作"列属性链去宿主前缀、定位参数抽进 params，
 * 每条头注标 S1 报告行号。S1 只验证了通路、读回未逐项验证的（§10.3），注释标
 * "通路已证、读回未逐项验证"。method 字符串形式本身无直接样例（B4 §4 缺口声明沿用），B3.5 探活后统一修正。
 * 本文件只在 et 型加载项里加载（addon-et/index.html 引用）；wps/wpp 型没有 Workbooks（S1 §6）。
 */
(function (global) {
  'use strict';

  var jsapi = (typeof module !== 'undefined' && module.exports)
    ? require('../jsapi.js')
    : global.WpsAddonJsapi;

  /** JSAPI 异常统一兜底：原文带回，不改写、不吞（jsAsynCall 的 rejection 也要接住） */
  function guard(impl) {
    return async function (params) {
      try {
        return await impl(params || {});
      } catch (e) {
        return { ok: false, error: String(e && e.message ? e.message : e) };
      }
    };
  }

  // ==========================================================================
  // ExcelAdapter —— 集合双语义统一层（本批的验收核心之一）
  // ==========================================================================

  /** 列号 → A1 列字母（1→A, 26→Z, 27→AA；ps1 同源换算） */
  function colToLetters(col) {
    var n = Number(col);
    var s = '';
    while (n > 0) {
      var rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  /** 行列号 → 'A1' 记法（cellRef 的编码侧：Item 语义的定位参数） */
  function a1(row, col) {
    return colToLetters(col) + Number(row);
  }

  var ExcelAdapter = {
    a1: a1,
    colToLetters: colToLetters,

    /**
     * cellRef：单元格定位的统一出口（S1 §4 #1，报告行 330）。
     * 返回 Item 语义参数 {row, col}；method 侧由动作拼 `Cells.Item`。
     * 【禁止】VBA 直调式——本函数存在就是为了让动作层永远不手写 `Cells(r,c)`。
     */
    cellRef: function (row, col) {
      return { row: Number(row), col: Number(col) };
    },

    /**
     * 从 params 提取 range 引用：range/address 显式优先，row+col 组 A1（Item 语义）。
     * 没有任何定位 → null（动作层据此显式拒绝）。
     */
    refOf: function (p) {
      if (p.range) return String(p.range);
      if (p.address) return String(p.address);
      if (p.row !== undefined && p.col !== undefined) return a1(p.row, p.col);
      return null;
    },

    /**
     * 写值 payload 强制 Value2 语义（S1 §4 #3，报告行 332：.Value = x 静默失效）。
     * 传入对象含 `value` 键 → 直接抛（调用方应写 `value2`）——防呆在编码层，不靠自觉。
     */
    writeValue2: function (payload) {
      if (payload && Object.prototype.hasOwnProperty.call(payload, 'value')) {
        throw new Error('[ExcelAdapter] 禁止 value 键：S1 §4 #3 实测 .Value = x 静默失效，写值一律 value2');
      }
      return payload;
    },

    /** 调用方给 value 的便利包装：自动落 value2 键 */
    withValue2: function (payload, x) {
      var out = {};
      for (var k in payload) {
        if (Object.prototype.hasOwnProperty.call(payload, k) && k !== 'value' && k !== 'value2') out[k] = payload[k];
      }
      out.value2 = x;
      return out;
    },

    /**
     * method 常量：Item/Value2 语义的规范化出口（防各动作手写漂移）。
     * 单元格级读写 = 'Cells.Item.Value2'（Item 定位 + Value2 读写，§4 #1/#3 的合成语义）。
     */
    M: {
      CELLS_ITEM_VALUE2: 'Cells.Item.Value2',
      RANGE_VALUE2: 'Range.Value2',
    },
  };

  /** 必须有 range 引用，没有就抛（guard 转 ok:false） */
  function mustRef(p, action) {
    var ref = ExcelAdapter.refOf(p);
    if (!ref) throw new Error('[勿重试·先改参数] 缺少 range/address/row+col 定位（' + action + ' 必填）');
    return ref;
  }

  // ==========================================================================
  // 60 动作（57 实现 + 3 PivotTables 桩），按 S1 §3 编号排列
  // ==========================================================================

  var excelActions = {

    /**
     * getActiveWorkbook —— S1-REPORT.md:28（§3 #2 支持：app.ActiveWorkbook.Name）
     * method 'ActiveWorkbook'（属性读，待补测——B4 §4 缺口沿用）。
     */
    getActiveWorkbook: guard(function () {
      return jsapi.jsAsynCall('ActiveWorkbook', {}).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getRangeData —— S1-REPORT.md:31（§3 #5 支持；§4 #4 报告行 333：读二维走 .Value2）
     * method 'Range.Value2'（读）。返回原样透传（二维结构 S1 §10 未逐项验证，以真机为准）。
     */
    getRangeData: guard(function (p) {
      var ref = mustRef(p, 'getRangeData');
      return jsapi.jsAsynCall(ExcelAdapter.M.RANGE_VALUE2, { ref: ref }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * setRangeData —— S1-REPORT.md:32（§3 #6 支持：Range 写二维数组）
     * method 'Range.Value2'（写）。value2 强制（§4 #3）；ref 缺省 A1 起写。
     */
    setRangeData: guard(function (p) {
      var ref = p.ref || p.range || 'A1';
      if (p.value2 === undefined && p.value === undefined) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 value2（setRangeData 要写的二维数组）' });
      }
      if (p.value2 === undefined) {
        // 兼容调用方传 value：经适配器强制改道 value2（不是静默接受——语义换道在编码层可见）
        p = ExcelAdapter.withValue2({ ref: ref }, p.value);
      }
      return jsapi.jsAsynCall(ExcelAdapter.M.RANGE_VALUE2,
        ExcelAdapter.writeValue2({ ref: ref, value2: p.value2 })).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * clearFormats —— S1-REPORT.md:42（§3 #16 支持：range.ClearFormats()）
     * method 'Range.ClearFormats'。通路已证、读回未逐项验证（S1 §10.3）。
     */
    clearFormats: guard(function (p) {
      var ref = mustRef(p, 'clearFormats');
      return jsapi.jsAsynCall('Range.ClearFormats', { ref: ref }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * addConditionalFormat —— S1-REPORT.md:43（§3 #17 支持：FormatConditions.Add(1,1,'>0')）
     * method 'FormatConditions.Add'。S1 证据的三参按位下发（type/operator/formula）。
     */
    addConditionalFormat: guard(function (p) {
      var ref = mustRef(p, 'addConditionalFormat');
      var payload = ExcelAdapter.writeValue2({
        ref: ref,
        type: p.type !== undefined ? Number(p.type) : 1,
        operator: p.operator !== undefined ? Number(p.operator) : 1,
      });
      if (p.formula !== undefined) payload.formula = String(p.formula);
      else if (p.value2 !== undefined) payload.formula = String(p.value2);
      return jsapi.jsAsynCall('FormatConditions.Add', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * removeConditionalFormat —— S1-REPORT.md:44（§3 #18 支持：FormatConditions.Delete()）
     * method 'FormatConditions.Delete'。
     */
    removeConditionalFormat: guard(function (p) {
      var ref = mustRef(p, 'removeConditionalFormat');
      return jsapi.jsAsynCall('FormatConditions.Delete', { ref: ref }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getConditionalFormats —— S1-REPORT.md:45（§3 #19 支持：FormatConditions.Count）
     * method 'FormatConditions.Count'（读）。
     */
    getConditionalFormats: guard(function (p) {
      var ref = mustRef(p, 'getConditionalFormats');
      return jsapi.jsAsynCall('FormatConditions.Count', { ref: ref }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * addDataValidation —— S1-REPORT.md:46（§3 #20 支持：Validation.Add(1,1,1,'10')）
     * method 'Validation.Add'。S1 证据四参按位下发。
     */
    addDataValidation: guard(function (p) {
      var ref = mustRef(p, 'addDataValidation');
      var payload = {
        ref: ref,
        type: p.type !== undefined ? Number(p.type) : 1,
        operator: p.operator !== undefined ? Number(p.operator) : 1,
        alertStyle: p.alertStyle !== undefined ? Number(p.alertStyle) : 1,
      };
      if (p.formula !== undefined) payload.formula = String(p.formula);
      return jsapi.jsAsynCall('Validation.Add', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * removeDataValidation —— S1-REPORT.md:47（§3 #21 支持：Validation.Delete()）
     * method 'Validation.Delete'。
     */
    removeDataValidation: guard(function (p) {
      var ref = mustRef(p, 'removeDataValidation');
      return jsapi.jsAsynCall('Validation.Delete', { ref: ref }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getDataValidations —— S1-REPORT.md:48（§3 #22 支持：Validation.Count/Type）
     * method 'Validation'（读）。
     */
    getDataValidations: guard(function (p) {
      var ref = mustRef(p, 'getDataValidations');
      return jsapi.jsAsynCall('Validation', { ref: ref }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * setArrayFormula —— S1-REPORT.md:50（§3 #24 支持：range.FormulaArray）
     * method 'Range.FormulaArray'（写）。公式串必填。
     */
    setArrayFormula: guard(function (p) {
      var ref = mustRef(p, 'setArrayFormula');
      var formula = p.formula2 !== undefined ? p.formula2 : p.formula;
      if (formula === undefined || formula === null || formula === '') {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 formula（setArrayFormula 的数组公式串）' });
      }
      return jsapi.jsAsynCall('Range.FormulaArray',
        ExcelAdapter.writeValue2({ ref: ref, formula: String(formula) })).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * refreshLinks —— S1-REPORT.md:52（§3 #26 支持：wb.UpdateLink）
     * method 'Workbook.UpdateLink'。通路已证、读回未逐项验证（S1 §10.3）。
     */
    refreshLinks: guard(function () {
      return jsapi.jsAsynCall('Workbook.UpdateLink', {}).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * consolidate —— S1-REPORT.md:53（§3 #27 支持：range.Consolidate）
     * method 'Range.Consolidate'。通路已证、读回未逐项验证（S1 §10.3）。
     */
    consolidate: guard(function (p) {
      var ref = mustRef(p, 'consolidate');
      return jsapi.jsAsynCall('Range.Consolidate', { ref: ref }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * calculateSheet —— S1-REPORT.md:54（§3 #28 支持：ws.Calculate()）
     * method 'Worksheet.Calculate'。
     */
    calculateSheet: guard(function (p) {
      return jsapi.jsAsynCall('Worksheet.Calculate', { sheet: p.sheet }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getExcelContext —— S1-REPORT.md:55（§3 #29 支持：wb/ws/维度汇总）
     * method 'Context'（读，params.includeSelection=false）。
     */
    getExcelContext: guard(function () {
      return jsapi.jsAsynCall('Context', { includeSelection: false }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getContext —— S1-REPORT.md:56（§3 #30 支持：wb/ws/sel 汇总；§4 #5：sel.Address 是函数）
     * method 'Context'（读，params.includeSelection=true）。
     */
    getContext: guard(function () {
      return jsapi.jsAsynCall('Context', { includeSelection: true }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * sortRange —— S1-REPORT.md:57（§3 #31 支持：range.Sort(key, order)）
     * method 'Range.Sort'。key/order 必填（S1 证据两参）。
     */
    sortRange: guard(function (p) {
      var ref = mustRef(p, 'sortRange');
      if (p.key === undefined || p.key === null) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 key（sortRange 的排序键）' });
      }
      return jsapi.jsAsynCall('Range.Sort', {
        ref: ref,
        key: String(p.key),
        order: p.order !== undefined ? Number(p.order) : 1,
      }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * autoFilter —— S1-REPORT.md:58（§3 #32 支持：range.AutoFilter / ws.AutoFilter）
     * method 'Range.AutoFilter'。
     */
    autoFilter: guard(function (p) {
      var ref = mustRef(p, 'autoFilter');
      var payload = { ref: ref };
      if (p.field !== undefined) payload.field = Number(p.field);
      if (p.criteria !== undefined) payload.criteria = String(p.criteria);
      return jsapi.jsAsynCall('Range.AutoFilter', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * updateChart —— S1-REPORT.md:60（§3 #34 判定"支持"，证据 string:无图表）
     * method 'Chart'（ChartType 读写）。⚠ S1 只证"通道可调"，图表对象语义待真机（无图表时探测未触达）。
     */
    updateChart: guard(function (p) {
      if (p.chartIndex === undefined && p.chartName === undefined) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 chartIndex/chartName（updateChart）' });
      }
      var payload = { chartIndex: p.chartIndex !== undefined ? Number(p.chartIndex) : undefined };
      if (p.chartName) payload.chartName = String(p.chartName);
      if (p.chartType !== undefined) payload.chartType = Number(p.chartType);
      return jsapi.jsAsynCall('Chart', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * exportChartAsImage —— S1-REPORT.md:61（§3 #35 判定"支持"，证据 string:无图表）
     * method 'Chart.Export'。⚠ 同 updateChart：对象语义待真机。
     */
    exportChartAsImage: guard(function (p) {
      var path = p.outputPath || p.path;
      if (!path) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 outputPath/path（exportChartAsImage 必填）' });
      }
      var payload = { path: String(path), filter: String(p.format || 'PNG') };
      if (p.chartIndex !== undefined) payload.chartIndex = Number(p.chartIndex);
      if (p.chartName) payload.chartName = String(p.chartName);
      return jsapi.jsAsynCall('Chart.Export', payload).then(function (r) {
        return { ok: true, data: { path: String(path), format: String(p.format || 'PNG') } };
      });
    }),

    /**
     * removeDuplicates —— S1-REPORT.md:63（§3 #37 支持：range.RemoveDuplicates(columns, header)）
     * method 'Range.RemoveDuplicates'。columns/header 按 S1 证据两参下发。
     */
    removeDuplicates: guard(function (p) {
      var ref = mustRef(p, 'removeDuplicates');
      var payload = { ref: ref };
      if (p.columns !== undefined) payload.columns = Number(p.columns);
      if (p.header !== undefined) payload.header = p.header;
      return jsapi.jsAsynCall('Range.RemoveDuplicates', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * deleteSheet —— S1-REPORT.md:65（§3 #39 支持：ws.Delete()）
     * method 'Worksheet.Delete'。定位 name/index 二选一。
     */
    deleteSheet: guard(function (p) {
      if (!p.name && p.index === undefined) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 name/index（deleteSheet 要删哪张表）' });
      }
      var payload = {};
      if (p.name) payload.name = String(p.name);
      if (p.index !== undefined) payload.index = Number(p.index);
      return jsapi.jsAsynCall('Worksheet.Delete', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * renameSheet —— S1-REPORT.md:66（§3 #40 支持：ws.Name = '探针表'）
     * method 'Worksheet.Name'（属性写编码）。
     */
    renameSheet: guard(function (p) {
      if (!p.name) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 name（renameSheet 的新表名）' });
      }
      var payload = { name: String(p.name) };
      if (p.target) payload.target = String(p.target);
      return jsapi.jsAsynCall('Worksheet.Name', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * copySheet —— S1-REPORT.md:67（§3 #41 支持：ws.Copy()）
     * method 'Worksheet.Copy'。
     */
    copySheet: guard(function (p) {
      var payload = {};
      if (p.target) payload.target = String(p.target);
      if (p.before !== undefined) payload.before = Number(p.before);
      if (p.after !== undefined) payload.after = Number(p.after);
      return jsapi.jsAsynCall('Worksheet.Copy', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getSheetList —— S1-REPORT.md:68（§3 #42 支持：Worksheets.Count + 名称）
     * method 'Worksheets'（读）。
     */
    getSheetList: guard(function () {
      return jsapi.jsAsynCall('Worksheets', {}).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getSheetInfo —— S1-REPORT.md:69（§3 #43 支持：ws.Name/UsedRange/Visible）
     * method 'Worksheet.Info'。
     */
    getSheetInfo: guard(function (p) {
      var payload = {};
      if (p.sheet) payload.sheet = String(p.sheet);
      return jsapi.jsAsynCall('Worksheet.Info', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * autoFit —— S1-REPORT.md:70（§3 #44 支持：ws.UsedRange.AutoFit）
     * method 'UsedRange.AutoFit'。与 autoFitAll（#61）同 method——S1 两条证据同源。
     */
    autoFit: guard(function (p) {
      return jsapi.jsAsynCall('UsedRange.AutoFit', { sheet: p.sheet }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * findInSheet —— S1-REPORT.md:72（§3 #46 支持：ws.Cells.Find(what)）
     * method 'Cells.Find'。S1 证据是 Cells 集合的 Find 方法（不是 Cells(r,c) 函数调用，
     * 与 §4 #1 不冲突）。what 必填。
     */
    findInSheet: guard(function (p) {
      if (p.what === undefined || p.what === null || p.what === '') {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 what（findInSheet 的查找串）' });
      }
      var payload = { what: String(p.what) };
      if (p.sheet) payload.sheet = String(p.sheet);
      if (p.after !== undefined) payload.after = String(p.after);
      return jsapi.jsAsynCall('Cells.Find', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * setSheetPageSetup —— S1-REPORT.md:77（§3 #51 支持：PageSetup.Orientation 读回）
     * method 'PageSetup'（多键写编码，待补测）。
     */
    setSheetPageSetup: guard(function (p) {
      if (p.orientation === undefined && p.paperSize === undefined && p.fitToPage === undefined) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] orientation/paperSize/fitToPage 至少给一个（setSheetPageSetup）' });
      }
      var payload = {};
      if (p.orientation !== undefined) payload.orientation = Number(p.orientation);
      if (p.paperSize !== undefined) payload.paperSize = Number(p.paperSize);
      if (p.fitToPage !== undefined) payload.fitToPage = !!p.fitToPage;
      return jsapi.jsAsynCall('PageSetup', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * setSheetPrintTitles —— S1-REPORT.md:78（§3 #52 支持：PageSetup.PrintTitleRows 读回）
     * method 'PageSetup.PrintTitleRows'。rows 串必填（S1 证据 '$1:$1'）。
     */
    setSheetPrintTitles: guard(function (p) {
      if (!p.rows) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 rows（setSheetPrintTitles，如 $1:$1）' });
      }
      var payload = { rows: String(p.rows) };
      if (p.columns) payload.columns = String(p.columns);
      return jsapi.jsAsynCall('PageSetup.PrintTitleRows', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * switchSheet —— S1-REPORT.md:79（§3 #53 支持：Worksheets(2).Activate()）
     * method 'Worksheets.Activate'。index/name 二选一。
     */
    switchSheet: guard(function (p) {
      if (!p.name && p.index === undefined) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 name/index（switchSheet 要切到哪张表）' });
      }
      var payload = {};
      if (p.name) payload.name = String(p.name);
      if (p.index !== undefined) payload.index = Number(p.index);
      return jsapi.jsAsynCall('Worksheets.Activate', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * moveSheet —— S1-REPORT.md:80（§3 #54 支持：ws.Move(before, after)）
     * method 'Worksheet.Move'。
     */
    moveSheet: guard(function (p) {
      var payload = {};
      if (p.target) payload.target = String(p.target);
      if (p.before !== undefined) payload.before = Number(p.before);
      if (p.after !== undefined) payload.after = Number(p.after);
      return jsapi.jsAsynCall('Worksheet.Move', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * mergeCells —— S1-REPORT.md:81（§3 #55 支持：range.Merge()）
     * method 'Range.Merge'。
     */
    mergeCells: guard(function (p) {
      var ref = mustRef(p, 'mergeCells');
      return jsapi.jsAsynCall('Range.Merge', { ref: ref }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * unmergeCells —— S1-REPORT.md:82（§3 #56 支持：range.UnMerge()）
     * method 'Range.UnMerge'。
     */
    unmergeCells: guard(function (p) {
      var ref = mustRef(p, 'unmergeCells');
      return jsapi.jsAsynCall('Range.UnMerge', { ref: ref }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * autoFitAll —— S1-REPORT.md:87（§3 #61 支持：UsedRange 行列全 AutoFit）
     * method 'UsedRange.AutoFit'（与 autoFit #44 同 method，S1 证据同源）。
     */
    autoFitAll: guard(function (p) {
      return jsapi.jsAsynCall('UsedRange.AutoFit', { sheet: p.sheet }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * setPrintArea —— S1-REPORT.md:90（§3 #64 支持：PageSetup.PrintArea 读回 $A$1:$D$5）
     * method 'PageSetup.PrintArea'。area 串必填。
     */
    setPrintArea: guard(function (p) {
      if (!p.area) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 area（setPrintArea，如 $A$1:$D$5）' });
      }
      return jsapi.jsAsynCall('PageSetup.PrintArea', { area: String(p.area) }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getSelection —— S1-REPORT.md:91（§3 #65 支持；§4 #5 报告行 334：sel.Address 是函数须调用）
     * method 'Selection.Address'（读）。
     */
    getSelection: guard(function () {
      return jsapi.jsAsynCall('Selection.Address', {}).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * clearRange —— S1-REPORT.md:92（§3 #66 支持：range.ClearContents()）
     * method 'Range.ClearContents'。
     */
    clearRange: guard(function (p) {
      var ref = mustRef(p, 'clearRange');
      return jsapi.jsAsynCall('Range.ClearContents', { ref: ref }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * freezePanes —— S1-REPORT.md:103（§3 #77 支持：ActiveWindow.FreezePanes = true）
     * method 'ActiveWindow.FreezePanes'（属性写编码，待补测）。
     */
    freezePanes: guard(function () {
      return jsapi.jsAsynCall('ActiveWindow.FreezePanes', { freeze: true }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * unfreezePanes —— S1-REPORT.md:104（§3 #78 支持：同属性 = false）
     * method 'ActiveWindow.FreezePanes'（freeze:false）。
     */
    unfreezePanes: guard(function () {
      return jsapi.jsAsynCall('ActiveWindow.FreezePanes', { freeze: false }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * copyRange —— S1-REPORT.md:106（§3 #80 支持：src.Copy()）
     * method 'Range.Copy'。源 ref 必填。
     */
    copyRange: guard(function (p) {
      var ref = mustRef(p, 'copyRange');
      var payload = { ref: ref };
      if (p.destination) payload.destination = String(p.destination);
      return jsapi.jsAsynCall('Range.Copy', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * pasteRange —— S1-REPORT.md:107（§3 #81 支持：dst.PasteSpecial()）
     * method 'Range.PasteSpecial'。目标 ref 必填。
     */
    pasteRange: guard(function (p) {
      var ref = mustRef(p, 'pasteRange');
      var payload = { ref: ref };
      if (p.pasteMode !== undefined) payload.pasteMode = Number(p.pasteMode);
      return jsapi.jsAsynCall('Range.PasteSpecial', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * fillSeries —— S1-REPORT.md:108（§3 #82 支持：range.DataSeries / FillDown）
     * method 'Range.DataSeries'。通路已证、读回未逐项验证（S1 §10.3）。
     */
    fillSeries: guard(function (p) {
      var ref = mustRef(p, 'fillSeries');
      var payload = { ref: ref };
      if (p.direction !== undefined) payload.direction = String(p.direction);
      return jsapi.jsAsynCall('Range.DataSeries', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * textToColumns —— S1-REPORT.md:110（§3 #84 支持：range.TextToColumns）
     * method 'Range.TextToColumns'。通路已证、读回未逐项验证（S1 §10.3）。
     */
    textToColumns: guard(function (p) {
      var ref = mustRef(p, 'textToColumns');
      var payload = { ref: ref };
      if (p.delimiter !== undefined) payload.delimiter = String(p.delimiter);
      return jsapi.jsAsynCall('Range.TextToColumns', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * subtotal —— S1-REPORT.md:111（§3 #85 支持：range.Subtotal(1, 1, …)）
     * method 'Range.Subtotal'。S1 证据两参按位下发。
     */
    subtotal: guard(function (p) {
      var ref = mustRef(p, 'subtotal');
      var payload = { ref: ref };
      if (p.groupBy !== undefined) payload.groupBy = Number(p.groupBy);
      if (p.function !== undefined) payload['function'] = Number(p['function']);
      if (p.totalList !== undefined) payload.totalList = p.totalList;
      return jsapi.jsAsynCall('Range.Subtotal', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * createNamedRange —— S1-REPORT.md:112（§3 #86 支持：Names.Add 实测成功 Count=2）
     * method 'Names.Add'。注意 §4 #16（报告行 345）：wb.Names 是集合不可当函数——Add 走
     * 'Names.Add' method，Item 系读同理；name/refersTo 必填。
     */
    createNamedRange: guard(function (p) {
      if (!p.name || !p.refersTo) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] name 与 refersTo 必填（createNamedRange，如 refersTo=Sheet1!$A$1:$B$2）' });
      }
      return jsapi.jsAsynCall('Names.Add', { name: String(p.name), refersTo: String(p.refersTo) }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getNamedRanges —— S1-REPORT.md:114（§3 #88 支持：Names.Count + 名称）
     * method 'Names'（读）。
     */
    getNamedRanges: guard(function () {
      return jsapi.jsAsynCall('Names', {}).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * protectSheet —— S1-REPORT.md:118（§3 #92 支持：ws.Protect()）
     * method 'Worksheet.Protect'。
     */
    protectSheet: guard(function (p) {
      var payload = {};
      if (p.password) payload.password = String(p.password);
      if (p.sheet) payload.sheet = String(p.sheet);
      return jsapi.jsAsynCall('Worksheet.Protect', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * unprotectSheet —— S1-REPORT.md:119（§3 #93 支持：ws.Unprotect()）
     * method 'Worksheet.Unprotect'。
     */
    unprotectSheet: guard(function (p) {
      var payload = {};
      if (p.password) payload.password = String(p.password);
      if (p.sheet) payload.sheet = String(p.sheet);
      return jsapi.jsAsynCall('Worksheet.Unprotect', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * protectWorkbook —— S1-REPORT.md:120（§3 #94 支持：wb.Protect()）
     * method 'Workbook.Protect'。
     */
    protectWorkbook: guard(function (p) {
      var payload = {};
      if (p.password) payload.password = String(p.password);
      if (p.structureOnly !== undefined) payload.structureOnly = !!p.structureOnly;
      return jsapi.jsAsynCall('Workbook.Protect', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * insertExcelImage —— S1-REPORT.md:121（§3 #95 支持：Shapes.AddPicture，Shapes.Count=1）
     * method 'Shapes.AddPicture'。path 必填；定位缺省 100/100/尺寸缺省由原生侧决定。
     */
    insertExcelImage: guard(function (p) {
      var path = p.path || p.imagePath;
      if (!path) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 path/imagePath（insertExcelImage 必填）' });
      }
      var payload = { path: String(path) };
      if (p.left !== undefined) payload.left = Number(p.left);
      if (p.top !== undefined) payload.top = Number(p.top);
      if (p.width !== undefined) payload.width = Number(p.width);
      if (p.height !== undefined) payload.height = Number(p.height);
      return jsapi.jsAsynCall('Shapes.AddPicture', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * setHyperlink —— S1-REPORT.md:122（§3 #96 支持：ws.Hyperlinks.Add(anchor, address)）
     * method 'Hyperlinks.Add'。anchor（anchorRef）与 address 必填。
     */
    setHyperlink: guard(function (p) {
      var anchor = p.anchorRef || p.anchor || ExcelAdapter.refOf(p);
      if (!anchor || !p.address) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] anchor（锚点引用）与 address 必填（setHyperlink）' });
      }
      var payload = { anchor: String(anchor), address: String(p.address) };
      if (p.textToDisplay) payload.textToDisplay = String(p.textToDisplay);
      return jsapi.jsAsynCall('Hyperlinks.Add', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * openWorkbook —— S1-REPORT.md:124（§3 #98 支持：Workbooks.Open）
     * method 'Workbooks.Open'。path 必填（TS 双键 path/filePath 都认）。
     */
    openWorkbook: guard(function (p) {
      var path = p.path || p.filePath;
      if (!path) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 path/filePath（openWorkbook 必填）' });
      }
      return jsapi.jsAsynCall('Workbooks.Open', { path: String(path) }).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * getOpenWorkbooks —— S1-REPORT.md:125（§3 #99 支持：Workbooks.Count + 名称）
     * method 'Workbooks'（读）。
     */
    getOpenWorkbooks: guard(function () {
      return jsapi.jsAsynCall('Workbooks', {}).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * switchWorkbook —— S1-REPORT.md:126（§3 #100 支持：Workbooks(2).Activate()）
     * method 'Workbooks.Activate'。index/name 二选一。
     */
    switchWorkbook: guard(function (p) {
      if (!p.name && p.index === undefined) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 name/index（switchWorkbook 要切到哪个簿）' });
      }
      var payload = {};
      if (p.name) payload.name = String(p.name);
      if (p.index !== undefined) payload.index = Number(p.index);
      return jsapi.jsAsynCall('Workbooks.Activate', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * closeWorkbook —— S1-REPORT.md:127（§3 #101 支持：Workbooks.Item(2).Close()）
     * method 'Workbooks.Close'。⚠ B2 已知问题（Windows 侧）"从未保存的簿关不掉"的
     * 对应语义在 et 侧未测——saveChanges 透传，B3.5 复核。
     */
    closeWorkbook: guard(function (p) {
      if (!p.name && p.index === undefined) {
        return Promise.resolve({ ok: false, error: '[勿重试·先改参数] 缺少 name/index（closeWorkbook 要关哪个簿）' });
      }
      var payload = {};
      if (p.name) payload.name = String(p.name);
      if (p.index !== undefined) payload.index = Number(p.index);
      if (p.saveChanges !== undefined) payload.saveChanges = !!p.saveChanges;
      return jsapi.jsAsynCall('Workbooks.Close', payload).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    /**
     * createWorkbook —— S1-REPORT.md:128（§3 #102 支持：Workbooks.Add，Count 0→1 实测）
     * method 'Workbooks.Add'。
     */
    createWorkbook: guard(function () {
      return jsapi.jsAsynCall('Workbooks.Add', {}).then(function (r) {
        return { ok: true, data: r };
      });
    }),

    // ========================================================================
    // PivotTables 系 3 桩：S1 判定"支持"但证据只是"集合存在"（Count=undefined、
    // Item 不可调，§3:35/74/75 + §4 #6）——任务卡 §0 明令标待补测，绝不猜 method。
    // ========================================================================

    /**
     * createPivotTable —— S1-REPORT.md:35（§3 #9：证据"ws.PivotTables 存在, Count=undefined（Add 未验证）"）
     * unsupported:'待S1补测'——S1 未验证 Add，method 无从编码。
     */
    createPivotTable: guard(function () {
      return Promise.resolve({
        ok: false,
        unsupported: '待S1补测',
        error: 'createPivotTable：S1 只证 PivotTables 集合存在（Add 未验证），method 待补测后实现',
      });
    }),

    /**
     * getPivotTables —— S1-REPORT.md:74（§3 #48：证据 Count=undefined）
     * unsupported:'待S1补测'。
     */
    getPivotTables: guard(function () {
      return Promise.resolve({
        ok: false,
        unsupported: '待S1补测',
        error: 'getPivotTables：S1 实测 Count=undefined，读侧语义待补测',
      });
    }),

    /**
     * refreshPivotTables —— S1-REPORT.md:75（§3 #49：证据 Count=undefined）
     * unsupported:'待S1补测'。
     */
    refreshPivotTables: guard(function () {
      return Promise.resolve({
        ok: false,
        unsupported: '待S1补测',
        error: 'refreshPivotTables：S1 实测 Count=undefined（无透视表则无需刷新），刷新语义待补测',
      });
    }),
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = { actions: excelActions, adapter: ExcelAdapter };
  } else {
    global.WpsAddonActionsExcel = excelActions;
    global.WpsExcelAdapter = ExcelAdapter;
  }
})(typeof window !== 'undefined' ? window : this);
