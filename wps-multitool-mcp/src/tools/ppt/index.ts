/**
 * Input: PPT 工具定义
 * Output: PPT 工具注册数组
 * Pos: PPT Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * - 动画Tools: remove_animation, get_animations（设置/顺序已并入 set_animation，第 8 轮）
 * - deckkit 引擎Tools: build_ppt_deck（JSON 页面描述 → 整套演示文稿；E2 骨架 + E3 主题三套/生成后溢出检查）、
 *   get_textbox_metrics（文本框排版度量，溢出检查读回通道，只读）
 */

import { RegisteredTool } from '../../types/tools';
import { slideTools } from './slide';
import { slideOpsTools } from './slide-ops';
import { presentationTools } from './presentation';
import { textboxTools } from './textbox';
import { pageInfoTools } from './page-info';
import { pageEffectsTools } from './page-effects';
import { shapeBasicTools } from './shape-basic';
import { shapeStyleTools } from './shape-style';
import { imageTools } from './image';
import { animationTools } from './animation';
import { chartFlowTools } from './chart-flow';
import { miscTools } from './misc';
import { dataVizTools } from './data-viz';
import { tableTools } from './table';
import { beautifyAdvancedTools } from './beautify-advanced';
import { deckTools } from './deckkit';
import { batchShapesTools } from './batch-shapes';

/**
 * 所有PPT相关的Tools
 * 包含：
 * - 幻灯片Tools: add_slide, beautify, unify_font, set_font_color, align_objects
 * - 幻灯片操作Tools: delete_slide, move_slide, get_slide_count, get_slide_info, switch_slide,
 *   set_slide_layout, get_slide_notes, add_shape, set_shape_style, add_textbox, set_slide_title,
 *   insert_image, set_shape_text, set_animation, set_background, set_slide_size, set_transition,
 *   set_shape_fill, add_speaker_notes
 * - 演示文稿管理Tools: create_presentation, open_presentation, close_presentation, get_open_presentations,
 *   switch_presentation, set_slide_theme, copy_slide, insert_slide_image, insert_slides_from_file, set_active_target
 * - 文本框Tools: delete_textbox, get_textboxes, set_textbox_text, set_textbox_style,
 *   get_slide_title, set_slide_subtitle, set_slide_content
 * - 页面信息Tools: set_slide_number, set_ppt_footer, set_ppt_date_time, duplicate_shape, set_shape_z_order
 * - 表格Tools: insert_table, set_table_cell, get_table_cell, set_table_style, set_table_cell_style, set_table_row_style
 * - 高级美化Tools: apply_color_scheme, create_kpi_cards,
 *   create_styled_table, add_title_decoration, add_page_indicator
 */
export const pptTools: RegisteredTool[] = [
  ...slideTools,
  ...slideOpsTools,
  ...presentationTools,
  ...textboxTools,
  ...pageInfoTools,
  ...pageEffectsTools,
  // 注册顺序与 shape 拆分前逐一对应：基础前 5 → 样式前 4 → 透明度/文字 → 填充 → 标题/对齐/分布/组合
  ...shapeBasicTools.slice(0, 5),
  ...shapeStyleTools.slice(0, 4),
  ...shapeBasicTools.slice(5, 7),
  ...shapeStyleTools.slice(4, 5),
  ...shapeBasicTools.slice(7),
  ...imageTools,
  ...animationTools,
  ...chartFlowTools,
  ...miscTools,
  ...dataVizTools,
  ...tableTools,
  ...beautifyAdvancedTools,
  ...deckTools,
  ...batchShapesTools,
];

// 分别导出，方便按需使用
export { slideTools } from './slide';
export { slideOpsTools } from './slide-ops';
export { presentationTools } from './presentation';
export { textboxTools } from './textbox';
export { pageInfoTools } from './page-info';
export { pageEffectsTools } from './page-effects';
export { shapeBasicTools } from './shape-basic';
export { shapeStyleTools } from './shape-style';
export { imageTools } from './image';
export { animationTools } from './animation';
export { chartFlowTools } from './chart-flow';
export { miscTools } from './misc';
export { dataVizTools } from './data-viz';
export { tableTools } from './table';
export { beautifyAdvancedTools } from './beautify-advanced';
export { deckTools } from './deckkit';
export { batchShapesTools } from './batch-shapes';

// 导出单独的定义和处理器，方便测试
export {
  addSlideDefinition,
  addSlideHandler,
  beautifyDefinition,
  beautifyHandler,
  unifyFontDefinition,
  unifyFontHandler,
  setFontColorDefinition,
  setFontColorHandler,
} from './slide';

export {
  deleteSlideDefinition,
  deleteSlideHandler,
  moveSlideDefinition,
  moveSlideHandler,
  getSlideCountDefinition,
  getSlideCountHandler,
  getSlideInfoDefinition,
  getSlideInfoHandler,
  switchSlideDefinition,
  switchSlideHandler,
  setSlideLayoutDefinition,
  setSlideLayoutHandler,
  getSlideNotesDefinition,
  getSlideNotesHandler,
  setSlideSizeDefinition,
  setSlideSizeHandler,
} from './slide-ops';

// 形状与文本 6 个自 slide-ops 迁入（拆分轮）；insert_image 自 slide-ops 迁入 image.ts
export {
  addShapeDefinition,
  addShapeHandler,
  addTextboxDefinition,
  addTextboxHandler,
  setSlideTitleDefinition,
  setSlideTitleHandler,
  setShapeTextDefinition,
  setShapeTextHandler,
} from './shape-basic';

// 形状样式类五件（style/shadow/gradient/border/fill）自 shape-basic 拆出
export {
  setShapeStyleDefinition,
  setShapeStyleHandler,
  setShapeShadowDefinition,
  setShapeShadowHandler,
  setShapeGradientDefinition,
  setShapeGradientHandler,
  setShapeBorderDefinition,
  setShapeBorderHandler,
  setShapeFillDefinition,
  setShapeFillHandler,
} from './shape-style';

export {
  setAnimationDefinition,
  setAnimationHandler,
  setBackgroundDefinition,
  setBackgroundHandler,
  setTransitionDefinition,
  setTransitionHandler,
  addSpeakerNotesDefinition,
  addSpeakerNotesHandler,
} from './page-effects';

export {
  createPresentationDefinition,
  createPresentationHandler,
  openPresentationDefinition,
  openPresentationHandler,
  closePresentationDefinition,
  closePresentationHandler,
  getOpenPresentationsDefinition,
  getOpenPresentationsHandler,
  switchPresentationDefinition,
  switchPresentationHandler,
  setSlideThemeDefinition,
  setSlideThemeHandler,
  copySlideDefinition,
  copySlideHandler,
  insertSlidesFromFileDefinition,
  insertSlidesFromFileHandler,
  setActiveTargetDefinition,
  setActiveTargetHandler,
} from './presentation';

export {
  deleteTextboxDefinition,
  deleteTextboxHandler,
  getTextboxesDefinition,
  getTextboxesHandler,
  setTextboxTextDefinition,
  setTextboxTextHandler,
  setTextboxStyleDefinition,
  setTextboxStyleHandler,
  getSlideTitleDefinition,
  getSlideTitleHandler,
  setSlideSubtitleDefinition,
  setSlideSubtitleHandler,
  setSlideContentDefinition,
  setSlideContentHandler,
} from './textbox';

export {
  setSlideNumberDefinition,
  setSlideNumberHandler,
  setPptFooterDefinition,
  setPptFooterHandler,
  setPptDateTimeDefinition,
  setPptDateTimeHandler,
  duplicateShapeDefinition,
  duplicateShapeHandler,
  setShapeZOrderDefinition,
  setShapeZOrderHandler,
} from './page-info';

export {
  deleteShapeDefinition,
  deleteShapeHandler,
  getShapesDefinition,
  getShapesHandler,
  setShapePositionDefinition,
  setShapePositionHandler,
  setShapeTransparencyDefinition,
  setShapeTransparencyHandler,
  alignShapesDefinition,
  alignShapesHandler,
  distributeShapesDefinition,
  distributeShapesHandler,
  groupShapesDefinition,
  groupShapesHandler,
} from './shape-basic';

export {
  insertImageDefinition,
  insertImageHandler,
  deletePptImageDefinition,
  deletePptImageHandler,
  setImageStyleDefinition,
  setImageStyleHandler,
  exportSlideAsImageDefinition,
  exportSlideAsImageHandler,
  replacePptImageDefinition,
  replacePptImageHandler,
} from './image';

export {
  removeAnimationDefinition,
  removeAnimationHandler,
  getAnimationsDefinition,
  getAnimationsHandler,
} from './animation';

export {
  insertPptChartDefinition,
  insertPptChartHandler,
  setPptChartDataDefinition,
  setPptChartDataHandler,
  setPptChartStyleDefinition,
  setPptChartStyleHandler,
  createFlowChartDefinition,
  createFlowChartHandler,
  createOrgChartDefinition,
  createOrgChartHandler,
  createTimelineDefinition,
  createTimelineHandler,
} from './chart-flow';

export {
  getSlideMasterDefinition,
  getSlideMasterHandler,
  addMasterElementDefinition,
  addMasterElementHandler,
  set3DRotationDefinition,
  set3DRotationHandler,
  set3DDepthDefinition,
  set3DDepthHandler,
  set3DMaterialDefinition,
  set3DMaterialHandler,
  create3DTextDefinition,
  create3DTextHandler,
  addPptHyperlinkDefinition,
  addPptHyperlinkHandler,
  removePptHyperlinkDefinition,
  removePptHyperlinkHandler,
  findPptTextDefinition,
  findPptTextHandler,
  replacePptTextDefinition,
  replacePptTextHandler,
  startSlideShowDefinition,
  startSlideShowHandler,
} from './misc';

export {
  createProgressBarDefinition,
  createProgressBarHandler,
  createGaugeDefinition,
  createGaugeHandler,
  createMiniChartsDefinition,
  createMiniChartsHandler,
  createDonutChartDefinition,
  createDonutChartHandler,
  autoLayoutDefinition,
  autoLayoutHandler,
  smartDistributeDefinition,
  smartDistributeHandler,
  createGridDefinition,
  createGridHandler,
} from './data-viz';

export {
  insertPptTableDefinition,
  insertPptTableHandler,
  setPptTableCellDefinition,
  setPptTableCellHandler,
  getPptTableCellDefinition,
  getPptTableCellHandler,
  setPptTableStyleDefinition,
  setPptTableStyleHandler,
  setPptTableCellStyleDefinition,
  setPptTableCellStyleHandler,
  setPptTableRowStyleDefinition,
  setPptTableRowStyleHandler,
} from './table';

export {
  applyColorSchemeDefinition,
  applyColorSchemeHandler,
  createKpiCardsDefinition,
  createKpiCardsHandler,
  createStyledTableDefinition,
  createStyledTableHandler,
  addTitleDecorationDefinition,
  addTitleDecorationHandler,
  addPageIndicatorDefinition,
  addPageIndicatorHandler,
} from './beautify-advanced';

export default pptTools;
