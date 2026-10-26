/**
 * skill 内容真伪门禁（第 10 轮 T1）
 * --------------------------------------------------
 * 为什么单独一个文件：SKILL.md 是 **Agent 读来决定怎么用这个 MCP 的**，
 * 它说错话 = Agent 对用户说错话。而 skill-verify.js 原来只验 frontmatter/长度/大小，
 * 20 条断言全绿时正放过「PPT 相关的工具没有真机覆盖验证过……整个 ppt 组都是 static」
 * 这句当时已经为假的话（第 9 轮 PPT 隔离验证做到 98/98）。
 *
 * 这里做两件事，缺一不可：
 *   A) **已知失效断言黑名单** —— 挡住那句原话，以及同义的改写。
 *      黑名单只收"曾经被实测证伪的具体句子"，不是收关键词。
 *   B) **拿台账真值逐条对账** —— 把 SKILL.md 里每一个「分组 + 验证结论」抽出来，
 *      和 scripts/manual/tool-verification.json 比对；对不上就红。
 *      B 是主力：A 只是补漏网的口语化说法。
 *
 * 为什么不直接 require 生成物：门禁要在 dist 没 build 时也能跑（CI 早一步），
 * 所以这里读 **真相源 json**（生成物由 gen-tool-verification.js 逐字搬过去，
 * 两边对不上由 verification-verify 负责）。
 *
 * 口径：live/static 两档是产品语义，static 不是"不能用"。
 * 所以门禁【不】禁止文档出现 static，只禁止**与台账矛盾**的说法
 * （例如台账里有 live 的工具，文档却说该组"没有真机覆盖"）。
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const LEDGER = path.join(ROOT, 'scripts/manual/tool-verification.json');

/** 台账真值 */
function loadLedger() {
  if (!fs.existsSync(LEDGER)) {
    return { error: `找不到验证台账 ${LEDGER}` };
  }
  return JSON.parse(fs.readFileSync(LEDGER, 'utf8'));
}

/**
 * 已知失效断言（第 10 轮实测证伪，逐条附出处）。
 * 只收"整句"，不收关键词 —— 关键词会误伤正确写法（比如文档里引用这句话本身时）。
 * 每条 retired 的右侧给出"当时为真的证据"，方便后来人判断该更新还是该删。
 */
const RETIRED_CLAIMS = [
  {
    id: 'ppt-no-mach-cover',
    re: /PPT\s*相关的工具没有真机覆盖验证过/,
    label: '「PPT 相关的工具没有真机覆盖验证过」',
    why: '第 9 轮 scripts/manual/ppt-isolated-verify.js 真机逐条跑通 98/98（矩阵 103 行 = 98 执行 + 5 跳过）',
  },
  {
    id: 'ppt-whole-group-static',
    re: /整个\s*ppt\s*组都是\s*static|目前整个\s*ppt\s*组都是\s*static/i,
    label: '「目前整个 ppt 组都是 static」',
    why: '台账 tools 里逐条记 live 的 PPT 工具有 ' +
         '若干个（create_presentation / add_slide / set_shape_text / set_shape_fill / ' +
         'set_table_cell / insert_ppt_chart / align_objects …），说"整个组都是 static"与台账直接矛盾',
  },
  {
    id: 'ppt-not-verified-blanket',
    re: /PPT\s*没有\s*真机\s*覆盖验证过/,
    label: '「PPT 没有真机覆盖验证过」',
    why: '同 ppt-no-mach-cover',
  },
];

/** 组名 → 台账里的组级状态 */
function groupStatusOf(ledger) {
  return ledger.groups || {};
}

/**
 * 从证据串里抠数字。
 *
 * 约定：evidence 串里用 `键=值` 的显式写法（如 `矩阵行数=103`），
 * 而不是靠中文散文正则去猜。这样两边不会各说各话 ——
 * 改 prose 不会悄悄让门禁失灵，写错值门禁立刻红。
 */
function numbersFromEvidence(evidence) {
  const pairs = {};
  for (const m of String(evidence).matchAll(/([\u4e00-\u9fa5A-Za-z]+)\s*=\s*(\d+)/g)) {
    pairs[m[1]] = Number(m[2]);
  }
  const pick = (...keys) => {
    for (const k of keys) if (pairs[k] !== undefined) return pairs[k];
    return undefined;
  };
  return {
    matrixRows: pick('矩阵行数'),
    executed: pick('执行条数'),
    skipped: pick('跳过条数'),
    groupTools: pick('PPT组工具数', 'PPT 组工具数'),
    liveTools: pick('逐条live数', 'live工具数'),
    pairs,
  };
}

/** 这句话是不是"在禁止某说法"，而不是"在断言某说法" */
function isMetaSentence(s) {
  return /(别说|不要说|不得说|禁止|别笼统|是错的|不对|切勿|勿)/.test(s);
}

/**
 * 找出文档里所有「组名 + 验证结论」的断言。
 * 认得的形态：同一句里同时出现组名（ppt/word/excel/common/core）和状态词
 * （live / static / 真机 / 未验证 / 没验证）。
 * 【跳过】"别这么说"的那几句 —— 第 10 轮实测：SKILL.md 里写
 * 「别说『PPT 没验证过』」是正确写法，却会被当成断言而误报红。
 */
function extractGroupClaims(text, groups) {
  const sentences = String(text)
    .split(/[\n。；;]/)
    .map((s) => s.trim())
    .filter(Boolean);
  const claims = [];
  for (const s of sentences) {
    if (isMetaSentence(s)) continue;
    const low = s.toLowerCase();
    for (const g of groups) {
      if (!low.includes(g)) continue;
      if (!/(live|static|真机|未验证|没验证|验证过|校验)/.test(low)) continue;
      claims.push({ group: g, sentence: s.slice(0, 160) });
      break;
    }
  }
  return claims;
}

/**
 * 判断一句断言是否与台账矛盾。
 * 规则（故意保守，宁可放过也不误伤）：
 *   1) 说"整组/全部 + static"而台账里该组有 live 工具 → 矛盾
 *   2) 说"没有真机/没验证过"而台账里该组有 live 工具 → 矛盾
 *   3) 说"整组/全部 + live"而台账组级是 static → 矛盾（反过来也会教错话）
 */
function contradictionOf(claim, ledger) {
  const s = claim.sentence;
  const low = s.toLowerCase();
  const groupLiveTools = Object.entries(ledger.tools || {}).filter(
    ([name, st]) => st === 'live' && name.includes(`_${claim.group}_`)
  );
  const gStatus = (ledger.groups || {})[claim.group];

  const blanket = /(整个|整组|全部|所有)/.test(s) || /^(?:整个|整组)/.test(s);
  const saysNoLive = /(没有|未|没)\s*真机/.test(s) || /(没有|未|没)验证过/.test(s);
  const saysStatic = /static/.test(low);
  const saysLive = /\blive\b/.test(low);

  if (saysNoLive && groupLiveTools.length > 0) {
    return `台账里 ${claim.group} 组有 ${groupLiveTools.length} 个 live 工具，文档却说"${s}"`;
  }
  if (blanket && saysStatic && groupLiveTools.length > 0) {
    return `台账里 ${claim.group} 组有 ${groupLiveTools.length} 个 live 工具，"${s}" 与之矛盾`;
  }
  if (blanket && saysLive && gStatus === 'static') {
    return `台账组级 ${claim.group} = static，文档却说"${s}"`;
  }
  return null;
}

module.exports = {
  RETIRED_CLAIMS,
  loadLedger,
  groupStatusOf,
  numbersFromEvidence,
  extractGroupClaims,
  isMetaSentence,
  contradictionOf,
  LEDGER,
};