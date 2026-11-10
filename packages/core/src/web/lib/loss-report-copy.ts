/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Loss reports in the reader's language.
 *
 * The file adapters write a report's summary (`user_facing_text`) and some
 * entry details as Chinese sentences, and the report is stored as written.
 * The sentences they compose are a closed set: the opener, the "nothing is
 * lost" line, the XLSX recommendations, the code editor's deferred
 * capabilities. Each has a `console-notebooks` key, `lossReport.server.*`,
 * and the report reads them in the reader's language when shown; anything
 * else (a file's own object names) stays as written.
 */

type Translate = {
  (key: string, values?: Record<string, string | number>): string;
  has?(key: string): boolean;
};

const known = (t: Translate, key: string) => typeof t.has === "function" && t.has(key);

const SERVER_SENTENCES: ReadonlyArray<readonly [server: string, key: string]> = [
  ["这次导出/投影里有以下损耗：", "lossReport.server.openerExport"],
  ["这次编辑/多文件协调里有以下损耗：", "lossReport.server.openerCode"],
  ["这次媒体导出/转码里有以下损耗：", "lossReport.server.openerMedia"],
  ["本次 XLSX 转换无损耗，所有字段保留原样。", "lossReport.server.xlsxLossless"],
  ["（本次无损耗）", "lossReport.server.noLoss"],
  ["（多文件改动需逐文件确认后才会应用）", "lossReport.server.codeConfirmEachFile"],
  [
    "建议在 LibreOffice 中复核图表系列与坐标轴；编辑时尽量保留原始系列数据。",
    "lossReport.server.xlsxChart",
  ],
  [
    "建议在原始 Excel 客户端中刷新透视缓存；本工具仅保留数据透视表的结构。",
    "lossReport.server.xlsxPivot",
  ],
  [
    "宏（VBA）按默认策略隔离：保留二进制但不会执行，也不可在本工具内编辑。",
    "lossReport.server.xlsxMacro",
  ],
  [
    "建议在原始客户端确认主题与表格样式；本工具会近似保留单元格字体、填充、边框。",
    "lossReport.server.xlsxStyle",
  ],
  [
    "公式文本按字符串保留；共享公式组不可逐单元格编辑。如需修改或求值，请在受信任的本地表格客户端中打开。",
    "lossReport.server.xlsxFormula",
  ],
  [
    "外部链接按默认策略隔离：保留引用但不会跟随；请在受信任的环境中手动刷新。",
    "lossReport.server.xlsxExternalLink",
  ],
  ["LSP 语义级重构（跨文件重命名 / 跳转定义 / 悬停）", "lossReport.server.codeLsp"],
  ["IDE 代码智能（补全 / 悬停 / 诊断）", "lossReport.server.codeIde"],
  ["跨语言依赖解析（超出逐导入解析的范围）", "lossReport.server.codeCrossLanguage"],
  ["git 提交 / 分支 / 推送操作", "lossReport.server.codeGit"],
  ["black 代码格式化", "lossReport.server.codeBlack"],
  ["prettier 代码格式化", "lossReport.server.codePrettier"],
  ["测试运行验证", "lossReport.server.codeTests"],
  ["编译器调用", "lossReport.server.codeCompiler"],
];

/** A media entry's advice, "（建议：…）". */
const ADVICE = /（建议：([^）]*)）/g;

/** The XLSX loss categories, named for a list (`lossReport.kind.*`). */
const KIND_KEYS: Readonly<Record<string, string>> = {
  chart_loss: "lossReport.kind.chart_loss",
  pivot_loss: "lossReport.kind.pivot_loss",
  macro_loss: "lossReport.kind.macro_loss",
  style_loss: "lossReport.kind.style_loss",
  formula_loss: "lossReport.kind.formula_loss",
  external_link_loss: "lossReport.kind.external_link_loss",
};

/** A report text with the adapters' own sentences in the reader's language. */
export function localizeLossReportText(text: string, t: Translate): string {
  if (!text) return text;
  let next = text;
  for (const [server, key] of SERVER_SENTENCES) {
    if (next.includes(server) && known(t, key)) next = next.split(server).join(t(key));
  }
  if (known(t, "lossReport.server.advice")) {
    next = next.replace(ADVICE, (_, advice: string) => t("lossReport.server.advice", { advice }));
  }
  // The Chinese punctuation the sentences were joined with stays only where
  // the reader's own catalogue writes it.
  if (known(t, "lossReport.server.noLoss") && !t("lossReport.server.noLoss").includes("（")) {
    next = next.replaceAll("：", ": ").replaceAll("（", " (").replaceAll("）", ")").replace(/ {2,}/g, " ");
  }
  return next;
}

/** A loss kind's name for a list, else the code in words. */
export function lossKindLabel(kind: string, t: Translate): string {
  const key = KIND_KEYS[kind];
  return key && known(t, key) ? t(key) : kind.replaceAll("_", " ");
}
