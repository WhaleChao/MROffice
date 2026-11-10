/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Simplified Chinese function help (XLS-A094 / XLS-A067) — the description
// Excel's zh-CN build shows beside a function in the AutoComplete ScreenTip, the
// Insert Function dialog and the Function Arguments dialog, plus the line for
// each argument box.
//
// PURE data, no imports. Loaded lazily (dynamic import from
// function-docs-locale.ts) so a non-Chinese session never downloads it.
//
// One sentence per function and one phrase per argument, following Excel's own
// zh-CN wording where it is known and otherwise translating the function's
// Microsoft Support reference page. Every engine function has an entry (a test
// asserts it): `summary` plus `args`, keyed by the bracket-stripped signature
// label from function-signatures.ts. A label a function does not spell out is
// looked up in ZH_COMMON_ARGUMENTS, which holds the arguments that mean the same
// thing in every function that takes them (settlement, basis, inumber, …).
// ---------------------------------------------------------------------------

export type ZhFunctionDoc = readonly [summary: string, args?: Readonly<Record<string, string>>];

/** Arguments with one meaning across every function that takes them. */
export const ZH_COMMON_ARGUMENTS: Readonly<Record<string, string>> = /* own data: ZH_COMMON_ARGUMENTS */ {};

// Argument sets shared by a family of functions.
const DB_CRITERIA = "包含指定条件的单元格区域。区域包括列标签及列标签下满足某个条件的单元格";
const FORECAST_ARGS = {
  x: "需要进行值预测的数据点，必须为数值",
  known_ys: "从属数组或数值数据区域",
  known_xs: "独立数组或数值数据区域，其方差不能为零",
};
const ETS_SERIES_ARGS = {
  values: "要预测的数值数组或区域",
  timeline: "独立的数值数据数组或区域，时间线中的日期之间的步长必须一致且不能为零",
  seasonality: "表示季节性模式长度的可选数值；默认值 1 表示自动检测季节性",
  data_completion: "处理缺失值的可选值；默认值 1 以插值方式替换缺失值，0 以零替换",
  aggregation: "聚合具有相同时间戳的多个值的可选数值；如果为空，则对其求平均值",
};
const ETS_ARGS = {
  ...ETS_SERIES_ARGS,
  target_date: "要预测值的数据点，应沿用时间线中值的模式",
};
const HYPGEOM_ARGS = {
  sample_s: "样本中成功的次数",
  number_sample: "样本容量",
  population_s: "样本总体中成功的次数",
  number_pop: "样本总体的容量",
};
const ODD_FIRST_ARGS = {
  first_coupon: "有价证券的首期付息日，以一串日期序列号表示",
  rate: "有价证券的利率",
};
const ODD_LAST_ARGS = {
  last_interest: "有价证券的末期付息日，以一串日期序列号表示",
  rate: "有价证券的利率",
};
const RANK_ARGS = {
  number: "指定的数字",
  ref: "一组数或对一个数据列表的引用。非数字值将被忽略",
  order: "指定排位的方式：0 或忽略 = 降序；非零值 = 升序",
};
const STDEV_SAMPLE_ARGS = {
  number1: "与总体抽样样本相应的 1 到 255 个数值，可以是数值或者包含数值的引用",
  number2: "其他与样本相应的数值或引用",
};
const VAR_SAMPLE_ARGS = {
  number1: "与总体抽样样本相应的 1 到 255 个数值参数",
  number2: "其他与样本相应的数值参数",
};
const SUBSTITUTE_ARGS = {
  text: "包含要替换字符的字符串或文本单元格引用",
  old_text: "要被替换的现有字符串。如果原有字符串中的大小写不匹配，将不进行替换",
  new_text: "用于替换 old_text 的新字符串",
  instance_num: "若指定的字符串 old_text 在父字符串中出现多次，则用本参数指定要替换第几个。如果省略，则全部替换",
};
const TTEST_ARGS = {
  array1: "第一个数据集",
  array2: "第二个数据集",
  tails: "指定所返回的分布函数是单尾分布（1）还是双尾分布（2）",
  type: "t 检验的类型：成对检验 = 1；双样本等方差假设 = 2；双样本异方差假设 = 3",
};
const TEXT_SPLIT_AT_ARGS = {
  text: "要在其中搜索的文本",
  delimiter: "标记要提取的点的文本",
  instance_num: "要在其后（前）提取文本的分隔符实例；默认为 1，负数表示从末尾开始搜索",
  match_mode: "是否区分大小写：0 = 区分（默认）；1 = 不区分",
  match_end: "是否将文本末尾视为分隔符：0 = 不视为（默认）；1 = 视为",
  if_not_found: "未找到匹配项时返回的值；默认返回 #N/A",
};
const WEIBULL_ARGS = {
  x: "用来计算函数的值，非负数",
  alpha: "分布参数，正数",
  beta: "分布参数，正数",
};
const ZTEST_ARGS = {
  array: "用来检验 x 的数组或数据区域",
  x: "要检验的值",
  sigma: "总体（已知）标准偏差。如果忽略，则使用样本标准偏差",
};

/**
 * Every engine function, keyed by its upper-case name. `args` gives each
 * argument's line where it is specific to the function; the rest come from
 * ZH_COMMON_ARGUMENTS.
 */
export const ZH_FUNCTION_DOCS: Readonly<Record<string, ZhFunctionDoc>> = /* own data: ZH_FUNCTION_DOCS */ {};
