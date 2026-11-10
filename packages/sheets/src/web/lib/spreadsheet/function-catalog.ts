/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Browse-by-category catalog for the "Insert Function (fx)" dialog (Excel / WPS
 * parity). The dialog lets the user pick a function by CATEGORY (Math, Text,
 * Date, …) and then insert it; this module owns the pure data — the mapping from
 * each engine-supported function name to a category — so the React dialog stays
 * a thin renderer. Categories mirror Excel's function-library tabs.
 *
 * Only functions the engine ACTUALLY evaluates are listed (the source of truth
 * is `authoringFunctionNames()` from formula-intellisense — the engine registry
 * minus what Excel for the web cannot author), so the picker can never offer
 * a phantom function. A function not explicitly mapped lands in "other" so it is
 * still discoverable. The arg syntax hint shown next to a selection reuses the
 * existing `signatureFor` / `formatSignature` helpers.
 */

import { authoringFunctionNames } from "./formula-intellisense";

/** Excel's function-library categories (plus an "all" pseudo-category). */
export type FunctionCategory =
  | "all"
  | "math"
  | "statistical"
  | "logical"
  | "text"
  | "dateTime"
  | "lookup"
  | "financial"
  | "engineering"
  | "information"
  | "database"
  | "web"
  | "other";

/** The ordered list of categories the dialog shows (excluding "other" unless
 * it has members; "all" always first). */
export const FUNCTION_CATEGORY_ORDER: readonly FunctionCategory[] = [
  "all",
  "math",
  "statistical",
  "logical",
  "text",
  "dateTime",
  "lookup",
  "financial",
  "engineering",
  "information",
  "database",
  "web",
  "other",
];

// Explicit category assignments for functions whose category is not obvious from
// a prefix rule. Keyed by UPPERCASE name. Anything not here is classified by the
// prefix rules below, then falls back to "other".
const EXPLICIT: Readonly<Record<string, FunctionCategory>> = {
  // Math & trig
  SUM: "math",
  SUMIF: "math",
  SUMIFS: "math",
  SUMPRODUCT: "math",
  SUMSQ: "math",
  PRODUCT: "math",
  ABS: "math",
  SIGN: "math",
  SQRT: "math",
  POWER: "math",
  EXP: "math",
  LN: "math",
  LOG: "math",
  LOG10: "math",
  MOD: "math",
  QUOTIENT: "math",
  GCD: "math",
  LCM: "math",
  FACT: "math",
  FACTDOUBLE: "math",
  COMBIN: "math",
  PERMUT: "math",
  PERMUTATIONA: "math",
  PI: "math",
  RAND: "math",
  RANDBETWEEN: "math",
  RANDARRAY: "math",
  SUBTOTAL: "math",
  AGGREGATE: "math",
  SEQUENCE: "math",

  // Statistical
  AVERAGE: "statistical",
  AVERAGEIF: "statistical",
  AVERAGEIFS: "statistical",
  AVERAGEA: "statistical",
  COUNT: "statistical",
  COUNTA: "statistical",
  COUNTBLANK: "statistical",
  COUNTIF: "statistical",
  COUNTIFS: "statistical",
  MIN: "statistical",
  MINA: "statistical",
  MINIFS: "statistical",
  MAX: "statistical",
  MAXA: "statistical",
  MAXIFS: "statistical",
  MEDIAN: "statistical",
  MODE: "statistical",
  LARGE: "statistical",
  SMALL: "statistical",
  RANK: "statistical",
  PERCENTILE: "statistical",
  QUARTILE: "statistical",
  FREQUENCY: "statistical",
  CORREL: "statistical",
  COVAR: "statistical",
  SLOPE: "statistical",
  INTERCEPT: "statistical",
  FORECAST: "statistical",
  GEOMEAN: "statistical",
  HARMEAN: "statistical",
  TRIMMEAN: "statistical",
  SKEW: "statistical",
  KURT: "statistical",

  // Logical
  IF: "logical",
  IFS: "logical",
  IFERROR: "logical",
  IFNA: "logical",
  AND: "logical",
  OR: "logical",
  NOT: "logical",
  XOR: "logical",
  TRUE: "logical",
  FALSE: "logical",
  SWITCH: "logical",

  // Lookup & reference
  VLOOKUP: "lookup",
  HLOOKUP: "lookup",
  XLOOKUP: "lookup",
  LOOKUP: "lookup",
  MATCH: "lookup",
  XMATCH: "lookup",
  INDEX: "lookup",
  OFFSET: "lookup",
  INDIRECT: "lookup",
  CHOOSE: "lookup",
  ROW: "lookup",
  ROWS: "lookup",
  COLUMN: "lookup",
  COLUMNS: "lookup",
  TRANSPOSE: "lookup",
  HYPERLINK: "lookup",
  // Excel files IMAGE under Lookup & Reference (not under the IM... engineering
  // functions its name would otherwise suggest).
  IMAGE: "lookup",
  // Excel's More Functions > Web shelf (ENCODEURL, FILTERXML, WEBSERVICE);
  // IMPORTRANGE sits on Google Sheets' Web shelf beside them.
  ENCODEURL: "web",
  FILTERXML: "web",
  WEBSERVICE: "web",
  IMPORTRANGE: "web",
  GETPIVOTDATA: "lookup",
  FILTER: "lookup",
  SORT: "lookup",
  SORTBY: "lookup",
  UNIQUE: "lookup",

  // Information
  ISBLANK: "information",
  ISERROR: "information",
  ISERR: "information",
  ISNA: "information",
  ISNUMBER: "information",
  ISTEXT: "information",
  ISNONTEXT: "information",
  ISLOGICAL: "information",
  ISREF: "information",
  ISFORMULA: "information",
  ISEVEN: "information",
  ISODD: "information",
  NA: "information",
  TYPE: "information",
  "ERROR.TYPE": "information",
  CELL: "information",
  INFO: "information",
  N: "information",
  SHEET: "information",
  SHEETS: "information",
  BOOKNAME: "information",
  SHEETSNAME: "information",
  "GET.CELL": "information",

  // Names whose stem a prefix rule below would misfile (DBCS is not a DB
  // depreciation function, T.DIST is not text, ODDFPRICE is not ODD, …), plus
  // every function no prefix rule reaches. Excel's own category is the source.
  ACOT: "math",
  ACOTH: "math",
  ARABIC: "math",
  BASE: "math",
  COMBINA: "math",
  COT: "math",
  COTH: "math",
  CSC: "math",
  CSCH: "math",
  DECIMAL: "math",
  "ECMA.CEILING": "math",
  "ISO.CEILING": "math",
  MULTINOMIAL: "math",
  PERCENTOF: "math",
  ROMAN: "math",
  SEC: "math",
  SECH: "math",
  SERIESSUM: "math",
  SQRTPI: "math",
  SUMX2MY2: "math",
  SUMX2PY2: "math",
  SUMXMY2: "math",

  CHIDIST: "statistical",
  CHIINV: "statistical",
  CHITEST: "statistical",
  "COVARIANCE.P": "statistical",
  "COVARIANCE.S": "statistical",
  CRITBINOM: "statistical",
  "F.DIST": "statistical",
  "F.DIST.RT": "statistical",
  "F.INV": "statistical",
  "F.INV.RT": "statistical",
  "F.TEST": "statistical",
  FDIST: "statistical",
  FINV: "statistical",
  "FORECAST.ETS": "statistical",
  "FORECAST.ETS.CONFINT": "statistical",
  "FORECAST.ETS.SEASONALITY": "statistical",
  "FORECAST.ETS.STAT": "statistical",
  "FORECAST.LINEAR": "statistical",
  FTEST: "statistical",
  LOGINV: "statistical",
  "LOGNORM.DIST": "statistical",
  "LOGNORM.INV": "statistical",
  LOGNORMDIST: "statistical",
  PEARSON: "statistical",
  RSQ: "statistical",
  "SKEW.P": "statistical",
  STEYX: "statistical",
  "T.DIST": "statistical",
  "T.DIST.2T": "statistical",
  "T.DIST.RT": "statistical",
  "T.INV": "statistical",
  "T.INV.2T": "statistical",
  "T.TEST": "statistical",
  TDIST: "statistical",
  TINV: "statistical",
  TTEST: "statistical",
  "Z.TEST": "statistical",
  ZTEST: "statistical",

  BYCOL: "logical",
  BYROW: "logical",
  ISOMITTED: "logical",
  LAMBDA: "logical",
  LET: "logical",
  MAKEARRAY: "logical",
  MAP: "logical",
  REDUCE: "logical",
  SCAN: "logical",

  ARRAYTOTEXT: "text",
  ASC: "text",
  BAHTTEXT: "text",
  DBCS: "text",
  JIS: "text",
  NUMBERSTRING: "text",
  PHONETIC: "text",
  REGEXEXTRACT: "text",
  REGEXMATCH: "text",
  REGEXP: "text",
  REGEXREPLACE: "text",
  REGEXTEST: "text",
  WIDECHAR: "text",

  DATESTRING: "dateTime",
  ISOWEEKNUM: "dateTime",

  ADDRESS: "lookup",
  AREAS: "lookup",
  CHOOSECOLS: "lookup",
  CHOOSEROWS: "lookup",
  DROP: "lookup",
  EXPAND: "lookup",
  FORMULATEXT: "lookup",
  GROUPBY: "lookup",
  HSTACK: "lookup",
  PIVOTBY: "lookup",
  RTD: "lookup",
  TAKE: "lookup",
  TOCOL: "lookup",
  TOROW: "lookup",
  TRIMRANGE: "lookup",
  VSTACK: "lookup",
  WRAPCOLS: "lookup",
  WRAPROWS: "lookup",

  AMORDEGRC: "financial",
  AMORLINC: "financial",
  COUPDAYBS: "financial",
  COUPDAYS: "financial",
  COUPDAYSNC: "financial",
  COUPNCD: "financial",
  COUPNUM: "financial",
  COUPPCD: "financial",
  DOLLARDE: "financial",
  DOLLARFR: "financial",
  INTRATE: "financial",
  ISPMT: "financial",
  MDURATION: "financial",
  ODDFPRICE: "financial",
  ODDFYIELD: "financial",
  ODDLPRICE: "financial",
  ODDLYIELD: "financial",
  PDURATION: "financial",
  RMB: "financial",
  RRI: "financial",
  USDOLLAR: "financial",

  COMPLEX: "engineering",

  DAVERAGE: "database",
  DCOUNT: "database",
  DCOUNTA: "database",
  DGET: "database",
  DMAX: "database",
  DMIN: "database",
  DPRODUCT: "database",
  DSTDEV: "database",
  DSTDEVP: "database",
  DSUM: "database",
  DVAR: "database",
  DVARP: "database",
};

// Prefix-based fallbacks (checked in order) for families whose names share a
// stem. Keeps the explicit map small without losing categorization.
const PREFIX_RULES: ReadonlyArray<{
  test: (name: string) => boolean;
  category: FunctionCategory;
}> = [
  { test: (n) => n.startsWith("BIN2") || n.startsWith("DEC2") || n.startsWith("HEX2") || n.startsWith("OCT2"), category: "engineering" },
  { test: (n) => n.startsWith("BIT"), category: "engineering" },
  { test: (n) => n.startsWith("IM"), category: "engineering" },
  { test: (n) => n.startsWith("CONVERT") || n.startsWith("DELTA") || n.startsWith("GESTEP") || n.startsWith("ERF"), category: "engineering" },
  { test: (n) => n.startsWith("BESSEL"), category: "engineering" },
  {
    test: (n) =>
      n.startsWith("DATE") ||
      n.startsWith("TIME") ||
      n.startsWith("DAY") ||
      n.startsWith("MONTH") ||
      n.startsWith("YEAR") ||
      n.startsWith("WEEK") ||
      n.startsWith("HOUR") ||
      n.startsWith("MINUTE") ||
      n.startsWith("SECOND") ||
      n.startsWith("EDATE") ||
      n.startsWith("EOMONTH") ||
      n.startsWith("WORKDAY") ||
      n.startsWith("NETWORKDAYS") ||
      n === "NOW" ||
      n === "TODAY",
    category: "dateTime",
  },
  {
    test: (n) =>
      n.startsWith("SIN") ||
      n.startsWith("COS") ||
      n.startsWith("TAN") ||
      n.startsWith("ASIN") ||
      n.startsWith("ACOS") ||
      n.startsWith("ATAN") ||
      n.startsWith("SINH") ||
      n.startsWith("COSH") ||
      n.startsWith("TANH") ||
      n.startsWith("DEGREES") ||
      n.startsWith("RADIANS") ||
      n.startsWith("ROUND") ||
      n.startsWith("CEILING") ||
      n.startsWith("FLOOR") ||
      n.startsWith("TRUNC") ||
      n.startsWith("INT") ||
      n.startsWith("EVEN") ||
      n.startsWith("ODD") ||
      n.startsWith("MROUND") ||
      n.startsWith("MMULT") ||
      n.startsWith("MINVERSE") ||
      n.startsWith("MDETERM") ||
      n.startsWith("MUNIT"),
    category: "math",
  },
  {
    test: (n) =>
      n.startsWith("LEFT") ||
      n.startsWith("RIGHT") ||
      n.startsWith("MID") ||
      n.startsWith("LEN") ||
      n.startsWith("FIND") ||
      n.startsWith("SEARCH") ||
      n.startsWith("REPLACE") ||
      n.startsWith("SUBSTITUTE") ||
      n.startsWith("CONCAT") ||
      n.startsWith("TEXT") ||
      n.startsWith("UPPER") ||
      n.startsWith("LOWER") ||
      n.startsWith("PROPER") ||
      n.startsWith("TRIM") ||
      n.startsWith("CLEAN") ||
      n.startsWith("REPT") ||
      n.startsWith("CHAR") ||
      n.startsWith("CODE") ||
      n.startsWith("UNICHAR") ||
      n.startsWith("UNICODE") ||
      n.startsWith("EXACT") ||
      n.startsWith("VALUE") ||
      n.startsWith("NUMBERVALUE") ||
      n.startsWith("DOLLAR") ||
      n.startsWith("FIXED") ||
      n === "T",
    category: "text",
  },
  {
    test: (n) =>
      n.startsWith("NORM") ||
      n.startsWith("BETA") ||
      n.startsWith("GAMMA") ||
      n.startsWith("CHISQ") ||
      n.startsWith("BINOM") ||
      n.startsWith("POISSON") ||
      n.startsWith("EXPON") ||
      n.startsWith("WEIBULL") ||
      n.startsWith("HYPGEOM") ||
      n.startsWith("NEGBINOM") ||
      n.startsWith("STDEV") ||
      n.startsWith("VAR") ||
      n.startsWith("CONFIDENCE") ||
      n.startsWith("PERCENTILE") ||
      n.startsWith("QUARTILE") ||
      n.startsWith("RANK") ||
      n.startsWith("MODE") ||
      n.startsWith("PERCENTRANK") ||
      n.startsWith("DEVSQ") ||
      n.startsWith("STANDARDIZE") ||
      n.startsWith("AVEDEV") ||
      n.startsWith("PROB") ||
      n.startsWith("COUNT") ||
      n.startsWith("LINEST") ||
      n.startsWith("LOGEST") ||
      n.startsWith("TREND") ||
      n.startsWith("GROWTH") ||
      n.startsWith("FISHER") ||
      n.startsWith("PHI") ||
      n.startsWith("GAUSS"),
    category: "statistical",
  },
  {
    test: (n) =>
      n.startsWith("PV") ||
      n.startsWith("FV") ||
      n.startsWith("NPV") ||
      n.startsWith("PMT") ||
      n.startsWith("IPMT") ||
      n.startsWith("PPMT") ||
      n.startsWith("RATE") ||
      n.startsWith("IRR") ||
      n.startsWith("MIRR") ||
      n.startsWith("XIRR") ||
      n.startsWith("XNPV") ||
      n.startsWith("NPER") ||
      n.startsWith("SLN") ||
      n.startsWith("SYD") ||
      n.startsWith("DB") ||
      n.startsWith("DDB") ||
      n.startsWith("VDB") ||
      n.startsWith("EFFECT") ||
      n.startsWith("NOMINAL") ||
      n.startsWith("ACCRINT") ||
      n.startsWith("COUPON") ||
      n.startsWith("DISC") ||
      n.startsWith("DURATION") ||
      n.startsWith("PRICE") ||
      n.startsWith("YIELD") ||
      n.startsWith("DOLLARDE") ||
      n.startsWith("DOLLARFR") ||
      n.startsWith("CUMIPMT") ||
      n.startsWith("CUMPRINC") ||
      n.startsWith("INTRATE") ||
      n.startsWith("RECEIVED") ||
      n.startsWith("TBILL"),
    category: "financial",
  },
];

/** The category a single (case-insensitive) function name belongs to. */
export function categoryOfFunction(name: string): FunctionCategory {
  const upper = name.toUpperCase();
  const explicit = EXPLICIT[upper];
  if (explicit) return explicit;
  for (const rule of PREFIX_RULES) {
    if (rule.test(upper)) return rule.category;
  }
  return "other";
}

/**
 * Build the full catalog: a map from category to the SORTED list of engine
 * function names in it. "all" holds every function. Categories with no members
 * are omitted (except "all", always present). Computed from the live
 * `authoringFunctionNames()`: what the engine supports AND Excel for the web
 * can author — a function evaluated for compatibility alone (FORECAST.ETS…)
 * is not offered (live audit B396).
 */
export function buildFunctionCatalog(): Record<FunctionCategory, string[]> {
  const all = authoringFunctionNames().sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
  const catalog = {
    all,
    math: [],
    statistical: [],
    logical: [],
    text: [],
    dateTime: [],
    lookup: [],
    financial: [],
    engineering: [],
    information: [],
    database: [],
    web: [],
    other: [],
  } as Record<FunctionCategory, string[]>;
  for (const name of all) {
    const cat = categoryOfFunction(name);
    catalog[cat].push(name);
  }
  return catalog;
}

/**
 * The categories that actually have members (for the dialog's category list),
 * in `FUNCTION_CATEGORY_ORDER`. "all" is always included.
 */
export function nonEmptyCategories(
  catalog: Record<FunctionCategory, string[]>,
): FunctionCategory[] {
  return FUNCTION_CATEGORY_ORDER.filter(
    (cat) => cat === "all" || catalog[cat].length > 0,
  );
}

/**
 * The function names to list for a category, optionally filtered by a
 * (case-insensitive) substring search the dialog's search box provides. Empty
 * search returns the category's full list. Always sorted.
 */
export function functionsInCategory(
  catalog: Record<FunctionCategory, string[]>,
  category: FunctionCategory,
  search = "",
): string[] {
  const base = catalog[category] ?? [];
  const q = search.trim().toUpperCase();
  if (q === "") return base;
  return base.filter((n) => n.includes(q));
}
