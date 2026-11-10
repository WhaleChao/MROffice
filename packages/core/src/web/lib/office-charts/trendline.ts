/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One trendline fit for every chart this product draws.
 *
 * Docs and the deck drew a trendline in the shared Chart V2 renderer, the sheet
 * in its own adapter, each with its own least-squares loop and its own idea of
 * which kinds exist (Docs and the deck: linear, then moving average; the sheet:
 * linear and moving average). Office's Trendline menu has six: Exponential,
 * Linear, Logarithmic, Polynomial (order 2–6), Power and Moving Average
 * (ECMA-376 `ST_TrendlineType`: exp, linear, log, movingAvg, poly, power). The
 * six are fitted here, once, and both renderers ask this module.
 *
 * What Office does, reproduced:
 * - On a category axis the x of a point is its ORDINAL, starting at 1 — which
 *   is what makes a logarithmic or a power fit possible there at all (ln 0 is
 *   undefined). On a value axis (scatter, bubble) x is the point's own x.
 * - Exponential and power need every y above zero, logarithmic and power every
 *   x above zero. Excel disables the option when the data breaks that; here the
 *   fit answers null and nothing is drawn, rather than a curve through garbage.
 * - A moving average starts at the first point that has a full window behind it.
 */
/**
 * Office's Format Trendline options, each ABSENT when it is Office's default
 * (so a plain linear trendline is still exactly `{ type: "linear" }`):
 * `c:name`, `c:forward` / `c:backward` (forecast — periods on a category axis,
 * x units on a value axis), `c:intercept`, `c:dispEq`, `c:dispRSqr`. Office
 * offers "Set Intercept" for a linear, polynomial or exponential trendline
 * only, and none of the five but the name for a moving average.
 */
export type OfficeTrendlineOptions = Readonly<{
  name?: string;
  forward?: number;
  backward?: number;
  intercept?: number;
  displayEquation?: true;
  displayRSquared?: true;
}>;

export type OfficeTrendline = OfficeTrendlineOptions &
  (
    | Readonly<{ type: "linear" }>
    | Readonly<{ type: "exponential" }>
    | Readonly<{ type: "logarithmic" }>
    | Readonly<{ type: "power" }>
    | Readonly<{ type: "polynomial"; order: number }>
    | Readonly<{ type: "movingAverage"; period: number }>
  );

export type OfficeTrendlineType = OfficeTrendline["type"];

export const OFFICE_TRENDLINE_TYPES = [
  "exponential",
  "linear",
  "logarithmic",
  "polynomial",
  "power",
  "movingAverage",
] as const satisfies readonly OfficeTrendlineType[];

/** The part's spelling of each kind (`ST_TrendlineType`). */
export const OFFICE_TRENDLINE_OOXML_TYPE: Record<OfficeTrendlineType, string> = {
  linear: "linear",
  exponential: "exp",
  logarithmic: "log",
  power: "power",
  polynomial: "poly",
  movingAverage: "movingAvg",
};

/**
 * The page's trendline for one part's `trendlineType` / `order` / `period`
 * and its options, or undefined when the part states a combination this
 * contract does not hold. The twin of `document_charts.trendline_from_part_*`.
 * An option at Office's default (`c:dispEq val="0"`, which Office writes on
 * every trendline it creates) states nothing; an absent one is displayed.
 */
export function officeTrendlineFromPart(
  kind: string | null,
  order: string | null,
  period: string | null,
  options: Readonly<{
    name?: string | null;
    forward?: string | null;
    backward?: string | null;
    intercept?: string | null;
    dispEq?: string | null;
    dispRSqr?: string | null;
    /** The part sits in a package Office 2007 wrote (absent options hidden). */
    office2007?: boolean;
  }> = {},
): OfficeTrendline | undefined {
  const type = (Object.keys(OFFICE_TRENDLINE_OOXML_TYPE) as OfficeTrendlineType[]).find(
    (key) => OFFICE_TRENDLINE_OOXML_TYPE[key] === kind,
  );
  if (!type) return undefined;
  const integer = (raw: string | null) =>
    raw !== null && /^[0-9]{1,3}$/.test(raw) ? Number(raw) : undefined;
  const stated: Record<string, unknown> = { type };
  if (type === "polynomial") {
    if (period !== null) return undefined;
    stated.order = integer(order);
  } else if (type === "movingAverage") {
    if (order !== null) return undefined;
    stated.period = integer(period);
  } else if (order !== null || period !== null) return undefined;
  for (const key of ["forward", "backward", "intercept"] as const) {
    const raw = options[key];
    if (raw === null || raw === undefined) continue;
    const value = raw.trim() === "" ? NaN : Number(raw);
    if (!Number.isFinite(value)) return undefined;
    if (key !== "intercept" && value === 0) continue;
    stated[key] = value;
  }
  for (const [partKey, pageKey] of [
    ["dispEq", "displayEquation"],
    ["dispRSqr", "displayRSquared"],
  ] as const) {
    const raw = options[partKey];
    // Excel reads an ABSENT option as displayed (X-225: C25_trendline states
    // neither and Excel draws "y = x + 1" and "R² = 1"); Office 2007 read it
    // as hidden. The twin of `document_charts.trendline_from_part_element`.
    if (raw === undefined) {
      // A moving average has neither option to display.
      if (type !== "movingAverage" && options.office2007 !== true) stated[pageKey] = true;
      continue;
    }
    // CT_Boolean: an element without `val` reads true.
    if (raw === null || raw === "1" || raw === "true") stated[pageKey] = true;
    else if (raw !== "0" && raw !== "false") return undefined;
  }
  const name = options.name?.trim();
  if (name) stated.name = name;
  return normalizeOfficeTrendline(stated);
}

/** CT_Trendline's children this contract reads or lets ride along. */
const TRENDLINE_PART_CHILDREN = new Set([
  "name",
  "spPr",
  "trendlineType",
  "order",
  "period",
  "forward",
  "backward",
  "intercept",
  "dispRSqr",
  "dispEq",
  "trendlineLbl",
  "extLst",
]);

/** The same reading, from the `c:trendline` element itself. */
export function officeTrendlineFromPartElement(
  element: Element,
  office2007 = false,
): OfficeTrendline | undefined {
  if (element.attributes.length) return undefined;
  const children = new Map<string, Element>();
  for (const child of Array.from(element.children)) {
    if (!TRENDLINE_PART_CHILDREN.has(child.localName)) return undefined;
    if (children.has(child.localName)) return undefined;
    children.set(child.localName, child);
  }
  const malformed = "\u0000";
  const value = (name: string): string | null | undefined => {
    const node = children.get(name);
    if (!node) return undefined;
    if (
      node.children.length ||
      Array.from(node.attributes).some((attribute) => attribute.name !== "val")
    )
      return malformed;
    return node.getAttribute("val");
  };
  const nameNode = children.get("name");
  if (nameNode && (nameNode.attributes.length || nameNode.children.length))
    return undefined;
  return officeTrendlineFromPart(
    value("trendlineType") ?? null,
    value("order") ?? null,
    value("period") ?? null,
    {
      name: nameNode?.textContent ?? null,
      forward: value("forward"),
      backward: value("backward"),
      intercept: value("intercept"),
      dispEq: value("dispEq"),
      dispRSqr: value("dispRSqr"),
      office2007,
    },
  );
}

/** `c:order` is ST_Order, 2–6; `c:period` is ST_Period, 2–255. */
export const OFFICE_TRENDLINE_ORDER = { min: 2, max: 6 } as const;
export const OFFICE_TRENDLINE_PERIOD = { min: 2, max: 255 } as const;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function boundedInteger(
  value: unknown,
  bound: { min: number; max: number },
): number | null {
  return typeof value === "number" &&
    Number.isInteger(value) &&
    value >= bound.min &&
    value <= bound.max
    ? value
    : null;
}

export const OFFICE_TRENDLINE_MAX_FORECAST = 1_000_000;
export const OFFICE_TRENDLINE_MAX_NAME_CHARS = 255;
const OPTION_KEYS = new Set([
  "name",
  "forward",
  "backward",
  "intercept",
  "displayEquation",
  "displayRSquared",
]);

/** The options on `value`, or null when one does not belong to `type`. */
function normalizeOptions(
  type: OfficeTrendlineType,
  value: Record<string, unknown>,
): OfficeTrendlineOptions | null {
  const out: {
    name?: string;
    forward?: number;
    backward?: number;
    intercept?: number;
    displayEquation?: true;
    displayRSquared?: true;
  } = {};
  if ("name" in value) {
    const name = typeof value.name === "string" ? value.name.trim() : "";
    if (!name || name.length > OFFICE_TRENDLINE_MAX_NAME_CHARS) return null;
    out.name = name;
  }
  for (const key of ["forward", "backward"] as const) {
    if (!(key in value)) continue;
    const amount = value[key];
    if (
      typeof amount !== "number" ||
      !Number.isFinite(amount) ||
      amount <= 0 ||
      amount > OFFICE_TRENDLINE_MAX_FORECAST
    )
      return null;
    out[key] = amount;
  }
  if ("intercept" in value) {
    const intercept = value.intercept;
    if (typeof intercept !== "number" || !Number.isFinite(intercept)) return null;
    if (!["linear", "polynomial", "exponential"].includes(type)) return null;
    if (type === "exponential" && intercept <= 0) return null;
    out.intercept = intercept;
  }
  for (const key of ["displayEquation", "displayRSquared"] as const) {
    if (!(key in value)) continue;
    if (value[key] !== true) return null;
    out[key] = true;
  }
  // A moving average takes the name and nothing else.
  if (
    type === "movingAverage" &&
    Object.keys(out).some((key) => key !== "name")
  )
    return null;
  return out;
}

/** The exact shape, or undefined: a stray key is not a trendline. */
export function normalizeOfficeTrendline(
  value: unknown,
): OfficeTrendline | undefined {
  if (!isRecord(value) || typeof value.type !== "string") return undefined;
  const own =
    value.type === "polynomial"
      ? "order"
      : value.type === "movingAverage"
        ? "period"
        : null;
  const keys = Object.keys(value);
  if (keys.some((key) => key !== "type" && key !== own && !OPTION_KEYS.has(key)))
    return undefined;
  switch (value.type) {
    case "linear":
    case "exponential":
    case "logarithmic":
    case "power": {
      const options = normalizeOptions(value.type, value);
      return options ? { type: value.type, ...options } : undefined;
    }
    case "polynomial": {
      const order = boundedInteger(value.order, OFFICE_TRENDLINE_ORDER);
      const options = normalizeOptions("polynomial", value);
      return order !== null && options
        ? { type: "polynomial", order, ...options }
        : undefined;
    }
    case "movingAverage": {
      const period = boundedInteger(value.period, OFFICE_TRENDLINE_PERIOD);
      const options = normalizeOptions("movingAverage", value);
      return period !== null && options
        ? { type: "movingAverage", period, ...options }
        : undefined;
    }
    default:
      return undefined;
  }
}

export interface OfficeTrendlinePoint {
  x: number;
  y: number;
}

function leastSquaresLine(
  xs: readonly number[],
  ys: readonly number[],
): { slope: number; intercept: number } | null {
  const n = xs.length;
  let sx = 0,
    sy = 0,
    sxx = 0,
    sxy = 0;
  for (let i = 0; i < n; i += 1) {
    sx += xs[i];
    sy += ys[i];
    sxx += xs[i] * xs[i];
    sxy += xs[i] * ys[i];
  }
  const denominator = n * sxx - sx * sx;
  if (!Number.isFinite(denominator) || denominator === 0) return null;
  const slope = (n * sxy - sx * sy) / denominator;
  const intercept = (sy - slope * sx) / n;
  return Number.isFinite(slope) && Number.isFinite(intercept)
    ? { slope, intercept }
    : null;
}

/** Least-squares polynomial coefficients, lowest power first; null when singular. */
function polynomialCoefficients(
  xs: readonly number[],
  ys: readonly number[],
  order: number,
): number[] | null {
  const size = order + 1;
  if (xs.length < size) return null;
  // Centre and scale x so the normal equations stay well conditioned at order 6.
  const mean = xs.reduce((sum, x) => sum + x, 0) / xs.length;
  const spread = Math.max(...xs.map((x) => Math.abs(x - mean))) || 1;
  const ts = xs.map((x) => (x - mean) / spread);
  const matrix = Array.from({ length: size }, () => new Array<number>(size + 1).fill(0));
  for (let i = 0; i < ts.length; i += 1) {
    const powers = [1];
    for (let p = 1; p <= 2 * order; p += 1) powers.push(powers[p - 1] * ts[i]);
    for (let row = 0; row < size; row += 1) {
      for (let col = 0; col < size; col += 1) matrix[row][col] += powers[row + col];
      matrix[row][size] += powers[row] * ys[i];
    }
  }
  for (let col = 0; col < size; col += 1) {
    let pivot = col;
    for (let row = col + 1; row < size; row += 1)
      if (Math.abs(matrix[row][col]) > Math.abs(matrix[pivot][col])) pivot = row;
    if (Math.abs(matrix[pivot][col]) < 1e-12) return null;
    [matrix[col], matrix[pivot]] = [matrix[pivot], matrix[col]];
    for (let row = 0; row < size; row += 1) {
      if (row === col) continue;
      const factor = matrix[row][col] / matrix[col][col];
      for (let k = col; k <= size; k += 1) matrix[row][k] -= factor * matrix[col][k];
    }
  }
  const scaled = matrix.map((row, i) => row[size] / row[i]);
  if (scaled.some((value) => !Number.isFinite(value))) return null;
  // Evaluated in the scaled variable; the caller gets a function, not raw terms.
  return [mean, spread, ...scaled];
}

/** Least squares through a fixed intercept: y − b = slope · x. */
function slopeThroughIntercept(
  xs: readonly number[],
  ys: readonly number[],
  intercept: number,
): number | null {
  let sxx = 0,
    sxy = 0;
  for (let i = 0; i < xs.length; i += 1) {
    sxx += xs[i] * xs[i];
    sxy += xs[i] * (ys[i] - intercept);
  }
  const slope = sxy / sxx;
  return sxx > 0 && Number.isFinite(slope) ? slope : null;
}

/**
 * Raw-x polynomial coefficients, lowest power first, optionally with the
 * constant pinned ("Set Intercept"). x is scaled (not centred) so the constant
 * keeps its meaning; this is what the equation label prints and what a pinned
 * fit draws.
 */
function rawPolynomial(
  xs: readonly number[],
  ys: readonly number[],
  order: number,
  intercept: number | undefined,
): number[] | null {
  const first = intercept === undefined ? 0 : 1;
  const size = order + 1 - first;
  if (xs.length < size) return null;
  const scale = Math.max(...xs.map(Math.abs)) || 1;
  const ts = xs.map((x) => x / scale);
  const matrix = Array.from({ length: size }, () => new Array<number>(size + 1).fill(0));
  for (let i = 0; i < ts.length; i += 1) {
    const powers = [1];
    for (let p = 1; p <= 2 * order; p += 1) powers.push(powers[p - 1] * ts[i]);
    const target = ys[i] - (intercept ?? 0);
    for (let row = 0; row < size; row += 1) {
      for (let col = 0; col < size; col += 1)
        matrix[row][col] += powers[row + first + col + first];
      matrix[row][size] += powers[row + first] * target;
    }
  }
  for (let col = 0; col < size; col += 1) {
    let pivot = col;
    for (let row = col + 1; row < size; row += 1)
      if (Math.abs(matrix[row][col]) > Math.abs(matrix[pivot][col])) pivot = row;
    if (Math.abs(matrix[pivot][col]) < 1e-14) return null;
    [matrix[col], matrix[pivot]] = [matrix[pivot], matrix[col]];
    for (let row = 0; row < size; row += 1) {
      if (row === col) continue;
      const factor = matrix[row][col] / matrix[col][col];
      for (let k = col; k <= size; k += 1) matrix[row][k] -= factor * matrix[col][k];
    }
  }
  const scaled = matrix.map((row, i) => row[size] / row[i]);
  const terms = [
    ...(first ? [intercept as number] : []),
    ...scaled,
  ].map((term, power) => (first && power === 0 ? term : term / scale ** power));
  return terms.every(Number.isFinite) ? terms : null;
}

/** What a fitted trendline is: its curve and what its label would say. */
export interface OfficeTrendlineFit {
  curve: (x: number) => number;
  /** Office's equation text, e.g. `y = 2.5x + 1.2`. */
  equation: string;
  /**
   * R² as Office displays it: on the transformed (linearised) model for an
   * exponential or power fit, and against Σ(y − b)² when the intercept is set.
   */
  rSquared: number;
}

const SUPERSCRIPT = ["", "", "²", "³", "⁴", "⁵", "⁶"];

function magnitude(value: number): string {
  const size = Math.abs(value);
  if (size === 0) return "0";
  // General in a label: five significant figures and at most four decimals,
  // whole numbers kept whole; the far ends in scientific with no decimals
  // (PowerPoint's labels: 0.4021x + 7.7338, 13.974x, 68.753, 125374, 8E-05,
  
  if (size >= 1e-4 && size < 1e6) {
    const decimals = Math.max(0, Math.min(4, 4 - Math.floor(Math.log10(size))));
    const fixed = size.toFixed(decimals);
    return (decimals ? fixed.replace(/\.?0+$/, "") : fixed) || "0";
  }
  return size
    .toExponential(0)
    .replace(/e([+-])(\d)$/, "E$10$2")
    .replace("e", "E");
}

function signedTerms(terms: readonly (readonly [number, string])[]): string {
  let out = "";
  for (const [value, suffix] of terms) {
    if (value === 0 && out) continue;
    // A unit coefficient is not written on a variable term: Excel labels the
    // fit of 2, 3, 4, 5 over 1..4 "y = x + 1" (C25_trendline, X-225).
    const size = magnitude(value);
    const body = suffix && size === "1" ? suffix : `${size}${suffix}`;
    out = out
      ? `${out} ${value < 0 ? "-" : "+"} ${body}`
      : `${value < 0 ? "-" : ""}${body}`;
  }
  return out || "0";
}

function determination(
  actual: readonly number[],
  fitted: readonly number[],
  pinned: number | undefined,
): number {
  const mean = actual.reduce((sum, y) => sum + y, 0) / actual.length;
  let residual = 0,
    total = 0;
  for (let i = 0; i < actual.length; i += 1) {
    residual += (actual[i] - fitted[i]) ** 2;
    total += (actual[i] - (pinned ?? mean)) ** 2;
  }
  return total > 0 ? 1 - residual / total : 1;
}

/**
 * Whether a kind can be fitted to these points at all. Office greys
 * Exponential and Power out when a value is zero or negative, and Logarithmic
 * and Power when an x is: the logarithm of either is undefined (PPT-B150:
 * Exponential was offered on 0, 2, 3, 4 and drew nothing, with no reason).
 */
export function officeTrendlineKindFits(
  type: OfficeTrendlineType,
  points: readonly OfficeTrendlinePoint[],
): boolean {
  const usable = points.filter(
    (point) => Number.isFinite(point.x) && Number.isFinite(point.y),
  );
  const ys = type === "exponential" || type === "power";
  const xs = type === "logarithmic" || type === "power";
  return (
    (!ys || usable.every((point) => point.y > 0)) &&
    (!xs || usable.every((point) => point.x > 0))
  );
}

/**
 * The fit with its equation and R², or null when Office would not offer the
 * kind for this data (or for a moving average, which is not a function of x).
 */
export function officeTrendlineFit(
  trendline: OfficeTrendline,
  points: readonly OfficeTrendlinePoint[],
): OfficeTrendlineFit | null {
  const usable = points.filter(
    (point) => Number.isFinite(point.x) && Number.isFinite(point.y),
  );
  if (usable.length < 2) return null;
  const xs = usable.map((point) => point.x);
  const ys = usable.map((point) => point.y);
  const positive = (values: readonly number[]) => values.every((value) => value > 0);
  const pinned = trendline.intercept;
  const line = (
    lineXs: readonly number[],
    lineYs: readonly number[],
    fixed: number | undefined,
  ) => {
    if (fixed === undefined) return leastSquaresLine(lineXs, lineYs);
    const slope = slopeThroughIntercept(lineXs, lineYs, fixed);
    return slope === null ? null : { slope, intercept: fixed };
  };
  switch (trendline.type) {
    case "linear": {
      const fit = line(xs, ys, pinned);
      if (!fit) return null;
      const curve = (x: number) => fit.slope * x + fit.intercept;
      return {
        curve,
        equation: `y = ${signedTerms([
          [fit.slope, "x"],
          [fit.intercept, ""],
        ])}`,
        rSquared: determination(ys, xs.map(curve), pinned),
      };
    }
    case "exponential": {
      if (!positive(ys)) return null;
      const logs = ys.map(Math.log);
      const fit = line(xs, logs, pinned === undefined ? undefined : Math.log(pinned));
      if (!fit) return null;
      const scale = Math.exp(fit.intercept);
      return {
        curve: (x) => scale * Math.exp(fit.slope * x),
        equation: `y = ${signedTerms([[scale, ""]])}e^(${signedTerms([[fit.slope, "x"]])})`,
        rSquared: determination(
          logs,
          xs.map((x) => fit.slope * x + fit.intercept),
          pinned === undefined ? undefined : Math.log(pinned),
        ),
      };
    }
    case "logarithmic": {
      if (!positive(xs)) return null;
      const fit = leastSquaresLine(xs.map(Math.log), ys);
      if (!fit) return null;
      const curve = (x: number) =>
        x > 0 ? fit.slope * Math.log(x) + fit.intercept : NaN;
      return {
        curve,
        equation: `y = ${signedTerms([
          [fit.slope, "ln(x)"],
          [fit.intercept, ""],
        ])}`,
        rSquared: determination(ys, xs.map(curve), undefined),
      };
    }
    case "power": {
      if (!positive(xs) || !positive(ys)) return null;
      const logXs = xs.map(Math.log);
      const logYs = ys.map(Math.log);
      const fit = leastSquaresLine(logXs, logYs);
      if (!fit) return null;
      const scale = Math.exp(fit.intercept);
      return {
        curve: (x) => (x > 0 ? scale * x ** fit.slope : NaN),
        equation: `y = ${signedTerms([[scale, ""]])}x^${signedTerms([[fit.slope, ""]])}`,
        rSquared: determination(
          logYs,
          logXs.map((x) => fit.slope * x + fit.intercept),
          undefined,
        ),
      };
    }
    case "polynomial": {
      const raw = rawPolynomial(xs, ys, trendline.order, pinned);
      if (!raw) return null;
      let curve = (x: number) => raw.reduceRight((sum, term) => sum * x + term, 0);
      if (pinned === undefined) {
        // The centred fit draws the free curve: better conditioned at order 6.
        const fitted = polynomialCoefficients(xs, ys, trendline.order);
        if (!fitted) return null;
        const [mean, spread, ...terms] = fitted;
        curve = (x: number) => {
          const t = (x - mean) / spread;
          return terms.reduceRight((sum, term) => sum * t + term, 0);
        };
      }
      return {
        curve,
        equation: `y = ${signedTerms(
          raw
            .map(
              (term, power) =>
                [term, power === 0 ? "" : `x${SUPERSCRIPT[power] ?? `^${power}`}`] as const,
            )
            .reverse(),
        )}`,
        rSquared: determination(ys, xs.map(curve), pinned),
      };
    }
    case "movingAverage":
      return null;
  }
}

/**
 * The fitted curve as a function of x, or null when Office would not offer the
 * kind for this data. Moving average is not a function of x: see
 * {@link officeTrendlinePolyline}.
 */
export function fitOfficeTrendline(
  trendline: OfficeTrendline,
  points: readonly OfficeTrendlinePoint[],
): ((x: number) => number) | null {
  return officeTrendlineFit(trendline, points)?.curve ?? null;
}

/**
 * The label Office draws beside the trendline — the equation, the R² value, or
 * both on two lines — or null when neither is displayed (or there is no fit).
 */
export function officeTrendlineLabel(
  trendline: OfficeTrendline,
  points: readonly OfficeTrendlinePoint[],
): string | null {
  if (!trendline.displayEquation && !trendline.displayRSquared) return null;
  const fit = officeTrendlineFit(trendline, points);
  if (!fit) return null;
  return [
    trendline.displayEquation ? fit.equation : null,
    // Four decimals in the General look, trailing zeros dropped: Excel's
    // label reads "R² = 0.996" for 0.99601 (chart matrix P09).
    trendline.displayRSquared ? `R² = ${String(Number(fit.rSquared.toFixed(4)))}` : null,
  ]
    .filter(Boolean)
    .join("\n");
}

/**
 * Office's words for a trendline's automatic name, by language: the kind's
 * short name, and a moving average's with its period.
 */
const TRENDLINE_NAME_WORDS: Readonly<
  Record<string, Readonly<Record<Exclude<OfficeTrendlineType, "movingAverage">, string>> & { movingAverage: (period: number) => string }>
> = {
  en: { linear: "Linear", exponential: "Expon.", logarithmic: "Log.", polynomial: "Poly.", power: "Power", movingAverage: (n) => `${n} per. Mov. Avg.` },
  zh: { linear: "线性", exponential: "指数", logarithmic: "对数", polynomial: "多项式", power: "乘幂", movingAverage: (n) => `${n} 期移动平均` },
  "zh-tw": { linear: "線性", exponential: "指數", logarithmic: "對數", polynomial: "多項式", power: "乘冪", movingAverage: (n) => `${n} 期移動平均` },
  ja: { linear: "線形", exponential: "指数", logarithmic: "対数", polynomial: "多項式", power: "累乗", movingAverage: (n) => `${n} 区間移動平均` },
  ko: { linear: "선형", exponential: "지수", logarithmic: "로그", polynomial: "다항식", power: "거듭제곱", movingAverage: (n) => `${n} 구간 이동 평균` },
};

/**
 * The name a trendline goes by in the legend, as Office gives it: its own
 * name (Format Trendline > Trendline Name > Custom, `c:name`), else the
 * automatic one -- its kind and the series it fits, "Linear (Series1)",
 * in the reader's language.
 */
export function officeTrendlineLegendName(
  trendline: OfficeTrendline,
  seriesName: string,
  locale?: string,
): string {
  const own = trendline.name?.trim();
  if (own) return own;
  const tag = (locale ?? "").toLowerCase();
  const words =
    TRENDLINE_NAME_WORDS[tag] ??
    TRENDLINE_NAME_WORDS[tag.startsWith("zh-hant") || tag === "zh-hk" ? "zh-tw" : tag.split("-")[0]] ??
    TRENDLINE_NAME_WORDS.en;
  const kind = trendline.type === "movingAverage" ? words.movingAverage(trendline.period) : words[trendline.type];
  return `${kind} (${seriesName})`;
}

/**
 * Blank category slots a forward forecast adds after the last category. Office
 * extends a category axis by whole periods; a value axis just grows.
 */
export function officeTrendlineForecastCategories(
  trendline: OfficeTrendline | null | undefined,
): number {
  return trendline?.forward ? Math.ceil(trendline.forward) : 0;
}

/**
 * The points to draw. `xs` are where the chart can place a point: every
 * ordinal on a category axis, or (when omitted) an even sampling of the data's
 * x range on a value axis, dense enough for a curve to read as one. A forecast
 * extends the range: by periods past the last ordinal on a category axis (and
 * at most half a period before the first, where Office's axis ends), by x
 * units on a value axis.
 */
export function officeTrendlinePolyline(
  trendline: OfficeTrendline,
  points: readonly OfficeTrendlinePoint[],
  xs?: readonly number[],
): OfficeTrendlinePoint[] {
  const usable = points.filter(
    (point) => Number.isFinite(point.x) && Number.isFinite(point.y),
  );
  if (trendline.type === "movingAverage") {
    const period = trendline.period;
    if (usable.length < period) return [];
    const out: OfficeTrendlinePoint[] = [];
    for (let index = period - 1; index < usable.length; index += 1) {
      let sum = 0;
      for (let back = 0; back < period; back += 1) sum += usable[index - back].y;
      out.push({ x: usable[index].x, y: sum / period });
    }
    return out;
  }
  const curve = fitOfficeTrendline(trendline, usable);
  if (!curve) return [];
  const forward = trendline.forward ?? 0;
  const backward = trendline.backward ?? 0;
  let sample = xs;
  if (sample) {
    if ((forward || backward) && sample.length) {
      const first = Math.min(...sample);
      const last = Math.max(...sample);
      const before = backward ? [first - Math.min(backward, 0.5)] : [];
      const after: number[] = [];
      for (let step = 1; step <= Math.floor(forward); step += 1) after.push(last + step);
      if (forward % 1) after.push(last + forward);
      sample = [...before, ...sample, ...after];
    }
  } else {
    const min = Math.min(...usable.map((point) => point.x)) - backward;
    const max = Math.max(...usable.map((point) => point.x)) + forward;
    if (!(max > min)) return [];
    const steps = trendline.type === "linear" ? 1 : 48;
    sample = Array.from({ length: steps + 1 }, (_, i) => min + ((max - min) * i) / steps);
  }
  return sample
    .map((x) => ({ x, y: curve(x) }))
    .filter((point) => Number.isFinite(point.y));
}
