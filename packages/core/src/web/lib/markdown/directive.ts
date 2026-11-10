/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Fenced directive codec for MRNote custom blocks.
 *
 * Wire format:
 *   ::: <type> [json-attrs]
 *   <optional inner markdown body>
 *   :::
 *
 * The single-line header keeps attrs roundtrip-safe (JSON parses
 * deterministically). Body is optional; types like `whiteboard` use
 * an empty body, while `ai-output` carries inline markdown.
 */

import { normalizeDocumentChartHost } from "../../components/console/editor/extensions/documentChartHost";
import { validateChartBlockAttrsV2 } from "../../components/console/editor/extensions/chartBlockShared";
import {
  MAX_CHART_POINTS_PER_SERIES,
  MAX_CHART_SERIES,
  MAX_CHART_TEXT_CHARS,
} from "../office-charts/chart-limits.generated";

export interface DirectiveBlock {
  type: string;
  attrs: Record<string, unknown>;
  body: string;
}

export interface DirectiveHeader {
  type: string;
  attrs: Record<string, unknown>;
}

export const MAX_DIRECTIVE_ATTR_CHARS = 64 * 1024;
export const MAX_DIRECTIVE_BODY_CHARS = 512 * 1024;
export const MAX_DIRECTIVE_BODY_LINES = 10_000;
const MAX_DIRECTIVE_JSON_NODES = 5_000;
const MAX_DIRECTIVE_JSON_DEPTH = 16;
const MAX_DIRECTIVE_STRING_CHARS = 100_000;
const MAX_AI_SOURCES = 200;

export class DirectiveLimitError extends Error {
  readonly kind: "attrs" | "body";
  readonly observed: number;
  readonly limit: number;

  constructor(kind: "attrs" | "body", observed: number, limit: number) {
    super(`Markdown directive ${kind} exceeds limit (${observed} > ${limit})`);
    this.name = "DirectiveLimitError";
    this.kind = kind;
    this.observed = observed;
    this.limit = limit;
  }
}

export class InvalidDirectiveError extends Error {
  readonly directiveType: string;

  constructor(directiveType: string) {
    super(`Invalid attrs for Markdown directive ${directiveType}`);
    this.name = "InvalidDirectiveError";
    this.directiveType = directiveType;
  }
}

const DIRECTIVE_OPEN_RE = /^(:{3,})\s*([A-Za-z][\w-]*)\s*(\{.*\})?\s*$/;
const DIRECTIVE_TYPE_RE = /^[A-Za-z][\w-]*$/;

interface DirectiveHeaderWithFence extends DirectiveHeader {
  fence: string;
}

function fenceForBody(body: string): string {
  let maxFence = 2;
  for (const line of body.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (/^:{3,}$/.test(trimmed)) {
      maxFence = Math.max(maxFence, trimmed.length);
    }
  }
  return ":".repeat(Math.max(3, maxFence + 1));
}

/**
 * Serialize a directive block back to its fenced markdown form.
 * Returns a multi-line string ending with `:::` (no trailing newline).
 * Invalid type tokens, schema-mismatched attrs, and over-budget payloads are
 * rejected before any Markdown is emitted.
 */
export function encodeDirective(block: DirectiveBlock): string {
  if (!DIRECTIVE_TYPE_RE.test(block.type)) {
    throw new InvalidDirectiveError(block.type);
  }
  const attrs = validateDirectiveAttrs(block.type, block.attrs);
  if (attrs === null) throw new InvalidDirectiveError(block.type);
  if (block.body.length > MAX_DIRECTIVE_BODY_CHARS) {
    throw new DirectiveLimitError(
      "body",
      block.body.length,
      MAX_DIRECTIVE_BODY_CHARS,
    );
  }
  const bodyLines = block.body ? block.body.split(/\r\n|\r|\n/).length : 0;
  if (bodyLines > MAX_DIRECTIVE_BODY_LINES) {
    throw new DirectiveLimitError("body", bodyLines, MAX_DIRECTIVE_BODY_LINES);
  }
  const attrPart =
    Object.keys(attrs).length === 0 ? "" : ` ${JSON.stringify(attrs)}`;
  if (attrPart.length > MAX_DIRECTIVE_ATTR_CHARS) {
    throw new DirectiveLimitError(
      "attrs",
      attrPart.length,
      MAX_DIRECTIVE_ATTR_CHARS,
    );
  }
  const fence = fenceForBody(block.body);
  const header = `${fence} ${block.type}${attrPart}`;
  if (header.length > MAX_DIRECTIVE_ATTR_CHARS) {
    throw new DirectiveLimitError(
      "attrs",
      header.length,
      MAX_DIRECTIVE_ATTR_CHARS,
    );
  }
  const bodyPart = block.body ? `\n${block.body}` : "";
  return `${header}${bodyPart}\n${fence}`;
}

/**
 * Parse a single line as a directive opening header.
 * Returns null when the line is not a directive (or attrs JSON is
 * malformed) — caller should fall back to treating it as plain text.
 *
 * The attrs match is greedy (`\{.*\}`), so any text after the closing `}` becomes
 * part of the JSON; if that JSON does not parse, the function returns null.
 */
function decodeDirectiveLineWithFence(
  line: string,
): DirectiveHeaderWithFence | null {
  if (line.length > MAX_DIRECTIVE_ATTR_CHARS) {
    throw new DirectiveLimitError(
      "attrs",
      line.length,
      MAX_DIRECTIVE_ATTR_CHARS,
    );
  }
  const match = line.match(DIRECTIVE_OPEN_RE);
  if (!match) return null;
  const type = match[2].trim();
  if (!type) return null;
  const fence = match[1];
  if (!match[3]) {
    const attrs = validateDirectiveAttrs(type, {});
    return attrs === null ? null : { type, attrs, fence };
  }
  try {
    const attrs = JSON.parse(match[3]);
    const validated = validateDirectiveAttrs(type, attrs);
    if (validated === null) return null;
    return { type, attrs: validated, fence };
  } catch {
    return null;
  }
}

export function decodeDirectiveLine(line: string): DirectiveHeader | null {
  const header = decodeDirectiveLineWithFence(line);
  if (!header) return null;
  return { type: header.type, attrs: header.attrs };
}

/**
 * Find the first directive block starting at lines[startIdx]. Returns
 * the parsed block plus the index of the closing `:::` line. Returns
 * null if startIdx is not a directive opening or the block is unclosed.
 *
 * The closing fence must match the opening fence length exactly. The encoder
 * chooses a longer fence when the body contains `:::` lines, so body content
 * round-trips without truncation.
 */
export function readDirectiveBlock(
  lines: string[],
  startIdx: number,
): { block: DirectiveBlock; endIdx: number } | null {
  const header = decodeDirectiveLineWithFence(lines[startIdx] ?? "");
  if (!header) return null;
  let bodyChars = 0;
  for (let i = startIdx + 1; i < lines.length; i++) {
    if (lines[i].trim() === header.fence) {
      const body = lines.slice(startIdx + 1, i).join("\n");
      return {
        block: { type: header.type, attrs: header.attrs, body },
        endIdx: i,
      };
    }
    bodyChars += lines[i].length + (i > startIdx + 1 ? 1 : 0);
    const bodyLines = i - startIdx;
    if (bodyChars > MAX_DIRECTIVE_BODY_CHARS) {
      throw new DirectiveLimitError(
        "body",
        bodyChars,
        MAX_DIRECTIVE_BODY_CHARS,
      );
    }
    if (bodyLines > MAX_DIRECTIVE_BODY_LINES) {
      throw new DirectiveLimitError(
        "body",
        bodyLines,
        MAX_DIRECTIVE_BODY_LINES,
      );
    }
  }
  return null;
}

type JsonObject = Record<string, unknown>;

function isPlainObject(value: unknown): value is JsonObject {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

/** Validate that an attrs tree is bounded JSON, not a cyclic/runtime object. */
function isBoundedJson(value: unknown): boolean {
  const stack: Array<{ value: unknown; depth: number }> = [{ value, depth: 0 }];
  const seen = new Set<object>();
  let nodes = 0;
  while (stack.length > 0) {
    const current = stack.pop()!;
    nodes += 1;
    if (
      nodes > MAX_DIRECTIVE_JSON_NODES ||
      current.depth > MAX_DIRECTIVE_JSON_DEPTH
    ) {
      return false;
    }
    const item = current.value;
    if (item === null || typeof item === "boolean") continue;
    if (typeof item === "string") {
      if (item.length > MAX_DIRECTIVE_STRING_CHARS) return false;
      continue;
    }
    if (typeof item === "number") {
      // JSON has no integer-width marker. Values outside JS's safe-integer
      // range are already rounded by JSON.parse, so accepting them would make
      // a directive appear lossless after its opaque data changed.
      if (
        !Number.isFinite(item) ||
        (Number.isInteger(item) && !Number.isSafeInteger(item))
      ) {
        return false;
      }
      continue;
    }
    if (typeof item !== "object") return false;
    if (seen.has(item)) return false;
    seen.add(item);
    if (Array.isArray(item)) {
      for (const child of item)
        stack.push({ value: child, depth: current.depth + 1 });
      continue;
    }
    if (!isPlainObject(item)) return false;
    for (const [key, child] of Object.entries(item)) {
      if (key === "__proto__" || key === "prototype" || key === "constructor") {
        return false;
      }
      stack.push({ value: child, depth: current.depth + 1 });
    }
  }
  return true;
}

function hasOnlyKeys(attrs: JsonObject, allowed: ReadonlySet<string>): boolean {
  return Object.keys(attrs).every((key) => allowed.has(key));
}

function optionalString(value: unknown): boolean {
  return (
    value === undefined ||
    (typeof value === "string" && value.length <= MAX_DIRECTIVE_STRING_CHARS)
  );
}

function optionalNullableString(value: unknown): boolean {
  return value === undefined || value === null || optionalString(value);
}

function optionalBoolean(value: unknown): boolean {
  return value === undefined || typeof value === "boolean";
}

function validateAiSources(value: unknown): boolean {
  if (value === undefined) return true;
  if (!Array.isArray(value) || value.length > MAX_AI_SOURCES) return false;
  return value.every(
    (source) =>
      isPlainObject(source) &&
      hasOnlyKeys(source, new Set(["type", "id", "title"])) &&
      typeof source.type === "string" &&
      typeof source.id === "string" &&
      typeof source.title === "string" &&
      source.type.length <= 200 &&
      source.id.length <= 2_000 &&
      source.title.length <= 10_000,
  );
}

// 图表 bounds are the one chart limits table (office_chart_limits, generated
// into chart-limits.generated.ts): the server reads a V1 chart directive onto
// the one model (`normalize_document_chart_attrs`), which takes a chart up to
// those bounds, never the old 50 categories / 10 series. 智能图形 bounds
// mirror the server codec (apps/api/app/services/markdown_directives.py).
const MAX_CHART_DIRECTIVE_CATEGORIES = MAX_CHART_POINTS_PER_SERIES;
const MAX_CHART_DIRECTIVE_SERIES = MAX_CHART_SERIES;
const MAX_SMART_ART_DIRECTIVE_ITEMS = 100;
const MAX_CHART_DIRECTIVE_TEXT_CHARS = MAX_CHART_TEXT_CHARS;
const CHART_DIRECTIVE_TYPES = new Set([
  "bar",
  "line",
  "pie",
  "barh",
  "area",
  "scatter",
  "radar",
  "combo",
  "stock",
]);
const CHART_DIRECTIVE_STYLE_PRESETS = new Set([
  "classic",
  "emerald-gradient",
  "warm",
  "mono",
  "vivid",
]);
// 股价图 contract (mirrors resolve_stock_series_roles server-side): exactly
// four series whose localized names map onto distinct OHLC roles.
const STOCK_DIRECTIVE_ROLE_ALIASES: ReadonlyArray<readonly string[]> = [
  ["开盘", "open"],
  ["最高", "high"],
  ["最低", "low"],
  ["收盘", "close"],
];

function stockDirectiveContractHolds(series: unknown): boolean {
  if (!Array.isArray(series) || series.length !== 4) return false;
  const seen = new Set<number>();
  for (const entry of series) {
    if (!isPlainObject(entry)) return false;
    const name = String(entry.name ?? "")
      .trim()
      .toLowerCase();
    const role = STOCK_DIRECTIVE_ROLE_ALIASES.findIndex((aliases) =>
      aliases.includes(name),
    );
    if (role < 0 || seen.has(role)) return false;
    seen.add(role);
  }
  return seen.size === 4;
}

// 智能图形 layout whitelist — the catalog ids from

// imported) so the markdown engine never pulls the layout engine; mirror the
// server whitelist exactly.
const SMART_ART_DIRECTIVE_LAYOUTS = new Set([
  "basic-list",
  "vertical-box-list",
  "grouped-list",
  "card-grid",
  "step-list",
  "basic-process",
  "chevron-process",
  "vertical-process",
  "alternating-process",
  "timeline",
  "step-up-process",
  "basic-cycle",
  "continuous-cycle",
  "radial-cycle",
  "segmented-cycle",
  "org-chart",
  "horizontal-hierarchy",
  "labeled-tree",
  "half-circle-org",
  "opposing-arrows",
  "balance",
  "funnel",
  "venn",
  "target",
  "radial",
  "gears",
  "basic-matrix",
  "titled-matrix",
  "grid-matrix",
  "cycle-matrix",
  "basic-pyramid",
  "inverted-pyramid",
  "segmented-pyramid",
  "pyramid-list",
  "picture-caption-list",
  "alternating-picture-blocks",
  "circular-picture-callout",
  "picture-accent-process",
  
  "list",
  "process",
  "cycle",
  "hierarchy",
  "matrix",
  "pyramid",
]);

function isChartText(value: unknown): boolean {
  return (
    typeof value === "string" && value.length <= MAX_CHART_DIRECTIVE_TEXT_CHARS
  );
}

function optionalChartText(value: unknown): boolean {
  return value === undefined || isChartText(value);
}

const SCHEMA_ENTRIES: Record<string, (attrs: JsonObject) => boolean> = {
  callout: (attrs) =>
    hasOnlyKeys(attrs, new Set(["variant"])) &&
    (attrs.variant === undefined ||
      (typeof attrs.variant === "string" &&
        new Set(["info", "warning", "success", "danger", "tip"]).has(
          attrs.variant,
        ))),
  "ai-output": (attrs) =>
    hasOnlyKeys(
      attrs,
      new Set(["action_type", "action_log_id", "model_id", "sources"]),
    ) &&
    optionalString(attrs.action_type) &&
    optionalString(attrs.action_log_id) &&
    optionalNullableString(attrs.model_id) &&
    validateAiSources(attrs.sources),
  flashcard: (attrs) =>
    hasOnlyKeys(attrs, new Set(["front", "back", "flipped", "card_id"])) &&
    optionalString(attrs.front) &&
    optionalString(attrs.back) &&
    optionalBoolean(attrs.flipped) &&
    optionalNullableString(attrs.card_id),
  file: (attrs) =>
    hasOnlyKeys(
      attrs,
      new Set([
        "attachment_id",
        "filename",
        "mime_type",
        "size_bytes",
        "editable_imported",
        "editable_page_id",
        "editable_page_type",
      ]),
    ) &&
    optionalString(attrs.attachment_id) &&
    optionalString(attrs.filename) &&
    optionalString(attrs.mime_type) &&
    (attrs.size_bytes === undefined ||
      (typeof attrs.size_bytes === "number" &&
        Number.isSafeInteger(attrs.size_bytes) &&
        attrs.size_bytes >= 0)) &&
    optionalBoolean(attrs.editable_imported) &&
    optionalString(attrs.editable_page_id) &&
    optionalString(attrs.editable_page_type),
  reference: (attrs) =>
    hasOnlyKeys(
      attrs,
      new Set(["target_type", "target_id", "title", "snippet"]),
    ) &&
    optionalString(attrs.target_type) &&
    optionalString(attrs.target_id) &&
    optionalString(attrs.title) &&
    optionalString(attrs.snippet),
  task: (attrs) =>
    hasOnlyKeys(
      attrs,
      new Set([
        "block_id",
        "title",
        "description",
        "due_date",
        "completed",
        "completed_at",
      ]),
    ) &&
    optionalString(attrs.block_id) &&
    optionalString(attrs.title) &&
    optionalNullableString(attrs.description) &&
    optionalNullableString(attrs.due_date) &&
    optionalBoolean(attrs.completed) &&
    optionalNullableString(attrs.completed_at),
  chart: (attrs) => attrs.schemaVersion===2
    ? (() => {
      const keys=new Set(["schemaVersion","chartType","chartSubtype","title","axisTitleX","axisTitleY","data","stylePreset","format","documentChart","textAlign","categories","series"]);
      if (!hasOnlyKeys(attrs,keys)) return false;
      if (attrs.categories!=null || attrs.series!=null) return false;
      // Paragraph alignment is the Docs host's, like its size (server twin:
      // `normalize_document_chart_attrs`).
      if (attrs.textAlign!==undefined && !["left","center","right"].includes(String(attrs.textAlign))) return false;
      try { normalizeDocumentChartHost(attrs.documentChart); } catch { return false; }
      const core={...attrs}; delete core.documentChart;delete core.textAlign;delete core.categories;delete core.series;
      return validateChartBlockAttrsV2(core).valid;
    })()
    :
    hasOnlyKeys(
      attrs,
      new Set(["chartType", "title", "categories", "series", "stylePreset"]),
    ) &&
    (attrs.chartType === undefined ||
      (typeof attrs.chartType === "string" &&
        CHART_DIRECTIVE_TYPES.has(attrs.chartType))) &&
    (attrs.stylePreset === undefined ||
      (typeof attrs.stylePreset === "string" &&
        CHART_DIRECTIVE_STYLE_PRESETS.has(attrs.stylePreset))) &&
    (attrs.chartType !== "stock" ||
      stockDirectiveContractHolds(attrs.series)) &&
    optionalChartText(attrs.title) &&
    (attrs.categories === undefined ||
      (Array.isArray(attrs.categories) &&
        attrs.categories.length <= MAX_CHART_DIRECTIVE_CATEGORIES &&
        attrs.categories.every((item) => isChartText(item)))) &&
    (attrs.series === undefined ||
      (Array.isArray(attrs.series) &&
        attrs.series.length <= MAX_CHART_DIRECTIVE_SERIES &&
        attrs.series.every(
          (entry) =>
            isPlainObject(entry) &&
            hasOnlyKeys(entry, new Set(["name", "values"])) &&
            optionalChartText(entry.name) &&
            (entry.values === undefined ||
              (Array.isArray(entry.values) &&
                entry.values.length <= MAX_CHART_DIRECTIVE_CATEGORIES &&
                entry.values.every(
                  (value) =>
                    typeof value === "number" && Number.isFinite(value),
                ))),
        ))),
  "smart-art": (attrs) =>
    hasOnlyKeys(attrs, new Set(["layout", "items"])) &&
    (attrs.layout === undefined ||
      (typeof attrs.layout === "string" &&
        SMART_ART_DIRECTIVE_LAYOUTS.has(attrs.layout))) &&
    (attrs.items === undefined ||
      (Array.isArray(attrs.items) &&
        attrs.items.length <= MAX_SMART_ART_DIRECTIVE_ITEMS &&
        attrs.items.every((item) => isChartText(item)))),
  whiteboard: (attrs) =>
    hasOnlyKeys(attrs, new Set(["elements", "appState", "width", "height", "sceneKind"])) &&
    (attrs.sceneKind === undefined || ["canvas", "shapes", "flowchart", "mindmap"].includes(String(attrs.sceneKind))) &&
    (attrs.elements === undefined ||
      (Array.isArray(attrs.elements) && attrs.elements.length <= 2_000)) &&
    (attrs.appState === undefined || isPlainObject(attrs.appState)) &&
    (attrs.width === undefined ||
      (typeof attrs.width === "number" &&
        Number.isFinite(attrs.width) &&
        attrs.width > 0 &&
        attrs.width <= 20_000)) &&
    (attrs.height === undefined ||
      (typeof attrs.height === "number" &&
        Number.isFinite(attrs.height) &&
        attrs.height > 0 &&
        attrs.height <= 20_000)),
};

const DIRECTIVE_SCHEMAS = Object.freeze(
  Object.assign(
    Object.create(null) as Record<string, (attrs: JsonObject) => boolean>,
    SCHEMA_ENTRIES,
  ),
);

/**
 * Validate attrs against the PM node's actual schema. Unknown directive types
 * remain syntactically decodable so the parser can preserve them verbatim as a
 * code block instead of deleting them.
 */
export function validateDirectiveAttrs(
  type: string,
  value: unknown,
): Record<string, unknown> | null {
  if (!isPlainObject(value) || !isBoundedJson(value)) return null;
  const schema = Object.prototype.hasOwnProperty.call(DIRECTIVE_SCHEMAS, type)
    ? DIRECTIVE_SCHEMAS[type]
    : undefined;
  if (schema && !schema(value)) return null;
  let encoded: string;
  try {
    encoded = JSON.stringify(value);
  } catch {
    return null;
  }
  if (encoded.length > MAX_DIRECTIVE_ATTR_CHARS) {
    throw new DirectiveLimitError(
      "attrs",
      encoded.length,
      MAX_DIRECTIVE_ATTR_CHARS,
    );
  }
  return value;
}
