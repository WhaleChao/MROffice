/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkGfm from "remark-gfm";
import remarkMath from "remark-math";
import remarkDirective from "remark-directive";
import { parseEntities } from "parse-entities";
import type { Root, Content } from "mdast";
import type { PmDoc, PmNode, PmMark } from "./types";
import {
  MAX_CODE_CHART_CAPTION_CHARS,
  MAX_CODE_CHART_SOURCE_CHARS,
  MAX_MERMAID_CAPTION_CHARS,
  MAX_MERMAID_SOURCE_CHARS,
  directiveTypeToPmType,
  isValidCodeChartOutputPng,
} from "./nodeMapping";
import { DirectiveLimitError, readDirectiveBlock } from "./directive";
import { chartNodeAttrsOnV2 } from "../../components/console/editor/extensions/chartBlockShared";
import {
  MAX_DOCUMENT_DEPTH,
  MAX_DOCUMENT_NODES,
  MAX_MARKDOWN_SOURCE_CHARS,
  MAX_MARKDOWN_TABLE_COLUMNS,
  MAX_MARKDOWN_TABLE_ROWS,
  MarkdownDocumentTooLargeError,
  MarkdownUnsupportedNodeError,
  enforcePmDocumentLimits,
  normalizeImageAltText,
  preflightMarkdownSource,
  sanitizeImageSrc,
  sanitizeLinkHref,
  type ImageSourceOptions,
} from "./policy";

export {
  MAX_DOCUMENT_DEPTH,
  MAX_DOCUMENT_NODES,
  MAX_MARKDOWN_SOURCE_CHARS,
  MAX_MARKDOWN_SOURCE_LINES,
  MarkdownDocumentTooLargeError,
  sanitizeImageSrc,
  sanitizeLinkHref,
} from "./policy";
export type { MarkdownDocumentTooLargeReason } from "./policy";

const processor = unified()
  .use(remarkParse)
  .use(remarkGfm)
  .use(remarkMath)
  .use(remarkDirective);

/** Reject a runaway remark tree before constructing the larger PM JSON tree. */
interface ParseBudget {
  sourceChars: number;
  mdastNodes: number;
}

function enforceMdastLimits(root: Root, budget?: ParseBudget): void {
  const stack: Array<{
    node: { type?: string; children?: unknown[] };
    depth: number;
  }> = [
    {
      node: root as unknown as { type?: string; children?: unknown[] },
      depth: 0,
    },
  ];
  let nodes = 0;
  while (stack.length > 0) {
    const current = stack.pop()!;
    nodes += 1;
    if (budget) budget.mdastNodes += 1;
    const observedNodes = budget?.mdastNodes ?? nodes;
    if (observedNodes > MAX_DOCUMENT_NODES) {
      throw new MarkdownDocumentTooLargeError(
        "nodes",
        observedNodes,
        MAX_DOCUMENT_NODES,
      );
    }
    if (current.depth > MAX_DOCUMENT_DEPTH) {
      throw new MarkdownDocumentTooLargeError(
        "depth",
        current.depth,
        MAX_DOCUMENT_DEPTH,
      );
    }

    const children = Array.isArray(current.node.children)
      ? current.node.children
      : [];
    if (current.node.type === "table") {
      if (children.length > MAX_MARKDOWN_TABLE_ROWS) {
        throw new MarkdownDocumentTooLargeError(
          "tableRows",
          children.length,
          MAX_MARKDOWN_TABLE_ROWS,
        );
      }
      for (const row of children) {
        const columns = Array.isArray(
          (row as { children?: unknown[] }).children,
        )
          ? (row as { children: unknown[] }).children.length
          : 0;
        if (columns > MAX_MARKDOWN_TABLE_COLUMNS) {
          throw new MarkdownDocumentTooLargeError(
            "tableColumns",
            columns,
            MAX_MARKDOWN_TABLE_COLUMNS,
          );
        }
      }
    }
    for (let i = children.length - 1; i >= 0; i -= 1) {
      const child = children[i];
      if (typeof child === "object" && child !== null) {
        stack.push({
          node: child as { type?: string; children?: unknown[] },
          depth: current.depth + 1,
        });
      }
    }
  }
}

const FRONTMATTER_RE = /^\uFEFF?---\r?\n([\s\S]*?\r?\n)?---\r?\n?/;
const MAX_TABLE_SPAN = 100;

interface CapturedDirective {
  type: string;
  attrs: Record<string, unknown>;
  body: string;
}

/**
 * Pick an internal marker prefix that provably does not occur in this source.
 * The bounded scans keep adversarial collision attempts linear, while the first
 * missing integer makes the fallback work in browsers without Web Crypto.
 */
function collisionFreeMarkerPrefix(
  source: string,
  namespace: "DIR" | "WIKI",
): string {
  const occupied = new Set<number>();
  // remark decodes one layer of CommonMark character references in text. Scan
  // that canonical view too, otherwise `&#68;` can manufacture the `D` in a
  // directive marker only after the raw-source collision check has passed.
  const collisionViews = source.includes("&")
    ? [source, parseEntities(source, { nonTerminated: false })]
    : [source];
  for (const view of collisionViews) {
    const pattern = new RegExp(`\\x01MRNOTE${namespace}(\\d{1,9})\\x01`, "g");
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(view)) !== null) {
      const candidate = Number.parseInt(match[1], 10);
      if (Number.isSafeInteger(candidate)) occupied.add(candidate);
    }
  }
  let candidate = 0;
  while (occupied.has(candidate)) candidate += 1;
  return `\x01MRNOTE${namespace}${candidate}\x01`;
}

function extractFrontmatter(md: string): { yaml: string | null; rest: string } {
  const m = md.match(FRONTMATTER_RE);
  if (!m) return { yaml: null, rest: md };
  // Strip trailing newline from captured body (if any). For empty body,
  // m[1] is undefined → yaml is "".
  const raw = m[1] ?? "";
  const yaml = raw.replace(/\r?\n$/, "");
  return { yaml, rest: md.slice(m[0].length) };
}

function preExtractDirectives(md: string): {
  stripped: string;
  directives: CapturedDirective[];
  markerPrefix: string;
} {
  // The prefix is selected outside the source's occupied marker namespace, so
  // source text cannot alias a generated placeholder. This is deterministic,
  // collision-safe, and does not require crypto.randomUUID in older browsers.
  const markerPrefix = collisionFreeMarkerPrefix(md, "DIR");
  const lines = md.split("\n");
  const out: string[] = [];
  const captured: CapturedDirective[] = [];
  let codeFence: { marker: "`" | "~"; length: number } | null = null;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const fenceMatch = line.match(/^\s{0,3}(`{3,}|~{3,})/);
    if (codeFence) {
      out.push(line);
      if (
        fenceMatch &&
        fenceMatch[1][0] === codeFence.marker &&
        fenceMatch[1].length >= codeFence.length
      ) {
        codeFence = null;
      }
      continue;
    }
    if (fenceMatch) {
      codeFence = {
        marker: fenceMatch[1][0] as "`" | "~",
        length: fenceMatch[1].length,
      };
      out.push(line);
      continue;
    }
    if (line.startsWith(":::")) {
      const dir = readDirectiveBlock(lines, i);
      if (dir) {
        const idx = captured.length;
        captured.push(dir.block);
        out.push(`${markerPrefix}${idx}END`);
        out.push("");
        i = dir.endIdx;
        continue;
      }
      // Malformed `:::` opener — wrap the offending block in a
      // fenced code block so the user's content survives verbatim.
      // Scan up to 50 lines ahead for a closing `:::`; if none, the
      // single opener line becomes the fenced block on its own.
      let endLine = i;
      for (let j = i + 1; j < Math.min(i + 50, lines.length); j++) {
        if (/^:{3,}$/.test(lines[j].trim())) {
          endLine = j;
          break;
        }
      }
      const block = lines.slice(i, endLine + 1).join("\n");
      // The fallback itself must not be terminable by bytes inside the source
      // being preserved. Pick a backtick fence longer than every run in the
      // malformed block instead of assuming a fixed triple fence is absent.
      let longestBacktickRun = 0;
      for (const match of block.matchAll(/`+/g)) {
        longestBacktickRun = Math.max(longestBacktickRun, match[0].length);
      }
      const fallbackFence = "`".repeat(Math.max(3, longestBacktickRun + 1));
      out.push(fallbackFence);
      out.push(block);
      out.push(fallbackFence);
      i = endLine;
      continue;
    }
    out.push(line);
  }
  return { stripped: out.join("\n"), directives: captured, markerPrefix };
}

/**
 * Promote a ```mermaid fence into the mermaid node (the ecosystem-native
 * markdown form of the block; see serializeMermaid).
 *
 * Returns null — leaving the fence as a lossless plain codeBlock — whenever
 * the payload cannot be represented faithfully in-bounds: source over
 * MAX_MERMAID_SOURCE_CHARS, or a fence info meta that is not exactly this
 * codec's {"caption": "..."} JSON (a foreign meta would be silently dropped
 * by the node's attrs).
 */
function mermaidNodeFromCodeFence(
  node: { value?: string | null; meta?: string | null },
  context: ParseContext,
): PmNode | null {
  const source = restoreSentinelsToRawText(
    node.value ?? "",
    context.wikiMarkers,
  );
  if (source.length > MAX_MERMAID_SOURCE_CHARS || source.includes("\r")) {
    return null;
  }
  const meta = (node.meta ?? "").trim();
  let caption = "";
  if (meta) {
    let parsed: unknown;
    try {
      parsed = JSON.parse(meta);
    } catch {
      return null;
    }
    if (
      typeof parsed !== "object" ||
      parsed === null ||
      Array.isArray(parsed) ||
      Object.keys(parsed).length !== 1 ||
      typeof (parsed as { caption?: unknown }).caption !== "string"
    ) {
      return null;
    }
    caption = restoreSentinelsToRawText(
      (parsed as { caption: string }).caption,
      context.wikiMarkers,
    );
    if (caption.length > MAX_MERMAID_CAPTION_CHARS) return null;
  }
  return {
    type: "mermaid",
    attrs: caption ? { source, caption } : { source },
  };
}

/**
 * codeChart nodes freshly promoted from a fence whose meta declared
 * "output":true — the immediately following standard image node (emitted by
 * this codec's serializer) carries their captured PNG and must be folded
 * back into outputPng. The set is consulted only within the same parse pass
 * and never serialized, so a WeakSet keeps the marker off the node JSON.
 */
const codeChartAwaitingOutput = new WeakSet<PmNode>();

/**
 * Promote a ```python fence carrying the {"mrnCodeChart":true,...} info-meta
 * marker into the codeChart node (代码图表, AI-native charting 第三层 step 3).
 *
 * Returns null — leaving the fence as a lossless plain codeBlock — whenever
 * the payload cannot be represented faithfully in-bounds: source over
 * MAX_CODE_CHART_SOURCE_CHARS, or a meta that is not exactly this codec's
 * marker JSON (allowed keys: mrnCodeChart:true, caption:string ≤300,
 * output:true). Mirrors mermaidNodeFromCodeFence's strictness so a foreign
 * meta is never silently dropped.
 */
function codeChartNodeFromCodeFence(
  node: { value?: string | null; meta?: string | null },
  context: ParseContext,
): PmNode | null {
  const source = restoreSentinelsToRawText(
    node.value ?? "",
    context.wikiMarkers,
  );
  if (source.length > MAX_CODE_CHART_SOURCE_CHARS || source.includes("\r")) {
    return null;
  }
  const meta = (node.meta ?? "").trim();
  if (!meta) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(meta);
  } catch {
    return null;
  }
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    return null;
  }
  const record = parsed as Record<string, unknown>;
  const keys = Object.keys(record);
  if (record.mrnCodeChart !== true) return null;
  const allowed = new Set(["mrnCodeChart", "caption", "output"]);
  if (keys.some((key) => !allowed.has(key))) return null;
  let caption = "";
  if ("caption" in record) {
    if (typeof record.caption !== "string") return null;
    caption = restoreSentinelsToRawText(record.caption, context.wikiMarkers);
    if (caption.length > MAX_CODE_CHART_CAPTION_CHARS) return null;
  }
  if ("output" in record && record.output !== true) return null;
  const result: PmNode = {
    type: "codeChart",
    attrs: caption ? { source, caption } : { source },
  };
  if (record.output === true) codeChartAwaitingOutput.add(result);
  return result;
}

/**
 * Fold the standard image node the serializer emits right after a marked
 * ```python fence back into the preceding codeChart's outputPng. Only a
 * codeChart whose meta declared "output":true consumes an image, and only
 * when that image is an in-bounds PNG data URI — an unrelated adjacent
 * picture (remote URL, JPEG, oversized) is left untouched.
 */
function recombineCodeChartOutputs(nodes: PmNode[]): PmNode[] {
  const result: PmNode[] = [];
  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    result.push(node);
    if (node.type !== "codeChart" || !codeChartAwaitingOutput.has(node)) {
      continue;
    }
    codeChartAwaitingOutput.delete(node);
    const next = nodes[index + 1];
    const src = (next?.attrs as { src?: unknown } | undefined)?.src;
    if (
      next?.type === "image" &&
      typeof src === "string" &&
      src !== "" &&
      isValidCodeChartOutputPng(src)
    ) {
      node.attrs = { ...(node.attrs ?? {}), outputPng: src };
      index += 1; // consume the image node
    }
  }
  return result;
}

function directiveFallbackCodeBlock(dir: CapturedDirective): PmNode {
  const attrs = Object.keys(dir.attrs).length
    ? ` ${JSON.stringify(dir.attrs)}`
    : "";
  const body = dir.body ? `\n${dir.body}` : "";
  return {
    type: "codeBlock",
    attrs: { language: null },
    content: [
      {
        type: "text",
        text: `::: ${dir.type}${attrs}${body}\n:::`,
      },
    ],
  };
}

function rehydrateDirectives(
  nodes: PmNode[],
  directives: CapturedDirective[],
  markerPrefix: string,
  context: ParseContext,
): PmNode[] {
  return nodes.flatMap((n) => {
    if (
      n.type !== "paragraph" ||
      n.content?.length !== 1 ||
      n.content[0].type !== "text"
    ) {
      return [n];
    }
    const text = (n.content[0] as { text?: string }).text ?? "";
    const m = text.trim().match(new RegExp(`^${markerPrefix}(\\d+)END$`));
    if (!m) return [n];
    const dir = directives[parseInt(m[1], 10)];
    if (!dir) return [n];
    const pmType = directiveTypeToPmType(dir.type);
    if (!pmType) {
      // Unknown directive type — preserve original text in a code block.
      return [directiveFallbackCodeBlock(dir)];
    }
    let attrs = { ...dir.attrs } as Record<string, unknown>;
    // A V1 chart directive is read onto the one chart model.
    if (pmType === "chart") attrs = chartNodeAttrsOnV2(attrs) ?? attrs;
    if (pmType === "ai_output" && dir.body) {
      attrs.content_markdown = dir.body;
    }
    if (pmType === "callout") {
      const directiveDepth = (context.directiveDepth ?? 0) + 1;
      if (
        context.enforceLimits !== false &&
        directiveDepth > MAX_DOCUMENT_DEPTH
      ) {
        throw new MarkdownDocumentTooLargeError(
          "depth",
          directiveDepth,
          MAX_DOCUMENT_DEPTH,
        );
      }
      const bodyDoc = parseMarkdownInternal(dir.body, {
        allowRemoteImages: context.allowRemoteImages,
        budget: context.budget,
        definitions: cloneReferenceDefinitions(context.definitions),
        directiveDepth,
        enforceLimits: context.enforceLimits,
        wikiMarkers: createWikiMarkers(""),
      });
      return [
        {
          type: pmType,
          attrs,
          content:
            bodyDoc.content && bodyDoc.content.length > 0
              ? bodyDoc.content
              : [{ type: "paragraph", content: [] }],
        },
      ];
    }
    // Atom directives have no PM child-content slot. A body on any atom other
    // than ai_output cannot be represented faithfully, so preserve the source
    // verbatim rather than silently deleting it.
    if (pmType !== "ai_output" && dir.body) {
      return [directiveFallbackCodeBlock(dir)];
    }
    return [{ type: pmType, attrs }];
  });
}

// Sentinel-based wikiLink preprocessing.
//
// remark does not natively understand [[Page]] / [[Target|Display]],
// so we substitute each match with a sentinel before remark sees the
// source. After remark walks the AST, sentinels appear inside `text`
// nodes, where `expandWikiSentinelsToNodes` splits them back into
// distinct text nodes carrying a `wikiLink` mark.
//
// Each parse chooses a marker prefix absent from its own source. U+0001 keeps
// remark from treating markers as Markdown punctuation, while the collision
// scan—not an assumption that control bytes cannot exist—provides safety.
const WIKI_LINK_RE = /\[\[([^\]|]+?)(?:\|([^\]]+?))?\]\]/g;

interface WikiMarkers {
  open: string;
  separator: string;
  close: string;
  outputPattern: RegExp;
}

function createWikiMarkers(source: string): WikiMarkers {
  const prefix = collisionFreeMarkerPrefix(source, "WIKI");
  const open = `${prefix}OPEN`;
  const separator = `${prefix}SEP`;
  const close = `${prefix}CLOSE`;
  const escape = (value: string) =>
    value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return {
    open,
    separator,
    close,
    outputPattern: new RegExp(
      `${escape(open)}([^\\u0000]+?)(?:${escape(separator)}([^\\u0000]+?))?${escape(close)}`,
      "g",
    ),
  };
}

function preprocessWikiLinks(md: string, markers: WikiMarkers): string {
  return md.replace(
    WIKI_LINK_RE,
    (match, target: string, display: string | undefined, offset: number) => {
      // Respect CommonMark backslash escaping. The regex begins at the first
      // `[`, so an odd run of immediately preceding backslashes means the wiki
      // opener is literal; an even run leaves one literal slash plus a live
      // opener, matching ordinary Markdown delimiter rules.
      let slashes = 0;
      for (
        let index = offset - 1;
        index >= 0 && md[index] === "\\";
        index -= 1
      ) {
        slashes += 1;
      }
      if (slashes % 2 === 1) return match;
      const components = display === undefined ? [target] : [target, display];
      if (
        components.some(
          (value) =>
            value.length === 0 ||
            value.length > 1_000 ||
            /[\u0000-\u001f\u007f-\u009f]|\||\[\[|\]\]/u.test(value),
        )
      ) {
        // Do not manufacture a wiki mark the serializer must later reject.
        // Leaving the original bytes lets remark treat the token as literal
        // bracketed prose without leaking internal sentinels across blocks.
        return match;
      }
      if (display)
        return `${markers.open}${target}${markers.separator}${display}${markers.close}`;
      return `${markers.open}${target}${markers.close}`;
    },
  );
}

const FORMATTED_TEXT_ALIGN_VALUES = new Set([
  "center",
  "right",
  "justify",
  // 分散对齐 (WPS "Distribute") — mirrors the serializer's set so the wrapper
  // round-trips instead of being washed on re-parse.
  "distribute",
]);
const MAX_FORMATTING_INDENT = 7;

/**
 * Restore sentinel-wrapped tokens to their raw markdown form
 * (`[[target]]` or `[[target|display]]`). Used for contexts where a
 * wikiLink mark is not appropriate — e.g. inline code, fenced code,
 * math expressions — but the original text must still survive
 * round-tripping intact.
 */
function restoreSentinelsToRawText(text: string, markers: WikiMarkers): string {
  return text.replace(
    markers.outputPattern,
    (_match, target: string, display?: string) => {
      if (display) return `[[${target}|${display}]]`;
      return `[[${target}]]`;
    },
  );
}

function expandWikiSentinelsToNodes(
  text: string,
  markers: WikiMarkers,
): PmNode[] {
  const nodes: PmNode[] = [];
  let lastIndex = 0;
  text.replace(
    markers.outputPattern,
    (match, target: string, display?: string, offset = 0) => {
      if (offset > lastIndex) {
        nodes.push({ type: "text", text: text.slice(lastIndex, offset) });
      }
      const display_ = display ?? null;
      const visible = display_ ?? target;
      nodes.push({
        type: "text",
        text: visible,
        marks: [
          {
            type: "wikiLink",
            attrs: { target, display: display_, resolved: false },
          },
        ],
      });
      lastIndex = offset + match.length;
      return match;
    },
  );
  if (lastIndex < text.length) {
    nodes.push({ type: "text", text: text.slice(lastIndex) });
  }
  return nodes;
}

const SAFE_INLINE_FORMATTING_TAGS: Readonly<Record<string, string>> = {
  u: "underline",
  strong: "bold",
  em: "italic",
  del: "strike",
};

type SafeInlineHtmlToken =
  | { kind: "open" | "close"; tag: string; markType: string }
  | { kind: "break" };

/**
 * Recognise only the exact attribute-free tags emitted by the serializer.
 * Keeping this grammar deliberately smaller than HTML prevents an event
 * handler, URL-bearing tag, or malformed nesting from crossing the Markdown
 * import boundary under the guise of formatting.
 */
function safeInlineHtmlToken(value: string): SafeInlineHtmlToken | null {
  const trimmed = value.trim();
  if (/^<br\s*\/?>$/i.test(trimmed)) return { kind: "break" };
  const match = trimmed.match(/^<(\/)?(u|strong|em|del)>$/i);
  if (!match) return null;
  const tag = match[2].toLowerCase();
  return {
    kind: match[1] ? "close" : "open",
    tag,
    markType: SAFE_INLINE_FORMATTING_TAGS[tag],
  };
}

function mdastInlineChildrenToPm(
  children: Content[] = [],
  inheritedMarks: PmMark[] = [],
  context: ParseContext = createParseContext(),
): PmNode[] {
  const out: PmNode[] = [];
  const formattingStack: Array<{ tag: string; markType: string }> = [];

  for (const child of children) {
    if (child.type === "html") {
      const token = safeInlineHtmlToken(child.value ?? "");
      if (!token) {
        throw new MarkdownUnsupportedNodeError("html", "inline");
      }
      if (token.kind === "break") {
        out.push({ type: "hardBreak" });
        continue;
      }
      if (token.kind === "open") {
        if (
          inheritedMarks.some((mark) => mark.type === token.markType) ||
          formattingStack.some((entry) => entry.markType === token.markType)
        ) {
          throw new MarkdownUnsupportedNodeError("html", "inline");
        }
        formattingStack.push({ tag: token.tag, markType: token.markType });
        continue;
      }
      const opened = formattingStack.pop();
      if (!opened || opened.tag !== token.tag) {
        throw new MarkdownUnsupportedNodeError("html", "inline");
      }
      continue;
    }
    const marks = [
      ...inheritedMarks,
      ...formattingStack.map(({ markType }) => ({ type: markType })),
    ];
    out.push(...mdastInlineToPm(child, marks, context));
  }

  if (formattingStack.length > 0) {
    throw new MarkdownUnsupportedNodeError("html", "inline");
  }

  return out;
}

interface ReferenceDefinitions {
  links: Map<string, { url: string; title?: string }>;
  footnotes: Map<string, Content[]>;
}

function createReferenceDefinitions(): ReferenceDefinitions {
  return { links: new Map(), footnotes: new Map() };
}

function cloneReferenceDefinitions(
  source: ReferenceDefinitions,
): ReferenceDefinitions {
  return {
    links: new Map(source.links),
    footnotes: new Map(source.footnotes),
  };
}

interface ParseContext extends ImageSourceOptions {
  allowRemoteImages: boolean;
  /** Per top-level parse; nested HTML/directives clone instead of resetting it. */
  definitions: ReferenceDefinitions;
  sourceText?: string;
  /** Shared by recursive directive parses so each body cannot reset limits. */
  budget?: ParseBudget;
  directiveDepth?: number;
  enforceLimits?: boolean;
  /** Parse-local, collision-free markers used only for wiki-link preprocessing. */
  wikiMarkers: WikiMarkers;
}

function createParseContext(): ParseContext {
  return {
    allowRemoteImages: false,
    definitions: createReferenceDefinitions(),
    wikiMarkers: createWikiMarkers(""),
  };
}

export interface ParseOptions extends ImageSourceOptions {
  /**
   * When true, skip the MAX_DOCUMENT_NODES / MAX_DOCUMENT_DEPTH guard.
   * Used only by perf benchmarks that intentionally exercise large
   * synthetic documents. Production callers should leave this at
   * default `false` so the structured error surfaces before the editor
   * tries to mount a runaway tree.
   */
  unsafeSkipSizeLimit?: boolean;
}

function parseMarkdownInternal(md: string, context: ParseContext): PmDoc {
  if (context.enforceLimits !== false) {
    preflightMarkdownSource(md);
    if (context.budget) {
      context.budget.sourceChars += md.length;
      if (context.budget.sourceChars > MAX_MARKDOWN_SOURCE_CHARS) {
        throw new MarkdownDocumentTooLargeError(
          "source",
          context.budget.sourceChars,
          MAX_MARKDOWN_SOURCE_CHARS,
        );
      }
    }
  }
  const { yaml, rest } = extractFrontmatter(md);
  let stripped: string;
  let directives: CapturedDirective[];
  let markerPrefix: string;
  try {
    ({ stripped, directives, markerPrefix } = preExtractDirectives(rest));
  } catch (error) {
    if (error instanceof DirectiveLimitError && error.kind === "body") {
      throw new MarkdownDocumentTooLargeError(
        "directiveBody",
        error.observed,
        error.limit,
      );
    }
    throw error;
  }
  const wikiMarkers = createWikiMarkers(stripped);
  const preprocessed = preprocessWikiLinks(stripped, wikiMarkers);
  if (context.enforceLimits !== false) {
    // Internal placeholders and preservation fences may expand the original.
    // Keep the parser's actual input under the same cap as the caller's source.
    preflightMarkdownSource(preprocessed);
  }
  context.wikiMarkers = wikiMarkers;
  context.sourceText = preprocessed;
  const tree = processor.parse(preprocessed) as Root;
  if (context.enforceLimits !== false) {
    enforceMdastLimits(tree, context.budget);
  }
  collectReferenceDefinitions(tree, context.definitions);
  const content: PmNode[] = [];
  if (yaml !== null) {
    content.push({ type: "frontmatter", attrs: { yaml } });
  }
  for (const child of tree.children) {
    const node = mdastBlockToPm(child, context);
    if (Array.isArray(node)) content.push(...node);
    else if (node) content.push(node);
  }
  const doc: PmDoc = {
    type: "doc",
    content: recombineCodeChartOutputs(
      rehydrateDirectives(content, directives, markerPrefix, context),
    ),
  };
  if (context.enforceLimits !== false) {
    enforcePmDocumentLimits(doc);
  }
  return doc;
}

export function parseMarkdown(md: string, options: ParseOptions = {}): PmDoc {
  const enforceLimits = options.unsafeSkipSizeLimit !== true;
  return parseMarkdownInternal(md, {
    allowRemoteImages: options.allowRemoteImages === true,
    budget: enforceLimits ? { sourceChars: 0, mdastNodes: 0 } : undefined,
    definitions: createReferenceDefinitions(),
    directiveDepth: 0,
    enforceLimits,
    wikiMarkers: createWikiMarkers(""),
  });
}

function splitInlineNodesIntoBlocks(
  inline: PmNode[],
  textBlockType: "paragraph" | "heading" = "paragraph",
  attrs?: Record<string, unknown>,
): PmNode[] {
  const blocks: PmNode[] = [];
  let textRun: PmNode[] = [];
  const flush = () => {
    if (textRun.length === 0) return;
    blocks.push({
      type: textBlockType,
      ...(attrs ? { attrs } : {}),
      content: textRun,
    });
    textRun = [];
  };
  for (const child of inline) {
    if (child.type === "image") {
      flush();
      blocks.push(child);
    } else {
      textRun.push(child);
    }
  }
  flush();
  return blocks;
}

function listItemBlocks(item: Content, context: ParseContext): PmNode[] {
  const content = ((item as { children?: Content[] }).children ?? []).flatMap(
    (block) => {
      const mapped = mdastBlockToPm(block, context);
      return Array.isArray(mapped) ? mapped : mapped ? [mapped] : [];
    },
  );
  // ProseMirror's listItem/taskItem schema requires a leading paragraph. A
  // syntactically empty Markdown marker has no MDAST child, so materialise that
  // required empty paragraph rather than returning an invalid/lossy item.
  return content.length > 0 ? content : [{ type: "paragraph", content: [] }];
}

function mdastBlockToPm(
  node: Content,
  context: ParseContext = createParseContext(),
): PmNode | PmNode[] | null {
  switch (node.type) {
    case "paragraph": {
      const inline = mdastInlineChildrenToPm(node.children ?? [], [], context);
      const blocks = splitInlineNodesIntoBlocks(inline);
      if (blocks.length === 0) return { type: "paragraph", content: [] };
      return blocks.length === 1 ? blocks[0] : blocks;
    }
    case "heading": {
      const attrs = { level: node.depth };
      const inline = mdastInlineChildrenToPm(node.children ?? [], [], context);
      const blocks = splitInlineNodesIntoBlocks(inline, "heading", attrs);
      if (blocks.length === 0) return { type: "heading", attrs, content: [] };
      return blocks.length === 1 ? blocks[0] : blocks;
    }
    case "list": {
      // GFM task list: list items have `checked` boolean property.
      const taskStates = (node.children ?? []).map((li) =>
        li.type === "listItem"
          ? (li as { checked?: boolean | null }).checked
          : null,
      );
      const isTaskList = taskStates.some(
        (checked) => checked !== null && checked !== undefined,
      );
      if (isTaskList) {
        // PM has only a homogeneous unordered taskList. Projecting an ordered
        // or mixed GFM list into it changes both numbering and item semantics.
        if (
          node.ordered ||
          taskStates.some(
            (checked) => checked === null || checked === undefined,
          )
        ) {
          throw new MarkdownUnsupportedNodeError(
            node.ordered ? "orderedTaskList" : "mixedTaskList",
            "block",
          );
        }
        return {
          type: "taskList",
          content: (node.children ?? []).map((li) => ({
            type: "taskItem",
            attrs: { checked: (li as { checked?: boolean }).checked === true },
            content: listItemBlocks(li, context),
          })),
        };
      }
      return {
        type: node.ordered ? "orderedList" : "bulletList",
        ...(node.ordered ? { attrs: { start: node.start ?? 1 } } : {}),
        content: (node.children ?? []).map((li) => ({
          type: "listItem",
          content: listItemBlocks(li, context),
        })),
      };
    }
    case "code": {
      if (node.lang === "mermaid") {
        const mermaidNode = mermaidNodeFromCodeFence(node, context);
        if (mermaidNode) return mermaidNode;
      }
      if (node.lang === "python" && (node.meta ?? "").trim()) {
        const codeChartNode = codeChartNodeFromCodeFence(node, context);
        if (codeChartNode) return codeChartNode;
      }
      return {
        type: "codeBlock",
        attrs: { language: node.lang ?? null },
        content: node.value
          ? [
              {
                type: "text",
                text: restoreSentinelsToRawText(
                  node.value,
                  context.wikiMarkers,
                ),
              },
            ]
          : [],
      };
    }
    case "blockquote":
      return {
        type: "blockquote",
        content: recombineCodeChartOutputs(
          (node.children ?? []).flatMap((b) => {
            const r = mdastBlockToPm(b, context);
            return Array.isArray(r) ? r : r ? [r] : [];
          }),
        ),
      };
    case "thematicBreak":
      return { type: "horizontalRule" };
    case "math":
      return {
        type: "mathBlock",
        attrs: {
          latex: restoreSentinelsToRawText(node.value, context.wikiMarkers),
        },
      };
    case "table": {
      const rows = node.children ?? [];
      const align =
        (node as { align?: Array<"left" | "right" | "center" | null> }).align ??
        [];
      return {
        type: "table",
        content: rows.map((row, rowIdx) => ({
          type: "tableRow",
          content: ((row as { children?: Content[] }).children ?? []).map(
            (cell, columnIdx) => {
              const inline = mdastInlineChildrenToPm(
                (cell as { children?: Content[] }).children ?? [],
                [],
                context,
              );
              const content = splitInlineNodesIntoBlocks(inline);
              const textAlign = align[columnIdx];
              return {
                type: rowIdx === 0 ? "tableHeader" : "tableCell",
                ...(textAlign ? { attrs: { textAlign } } : {}),
                content:
                  content.length > 0
                    ? content
                    : [{ type: "paragraph", content: [] }],
              };
            },
          ),
        })),
      };
    }
    case "html": {
      const rendered =
        parseSafeHtmlTableBlock(node.value ?? "", context) ??
        parseSafeFormattingHtmlBlock(node.value ?? "", context);
      if (rendered) return rendered;
      throw new MarkdownUnsupportedNodeError("html", "block");
    }
    case "definition":
      // The URL is consumed by linkReference/imageReference resolution; the
      // definition line itself renders nothing.
      return null;
    case "footnoteDefinition": {
      // Surface the note body as real blocks so its text is never lost (the
      // inline `[^id]` marker is preserved by the footnoteReference case).
      const blocks: PmNode[] = [];
      for (const child of (node as { children?: Content[] }).children ?? []) {
        const pm = mdastBlockToPm(child, context);
        if (Array.isArray(pm)) blocks.push(...pm);
        else if (pm) blocks.push(pm);
      }
      return blocks.length ? blocks : null;
    }
    default: {
      if (node.type === "containerDirective" || node.type === "leafDirective") {
        const source = sourceTextForNode(node, context);
        if (source !== null) {
          return {
            type: "codeBlock",
            attrs: { language: null },
            content: source ? [{ type: "text", text: source }] : [],
          };
        }
      }
      throw new MarkdownUnsupportedNodeError(node.type, "block");
    }
  }
}

function parseSafeHtmlTableBlock(
  value: string,
  context: ParseContext = createParseContext(),
): PmNode | null {
  const trimmed = value.trim();
  const tableMatch = trimmed.match(/^<table>\s*([\s\S]*?)\s*<\/table>$/i);
  if (!tableMatch) return null;

  const rows: PmNode[] = [];
  const tableBody = tableMatch[1];
  const rowRe = /<tr>\s*([\s\S]*?)\s*<\/tr>/gi;
  let rowMatch: RegExpExecArray | null;
  let rowCursor = 0;
  while ((rowMatch = rowRe.exec(tableBody)) !== null) {
    if (tableBody.slice(rowCursor, rowMatch.index).trim() !== "") return null;
    if (rows.length >= MAX_MARKDOWN_TABLE_ROWS) {
      throw new MarkdownDocumentTooLargeError(
        "tableRows",
        rows.length + 1,
        MAX_MARKDOWN_TABLE_ROWS,
      );
    }
    const cells: PmNode[] = [];
    const rowBody = rowMatch[1];
    const cellRe = /<(td|th)((?:\s+[^>]*)?)>\s*([\s\S]*?)\s*<\/\1>/gi;
    let cellMatch: RegExpExecArray | null;
    let cellCursor = 0;
    while ((cellMatch = cellRe.exec(rowBody)) !== null) {
      if (rowBody.slice(cellCursor, cellMatch.index).trim() !== "") return null;
      if (cells.length >= MAX_MARKDOWN_TABLE_COLUMNS) {
        throw new MarkdownDocumentTooLargeError(
          "tableColumns",
          cells.length + 1,
          MAX_MARKDOWN_TABLE_COLUMNS,
        );
      }
      const attrs = parseTableCellHtmlAttrs(cellMatch[2] ?? "");
      if (attrs === null) return null;
      cells.push({
        type: cellMatch[1].toLowerCase() === "th" ? "tableHeader" : "tableCell",
        ...(Object.keys(attrs).length > 0 ? { attrs } : {}),
        content: [
          {
            type: "paragraph",
            content: parseTableCellHtmlInlineContent(
              cellMatch[3] ?? "",
              context,
            ),
          },
        ],
      });
      cellCursor = cellRe.lastIndex;
    }
    if (cells.length === 0 || rowBody.slice(cellCursor).trim() !== "") {
      return null;
    }
    rows.push({ type: "tableRow", content: cells });
    rowCursor = rowRe.lastIndex;
  }

  if (tableBody.slice(rowCursor).trim() !== "") return null;

  return rows.length > 0 ? { type: "table", content: rows } : null;
}

function parseTableCellHtmlAttrs(
  attrText: string,
): Record<string, number> | null {
  const trimmed = attrText.trim();
  if (!trimmed) return {};

  const attrs: Record<string, number> = {};
  const seen = new Set<string>();
  let rest = trimmed;
  while (rest.length > 0) {
    const match = rest.match(
      /^(colspan|rowspan)\s*=\s*(?:"([^"]*)"|'([^']*)')\s*/i,
    );
    if (!match) return null;

    const name = match[1].toLowerCase();
    if (seen.has(name)) return null;
    seen.add(name);

    const span = parseTableSpan(match[2] ?? match[3] ?? "");
    if (span === null) return null;
    if (span > 1) attrs[name] = span;

    rest = rest.slice(match[0].length).trimStart();
  }

  return attrs;
}

function parseTableSpan(value: string): number | null {
  if (!/^\d+$/.test(value)) return null;
  const parsed = Number.parseInt(value, 10);
  // Clamping changes the table grid while making the import look successful.
  // Reject unsupported geometry so the caller can choose another format.
  if (!Number.isFinite(parsed) || parsed < 1 || parsed > MAX_TABLE_SPAN) {
    return null;
  }
  return parsed;
}

function parseTableCellHtmlInlineContent(
  value: string,
  context: ParseContext = createParseContext(),
): PmNode[] {
  const chunks = value.split(/<br\s*\/?>/i);
  const out: PmNode[] = [];
  chunks.forEach((chunk, index) => {
    if (index > 0) out.push({ type: "hardBreak" });
    out.push(...parseFormattingHtmlInlineContent(chunk, context));
  });
  return out;
}

function parseSafeFormattingHtmlBlock(
  value: string,
  context: ParseContext = createParseContext(),
): PmNode | null {
  const match = value
    .trim()
    .match(/^<(p|h[1-6])((?:\s+[^>]*)?)>([\s\S]*)<\/\1>$/i);
  if (!match) return null;

  const tag = match[1].toLowerCase();
  const attrs = parseFormattingHtmlAttrs(match[2] ?? "");
  if (!attrs) return null;

  const content = parseFormattingHtmlInlineContent(match[3] ?? "", context);
  if (tag === "p") {
    return Object.keys(attrs).length > 0
      ? { type: "paragraph", attrs, content }
      : { type: "paragraph", content };
  }
  const level = Number.parseInt(tag.slice(1), 10);
  return {
    type: "heading",
    attrs: { level, ...attrs },
    content,
  };
}

function parseFormattingHtmlAttrs(
  attrText: string,
): Record<string, unknown> | null {
  const trimmed = attrText.trim();
  if (!trimmed) return {};

  const attrs: Record<string, unknown> = {};
  const seen = new Set<string>();
  let rest = trimmed;
  while (rest.length > 0) {
    const match = rest.match(
      /^([A-Za-z_:][A-Za-z0-9:._-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')\s*/,
    );
    if (!match) return null;

    const name = match[1].toLowerCase();
    const value = match[2] ?? match[3] ?? "";
    if (seen.has(name)) return null;
    seen.add(name);

    if (name === "data-text-align") {
      if (value === "left") {
        rest = rest.slice(match[0].length).trimStart();
        continue;
      }
      if (!FORMATTED_TEXT_ALIGN_VALUES.has(value)) return null;
      attrs.textAlign = value;
    } else if (name === "data-indent") {
      const indent = parseFormattingIndent(value);
      if (indent === null) return null;
      if (indent > 0) attrs.indent = indent;
    } else {
      return null;
    }

    rest = rest.slice(match[0].length).trimStart();
  }

  return attrs;
}

function parseFormattingIndent(value: string): number | null {
  if (!/^\d+$/.test(value)) return null;
  const parsed = Number.parseInt(value, 10);
  if (
    !Number.isFinite(parsed) ||
    parsed < 0 ||
    parsed > MAX_FORMATTING_INDENT
  ) {
    return null;
  }
  return parsed;
}

function parseFormattingHtmlInlineContent(
  value: string,
  context: ParseContext = createParseContext(),
): PmNode[] {
  // `value` can already contain markers inserted by the enclosing parse before
  // remark classified the safe HTML block. Keep that parse-local capability;
  // replacing it here would leak the old marker as visible control text.
  const preprocessed = preprocessWikiLinks(value, context.wikiMarkers);
  const tree = processor.parse(preprocessed) as Root;
  const definitions = cloneReferenceDefinitions(context.definitions);
  const inlineContext: ParseContext = {
    ...context,
    definitions,
    sourceText: preprocessed,
  };
  collectReferenceDefinitions(tree, definitions);
  if (context.enforceLimits !== false) {
    enforceMdastLimits(tree, context.budget);
  }
  if (tree.children.length === 0) return [];
  if (
    tree.children.length !== 1 ||
    (tree.children[0].type !== "paragraph" &&
      tree.children[0].type !== "heading")
  ) {
    throw new MarkdownUnsupportedNodeError("html", "inline");
  }
  const out: PmNode[] = [];
  const child = tree.children[0];
  out.push(...mdastInlineChildrenToPm(child.children ?? [], [], inlineContext));
  return out;
}

// Canonical mark ordering so equivalent inputs (***x*** vs **_x_**) yield the
// same marks-array order and the parse↔serialize round-trip is faithful. The
// order mirrors the serializer's OUTER → INNER nesting (wikiLink/link wrap
// outside style marks, which nest strike > underline > bold > italic, with code
// innermost) so re-parsing the serialized form yields this same order.
const MARK_SORT_ORDER = [
  "wikiLink",
  "link",
  "strike",
  "underline",
  "bold",
  "italic",
  "code",
];

function markSortKey(type: string): number {
  const i = MARK_SORT_ORDER.indexOf(type);
  return i === -1 ? MARK_SORT_ORDER.length : i;
}

function sortMarks(marks: PmMark[]): PmMark[] {
  return [...marks].sort((a, b) => markSortKey(a.type) - markSortKey(b.type));
}

// Reference-style links/images (`[text][id]` + `[id]: url`) and footnotes
// (`[^id]` + `[^id]: …`) are parsed by remark as SEPARATE reference + definition
// nodes; because we call `processor.parse()` without the resolve transform, those
// nodes survive un-inlined. Definitions belong to one ParseContext. Nested
// directive/HTML parses clone that context and can add local definitions without
// resetting the outer document's references.
function collectReferenceDefinitions(
  tree: Root,
  definitions: ReferenceDefinitions,
): void {
  const visit = (node: {
    type?: string;
    identifier?: string;
    url?: string;
    title?: string;
    children?: unknown[];
  }): void => {
    if (node.type === "definition" && typeof node.identifier === "string") {
      // CommonMark resolves duplicate definitions using the first occurrence.
      // Overwriting here would make our reference target disagree with every
      // conforming renderer while still looking like a successful import.
      if (!definitions.links.has(node.identifier)) {
        definitions.links.set(node.identifier, {
          url: node.url ?? "",
          title: node.title ?? undefined,
        });
      }
    } else if (
      node.type === "footnoteDefinition" &&
      typeof node.identifier === "string"
    ) {
      if (!definitions.footnotes.has(node.identifier)) {
        definitions.footnotes.set(
          node.identifier,
          (node.children ?? []) as Content[],
        );
      }
    }
    if (Array.isArray(node.children)) {
      for (const child of node.children) {
        visit(child as Parameters<typeof visit>[0]);
      }
    }
  };
  visit(tree as unknown as Parameters<typeof visit>[0]);
}

// Max chars for a parsed link title (Word ScreenTip). Mirrors
// MAX_LINK_TITLE_CHARS in native-document/linkPlaces without a components import
// (lib/markdown must not depend on the editor layer).
const MAX_PARSED_LINK_TITLE_CHARS = 260;

/**
 * Build a link mark's attrs, attaching the optional CommonMark title (the link
 * mark's modelled `title` attr / Word ScreenTip) when present. The title is
 * sentinel-restored, whitespace-collapsed, trimmed and clamped — mirroring
 * normalizeLinkTitle — so a hostile .md cannot inject a huge or odd tooltip.
 */
function linkAttrsWithTitle(
  href: string,
  rawTitle: unknown,
  markers: WikiMarkers,
): Record<string, unknown> {
  const title =
    typeof rawTitle === "string"
      ? restoreSentinelsToRawText(rawTitle, markers)
          .replace(/\s+/g, " ")
          .trim()
          .slice(0, MAX_PARSED_LINK_TITLE_CHARS)
      : "";
  return title ? { href, title } : { href };
}

function mdastInlineToPm(
  node: Content,
  inheritedMarks: PmMark[] = [],
  context: ParseContext = createParseContext(),
): PmNode[] {
  const addMark = (mark: PmMark) => [...inheritedMarks, mark];
  const keepMarks = () =>
    inheritedMarks.length ? sortMarks(inheritedMarks) : undefined;
  switch (node.type) {
    case "text": {
      const text = node.value;
      const expanded = expandWikiSentinelsToNodes(text, context.wikiMarkers);
      return expanded.map((n) => {
        if (n.type !== "text") return n;
        const wikiMark = (n as { marks?: PmMark[] }).marks ?? [];
        const allMarks = [...inheritedMarks, ...wikiMark];
        return allMarks.length
          ? { ...n, marks: sortMarks(allMarks) }
          : { ...n, marks: undefined };
      });
    }
    case "strong":
      return mdastInlineChildrenToPm(
        node.children ?? [],
        addMark({ type: "bold" }),
        context,
      );
    case "emphasis":
      return mdastInlineChildrenToPm(
        node.children ?? [],
        addMark({ type: "italic" }),
        context,
      );
    case "delete":
      return mdastInlineChildrenToPm(
        node.children ?? [],
        addMark({ type: "strike" }),
        context,
      );
    case "inlineCode":
      return [
        {
          type: "text",
          text: restoreSentinelsToRawText(node.value, context.wikiMarkers),
          marks: sortMarks([...inheritedMarks, { type: "code" }]),
        },
      ];
    case "link": {
      const safeHref = sanitizeLinkHref(node.url ?? "");
      if (safeHref === null) {
        // Disallowed protocol — strip the link mark, keep visible text
        // so user content isn't lost. Prevents `javascript:` /
        // `data:text/html` / `vbscript:` / `file:` URLs from a
        // malicious paste or .md import surviving into PM JSON.
        return mdastInlineChildrenToPm(
          node.children ?? [],
          inheritedMarks,
          context,
        );
      }
      return mdastInlineChildrenToPm(
        node.children ?? [],
        addMark({
          type: "link",
          attrs: linkAttrsWithTitle(
            restoreSentinelsToRawText(safeHref, context.wikiMarkers),
            (node as { title?: string | null }).title,
            context.wikiMarkers,
          ),
        }),
        context,
      );
    }
    case "inlineMath":
      return [
        {
          type: "inlineMath",
          attrs: {
            latex: restoreSentinelsToRawText(node.value, context.wikiMarkers),
          },
        },
      ];
    case "image": {
      const safeSrc = sanitizeImageSrc(node.url ?? "", context);
      if (safeSrc === null) {
        // Never auto-load an untrusted remote URL. Preserve the visible label
        // and, for safe http(s), retain a click-only link so the source is not
        // silently lost while no network request happens on editor mount.
        const safeHref = sanitizeLinkHref(node.url ?? "");
        const label =
          normalizeImageAltText(node.alt ?? "") ||
          (safeHref ? normalizeImageAltText(safeHref) : "");
        if (!label) return [];
        return [
          {
            type: "text",
            text: label,
            marks: safeHref
              ? sortMarks([
                  ...inheritedMarks,
                  { type: "link", attrs: { href: safeHref } },
                ])
              : keepMarks(),
          },
        ];
      }
      return [
        {
          type: "image",
          attrs: { src: safeSrc, alt: normalizeImageAltText(node.alt ?? "") },
        },
      ];
    }
    case "break":
      return [{ type: "hardBreak" }];
    case "linkReference": {
      // `[text][id]` — resolve the URL from the collected definition. Falls back
      // to the visible text (never dropping it) when the ref or scheme is bad.
      const ref = (node as { identifier?: string }).identifier;
      const children = (node as { children?: Content[] }).children ?? [];
      const def = ref ? context.definitions.links.get(ref) : undefined;
      if (def) {
        const safeHref = sanitizeLinkHref(def.url);
        if (safeHref !== null) {
          return mdastInlineChildrenToPm(
            children,
            addMark({
              type: "link",
              attrs: linkAttrsWithTitle(
                restoreSentinelsToRawText(safeHref, context.wikiMarkers),
                def.title,
                context.wikiMarkers,
              ),
            }),
            context,
          );
        }
      }
      return mdastInlineChildrenToPm(children, inheritedMarks, context);
    }
    case "imageReference": {
      // `![alt][id]` — resolve the src; keep the alt text if the ref is missing.
      const ref = (node as { identifier?: string }).identifier;
      const alt = (node as { alt?: string }).alt ?? "";
      const def = ref ? context.definitions.links.get(ref) : undefined;
      if (def) {
        const safeSrc = sanitizeImageSrc(def.url, context);
        if (safeSrc !== null) {
          return [
            {
              type: "image",
              attrs: { src: safeSrc, alt: normalizeImageAltText(alt) },
            },
          ];
        }
      }
      const safeHref = def ? sanitizeLinkHref(def.url) : null;
      const label =
        normalizeImageAltText(alt) ||
        (safeHref ? normalizeImageAltText(safeHref) : "");
      if (!label) return [];
      return [
        {
          type: "text",
          text: label,
          marks: safeHref
            ? sortMarks([
                ...inheritedMarks,
                { type: "link", attrs: { href: safeHref } },
              ])
            : keepMarks(),
        },
      ];
    }
    case "footnoteReference": {
      // `[^id]` — keep the marker inline; the note body is surfaced as blocks by
      // the footnoteDefinition case, so neither the marker nor the text is lost.
      const label =
        (node as { label?: string }).label ??
        (node as { identifier?: string }).identifier ??
        "";
      return [{ type: "text", text: `[^${label}]`, marks: keepMarks() }];
    }
    case "textDirective": {
      // remark-directive also classifies ordinary text such as `14:34` and the
      // `:image` fragment in a malformed data URL as textDirective. Preserve
      // the exact literal slice instead of deleting it or rejecting an otherwise
      // valid document merely because the optional directive grammar claimed it.
      const source = sourceTextForNode(node, context);
      if (source !== null) {
        return [
          {
            type: "text",
            text: restoreSentinelsToRawText(source, context.wikiMarkers),
            marks: keepMarks(),
          },
        ];
      }
      throw new MarkdownUnsupportedNodeError(node.type, "inline");
    }
    default:
      throw new MarkdownUnsupportedNodeError(node.type, "inline");
  }
}

function sourceTextForNode(
  node: Content,
  context: ParseContext,
): string | null {
  const position = (
    node as {
      position?: { start?: { offset?: number }; end?: { offset?: number } };
    }
  ).position;
  const start = position?.start?.offset;
  const end = position?.end?.offset;
  if (
    typeof context.sourceText !== "string" ||
    typeof start !== "number" ||
    typeof end !== "number" ||
    start < 0 ||
    end < start ||
    end > context.sourceText.length
  ) {
    return null;
  }
  return context.sourceText.slice(start, end);
}
