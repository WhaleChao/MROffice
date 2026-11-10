/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Syntax highlighting for the in-browser code editor.
 *
 * lowlight (the highlight.js AST, already a dependency for document code
 * blocks) tokenizes the source into a hast tree. We flatten that tree to a
 * list of leaf segments, each carrying the innermost token class, and map the
 * token class to an inline colour. This keeps highlighting framework-agnostic
 * (no global stylesheet, no `dangerouslySetInnerHTML`) and unit-testable: the
 * segment list is a plain serialisable structure.
 */

import { common, createLowlight } from "lowlight";

const lowlight = createLowlight(common);

export interface HighlightSegment {
  text: string;
  /** Space-joined highlight.js class list of the innermost token, if any. */
  className?: string;
}

interface HastText {
  type: "text";
  value: string;
}
interface HastElement {
  type: "element";
  properties?: { className?: string[] | string };
  children: HastNode[];
}
type HastNode = HastText | HastElement | { type: string };

function classNameOf(node: HastElement): string | undefined {
  const cn = node.properties?.className;
  if (Array.isArray(cn)) return cn.join(" ");
  if (typeof cn === "string") return cn;
  return undefined;
}

function flatten(
  node: HastNode,
  inherited: string | undefined,
  out: HighlightSegment[],
): void {
  if (node.type === "text") {
    const value = (node as HastText).value;
    if (value) out.push(inherited ? { text: value, className: inherited } : { text: value });
    return;
  }
  if (node.type === "element") {
    const el = node as HastElement;
    const own = classNameOf(el) ?? inherited;
    for (const child of el.children) flatten(child, own, out);
  }
}

/**
 * Tokenize `code` for `language` into highlight segments. Falls back to a
 * single uncoloured segment for unknown languages or on any failure, and is
 * lossless: the concatenated segment text always equals the input.
 */
export function highlightSegments(
  language: string,
  code: string,
): HighlightSegment[] {
  if (!code) return [];
  let tree: { children: HastNode[] } | null = null;
  try {
    if (lowlight.registered(language)) {
      tree = lowlight.highlight(language, code) as unknown as {
        children: HastNode[];
      };
    }
  } catch {
    tree = null;
  }
  if (!tree) return [{ text: code }];
  const out: HighlightSegment[] = [];
  for (const child of tree.children) flatten(child, undefined, out);
  // Safety net: never drop or reorder characters relative to the source.
  if (out.map((s) => s.text).join("") !== code) return [{ text: code }];
  return out;
}

/**
 * GitHub-light token palette, keyed by highlight.js class (minus the `hljs-`
 * prefix). Matches the editor's light surface; unknown tokens render in the
 * default foreground.
 */
const TOKEN_COLOR: Record<string, string> = {
  keyword: "#cf222e",
  "variable language_": "#cf222e",
  built_in: "#0550ae",
  type: "#0550ae",
  literal: "#0550ae",
  number: "#0550ae",
  operator: "#0550ae",
  attr: "#0550ae",
  attribute: "#0550ae",
  property: "#0550ae",
  selector_attr: "#0550ae",
  selector_pseudo: "#0550ae",
  string: "#0a3069",
  regexp: "#0a3069",
  "selector-attr": "#0a3069",
  symbol: "#0a3069",
  "char.escape_": "#0a3069",
  comment: "#6e7781",
  quote: "#6e7781",
  doctag: "#6e7781",
  meta: "#6e7781",
  "meta keyword": "#cf222e",
  title: "#8250df",
  "title function_": "#8250df",
  "title class_": "#8250df",
  "title function_ invoke__": "#8250df",
  function: "#8250df",
  section: "#0550ae",
  name: "#116329",
  selector_tag: "#116329",
  "selector-tag": "#116329",
  tag: "#116329",
  bullet: "#0550ae",
  variable: "#e36209",
  "template-variable": "#e36209",
  params: "#24292f",
  emphasis: "#24292f",
  strong: "#24292f",
  deletion: "#82071e",
  addition: "#116329",
};

/** The colour roles of the palette above, for a stylesheet to paint. */
export type TokenRole =
  | "keyword"
  | "constant"
  | "string"
  | "comment"
  | "function"
  | "tag"
  | "variable"
  | "plain"
  | "deletion";

const TOKEN_ROLE_BY_COLOR: Record<string, TokenRole> = {
  "#cf222e": "keyword",
  "#0550ae": "constant",
  "#0a3069": "string",
  "#6e7781": "comment",
  "#8250df": "function",
  "#116329": "tag",
  "#e36209": "variable",
  "#24292f": "plain",
  "#82071e": "deletion",
};

/**
 * Resolve a highlight.js class list to its colour role (the CSS paints it:
 * CodeEditorChrome.module.css `[data-token]`), if recognised.
 */
export function tokenRole(className: string | undefined): TokenRole | undefined {
  const color = tokenColor(className);
  return color ? TOKEN_ROLE_BY_COLOR[color] : undefined;
}

/** Resolve a highlight.js class list to an inline colour, if recognised. */
export function tokenColor(className: string | undefined): string | undefined {
  if (!className) return undefined;
  if (className in TOKEN_COLOR) return TOKEN_COLOR[className];
  for (const raw of className.split(/\s+/)) {
    const key = raw.replace(/^hljs-/, "");
    if (key in TOKEN_COLOR) return TOKEN_COLOR[key];
  }
  return undefined;
}
