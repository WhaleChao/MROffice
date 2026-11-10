/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A focused Emmet abbreviation expander for the code editor (VSCode parity).
 *
 * No Emmet npm package is installed, so this implements the common subset that
 * covers day-to-day use, framework-free and fully unit-tested:
 *
 *  - **Markup (HTML / JSX)** — child `>`, sibling `+`, multiply `*N`, grouping
 *    `()`, class `.foo`, id `#bar`, attributes `[href=#]`, and literal text
 *    `{hello}`. Implicit tag names: a bare `.cls` / `#id` becomes `div`, except
 *    inside a `ul`/`ol` (→ `li`), `table`/`tbody`/`thead` (→ `tr`), `tr` (→
 *    `td`), and `select`/`optgroup` (→ `option`). Numbering `$` in a `*N` repeat
 *    is substituted with the 1-based item index.
 *  - **CSS** — property shorthands (`m10` → `margin: 10px;`, `p5` → `padding:
 *    5px;`, `w100` → `width: 100px;`, `df` → `display: flex;`, …). A trailing
 *    number gets the `px` unit unless the property is unitless.
 *
 * Everything here is a PURE string transform: {@link expandEmmet} takes an
 * abbreviation + a syntax ("markup" | "css") and returns the expanded snippet,
 * or `null` when the input is not a valid/known abbreviation (so the caller can
 * fall through to a normal Tab/indent). The editor decides the syntax from the
 * active language and only calls this when the token left of the caret looks
 * like an abbreviation.
 */

/** The two expansion dialects the editor drives. */
export type EmmetSyntax = "markup" | "css";

/** Self-closing (void) HTML elements — emitted as `<tag />` with no close. */
const VOID_ELEMENTS = new Set([
  "area",
  "base",
  "br",
  "col",
  "embed",
  "hr",
  "img",
  "input",
  "link",
  "meta",
  "param",
  "source",
  "track",
  "wbr",
]);

/**
 * The implicit tag a bare `.class` / `#id` abbreviation resolves to, given the
 * PARENT tag. Mirrors Emmet's context-aware defaults; falls back to `div`.
 */
function implicitTag(parentTag: string | null): string {
  switch (parentTag) {
    case "ul":
    case "ol":
      return "li";
    case "table":
    case "tbody":
    case "thead":
    case "tfoot":
      return "tr";
    case "tr":
      return "td";
    case "select":
    case "optgroup":
      return "option";
    case "em":
      return "span";
    default:
      return "div";
  }
}

/** A parsed markup node: a tag plus its attributes / text / children / repeat. */
interface MarkupNode {
  /** Explicit tag, or null for an implicit (`.cls`-only) node resolved later. */
  tag: string | null;
  id?: string;
  classes: string[];
  /** Raw attribute pairs from `[a=b c]`. */
  attrs: { name: string; value: string | null }[];
  /** Literal text from `{...}`. */
  text?: string;
  /** Repeat count from `*N` (1 when absent). */
  repeat: number;
  children: MarkupNode[];
}

/** True for the small set of identifier chars Emmet tag/class/id tokens use. */
function isNameChar(ch: string): boolean {
  return /[A-Za-z0-9_-]/.test(ch);
}

/**
 * Name chars for a class / id / attribute-value token, which additionally allow
 * the `$` numbering placeholder (substituted with the item index at serialize
 * time). Tag names use {@link isNameChar} (no `$`).
 */
function isClassIdChar(ch: string): boolean {
  return /[A-Za-z0-9_\-$]/.test(ch);
}

/**
 * Tokenize + parse a markup abbreviation into a node tree. Returns null on any
 * structural error (unbalanced group/attr/text, empty token where one is
 * required) so the caller falls through rather than inserting garbage.
 *
 * Grammar (informal):
 *   group   := item (('>' | '+') item)*
 *   item    := '(' group ')' repeat? | element repeat?
 *   element := tag? ('#' id | '.' class | '[' attrs ']' | '{' text '}')*
 */
function parseMarkup(abbr: string): MarkupNode[] | null {
  let i = 0;
  const src = abbr;

  function fail(): null {
    return null;
  }

  // Parse a single element (no combinators). Returns the node or null.
  function parseElement(): MarkupNode | null {
    const node: MarkupNode = {
      tag: null,
      classes: [],
      attrs: [],
      repeat: 1,
      children: [],
    };
    let sawAnything = false;

    // Optional leading tag name.
    let tag = "";
    while (i < src.length && isNameChar(src[i])) {
      tag += src[i];
      i += 1;
    }
    if (tag) {
      node.tag = tag;
      sawAnything = true;
    }

    // Zero or more class/id/attr/text segments.
    for (;;) {
      const ch = src[i];
      if (ch === "#") {
        i += 1;
        let id = "";
        while (i < src.length && isClassIdChar(src[i])) {
          id += src[i];
          i += 1;
        }
        if (!id) return fail();
        node.id = id;
        sawAnything = true;
      } else if (ch === ".") {
        i += 1;
        let cls = "";
        while (i < src.length && isClassIdChar(src[i])) {
          cls += src[i];
          i += 1;
        }
        if (!cls) return fail();
        node.classes.push(cls);
        sawAnything = true;
      } else if (ch === "[") {
        i += 1;
        let raw = "";
        while (i < src.length && src[i] !== "]") {
          raw += src[i];
          i += 1;
        }
        if (src[i] !== "]") return fail();
        i += 1; // consume ']'
        const parsed = parseAttrs(raw);
        if (parsed === null) return fail();
        node.attrs.push(...parsed);
        sawAnything = true;
      } else if (ch === "{") {
        i += 1;
        let text = "";
        // Text runs to the matching '}'; no nesting (Emmet text is flat here).
        while (i < src.length && src[i] !== "}") {
          text += src[i];
          i += 1;
        }
        if (src[i] !== "}") return fail();
        i += 1; // consume '}'
        node.text = text;
        sawAnything = true;
      } else {
        break;
      }
    }

    if (!sawAnything) return fail();
    return node;
  }

  // Parse a repeat suffix `*N` (default 1). Returns the count or null on error.
  function parseRepeat(): number | null {
    if (src[i] !== "*") return 1;
    i += 1;
    let digits = "";
    while (i < src.length && /[0-9]/.test(src[i])) {
      digits += src[i];
      i += 1;
    }
    if (!digits) return null;
    const n = Number(digits);
    // Cap the repeat so a fat-fingered `li*9999` can't blow up the buffer.
    if (n < 1 || n > 1000) return null;
    return n;
  }

  // Parse an item: a parenthesized group (becomes a synthetic node holding the
  // group's siblings as children-less repeated block) or a single element. The
  // returned node's `children` may carry a sibling-chain from a group.
  function parseItem(): MarkupNode | null {
    if (src[i] === "(") {
      i += 1;
      const inner = parseGroup();
      if (inner === null) return fail();
      if (src[i] !== ")") return fail();
      i += 1; // consume ')'
      const repeat = parseRepeat();
      if (repeat === null) return fail();
      // Wrap the group's siblings in a transparent container so a `*N` on the
      // group repeats the whole sequence. A null tag + null id/class + no text
      // marks it transparent; the serializer flattens it.
      return {
        tag: null,
        classes: [],
        attrs: [],
        repeat,
        children: inner,
        text: undefined,
        // Mark as a group wrapper via a private flag on the object.
        ...( { __group: true } as object),
      } as MarkupNode;
    }
    const el = parseElement();
    if (el === null) return fail();
    const repeat = parseRepeat();
    if (repeat === null) return fail();
    el.repeat = repeat;
    return el;
  }

  // Parse a group: item (('>'|'+') item)*. A '>' nests the next item under the
  // previous one; a '+' makes it a sibling at the current level.
  function parseGroup(): MarkupNode[] | null {
    const first = parseItem();
    if (first === null) return fail();
    const siblings: MarkupNode[] = [first];
    let current = first;
    for (;;) {
      const op = src[i];
      if (op === ">") {
        i += 1;
        const child = parseItem();
        if (child === null) return fail();
        (current.children ||= []).push(child);
        current = child;
      } else if (op === "+") {
        i += 1;
        const sib = parseItem();
        if (sib === null) return fail();
        siblings.push(sib);
        current = sib;
      } else {
        break;
      }
    }
    return siblings;
  }

  const tree = parseGroup();
  if (tree === null) return null;
  if (i !== src.length) return null; // trailing garbage -> not a valid abbr
  return tree;
}

/** Parse a `[name=value name2="v2" bool]` attribute list. */
function parseAttrs(raw: string): { name: string; value: string | null }[] | null {
  const out: { name: string; value: string | null }[] = [];
  const parts = raw.trim().split(/\s+/).filter(Boolean);
  for (const part of parts) {
    const eq = part.indexOf("=");
    if (eq === -1) {
      if (!/^[A-Za-z_:][\w:.-]*$/.test(part)) return null;
      out.push({ name: part, value: null });
    } else {
      const name = part.slice(0, eq);
      let value = part.slice(eq + 1);
      if (!/^[A-Za-z_:][\w:.-]*$/.test(name)) return null;
      // Strip optional surrounding quotes.
      if (
        (value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'"))
      ) {
        value = value.slice(1, -1);
      }
      out.push({ name, value });
    }
  }
  return out;
}

/** Substitute Emmet `$` numbering tokens with the 1-based item index. */
function substituteNumbering(s: string, index: number): string {
  // Each run of `$` becomes the index zero-padded to that run's length.
  return s.replace(/\$+/g, (run) => String(index).padStart(run.length, "0"));
}

/** Serialize one node (and its children) to markup at the given indent depth. */
function serializeNode(
  node: MarkupNode,
  parentTag: string | null,
  depth: number,
  indentUnit: string,
): string {
  const isGroup = (node as unknown as { __group?: boolean }).__group === true;
  const pieces: string[] = [];
  for (let rep = 1; rep <= node.repeat; rep += 1) {
    if (isGroup) {
      // A transparent group: emit its children as siblings at THIS depth.
      for (const child of node.children) {
        pieces.push(serializeNode(child, parentTag, depth, indentUnit));
      }
      continue;
    }
    const tag = node.tag ?? implicitTag(parentTag);
    const pad = indentUnit.repeat(depth);

    // Build the attribute string (id, class, then explicit attrs), with $-numbering.
    const attrStrs: string[] = [];
    if (node.id) attrStrs.push(`id="${substituteNumbering(node.id, rep)}"`);
    if (node.classes.length > 0) {
      const cls = node.classes.map((c) => substituteNumbering(c, rep)).join(" ");
      attrStrs.push(`class="${cls}"`);
    }
    for (const a of node.attrs) {
      if (a.value === null) {
        attrStrs.push(a.name);
      } else {
        attrStrs.push(`${a.name}="${substituteNumbering(a.value, rep)}"`);
      }
    }
    const attrPart = attrStrs.length > 0 ? ` ${attrStrs.join(" ")}` : "";
    const text = node.text !== undefined ? substituteNumbering(node.text, rep) : "";

    if (VOID_ELEMENTS.has(tag)) {
      pieces.push(`${pad}<${tag}${attrPart} />`);
      continue;
    }

    const open = `<${tag}${attrPart}>`;
    const close = `</${tag}>`;

    if (node.children.length === 0) {
      // Leaf: open + text + close on one line.
      pieces.push(`${pad}${open}${text}${close}`);
    } else {
      // Has children: open on its own line, children indented, close aligned.
      const inner = node.children
        .map((c) => serializeNode(c, tag, depth + 1, indentUnit))
        .join("\n");
      const textLine = text ? `${indentUnit.repeat(depth + 1)}${text}\n` : "";
      pieces.push(`${pad}${open}\n${textLine}${inner}\n${pad}${close}`);
    }
  }
  return pieces.join("\n");
}

/** Serialize a sibling list of nodes at a depth. */
function serializeTree(
  nodes: MarkupNode[],
  parentTag: string | null,
  depth: number,
  indentUnit: string,
): string {
  return nodes
    .map((n) => serializeNode(n, parentTag, depth, indentUnit))
    .join("\n");
}

/** Expand a markup abbreviation; null when it doesn't parse. */
function expandMarkup(abbr: string, indentUnit: string): string | null {
  const tree = parseMarkup(abbr);
  if (tree === null) return null;
  return serializeTree(tree, null, 0, indentUnit);
}

// --- CSS expansion ---------------------------------------------------------

/** Map of CSS property prefixes -> full property name. Longest-prefix wins. */
const CSS_PROPERTIES: { prefix: string; prop: string; defaultValue?: string }[] = [
  // Order longest-first within a shared leading letter so e.g. `mar` doesn't
  // shadow `margin`; the lookup walks this list and takes the first whose
  // prefix the token starts with (after a longest-first sort below).
  { prefix: "margin", prop: "margin" },
  { prefix: "padding", prop: "padding" },
  { prefix: "width", prop: "width" },
  { prefix: "height", prop: "height" },
  { prefix: "top", prop: "top" },
  { prefix: "left", prop: "left" },
  { prefix: "right", prop: "right" },
  { prefix: "bottom", prop: "bottom" },
  { prefix: "m", prop: "margin" },
  { prefix: "p", prop: "padding" },
  { prefix: "w", prop: "width" },
  { prefix: "h", prop: "height" },
  { prefix: "t", prop: "top" },
  { prefix: "l", prop: "left" },
  { prefix: "r", prop: "right" },
  { prefix: "b", prop: "bottom" },
  { prefix: "fz", prop: "font-size" },
  { prefix: "fw", prop: "font-weight" },
  { prefix: "z", prop: "z-index", defaultValue: "" }, // unitless
  { prefix: "op", prop: "opacity", defaultValue: "" }, // unitless
  { prefix: "lh", prop: "line-height", defaultValue: "" }, // unitless
];

/** Properties whose numeric value is unitless (no `px` appended). */
const UNITLESS = new Set(["z-index", "opacity", "line-height", "font-weight"]);

/** Static CSS keyword shorthands: `df` -> `display: flex;`. */
const CSS_KEYWORDS: Record<string, string> = {
  df: "display: flex;",
  db: "display: block;",
  dib: "display: inline-block;",
  dn: "display: none;",
  dg: "display: grid;",
  posr: "position: relative;",
  posa: "position: absolute;",
  posf: "position: fixed;",
  poss: "position: sticky;",
  tac: "text-align: center;",
  tal: "text-align: left;",
  tar: "text-align: right;",
  fll: "float: left;",
  flr: "float: right;",
  curp: "cursor: pointer;",
  ovh: "overflow: hidden;",
  ova: "overflow: auto;",
  bdb: "border-bottom: 1px solid;",
  bdt: "border-top: 1px solid;",
  bxz: "box-sizing: border-box;",
};

/**
 * Expand a CSS abbreviation:
 *  - exact keyword shorthands (`df`, `posr`, …)
 *  - `<prop-prefix><number>` → `prop: <number>px;` (or unitless), e.g. `m10`,
 *    `w100`, `fz14`, `z5`, `op0`. A negative number is supported (`m-10`).
 * Returns null when nothing matches.
 */
function expandCss(abbr: string): string | null {
  const token = abbr.trim();
  if (!token) return null;

  // Static keyword first (exact match).
  if (Object.prototype.hasOwnProperty.call(CSS_KEYWORDS, token)) {
    return CSS_KEYWORDS[token];
  }

  // <letters><optional number>. Letters identify the property; the number is
  // the value. With no number, only treat as a property when a default exists
  // (we don't here), so require a numeric value for the generic path.
  const match = /^([a-z]+)(-?\d+(?:\.\d+)?)$/.exec(token);
  if (!match) return null;
  const [, letters, numStr] = match;

  // Longest-prefix lookup: sort the table longest-first, take the first whose
  // prefix exactly equals the letters (we require exact, not startsWith, so
  // `mx` doesn't silently map to `margin`).
  const sorted = [...CSS_PROPERTIES].sort((a, b) => b.prefix.length - a.prefix.length);
  const hit = sorted.find((e) => e.prefix === letters);
  if (!hit) return null;

  const unit = UNITLESS.has(hit.prop) ? "" : "px";
  return `${hit.prop}: ${numStr}${unit};`;
}

/** True when `s` has a space/tab that is NOT inside a `[..]` or `{..}` span. */
function hasTopLevelWhitespace(s: string): boolean {
  let depth = 0;
  for (let i = 0; i < s.length; i += 1) {
    const ch = s[i];
    if (ch === "[" || ch === "{") depth += 1;
    else if (ch === "]" || ch === "}") depth = Math.max(0, depth - 1);
    else if ((ch === " " || ch === "\t") && depth === 0) return true;
  }
  return false;
}

/**
 * Expand an Emmet abbreviation for the given syntax. Returns the expanded
 * snippet, or `null` when `abbr` is not a recognised abbreviation (the caller
 * then falls through to a literal Tab / indent).
 *
 * `indentUnit` is the editor's active indent string (e.g. `"  "` or `"\t"`),
 * used to indent nested markup so the result matches the buffer's style.
 */
export function expandEmmet(
  abbr: string,
  syntax: EmmetSyntax,
  indentUnit = "  ",
): string | null {
  const trimmed = abbr.trim();
  if (!trimmed) return null;
  // Reject whitespace, but ONLY outside attribute/text brackets — a `[a=b c]`
  // attr list or `{some text}` legitimately contains spaces, whereas a space at
  // the top level means this isn't a single abbreviation (fall through to Tab).
  if (hasTopLevelWhitespace(trimmed)) return null;
  if (syntax === "css") return expandCss(trimmed);
  // Reject a bare number for markup (it would expand to nothing meaningful).
  if (/^\d+$/.test(trimmed)) return null;
  return expandMarkup(trimmed, indentUnit);
}

/**
 * The offset within an expanded snippet where the caret should land — a
 * lightweight stand-in for Emmet's tab-stops. We pick the content position of
 * the FIRST empty element (the index right after the first `>` that is
 * immediately followed by `</`), so `ul>li` leaves the caret inside the `<li>`.
 * Falls back to `null` (caller puts the caret at the snippet end) when there is
 * no empty element — e.g. a self-closed `<img />` or a CSS declaration.
 */
export function firstEmptyTagCaret(snippet: string): number | null {
  // Match `>` directly followed by `</` (an empty element's content slot).
  const idx = snippet.indexOf("></");
  if (idx === -1) return null;
  return idx + 1; // position just after the '>'
}

/**
 * Extract the Emmet abbreviation token immediately to the LEFT of the caret on
 * `line` (caret at 0-based `col`). Returns the token + its start column, or
 * null when there's nothing abbreviation-like there.
 *
 * The token is the maximal run of abbreviation characters ending at the caret.
 * We allow the full markup/CSS abbreviation alphabet (letters, digits, and the
 * structural operators `>+*().#[]{}=:-_"'$` plus space INSIDE brackets). To keep
 * it simple and safe, the scan walks left while the char is part of an
 * abbreviation, but STOPS at a bracket-depth-aware space so a normal sentence
 * before the caret isn't swallowed.
 */
export function emmetAbbreviationLeftOf(
  line: string,
  col: number,
): { abbr: string; start: number } | null {
  const end = Math.max(0, Math.min(col, line.length));
  if (end === 0) return null;
  // Characters that may legally appear in an abbreviation OUTSIDE brackets.
  const tokenChar = /[A-Za-z0-9>+*().#[\]{}=:\-_$!]/;
  let start = end;
  let bracketDepth = 0;
  while (start > 0) {
    const ch = line[start - 1];
    if (ch === "]" || ch === "}") bracketDepth += 1;
    else if (ch === "[" || ch === "{") bracketDepth = Math.max(0, bracketDepth - 1);
    if (ch === " " || ch === "\t") {
      // A space is only allowed inside an attribute/text bracket span.
      if (bracketDepth > 0) {
        start -= 1;
        continue;
      }
      break;
    }
    if (bracketDepth > 0) {
      // Inside brackets, accept any char (attr values, quotes, etc.).
      start -= 1;
      continue;
    }
    if (!tokenChar.test(ch)) break;
    start -= 1;
  }
  const abbr = line.slice(start, end);
  if (!abbr) return null;
  return { abbr, start };
}
