/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The one door every KaTeX rendering in the app goes through.
 *
 * AI answers write Chinese words, full-width punctuation and dashes straight
 * into math mode (`$总成本 = C_f + C_v$`, `$a — b$`). KaTeX draws them (CJK
 * in its fallback font, upright), but under its default `strict: "warn"`
 * each one is a console warning (「LaTeX-incompatible input and strict mode
 * is set to 'warn': Unicode text character "对" used in math mode」,
 * 「Unrecognized Unicode character "—" (8212)」), dozens per answer; and a
 * character LaTeX would only accept in text (a dash, a curly quote) is drawn
 * without metrics, in the wrong box.
 *
 * - {@link prepareLatexForKatex} puts such text where LaTeX would: each run
 *   of CJK characters, full-width punctuation, dashes and curly quotes at
 *   math level becomes `\text{…}` (KaTeX's text mode has metrics for all of
 *   them: no missing glyph, never a math-italic face), and the few symbols
 *   with a math spelling (′ ″ ℃ ￥ ￡) take it. Text already in a text or
 *   font command's argument (`\text{}`, `\mathrm{}`, `\tag{}`) is left as
 *   written. The stored source is never changed; this is display only.
 * - {@link katexStrict} ignores the codes that only say "LaTeX would not
 *   accept this" (KaTeX draws them all the same) and still warns about any
 *   other.
 */

import katex from "katex";

/** KaTeX strict-mode codes that only say "LaTeX would not accept this"; it draws them all the same. */
export const KATEX_DRAWABLE_STRICT_CODES: ReadonlySet<string> = new Set([
  "unknownSymbol",
  "unicodeTextInMathMode",
  "htmlExtension",
  "mathVsTextAccents",
  "mathVsTextUnits",
  "newLineInDisplayMode",
  "commentAtEnd",
]);

/** The `strict` option of a passive rendering (an answer, a note, a preview). */
export function katexStrict(errorCode: string): "ignore" | "warn" {
  return KATEX_DRAWABLE_STRICT_CODES.has(errorCode) ? "ignore" : "warn";
}

/** A command whose next `{…}` argument is text mode: nothing in it is moved. */
const TEXT_ARGUMENT_COMMANDS: ReadonlySet<string> = new Set([
  "text", "textrm", "textsf", "texttt", "textnormal", "textbf", "textit", "textmd",
  "textup", "textsl", "textsc", "emph", "mbox", "hbox", "tag", "label", "href", "url",
]);

/** A command whose next `{…}` argument sets a math font: what is written there keeps it. */
const FONT_ARGUMENT_COMMANDS: ReadonlySet<string> = new Set([
  "mathrm", "mathbf", "mathit", "mathsf", "mathtt", "mathnormal", "mathbb", "mathcal",
  "mathfrak", "mathscr", "boldsymbol", "bm", "pmb", "operatorname",
]);

/**
 * Characters KaTeX's text mode has metrics for and its math mode has not:
 * CJK punctuation and kana (U+3000-30FF), the CJK ideographs (U+4E00-9FAF),
 * full-width ASCII forms (U+FF00-FF60), Hangul syllables (U+AC00-D7AF), and
 * the dashes and curly quotes of running text.
 */
function isTextModeCodePoint(cp: number): boolean {
  return (
    (cp >= 0x3000 && cp <= 0x30ff) ||
    (cp >= 0x4e00 && cp <= 0x9faf) ||
    (cp >= 0xff00 && cp <= 0xff60) ||
    (cp >= 0xac00 && cp <= 0xd7af) ||
    cp === 0x2013 || cp === 0x2014 ||
    cp === 0x2018 || cp === 0x2019 || cp === 0x201c || cp === 0x201d
  );
}

/** Symbols with a math spelling of their own. */
const MATH_SPELLINGS: ReadonlyMap<string, string> = new Map([
  ["′", "'"],
  ["″", "''"],
  ["‴", "'''"],
  ["℃", "{}^{\\circ}\\mathrm{C}"],
  ["￥", "{\\yen}"],
  ["￡", "{\\pounds}"],
]);

const NEEDS_PREPARATION =
  /[–—‘’“”′-‴℃　-ヿ一-龯가-힯＀-｠￡￥]/u;

type GroupKind = "math" | "text" | "font";

/** `latex` with its math-level text put in `\text{}` (see the module note). */
/** A delimiter's control word run into the letters after it (\\rangledx): an
 * import read a closing bracket and the run after it as one token, and such a
 * stored equation printed as its source. No TeX command starts with these. */
const FUSED_DELIMITER_WORD = /\\(rangle|langle|rfloor|lfloor|rceil|lceil)(?=[A-Za-z])/g;

export function prepareLatexForKatex(stored: string): string {
  const latex = stored.replace(FUSED_DELIMITER_WORD, "\\$1 ");
  if (!NEEDS_PREPARATION.test(latex)) return latex;
  const groups: GroupKind[] = [];
  // The kind the `{` right after a command name opens (null: inherit).
  let pending: GroupKind | null = null;
  let out = "";
  let i = 0;
  const n = latex.length;
  const current = (): GroupKind => groups[groups.length - 1] ?? "math";
  while (i < n) {
    const ch = latex[i];
    if (ch === "\\") {
      let j = i + 1;
      if (j < n && /[A-Za-z]/u.test(latex[j])) {
        while (j < n && /[A-Za-z]/u.test(latex[j])) j += 1;
        const name = latex.slice(i + 1, j);
        if (latex[j] === "*") j += 1;
        pending = TEXT_ARGUMENT_COMMANDS.has(name) ? "text" : FONT_ARGUMENT_COMMANDS.has(name) ? "font" : null;
        out += latex.slice(i, j);
        i = j;
        continue;
      }
      // A control symbol (\\ \{ \} \, \% …): two characters, taken as they are.
      const symbol = latex.slice(i, i + 1 + (j < n ? String.fromCodePoint(latex.codePointAt(j)!).length : 0));
      out += symbol;
      i += symbol.length;
      pending = null;
      continue;
    }
    if (ch === "%") {
      // A comment runs to the end of its line.
      const end = latex.indexOf("\n", i);
      const stop = end < 0 ? n : end;
      out += latex.slice(i, stop);
      i = stop;
      continue;
    }
    if (ch === "{") {
      groups.push(pending ?? current());
      pending = null;
      out += ch;
      i += 1;
      continue;
    }
    if (ch === "}") {
      groups.pop();
      pending = null;
      out += ch;
      i += 1;
      continue;
    }
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r") {
      // `\text {…}` still names the group after the space.
      out += ch;
      i += 1;
      continue;
    }
    pending = null;
    const cp = latex.codePointAt(i)!;
    const char = String.fromCodePoint(cp);
    if (current() === "math") {
      if (isTextModeCodePoint(cp)) {
        let j = i + char.length;
        while (j < n) {
          const next = latex.codePointAt(j)!;
          if (!isTextModeCodePoint(next)) break;
          j += String.fromCodePoint(next).length;
        }
        out += `\\text{${latex.slice(i, j)}}`;
        i = j;
        continue;
      }
      const spelling = MATH_SPELLINGS.get(char);
      if (spelling !== undefined) {
        out += spelling;
        i += char.length;
        continue;
      }
    }
    out += char;
    i += char.length;
  }
  return out;
}

/** The options of a passive rendering, with the caller's own on top. */
function passiveOptions(options: katex.KatexOptions | undefined): katex.KatexOptions {
  return { throwOnError: false, ...options, strict: options?.strict ?? katexStrict };
}

/** `katex.renderToString` through this door. */
export function renderKatexToString(latex: string, options?: katex.KatexOptions): string {
  return katex.renderToString(prepareLatexForKatex(latex), passiveOptions(options));
}

/** `katex.render` through this door. */
export function renderKatexInto(element: HTMLElement, latex: string, options?: katex.KatexOptions): void {
  katex.render(prepareLatexForKatex(latex), element, passiveOptions(options));
}

/** The options rehype-katex is given wherever an answer or a note is rendered. */
export const REHYPE_KATEX_OPTIONS: Readonly<katex.KatexOptions> = Object.freeze({ strict: katexStrict });

interface HastText {
  type: "text";
  value: string;
}
interface HastNode {
  type: string;
  tagName?: string;
  value?: string;
  properties?: { className?: unknown };
  children?: HastNode[];
}

const MATH_CLASSES = ["language-math", "math-inline", "math-display"];

function isMathElement(node: HastNode): boolean {
  if (node.type !== "element") return false;
  const classes = node.properties?.className;
  return Array.isArray(classes) && classes.some((name) => MATH_CLASSES.includes(String(name)));
}

function prepareMathElements(node: HastNode): void {
  if (!node.children) return;
  for (const child of node.children) {
    if (isMathElement(child) && child.children?.every((grand) => grand.type === "text")) {
      const source = child.children.map((grand) => grand.value ?? "").join("");
      const prepared = prepareLatexForKatex(source);
      if (prepared !== source) child.children = [{ type: "text", value: prepared } satisfies HastText];
      continue;
    }
    prepareMathElements(child);
  }
}

/** A rehype plugin, run before rehype-katex: the math elements remark-math
 * made read as {@link prepareLatexForKatex} has them. */
export function rehypePrepareKatexMath() {
  return (tree: HastNode): void => {
    prepareMathElements(tree);
  };
}
