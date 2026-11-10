/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bold and Italic on the letters of an equation, as Office Math applies them.
 *
 * Microsoft (Murray Sargent, "Using Math Italic and Bold in Word 2007"): in a
 * math zone the italic and bold buttons (Ctrl+I, Ctrl+B) change LETTERS only
 * -- Latin and Greek -- between math italic, upright, bold and bold italic;
 * operators such as + and = and delimiters stay upright and normal weight.
 * The state is a property of each math run (`m:sty` p / b / i / bi), so
 * formatting part of an equation never changes its structure: a radical,
 * fraction or script keeps its arguments and only the letters in them
 * change. A letter is math italic unless it was made upright.
 *
 * Here that is done on the LaTeX an equation is stored as. The letters'
 * current state is read through the font wrappers around them (`\mathrm`,
 * `\mathbf`, `\boldsymbol`, MathLive's `\mathbfit` / `\bm` /
 * `\mathbf{\mathbfit{}}`), every wrapper of that kind is taken off, and each
 * run of letters that ends up upright or bold is wrapped again in the
 * document's own spelling: `\mathrm` (upright), `\boldsymbol` (bold italic),
 * `\mathbf` (bold upright). Bold also takes digits (Office's math bold
 * digits); italic never makes a digit italic. Everything else -- commands
 * with arguments, scripts, `\left`/`\right`, environments, `\text{…}` --
 * is copied as it was, its arguments formatted in place.
 *
 * Button state as Office shows it: Italic is on when every letter is italic,
 * Bold when every letter (and digit) is bold; pressing turns it off when it
 * is on and on otherwise.
 */

export type EquationLetterFormatKind = "bold" | "italic" | "underline";

/**
 * A run's look: math italic or upright, bold, and Word's run underline
 * (`w:u` / PowerPoint's `u`), which the document spells `\uline{…}` around
 * each run -- one line under each run on its own baseline, a fraction's
 * numerator and denominator each underlined, as Word draws Ctrl+U on part of
 * an equation. `\underline{…}` is the other thing Word has, the Bar
 * structure (`m:bar`): one line under all it holds, kept as a structure here.
 */
type LetterStyle = { readonly upright: boolean; readonly bold: boolean; readonly under: boolean };
type FontStyle = { readonly upright: boolean; readonly bold: boolean };

const PLAIN: LetterStyle = { upright: false, bold: false, under: false };

/** The document's spelling of the run underline. */
export const EQUATION_RUN_UNDERLINE_COMMAND = "\\uline";

/** Font wrappers: what each states for the letters in it. */
const STYLE_WRAPPERS: Record<string, FontStyle> = {
  mathrm: { upright: true, bold: false },
  mathup: { upright: true, bold: false },
  mathit: { upright: false, bold: false },
  mathnormal: { upright: false, bold: false },
  mathbf: { upright: true, bold: true },
  boldsymbol: { upright: false, bold: true },
  bm: { upright: false, bold: true },
  mathbfit: { upright: false, bold: true },
  pmb: { upright: false, bold: true },
};

/** Commands whose argument is not math letters: copied as they are. */
const OPAQUE_ARGUMENT = new Set([
  "text",
  "textrm",
  "textbf",
  "textit",
  "textsf",
  "texttt",
  "textnormal",
  "mbox",
  "operatorname",
  "mathbb",
  "mathcal",
  "mathscr",
  "mathfrak",
  "mathsf",
  "mathtt",
  "begin",
  "end",
  "label",
  "tag",
  "hspace",
  "vspace",
  "kern",
  "mkern",
  "mspace",
  "htmlClass",
  "htmlId",
  "htmlData",
  "htmlStyle",
  "cssId",
  "class",
]);

/** Lowercase Greek: math italic by default, like Latin letters. */
const LOWER_GREEK = new Set([
  "alpha", "beta", "gamma", "delta", "epsilon", "varepsilon", "zeta", "eta", "theta", "vartheta",
  "iota", "kappa", "varkappa", "lambda", "mu", "nu", "xi", "omicron", "pi", "varpi", "rho", "varrho",
  "sigma", "varsigma", "tau", "upsilon", "phi", "varphi", "chi", "psi", "omega", "digamma",
]);
/** Capital Greek: upright by default (Office), but bold takes them. */
const UPPER_GREEK = new Set([
  "Gamma", "Delta", "Theta", "Lambda", "Xi", "Pi", "Sigma", "Upsilon", "Phi", "Psi", "Omega",
]);

// ---- Tokens ------------------------------------------------------------

type Token =
  | { readonly t: "cmd"; readonly name: string; readonly raw: string }
  | { readonly t: "open" }
  | { readonly t: "close" }
  | { readonly t: "char"; readonly raw: string };

function tokenize(latex: string): Token[] {
  const tokens: Token[] = [];
  for (let i = 0; i < latex.length; ) {
    const c = latex[i];
    if (c === "\\") {
      const rest = latex.slice(i + 1);
      const word = /^[A-Za-z]+/.exec(rest);
      if (word) {
        tokens.push({ t: "cmd", name: word[0], raw: `\\${word[0]}` });
        i += 1 + word[0].length;
      } else {
        const sym = rest[0] ?? "";
        tokens.push({ t: "cmd", name: sym, raw: `\\${sym}` });
        i += 1 + sym.length;
      }
      continue;
    }
    if (c === "{") tokens.push({ t: "open" });
    else if (c === "}") tokens.push({ t: "close" });
    else tokens.push({ t: "char", raw: c });
    i += 1;
  }
  return tokens;
}

// ---- Tree --------------------------------------------------------------

/** A letter/digit/symbol the buttons may change, or anything copied as is. */
type Leaf = {
  readonly k: "leaf";
  readonly raw: string;
  /**
   * "letter": italic, bold and underline change it; "upperGreek" and
   * "digit": bold and underline; "symbol" (an operator, a delimiter, a large
   * operator, a function name): underline only.
   */
  readonly kind: "letter" | "upperGreek" | "digit" | "symbol";
  style: LetterStyle;
};
/** `{…}`, or (`splice`) a font wrapper's body, written back without braces. */
type Group = { readonly k: "group"; readonly items: Node[]; readonly splice?: boolean };
/** Verbatim text (a command and its untouched arguments, an operator). */
type Raw = { readonly k: "raw"; readonly raw: string };
/** A command whose arguments hold math: each argument formatted in place. */
type Structure = {
  readonly k: "structure";
  readonly head: string;
  readonly args: readonly (Group | Raw)[];
};
type Node = Leaf | Group | Raw | Structure;

const ARITY: Record<string, number> = {
  frac: 2, dfrac: 2, tfrac: 2, cfrac: 2, binom: 2, dbinom: 2, tbinom: 2,
  overset: 2, underset: 2, stackrel: 2, sqrt: 1, overline: 1, underline: 1,
  hat: 1, widehat: 1, bar: 1, vec: 1, tilde: 1, widetilde: 1, dot: 1, ddot: 1,
  dddot: 1, acute: 1, grave: 1, breve: 1, check: 1, mathring: 1,
  overrightarrow: 1, overleftarrow: 1, overleftrightarrow: 1, underrightarrow: 1,
  underleftarrow: 1, overbrace: 1, underbrace: 1, boxed: 1, cancel: 1, bcancel: 1,
  xcancel: 1, phantom: 1, hphantom: 1, vphantom: 1, smash: 1, displaystyle: 0,
  xrightarrow: 1, xleftarrow: 1,
};

class Parser {
  private at = 0;
  constructor(private readonly tokens: readonly Token[]) {}

  parseSequence(style: LetterStyle, untilClose: boolean): Node[] {
    const items: Node[] = [];
    while (this.at < this.tokens.length) {
      const token = this.tokens[this.at];
      if (token.t === "close") {
        if (untilClose) return items;
        // An unbalanced brace: copied, never dropped.
        this.at += 1;
        items.push({ k: "raw", raw: "}" });
        continue;
      }
      items.push(this.parseItem(style));
    }
    return items;
  }

  /** One argument: a braced group, or the single token after a command. */
  private parseArgument(style: LetterStyle): Group | Raw {
    const token = this.tokens[this.at];
    if (!token) return { k: "raw", raw: "" };
    if (token.t === "open") {
      this.at += 1;
      const items = this.parseSequence(style, true);
      if (this.tokens[this.at]?.t === "close") this.at += 1;
      return { k: "group", items };
    }
    return { k: "group", items: [this.parseItem(style)] };
  }

  /** An argument copied verbatim (balanced braces). */
  private rawArgument(): string {
    const token = this.tokens[this.at];
    if (!token) return "";
    if (token.t !== "open") {
      this.at += 1;
      return tokenRaw(token);
    }
    let depth = 0;
    let out = "";
    while (this.at < this.tokens.length) {
      const next = this.tokens[this.at];
      this.at += 1;
      out += tokenRaw(next);
      if (next.t === "open") depth += 1;
      else if (next.t === "close") {
        depth -= 1;
        if (depth === 0) break;
      }
    }
    return out;
  }

  /** `[...]` after a command (\sqrt's degree), formatted in place. */
  private optionalArgument(style: LetterStyle): Group | null {
    const token = this.tokens[this.at];
    if (!token || token.t !== "char" || token.raw !== "[") return null;
    this.at += 1;
    const items: Node[] = [];
    while (this.at < this.tokens.length) {
      const next = this.tokens[this.at];
      if (next.t === "char" && next.raw === "]") {
        this.at += 1;
        break;
      }
      if (next.t === "close") break;
      items.push(this.parseItem(style));
    }
    return { k: "group", items };
  }

  private parseItem(style: LetterStyle): Node {
    const token = this.tokens[this.at];
    this.at += 1;
    if (token.t === "open") {
      const items = this.parseSequence(style, true);
      if (this.tokens[this.at]?.t === "close") this.at += 1;
      return { k: "group", items };
    }
    if (token.t === "char") {
      if (token.raw === "^" || token.raw === "_") {
        return { k: "structure", head: token.raw, args: [this.parseArgument(style)] };
      }
      if (/[A-Za-z]/.test(token.raw)) return { k: "leaf", raw: token.raw, kind: "letter", style };
      if (/[0-9]/.test(token.raw)) return { k: "leaf", raw: token.raw, kind: "digit", style };
      // Space and a matrix's column break are not characters of a run.
      if (/\s/.test(token.raw) || token.raw === "&") return { k: "raw", raw: token.raw };
      return { k: "leaf", raw: token.raw, kind: "symbol", style };
    }
    if (token.t === "close") return { k: "raw", raw: "}" };
    const name = token.name;
    const wrapper = STYLE_WRAPPERS[name];
    if (wrapper || name === "uline") {
      // A font wrapper or the run underline: its characters take its look
      // (the innermost font wins, so MathLive's \mathbf{\mathbfit{x}} reads
      // as bold italic; the underline adds to it). It is not kept: the runs
      // are wrapped again once formatted.
      const inner = this.parseArgument(wrapper ? { ...wrapper, under: style.under } : { ...style, under: true });
      return { k: "group", items: inner.k === "group" ? inner.items : [inner], splice: true };
    }
    if (LOWER_GREEK.has(name)) return { k: "leaf", raw: token.raw, kind: "letter", style };
    if (UPPER_GREEK.has(name)) return { k: "leaf", raw: token.raw, kind: "upperGreek", style };
    if (name === "left" || name === "right" || name === "middle" || name === "big" || name === "Big" || name === "bigg" || name === "Bigg" || name === "bigl" || name === "bigr" || name === "Bigl" || name === "Bigr") {
      // The delimiter after it belongs to it.
      const delimiter = this.tokens[this.at];
      if (delimiter && delimiter.t !== "open" && delimiter.t !== "close") {
        this.at += 1;
        return { k: "raw", raw: `${token.raw}${tokenRaw(delimiter)}` };
      }
      return { k: "raw", raw: token.raw };
    }
    if (name === "textcolor" || name === "colorbox") {
      const color = this.rawArgument();
      return { k: "structure", head: `${token.raw}${color}`, args: [this.parseArgument(style)] };
    }
    if (name === "color") return { k: "raw", raw: `${token.raw}${this.rawArgument()}` };
    if (OPAQUE_ARGUMENT.has(name)) {
      let raw = token.raw;
      // \begin{env} keeps its name (and a matrix's column spec) verbatim.
      raw += this.rawArgument();
      if (name === "begin" && this.tokens[this.at]?.t === "open" && /^(array|alignedat|subarray)$/.test(raw.slice(7, -1))) {
        raw += this.rawArgument();
      }
      return { k: "raw", raw };
    }
    const arity = ARITY[name];
    if (arity !== undefined) {
      const optional = name === "sqrt" ? this.optionalArgument(style) : null;
      const args: (Group | Raw)[] = [];
      for (let index = 0; index < arity; index += 1) args.push(this.parseArgument(style));
      return {
        k: "structure",
        head: token.raw,
        args: optional ? [{ k: "raw", raw: "[" }, optional, { k: "raw", raw: "]" }, ...args] : args,
      };
    }
    // Spacing and line structure are no characters of a run.
    if (LAYOUT_COMMANDS.has(name)) return { k: "raw", raw: token.raw };
    // Any other command: a symbol, an operator or a function name (\pm,
    // \infty, \sum, \sin): a character of its run.
    return { k: "leaf", raw: token.raw, kind: "symbol", style };
  }
}

/** Commands that space or break the formula rather than draw a character. */
const LAYOUT_COMMANDS = new Set([
  ",", ";", ":", "!", " ", ">", "quad", "qquad", "\\", "cr", "newline", "hline", "displaystyle",
  "textstyle", "scriptstyle", "scriptscriptstyle", "limits", "nolimits", "enspace", "thinspace",
  "medspace", "thickspace", "negthinspace", "nobreak", "allowbreak",
]);

function tokenRaw(token: Token): string {
  return token.t === "open" ? "{" : token.t === "close" ? "}" : token.raw;
}

function parse(latex: string): Node[] {
  return new Parser(tokenize(latex)).parseSequence(PLAIN, false);
}

// ---- Reading and writing ----------------------------------------------

function* leaves(nodes: readonly Node[]): Generator<Leaf> {
  for (const node of nodes) {
    if (node.k === "leaf") yield node;
    else if (node.k === "group") yield* leaves(node.items);
    else if (node.k === "structure") for (const arg of node.args) if (arg.k === "group") yield* leaves(arg.items);
  }
}

/** The leaves a kind of formatting changes. */
function affected(leaf: Leaf, kind: EquationLetterFormatKind): boolean {
  if (kind === "italic") return leaf.kind === "letter";
  if (kind === "bold") return leaf.kind !== "symbol";
  return true;
}

/**
 * Whether the button shows on for these letters: Italic when every letter is
 * math italic, Bold when every letter and digit is bold. `null` when the
 * equation holds nothing the button changes.
 */
export function equationLettersFormatState(latex: string, kind: EquationLetterFormatKind): boolean | null {
  let any = false;
  for (const leaf of leaves(parse(latex))) {
    if (!affected(leaf, kind)) continue;
    any = true;
    const on = kind === "italic" ? !leaf.style.upright : kind === "bold" ? leaf.style.bold : leaf.style.under;
    if (!on) return false;
  }
  return any ? true : null;
}

const WRAPPER_FOR = (style: FontStyle): string | null =>
  style.bold ? (style.upright ? "\\mathbf" : "\\boldsymbol") : style.upright ? "\\mathrm" : null;

/** A leaf's font wrapper: a symbol never takes one (Office keeps operators upright, normal weight). */
const FONT_OF = (leaf: Leaf): string | null => (leaf.kind === "symbol" ? null : WRAPPER_FOR(leaf.style));

/** LaTeX pieces joined, a control word kept apart from a letter after it. */
function joiner(): { append(text: string): void; text(): string } {
  let out = "";
  return {
    append(text: string) {
      if (!text) return;
      if (/\\[A-Za-z]+$/.test(out) && /^[A-Za-z]/.test(text)) out += " ";
      out += text;
    },
    text: () => out,
  };
}

/**
 * One run's characters in their fonts: consecutive leaves of the same font
 * in one wrapper, an upright run taking the symbols between its letters
 * (they draw the same inside \mathrm), a bold one never.
 */
function serializeFonts(run: readonly Leaf[]): string {
  const out = joiner();
  let at = 0;
  while (at < run.length) {
    const wrapper = FONT_OF(run[at]);
    if (!wrapper) {
      out.append(run[at].raw);
      at += 1;
      continue;
    }
    let end = at + 1;
    let last = at;
    while (end < run.length) {
      const font = FONT_OF(run[end]);
      if (font === wrapper) {
        last = end;
        end += 1;
        continue;
      }
      if (font === null && wrapper === "\\mathrm") {
        end += 1;
        continue;
      }
      break;
    }
    const body = joiner();
    for (const leaf of run.slice(at, last + 1)) body.append(leaf.raw);
    out.append(`${wrapper}{${body.text()}}`);
    at = last + 1;
  }
  return out.text();
}

/** A font wrapper's body back in the sequence it sat in. */
function spliced(nodes: readonly Node[]): Node[] {
  return nodes.flatMap((node) => (node.k === "group" && node.splice ? spliced(node.items) : [node]));
}

function serialize(input: readonly Node[]): string {
  const nodes = spliced(input);
  const out = joiner();
  const scripted = (at: number) => {
    const after = nodes[at + 1];
    return after?.k === "structure" && (after.head === "^" || after.head === "_");
  };
  let index = 0;
  while (index < nodes.length) {
    const node = nodes[index];
    if (node.k === "leaf") {
      // A run: consecutive characters with the same underline. A character
      // a script follows is a run of its own, so the script stays on it.
      const run: Leaf[] = [node];
      let end = index + 1;
      if (!scripted(index)) {
        while (end < nodes.length) {
          const next = nodes[end];
          if (next.k !== "leaf" || next.style.under !== node.style.under || scripted(end)) break;
          run.push(next);
          end += 1;
        }
      }
      const body = serializeFonts(run);
      out.append(node.style.under ? `${EQUATION_RUN_UNDERLINE_COMMAND}{${body}}` : body);
      index = end;
      continue;
    }
    if (node.k === "raw") out.append(node.raw);
    else if (node.k === "group") out.append(`{${serialize(node.items)}}`);
    else {
      out.append(node.head);
      for (const arg of node.args) {
        if (arg.k === "raw") out.append(arg.raw);
        else if (node.head.startsWith("\\sqrt") && arg === node.args[1] && node.args.length === 4) out.append(serialize(arg.items));
        else out.append(`{${serialize(arg.items)}}`);
      }
    }
    index += 1;
  }
  return out.text();
}

/**
 * The equation with Bold, Italic or Underline set (`on`) or cleared on every character it
 * changes -- or toggled as the button does when `on` is omitted (off when
 * it shows on, else on). Structure, operators and everything else are
 * kept. The LaTeX is returned unchanged when nothing changes.
 */
export function equationLatexWithLetterFormat(
  latex: string,
  kind: EquationLetterFormatKind,
  on?: boolean,
): string {
  const nodes = parse(latex);
  const target = on ?? !(equationLettersFormatState(latex, kind) ?? false);
  let changed = false;
  for (const leaf of leaves(nodes)) {
    if (!affected(leaf, kind)) continue;
    const next: LetterStyle =
      kind === "italic"
        ? { ...leaf.style, upright: leaf.kind === "upperGreek" ? leaf.style.upright : !target }
        : kind === "bold"
          ? { ...leaf.style, upright: leaf.kind === "digit" ? true : leaf.style.upright, bold: target }
          : { ...leaf.style, under: target };
    if (next.upright !== leaf.style.upright || next.bold !== leaf.style.bold || next.under !== leaf.style.under) {
      changed = true;
    }
    leaf.style = next;
  }
  // Digits: upright always; bold digits wrap as \mathbf, plain ones bare.
  for (const leaf of leaves(nodes)) {
    if (leaf.kind === "digit") leaf.style = { upright: leaf.style.bold, bold: leaf.style.bold, under: leaf.style.under };
  }
  if (!changed) return latex;
  return serialize(nodes);
}
